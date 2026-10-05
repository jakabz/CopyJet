import type { SPFI } from '@pnp/sp';
import '@pnp/sp/webs';
import '@pnp/sp/lists';
import '@pnp/sp/items';
import '@pnp/sp/fields';
import '@pnp/sp/attachments';
import '@pnp/sp/security';
import { readAssociatedGroups } from '../groups';
import { isHttpStatus } from '../http/status';
import type { Logger } from '../logger/Logger';
import type { IRunState, IRunStore } from './runState';

export const RUN_LOG_LIST = 'CopyJetLog';
const STATE_FILE = 'state.json';
const FIELDS: Array<[string, string]> = [
  ['CJRunId', 'RunId'],
  ['CJTemplate', 'Template'],
  ['CJChecksum', 'Checksum'],
  ['CJStatus', 'Status']
];
/** SPAddFieldOptions.AddFieldInternalNameHint: without it SharePoint makes the internal name from DisplayName. */
const INTERNAL_NAME_HINT = 8;

interface IRunItem {
  Id: number;
  CJRunId: string;
  CJStatus: string;
}

/**
 * Run states in the target site's hidden CopyJetLog list (spike 14): one item per run with its ID, template,
 * checksum and status in columns, and the full state as the item's state.json attachment, overwritten in place
 * on every save (a note field would also hold 300 000 characters, but item ID maps of a large list can be
 * bigger). The list is created on the first save, without inheritance: the owners group and the installer.
 */
export class SpRunStore implements IRunStore {
  private _list: Promise<void> | undefined;
  private readonly _items: { [runId: string]: number } = {};
  private readonly _status: { [runId: string]: string } = {};

  constructor(
    private readonly _sp: SPFI,
    private readonly _log?: Logger
  ) {}

  public async findUnfinished(checksum: string): Promise<IRunState | undefined> {
    if (!(await this._exists())) return undefined; // no install has run here yet
    await this._ensureList();
    const found = await this._listRef()
        .items.select('Id', 'CJRunId', 'CJStatus')
        .filter(`CJChecksum eq '${checksum}' and CJStatus ne 'completed'`)
        .orderBy('Modified', false)
        .top(1)<IRunItem[]>();
    if (!found.length) return undefined;
    const text = await this._listRef().items.getById(found[0].Id).attachmentFiles.getByName(STATE_FILE).getText();
    const state = JSON.parse(text) as IRunState;
    this._items[state.runId] = found[0].Id;
    this._status[state.runId] = found[0].CJStatus;
    return state;
  }

  public async save(state: IRunState): Promise<void> {
    await this._ensureList();
    const body = new Blob([JSON.stringify(state)], { type: 'application/json' });
    let id = this._items[state.runId];
    if (id === undefined) {
      const found = await this._listRef().items.select('Id', 'CJStatus').filter(`CJRunId eq '${state.runId}'`).top(1)<IRunItem[]>();
      if (found.length) {
        id = this._items[state.runId] = found[0].Id;
        this._status[state.runId] = found[0].CJStatus;
      }
    }
    if (id === undefined) {
      const added: { Id: number } = await this._listRef().items.add({ Title: `${state.templateName} – ${state.started}`, CJRunId: state.runId, CJTemplate: state.templateName, CJChecksum: state.checksum, CJStatus: state.status });
      id = this._items[state.runId] = added.Id;
      this._status[state.runId] = state.status;
      await this._listRef().items.getById(id).attachmentFiles.add(STATE_FILE, body);
      return;
    }
    await this._listRef().items.getById(id).attachmentFiles.getByName(STATE_FILE).setContent(body);
    if (this._status[state.runId] !== state.status) {
      await this._listRef().items.getById(id).update({ CJStatus: state.status });
      this._status[state.runId] = state.status;
    }
  }

  private _listRef(): ReturnType<SPFI['web']['lists']['getByTitle']> {
    return this._sp.web.lists.getByTitle(RUN_LOG_LIST);
  }

  private _ensureList(): Promise<void> {
    if (!this._list) {
      this._list = this._create().catch((e) => {
        this._list = undefined;
        throw e;
      });
    }
    return this._list;
  }

  private async _exists(): Promise<boolean> {
    try {
      await this._listRef().select('Id')();
      return true;
    } catch (e) {
      if (isHttpStatus(e, 404)) return false;
      throw e;
    }
  }

  /** Adds the columns the list lacks: 1.8.0.0 created them under their display names (RunId …). */
  private async _ensureFields(): Promise<void> {
    const have = (await this._listRef().fields.select('InternalName')<Array<{ InternalName: string }>>()).map((f) => f.InternalName);
    for (const [name, title] of FIELDS) {
      if (have.indexOf(name) >= 0) continue;
      await this._listRef().fields.createFieldAsXml({
        SchemaXml: `<Field Type="Text" Name="${name}" StaticName="${name}" DisplayName="${title}" />`,
        Options: INTERNAL_NAME_HINT
      });
    }
  }

  private async _create(): Promise<void> {
    if (await this._exists()) {
      await this._ensureFields();
      return;
    }
    await this._sp.web.lists.add(RUN_LOG_LIST, 'CopyJet install runs (resume state).', 100, false, { Hidden: true, NoCrawl: true, OnQuickLaunch: false });
    await this._ensureFields();
    try {
      // Spike 14: without copying, the installer keeps Full Control; the owners group is added beside.
      const owner = (await readAssociatedGroups(this._sp)).owner;
      await this._listRef().breakRoleInheritance(false, true);
      if (owner !== undefined) {
        const full = await this._sp.web.roleDefinitions.getByType(5).select('Id')<{ Id: number }>();
        await this._listRef().roleAssignments.add(owner, full.Id);
      }
    } catch (e) {
      if (this._log) {
        this._log.warn(`The ${RUN_LOG_LIST} list keeps the site's permissions: ${e instanceof Error ? e.message : String(e)}`, { code: 'STATE_LIST_PERMISSIONS' });
      }
    }
  }
}
