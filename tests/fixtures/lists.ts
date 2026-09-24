// REST responses for /_api/web/lists (odata=nometadata), modelled on the Forrás site (spike 03 A).
import type { IListInfoLike } from '../../src/core/lists';

const WEB = '/sites/Forras';

const list = (Id: string, Title: string, url: string, BaseTemplate: number, flags: Partial<IListInfoLike> = {}): IListInfoLike => ({
  Id,
  Title,
  BaseTemplate,
  ItemCount: 0,
  RootFolder: { ServerRelativeUrl: `${WEB}/${url}` },
  ...flags
});

export const documents = list('11111111-0000-4000-8000-000000000001', 'Dokumentumok', 'Shared Documents', 101, {
  EnableVersioning: true,
  MajorVersionLimit: 500,
  EnableMinorVersions: false,
  ForceCheckout: false,
  OnQuickLaunch: true
});
export const events = list('11111111-0000-4000-8000-000000000002', 'Események', 'Lists/Events', 106);
export const tesztLista = list('11111111-0000-4000-8000-000000000003', 'Teszt lista', 'Lists/Teszt lista', 100, {
  Description: 'Próba',
  ItemCount: 1,
  OnQuickLaunch: true,
  EnableAttachments: true,
  EnableVersioning: true,
  MajorVersionLimit: 50,
  ContentTypesEnabled: true
});
// SharePoint dropped the "á" from the URL.
export const lookupForras = list('11111111-0000-4000-8000-000000000004', 'Teszt lookup forrás', 'Lists/Teszt lookup forrs', 100, { ItemCount: 7, EnableAttachments: true });

export const systemLists: IListInfoLike[] = [
  list('22222222-0000-4000-8000-000000000001', 'Űrlapsablonok', 'FormServerTemplates', 101, { IsSystemList: true }),
  list('22222222-0000-4000-8000-000000000002', 'Webhelyeszközök', 'SiteAssets', 101, { IsSystemList: true }),
  list('22222222-0000-4000-8000-000000000003', 'Weblapok', 'SitePages', 119, { IsSystemList: true }),
  list('22222222-0000-4000-8000-000000000004', 'Stílustár', 'Style Library', 101, { IsSystemList: true, IsCatalog: true }),
  list('22222222-0000-4000-8000-000000000005', 'TaxonomyHiddenList', 'Lists/TaxonomyHiddenList', 100, { Hidden: true, IsSystemList: true })
];

export const sourceLists: IListInfoLike[] = [documents, events, ...systemLists, tesztLista, lookupForras];
export const SOURCE_WEB = { Url: 'https://contoso.sharepoint.com/sites/Forras', ServerRelativeUrl: WEB, Title: 'Forrás' };

// Structure of the source lists: folders (relative to the list root, incl. system ones) and list content types.
export const PROJEKT_SITE_CT = '0x0100A1B2C3D4E5F60718293A4B5C6D7E8F90';
export const sourceFolders: { [listUrl: string]: string[] } = {
  '/sites/Forras/Shared Documents': ['Forms', '2026', '2026/Q3'],
  '/sites/Forras/Lists/Teszt lista': ['Attachments', 'Mappa']
};
export const sourceListCts: { [listUrl: string]: { ordered: string[]; all: string[] } } = {
  '/sites/Forras/Lists/Teszt lista': {
    ordered: [`${PROJEKT_SITE_CT}00AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA`, '0x0100BBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBB'],
    all: ['0x0100BBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBB', `${PROJEKT_SITE_CT}00AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA`, '0x012000CCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCC']
  }
};
