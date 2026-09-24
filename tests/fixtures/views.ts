// Views of the Forrás lists as REST returns them (spike 06 A); a GRID view is added for the unsupported path.
import type { IViewInfoLike } from '../../src/core/lists';

const v = (Id: string, Title: string, url: string, extra: Partial<IViewInfoLike> = {}): IViewInfoLike => ({
  Id,
  Title,
  ServerRelativeUrl: `/sites/Forras/${url}`,
  DefaultView: false,
  Hidden: false,
  PersonalView: false,
  ViewType: 'HTML',
  RowLimit: 30,
  Paged: true,
  Scope: 0,
  ViewQuery: '',
  CustomFormatter: '',
  ...extra
});

export const allItems = v('88888888-0000-4000-8000-000000000001', 'Minden elem', 'Lists/Teszt lista/AllItems.aspx', { DefaultView: true });
export const nyitott = v('88888888-0000-4000-8000-000000000002', 'Nyitott elemek', 'Lists/Teszt lista/Nyitott elemek.aspx', {
  ViewQuery: '<OrderBy><FieldRef Name="ListaSzoveg" /></OrderBy><Where><IsNotNull><FieldRef Name="ListaSzoveg" /></IsNotNull></Where>'
});
export const ugyfel = v('88888888-0000-4000-8000-000000000003', 'Ügyfél nézet', 'Lists/Teszt lista/gyfl nzet.aspx', { Scope: 1, RowLimit: 100 });
export const racs = v('88888888-0000-4000-8000-000000000004', 'Rács', 'Lists/Teszt lista/Racs.aspx', { ViewType: 'GRID' });
export const allDocs = v('88888888-0000-4000-8000-000000000005', 'Minden dokumentum', 'Shared Documents/Forms/AllItems.aspx', {
  DefaultView: true,
  ViewQuery: '<OrderBy><FieldRef Name="FileLeafRef" /></OrderBy>'
});

export const viewFields: { [id: string]: string[] } = {
  [allItems.Id]: ['LinkTitle', 'Valaki', 'V_x00e1_lassz', 'ListaLookup', 'SiteColumn1'],
  [nyitott.Id]: ['LinkTitle', 'ListaSzoveg'],
  [ugyfel.Id]: ['LinkTitle', 'SiteColumn1', 'ListaLookup'],
  [racs.Id]: ['LinkTitle'],
  [allDocs.Id]: ['DocIcon', 'LinkFilename', 'Modified', 'Editor']
};
