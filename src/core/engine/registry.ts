import { ContentTypeExtractor } from '../extractors/ContentTypeExtractor';
import { FileExtractor } from '../extractors/FileExtractor';
import { GroupExtractor } from '../extractors/GroupExtractor';
import { ItemExtractor } from '../extractors/ItemExtractor';
import { NavigationExtractor } from '../extractors/NavigationExtractor';
import { PageExtractor } from '../extractors/PageExtractor';
import { ListExtractor } from '../extractors/ListExtractor';
import { ListFieldExtractor } from '../extractors/ListFieldExtractor';
import { ListSecurityExtractor } from '../extractors/ListSecurityExtractor';
import { SiteFieldExtractor } from '../extractors/SiteFieldExtractor';
import { ViewExtractor } from '../extractors/ViewExtractor';
import type { FetchLike } from '../http/raw';
import type { IExtractor } from '../model';
import { ContentTypeProvider } from '../providers/ContentTypeProvider';
import { FileProvider } from '../providers/FileProvider';
import { GroupProvider } from '../providers/GroupProvider';
import { ItemLookupProvider } from '../providers/ItemLookupProvider';
import { ItemProvider, type IItemProviderOptions } from '../providers/ItemProvider';
import { NavigationProvider } from '../providers/NavigationProvider';
import { PageProvider } from '../providers/PageProvider';
import { ListFieldProvider } from '../providers/ListFieldProvider';
import { ListProvider } from '../providers/ListProvider';
import { ListSecurityProvider } from '../providers/ListSecurityProvider';
import { SiteFieldProvider } from '../providers/SiteFieldProvider';
import { ViewProvider } from '../providers/ViewProvider';
import type { ProviderMap } from './engine';

/**
 * Providers for every supported artifact kind. `fetchImpl` is used for CSOM/raw REST calls, `items` for the
 * item writes (both injected in tests).
 */
export function createProviders(fetchImpl?: FetchLike, items?: IItemProviderOptions): ProviderMap {
  return {
    group: new GroupProvider(fetchImpl),
    siteField: new SiteFieldProvider(fetchImpl),
    contentType: new ContentTypeProvider(fetchImpl),
    list: new ListProvider(fetchImpl),
    listField: new ListFieldProvider(fetchImpl),
    view: new ViewProvider(),
    listSecurity: new ListSecurityProvider(),
    items: new ItemProvider(items),
    itemLookups: new ItemLookupProvider(items),
    files: new FileProvider(),
    page: new PageProvider(),
    navigation: new NavigationProvider()
  } as ProviderMap;
}

/**
 * Extractors in the order the Setup runs them: lists before list columns, views, items and files, which
 * write into the template's list entries.
 */
export function createExtractors(): Array<IExtractor<unknown>> {
  return [
    new GroupExtractor(),
    new SiteFieldExtractor(),
    new ContentTypeExtractor(),
    new ListExtractor(),
    new ListFieldExtractor(),
    new ViewExtractor(),
    new ListSecurityExtractor(),
    new ItemExtractor(),
    new FileExtractor(),
    new PageExtractor(),
    new NavigationExtractor()
  ] as Array<IExtractor<unknown>>;
}
