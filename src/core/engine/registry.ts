import { ContentTypeExtractor } from '../extractors/ContentTypeExtractor';
import { GroupExtractor } from '../extractors/GroupExtractor';
import { ItemExtractor } from '../extractors/ItemExtractor';
import { ListExtractor } from '../extractors/ListExtractor';
import { ListFieldExtractor } from '../extractors/ListFieldExtractor';
import { SiteFieldExtractor } from '../extractors/SiteFieldExtractor';
import { ViewExtractor } from '../extractors/ViewExtractor';
import type { FetchLike } from '../http/raw';
import type { IExtractor } from '../model';
import { ContentTypeProvider } from '../providers/ContentTypeProvider';
import { GroupProvider } from '../providers/GroupProvider';
import { ItemLookupProvider } from '../providers/ItemLookupProvider';
import { ItemProvider, type IItemProviderOptions } from '../providers/ItemProvider';
import { ListFieldProvider } from '../providers/ListFieldProvider';
import { ListProvider } from '../providers/ListProvider';
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
    siteField: new SiteFieldProvider(),
    contentType: new ContentTypeProvider(fetchImpl),
    list: new ListProvider(fetchImpl),
    listField: new ListFieldProvider(),
    view: new ViewProvider(),
    items: new ItemProvider(items),
    itemLookups: new ItemLookupProvider(items)
  } as ProviderMap;
}

/**
 * Extractors in the order the Setup runs them: lists before list columns, views and items, which write into
 * the template's list entries.
 */
export function createExtractors(): Array<IExtractor<unknown>> {
  return [
    new GroupExtractor(),
    new SiteFieldExtractor(),
    new ContentTypeExtractor(),
    new ListExtractor(),
    new ListFieldExtractor(),
    new ViewExtractor(),
    new ItemExtractor()
  ] as Array<IExtractor<unknown>>;
}
