import { ContentTypeExtractor } from '../extractors/ContentTypeExtractor';
import { GroupExtractor } from '../extractors/GroupExtractor';
import { ListExtractor } from '../extractors/ListExtractor';
import { ListFieldExtractor } from '../extractors/ListFieldExtractor';
import { SiteFieldExtractor } from '../extractors/SiteFieldExtractor';
import { ViewExtractor } from '../extractors/ViewExtractor';
import type { FetchLike } from '../http/raw';
import type { IExtractor } from '../model';
import { ContentTypeProvider } from '../providers/ContentTypeProvider';
import { GroupProvider } from '../providers/GroupProvider';
import { ListFieldProvider } from '../providers/ListFieldProvider';
import { ListProvider } from '../providers/ListProvider';
import { SiteFieldProvider } from '../providers/SiteFieldProvider';
import { ViewProvider } from '../providers/ViewProvider';
import type { ProviderMap } from './engine';

/** Providers for every artifact kind of phase 1. `fetchImpl` is used for CSOM/raw REST calls (tests). */
export function createProviders(fetchImpl?: FetchLike): ProviderMap {
  return {
    group: new GroupProvider(fetchImpl),
    siteField: new SiteFieldProvider(),
    contentType: new ContentTypeProvider(fetchImpl),
    list: new ListProvider(fetchImpl),
    listField: new ListFieldProvider(),
    view: new ViewProvider()
  } as ProviderMap;
}

/**
 * Extractors in the order the Setup runs them: lists before list columns and views, which write into the
 * template's list entries.
 */
export function createExtractors(): Array<IExtractor<unknown>> {
  return [
    new GroupExtractor(),
    new SiteFieldExtractor(),
    new ContentTypeExtractor(),
    new ListExtractor(),
    new ListFieldExtractor(),
    new ViewExtractor()
  ] as Array<IExtractor<unknown>>;
}
