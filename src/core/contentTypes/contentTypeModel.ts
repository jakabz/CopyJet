import type { IContentType } from '../model';

/** Content type properties CopyJet reads from the REST API. */
export interface IContentTypeInfoLike {
  StringId: string;
  Name: string;
  Group: string;
  Description?: string;
  Hidden?: boolean;
  ReadOnly?: boolean;
  Sealed?: boolean;
  SchemaXml: string;
}

export interface IFieldLinkInfoLike {
  /** The linked field's GUID; needed to change flags on an existing link. */
  Id?: string;
  // eslint-disable-next-line @rushstack/no-new-null -- REST returns null for links without an internal name
  FieldInternalName: string | null;
  Name: string;
  Required: boolean;
  Hidden: boolean;
}

/** SharePoint's canonical form: lowercase 0x prefix, uppercase hex digits. */
export function normalizeContentTypeId(id: string): string {
  return `0x${id.slice(2).toUpperCase()}`;
}

export const contentTypeKey = (id: string): string => `contentType:${normalizeContentTypeId(id)}`;

/**
 * Parent content type ID. Child IDs are either parent + "00" + GUID (32 hex digits, how SharePoint names
 * user-created types) or parent + two hex digits (built-in hierarchy, e.g. 0x0101 → 0x010101).
 */
export function parentIdOf(id: string): string | undefined {
  const upper = id.toUpperCase();
  if (upper === '0X' || upper === '0X01' || upper.length <= 4) {
    return upper === '0X01' ? '0x' : undefined;
  }
  const guidChild = /^(0X[0-9A-F]+)00[0-9A-F]{32}$/.exec(upper);
  const parent = guidChild ? guidChild[1] : upper.slice(0, -2);
  return `0x${parent.slice(2)}`;
}

/**
 * Built-in content types that SharePoint Online provisions without a FeatureId (spike 02): the modern page
 * family. Listed by exact ID so a user-created child of Site Page still counts as custom.
 */
const BUILT_IN_WITHOUT_FEATURE: string[] = [
  '0x0101009D1CB255DA76424F860D91F20E6C4118', // Site Page (Webhelylap)
  '0x0101009D1CB255DA76424F860D91F20E6C4118002A50BFCFB7614729B56886FADA02339B', // Repost Page
  '0x0101009D1CB255DA76424F860D91F20E6C411800EE3B4FE2C29A4DCDB55229ADEB3DBFE4', // HTML page
  '0x0101009D1CB255DA76424F860D91F20E6C411800C57E7DC123744954AACC08D1D5260154', // Content freshness (hidden)
  '0x0101002039C03B61C64EC4A04F5361F3851066', // Display Template (_Hidden)
  '0x0101002039C03B61C64EC4A04F5361F385106605', // Display Template Code (_Hidden)
  '0x010100C5033D6CFB8447359FB795C8A73A2B19' // Design File (_Hidden)
];

/**
 * True for content types created by users or solutions. Feature-deployed (built-in) types carry a
 * FeatureId in their SchemaXml; the few built-ins without one are listed above (spike 02).
 */
export function isCustomContentType(info: IContentTypeInfoLike): boolean {
  if (info.Group === '_Hidden' || info.Hidden) {
    return false;
  }
  if (BUILT_IN_WITHOUT_FEATURE.indexOf(normalizeContentTypeId(info.StringId)) >= 0) {
    return false;
  }
  return !/\sFeatureId="/i.test(info.SchemaXml);
}

function linkName(l: IFieldLinkInfoLike): string {
  return l.FieldInternalName || l.Name;
}

/**
 * Template definition of a content type. Only the field links the type adds on top of its parent are
 * stored; inherited ones come from the parent on the target.
 */
export function toContentTypeDef(info: IContentTypeInfoLike, links: IFieldLinkInfoLike[], parentLinks: IFieldLinkInfoLike[]): IContentType {
  // Keep links the parent lacks, and inherited links whose Required/Hidden the child overrides.
  const isInheritedAsIs = (l: IFieldLinkInfoLike): boolean =>
    parentLinks.some((p) => linkName(p) === linkName(l) && !!p.Required === !!l.Required && !!p.Hidden === !!l.Hidden);
  const def: IContentType = {
    id: normalizeContentTypeId(info.StringId),
    name: info.Name,
    fieldRefs: links
      .filter((l) => !isInheritedAsIs(l))
      .map((l) => {
        const ref: IContentType['fieldRefs'][number] = { internalName: linkName(l) };
        if (l.Required) ref.required = true;
        if (l.Hidden) ref.hidden = true;
        return ref;
      })
  };
  const parentId = parentIdOf(info.StringId);
  if (parentId) def.parentId = parentId;
  if (info.Group) def.group = info.Group;
  if (info.Description) def.description = info.Description;
  return def;
}

/**
 * Differences between a template content type and the one on the target:
 * name / group / description, missing field links ('fieldRefs') and differing Required/Hidden flags on
 * links present on both ('fieldRefFlags'). Extra links on the target are not a difference.
 */
export function compareContentTypes(def: IContentType, target: IContentTypeInfoLike, targetLinks: IFieldLinkInfoLike[]): string[] {
  const changes: string[] = [];
  if (def.name !== target.Name) changes.push('name');
  if ((def.group || '') !== (target.Group || '')) changes.push('group');
  if ((def.description || '') !== (target.Description || '')) changes.push('description');

  let missing = false;
  let flags = false;
  def.fieldRefs.forEach((ref) => {
    const link = targetLinks.filter((l) => linkName(l) === ref.internalName)[0];
    if (!link) {
      missing = true;
    } else if (!!ref.required !== !!link.Required || !!ref.hidden !== !!link.Hidden) {
      flags = true;
    }
  });
  if (missing) changes.push('fieldRefs');
  if (flags) changes.push('fieldRefFlags');
  return changes;
}

/** Field links of `def` that the target content type does not have yet. */
export function missingFieldRefs(def: IContentType, targetLinks: IFieldLinkInfoLike[]): IContentType['fieldRefs'] {
  const present = targetLinks.map(linkName);
  return def.fieldRefs.filter((r) => present.indexOf(r.internalName) < 0);
}
