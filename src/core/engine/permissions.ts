import type { SPFI } from '@pnp/sp';
import '@pnp/sp/webs';
import '@pnp/sp/security';
import { PermissionKind } from '@pnp/sp/security';

export type InstallPermission = 'ManageLists' | 'ManageWeb' | 'ManagePermissions';

const REQUIRED: Array<[InstallPermission, PermissionKind]> = [
  ['ManageLists', PermissionKind.ManageLists],
  ['ManageWeb', PermissionKind.ManageWeb],
  ['ManagePermissions', PermissionKind.ManagePermissions]
];

/** Permissions the current user lacks on the target web for an install (rendszerterv §9); empty = ok. */
export async function missingInstallPermissions(sp: SPFI): Promise<InstallPermission[]> {
  const has = await Promise.all(REQUIRED.map(([, kind]) => sp.web.currentUserHasPermissions(kind)));
  return REQUIRED.filter((_, i) => !has[i]).map(([name]) => name);
}
