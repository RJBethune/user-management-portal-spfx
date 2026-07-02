/**
 * O365 group ids are GUIDs; SharePoint site-group ids are integers.
 * Used to branch the web part between the PROD (Graph) path and the DEV
 * (direct SharePoint REST) path on the same build.
 */
export function isSharePointGroup(groupId: string | undefined): boolean {
  return /^\d+$/.test((groupId || '').toString().trim());
}

/** Human label for the kind of group, derived from the id format (GUID = M365, integer = SharePoint). */
export function groupKindLabel(groupId: string | undefined): string {
  return isSharePointGroup(groupId) ? 'SharePoint Group' : 'Microsoft 365 Group';
}
