/**
 * O365 group ids are GUIDs; SharePoint site-group ids are integers.
 * Used to branch the web part between the PROD (Graph) path and the DEV
 * (direct SharePoint REST) path on the same build.
 */
export function isSharePointGroup(groupId: string | undefined): boolean {
  return /^\d+$/.test((groupId || '').toString().trim());
}
