# Admin-to-group assignments

Use **Access assignments** to see which groups each admin can manage, search by admin or group, and edit assignments directly in the app. Membership administration remains on the **Membership** view.

The view is available to users with **Manage Lists** and **Edit Items** on the configured `Group Management Authorized Admins` list. Adding a record also requires **Add Items**. Being assigned to manage a Microsoft 365 group does not grant these list permissions. Page edit mode does not grant access to the view. SharePoint still enforces item-level permissions on each save, including records with unique permissions.

## Assign or remove access

1. Open **Access assignments** and search for the admin or group.
2. Select an admin in the left list, or choose **Add an admin assignment** for a new admin. Their editable groups appear on the right. Narrow columns show one pane at a time with **Back to admins**.
3. For a new assignment, search by name or email and choose an individual person. Existing records keep their admin identity. Keep one record per admin where possible.
4. Select the groups this admin can manage, then choose **Save assignment**. Group search retains selections that are hidden by the search. Historical single-value lookup fields allow one selection.
5. The overview refreshes automatically. The affected admin can reload their membership page to see the updated group list.

To remove access to a group, clear it from **every** assignment record for that admin. Multiple records combine their groups; removing one occurrence may leave another grant. If the lookup field is optional, clearing all groups removes access from that record without deleting it. The overview flags duplicate admin records and shows groups without an assigned admin. Unavailable groups must be explicitly removed before saving. Records with missing users need repair through **Open assignments list**.

Drafts survive switching admins, views and themes, including a new admin assignment. Each assignment record keeps its own draft; unsaved changes are marked in the admin list. Saving one record preserves drafts for other records. Canceling a changed draft asks before discarding it. Refresh is disabled until drafts are saved or discarded. A concurrent edit or deletion blocks saving until you explicitly reload the assignment. Permission errors keep the draft. If a connection failure leaves the save unconfirmed, close the editor and refresh before retrying; the original save may have succeeded. There are no automatic write retries.

The overview reads all available pages of the two configured lists. Search filters the admin list without hiding the selected editor or discarding its draft. Duplicate records show their record number and a warning in the editor. Expand **Groups without an assigned admin** to review the full loaded inventory, regardless of search. The page's `visibleOffices` setting does not filter this administrative view.

## What the assignment means

| Configuration | Effect |
| --- | --- |
| `Managed Groups` row | Identifies a group the application can manage. Its existing `GroupId` and type determine supported actions. |
| `User` plus `OfficeGroupRecord` | Maps an individual admin to one or more managed groups. The editor supports both single-value and multi-value lookup fields. |
| `visibleOffices` web-part property | Narrows the groups shown on a membership page. It does not grant access. |
| `IsTeamsConnected` | Indicates that the Microsoft 365 group is connected to Teams. The overview does not inventory private or shared channel permissions. |

The app submits Microsoft 365 membership requests to the existing request list. Power Automate must independently authorize the request using its system-created `AuthorId` and current assignment records before writing membership changes. SharePoint group changes use the signed-in user's SharePoint permissions. An assignment alone does not grant site access.

The current model treats member and owner actions as group-level authorization. It does not expand admin-team membership or distinguish member managers from owner managers. Those features would require a separately designed and tested server-side authorization change; hiding buttons would not enforce them.

## Existing data and workflow contracts

This upgrade adds no list, column, Graph scope or flow connection. Saves use the signed-in user's `SPHttpClient` against the configured list site. The service rechecks management permissions before each write and leaves final list/item authorization, versioning, validation and configured workflows to SharePoint. Only trusted access administrators should be able to edit the assignment list.

Existing records update only `OfficeGroupRecordId`. New records set `UserId`, `OfficeGroupRecordId` and `Title`. Directory search reuses the app's existing read-only Graph service; SharePoint `ensureuser` resolves the chosen account to the list site's numeric user ID without granting site or group permissions. Person-field selection-group restrictions are checked. The editor validates field types, read-only flags, the lookup's target list and required-group constraints. Unsupported/custom configurations remain available through **Open assignments list**.

Updates send the item's exact ETag with `IF-MATCH` and `MERGE`, following [Microsoft's SharePoint REST concurrency contract](https://learn.microsoft.com/en-us/sharepoint/dev/sp-add-ins/working-with-lists-and-list-items-with-rest). A 412 conflict never triggers a wildcard overwrite.

List titles and site location still come from existing web-part settings. The fallback list URL comes from the configured list, so renamed list URLs do not require guessed paths. The administration view does not create lists, modify permissions, deploy a flow, delete assignment records, or automatically consolidate duplicate records.

For schema and flow details, see [list setup](SharePoint-Lists-Setup.md) and the [flow contract](PowerAutomate-Flow-Spec.md).
