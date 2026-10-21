# Admin-to-group assignments

Use **Access assignments** to see which groups each admin can manage, search by admin or group, and open the existing SharePoint assignment form. Membership administration remains on the **Membership** view.

The overview is available to users with **Manage Lists** and **Edit Items** on the configured `Group Management Authorized Admins` list. Being assigned to manage a Microsoft 365 group does not grant these list permissions. Page edit mode does not grant access to the overview.

## Assign or remove access

1. Open **Access assignments** and search for the admin or group.
2. Choose **Edit assignment**, or **Add an admin assignment** for a new admin. The configured list's native form opens in another tab.
3. Set `User` to the individual admin. Select all their groups in `OfficeGroupRecord`, then save. Keep one record per admin where possible.
4. Return to the app and select **Refresh assignments**. The affected admin can reload their membership page to see the updated group list.

To remove access to a group, remove it from **every** assignment record for that admin. Multiple records combine their groups; removing one occurrence may leave another grant. The overview flags duplicate admin records and shows groups without an assigned admin. Missing users and unavailable group records need review in the native list.

The overview reads all available pages of the two configured lists. Search filters the assignment table. The groups-without-an-admin section remains a view of the full loaded inventory, regardless of search. The page's `visibleOffices` setting does not filter this administrative view.

## What the assignment means

| Configuration | Effect |
| --- | --- |
| `Managed Groups` row | Identifies a group the application can manage. Its existing `GroupId` and type determine supported actions. |
| `User` plus `OfficeGroupRecord` | Maps an individual admin to one or more managed groups. The existing single-value lookup format remains readable. |
| `visibleOffices` web-part property | Narrows the groups shown on a membership page. It does not grant access. |
| `IsTeamsConnected` | Indicates that the Microsoft 365 group is connected to Teams. The overview does not inventory private or shared channel permissions. |

The app submits Microsoft 365 membership requests to the existing request list. Power Automate must independently authorize the request using its system-created `AuthorId` and current assignment records before writing membership changes. SharePoint group changes use the signed-in user's SharePoint permissions. An assignment alone does not grant site access.

The current model treats member and owner actions as group-level authorization. It does not expand admin-team membership or distinguish member managers from owner managers. Those features would require a separately designed and tested server-side authorization change; hiding buttons would not enforce them.

## Existing data and workflow contracts

This upgrade adds no list, column, Graph scope or flow connection. The overview is read-only. SharePoint enforces every edit made through its native forms, using existing permissions and versioning settings. Only trusted access administrators should be able to edit the assignment list.

List titles and site location still come from existing web-part settings. Form URLs come from the configured list, so renamed list URLs do not require guessed paths. The administration view does not create lists, modify permissions, deploy a flow, or automatically consolidate duplicate records.

For schema and flow details, see [list setup](SharePoint-Lists-Setup.md) and the [flow contract](PowerAutomate-Flow-Spec.md).
