# 365 Account Management — SharePoint List Setup Spec

Build these **4 SharePoint lists by hand** (this tenant blocks PnP, so there is no provisioning
script — this spec is the source of truth). The web part reads/writes them at runtime via the
signed-in user; a Power Automate flow processes the O365 requests.

---

## ⚠️ Read this first — 5 rules that will save you a broken deployment

1. **Internal column names must match EXACTLY (and have no spaces).** SharePoint freezes a column's
   *internal* name from its display name **at creation time**, and turns any space into `_x0020_`
   (e.g. a column first named "Group Id" gets the internal name `Group_x0020_Id`). The app requests
   columns by the internal names below, so a mismatch makes the whole query fail or silently return
   nothing. **Create each column using the exact space-free name shown, then rename the display label
   afterward if you want a friendlier caption** — renaming the display does not change the internal name.

2. **Every column listed must exist.** The app `$select`s these columns by name; a missing column
   makes SharePoint reject the entire request. The only exceptions are the few marked **"Optional to
   create."**

3. **Person and Lookup columns are written/read by their `…Id` field.** Create the column with the
   base name (e.g. `RequestedBy`); the app automatically uses `RequestedById`. Don't create a column
   literally named `RequestedById`.

4. **Built-in columns already exist — do not re-create them:** `Title`, `Created By` (Author),
   `Modified`, `ID`.

5. **List titles must match** the values below, **or** be set per page in the web part's property pane
   → *Data source (lists)*. Defaults are shown.

**Legend — "Filled by":** `Admin` = you populate rows by hand · `App` = the web part writes it ·
`Flow` = the Power Automate flow writes it back · `System` = SharePoint maintains it.

---

## List 1 — `Managed Groups`  (property: `groupListTitle`)

One row per group/office the tool manages. **Read-only to the app — admins populate every row.**
All columns must exist **except `GroupType`**.

| Column (internal name) | Type | Filled by | Notes |
|---|---|---|---|
| `Title` *(built-in)* | Single line of text | Admin | Group display name. |
| `GroupId` | Single line of text | Admin | **Required.** Entra/O365 group **GUID** (production) or SharePoint site-group **integer** (dev/test). The app can't act on a row without it. |
| `Description` | Multiple lines of text | Admin | Shown in the UI. |
| `Mail` | Single line of text | Admin | Group email address. |
| `MailNickname` | Single line of text | Admin | Alias; used as a title fallback. |
| `Visibility` | Single line of text | Admin | e.g. Public / Private (free text). |
| `CreatedDateTime` | Date and Time | Admin | Group creation date (separate from built-in *Created*). |
| `IsTeamsConnected` | Yes/No | Admin | Teams-connected flag. |
| `SiteUrl` | Hyperlink | Admin | Optional. **Same-tenant** site the group's permissions live on; used for SharePoint-group actions. Must be a real site-collection URL (see note). Single line of text also works. |
| `SiteTitle` | Single line of text | Admin | Friendly site/group title; title fallback. |
| `GroupType` | Single line of text | Admin | **Optional to create** — the app retries without it. Flags un-manageable types (dynamic, distribution). |

> `Title`, `GroupId`, and `SiteTitle` are surfaced to the Authorized Admins lookup — keep them as text.
> For a SharePoint-group row, `SiteUrl` must be the **site-collection root that owns the group**
> (e.g. `https://…/sites/M-EX`), never a subsite/folder.

---

## List 2 — `Group Membership Requests`  (property: `requestListTitle`)

The queue **and** audit log for O365 group changes: the app writes a `Pending` row, the Power Automate
flow performs the Microsoft Graph change and writes the result back. All columns must exist **except
`RequestedBy`** (the app retries without it).

| Column (internal name) | Type | Filled by | Notes |
|---|---|---|---|
| `Title` *(built-in)* | Single line of text | App | Label, e.g. "Add Member: Jane Doe". |
| `Action` | **Choice** | App | Exact values: **`Add Member`**, **`Remove Member`**. |
| `GroupId` | Single line of text | App | Target O365 group **GUID** the flow acts on. |
| `GroupName` | Single line of text | App | Group display name. |
| `MemberId` | Single line of text | App | Target member's Entra object id (GUID). |
| `MemberDisplayName` | Single line of text | App | |
| `Status` | **Choice** | App + Flow | Exact values: **`Pending`**, **`Completed`**, **`Failed`**. Default = `Pending`. App writes `Pending`; the flow writes `Completed`/`Failed`. |
| `ResultMessage` | Multiple lines of text | Flow (+App) | Outcome / error text. |
| `RequestedOn` | Date and Time | App | ISO submit timestamp (include time). |
| `TargetUserPrincipalName` | Single line of text | App | Member UPN. |
| `TargetUserEmail` | Single line of text | App | Member email. |
| `CorrelationId` | Single line of text | App | GUID for app↔flow correlation / audit. |
| `OfficeGroupRecord` | **Lookup → `Managed Groups`** (show `Title`) | App | Links the request to its Managed Groups row. App writes `OfficeGroupRecordId`. |
| `Justification` | Multiple lines of text | App | Reason for the change. **Must exist** — it is *not* dropped on failure, so a missing column errors the submit. Backs the "Require a reason" toggle. |
| `MemberEntraId` | **Person or Group** | App | Target member as a site user (REST `MemberEntraIdId`). **Must exist.** |
| `RequestedBy` | **Person or Group** | App | Signed-in initiator (REST `RequestedById`). **Optional to create** — app retries without it. |
| `AuthorizationChecked` | Yes/No | Flow | Flow records whether it authorized the change. App reads it. |
| `AuthorizationResult` | Multiple lines of text | Flow | Flow's authorization outcome/message. App reads it. |
| `Created By` (`Author`) *(built-in)* | Person | System | The app filters recent requests by the creator — built-in, no action. |
| `Modified` *(built-in)* | Date and Time | System | The app shows when the flow last updated the row. |

> **Flow team:** locate the row the app created and write back `Status` (`Completed`/`Failed`),
> `ResultMessage`, and optionally `AuthorizationChecked` / `AuthorizationResult`. Point the flow's
> trigger at **this** list (mind the exact title).

---

## List 3 — `Group Management Authorized Admins`  (property: `authorizedAdminsListTitle`)

Authorization table: which users may manage which groups. **Read-only to the app — admins populate it.**
**One row per admin** — `OfficeGroupRecord` is a **multiple-value** lookup, so you select *all* the groups
that admin manages in a single row. (The app also still reads the old one-row-per-group layout, so you can
migrate at your own pace.)

| Column (internal name) | Type | Filled by | Notes |
|---|---|---|---|
| `User` | **Person or Group** (single) | Admin | The authorized admin. The app filters on this (REST `UserId`) — **index it.** |
| `OfficeGroupRecord` | **Lookup → `Managed Groups`** (show `Title`), **Allow multiple values ✓** | Admin | All groups this admin may manage — pick several in the one row. The app reads each and pulls `Title`/`GroupId`/`SiteTitle`. A single-value lookup still works (old layout). |

*(The built-in `ID` is used internally — no action.)*

> **Switching an existing list to multi-select:** SharePoint won't let you toggle "Allow multiple values" on a
> lookup that already has data. Delete the existing `OfficeGroupRecord` column and re-create it with the **same
> name** and **Allow multiple values** checked, then consolidate to one row per admin.
>
> ⚠️ **Power Automate flow:** the flow's server-side authorization re-check reads this same list. It must be
> updated to test whether the requested group is **in** the requester's multi-value `OfficeGroupRecord`
> collection (instead of matching a single lookup). Update the flow together with this list change, or O365
> membership requests will fail their authorization check.

---

## List 4 — `Group Site Permissions`  (property: `sitePermissionsListTitle`)  — optional list

Feeds the "Used on these sites" panel. **Read-only to the app — admins populate it.** The whole list is
optional (if absent the panel just doesn't render), but to make the panel work, create all four columns.
One row per (group → site).

| Column (internal name) | Type | Filled by | Notes |
|---|---|---|---|
| `GroupId` | Single line of text | Admin | The group's **GUID** (matches `Managed Groups.GroupId`). The app filters on this — **index it.** Plain text, **not** a lookup. |
| `SiteName` | Single line of text | Admin | Site label shown in the panel. Results are sorted by this — **index it.** |
| `SiteUrl` | Hyperlink | Admin | Clickable site link. (Single line of text also works — it renders as plain text.) |
| `Permission` | Single line of text | Admin | e.g. `Full Control`, `Edit`, `Read` (free text; the app doesn't validate values). |

*(The built-in `ID` is used internally — no action.)*

---

## After the lists exist — finish the wiring

- [ ] On each page, open the web part property pane → **Data source (lists)** and confirm the four
      titles match the lists you created (or type your custom titles).
- [ ] Point the **Power Automate flow** trigger at the `Group Membership Requests` list.
- [ ] Have a tenant admin approve the web part's **4 Microsoft Graph permissions** (SharePoint admin →
      Advanced → API access): `GroupMember.Read.All`, `User.ReadBasic.All`, `User.Read.All`,
      `ProfilePhoto.Read.All`.
- [ ] Index the columns noted above (`User` on Authorized Admins; `GroupId` and `SiteName` on Group
      Site Permissions) if those lists may exceed ~5,000 items.
