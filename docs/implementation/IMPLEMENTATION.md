# 365 Account Management — Implementation Guide

<!-- GENERATED FILE. Do not edit by hand: update implementation.json and re-run
     spfx-docs --app <app-path>   (from the spfx-implementation-docs repo)      -->
<!-- spfx-docs-meta {"slug":"user-management-portal-spfx","name":"365 Account Management","updated":"2026-07-09","lists":4,"dataverse":0,"flows":1,"webParts":1} -->

> One web part that lets designated office admins add and remove members of Microsoft 365 groups and classic SharePoint site groups from a SharePoint page. Four SharePoint lists drive it: Managed Groups (the offices), Group Management Authorized Admins (who may manage what), Group Membership Requests (the queue and audit trail), and an optional Group Site Permissions list. Microsoft 365 group changes are performed by one Power Automate flow running as a service account; SharePoint site-group changes are applied directly by the web part as the signed-in user.

| Field | Value |
| --- | --- |
| Solution | 365 Account Management |
| Solution version | 1.8.1.0 |
| Package | solution/365-account-management.sppkg |
| Environment | Production SharePoint (gov tenant — CDN-hosted assets required) |
| Last updated | 2026-07-09 |
| Owners | Executive Office dev team |
| Audience | Anyone — no SharePoint or Power Automate experience is assumed. |

**How to use this guide**

Work through the sections **in order** — each one builds on the previous. Every step is written
so that it can be followed without prior experience: exact menu names appear in **bold**, and
every value you must type or pick is shown in a table or in `code style`. Check off the
verification list in section 10 before calling the implementation done.

**Contents**

- 1. Solution overview
- 2. Prerequisites & access
- 3. Architecture at a glance
- 4. SharePoint lists
- 5. Dataverse tables
- 6. Power Automate flows
- 7. App deployment
- 8. Web part setup
- 9. Site configuration
- 10. Post-deployment verification
- 11. Rollback
- 12. Glossary

## 1. Solution overview

One web part that lets designated office admins add and remove members of Microsoft 365 groups and classic SharePoint site groups from a SharePoint page. Four SharePoint lists drive it: Managed Groups (the offices), Group Management Authorized Admins (who may manage what), Group Membership Requests (the queue and audit trail), and an optional Group Site Permissions list. Microsoft 365 group changes are performed by one Power Automate flow running as a service account; SharePoint site-group changes are applied directly by the web part as the signed-in user.

Everything keys off the GroupId column of the Managed Groups list. The web part reads Group Management Authorized Admins to find the signed-in user's grants, loads those Managed Groups rows, and shows one card per office. When an admin adds or removes a member, the web part branches on the GroupId FORMAT. GUID (a Microsoft 365 group): the web part — which holds only read-only Graph permissions — writes a Pending row to Group Membership Requests and polls it; the flow triggers on the new row, re-checks authorization server-side against the Authorized Admins list (matching the row's server-stamped Created By to the exact target group), applies the change through the Office 365 Groups connector as the service account, and writes Status = Completed or Failed plus a result message back onto the row. Integer (a classic SharePoint site group): NO flow at all — the web part calls SharePoint REST directly as the signed-in user (bounded by that user's own permissions on the target site) and logs a pre-completed audit row to the same request list so both paths share one history. The optional Group Site Permissions list feeds a curated “which sites does this group unlock” panel on each card.

## 2. Prerequisites & access

Confirm every row below **before starting**. Implementation stalls are almost always a missing permission.

| You need | Why | How to get it |
| --- | --- | --- |
| Site owner on the site that will host the four lists | You will create the Managed Groups, Group Management Authorized Admins, Group Membership Requests, and Group Site Permissions lists by hand | Site collection administrator |
| A SharePoint Administrator to approve the web part's Microsoft Graph permissions | The package requests four DELEGATED, READ-ONLY Graph permissions — GroupMember.Read.All, User.ReadBasic.All, User.Read.All, ProfilePhoto.Read.All — which must be approved on the SharePoint admin center's API access page before member lists, people search, and group photos work. The web part is read-only on Graph by design; all Microsoft 365 group WRITES go through the flow. | SharePoint admin team — the requests appear under SharePoint admin center → Advanced → API access after the package is deployed |
| A service account to own the flow and its connections | The Office 365 Groups connector changes membership as its connection identity, so the account must be an OWNER of every Microsoft 365 group the tool manages. It also needs Edit on Group Membership Requests plus Read on Managed Groups and Group Management Authorized Admins, and it must sign in unattended (exclude it from interactive MFA with a Conditional Access policy scoped to that one account — never tenant-wide). | IT service-account request — a licensed Entra user account (not a shared mailbox), password vaulted, owned by a team rather than a person |
| App Catalog administrator + CDN team contact (deployment section) | Package upload is manual; production assets are CDN-hosted | SharePoint admin team / workspace deployment standard |

## 3. Architecture at a glance

| Component | Kind | What it does |
| --- | --- | --- |
| 365 Account Management web part | SPFx web part | Office cards with member lists, tenant-wide people search, and Add/Remove — files flow requests for Microsoft 365 groups, applies SharePoint-group changes directly |
| Managed Groups list | SharePoint list | One row per office (group) the tool can manage; the GroupId format (GUID vs integer) selects the change path |
| Group Management Authorized Admins list | SharePoint list | Grants: which person may manage which Managed Groups row(s); read by the web part for display and re-checked server-side by the flow |
| Group Membership Requests list | SharePoint list | Queue and audit trail: the web part files Pending rows, the flow writes the outcome back, the web part polls the row |
| Group Site Permissions list (optional) | SharePoint list | Curated “sites this group is used on” rows shown on each office card; the panel hides itself when the list is absent |
| Process group membership requests flow | Cloud flow (Power Automate) | Authorizes and applies Microsoft 365 group membership changes as the service account and writes the result back |
| Service account | Entra user account | Owns the flow and its two connections; must be an owner of every managed Microsoft 365 group |

## 4. SharePoint lists

This solution uses **4 lists**. Lists in this tenant are created **by hand in the browser** — no scripts. Each subsection below is self-contained: create the list, add its columns in order, then verify.

### 4.1 List: “Managed Groups”

**Purpose.** One row per office (group) the tool manages. The web part shows a card per row the signed-in admin is authorized for; the GroupId format decides how changes are applied (GUID = Microsoft 365 group via the flow, integer = classic SharePoint site group via direct REST).
**Where.** the site the web part's “List site URL” property points at (blank = the site where the page lives)

#### 4.1.1 Create the list

1. Open the site (the site the web part's “List site URL” property points at (blank = the site where the page lives)) in the browser.
2. Click the **gear icon** (top right) → **Site contents**.
3. **+ New** → **List** → **Blank list**.
4. Name: `Managed Groups` → **Create**.

#### 4.1.2 Columns

| # | Display name | Internal name | Type | Required | Indexed | Details |
| --- | --- | --- | --- | --- | --- | --- |
| 1 | Title | `Title` | Built-in | Yes | No | Built-in Title — the office name shown on the card. If blank the app falls back to SiteTitle, then MailNickname. |
| 2 | GroupId | `GroupId` | Single line of text | Yes | No | THE key column. For a Microsoft 365 group: the Entra group object id GUID (Entra admin center → Groups → the group → Overview → Object Id). For a classic SharePoint site group: the integer group id (Site settings → People and groups — the ID in the page URL). The format alone picks the path: GUID rows go through the flow; integer rows are changed directly by the web part. |
| 3 | Description | `Description` | Multiple lines of text | No | No | Shown on the office card. |
| 4 | Mail | `Mail` | Single line of text | No | No | The group's email address — display only. |
| 5 | MailNickname | `MailNickname` | Single line of text | No | No | The group's alias; also a card-title fallback. |
| 6 | Visibility | `Visibility` | Single line of text | No | No | Private or Public — display only. |
| 7 | CreatedDateTime | `CreatedDateTime` | Date and time | No | No | When the group was created — display only. |
| 8 | IsTeamsConnected | `IsTeamsConnected` | Yes/No | No | No | Yes shows the Teams badge on the card. |
| 9 | SiteUrl | `SiteUrl` | Single line of text | No | No | Absolute URL of the group's SharePoint site. For integer (SharePoint-group) rows this is the site that HOLDS the site group — the app sends its REST calls there and only honors same-tenant URLs. |
| 10 | SiteTitle | `SiteTitle` | Single line of text | No | No | The site's display name; card-title fallback. |
| 11 | GroupType | `GroupType` | Single line of text | No | No | OPTIONAL flag for group kinds this tool cannot manage. Leave blank for normal groups. Recognized values (case/spacing-insensitive): Dynamic, MailEnabledSecurity, Distribution, RoleAssignable, OnPrem — each disables the card with an explanatory message. |

#### 4.1.3 Create each column

> **Why the two-step naming below matters:** SharePoint permanently locks a column’s *internal*
> name to the first name it is given. The app reads columns by internal name, so create each
> column named exactly as shown in step 1, and only rename it to the display name afterwards.

**Column step 1 — `GroupId` (Single line of text)**

1. In the list, click **+ Add column** → **Text** → **Next**.
2. Name: type exactly `GroupId` (this locks the internal name).
3. Expand **More options** → turn **Require that this column contains information** **On**.
4. **Save**.

> **Note.** The flow refuses non-GUID ids, and the app disables the card up front when the value is blank or malformed.

**Column step 2 — `Description` (Multiple lines of text)**

1. In the list, click **+ Add column** → **Multiple lines of text** → **Next**.
2. Name: type exactly `Description` (this locks the internal name).
3. **Save**.

> **Note.** Multiple lines of text in PLAIN TEXT mode (the app renders it as-is). Type inferred from display-only use — a single-line Text column also works.

**Column step 3 — `Mail` (Single line of text)**

1. In the list, click **+ Add column** → **Text** → **Next**.
2. Name: type exactly `Mail` (this locks the internal name).
3. **Save**.

**Column step 4 — `MailNickname` (Single line of text)**

1. In the list, click **+ Add column** → **Text** → **Next**.
2. Name: type exactly `MailNickname` (this locks the internal name).
3. **Save**.

**Column step 5 — `Visibility` (Single line of text)**

1. In the list, click **+ Add column** → **Text** → **Next**.
2. Name: type exactly `Visibility` (this locks the internal name).
3. **Save**.

**Column step 6 — `CreatedDateTime` (Date and time)**

1. In the list, click **+ Add column** → **Date and time** → **Next**.
2. Name: type exactly `CreatedDateTime` (this locks the internal name).
3. **Save**.

**Column step 7 — `IsTeamsConnected` (Yes/No)**

1. In the list, click **+ Add column** → **Yes/No** → **Next**.
2. Name: type exactly `IsTeamsConnected` (this locks the internal name).
3. **Save**.

**Column step 8 — `SiteUrl` (Single line of text)**

1. In the list, click **+ Add column** → **Text** → **Next**.
2. Name: type exactly `SiteUrl` (this locks the internal name).
3. **Save**.

> **Note.** A Hyperlink column also works (the app reads either shape); plain Text keeps hand-entry simple.

**Column step 9 — `SiteTitle` (Single line of text)**

1. In the list, click **+ Add column** → **Text** → **Next**.
2. Name: type exactly `SiteTitle` (this locks the internal name).
3. **Save**.

**Column step 10 — `GroupType` (Single line of text)**

1. In the list, click **+ Add column** → **Text** → **Next**.
2. Name: type exactly `GroupType` (this locks the internal name).
3. **Save**.

> **Note.** Create the column even though the app treats it as optional: when it is missing, the app's first query fails (HTTP 500) and it silently retries without the column on every load. The Office 365 Groups connector genuinely cannot manage those group kinds, so flagging them here prevents doomed requests.

#### 4.1.5 Permissions

Everyone who uses the web part needs Read. Restrict editing to the owning team — this list is trusted, privileged configuration (its SiteUrl values are where SharePoint-group REST calls are sent). The flow's service account needs Read.

#### 4.1.6 Starter data

One row per office. Fill Title, GroupId, and SiteUrl/SiteTitle at minimum. For every GUID row, also make the service account an OWNER of that Microsoft 365 group — the flow's connector actions fail otherwise.

> **Note.** Column names here double as display AND internal names — create each column with exactly this name and do not rename.

### 4.2 List: “Group Management Authorized Admins”

**Purpose.** The authorization table — one row grants one person the right to manage one or more Managed Groups rows. The web part shows an admin only the offices granted here, and the flow re-checks this same table server-side before changing anything.
**Where.** same site as Managed Groups

#### 4.2.1 Create the list

1. Open the site (same site as Managed Groups) in the browser.
2. Click the **gear icon** (top right) → **Site contents**.
3. **+ New** → **List** → **Blank list**.
4. Name: `Group Management Authorized Admins` → **Create**.

#### 4.2.2 Columns

| # | Display name | Internal name | Type | Required | Indexed | Details |
| --- | --- | --- | --- | --- | --- | --- |
| 1 | Title | `Title` | Built-in | Yes | No | Not read by the app — type the admin's name (or any label) to keep the list readable. |
| 2 | User | `User` | Person | Yes | Yes | The admin being granted rights. The app finds the signed-in user's grants with a numeric-id filter on this Person column (UserId eq <current user id>). |
| 3 | OfficeGroupRecord | `OfficeGroupRecord` | Lookup | Yes | No | Lookup → “Managed Groups” (Title). Which Managed Groups row(s) this person may manage. Allow multiple values if you want one row to grant several offices — the app handles single- and multi-value lookups. |

#### 4.2.3 Create each column

> **Why the two-step naming below matters:** SharePoint permanently locks a column’s *internal*
> name to the first name it is given. The app reads columns by internal name, so create each
> column named exactly as shown in step 1, and only rename it to the display name afterwards.

**Column step 1 — `User` (Person)**

1. In the list, click **+ Add column** → **Person** → **Next**.
2. Name: type exactly `User` (this locks the internal name).
3. Expand **More options** → turn **Require that this column contains information** **On**.
4. **Save**.

**Column step 2 — `OfficeGroupRecord` (Lookup)**

1. In the list, click **+ Add column** → **Lookup** → **Next**.
2. Name: type exactly `OfficeGroupRecord` (this locks the internal name).
3. Choose list **Managed Groups**, column **Title**.
4. Expand **More options** → turn **Require that this column contains information** **On**.
5. **Save**.

> **Note.** IMPORTANT: when creating the lookup, ALSO tick GroupId under “Add a column to show each of these additional fields”. The flow's authorization query reads OfficeGroupRecord/GroupId and silently authorizes nobody without that projection. Other projections (e.g. SiteTitle) are optional and only aid readability.

**Column step 3 — add the indexes**

1. Gear icon → **List settings** (if you don’t see it: gear → **Site contents** → “…” next to the list → **Settings**).
2. Under **Columns**, click **Indexed columns** → **Create a new index**.
3. Pick `User` as the primary column → **Create**. Repeat for each: `User`.

#### 4.2.5 Permissions

Everyone who uses the web part needs Read (the app looks up the signed-in user's own grants). Restrict editing to the owning team — this list IS the permission model. The flow's service account needs Read.

#### 4.2.6 Starter data

One row per admin: pick the person in User and the office(s) in OfficeGroupRecord. An admin with no row sees an empty web part — that is the intended “not authorized” state.

> **Note.** Column names here double as display AND internal names — create each column with exactly this name and do not rename.

### 4.3 List: “Group Membership Requests”

**Purpose.** The flow's work queue and the system of record for who asked for what. The web part files a Pending row per Microsoft 365 group change and polls it until the flow writes Completed or Failed back. Direct SharePoint-group changes are logged here too, pre-marked Completed so the flow skips them.
**Where.** same site as Managed Groups

#### 4.3.1 Create the list

1. Open the site (same site as Managed Groups) in the browser.
2. Click the **gear icon** (top right) → **Site contents**.
3. **+ New** → **List** → **Blank list**.
4. Name: `Group Membership Requests` → **Create**.

#### 4.3.2 Columns

| # | Display name | Internal name | Type | Required | Indexed | Details |
| --- | --- | --- | --- | --- | --- | --- |
| 1 | Title | `Title` | Built-in | Yes | No | Written by the app as “<Action>: <member name>”, e.g. “Add Member: Jane Smith”. |
| 2 | Action | `Action` | Single line of text | No | No | Exactly “Add Member” or “Remove Member” — the flow's Switch matches these case-sensitively. |
| 3 | GroupId | `GroupId` | Single line of text | No | No | The target group's id copied from Managed Groups. The flow only processes GUID values; integer (SharePoint-group) rows arrive pre-completed. |
| 4 | GroupName | `GroupName` | Single line of text | No | No | The group's display name — used in result messages. |
| 5 | MemberId | `MemberId` | Single line of text | No | No | Entra object id (GUID) of the target user. Audit only on the connector path; it is what an app-registration variant of the flow would use. |
| 6 | MemberDisplayName | `MemberDisplayName` | Single line of text | No | No | Display name of the target user — used in result messages. |
| 7 | Status | `Status` | Single line of text | No | No | Lifecycle. The app writes Pending (or Completed for direct SharePoint-group changes); the flow ends every row at exactly Completed or Failed. |
| 8 | RequestedOn | `RequestedOn` | Date and time | No | No | When the admin submitted the request (the app stamps the current date and time). |
| 9 | TargetUserPrincipalName | `TargetUserPrincipalName` | Single line of text | No | No | UPN of the user being added or removed — the value the flow feeds to the Office 365 Groups connector. |
| 10 | TargetUserEmail | `TargetUserEmail` | Single line of text | No | No | The target user's email address (audit/messages). |
| 11 | CorrelationId | `CorrelationId` | Single line of text | No | No | A GUID the app generates per request — traces a click in the web part to a flow run. |
| 12 | OfficeGroupRecord | `OfficeGroupRecord` | Lookup | No | No | Lookup → “Managed Groups” (Title). The Managed Groups row the request came from (the app sets it via OfficeGroupRecordId). |
| 13 | Justification | `Justification` | Multiple lines of text | No | No | The reason the admin typed. Written only when provided; the web part property “Require a reason for each change” makes it mandatory in the UI. |
| 14 | MemberEntraId | `MemberEntraId` | Person | No | No | Person column for the target user — the app resolves the UPN via ensureuser and sets MemberEntraIdId. Best-effort: left empty when the account cannot be resolved. |
| 15 | RequestedBy | `RequestedBy` | Person | No | No | Person column for the requesting admin. |
| 16 | ResultMessage | `ResultMessage` | Multiple lines of text | No | No | The human sentence shown to the user. Written by the FLOW on the connector path; on the direct path the app writes “Applied directly to the SharePoint group.” |
| 17 | AuthorizationChecked | `AuthorizationChecked` | Yes/No | No | No | Written by the FLOW: Yes once the server-side authorization re-check has run. |
| 18 | AuthorizationResult | `AuthorizationResult` | Multiple lines of text | No | No | Written by the FLOW: the re-check outcome, e.g. “Authorized for <group>.” or “Not authorized to manage this group.” |

#### 4.3.3 Create each column

> **Why the two-step naming below matters:** SharePoint permanently locks a column’s *internal*
> name to the first name it is given. The app reads columns by internal name, so create each
> column named exactly as shown in step 1, and only rename it to the display name afterwards.

**Column step 1 — `Action` (Single line of text)**

1. In the list, click **+ Add column** → **Text** → **Next**.
2. Name: type exactly `Action` (this locks the internal name).
3. **Save**.

> **Note.** Single line of text. The app only ever writes those two values; a Choice column is unnecessary and risks a mismatch.

**Column step 2 — `GroupId` (Single line of text)**

1. In the list, click **+ Add column** → **Text** → **Next**.
2. Name: type exactly `GroupId` (this locks the internal name).
3. **Save**.

**Column step 3 — `GroupName` (Single line of text)**

1. In the list, click **+ Add column** → **Text** → **Next**.
2. Name: type exactly `GroupName` (this locks the internal name).
3. **Save**.

**Column step 4 — `MemberId` (Single line of text)**

1. In the list, click **+ Add column** → **Text** → **Next**.
2. Name: type exactly `MemberId` (this locks the internal name).
3. **Save**.

**Column step 5 — `MemberDisplayName` (Single line of text)**

1. In the list, click **+ Add column** → **Text** → **Next**.
2. Name: type exactly `MemberDisplayName` (this locks the internal name).
3. **Save**.

**Column step 6 — `Status` (Single line of text)**

1. In the list, click **+ Add column** → **Text** → **Next**.
2. Name: type exactly `Status` (this locks the internal name).
3. **Save**.

> **Note.** MUST be single line of text, NOT a Choice column — the app posts plain strings. The terminal values are a case-sensitive contract: the web part treats only the exact strings “Completed” and “Failed” as terminal; anything else (Complete, Success, InProgress) makes it poll to its timeout and report a false failure.

**Column step 7 — `RequestedOn` (Date and time)**

1. In the list, click **+ Add column** → **Date and time** → **Next**.
2. Name: type exactly `RequestedOn` (this locks the internal name).
3. **Save**.

**Column step 8 — `TargetUserPrincipalName` (Single line of text)**

1. In the list, click **+ Add column** → **Text** → **Next**.
2. Name: type exactly `TargetUserPrincipalName` (this locks the internal name).
3. **Save**.

**Column step 9 — `TargetUserEmail` (Single line of text)**

1. In the list, click **+ Add column** → **Text** → **Next**.
2. Name: type exactly `TargetUserEmail` (this locks the internal name).
3. **Save**.

**Column step 10 — `CorrelationId` (Single line of text)**

1. In the list, click **+ Add column** → **Text** → **Next**.
2. Name: type exactly `CorrelationId` (this locks the internal name).
3. **Save**.

**Column step 11 — `OfficeGroupRecord` (Lookup)**

1. In the list, click **+ Add column** → **Lookup** → **Next**.
2. Name: type exactly `OfficeGroupRecord` (this locks the internal name).
3. Choose list **Managed Groups**, column **Title**.
4. **Save**.

> **Note.** Client-supplied convenience — the flow deliberately authorizes on GroupId plus the server-stamped Created By, never on this lookup.

**Column step 12 — `Justification` (Multiple lines of text)**

1. In the list, click **+ Add column** → **Multiple lines of text** → **Next**.
2. Name: type exactly `Justification` (this locks the internal name).
3. **Save**.

> **Note.** Multiple lines of text, PLAIN TEXT mode. Do not skip this column: when a reason is required, the app deliberately fails the submit rather than silently dropping the reason if the column is missing.

**Column step 13 — `MemberEntraId` (Person)**

1. In the list, click **+ Add column** → **Person** → **Next**.
2. Name: type exactly `MemberEntraId` (this locks the internal name).
3. **Save**.

**Column step 14 — `RequestedBy` (Person)**

1. In the list, click **+ Add column** → **Person** → **Next**.
2. Name: type exactly `RequestedBy` (this locks the internal name).
3. **Save**.

> **Note.** Optional metadata: if this column is missing the app retries the write without it. The authoritative requester identity is always the built-in Created By.

**Column step 15 — `ResultMessage` (Multiple lines of text)**

1. In the list, click **+ Add column** → **Multiple lines of text** → **Next**.
2. Name: type exactly `ResultMessage` (this locks the internal name).
3. **Save**.

> **Note.** Plain text mode.

**Column step 16 — `AuthorizationChecked` (Yes/No)**

1. In the list, click **+ Add column** → **Yes/No** → **Next**.
2. Name: type exactly `AuthorizationChecked` (this locks the internal name).
3. **Save**.

**Column step 17 — `AuthorizationResult` (Multiple lines of text)**

1. In the list, click **+ Add column** → **Multiple lines of text** → **Next**.
2. Name: type exactly `AuthorizationResult` (this locks the internal name).
3. **Save**.

> **Note.** Plain text mode.

#### 4.3.5 Permissions

Requesting admins need Add items plus Read — the app lists “my recent requests” filtered to rows they created, so item-level “Read items that were created by the user” hardening is compatible. The flow's service account needs Edit. Ideally only the service account edits rows after creation: Status and the Authorization columns are its write-back.

> **Note.** Who writes what — the APP writes Title, Action, GroupId, GroupName, MemberId, MemberDisplayName, Status (Pending, or Completed on the direct path), RequestedOn, TargetUserPrincipalName, TargetUserEmail, CorrelationId, OfficeGroupRecord, Justification, MemberEntraId, and RequestedBy at submit. The FLOW writes back exactly four: Status (Completed/Failed), ResultMessage, AuthorizationChecked, AuthorizationResult. Column names double as display AND internal names — create each column with exactly this name.

### 4.4 List: “Group Site Permissions”

**Purpose.** OPTIONAL. A curated, hand-maintained map of which SharePoint sites each group grants access to (and at what level), shown as a panel on the office card. The web part hides the panel when the list is missing or unreadable — nothing else depends on it.
**Where.** same site as Managed Groups

#### 4.4.1 Create the list

1. Open the site (same site as Managed Groups) in the browser.
2. Click the **gear icon** (top right) → **Site contents**.
3. **+ New** → **List** → **Blank list**.
4. Name: `Group Site Permissions` → **Create**.

#### 4.4.2 Columns

| # | Display name | Internal name | Type | Required | Indexed | Details |
| --- | --- | --- | --- | --- | --- | --- |
| 1 | Title | `Title` | Built-in | Yes | No | Not read by the app — type any label (e.g. “<group> on <site>”). |
| 2 | SiteName | `SiteName` | Single line of text | No | No | Friendly site name shown in the panel. |
| 3 | SiteUrl | `SiteUrl` | Hyperlink | No | No | Hyperlink column pointing at the site. The app reads the link's URL, and uses the link's description as a fallback name when SiteName is blank. |
| 4 | Permission | `Permission` | Single line of text | No | No | The permission level shown, e.g. Read, Contribute, Full Control — free text. |
| 5 | GroupId | `GroupId` | Single line of text | No | Yes | Matches the group's GroupId in Managed Groups — the app filters this list on it for each card. |

#### 4.4.3 Create each column

> **Why the two-step naming below matters:** SharePoint permanently locks a column’s *internal*
> name to the first name it is given. The app reads columns by internal name, so create each
> column named exactly as shown in step 1, and only rename it to the display name afterwards.

**Column step 1 — `SiteName` (Single line of text)**

1. In the list, click **+ Add column** → **Text** → **Next**.
2. Name: type exactly `SiteName` (this locks the internal name).
3. **Save**.

**Column step 2 — `SiteUrl` (Hyperlink)**

1. In the list, click **+ Add column** → **Hyperlink** → **Next**.
2. Name: type exactly `SiteUrl` (this locks the internal name).
3. **Save**.

> **Note.** Create as “Hyperlink or Picture” → Hyperlink. (The app also tolerates a plain Text column, but Hyperlink is the intended type.)

**Column step 3 — `Permission` (Single line of text)**

1. In the list, click **+ Add column** → **Text** → **Next**.
2. Name: type exactly `Permission` (this locks the internal name).
3. **Save**.

**Column step 4 — `GroupId` (Single line of text)**

1. In the list, click **+ Add column** → **Text** → **Next**.
2. Name: type exactly `GroupId` (this locks the internal name).
3. **Save**.

**Column step 5 — add the indexes**

1. Gear icon → **List settings** (if you don’t see it: gear → **Site contents** → “…” next to the list → **Settings**).
2. Under **Columns**, click **Indexed columns** → **Create a new index**.
3. Pick `GroupId` as the primary column → **Create**. Repeat for each: `GroupId`.

#### 4.4.5 Permissions

Web part users need Read; restrict editing to the owning team (it is trusted display data).

#### 4.4.6 Starter data

One row per (group, site) pair you want shown. Rows are display-only — they grant nothing; they document grants made elsewhere.

> **Note.** Skip this list entirely if you do not want the panel — the web part degrades gracefully. Column names double as display AND internal names.

## 5. Dataverse tables

_None for this solution._

## 6. Power Automate flows

This solution depends on **1 cloud flow**. Build them in [make.powerautomate.com](https://make.powerautomate.com) after the lists exist (flows reference the lists). Sign in as the intended flow owner before starting.

### 6.1 Flow: “Process group membership requests”

**Purpose.** Watches the Group Membership Requests list. For each new Pending row targeting a Microsoft 365 (GUID) group it re-checks that the requester is authorized for that exact group, applies the add/remove through the Office 365 Groups connector as the service account (idempotently), and writes Completed or Failed plus a human message back onto the row the web part is polling. SharePoint-group changes never reach this flow.

| Setting | Value |
| --- | --- |
| Flow type | Automated cloud flow |
| Environment | Default |
| Solution | — |
| Recommended owner | Service account (licensed Entra user; OWNER of every managed Microsoft 365 group; Edit on Group Membership Requests, Read on Managed Groups and Group Management Authorized Admins) |
| Trigger | When an item is created (SharePoint) |

#### 6.1.1 Connections needed

Create or confirm these before building — the designer prompts for them the first time each connector is used.

| Connector | Sign-in account | Notes |
| --- | --- | --- |
| SharePoint | The flow owner (service account) | Used by the trigger, the authorization re-check, and the write-back. |
| Office 365 Groups | The flow owner (service account) | Standard (non-premium) connector. Membership actions succeed because the connection identity is an OWNER of the target group. It cannot manage mail-enabled security groups, distribution lists, or role-assignable groups — flag those rows with GroupType in Managed Groups instead. |

#### 6.1.2 Create the flow and trigger

1. Open [make.powerautomate.com](https://make.powerautomate.com); confirm the **environment** (top right).
2. **+ Create** (left nav).
3. Choose **Automated cloud flow**.
4. Flow name: `Process group membership requests`.
5. Trigger: search for and select **When an item is created (SharePoint)** → **Create**.

Configure the trigger:

| Trigger field | Set to |
| --- | --- |
| Site Address | the site that hosts Group Membership Requests |
| List Name | Group Membership Requests |

> **Note.** Use the plain “created” trigger, NOT “created or modified” — the flow updates the same row and must not re-trigger itself. Trigger settings (… → Settings): Concurrency Control On, Degree of Parallelism 1. If a page's “Request list title” property points at a different list, that list needs its own copy of this flow triggering on it.

#### 6.1.3 Build the steps

Add each action in order with **+ New step**. “Action” below is what to type into the action search box.

**Step 1 — Initialize the four result variables**

1. Action: **Variable → Initialize variable (four times)**.

| Field | Set to |
| --- | --- |
| varTerminalStatus | Type String, no initial value |
| varResultMessage | Type String, no initial value |
| varAuthorizationResult | Type String, no initial value |
| varAuthorizationChecked | Type Boolean, initial value false |

> **Note.** The rest of the flow only fills these variables; a single Update item at the end (step 11) writes them to the row — every path funnels through that one write-back.

**Step 2 — Ignore rows that are not Pending**

1. Action: **Control → Condition**.

| Field | Set to |
| --- | --- |
| Condition | Status (dynamic content) is equal to Pending |
| If no | Control → Terminate, status Succeeded — the row is not work for this flow (e.g. the app's pre-completed SharePoint-group audit rows). Write nothing. |

> **Note.** Everything from step 3 on goes inside the “If yes” branch. This is the only mid-flow Terminate.

**Step 3 — Refuse SharePoint-group ids (insurance guard)**

1. Action: **Control → Condition**.

| Field | Set to |
| --- | --- |
| Condition | GroupId (dynamic content) contains - |
| If no | Set variable varTerminalStatus = Failed, then Set variable varResultMessage = “SharePoint groups are handled by the widget, not this flow.” Do NOT call the connector; execution falls through to step 11. |

> **Note.** Entra group GUIDs always contain hyphens; classic SharePoint site-group ids are pure integers and are changed directly by the web part, never by this flow.

**Step 4 — Open a Try scope**

1. Action: **Control → Scope (rename it Try)**.

> **Note.** Steps 5–9 all go INSIDE this scope so any unexpected error lands in the Catch scope (step 10) instead of stranding the row at Pending.

**Step 5 — Re-check authorization server-side**

1. Action: **SharePoint → Send an HTTP request to SharePoint (rename it HTTP_GetGrant)**.

| Field | Set to |
| --- | --- |
| Site Address | the site that hosts the lists |
| Method | GET |
| Uri | _api/web/lists/getbytitle('Group Management Authorized Admins')/items?$select=Id,UserId,OfficeGroupRecord/Id,OfficeGroupRecord/GroupId&$expand=OfficeGroupRecord&$filter=UserId eq @{triggerOutputs()?['body/Author/Id']} and OfficeGroupRecord/GroupId eq '@{triggerBody()?['GroupId']}' |
| Headers | Accept: application/json;odata=nometadata |

> **Note.** This authorizes the row's server-stamped Created By (never a client-supplied field) against the EXACT group the request targets — a forged OfficeGroupRecord lookup cannot help. A plain Get items action cannot traverse the lookup, hence the REST call. It only returns rows when the Authorized Admins lookup projects GroupId (see that list's column notes).

**Step 6 — Record that the check ran, then decide**

1. Action: **Variable → Set variable, then Control → Condition**.

| Field | Set to |
| --- | --- |
| Set variable | varAuthorizationChecked = true |
| Condition (expression) | length(body('HTTP_GetGrant')?['value']) is greater than 0 |
| If no (not authorized) | varTerminalStatus = Failed; varAuthorizationResult = “Not authorized to manage this group.”; varResultMessage = “You are not authorized to manage membership for @{triggerBody()?['GroupName']}.” Do NOT call the connector; fall through to step 11. |
| If yes | varAuthorizationResult = “Authorized for @{triggerBody()?['GroupName']}.” — continue to step 7. |

> **Note.** Fail-closed: if HTTP_GetGrant itself errors (throttling, list renamed), the Try scope fails into Catch and the row is marked Failed with no change made — the intended safe default.

**Step 7 — Switch on the requested action**

1. Action: **Control → Switch**.

| Field | Set to |
| --- | --- |
| On | Action (dynamic content) |
| Case “Add Member” | step 8 |
| Case “Remove Member” | step 9 |
| Default | varTerminalStatus = Failed; varResultMessage = “Unknown action: @{triggerOutputs()?['body/Action']}.” |

> **Note.** The case labels are case-sensitive and must read exactly Add Member and Remove Member.

**Step 8 — Case Add Member — idempotent add**

1. Action: **Office 365 Groups → List group members, then Control → Condition, then Office 365 Groups → Add member to group**.

| Field | Set to |
| --- | --- |
| List group members | Group Id = GroupId (dynamic content) |
| Condition “already a member?” | length of the member list filtered where userPrincipalName equals TargetUserPrincipalName (compare lower-cased) is greater than 0 |
| If yes (already a member) | varTerminalStatus = Completed; varResultMessage = “@{triggerBody()?['MemberDisplayName']} is already a member of @{triggerBody()?['GroupName']}.” |
| If no — Add member to group | Group Id = GroupId; User Principal Name = TargetUserPrincipalName |
| After the add (run after: is successful) | varTerminalStatus = Completed; varResultMessage = “… was added to @{triggerBody()?['GroupName']}.” |
| Parallel branch (run after: has failed / has timed out) | varTerminalStatus = Failed; varResultMessage = “Could not add …: ” + coalesce(body('Add_member_to_group')?['error']?['message'], 'the request failed or timed out') |

> **Note.** Idempotency comes from the member list — never decide a no-op by parsing a localized error string. Any real failure stays Failed with the real message; never mask an error as success.

**Step 9 — Case Remove Member — idempotent remove**

1. Action: **Office 365 Groups → List group members, then Control → Condition, then Office 365 Groups → Remove member from group**.

| Field | Set to |
| --- | --- |
| List group members | Group Id = GroupId (dynamic content) |
| Condition “is a member?” | same length(filter(...)) test as step 8 |
| If no (not a member) | varTerminalStatus = Completed; varResultMessage = “@{triggerBody()?['MemberDisplayName']} was not a member of @{triggerBody()?['GroupName']}.” |
| If yes — Remove member from group | Group Id = GroupId; User Principal Name = TargetUserPrincipalName |
| After the remove (run after: is successful) | varTerminalStatus = Completed; varResultMessage = “… was removed from @{triggerBody()?['GroupName']}.” |
| Parallel branch (run after: has failed / has timed out) | varTerminalStatus = Failed; varResultMessage = “Could not remove …: ” + the connector's error message (same coalesce pattern as step 8) |

**Step 10 — Catch scope — any unexpected error**

1. Action: **Control → Scope (rename it Catch)**.

| Field | Set to |
| --- | --- |
| Configure run after | has failed, has timed out (of the Try scope) |
| Inside | Set variable varTerminalStatus = Failed; Set variable varResultMessage = “The request could not be completed due to an unexpected error.” |

> **Note.** Optional: surface the real error with first(filter(result('Try'), …status Failed…))?['outputs']?['body']?['error']?['message'] — do not use result('Try')[0], which is the first action, not the failed one.

**Step 11 — Write the result back — exactly once**

1. Action: **SharePoint → Update item**.

| Field | Set to |
| --- | --- |
| Site Address / List Name | the request list (same as the trigger) |
| Id | ID (dynamic content from the trigger) |
| Status | @{variables('varTerminalStatus')} |
| ResultMessage | @{variables('varResultMessage')} |
| AuthorizationChecked | @{variables('varAuthorizationChecked')} |
| AuthorizationResult | @{variables('varAuthorizationResult')} |
| Every other field | leave blank so Update item does not clear it |

> **Note.** Place this AFTER the Try and Catch scopes and set its Configure run after to is successful + has failed + is skipped so every path reaches it. Status must resolve to exactly Completed or Failed — never an interim value (InProgress, Complete, Success): the web part polls this row and only those two exact strings stop the poll. If the update returns 400, check the trigger's item-id token (ID) first — a failed write here strands the row at Pending.

**Step 12 — End the run**

1. Action: **Control → Terminate (status Succeeded)**.

> **Note.** The run succeeded at recording a business outcome — even when that outcome is Failed.

#### 6.1.4 Error handling & settings

Steps 5–9 live inside a Try scope; a Catch scope (Configure run after: Try has failed / has timed out) marks the row Failed with a generic message. The single Update item after both scopes runs on is successful / has failed / is skipped, so every branch — unauthorized, unknown action, guard, success, caught error — writes back. Invariant to verify before go-live: NO path may leave a row at Status = Pending; a stranded Pending row makes the web part poll to its timeout and report a false failure.

Trigger → Settings → Concurrency Control: On, Degree of Parallelism: 1 — serializes runs so the already-a-member check is reliable and avoids connector throttling. Keep end-to-end runtime well under the web part's “Status wait timeout” property (default 120 s; the created-trigger itself can add ~5–30 s of latency) — measure in UAT and raise the property if needed.

#### 6.1.5 Test the flow

1. Happy add: in the web part, add a member to a Microsoft 365 (GUID) group → the flow run succeeds; the request row flips Pending → Completed with “… was added …”; the web part shows success within seconds and the member appears in the list.
2. Idempotent add: submit the same add again → Completed with “… is already a member …” (decided from the member list, not from a connector error).
3. Happy remove, then idempotent remove of a non-member → Completed with “… was removed …” / “… was not a member …”.
4. Unauthorized: hand-create a Pending row as (or forge one for) an admin with no grant for that group → Failed, AuthorizationResult “Not authorized to manage this group.”, and the run history shows NO connector action ran.
5. SharePoint-group guard: hand-create a Pending row whose GroupId is an integer → Failed with “SharePoint groups are handled by the widget, not this flow.”, no connector call.
6. Unknown action: hand-create a Pending row with Action = Foo → Failed, “Unknown action: Foo”.
7. Contract regression (critical): temporarily make the write-back write Status = “Complete” (no d) on a happy add → the web part polls to its timeout and then shows an error, proving the case-sensitive Completed/Failed contract. Revert immediately.

> **Note.** Ownership drift is this model's number-one failure mode: if the service account is not an owner of a newly added group (or is removed from one), adds and removes for that group start failing — put “add the service account as owner” in the office-onboarding checklist and audit monthly. Entra's audit log attributes changes to the service account; the request list (Created By + CorrelationId) is the system of record for who asked. Watch for rows stuck at Pending (the trigger did not fire, or the flow died before write-back). If the tool later outgrows group-by-group ownership, swap the two connector actions for HTTP calls under an app registration with application permission GroupMember.ReadWrite.All — the trigger, authorization re-check, and write-back stay identical.

## 7. App deployment

| Fact | Value |
| --- | --- |
| Solution name | 365 Account Management |
| Solution id | ee656af9-b263-4758-b163-a56663bc6cbd |
| Solution version | 1.8.1.0 |
| Feature version | 1.8.1.0 |
| Assets bundled in package | true |
| CDN base path | https://irm.azureedge.us/M/user-management-portal-spfx/ |
| Deployment mode documented here | Production (CDN-hosted assets) |

**Production rule for this tenant:** SharePoint pages cannot load web part JavaScript from SharePoint itself (a Conditional Access policy intercepts those downloads). Production packages must load assets from the CDN. Deploying in the wrong order breaks every page that uses the web part — follow the sequence exactly.

1. Bump **all three versions** before building: `version` in `package.json`, and both the solution `version` and feature `version` in `config/package-solution.json`.
2. Confirm production flags: `includeClientSideAssets: false` in `config/package-solution.json`; `cdnBasePath` in `config/write-manifests.json` points at the CDN folder for this project.
3. Build: `gulp clean && gulp bundle --ship && gulp package-solution --ship`.
4. Zip everything in `release/assets/` (flat, original filenames, plus SHA-256 checksums) and send it to the CDN team for this project’s folder. **Never rename the files.**
5. **WAIT** for CDN confirmation, then open one bundle URL in a browser and confirm it returns JavaScript, not an HTML sign-in/policy page.
6. Only then upload `sharepoint/solution/*.sppkg` to the **App Catalog** → **Deploy**.
7. If the catalog ever shows two entries for this solution, delete both, empty **both** recycle bins on the catalog site, and re-upload one — duplicates serve stale manifests.

## 8. Web part setup

After deployment (section 7), add each web part to a page:

1. Open the target page → **Edit**.
2. Click **+** in the section where the web part belongs.
3. Search for the web part by the name shown below and select it.
4. Open its **property pane** (pencil icon) and set the properties, then **Republish** the page.

### 8.1 365 Account Management

Manage Microsoft 365 account and group membership requests

| Fact | Value |
| --- | --- |
| Toolbox name | 365 Account Management |
| Alias | AccountManagementWebPart |
| Component id | 762df7fa-0fd0-4019-b227-d184ab4abc67 |
| Version | * |

Default property values (change in the property pane as needed):

| Property | Default |
| --- | --- |
| `description` | `365 Account Management` |
| `introText` | _(blank)_ |
| `helpText` | _(blank)_ |
| `helpUrl` | _(blank)_ |
| `requestListTitle` | `Group Membership Requests` |
| `groupListTitle` | `Managed Groups` |
| `authorizedAdminsListTitle` | `Group Management Authorized Admins` |
| `sitePermissionsListTitle` | `Group Site Permissions` |
| `listSiteUrl` | _(blank)_ |
| `visibleOffices` | _(blank)_ |
| `pollTimeoutSeconds` | `120` |
| `startCollapsed` | `true` |
| `showGroupPhotos` | `true` |
| `requireJustification` | `false` |
| `verboseLogging` | `false` |

## 9. Site configuration

### 9.1 Approve the web part's Microsoft Graph permissions (once per tenant)

1. After the package is deployed (section 7), open the SharePoint admin center → Advanced → API access.
2. Under Pending requests, approve the four Microsoft Graph delegated permissions the package requests: GroupMember.Read.All, User.ReadBasic.All, User.Read.All, and ProfilePhoto.Read.All.
3. All four are read-only: they let the web part list group members and owners, search the directory, and show group photos. Membership WRITES are performed by the flow, not the web part.

> **Note.** Until these are approved, office cards load but member lists and people search fail with a Graph permissions error.

### 9.2 Add and wire the web part

1. Edit the target page, add the “365 Account Management” web part (section 8), and open its property pane.
2. Data source: the four list-title properties already default to Managed Groups, Group Management Authorized Admins, Group Membership Requests, and Group Site Permissions — change them only if you named the lists differently. Set “List site URL” to the site that hosts the lists when the page lives elsewhere in the tenant (blank = the current site).
3. Offices on this page: leave “Offices to show” blank to show every office the viewer is authorized for, or enter a comma-separated subset of office names / Group IDs (authorization is still enforced either way — listing an unauthorized office has no effect).
4. Behavior: set “Status wait timeout” (default 120 s) comfortably above the flow's observed end-to-end time; turn on “Require a reason for each change” if submissions must carry a justification; leave “Show Microsoft 365 group photos” on unless photos are unwanted.
5. Republish the page.

> **Note.** Pointing different pages at different request lists routes them to different flows — each such list needs its own copy of the flow triggering on it.

## 10. Post-deployment verification

Run every check. The implementation is complete only when all boxes are checked.

- [ ] Open the **Managed Groups** list on the site the web part's “List site URL” property points at (blank = the site where the page lives) — every column in its table (section 4) exists with the exact internal name and type.
- [ ] Open the **Group Management Authorized Admins** list on same site as Managed Groups — every column in its table (section 4) exists with the exact internal name and type.
- [ ] Open the **Group Membership Requests** list on same site as Managed Groups — every column in its table (section 4) exists with the exact internal name and type.
- [ ] Open the **Group Site Permissions** list on same site as Managed Groups — every column in its table (section 4) exists with the exact internal name and type.
- [ ] **Process group membership requests** exists, is turned **On**, and its run history shows a successful test run.
- [ ] The **365 Account Management** web part appears in the page toolbox and loads without errors on a real page.
- [ ] Open the page as an admin who has Group Management Authorized Admins rows → expected: One card per granted office; offices without a grant do not appear
- [ ] Open the page as a user with NO Authorized Admins row → expected: The web part shows its empty “no offices” state — no cards, no errors
- [ ] Search for a person in a card's add-member box → expected: Directory results appear (proves the Graph permissions were approved)
- [ ] Add a member to a Microsoft 365 (GUID) group → expected: Success message within seconds; the request row ends at Status = Completed with AuthorizationChecked = Yes; the member appears in the card and in the group
- [ ] Add or remove a member on a classic SharePoint (integer) group → expected: The change applies immediately with NO flow run; an audit row appears in Group Membership Requests already marked Completed (“Applied directly to the SharePoint group.”)
- [ ] Open a card for a group that has Group Site Permissions rows → expected: The sites panel lists them; groups without rows (or with the list absent) simply show no panel

## 11. Rollback

1. In the App Catalog, re-upload the **previous** .sppkg version and choose **Replace** — sites pick up the old version on next load.
2. CDN uploads are additive; earlier asset versions stay live, so re-deploying an old package needs no CDN change.
3. Turn the flow Off in Power Automate to stop all Microsoft 365 group changes instantly. Caution: rows created while it is off are NOT picked up when it is turned back on (a created-trigger does not fire retroactively) — the web part reports a timeout for them; resubmit those requests or close them by hand (Status = Failed plus a note in ResultMessage).
4. Remove the web part from its page(s) to stop new requests; the four lists keep the full audit history and can be left in place.

## 12. Glossary

| Term | Meaning |
| --- | --- |
| .sppkg | A SharePoint solution package file. Uploading it to the App Catalog is how a web part is installed in the tenant. |
| App Catalog | The special SharePoint site ("Apps for SharePoint") where administrators upload SPFx packages (.sppkg files) to make them available to sites. |
| CDN | The web host that serves this solution’s JavaScript files in production. In this tenant, packages must load their assets from the CDN, not from SharePoint. |
| Cloud flow | A Power Automate workflow. Automated flows react to events (like a new list item), scheduled flows run on a timer, instant flows are started manually or by an app. |
| Connection | A saved sign-in that a flow uses to talk to a service (SharePoint, Outlook…). Each connector in a flow needs one. |
| Delegated permission / admin consent | SPFx web parts call Microsoft Graph as the signed-in user, but a SharePoint admin must approve each requested permission once (SharePoint admin center → Advanced → API access) before the calls succeed. |
| Idempotent | Safe to run twice: adding someone who is already a member (or removing a non-member) ends as Completed with an explanatory message instead of an error. |
| Indexed column | A column SharePoint pre-sorts behind the scenes so filters on it stay fast on large lists. |
| Internal name | The permanent, code-facing name of a column. It is set from the FIRST name a column is given and never changes afterwards — which is why these instructions create columns using the internal name first, then rename. |
| List | A SharePoint table. Rows are called items; columns are called fields. |
| Microsoft 365 group | An Entra ID (Azure AD) group with a GUID object id — the kind behind Teams and modern SharePoint team sites. In this solution, changed only by the flow. |
| Property pane | The settings panel that opens on the right when you edit a web part on a page. |
| Queue list | A list used as a to-do line for a flow: the app adds a row, the flow processes it and stamps the result on the same row, and the app polls that row to show the outcome. |
| Service account | A licensed, non-personal user account that owns the flow and its connections so the automation does not break when a person leaves. Here it must also be an owner of every managed Microsoft 365 group. |
| SharePoint site group | A classic permission group that lives inside one SharePoint site, identified by a small integer id. Changed directly by the web part as the signed-in user — the flow never touches these. |
| Trigger | The event that starts a flow. Every flow has exactly one trigger followed by one or more actions. |
| UPN (User Principal Name) | A user's sign-in name, e.g. jane@agency.gov — what the flow hands the connector to identify who to add or remove. |
| Web part | A building block you add to a SharePoint page. This solution ships one or more web parts. |
