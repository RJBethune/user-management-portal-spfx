# Power Automate Flow Spec — 365 Account Management (O365 group requests)

This is the **app-side contract** the flow must satisfy. It only handles **O365 (GUID) group** requests —
SharePoint site-group changes are applied directly by the web part (not via Graph), but they ARE still logged
to this same list as **`Status = Completed`** audit rows, so the flow MUST skip any non-Pending row (see the
guard in section 1) or it will re-process them and fail. Reconcile with
your existing flow build notes; the flow-side changes for the current release (v1.11.0) are summarized in **§7**.

---

## 1. Trigger & identity
- **Trigger:** *When an item is created* on the **Group Membership Requests** list (the request/audit list).
- **FIRST ACTION (required) — guard on Status:** if `Status` is **not** `Pending`, **terminate immediately**
  (do nothing / Cancelled). The SharePoint-direct path writes `Completed` audit rows to this same list; without
  this guard the flow re-processes them, the Graph call fails on the SharePoint group's **integer** `GroupId`,
  and it overwrites `Status` to `Failed` — the user sees "Completed, then Failed." Only `Status = Pending` rows
  are real O365 work. (Optionally also skip when `GroupId` is not a GUID, as belt-and-suspenders.)
- **Concurrency:** Degree of Parallelism = **1** (serialize; avoids interleaving on the same group).
- **Runs as:** a service account that is **Owner of every managed O365 group**, or an app registration with
  application permission **`GroupMember.ReadWrite.All`**. (This identity is the elevation — the web part itself
  holds only read scopes.)
- **Owner actions need a broader scope:** `Add Owner` / `Remove Owner` (see §5) act on the group's **owners**
  collection, which requires **`Group.ReadWrite.All`** (delegated or application) — `GroupMember.ReadWrite.All`
  covers members only. Add and admin-consent `Group.ReadWrite.All` on the flow identity **before** enabling
  owner requests, or those rows fail with an authorization error.

## 2. Request item — fields the flow READS
| Field (internal name) | Type | Use |
|---|---|---|
| `Action` | Choice | `Add Member`, `Remove Member`, `Add Owner`, or `Remove Owner` |
| `GroupId` | Text | **Target O365 group object id (GUID)** — the group to change |
| `MemberId` | Text | **Member's Entra object id (GUID)** — the person to add/remove (use for the Graph call) |
| `MemberDisplayName`, `GroupName` | Text | Display / logging only |
| `TargetUserPrincipalName`, `TargetUserEmail` | Text | Member UPN/email (verification / fallback) |
| `Author` (Created By) → `AuthorId` | Person | **The requester. The ONLY trusted identity for authorization.** |
| `CorrelationId` | Text | Correlate the row to the web part submission / flow run |
| `Justification` | Multi-line | Reason (audit) |

> **Never authorize on client-writable fields** (`RequestedBy`/`RequestedById`, `OfficeGroupRecordId`, `GroupId`
> as a trust anchor). Only `Author`/`AuthorId` is server-stamped and trustworthy.

## 3. Request item — fields the flow WRITES back
| Field | Value |
|---|---|
| `Status` | **`Completed`** or **`Failed`** — exact casing (the web part polls for these literally; a mis-cased value hangs the UI at timeout) |
| `ResultMessage` | Outcome or error text |
| `AuthorizationChecked` | Yes/No |
| `AuthorizationResult` | Authorization outcome text |

**Invariant:** no path may leave `Status = Pending`. Every branch (authorized, not-authorized, Graph error,
exception) must end by writing a terminal `Status`.

## 4. ⭐ Authorization re-check — UPDATED for the multi-value lookup

The **Group Management Authorized Admins** list is now **one row per admin**, and `OfficeGroupRecord` is a
**multi-value lookup** into Managed Groups (an admin's whole group set is on one row). Authorize like this:

**Step 4a — resolve the requester's SharePoint user id.** Read the created item's numeric `AuthorId`:
```
GET {listsSite}/_api/web/lists/getbytitle('Group Membership Requests')/items({ID})?$select=AuthorId
```

**Step 4b — read the requester's authorized groups (expanding the multi-value lookup):**
```
GET {listsSite}/_api/web/lists/getbytitle('Group Management Authorized Admins')/items
    ?$select=Id,OfficeGroupRecord/GroupId
    &$expand=OfficeGroupRecord
    &$filter=UserId eq {AuthorId}
Header:  Accept: application/json;odata=nometadata
```
> Filter by **`UserId`** (the single-value `User` Person column — filterable, indexed), **not** by the group.
> SharePoint can't reliably `$filter` a *multi-value* lookup's sub-field, so you fetch the requester's row(s)
> and test the group membership **inside the flow**.

**Step 4c — test membership.** From the response, build the set of **all** `OfficeGroupRecord[].GroupId`
values across the returned row(s). The request is **authorized** iff the request item's `GroupId` is in that set
(case-insensitive GUID compare). Under `odata=nometadata`, `OfficeGroupRecord` is a plain **array** per row
(`item.OfficeGroupRecord` → `[ { GroupId }, { GroupId }, … ]`); under `odata=verbose` it is `{ "results": [ … ] }`.

- **Authorized →** proceed to §5, set `AuthorizationChecked = Yes`, `AuthorizationResult = "Authorized"`.
- **Not authorized (group not in the set, or no rows) →** **fail closed:** `AuthorizationChecked = Yes`,
  `AuthorizationResult = "Requester not authorized for this group"`, `Status = Failed`, and a clear
  `ResultMessage`. **Make no Graph call.**

## 5. Perform the membership change (only if authorized)
Using `GroupId` (target group) and `MemberId` (member object id):

- **Add Member:**
  ```
  POST https://graph.microsoft.com/v1.0/groups/{GroupId}/members/$ref
  Body: { "@odata.id": "https://graph.microsoft.com/v1.0/directoryObjects/{MemberId}" }
  ```
  Treat **HTTP 400 "already exists"** as success (idempotent).
- **Remove Member:**
  ```
  DELETE https://graph.microsoft.com/v1.0/groups/{GroupId}/members/{MemberId}/$ref
  ```
  Treat **HTTP 404 "not found"** as success (already not a member).
- **Add Owner:** (needs `Group.ReadWrite.All` — see §1)
  ```
  POST https://graph.microsoft.com/v1.0/groups/{GroupId}/owners/$ref
  Body: { "@odata.id": "https://graph.microsoft.com/v1.0/directoryObjects/{MemberId}" }
  ```
  Treat **HTTP 400 "already exists"** as success. (An owner should also be a member — optionally also POST to
  `/members/$ref` so the new owner is a member too.)
- **Remove Owner:** (needs `Group.ReadWrite.All`)
  **Guard first — refuse to remove the last owner:** GET `https://graph.microsoft.com/v1.0/groups/{GroupId}/owners?$count=true`
  (header `ConsistencyLevel: eventual`); if the group has only **one** owner, write `Status = Failed`,
  `ResultMessage = "Cannot remove the last owner"` and make **no** change. Otherwise:
  ```
  DELETE https://graph.microsoft.com/v1.0/groups/{GroupId}/owners/{MemberId}/$ref
  ```
  Treat **HTTP 404 "not found"** as success (already not an owner).

On any other Graph error → `Status = Failed`, `ResultMessage` = the Graph error.
On success → `Status = Completed`, `ResultMessage` = e.g. "Added"/"Removed".

## 6. Invariants (contract the web part relies on)
1. **Authorize on `AuthorId` only** (server-stamped Created By); ignore all client-writable fields for security.
2. **Fail closed** — any error, missing row, or ambiguity → `Status = Failed`, no Graph call.
3. **Exact Status casing** — `Completed` / `Failed` (never `Complete`, `Success`, etc.).
4. **No row left `Pending`** — every branch writes a terminal status. (A scheduled "sweep" flow that fails
   stale `Pending` rows older than N minutes is recommended as a backstop.)
5. **Idempotent** — tolerate already-member / not-member (see §5).
6. Only the flow identity should be able to edit `Status`/`Authorization*` (list permissions).

## 7. What changed — flow work for the current release (v1.11.0)

**NEW — owner management (needs a new Graph permission).** The web part now files `Add Owner` and
`Remove Owner` requests (O365 groups only). To make them work the flow must:
- Handle the two new `Action` values (§2) with the group **owners** Graph calls in §5.
- Run the **last-owner guard** before every Remove Owner (§5) — never leave a group ownerless.
- Use **`Group.ReadWrite.All`** for owner calls — `GroupMember.ReadWrite.All` covers members only (§1). Add and
  admin-consent this scope **before** enabling owner requests, or owner rows fail with an authorization error.
- Member add/remove branches (§5) are unchanged; owner requests reach the flow through the same trigger, carry
  the same fields (§2), and write back the same way (§3) — only the perform-change branch and permission differ.

**Still required from v1.8.0 — the multi-value authorization re-check (§4).** Authorized Admins is one row per
admin with a **multi-value** `OfficeGroupRecord`; §4 `$expand`s it, filters by `UserId eq {AuthorId}`, and tests
whether the request's `GroupId` is **in** the requester's set. The list's multi-value column and the flow's §4
check must change **together**. The web part reads both list layouts, so the list can migrate gradually.

**Unchanged:** the trigger + `Status = Pending` guard (§1), write-back (§3), and all invariants (§6).

---

## 8. Nightly Group Site Permissions sync (optional, separate flow)

A scheduled flow can keep the **Group Site Permissions** list current instead of hand-maintaining it, by
discovering which sites each managed O365/security group has permissions on. This is a **separate flow** from
the O365 request flow above.

**Identity & scope.** A dedicated **app registration** with **`Sites.Selected`**, using an **app-only bearer
token** acquired via an HTTP action against the gov endpoints (`login.microsoftonline.us` / `*.sharepoint.us`)
and passed as an explicit `Authorization` header on raw SharePoint REST calls (the stock "Send an HTTP request
to SharePoint" runs as the connection user, which usually can't read role assignments). Scan a **curated
"Managed Sites" list**, never the whole tenant (Power Automate action limits / throttling).

**Sync-only columns** on Group Site Permissions (the web part ignores them): `Source` (Choice Auto|Manual,
default Manual), `SyncKey` (indexed text = `lower(GroupId)+'|'+canonical(SiteUrl)`), `LastSyncedRunId` (text),
`SyncStatus` (Choice Active|Stale), `StaleRuns` (Number), `LastSyncedUtc` (Date/Time).

**Per site (root web):**
1. `GET {site}/_api/web/roleAssignments?$expand=Member,RoleDefinitionBindings&$select=Member/PrincipalType,Member/LoginName,PrincipalId,RoleDefinitionBindings/Name,RoleDefinitionBindings/RoleTypeKind` + `GET {site}/_api/web?$select=Title`.
2. Match group principals to Managed Groups GroupIds by **known claim prefix** - `c:0t.c|tenant|` = security group, `c:0o.c|federateddirectoryclaimprovider|` = M365 (strip a trailing `_o`), lowercase both sides.
3. **Mandatory:** for each `PrincipalType=8` SharePoint site group, `GET sitegroups(<PrincipalId>)/users` and match its members too - most group grants are nested there, not direct.
4. `Permission` = the highest `RoleDefinitionBindings` by `RoleTypeKind`, **dropping "Limited Access."**
5. Diff vs the site's existing `Auto` rows; create/update via OData `$batch`; stamp `LastSyncedRunId` / `SyncStatus=Active`.

**Reconcile (soft delete).** For **successfully-scanned sites only**, unstamped `Auto` rows -> `SyncStatus=Stale`,
`StaleRuns++`; hard-delete only at `StaleRuns >= 3`. **Abort** the sweep if any site GET failed or the stale set
exceeds ~50% - this prevents a one-bad-night mass wipe. Every query/write/delete is filtered `Source eq 'Auto'`,
so hand-entered `Manual` rows are never touched.

**Safeguards:** a run-lock (no overlapping runs), a `DryRun` report-only mode for first validation, canonical
`SiteUrl` in the key, fetch `ListItemEntityTypeFullName` at runtime (don't hardcode it), and a run-log list +
failure alerts.

**Known limits:** site **root web only** (subsite/list/item unique permissions not captured); Site Collection
Admins do not appear in role assignments; one level of group nesting. See **Setup-Guide.html -> "Nightly sync"**
for the full walkthrough.
