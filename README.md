# 365 Account Management (user-management-portal-spfx)

SharePoint Framework (SPFx) web part that lets designated office admins **add and remove members of
Microsoft 365 groups and SharePoint site groups** from a SharePoint page — with a live Entra profile
card per member, per-page office scoping, and a configurable backend.

## How it works

The web part branches on the **Group ID format** of each configured office:

| Group ID in the Group list | Path | What happens |
|---|---|---|
| **GUID** (Entra/O365 group) | Request + flow | Files a `Pending` item to a SharePoint **request list**; a **Power Automate flow** (running as a service account or app registration with `GroupMember.ReadWrite.All`) performs the Graph change and writes the status back. The web part polls for the result. |
| **Integer** (SharePoint site group) | Direct REST | Changes membership directly via SharePoint REST in the signed-in user's context — no flow. |

The web part holds **read-only** Graph scopes (`GroupMember.Read.All`, `User.ReadBasic.All` — escalate the
search to `User.Read.All` for production); all O365 writes are delegated to the flow.

Authorization is data-driven via a SharePoint **Authorized Admins** list (per user → office), enforced
client-side for display and **re-checked server-side by the flow** on the request item's `Created By`.

## Required SharePoint lists

Provisioned by hand (titles are configurable in the property pane; defaults shown):

- **Group (offices)** — `Group Management Test Data`: one row per office. Key column `GroupId` (Entra GUID
  for O365, integer for SharePoint groups). Optional `GroupType` column flags unmanageable types
  (`Dynamic`, `Distribution`, `MailEnabledSecurity`, `RoleAssignable`, `OnPrem`).
- **Authorized Admins** — `Group Management Authorized Admins`: Person column `User` + Lookup
  `OfficeGroupRecord` into the Group list (must project `GroupId`).
- **Request** — `Group Membership Audit Test`: the queue the O365 flow processes. `Status` is single-line
  text (`Pending` → `Completed`/`Failed`); plus `Action`, `GroupId`, `MemberId`, `TargetUserPrincipalName`,
  `ResultMessage`, `AuthorizationChecked` (Yes/No), `AuthorizationResult`, `CorrelationId`, etc.

## Property pane

- **Display** — Title.
- **Offices on this page** — comma-separated office names/Group IDs to show (blank = all you're authorized for).
- **Data source** — the three list titles (point different pages at different request lists to route to
  different flows).
- **Behavior** — O365 status-wait timeout (seconds) and a verbose-logging toggle (off by default).

## Build

```bash
npm install
gulp serve                                   # local workbench
gulp clean && gulp bundle --ship && gulp package-solution --ship
```

The `.sppkg` is written to `sharepoint/solution/`. For production in a CDN-fronted tenant, set
`config/package-solution.json` `includeClientSideAssets=false` and `config/write-manifests.json`
`cdnBasePath`, then host the assets on the CDN.

## Not in this repo

The **Power Automate flow** and the **manual list provisioning** are separate deliverables (the flow design
and provisioning specs are maintained outside the source tree).

## Stack

SPFx 1.21.1 · React 17 · Fluent UI 8 · `@pnp/spfx-controls-react` (LivePersona) · TypeScript 5.3.
