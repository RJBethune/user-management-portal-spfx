# Known issues and acceptance

- BUILD-PEERS-01 — RESOLVED during Heft migration: incoming legacy-peer-deps skips build-tool peers. Ajv 8.18.0, PostCSS 8.5.28 and webpack 5.105.4 are explicit dependencies. Verify both release builds and heft start after lock changes; release success alone missed webpack-dev-server's hot runtime dependency.
- BUILD-ORDER-01 — RESOLVED: the print-document helper now precedes its use to satisfy the newer linter; its body and print behavior are unchanged.

- CDN-02 — Classic script loading patch retained for lazy chunks. PROD uses the established CDN; no runtime CDN fallback is assumed. Local package checks do not prove server headers or live asset availability.
- CDN-11 — Compare exact asset hashes after authorized upload and before catalog installation. Never infer safety from SPPKG size or delete duplicate catalog entries automatically.
- UI-ACCEPTANCE — Validate Fluent focus, menus, dialogs, remount and adjacent web parts in authenticated SharePoint after the baseline alignment.
- ACCOUNT-01 — External Power Automate flow authorization and Microsoft 365 membership writes require tenant acceptance. Keep the existing lists, Graph read scopes and configured flow connections; this migration does not provision or replace them.

- ACCESS-01 — Version 1.13.0 adds inline creation and group editing with existing SharePoint permissions. Validate with a list owner, an ordinary assigned admin, an owner with no membership assignment, Add Items withheld, duplicate rows, unique item permissions, and revoked permissions. Confirm saves target the configured list site; exercise single/multiple lookups, clearing optional groups, required/custom field validation, restricted Person fields and guest resolution. Race two editors to confirm ETag conflicts preserve drafts. Verify existing version history and list-triggered workflows after a save. Local fixtures cannot prove tenant enforcement or flow behavior; the app does not elevate or repair permissions.
- KIT-01 — Version 1.12.0 adopts the compatible published Kit cohort documented in BASELINE.md. Validate SharePoint/Light/Dark, custom confirmation focus, two adjacent instances and live section-theme changes in the tenant. Browser fixtures cannot prove the SharePoint shell or CDN behavior.
- GROUPS-01 — Version 1.14.0 replaces membership accordion cards with a group list and details pane. Validate desktop and narrow SharePoint columns, long group names, many groups, list scrolling, Back to groups focus, retained drafts, startCollapsed, printing and all supported group types. Authorization, office scope, membership writes and Power Automate contracts are unchanged.

Record new defects with version/channel, steps, expected/actual results, evidence, affected module, resolution and a regression check. Keep credentials and personal data out of shared issue notes.
