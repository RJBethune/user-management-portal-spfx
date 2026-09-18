# Known issues and acceptance

- BUILD-PEERS-01 — RESOLVED during Heft migration: incoming legacy-peer-deps skips build-tool peers. Ajv 8.18.0, PostCSS 8.5.28 and webpack 5.105.4 are explicit dependencies. Verify both release builds and heft start after lock changes; release success alone missed webpack-dev-server's hot runtime dependency.
- BUILD-ORDER-01 — RESOLVED: the print-document helper now precedes its use to satisfy the newer linter; its body and print behavior are unchanged.

- CDN-02 — Classic script loading patch retained for lazy chunks. PROD uses the established CDN; no runtime CDN fallback is assumed. Local package checks do not prove server headers or live asset availability.
- CDN-11 — Compare exact asset hashes after authorized upload and before catalog installation. Never infer safety from SPPKG size or delete duplicate catalog entries automatically.
- UI-ACCEPTANCE — Validate Fluent focus, menus, dialogs, remount and adjacent web parts in authenticated SharePoint after the baseline alignment.
- ACCOUNT-01 — External Power Automate flow authorization and Microsoft 365 membership writes require tenant acceptance. Keep the existing lists, Graph read scopes and configured flow connections; this migration does not provision or replace them.

Record new defects with version/channel, steps, expected/actual results, evidence, affected module, resolution and a regression check. Keep credentials and personal data out of shared issue notes.
