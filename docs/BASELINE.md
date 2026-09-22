# Baseline upgrades

Current delivery toolkit: **1.5.3**, SHA-256 `3390fa4ec29fec6473fd0ff6c2e4ef23bc3f02db35c14558cae328cc6de2582b`. Both channels retain the original PROD identities and product names. See [package identity and filenames](PACKAGE-IDENTITY.md). Runtime/toolchain dependencies are unchanged. The records below describe earlier adoption steps.

Pinned delivery toolkit 1.4.0; SPFx 1.23.2; Heft 1.2.19; TypeScript ~5.8.0; React 17.0.1; Fluent 9.70.0 and the complete baseline Griffel/Tabster 8.5.5/Keyborg 2.6.0 closure. Both Windows and Linux Sass binaries are declared. npm run check verifies the archive, identities, installed lock closure, source build and app regressions.

For an update, review the released contracts and migration notes, preserve local work, create an app branch, vendor the immutable archive/hash, update explicit pins and lock, run a fresh npm ci, all checks and both deliveries. Compare PROD IDs, permissions, defaults and assets, then accept in SharePoint before promotion. Keep earlier releases. Do not use npm audit fix as an implicit upgrade.

Gulp was migrated following [Microsoft’s manual Heft migration guide](https://learn.microsoft.com/en-us/sharepoint/dev/spfx/toolchain/migrate-gulptoolchain-hefttoolchain-manual). Previous Gulp/ESLint configuration is retained under migration/legacy for reference.


## Dashboard tooling update — 2026-09-21

Historical delivery toolkit: **1.4.3** (previously 1.4.0). Archive SHA-256: `f4020a6f0cc027de6f6d4645278e266ff74f892151a48d7badf65abf63259026`. Earlier rollout descriptions remain above as history. Runtime contracts, dependency pins and identities are unchanged. See the dashboard update record for the check result and recovery path. Review and commit the update, then run both channel builds and SharePoint acceptance.

Historical delivery toolkit: **1.4.4**. Both channels enable the all-sites deployment option. Runtime contracts and dependency versions are unchanged. Archive SHA-256: `78e0ba70f26d62f44e263a556441d783712d5a2fa652223ca5049cf47e1157e5`. Source checks and migration evidence: local portfolio audit dated 2026-09-21.
