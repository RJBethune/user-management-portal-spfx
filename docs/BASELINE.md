# Baseline upgrades

Pinned delivery toolkit 1.4.0; SPFx 1.23.2; Heft 1.2.19; TypeScript ~5.8.0; React 17.0.1; Fluent 9.70.0 and the complete baseline Griffel/Tabster 8.5.5/Keyborg 2.6.0 closure. Both Windows and Linux Sass binaries are declared. npm run check verifies the archive, identities, installed lock closure, source build and app regressions.

For an update, review the released contracts and migration notes, preserve local work, create an app branch, vendor the immutable archive/hash, update explicit pins and lock, run a fresh npm ci, all checks and both deliveries. Compare PROD IDs, permissions, defaults and assets, then accept in SharePoint before promotion. Keep earlier releases. Do not use npm audit fix as an implicit upgrade.

Gulp was migrated following [Microsoft’s manual Heft migration guide](https://learn.microsoft.com/en-us/sharepoint/dev/spfx/toolchain/migrate-gulptoolchain-hefttoolchain-manual). Previous Gulp/ESLint configuration is retained under migration/legacy for reference.
