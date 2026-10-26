# Baseline and shared UI

Account Management 1.16.2 uses published delivery toolkit **1.5.3**, SPFx **1.23.2**, Heft **1.2.19**, TypeScript **~5.8.0**, React **17.0.1**, and Node **22.23.3**. Both channels retain the original PROD identities and product names. See [package identity and filenames](PACKAGE-IDENTITY.md).

The toolkit archive is unchanged: SHA-256 `3390fa4ec29fec6473fd0ff6c2e4ef23bc3f02db35c14558cae328cc6de2582b`. The platform checkout's newer baseline guidance is unreleased; it does not replace this immutable archive. The app adopts its Node 22.23.3 minimum in the engine, local version file and CI.

## Pinned UX Kit cohort

These compatible published archives are vendored locally and verified against the UX Kit release ledger. They are the adopted cohort, not a claim that no newer Kit versions exist.

| Package | Version | Archive SHA-256 |
| --- | --- | --- |
| `@spfx-kit/ui-profile` | 0.3.0 | `f8cf993e8a295d56a918e440ed524552c8eed9392f3174da799c98f7d782cac0` |
| `@spfx-kit/ui-theme` | 0.9.0 | `b9c7803c917a0a7c9d99fe1d29fc75c633dada078abc953d12de2799ef29c7ff` |
| `@spfx-kit/ui-patterns` | 0.5.0 | `2dbc30fc85addb53a51ef0585ee3ff54dfb26f48bdb07e847d63c8c81500a9c6` |

`vendor/ux-kit-svelte.json` is the shared Svelte style, ID `8f6bcff6-bc73-4116-811a-d40a6065a20c`. Its exact bytes supply the artifact revision digest. `npm run theme:compile` produces the runtime snapshot; `theme:check` verifies it and all three archives. No runtime dependency points at a sibling checkout or GitLab.

Kit controls, header, tabs, notices and status badges replace the Fluent component consumers. Compatibility pins remain at Fluent 9.70.0, Tabster 8.5.5 and Keyborg 2.6.0 until a separate dependency cleanup. The custom confirmation dialog remains as required by the app contract. Its portal now belongs to the Kit host and receives the same theme.

The webpack adapter ships only Luma profile styles, inline with the web part, plus shared theme and pattern CSS and local fonts. Existing classic CDN loading patches remain. SharePoint, Light and Dark preferences are stored per site and user. Live SharePoint theme changes do not remount the app.

## Validation and future upgrades

Use Node 22.23.3, `npm ci`, and `npm run check`. For UI changes, run `npm run preview:build` and `npm run test:browser`; the browser harness uses fictional records and simulates assignment writes in memory. It makes no tenant writes. Install Playwright's Chromium once with `npx playwright install chromium` when needed. Commit source before running both `package:dev` and `package:prod`. Finished builds are in `releases/<channel>/latest/`.

For another upgrade, review released contracts and migration notes, back up dirty or untracked work, create an app branch, and vendor exact archives and hashes. Refresh the lock deliberately. Preserve PROD IDs, Graph scopes, list and property names, configured flows and earlier releases. Do not use `npm audit fix` as an implicit upgrade.

Local tests do not prove tenant acceptance. Validate focus, dialog cancellation, live themes, remount, adjacent web parts, inline assignment saves and existing Power Automate writes in authenticated SharePoint. See [known issues](KNOWN-ISSUES.md).

Version 1.16.0 retains this Kit cohort and adds its published `AppHeaderExpandButton`. Validation covers 17 service regressions, 29 browser cases, and the actual Portfolio Vite preview. New coverage includes the SharePoint group gate, revoked membership, fullscreen portals and cleanup, narrow and adjacent instances, browser-fullscreen success/failure, and Recent requests outcomes and empty/error states. The setting defaults to `365AccessManagement-Admins`; no tenant permissions are provisioned or changed.

Version 1.16.1 retains the pinned Kit components and adds a scoped expanded-header surface using `--spfx-ui-color-primary` and `--spfx-ui-color-primary-foreground`. In SharePoint mode these resolve from the current site theme through the Kit's existing theme layer. Hover and open controls use the Kit's paired neutral surface and foreground; keyboard focus uses the banner foreground. Restoring the workspace removes the banner styling, and portaled menus retain their normal theme. No Kit files, spacing, typography, or component behavior are overridden.

Version 1.16.2 corrects the SPFx stylesheet pipeline for `@spfx-kit/ui-patterns`. The app's webpack adapter sends only that exact package stylesheet through the existing global CSS rule, retaining its Kit host scope and literal React class names. App CSS remains subject to CSS-module isolation. The installed SPFx production loader now runs in a regression test that checks every pattern class, ordinary module/global CSS, and duplicate adapter application. This check fails against the previous adapter with `Kit selector was renamed or removed: spfx-uip-app-header`; Vite and esbuild previews alone could not detect the packaging defect. Kit pins remain unchanged.

Local validation for 1.15.0 uses Node 22.23.3: a fresh `npm ci`, `npm run check` with 14 regression tests, and 22 Chromium browser tests. Coverage includes current-user permission gates, conditional writes, conflicts, failed saves, field compatibility, directory selection, keyboard behavior, draft retention and desktop/mobile layouts. Both workspaces are checked in narrow containers inside a wide browser. Access coverage includes separate drafts for duplicate records, new-assignment drafts, unique field IDs, save locking, filtering, and drafts remaining reachable after a concurrent deletion or permission change. Group navigation covers long names and `startCollapsed`. The published baseline audit reports one nonblocking advisory because its recorded Node minimum is 22.14.0 and this app now requires 22.23.3. Tenant acceptance is pending.
