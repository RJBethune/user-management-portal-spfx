# Account Management

Independent SPFx 1.23.2 / React 17 / Heft repository. Source is under src/, build configuration under config/, delivery contracts under delivery/, repeatable tooling under scripts/ and vendor/, documentation under docs/, and ignored finished builds under releases/.

The shared UI uses the pinned UX Kit profile, Svelte theme and patterns. Each web part owns one Kit root and portal host; theme changes preserve the mounted application. `src/ui/` contains the compiled theme and adapters for existing membership controls. The custom confirmation dialog remains, inside the owned portal host.

The accountManagement web part uses read-only Graph queries and SharePoint request lists for Microsoft 365 membership changes. Existing Power Automate flows enforce authorization and perform the writes. SharePoint-group membership uses the current user context. Preserve the custom confirmation dialog, list schema and admin/office boundaries.

`AccessAdministrationService` requires direct membership in the configured SharePoint group on the list site (default `365AccessManagement-Admins`), plus Manage Lists and Edit Items on the Authorized Admins list. Creating an assignment also requires Add Items. Reads and writes recheck membership; missing, unreadable or blank group configuration denies access. Writes use the signed-in SharePoint context and exact item ETags; they preserve list permissions, field names and configured workflows. The list-and-details workspace keeps separate drafts across admin, view, theme and fullscreen changes. See [Access administration](docs/ACCESS-ADMINISTRATION.md).

[Delivery](docs/DELIVERY.md) · [Baseline](docs/BASELINE.md) · [Known issues](docs/KNOWN-ISSUES.md) · [Migration](docs/MIGRATION.md).
