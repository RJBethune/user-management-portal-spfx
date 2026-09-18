# Account Management

Independent SPFx 1.23.2 / React 17 / Heft repository. Source is under src/, build configuration under config/, delivery contracts under delivery/, repeatable tooling under scripts/ and vendor/, documentation under docs/, and ignored finished builds under releases/.

The accountManagement web part uses read-only Graph queries and SharePoint request lists for Microsoft 365 membership changes. Existing Power Automate flows enforce authorization and perform the writes. SharePoint-group membership uses the current user context. Preserve the custom confirmation dialog, list schema and admin/office boundaries.

[Delivery](docs/DELIVERY.md) · [Baseline](docs/BASELINE.md) · [Known issues](docs/KNOWN-ISSUES.md) · [Migration](docs/MIGRATION.md).
