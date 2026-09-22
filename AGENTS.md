# Account Management

Read ARCHITECTURE.md and the relevant issue/module only.

- Preserve PROD identities, list/property internal names, permissions and configured workflows. Tenant URLs are deployment inputs.
- Back up dirty/untracked files before editing. Use Node 22, npm ci, npm run check, then package:dev or package:prod from the root. Commit source before delivery.
- One SPPKG per channel includes 1 web part(s). Finished builds: releases/<channel>/latest/. No implicit deployment, remote push or original-checkout removal.
- Keep version/channel and M/EX attribution at the bottom of every web-part settings pane.
- Review docs/BASELINE.md for upgrades and docs/KNOWN-ISSUES.md for acceptance. Do not read generated assets/lockfiles wholesale.

- Both channels use the original PROD IDs and names. Follow [package identity and filenames](docs/PACKAGE-IDENTITY.md).
