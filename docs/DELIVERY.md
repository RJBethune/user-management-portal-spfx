# Build and release

Run npm ci on Node 22.14 or later in the 22.x line, then npm run check. Commit the source and run npm run package:dev or npm run package:prod. These commands validate and package isolated source; they do not deploy.

The single SPPKG and matching assets, hashes, build manifest, check logs and DEPLOYMENT.md are in releases/dev/latest/ or releases/prod/latest/. Immutable history is in releases/<channel>/<version>/<build-id>/. One package contains 1 web part(s). DEV IDs allow coexistence with PROD; no extra package per web part is required.

PROD retains all existing IDs and CDN https://irm.azureedge.us/M/user-management-portal-spfx/. The folder/product rename does not rename that published asset route. DEV embeds assets and is site-scoped. Where block-download policy prevents embedded loading, choose a reviewed DEV CDN profile and rebuild. Tenant origins in profiles are verification defaults, not runtime restrictions.

Upload assets only when authorized, then run npm run verify:cdn -- --build releases/prod/1.11.3/BUILD-ID --origin https://destination.sharepoint.com against the exact uploaded bytes before catalog installation. Validate cold/warm loading, permissions, settings and dialogs on an authenticated SharePoint DEV page. Retain old assets and packages for rollback.

GitHub/GitLab CI builds both channels and retains releases. Configure GitLab runner tags to suit the team. Hosted repo names/remotes are retained; rename them only as a separate deliberate operation. Historical build commands and CDN guidance are superseded by this guide.
