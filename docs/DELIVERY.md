# Build and release

Both channels use the original PROD IDs and product names. [Package identity and filenames](PACKAGE-IDENTITY.md) is authoritative for channel identity, catalog isolation and renamed upload files; it supersedes earlier split-channel guidance below.

Run npm ci on Node 22.14 or later in the 22.x line, then npm run check. Commit the source and run npm run package:dev or npm run package:prod. These commands validate and package isolated source; they do not deploy.

The single SPPKG and matching assets, hashes, build manifest, check logs and DEPLOYMENT.md are in releases/dev/latest/ or releases/prod/latest/. Immutable history is in releases/<channel>/<version>/<build-id>/. One package contains 1 web part(s). Both channels reuse the original PROD IDs; no extra package per web part is required.

PROD retains all existing IDs and CDN https://irm.azureedge.us/M/user-management-portal-spfx/. The folder/product rename does not rename that published asset route. DEV embeds assets; both channels support the all-sites deployment option. Where block-download policy prevents embedded loading, choose a reviewed DEV CDN profile and rebuild. Tenant origins in profiles are verification defaults, not runtime restrictions.

Upload assets only when authorized, then run npm run verify:cdn -- --build releases/prod/1.11.3/BUILD-ID --origin https://destination.sharepoint.com against the exact uploaded bytes before catalog installation. Validate cold/warm loading, permissions, settings and dialogs on an authenticated SharePoint DEV page. Retain old assets and packages for rollback.

GitHub/GitLab CI builds both channels and retains releases. Configure GitLab runner tags to suit the team. Hosted repo names/remotes are retained; rename them only as a separate deliberate operation. Historical build commands and CDN guidance are superseded by this guide.


## All-sites eligibility in both channels — 2026-09-21

The reviewed delivery toolkit is 1.4.4. Every DEV and PROD profile enables `skipFeatureDeployment: true`. The toolkit defaults omitted settings to true, rejects false, and checks the emitted package flag. Permanent IDs, names, CDN routes, permissions and runtime dependencies stay unchanged.

In the tenant App Catalog, an administrator can select **Enable this app and add it to all sites**, or **Add to all sites** for an existing app. This makes the selected channel’s web parts available to sites and subsites; it does not add them to pages, provision lists or isolate backend data. Site-collection catalogs remain limited to their collection. Confirm the catalog scope, channel identity and affected-subsite availability during acceptance.

This configuration update does not rewrite retained packages. Commit reviewed source and use the standard packaging command to create a new package. Check the emitted deployment flag and catalog version before installation; an older package with a valid file inventory may still lack all-sites eligibility. No tenant deployment is performed by this update.
