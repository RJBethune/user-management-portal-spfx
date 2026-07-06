# Deployment & CDN Runbook — 365 Account Management

Ops guide for hosting and releasing the web part. (For building the SharePoint lists and the Power
Automate flow, see **Setup-Guide.html** / **SharePoint-Lists-Setup.md** / **PowerAutomate-Flow-Spec.md**.)

## Why the CDN (the block-download policy)
Production sites in this tenant sit behind a Conditional Access / block-download policy that 302-redirects
file downloads from SharePoint libraries to an HTML policy page. SPFx assets served the normal way — from the
App Catalog's **ClientSideAssets** library — therefore **do not load** on production sites (RequireJS
"scripterror"; the asset request comes back as `text/html`). So production assets are hosted on the Azure gov
CDN and the package manifest points there.

    https://irm.azureedge.us/M/user-management-portal-spfx/

## Two build modes
| | Dev / test | Production |
|---|---|---|
| `config/package-solution.json` → `includeClientSideAssets` | `true` (assets embedded) | `false` (CDN) |
| `config/write-manifests.json` → `cdnBasePath` | (ignored) | the CDN folder above |
| Loads assets from | App Catalog **ClientSideAssets** | **Azure CDN** |
| Works on | policy-exempt sites / workbench | **all** sites, incl. block-download |

The PROD (CDN) build is effectively **universal** — the CDN is a public HTTPS endpoint reachable from every
site, block-download or not — so the single CDN package works everywhere. The dev/embedded build exists for the
workbench and policy-exempt test sites where you don't want a CDN dependency.

## Is there a "CDN → ClientSideAssets" fallback? No — and it would not help.
- **Not supported:** `cdnBasePath` **replaces** the manifest's base URL; it does not chain a "primary + fallback
  on error." There is no dependable runtime CDN→ClientSideAssets error-recovery for a *custom* CDN, and the
  bundle load fails at the SPFx **loader** (before any app code runs), so nothing in the app can self-recover.
- **Would not help where it matters:** on production (block-download) sites, ClientSideAssets is exactly what is
  blocked — the fallback target fails there too. A CDN outage on those sites means the app is unavailable,
  fallback or not. This is inherent to the tenant's download policy.
- **Where a fallback would help** (a CDN outage on a policy-**exempt** site) is already covered by the
  dev/embedded build. So there is no real gap for an app-level fallback to fill.
- **Put resilience at the layers you control:** the CDN's own availability / origin-failover (a CDN-team / SLA
  concern), and the deploy-order gate below. Do not attempt an app-level fallback.

## The CDN is a hosting tier — treat it like one
- Azure gov CDN is highly available and geo-distributed; a true outage is rare and usually coincides with
  broader Azure issues.
- Uploads are **additive** (content-hashed filenames) — older versions stay live and **are** the rollback path.

## Production release — the deploy-order gate
1. Bump all three versions (`package.json`, solution version, feature version in `package-solution.json`).
2. Set `includeClientSideAssets: false`; confirm `cdnBasePath` is the CDN folder.
3. Build: `gulp clean && gulp bundle --ship && gulp package-solution --ship`.
4. **Guardrail** — the `.sppkg` manifest must show the CDN URL and **none** of `SPCLIENTSIDEASSETLIBRARY`,
   `localhost`, or `REPLACE-WITH`. (The 4 KB `.sppkg` size confirms assets are externalized.)
5. Hand `release/v<version>/` to the CDN team for the project folder — **additive only; never rename or
   overwrite** published files.
6. **WAIT** for CDN confirmation, then verify in a browser that
   `https://irm.azureedge.us/M/user-management-portal-spfx/<bundle>.js` returns **HTTP 200 JavaScript** (not an
   HTML policy page) **and** carries a CORS `Access-Control-Allow-Origin` header.
7. **Only then** upload the `.sppkg` to the App Catalog and deploy. (Deploying before the CDN files are live
   breaks the web part for all users with an opaque RequireJS error — there is no fallback.)

## CORS (required on the CDN/origin)
SPFx's loader issues a cross-origin **fetch** for the bundle, so the CDN/origin must return CORS headers, or the
console logs "blocked by CORS policy" (even though the `<script>` load itself succeeds). Ask the CDN team for:
Allowed origins `https://usdossiolab.sharepoint.com` (or `*`), methods `GET, HEAD, OPTIONS`, headers `*`. This
is an endpoint-level setting — one-time, benefits every SPFx project on the CDN.

## Rollback
CDN uploads are additive, so the previous version's hashed files are still live. To roll back the app tier,
re-deploy the **previous `.sppkg`** from the App Catalog — its CDN assets remain reachable. Keep prior `.sppkg`
files archived.

## Duplicate-registration hazard
If the App Catalog ever holds two entries registering the same component id, they serve stale manifests. Delete
**all** of them, empty **both** recycle bins on the catalog site, then re-upload **one**.
