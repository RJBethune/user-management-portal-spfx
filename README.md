# Account Management

Version 1.15.0. An independent SPFx app for authorized Microsoft 365 and SharePoint group membership changes.

The app uses the shared UX Kit with SharePoint, Light and Dark themes. Its **Access assignments** view helps trusted list administrators review who manages each group, find duplicate assignments and groups without an admin, and add or edit assignments directly in the app.

**Access assignments** places admins on the left and editable group assignments on the right. Switching admins preserves each record's draft, including a new assignment. Narrow columns use **Back to admins**. The Kit header uses the same icon and theme-menu treatment as Better apps.

**Groups you manage** uses an Outlook-style group list and details pane. Search or choose a group on the left; manage its members and owners on the right. Narrow web parts switch between the list and details with **Back to groups**. Changing groups preserves drafts and the selected tab.

## Build and deliver

Use Node 22.23.3 and run these commands from the repository root:

```sh
npm ci
npm run check
npm run preview:build
npm run test:browser
```

Commit source, then run `npm run package:dev` and `npm run package:prod`. Each channel produces one SPPKG in `releases/<channel>/latest/`. Packaging does not deploy. Both channels retain the original PROD identities.

[Architecture](ARCHITECTURE.md) · [Delivery](docs/DELIVERY.md) · [Baseline](docs/BASELINE.md) · [Known issues](docs/KNOWN-ISSUES.md).

## Local preview

SPFx Portfolio loads the actual web part through `preview/adapter.json`. Its app-owned `sourceAliases` map `@microsoft/sp-page-context` to the local permission shim; this mapping does not enter SharePoint packages. Restart this app's local preview after changing the adapter or dependencies.

Portfolio fixtures include sample membership groups and Access assignments. You can select an admin, change draft groups, or search for a sample admin to add. Saves display an explicit local-preview message and never contact SharePoint or Graph. Use authenticated SharePoint DEV testing to verify permissions and persisted changes.

With the Portfolio preview running, test its **Open standalone preview** URL:

```sh
npm run test:portfolio-preview -- http://127.0.0.1:<preview-port>/
```

This check exercises Portfolio's Vite entry. `npm run test:browser` exercises the separate esbuild harness and its simulated REST saves; it cannot verify Portfolio's import resolution.

## Membership and authorization

A Microsoft 365 group GUID routes changes through the existing SharePoint request list and Power Automate flow. The web part has read-only Graph scopes. The flow must independently authorize the request's system-created `AuthorId` before writing membership changes.

An integer SharePoint group ID routes changes through SharePoint REST in the signed-in user's context. The Authorized Admins list controls which groups the UI shows; SharePoint permissions enforce the actual change.

An individual admin can have several groups assigned through `OfficeGroupRecord`. The page's `visibleOffices` setting only narrows what that admin sees. Access administration requires Manage Lists and Edit Items on the Authorized Admins list, plus Add Items for new records. SharePoint enforces list and item permissions on every save. Concurrent changes require an explicit reload; the app never overwrites a newer assignment silently.

## Configuration

The four existing list defaults are `Managed Groups`, `Group Membership Requests`, `Group Management Authorized Admins`, and `Group Site Permissions`. List titles and the list site URL remain configurable in the web-part settings. Existing property and column internal names are preserved.

The property pane also controls introductory and support text, office scope, request timeout, group photos, initial expansion, justification and diagnostics. Version, channel and M/EX attribution remain at the bottom.

[Access administration](docs/ACCESS-ADMINISTRATION.md) explains assignments and their limits. Use [list setup](docs/SharePoint-Lists-Setup.md) and the [flow contract](docs/PowerAutomate-Flow-Spec.md) for configuration. Provisioning and the deployed flows remain external to this application.
