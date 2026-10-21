# Account Management UI

The [UX Kit design guide](https://gitlab.state.gov/BroomfieldPN/ux-kit/-/blob/main/DESIGN.md) governs shared appearance and interaction. The pinned package cohort and Svelte source are recorded in [Baseline](docs/BASELINE.md). This file records app-specific choices.

The app has two views: Membership for authorized group changes, and Access assignments for trusted list administrators. Keep both views' drafts mounted when switching views or themes. Use the existing office/group terminology without renaming stored properties, columns or flow fields.

Membership uses a compact group list beside a single details pane, following the list-and-reading-pane arrangement familiar from Outlook on the web. The list contains search, a selected-record state and Print all groups. Selecting a group shows its members, owners and existing actions without expanding a stack of cards. People use flat rows separated by rules. Shared Kit controls, typography, tokens and themes remain authoritative; this is not a copy of Outlook branding.

The group list is 288px wide when the web part has more than 800px available. At narrower container widths, show either the list or the selected group's details with Back to groups. Return keyboard focus to the selected record, or the search field if the record is filtered out. Filtering does not discard the selected details or draft. Per-group search, selection, justification and active-tab state persist when switching groups. `startCollapsed` leaves no initial selection. In-app navigation marks its current view with `aria-current`; record navigation uses `aria-current="true"`.

Access assignments presents an admin-to-group table, search, duplicate-record guidance and groups without assigned admins. Add and edit actions open an inline form with Kit inputs, directory results and searchable group checkboxes. Existing records keep their admin identity; the editor changes their groups. Show selection counts and added/removed counts, preserve failed drafts, and require explicit discard before canceling or reloading a changed record. Empty membership access must not hide administration from an authorized list manager.

Use Kit controls and semantic tokens. `src/ui/Controls.tsx` translates existing app event shapes to Kit primitives; it must not restyle those primitives. The stylesheet holds app-specific membership rows, layout and the retained confirmation dialog. The print report intentionally uses a self-contained light stylesheet for paper.

The custom confirmation dialog is the app's explicit exception to the shared AlertDialog preference. Preserve its confirmation wording, keyboard trap, cancellation and return focus. Portal into the instance-owned Kit host. SharePoint still requires acceptance for modal interaction with surrounding page controls.

Keep version, channel and M/EX attribution at the bottom of the settings pane. Do not display build diagnostics above the user task.
