# Account Management rollout and communications plan

The 20-slide PowerPoint plan prepares the M Executive Office (M/EX) to introduce
365 Account Management. It covers audiences, ownership, readiness, channels,
draft messages, training, support, and adoption review. Every slide includes
speaker notes with additional guidance or source references.

The plan contains no launch dates or fixed intervals. Communications owns the
schedule and distribution. The sequence uses readiness conditions so the office
can set dates after confirming the service, people, and materials.

## Use the plan

Open `Account-Management-Rollout-and-Communications-Plan.pptx` in PowerPoint.
The text and five tables remain editable. Screenshots use fictional local
preview records from Account Management 1.18.0.

- Slides 1–7 explain the service, audiences, and important access distinctions.
- Slides 8–12 set out the rollout phases, messages, channels, and asset owners.
- Slides 13–18 provide announcement and launch emails, training copy, staff
  talking points, a follow-up post, FAQs, and a service notice.
- Slides 19–20 cover evidence, ongoing ownership, and the handoff to Communications.

The office identity follows the supplied Better Task reference. Role titles are
proposed responsibilities, not confirmed staffing assignments. Confirm the
office scope and owners before distributing the plan.

Replace the bracketed fields in the draft copy before sending:

| Field | Confirm with |
| --- | --- |
| `[APP LINK]` | Product owner / page owner |
| `[OFFICE REQUEST ROUTE]` | Office sponsor / rollout lead |
| `[SUPPORT CONTACT]` | Support lead |
| `[GUIDE LINK]` | Product owner / Communications |
| `[TRAINING DETAILS]` | Trainer / Communications |
| `[APPROVED FALLBACK]` | Product owner / support lead |
| `[STATUS LOCATION]` | Communications / support lead |
| Incident-specific fields | Support lead, using verified facts |

The launch email is draft copy for use after release approval. The plan does not
assert that tenant acceptance is complete. Use the current acceptance record in
`docs/KNOWN-ISSUES.md` when deciding readiness. Reconcile existing user guidance
with the launch build before sharing it as the current guide.

## Design and source

The deck adapts the design supplied in the Better Task rollout folder:
Garamond headings, Calibri body text, the Brand America palette, the U.S. flag,
department masthead, pale panels, and the M/EX footer. It retains the reference's
16:9 canvas and unnumbered slide treatment. No official seal was added.

`source/brand-template.pptx` contains the imported reference layouts with the
Better Task slide content removed. The four white line icons come from the
supplied reference. The original reference remains unchanged.

| Source file | Purpose |
| --- | --- |
| `source/plan.json` | Slide copy, draft communications, and speaker notes |
| `source/build-deck.mjs` | Editable slide authoring, rendering, and validation |
| `source/brand-template.pptx` | Reusable reference layouts, palette, fonts, flag, and footer |
| `source/capture-screens.cjs` | Screenshots of the local fictional preview |
| `source/images/` | Screenshots and the reused reference icons |

Content is grounded in `ARCHITECTURE.md`, `docs/BASELINE.md`,
`docs/ACCESS-ADMINISTRATION.md`, `docs/KNOWN-ISSUES.md`, and the relevant
Membership, Access assignments, and Recent requests components. The plan makes
no promise about request processing time. It distinguishes current-user request
history from office-wide adoption evidence.

## Rebuild

Use the bundled Codex presentation runtime and the Presentations skill. Ask
`load_workspace_dependencies` for the current paths; no presentation packages
need to be added to this application's dependencies.

Set these environment variables, then run the builder with the bundled Node
executable from the application root:

```powershell
$env:ARTIFACT_NODE_MODULES = '<bundled Node packages path>'
$env:ARTIFACT_PYTHON = '<bundled Python executable>'
$env:PRESENTATIONS_SKILL = '<installed presentations skill directory>'
$env:DECK_OUT = '<new absolute .pptx path in docs/rollout or scratch/rollout-build/export>'
& '<bundled Node executable>' docs/rollout/source/build-deck.mjs
```

The finalizer requires a new output filename. To replace the checked-in deck,
first back up that file, generate a new validated output, inspect it, and copy
the validated bytes to the delivery filename. Follow the skill's lifecycle
instructions when starting a new authoring operation.

The builder writes private renders, candidate files, and validation receipts to
the ignored `scratch/rollout-build/` directory. Inspect every slide after changes,
especially longer email copy and table cells.

To refresh the screenshots, use Node 22.23.3 and the existing app dependencies:

```powershell
npm ci
npm run preview:build
node preview/serve.cjs
```

In a second terminal, run `node docs/rollout/source/capture-screens.cjs` from the
root. It uses the existing Playwright browser installation, blocks external
requests, and captures fictional data only. Stop the preview server when done.
Rebuild the deck to include the refreshed screenshots.
