# Implementation docs for user-management-portal-spfx

**Source of truth:** `implementation.json` (this folder). **Deliverable:** `IMPLEMENTATION.md`, generated — never edit it by hand.

Workflow:

1. Fill in `implementation.json`. The schema that validates it (and documents every field) is
   `schemas/implementation.schema.json` in the **spfx-implementation-docs** repo.
2. Render from your clone of spfx-implementation-docs:
   `node bin/spfx-docs.mjs --app <path-to-this-project>`
   (or `spfx-docs --app …` if you ran `npm link`). The guide is also auto-published to the
   workspace `_implementation-docs` folder when one exists.
3. Commit BOTH files together.

Authoring rules that keep documents consistent and novice-safe:

- **You supply data, the generator supplies procedure.** Never write click-paths in the JSON —
  list a column's type/name/choices and the generator emits the exact create-steps (including the
  internal-name-first naming trick).
- **Internal names are contract.** Copy them from the app's code (service $select clauses,
  constants), not from what the list "should" have.
- **Flows: one step object per designer action**, in order, with the exact action name
  ("SharePoint → Get items") and a configuration table. Someone who has never opened
  Power Automate will follow it literally.
- **Plain language.** Summaries and purposes should make sense to a non-developer.
- The deployment and web-part sections are generated from `config/package-solution.json`,
  `config/write-manifests.json`, and the web part manifests — keep those accurate and the doc
  stays accurate.
