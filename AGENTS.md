# AGENTS.md

Guidance for future coding agents working on ArchLens.

## Prime Directive

ArchLens is a working Streamlit SaaS-style platform for UK architectural planning and building regulations review. Preserve the current working UI and behavior. Make incremental safe updates only unless the user explicitly requests a redesign or larger refactor.

Design priorities:

- Dark luxury SaaS visual direction.
- Gold accent must remain `#D4C29A`.
- SY Design Studio branding should remain visible and consistent.
- Do not redesign the full app.
- Do not break the project wizard, report generation, downloads, credit display, or saved report flow.

## Read First

Before editing, read the relevant files in this order:

1. `PROJECT_CONTEXT.md`
2. `app.py`
3. `pdf_summary.py`
4. `planning_rules.py`
5. `api.py` if credits, auth, webhooks, or API behavior are involved
6. `README_DEPLOY.md` and `requirements.txt` if deployment or dependency changes are involved

Treat the current dirty worktree as user work unless proven otherwise. Do not revert existing changes.

## Architecture Map

- `app.py` is the Streamlit app shell and should stay the main UI entry point.
- `pdf_summary.py` is the PDF extraction, OpenAI prompt, local policy context, and planning statement layer.
- `planning_rules.py` is the deterministic permitted-development/planning rule engine.
- `api.py` is the FastAPI credit service.
- `planning_policies/` stores local policy PDFs used as retrieval context.
- `assets/` stores branding assets, currently including the SY Design Studio logo.
- `archlens_rule_engine_update/` is reference/update material. Do not merge from it automatically.

## Safe Editing Rules

- Keep changes tightly scoped to the user request.
- Do not perform cosmetic rewrites across large sections.
- Do not reorganize files unless explicitly asked.
- Do not rename existing Streamlit session state keys casually.
- Do not remove required report headings or change their exact wording without updating validation, parsing, display, and exports together.
- Do not alter credit costs, plan gates, or business rules unless the user asks.
- Do not hard-code new secrets, API keys, tokens, or private URLs.
- Preserve local policy support through `ARCHLENS_POLICY_FOLDER` and `planning_policies/`.

## UI Rules

The current UI is controlled largely by `inject_custom_css()` in `app.py`. The app uses custom `sy-` classes and Streamlit widgets.

When updating UI:

- Preserve the existing navigation pages: Dashboard, Projects, Reports, Settings.
- Preserve the project wizard and its step flow.
- Preserve the dark theme and gold accent.
- Reuse existing `sy-` components/classes where possible.
- Prefer small refinements inside existing containers rather than new page structures.
- Avoid large CSS rewrites.
- Verify text remains readable in dark mode.
- Keep download buttons, upload preview, step progress, side panel, and report cards functional.

Important existing classes include:

- `sy-topbar`
- `sy-hero`
- `sy-project-hero`
- `sy-form-card`
- `sy-sidepanel`
- `sy-subtle-card`
- `sy-report-card`
- `sy-step-row`
- `sy-step-item`
- `sy-new-project-btn`

## Report and Prompt Rules

Report generation depends on exact headings.

Building Regulations headings live in:

- `BUILDING_REQUIRED_HEADINGS` in `app.py`
- `REQUIRED_HEADINGS` in `pdf_summary.py`

Planning headings live in:

- `PLANNING_REQUIRED_HEADINGS` in `app.py`
- `PLANNING_REQUIRED_HEADINGS` in `pdf_summary.py`

If changing headings, update all of these together:

- Required heading lists
- Section order lists
- Parsing/rendering logic
- Word/PDF export behavior
- Prompt repair logic
- Any tests or validation assumptions

Client-facing reports must:

- Use professional UK planning/building-control language.
- Avoid internal/system wording.
- Avoid unsupported certainty.
- Avoid invented dimensions, policy references, or submitted documents.
- Treat drawing content as stronger evidence than user-entered assumptions.
- Use deterministic rule output to stabilize planning-route reasoning.

## Planning Rule Engine Rules

`planning_rules.py` should stay deterministic and conservative.

- Return `NEEDS CONFIRMATION` when information is missing.
- Do not guess compliance.
- Keep rule checks explainable with evidence and action text.
- Preserve compatibility helpers used by `app.py` and `pdf_summary.py`.
- When adding rules, keep public output stable: `RuleCheck`, `RuleEngineResult`, and formatting helpers should remain backward compatible.

## Credits/API Rules

`api.py` provides a local JSON-backed credits service. It is simple by design.

- Keep duplicate order protection.
- Keep `STORE_LOCK` around load/update/save operations.
- Keep protected endpoints protected by `x-archlens-secret`.
- Prefer environment variables for production configuration.
- Do not commit generated `data/credits_store.json`.

## Verification Expectations

Documentation-only changes:

- No runtime test required.

Python logic changes:

- Run a focused syntax/import check where possible.
- Smoke test the affected path.
- For Streamlit behavior, run `streamlit run app.py` and inspect the changed screen.
- For API behavior, run `uvicorn api:app` and test the relevant endpoint.

Prompt/report changes:

- Verify required heading validation still passes.
- Check both Planning Review and Building Regulations Review if the shared reporting path changed.

UI changes:

- Inspect the app visually.
- Check desktop and a narrow/mobile viewport if layout changed.
- Confirm no text overlap and no unreadable controls in dark mode.

## Deployment Notes

Streamlit deployment starts from:

```powershell
streamlit run app.py
```

Render deployment for the Streamlit app uses:

```powershell
streamlit run app.py --server.port $PORT --server.address 0.0.0.0
```

FastAPI credits service can run with:

```powershell
uvicorn api:app --host 0.0.0.0 --port 8000
```

Required dependency list is in `requirements.txt`.

## Final Reminder

This codebase already contains the working product direction. Improve it carefully from inside the existing structure. Small, verified, brand-consistent changes are preferred over broad rewrites.
