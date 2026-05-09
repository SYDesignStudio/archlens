# ArchLens Project Context

ArchLens is a Streamlit-based AI architectural planning and building regulations platform for UK residential projects. The current product is a dark luxury SaaS interface for SY Design Studio with a gold accent (`#D4C29A`). Future changes should preserve the working UI and make incremental, low-risk updates rather than redesigning the application.

## Product Purpose

ArchLens helps users upload PDF drawing packs, complete a structured project intake, and generate professional AI-assisted reports for:

- Planning Review
- Building Regulations Review

The app is intended for UK projects and currently focuses on householder planning routes, permitted development, prior approval, local policy context, drawing-pack consistency, and building regulations readiness.

## Current Repository Structure

- `app.py` - Main Streamlit application. Owns session state, navigation, wizard flow, custom CSS, PDF preview, report display, branded PDF/Word exports, saved project records, subscription gating, and credit display.
- `pdf_summary.py` - PDF extraction and AI report-generation layer. Extracts text and rendered page images, detects sheets/local authority/proposal features, retrieves local planning policy context, calls the OpenAI Responses API, polishes report text, and can generate planning statements.
- `planning_rules.py` - Deterministic UK householder planning rule engine. Provides pass/fail/needs-confirmation checks for common PD and prior approval routes and formats rule output for prompts/reports.
- `api.py` - FastAPI credits service. Provides credit balance, transaction, Wix credit top-up, credit deduction, and admin restore endpoints backed by local JSON storage.
- `requirements.txt` - Python dependencies for Streamlit, PDF parsing, report exports, OpenAI, JWT, FastAPI, and Uvicorn.
- `README_DEPLOY.md` - Basic Streamlit Community Cloud and Render deployment notes.
- `config.toml` and `S_config.toml` - Streamlit upload limit config (`maxUploadSize = 20`).
- `assets/sy_design_studio_logo.png` - Brand logo used in UI and branded report exports.
- `planning_policies/` - Local policy PDF library used by `pdf_summary.py` to retrieve authority/project-relevant planning context.
- `archlens_rule_engine_update/` - Reference/update copy of app, PDF summary, and updated planning rules files. Treat as supporting material unless the user explicitly asks to merge from it.
- `archlens_sidebar_only_fix.zip` - Archive artifact. Do not modify unless explicitly requested.

## Runtime Entry Points

Streamlit app:

```powershell
streamlit run app.py
```

FastAPI credits service:

```powershell
uvicorn api:app --host 0.0.0.0 --port 8000
```

Important environment variables:

- `OPENAI_API_KEY` - Required by `pdf_summary.py` for AI report generation.
- `ARCHLENS_POLICY_FOLDER` - Optional override for the local planning policy PDF folder. Defaults to `planning_policies`.
- `ARCHLENS_DATA_DIR` - Optional credits API data directory. Defaults to `data`.
- `ARCHLENS_WEBHOOK_SECRET` - Optional override for the FastAPI webhook/admin secret.

## Streamlit App Flow

`app.py` initializes `st.session_state` from `DEFAULT_STATE`, configures a wide Streamlit page, injects custom CSS, then renders the authenticated app shell.

Primary pages:

- Dashboard - saved project/report metrics and recent projects.
- Projects - main structured wizard for project intake, uploads, preview, AI analysis, and report download.
- Reports - generated report library and download history.
- Settings - account, theme, and branding controls.

Project wizard:

1. Select review module and report mode.
2. Add project details, client, address, and proposal description.
3. Select project and property type.
4. Select scope items for Building Regulations mode.
5. Add planning-specific dimensions and permitted-development accuracy answers.
6. Upload PDF drawing files and preview them.
7. Generate the report and download PDF/Word outputs.

`run_archlens_analysis()` is the main handoff from UI to analysis. It validates limits, writes the uploaded PDF to a temporary file, calls either `pdf_summary.analyze_pdf()` or `pdf_summary.analyze_planning_pdf()`, validates required headings, parses sections, builds PDF and Word reports, stores outputs in session state, and adds a saved project record.

## Analysis and Reporting Flow

Building Regulations Review:

- Uses `pdf_summary.analyze_pdf()`.
- Required headings are defined in `BUILDING_REQUIRED_HEADINGS`.
- Report output is rendered and exported by `app.py`.

Planning Review:

- Uses `pdf_summary.analyze_planning_pdf()`.
- Required headings are defined in `PLANNING_REQUIRED_HEADINGS`.
- Combines PDF text extraction, page classification, local authority detection, local policy snippets, proposal feature detection, planning route inference, deterministic rule checks, and OpenAI-generated professional report text.

Rule engine:

- `planning_rules.run_planning_rule_checks()` returns a `RuleEngineResult`.
- Public compatibility helpers include `facts_from_app_context()`, `run_householder_pd_rules()`, and `format_rule_result_for_prompt()`.
- Missing data should usually produce `NEEDS CONFIRMATION`, not guessed compliance.

Exports:

- `build_pdf_report()` creates branded PDF reports using ReportLab.
- `build_word_report()` creates branded Word reports using `python-docx`.
- Gold brand color appears in export styling as `#D4C29A`.

## UI and Design Guardrails

Preserve the current working UI. Do not redesign the entire app unless the user explicitly asks.

Current visual direction:

- Dark luxury SaaS interface.
- Gold accent: `#D4C29A`.
- SY Design Studio branding.
- Streamlit-native controls with custom CSS wrappers.
- Main custom class prefix: `sy-`.
- Existing UI surfaces include `sy-topbar`, `sy-hero`, `sy-project-hero`, `sy-form-card`, `sy-sidepanel`, `sy-subtle-card`, `sy-report-card`, `sy-step-row`, and `sy-new-project-btn`.

When changing UI:

- Make small scoped updates.
- Preserve page structure and wizard flow.
- Keep dark mode readable.
- Preserve gold accent behavior for primary actions, focus states, progress bars, and badges.
- Avoid broad CSS rewrites because `inject_custom_css()` controls much of the current app feel.
- Check mobile-sensitive areas, especially the step row, side panels, upload preview, and report cards.
- Do not remove existing HTML/CSS class names unless replacing every use safely.

## Data and State Notes

The app currently stores saved projects, generated reports, credit display, report files, and wizard fields in Streamlit session state. This is intentionally lightweight and not a persistent database.

The FastAPI credit service persists credit data to JSON:

- Default path: `data/credits_store.json`
- Writes are protected by `STORE_LOCK`.
- Duplicate Wix orders are blocked through `processed_orders`.

Be careful with credit-related behavior. Report download and subscription gating are user-facing business logic.

## Current Constraints and Safety Notes

- PDF upload limit is 20 MB.
- Live analysis page count is limited in `pdf_summary.py`.
- The app relies on exact report headings. If prompt/report changes alter headings, `validate_report_headings()` can block report generation.
- Planning reports should not invent unsupported measurements, fire statements, policy references, or planning route certainty.
- User-entered measurements are reference context only; drawing content takes priority where there is a conflict.
- Local planning policy PDFs are optional support context, not the sole source of truth.
- Do not commit secrets. Several defaults exist in code for local operation, but production should use environment variables/secrets.

## Testing and Verification

For small documentation-only changes, no app runtime test is usually required.

For Python behavior changes:

- Run the relevant module checks where possible.
- Smoke test Streamlit with `streamlit run app.py`.
- For analysis changes, test both Planning Review and Building Regulations Review paths if feasible.
- For UI/CSS changes, inspect the app in a browser at desktop and mobile widths.
- For FastAPI changes, smoke test the health endpoint and at least one protected endpoint with a test secret.

## Future Update Principles

- Preserve the existing app before improving it.
- Prefer narrow patches over broad rewrites.
- Keep business logic and report headings stable unless the user asks for a specific change.
- Keep AI prompt changes conservative and aligned with UK architectural/planning terminology.
- Use deterministic rule checks to stabilize planning route output.
- Keep client-facing report language professional, concise, and free of system/internal wording.
