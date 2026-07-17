# `batch_edit.py` — Phased Development Plan

**App type:** Single-window PyQt5 desktop app, launched via `./venv/Scripts/python.exe batch_edit.py`
**Purpose:** Run a configurable, runtime-extensible pipeline of Word document operations (find/replace, TOC generation, more later) against a single `.docx` file.

This plan is based on the architecture patterns already proven in `word_find_replace_enhanced.py` (dual-engine COM + python-docx design), extended to support a generic, pluggable operation pipeline rather than a single-purpose tool.

---

## 0. Key Architectural Decisions (settle before coding)

These choices ripple through every phase, so they're called out up front rather than buried in a phase.

| Decision | Recommendation | Why |
|---|---|---|
| **Engine strategy** | Dual-engine, same as reference app: **COM (`pywin32` + Word)** as primary engine, **`python-docx`** as fallback | TOC generation via `Document.TablesOfContents.Add(...)` is native Word functionality with no reliable python-docx equivalent for custom heading-level selection + full formatting control. python-docx can only insert a static TOC *field code* that Word must update on open — acceptable as a degraded fallback, not as primary. |
| **Operation model** | Each operation is a self-contained class implementing a common `Operation` interface (`validate()`, `run(doc_handle)`, `describe()`) registered in an **operation registry** | Enables "add operations at runtime" cleanly — new operation types register themselves; the pipeline/GUI don't need to know concrete types in advance. |
| **Pipeline execution** | Operations run sequentially against one open document handle (COM or python-docx), each producing a structured result (success/warning/error + details) | Matches "series of processing operations on a single document" requirement; avoids reopening the file per operation. |
| **Threading** | Long-running pipeline execution happens on a `QThread` worker, GUI stays responsive, progress/log signals stream back | Same pattern as reference app (`_refresh_availability`, worker threads). Needed since COM automation and large docs can be slow. |
| **Config persistence** | JSON config file next to the script (e.g. `batch_edit_config.json`) for last-used paths, default engine preference, saved operation presets | Matches "Config" tab requirement; also enables re-running a saved pipeline later. |
| **Style discovery** | Read available styles from the *currently selected* document at load time (via COM `Document.Styles` if available, else python-docx `document.styles`), not from a static list | Required so the GUI always reflects the actual doc's styles, including custom/non-standard ones. |

**Decision confirmed:** TOC generation requires Word/COM and is disabled (with a clear error) when unavailable — no python-docx fallback. This is reflected in Phase 6 below.

---

## Phase 1 — Project Scaffolding & Environment Validation

**Goal:** Get a runnable skeleton before any real logic.

- Confirm venv layout matches `./venv/Scripts/python.exe batch_edit.py` (Windows venv).
- Verify `PyQt5`, `pywin32`, `python-docx` are installed in that venv; write a small startup dependency-check routine (mirroring the reference app's `_refresh_availability`) that:
  - Hard-fails with a clear message if `PyQt5` is missing (nothing runs without it).
  - Soft-disables COM-dependent features if `pywin32` import or Word COM dispatch fails.
  - Soft-disables python-docx fallback features if `python-docx` import fails.
- Project layout:
  ```
  batch_edit.py            # entry point, thin — creates QApplication + MainWindow
  /core/
      engines.py           # COM engine, python-docx engine, common interface
      operations.py        # Operation base class + registry
      pipeline.py           # pipeline runner, QThread worker
  /operations/
      find_replace.py
      toc_builder.py
  /ui/
      main_window.py
      tabs_config.py
      tabs_pipeline.py
      widgets_common.py    # file/folder pickers, style list widget, log panel
  /config/
      config_manager.py
  batch_edit_config.json  # generated at runtime
  ```
- Deliverable: empty-but-launching main window with tab bar (Pipeline, Config) and a status bar.

---

## Phase 2 — Document Engine Abstraction Layer

**Goal:** One clean interface both operation types (and future ones) call into, hiding COM vs. python-docx differences.

- Define `DocumentEngine` interface: `open(path)`, `save(path=None)`, `close()`, `list_styles()`, `get_paragraphs()`/equivalent, `capabilities` (a set of supported features so the pipeline/UI can gray out unsupported operations).
- `ComEngine`: wraps `pythoncom.CoInitialize()` + `win32com.client.dynamic.Dispatch("Word.Application")`, opens doc via `Documents.Open`, exposes style enumeration via `Document.Styles`, exposes TOC insertion via `Document.TablesOfContents.Add`.
- `DocxEngine`: wraps `python-docx`'s `Document()`, exposes `document.styles`, body paragraph/run access for exact-match replace.
- Engine selection logic: try COM first (if available and Word installed), fall back to python-docx, and surface *why* a given engine wasn't used in the log panel (reuse the reference app's availability-check pattern).
- Deliverable: can open a doc, list its styles, and close it cleanly via both engines, exercised with a small manual test script (not yet wired to GUI).

---

## Phase 3 — Operation Framework (the "add operations at runtime" core)

**Goal:** A pipeline that can hold an ordered list of heterogeneous operations, added/removed/reordered by the user before running.

- `Operation` base class: `name`, `required_capabilities`, `params` (schema-described, so the GUI can auto-generate a form later if desired — or at minimum validate params), `validate(engine)`, `run(engine) -> OperationResult`.
- `OperationResult`: status (`ok` / `warning` / `error`), message, optional details (e.g., number of replacements made).
- `OperationRegistry`: maps a display name → operation class + a factory for its parameter-editing widget. New operation types self-register via decorator, so adding a third operation type later doesn't touch the pipeline code.
- `Pipeline`: ordered list of configured operation instances; `run(engine)` executes them in order, stops or continues past an error based on a user-chosen policy ("stop on first error" vs "continue and report all").
- Deliverable: headless test — build a pipeline in code with 2 dummy operations, run it against a sample doc, confirm ordering and error propagation work.

---

## Phase 4 — Find/Replace Operation

**Goal:** Port the reference app's proven logic into the new operation-framework shape.

- Parameters: find string, replace string, scope (body / headers-footers / comments — COM only), match case, whole word, wildcard/regex toggle, optional source formatting spec (font, bold/italic/underline, color, highlight), optional target formatting spec.
- COM path: native Word `Find`/`Replace` via `Range.Find`, supporting wildcard mode and formatting-only replace (mirrors reference app's regex/format-only modes).
- python-docx path: exact-string match only, body paragraphs/runs only, formatting applied via run properties post-replace — no regex, no headers/footers/comments (matches the reference app's documented limitation).
- Capability gating: if selected engine is `DocxEngine`, wildcard/regex option and header/footer/comment scope are disabled in the UI, not just silently ignored.
- Deliverable: operation runs standalone against a test doc under both engines, with a diff/count of replacements reported back.

---

## Phase 5 — Arabic RTL Normalize Operation

**Goal:** Add a first-class operation for normalizing Word content for Arabic right-to-left presentation and editing.

- Parameters: scope (body / headers-footers / comments — COM only), paragraph alignment policy, reading order/direction (`RTL`), optional font/style preset for Arabic content, optional list/table normalization flags, and whether to touch only paragraphs already containing Arabic-script characters vs. force all selected content.
- COM path (primary): apply native Word paragraph and table direction settings such as right alignment, bidi/reading-order properties, and table/cell RTL formatting where available through Word automation. This is the preferred path because Word exposes the actual layout properties used by the final document.
- python-docx path (fallback): limited body-only support. Apply paragraph alignment and bidi-related XML where safely supported, but document clearly that full parity with Word COM is not guaranteed for headers/footers/comments, list behavior, or table direction.
- Capability gating: if selected engine is `DocxEngine`, disable unsupported scope targets and any table/comment/header-footer RTL options rather than silently ignoring them.
- Deliverable: operation runs standalone against an Arabic sample doc, produces visibly correct RTL paragraph direction/alignment in body text under both engines, and documents any fallback limitations for python-docx.

---

## Phase 6 — Custom TOC Builder Operation

**Goal:** Insert/replace a table of contents with user-controlled formatting and heading-level selection.

- Parameters: heading/style levels to include (multi-select from the doc's discovered styles, not just "Heading 1-9" — must support custom style names since the app lists actual doc styles), TOC title text, tab leader style, whether to show page numbers, whether to make entries hyperlinked, right-align page numbers, formatting (font/size) for the TOC text itself, insertion location (cursor/bookmark/start of doc).
- COM path (primary): `Document.TablesOfContents.Add(Range, UseHeadingStyles, UpperHeadingLevel, LowerHeadingLevel, ...)`, then post-process formatting via the TOC field's `Range.Font`/paragraph formatting; handle replacing an existing TOC vs. inserting new.
- **Decision (confirmed): no python-docx fallback.** If COM/Word is unavailable, the TOC operation is disabled and reports a clear "TOC generation requires Microsoft Word — unavailable on this system" message; no file is touched. This keeps TOC generation consistent with how other COM-only features (regex, comments, headers/footers) are already gated, and avoids producing an unverifiable field-code TOC that only resolves correctly once manually opened and updated in Word.
- Style-level selection UI reuses the style list discovered in Phase 2/7, so it's always accurate to the loaded doc, including mapping custom paragraph styles to an "include in TOC" flag with a chosen indent/level.
- Deliverable: operation runs standalone, produces a doc with a correctly leveled, formatted TOC under COM; documented fallback behavior under python-docx.

---

## Phase 7 — Main GUI Shell & File/Folder Selection

**Goal:** Wire Phases 2–6 into an actual usable window.

- Tab bar: **Pipeline** (main working tab) and **Config** (per requirement #4).
- File selection: single-file picker for the target `.docx` (per requirement — "a single MS Word document"), plus a folder picker for default working/output directory, both remembered in config between runs.
- On file load: open via the resolved engine, populate the style list, enable/disable operation types based on engine capabilities, show engine-in-use ("Word COM" vs "python-docx fallback") in the status bar.
- **Decision (confirmed): save path is requested upfront**, at the start of the workflow alongside the input file selection — not deferred to end-of-run. The source file is never silently overwritten; the user picks the destination path before the pipeline runs.
- Deliverable: can load a real doc, see its styles listed, see which engine is active.

---

## Phase 8 — Pipeline Tab: Add/Reorder/Remove Operations at Runtime

**Goal:** Satisfy requirement #2 directly in the UI.

- "Add Operation" control listing registered operation types (from the Phase 3 registry) → opens that operation's parameter form (modal or inline panel).
- List view of the currently configured pipeline, each entry showing a summary (e.g., "Find/Replace: 'foo' → 'bar' (regex)"), with edit / remove / reorder (drag or up/down buttons) controls.
- "Run Pipeline" button executes the full ordered list via the Phase 3 `Pipeline.run`, on a background `QThread` so the UI doesn't freeze — mirrors the reference app's worker-thread pattern.
- Progress feedback: per-operation status as it completes (running → ok/warning/error), not just a single end-of-run result.
- Deliverable: user can build a 2+ operation pipeline (e.g., a find/replace, an Arabic RTL normalize step, then a TOC insert) entirely through the GUI and run it.

---

## Phase 9 — Error Reporting (requirement #3)

**Goal:** Errors are visible, actionable, and don't silently corrupt the document.

- Structured log panel (persistent across the session) showing timestamped entries per operation: started / succeeded / warning / failed, with exception details expandable (reuse reference app's `traceback` handling).
- End-of-run summary dialog: N operations succeeded, N warnings, N failed, with a "save log to file" option.
- Safety behavior: since the save path is chosen upfront (Phase 7), the source document is never at risk of silent overwrite from the pipeline itself; on failure, the log/summary still makes clear whether any partial writes occurred before the error.
- Distinguish engine-level failures (Word crashed, COM dispatch failed, file locked/open elsewhere) from operation-level failures (bad regex, style not found) in the message shown to the user.
- Deliverable: intentionally-broken test operations (bad regex, missing style) produce clear, non-crashing error reports.

---

## Phase 10 — Config Tab (requirement #4, config-specific content)

**Goal:** A dedicated place for app-wide settings, separate from per-run pipeline setup.

- Engine preference: auto (COM if available, else python-docx) / force COM / force python-docx-only.
- Default folders: last-used input folder, default output/save folder.
- Error-handling policy default: stop-on-first-error vs. continue-and-report-all.
- Logging verbosity / log file location.
- Persisted to `batch_edit_config.json`, loaded on startup, editable and saved from this tab.
- Deliverable: settings survive an app restart.

---

## Phase 11 — Style Listing Integration Polish (requirement #6)

**Goal:** Make sure style discovery is robust, not just a Phase 2 proof-of-concept.

- Refresh style list automatically whenever a new document is loaded.
- Distinguish style types (paragraph / character / table / list) since only paragraph styles are generally relevant to headings/TOC — surface this distinction in the multi-select used by the TOC operation.
- Handle documents with no custom styles gracefully (built-ins only) and documents where COM/python-docx style enumeration disagrees (rare, but worth a defensive check given the dual-engine design).
- Deliverable: style list is accurate and correctly filtered across a handful of real-world test documents with varied style sets.

---

## Phase 12 — Testing & Hardening

- Unit-level: engine abstraction methods, operation `validate()`/`run()` logic, registry registration — testable without a real Word instance where possible (mock COM calls).
- Integration-level: full pipelines against a small suite of representative sample `.docx` files (simple doc, doc with existing TOC, doc with custom styles, doc with headers/footers/comments, protected doc).
- Manual test matrix across the two engines × all operation types × the capability-gating rules from Phases 4/5/6.
- Edge cases to explicitly test: file open in Word already (COM conflict), file path with spaces/unicode, empty document, Arabic-only document with mixed punctuation, and document with no headings at all (TOC operation on nothing to include).

---

## Phase 13 — Packaging & Documentation

- Confirm final `requirements.txt` (or equivalent) — mirroring the reference app's practical install set (`PyQt5`, `pywin32`, `python-docx`) plus anything new introduced.
- Write a short README covering the launch command, engine requirements (Windows + Word for full functionality, exactly like the reference app), and how to add a new operation type to the registry for future extension.
- Optional: packaging as a standalone executable (e.g., PyInstaller) if distribution beyond the dev venv is ever needed — flagged as optional/future, not in scope unless you ask for it.

---

## Suggested Build Order Summary

1 → 2 → 3 are foundational and unavoidable in sequence (can't build operations without the engine layer, can't build the pipeline without operations existing).
4, 5, and 6 can be developed in parallel once Phase 3 is done, since they're independent operation implementations.
7 and 8 depend on 4/5/6 having at least one working operation to test against.
9, 10, 11 layer on top of 7/8 and can be interleaved rather than strictly sequential.
12 and 13 close out the project.

---

## Decisions confirmed

1. **TOC fallback**: disabled entirely without Word/COM — no best-effort python-docx field-code path.
2. **Save behavior**: the app asks for the save path upfront, alongside input file selection, before the pipeline runs.
3. **Scope**: no plans to extend beyond docx-level operations — the engine/operation interfaces don't need to be built generic enough for other file formats (e.g. PDF export).

With these settled, the plan above is ready to move into implementation starting at Phase 1.
