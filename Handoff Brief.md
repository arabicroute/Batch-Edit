I’m pulling the exact current code references for the add-operation workflow so I can give you a clean handoff brief that a co-programmer can implement without re-discovering the UI flow.
**Preparing task handoff**

It seems like the user doesn't need any code changes. They want to hand off the task to a co-programmer by explaining the objectives and sharing the current code related to that workflow. I'll create a concise handoff brief that references the existing code. 

I think it’s essential to inspect the relevant lines around operations, possibly including file references, acceptance criteria, and a suggested implementation plan. It's important to create clickable code references using absolute paths for clarity and convenience. Let's make sure everything is well-organized!

**Creating a task brief**

I think I could generate a brief that’s ready to send. It might include sample operation fields like Find/Replace, Arabic RTL Normalize, and Custom TOC Builder based on schemas. I want to explain the dialog in plain English and highlight the form-driven editors. 

I should recommend no direct JSON editing by default, with advanced options hidden. Including examples of plain English descriptions for operations and acceptance criteria seems essential. Additionally, I can mention the workflow related to adding operations and share specific code references for clarity on the relevant lines in the codebase.



**Crafting developer handoff**

I'm focusing on creating a handoff for the user to send to a co-programmer. They need to explain objectives and share the current code for their workflow. I'll provide a brief that includes sections like "Current UX problem," "Desired UX," and "Current code hotspots." I'll also mention "Operation-specific form requirements" and "Acceptance criteria" while suggesting an implementation approach. Importantly, I’ll avoid direct JSON interaction mentions, preserving the backend contract of parameters to validate existing operations. Let's put together the details along with code reference links.


**Handoff Brief**
- Here is a ready-to-send handoff for a co-programmer to replace the current JSON-heavy operation setup flow with a guided form-based workflow.
- Goal: keep the existing backend contract of `display_name + params dict`, but make the UI user-friendly and typo-resistant.

**Task Objective**
- Replace the current raw JSON parameter editing workflow with a guided, form-based operation builder.
- Keep the existing operation execution backend intact:
  - operation registration
  - `params_schema`
  - `OperationRegistry.create(...)`
  - `operation.validate(...)`
  - configured pipeline storage as `{"display_name": ..., "params": ...}`
- Improve the `Add Operation` workflow in two stages:
  - Stage 1: operation picker dialog should explain each available operation in plain English, with examples.
  - Stage 2: selecting an operation should open a custom form UI for that operation, not a JSON editor.

**Why This Change Is Needed**
- The current UX requires the user to type JSON manually.
- That is too technical for normal users and invites typos, invalid keys, wrong types, and avoidable validation errors.
- The backend is already good enough; the main gap is the input experience.

**Current Workflow**
- `Add Operation` currently opens a simple `QInputDialog` list, then sends the user to a generic JSON editor dialog.
- The current JSON editor is `OperationParamsDialog` in [tabs_pipeline.py](file:///E:/laragon/www/word-find-replace/ui/tabs_pipeline.py#L10-L51).
- The add/edit flow is in [main_window.py](file:///E:/laragon/www/word-find-replace/ui/main_window.py#L517-L590):
  - [_edit_operation_params()](file:///E:/laragon/www/word-find-replace/ui/main_window.py#L517-L540)
  - [add_operation_from_ui()](file:///E:/laragon/www/word-find-replace/ui/main_window.py#L542-L573)
  - [edit_selected_operation()](file:///E:/laragon/www/word-find-replace/ui/main_window.py#L575-L590)
- The configured pipeline is still built correctly downstream in [run_pipeline_from_ui()](file:///E:/laragon/www/word-find-replace/ui/main_window.py#L620-L694), and that should remain unchanged.

**Current Code To Review**
- Generic JSON dialog:
  - [OperationParamsDialog](file:///E:/laragon/www/word-find-replace/ui/tabs_pipeline.py#L10-L51)
- Add/edit operation entry points:
  - [_edit_operation_params()](file:///E:/laragon/www/word-find-replace/ui/main_window.py#L517-L540)
  - [add_operation_from_ui()](file:///E:/laragon/www/word-find-replace/ui/main_window.py#L542-L573)
  - [edit_selected_operation()](file:///E:/laragon/www/word-find-replace/ui/main_window.py#L575-L590)
- Operation schemas and validation rules:
  - [FindReplaceOperation.params_schema](file:///E:/laragon/www/word-find-replace/operations/find_replace.py#L226-L240)
  - [ArabicRtlNormalizeOperation.params_schema](file:///E:/laragon/www/word-find-replace/operations/rtl_normalize.py#L176-L188)
  - [TocBuilderOperation.params_schema](file:///E:/laragon/www/word-find-replace/operations/toc_builder.py#L129-L148)

**Requested UX**
- When the user clicks `Add Operation`:
  - show a richer dialog listing available operations
  - each item should include a plain-English description
  - each item should include a small example
- When the user selects an operation:
  - open a dedicated helper form
  - use text boxes for strings
  - use checkboxes/toggles for booleans
  - use dropdowns for enums/presets
  - use numeric spinboxes for numbers
  - use style pickers for style names
  - use color/font/alignment pickers where relevant
- The form should build the JSON `params` behind the scenes.
- The user should not need to see or edit raw JSON during normal use.

**Operation-Specific Form Expectations**
- `Find/Replace`
  - fields for `find_text`, `replace_text`
  - scope checkboxes: `body`, `comments`, `headers/footers`
  - toggles for `match_case`, `whole_word`, `use_regex`, `use_wildcards`
  - formatting sections for source/target formatting, ideally collapsed/optional
  - hide or disable unsupported combinations based on current engine
- `Arabic RTL Normalize`
  - scope checkboxes
  - alignment dropdown: `right`, `center`, `left`
  - direction fixed to `rtl` or shown as a locked field
  - font picker/text field
  - toggles for `arabic_only`, `normalize_tables`, `normalize_lists`
  - disable unsupported scope options under `python-docx`
- `Custom TOC Builder`
  - style-level selector using paragraph styles only
  - allow multi-row style mapping: style name + level
  - title text field
  - tab leader dropdown
  - page number/hyperlink toggles
  - font/size fields for title and TOC text
  - insertion location dropdown
  - bookmark field only when `bookmark` is selected
  - if current engine is `python-docx`, this operation should remain unavailable as it is now

**Implementation Guidance**
- Keep the current backend objects and validation flow.
- Replace the generic editor with:
  - a richer operation picker dialog
  - per-operation form widgets/dialogs
- Suggested structure:
  - `OperationPickerDialog`
  - `FindReplaceFormDialog`
  - `ArabicRtlNormalizeFormDialog`
  - `TocBuilderFormDialog`
- Optional but recommended:
  - keep an `Advanced JSON` collapsible section for debugging only
  - not visible by default
- Reuse existing main-window helpers:
  - current styles from `self.current_style_catalog`
  - current engine from `self.current_engine`
  - existing template generation in [_build_operation_template()](file:///E:/laragon/www/word-find-replace/ui/main_window.py#L327-L375)
- Preserve existing output format:
  - `params` must still be a dict passed into `OperationRegistry.create(display_name, params)`

**Plain-English Operation Copy To Add**
- `Find/Replace`
  - “Find text and replace it throughout the document. Can target body text, comments, and headers/footers depending on engine support.”
  - Example: “Replace `Acme Inc.` with `Acme Corporation` in the document body.”
- `Arabic RTL Normalize`
  - “Normalize Arabic content for right-to-left reading by setting alignment/direction and optional Arabic-friendly formatting.”
  - Example: “Set Arabic paragraphs to right-aligned RTL and apply an Arabic font.”
- `Custom TOC Builder`
  - “Insert or replace a table of contents using selected heading or paragraph styles.”
  - Example: “Build a TOC from `Heading 1` and `Heading 2` at the start of the document.”

**Acceptance Criteria**
- No raw JSON is required for normal add/edit operation workflows.
- User can add each operation type using forms only.
- Form submission produces the same `params` structure the backend expects.
- Existing pipeline execution continues to work without backend changes.
- Validation errors are shown in user language at form level, not as generic JSON mistakes.
- TOC style mapping uses paragraph styles from the loaded document.
- Unsupported options are disabled or hidden based on current engine capability.
- Existing configured pipeline add/edit/remove/reorder/run behavior continues to function.

**Non-Goals**
- Do not redesign the pipeline execution backend.
- Do not change operation semantics unless required for UI compatibility.
- Do not remove the existing validation logic inside operation classes.

**Recommended First Steps**
- Implement a richer `OperationPickerDialog` first.
- Replace `_edit_operation_params()` so it routes by operation name to specialized dialogs.
- Start with the three current operation types only.
- Leave a small fallback path for unknown future operations, but not as the default UX.

If you want, I can also turn this into a compact developer ticket format with:
- background
- scope
- tasks
- acceptance criteria
- file checklist