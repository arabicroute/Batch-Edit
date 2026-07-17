from __future__ import annotations

import json
import re
from datetime import datetime
from pathlib import Path

from PyQt5 import QtWidgets

import operations  # noqa: F401
from config.config_manager import ConfigManager
from core.engines import EngineUnavailableError, get_runtime_status, select_engine
from core.operations import OperationRegistry, OperationValidationError
from core.pipeline import Pipeline, PipelineExecutionWorker
from ui.tabs_config import ConfigTab
from ui.tabs_pipeline import OperationParamsDialog, PipelineTab

HEADING_STYLE_RE = re.compile(r"^heading\s+([1-9])$", re.IGNORECASE)


class MainWindow(QtWidgets.QMainWindow):
    def __init__(
        self,
        app_title: str,
        show_summary_dialog: bool = True,
        config_path: str | Path | None = None,
    ) -> None:
        super().__init__()
        self.setWindowTitle(app_title)
        self.resize(960, 640)
        self.show_summary_dialog = show_summary_dialog

        resolved_config_path = (
            Path(config_path)
            if config_path is not None
            else Path(__file__).resolve().parents[1] / "batch_edit_config.json"
        )
        self.config_manager = ConfigManager(resolved_config_path)
        self.config = self.config_manager.load()
        self.runtime_status = get_runtime_status()
        self.current_engine = None
        self.current_styles = []
        self.current_style_catalog: dict = {}
        self.current_input_path = ""
        self.current_output_path = ""
        self.configured_operations: list[dict] = []
        self.pipeline_worker = None
        self.session_log_entries: list[dict] = []
        self.last_run_summary: dict | None = None
        self.last_failure_payload: dict | None = None

        self._build_ui()
        self._connect_signals()
        self.config_tab.load_config(self.config)
        self._refresh_availability()

    def _build_ui(self) -> None:
        central = QtWidgets.QWidget()
        self.setCentralWidget(central)
        layout = QtWidgets.QVBoxLayout(central)

        self.main_tabs = QtWidgets.QTabWidget()
        self.pipeline_tab = PipelineTab()
        self.config_tab = ConfigTab()
        self.main_tabs.addTab(self.pipeline_tab, "Pipeline")
        self.main_tabs.addTab(self.config_tab, "Config")
        layout.addWidget(self.main_tabs, 1)

        status_bar = self.statusBar()
        status_bar.showMessage("Initializing...")

    def _connect_signals(self) -> None:
        self.pipeline_tab.input_row.browse_button.clicked.connect(self._browse_input_file)
        self.pipeline_tab.output_row.browse_button.clicked.connect(self._browse_output_file)
        self.pipeline_tab.load_button.clicked.connect(self.load_document_from_ui)
        self.pipeline_tab.clear_button.clicked.connect(self.clear_loaded_document)
        self.pipeline_tab.add_operation_button.clicked.connect(self.add_operation_from_ui)
        self.pipeline_tab.edit_operation_button.clicked.connect(self.edit_selected_operation)
        self.pipeline_tab.remove_operation_button.clicked.connect(
            self.remove_selected_operation
        )
        self.pipeline_tab.move_up_button.clicked.connect(
            lambda: self.move_selected_operation(-1)
        )
        self.pipeline_tab.move_down_button.clicked.connect(
            lambda: self.move_selected_operation(1)
        )
        self.pipeline_tab.run_pipeline_button.clicked.connect(self.run_pipeline_from_ui)
        self.pipeline_tab.configured_operations_list.itemDoubleClicked.connect(
            lambda *_: self.edit_selected_operation()
        )
        self.pipeline_tab.style_type_filter.currentIndexChanged.connect(
            self.pipeline_tab.apply_style_filters
        )
        self.pipeline_tab.style_search_edit.textChanged.connect(
            self.pipeline_tab.apply_style_filters
        )
        self.config_tab.default_output_row.browse_button.clicked.connect(
            self._browse_default_output_dir
        )
        self.config_tab.log_file_row.browse_button.clicked.connect(self._browse_log_file)
        self.config_tab.save_button.clicked.connect(self.save_config_from_ui)
        self.config_tab.save_preset_button.clicked.connect(self.save_current_pipeline_preset)
        self.config_tab.load_preset_button.clicked.connect(self.load_selected_preset)
        self.config_tab.delete_preset_button.clicked.connect(self.delete_selected_preset)

    def _refresh_availability(self) -> None:
        self.config_tab.update_runtime_status(self.runtime_status)
        self._refresh_operation_availability(None)
        if self.runtime_status.com_available:
            self.statusBar().showMessage("Ready. Word COM available.")
            return
        if self.runtime_status.docx_available:
            self.statusBar().showMessage(
                f"Ready in fallback mode. {self.runtime_status.com_reason}"
            )
            return
        self.statusBar().showMessage(
            "Startup limited: neither Word COM nor python-docx is available."
        )

    def closeEvent(self, event) -> None:
        try:
            self._close_current_engine()
        finally:
            super().closeEvent(event)

    def _browse_input_file(self) -> None:
        start_dir = self._preferred_input_dir()
        path, _ = QtWidgets.QFileDialog.getOpenFileName(
            self,
            "Select input .docx",
            start_dir,
            "Word Documents (*.docx)",
        )
        if not path:
            return
        self.pipeline_tab.input_row.set_text(path)
        if not self.pipeline_tab.output_row.text():
            self.pipeline_tab.output_row.set_text(self._derive_output_path(path))

    def _browse_output_file(self) -> None:
        start_path = self.pipeline_tab.output_row.text() or self._derive_output_path(
            self.pipeline_tab.input_row.text()
        )
        path, _ = QtWidgets.QFileDialog.getSaveFileName(
            self,
            "Choose save path",
            start_path,
            "Word Documents (*.docx)",
        )
        if path:
            self.pipeline_tab.output_row.set_text(path)

    def _browse_default_output_dir(self) -> None:
        start_dir = self.config_tab.default_output_row.text() or self._preferred_input_dir()
        folder = QtWidgets.QFileDialog.getExistingDirectory(
            self,
            "Choose default output folder",
            start_dir,
        )
        if folder:
            self.config_tab.default_output_row.set_text(folder)

    def _browse_log_file(self) -> None:
        start_path = self.config_tab.log_file_row.text() or self._default_log_path()
        path, _ = QtWidgets.QFileDialog.getSaveFileName(
            self,
            "Choose session log file",
            start_path,
            "Log Files (*.log);;JSON Lines (*.jsonl);;All Files (*.*)",
        )
        if path:
            self.config_tab.log_file_row.set_text(path)

    def _preferred_input_dir(self) -> str:
        last_input = str(self.config.get("last_input_path", ""))
        if last_input and Path(last_input).exists():
            return str(Path(last_input).parent)
        default_output_dir = str(self.config.get("default_output_dir", ""))
        if default_output_dir and Path(default_output_dir).is_dir():
            return default_output_dir
        return str(Path(__file__).resolve().parents[1])

    def _derive_output_path(self, input_path: str) -> str:
        if not input_path:
            default_output_dir = str(self.config.get("default_output_dir", ""))
            if default_output_dir:
                return str(Path(default_output_dir) / "batch_edit_output.docx")
            return ""
        source = Path(input_path)
        output_dir = str(self.config_tab.default_output_row.text() or self.config.get("default_output_dir", ""))
        target_dir = Path(output_dir) if output_dir else source.parent
        return str(target_dir / f"{source.stem}-batch-edit{source.suffix}")

    def save_config_from_ui(self) -> None:
        self.config.update(self.config_tab.build_config_update())
        self.config_manager.save(self.config)
        self._append_session_log("info", "config", "Saved config settings.")
        self.statusBar().showMessage("Settings saved.")

    def clear_loaded_document(self) -> None:
        self._close_current_engine()
        self.current_styles = []
        self.current_style_catalog = {}
        self.current_input_path = ""
        self.current_output_path = ""
        self.pipeline_tab.set_document_paths("", "")
        self.pipeline_tab.reset_document_state()
        self.configured_operations = []
        self._refresh_configured_pipeline()
        self.config_tab.set_active_engine("none")
        self.config_tab.set_last_input_path(str(self.config.get("last_input_path", "")))
        self._refresh_operation_availability(None)
        self._append_session_log("info", "document", "Cleared loaded document.")
        self.statusBar().showMessage("Document cleared.")

    def load_document_from_ui(self) -> None:
        input_path = self.pipeline_tab.input_row.text()
        output_path = self.pipeline_tab.output_row.text() or self._derive_output_path(input_path)
        try:
            self.load_document(input_path=input_path, output_path=output_path)
        except (ValueError, FileNotFoundError, EngineUnavailableError) as exc:
            QtWidgets.QMessageBox.warning(self, "Load Document", str(exc))
        except Exception as exc:  # noqa: BLE001
            QtWidgets.QMessageBox.critical(self, "Load Document", str(exc))

    def load_document(self, input_path: str, output_path: str | None = None) -> None:
        input_path = str(Path(input_path).resolve()) if input_path else ""
        if not input_path:
            raise ValueError("Choose an input .docx file first.")
        if not Path(input_path).is_file():
            raise FileNotFoundError("The selected input file does not exist.")

        resolved_output = output_path or self._derive_output_path(input_path)
        if not resolved_output:
            raise ValueError("Choose a save path before loading the document.")
        resolved_output = str(Path(resolved_output).resolve())
        Path(resolved_output).parent.mkdir(parents=True, exist_ok=True)

        self._close_current_engine()
        preference = self.config_tab.engine_preference_combo.currentData() or "auto"
        engine, reason = select_engine(str(preference))
        try:
            engine.open(input_path)
            styles = engine.list_styles()
        except Exception:
            engine.close()
            raise

        self.current_engine = engine
        self.current_styles = styles
        self.current_style_catalog = self._build_style_catalog(styles)
        self.current_input_path = input_path
        self.current_output_path = resolved_output

        self.pipeline_tab.set_document_paths(input_path, resolved_output)
        self.pipeline_tab.set_engine_status(engine.engine_name, reason)
        self.pipeline_tab.set_styles(
            self.current_style_catalog["rows"],
            summary_text=self.current_style_catalog["summary_text"],
            detail_text=self.current_style_catalog["detail_text"],
            toc_hint_text=self.current_style_catalog["toc_hint_text"],
        )
        self._refresh_operation_availability(engine)
        self._append_session_log(
            "info",
            "document",
            f"Loaded document with {engine.engine_name}: {input_path}",
            {"output_path": resolved_output, "style_count": len(styles)},
        )
        self.config_tab.set_active_engine(engine.engine_name, preference)
        self.config_tab.set_last_input_path(input_path)
        self.statusBar().showMessage(f"Loaded via {engine.engine_name}.")

        self.config.update(self.config_tab.build_config_update())
        self.config["last_input_path"] = input_path
        self.config_manager.save(self.config)
        self._refresh_configured_pipeline()

    def _refresh_operation_availability(self, engine) -> None:
        rows = []
        for display_name in OperationRegistry.list_operations():
            registration = OperationRegistry.get(display_name)
            required = sorted(registration.operation_cls.required_capabilities)
            if engine is None:
                status = "Not evaluated"
            else:
                missing = sorted(set(required) - set(getattr(engine, "capabilities", set())))
                if missing:
                    status = f"Unavailable (missing: {', '.join(missing)})"
                else:
                    status = "Available"
            capability_text = ", ".join(required) if required else "none"
            rows.append(f"{display_name}: {status} | requires {capability_text}")
        self.pipeline_tab.set_operations(rows)

    def _refresh_configured_pipeline(self) -> None:
        rows = []
        for index, item in enumerate(self.configured_operations, start=1):
            display_name = item["display_name"]
            params = item["params"]
            try:
                operation = OperationRegistry.create(display_name, params)
                summary = operation.describe()
            except Exception as exc:  # noqa: BLE001
                summary = f"{display_name}: invalid configuration ({exc})"
            last_status = str(item.get("last_status", "")).strip()
            status_prefix = f"[{last_status}] " if last_status else ""
            rows.append(f"{index}. {status_prefix}{summary}")
        self.pipeline_tab.set_configured_operations(rows)

    def _available_operation_names(self) -> list[str]:
        if self.current_engine is None:
            return []
        names: list[str] = []
        for display_name in OperationRegistry.list_operations():
            registration = OperationRegistry.get(display_name)
            missing = sorted(
                set(registration.operation_cls.required_capabilities)
                - set(getattr(self.current_engine, "capabilities", set()))
            )
            if not missing:
                names.append(display_name)
        return names

    def _build_operation_template(self, display_name: str) -> dict:
        schema = OperationRegistry.get(display_name).operation_cls.params_schema
        template = {}
        for key, meta in schema.items():
            if "default" in meta:
                template[key] = meta["default"]
                continue
            param_type = meta.get("type")
            if param_type == "boolean":
                template[key] = False
            elif param_type == "array":
                template[key] = []
            elif param_type == "object":
                template[key] = {}
            elif meta.get("nullable"):
                template[key] = None
            else:
                template[key] = ""

        if display_name == "Find/Replace":
            template.update(
                {
                    "find_text": "",
                    "replace_text": "",
                    "scope": {"body": True, "comments": False, "headers_footers": False},
                }
            )
        elif display_name == "Arabic RTL Normalize":
            template.update(
                {
                    "scope": {"body": True, "comments": False, "headers_footers": False},
                    "alignment": "right",
                    "direction": "rtl",
                    "arabic_only": True,
                    "normalize_tables": True,
                    "normalize_lists": True,
                }
            )
        elif display_name == "Custom TOC Builder":
            heading_levels = self._recommended_toc_style_levels()
            template.update(
                {
                    "style_levels": heading_levels,
                    "title_text": "Contents",
                    "insertion_location": "start",
                    "replace_existing": True,
                }
            )
        return template

    def _operation_helper_text(self, display_name: str) -> str:
        available_styles = ", ".join(style.name for style in self.current_styles[:20])
        if len(self.current_styles) > 20:
            available_styles = f"{available_styles}, ..."
        if not available_styles:
            available_styles = "No styles loaded."
        if display_name == "Custom TOC Builder":
            recommended = json.dumps(
                self._recommended_toc_style_levels(),
                indent=2,
                ensure_ascii=False,
            )
            paragraph_text = ", ".join(
                style.name
                for style in self.current_style_catalog.get("paragraph_styles", [])[:20]
            ) or "No paragraph styles detected."
            return (
                f"Edit JSON parameters for '{display_name}'. "
                f"Paragraph styles: {paragraph_text}\n"
                f"Recommended style_levels template:\n{recommended}"
            )
        return (
            f"Edit JSON parameters for '{display_name}'. "
            f"Loaded styles include: {available_styles}"
        )

    def _build_style_catalog(self, styles: list) -> dict:
        ordered_styles = sorted(
            styles,
            key=lambda item: (item.style_type, item.name.lower(), item.source.lower()),
        )
        type_counts: dict[str, int] = {}
        paragraph_styles = []
        custom_count = 0
        heading_candidates = []
        rows = []
        seen_names: set[str] = set()
        for style in ordered_styles:
            type_counts[style.style_type] = type_counts.get(style.style_type, 0) + 1
            if not style.builtin:
                custom_count += 1
            if style.style_type == "paragraph":
                paragraph_styles.append(style)
                heading_match = HEADING_STYLE_RE.match(style.name)
                if heading_match:
                    heading_candidates.append(
                        {"style_name": style.name, "level": int(heading_match.group(1))}
                    )
            label_parts = [style.name, f"[{style.style_type}]"]
            if style.builtin:
                label_parts.append("(built-in)")
            rows.append(
                {
                    "name": style.name,
                    "style_type": style.style_type,
                    "builtin": style.builtin,
                    "display_name": " ".join(label_parts),
                }
            )
            seen_names.add(style.name.casefold())

        paragraph_count = len(paragraph_styles)
        custom_paragraph_count = sum(1 for style in paragraph_styles if not style.builtin)
        summary_parts = [
            f"Loaded {len(ordered_styles)} unique style(s)",
            f"paragraph {type_counts.get('paragraph', 0)}",
            f"character {type_counts.get('character', 0)}",
            f"table {type_counts.get('table', 0)}",
            f"list {type_counts.get('list', 0)}",
        ]
        if type_counts.get("unknown", 0):
            summary_parts.append(f"unknown {type_counts['unknown']}")
        summary_text = " | ".join(summary_parts)

        detail_lines = [
            f"Custom styles: {custom_count}",
            f"Built-in only document: {'yes' if custom_count == 0 else 'no'}",
            f"Paragraph styles: {paragraph_count} ({custom_paragraph_count} custom)",
        ]
        if paragraph_count == 0:
            detail_lines.append(
                "No paragraph styles detected. TOC style mapping may be limited."
            )
        elif not heading_candidates:
            detail_lines.append(
                "No built-in heading styles detected. Use paragraph styles manually for TOC mapping."
            )
        detail_text = " | ".join(detail_lines)

        recommended_levels = self._recommended_toc_style_levels_from(paragraph_styles)
        if recommended_levels:
            formatted = ", ".join(
                f"{item['style_name']} -> L{item['level']}" for item in recommended_levels
            )
            toc_hint_text = f"Recommended TOC paragraph styles: {formatted}"
        else:
            toc_hint_text = (
                "No TOC-ready paragraph style recommendations found yet. "
                "Load a document with heading or custom paragraph styles."
            )

        return {
            "rows": rows,
            "summary_text": summary_text,
            "detail_text": detail_text,
            "toc_hint_text": toc_hint_text,
            "paragraph_styles": paragraph_styles,
            "heading_candidates": heading_candidates,
            "type_counts": type_counts,
        }

    def _recommended_toc_style_levels(self) -> list[dict]:
        return self._recommended_toc_style_levels_from(
            self.current_style_catalog.get("paragraph_styles", [])
        )

    def _recommended_toc_style_levels_from(self, paragraph_styles: list) -> list[dict]:
        heading_levels = []
        custom_paragraphs = []
        for style in paragraph_styles:
            match = HEADING_STYLE_RE.match(style.name)
            if match:
                heading_levels.append(
                    {"style_name": style.name, "level": int(match.group(1))}
                )
            elif not style.builtin:
                custom_paragraphs.append(style)
        if heading_levels:
            return sorted(heading_levels, key=lambda item: (item["level"], item["style_name"].lower()))
        if custom_paragraphs:
            return [
                {"style_name": style.name, "level": min(index + 1, 3)}
                for index, style in enumerate(
                    sorted(custom_paragraphs, key=lambda item: item.name.lower())[:3]
                )
            ]
        if paragraph_styles:
            return [{"style_name": paragraph_styles[0].name, "level": 1}]
        return [{"style_name": "Heading 1", "level": 1}]

    def _edit_operation_params(
        self,
        display_name: str,
        initial_params: dict | None = None,
    ) -> dict | None:
        if self.current_engine is None:
            raise ValueError("Load a document before configuring operations.")
        schema = OperationRegistry.get(display_name).operation_cls.params_schema
        dialog = OperationParamsDialog(
            operation_name=display_name,
            schema=schema,
            initial_params=initial_params or self._build_operation_template(display_name),
            helper_text=self._operation_helper_text(display_name),
            parent=self,
        )
        while dialog.exec_() == QtWidgets.QDialog.Accepted:
            try:
                params = dialog.get_params()
                operation = OperationRegistry.create(display_name, params)
                operation.validate(self.current_engine)
                return params
            except (ValueError, OperationValidationError, KeyError) as exc:
                QtWidgets.QMessageBox.warning(self, "Operation Parameters", str(exc))
        return None

    def add_operation_from_ui(self) -> None:
        if self.current_engine is None:
            QtWidgets.QMessageBox.warning(
                self, "Add Operation", "Load a document before adding operations."
            )
            return
        available = self._available_operation_names()
        if not available:
            QtWidgets.QMessageBox.warning(
                self,
                "Add Operation",
                "No operations are available for the currently loaded engine.",
            )
            return
        display_name, ok = QtWidgets.QInputDialog.getItem(
            self,
            "Add Operation",
            "Operation type:",
            available,
            0,
            False,
        )
        if not ok or not display_name:
            return
        params = self._edit_operation_params(display_name)
        if params is None:
            return
        self.configured_operations.append(
            {"display_name": display_name, "params": params, "last_status": ""}
        )
        self._refresh_configured_pipeline()
        self._append_session_log("info", "pipeline", f"Added operation: {display_name}")

    def edit_selected_operation(self) -> None:
        index = self.pipeline_tab.selected_operation_index()
        if index is None:
            QtWidgets.QMessageBox.information(
                self, "Edit Operation", "Select a configured operation first."
            )
            return
        item = self.configured_operations[index]
        params = self._edit_operation_params(item["display_name"], item["params"])
        if params is None:
            return
        item["params"] = params
        self._refresh_configured_pipeline()
        self._append_session_log(
            "info", "pipeline", f"Updated operation: {item['display_name']}"
        )

    def remove_selected_operation(self) -> None:
        index = self.pipeline_tab.selected_operation_index()
        if index is None:
            QtWidgets.QMessageBox.information(
                self, "Remove Operation", "Select a configured operation first."
            )
            return
        removed = self.configured_operations.pop(index)
        self._refresh_configured_pipeline()
        self._append_session_log(
            "info", "pipeline", f"Removed operation: {removed['display_name']}"
        )

    def move_selected_operation(self, delta: int) -> None:
        index = self.pipeline_tab.selected_operation_index()
        if index is None:
            QtWidgets.QMessageBox.information(
                self, "Move Operation", "Select a configured operation first."
            )
            return
        target = index + delta
        if target < 0 or target >= len(self.configured_operations):
            return
        item = self.configured_operations.pop(index)
        self.configured_operations.insert(target, item)
        self._refresh_configured_pipeline()
        self.pipeline_tab.configured_operations_list.setCurrentRow(target)

    def run_pipeline_from_ui(self) -> None:
        if self.current_engine is None or not self.current_input_path:
            QtWidgets.QMessageBox.warning(
                self, "Run Pipeline", "Load a document before running the pipeline."
            )
            return
        if not self.configured_operations:
            QtWidgets.QMessageBox.warning(
                self, "Run Pipeline", "Add at least one operation to the pipeline."
            )
            return
        output_path = self.pipeline_tab.output_row.text()
        if not output_path:
            QtWidgets.QMessageBox.warning(
                self, "Run Pipeline", "Choose a save path before running the pipeline."
            )
            return
        if str(Path(output_path).resolve()) == str(Path(self.current_input_path).resolve()):
            QtWidgets.QMessageBox.warning(
                self,
                "Run Pipeline",
                "Choose a save path different from the source document.",
            )
            return

        continue_on_error = (
            self.config_tab.error_policy_combo.currentData() == "continue_on_error"
        )
        try:
            pipeline = Pipeline(
                operations=[
                    OperationRegistry.create(item["display_name"], item["params"])
                    for item in self.configured_operations
                ],
                continue_on_error=continue_on_error,
            )
        except Exception as exc:  # noqa: BLE001
            QtWidgets.QMessageBox.critical(self, "Run Pipeline", str(exc))
            return

        preference = self.config_tab.engine_preference_combo.currentData() or "auto"

        def engine_factory():
            engine, _ = select_engine(str(preference))
            return engine

        self._close_current_engine()
        for item in self.configured_operations:
            item["last_status"] = ""
        self._refresh_configured_pipeline()
        self.pipeline_tab.set_pipeline_running(True)
        self._append_session_log(
            "info",
            "pipeline",
            "Starting pipeline run.",
            {
                "input_path": self.current_input_path,
                "output_path": output_path,
                "operation_count": len(self.configured_operations),
                "continue_on_error": continue_on_error,
            },
        )
        self.pipeline_worker = PipelineExecutionWorker(
            pipeline=pipeline,
            engine_factory=engine_factory,
            input_path=self.current_input_path,
            output_path=output_path,
        )
        self.pipeline_worker.operation_started.connect(self._on_pipeline_operation_started)
        self.pipeline_worker.operation_finished.connect(
            self._on_pipeline_operation_finished
        )
        self.pipeline_worker.pipeline_finished.connect(self._on_pipeline_finished)
        self.pipeline_worker.pipeline_failed.connect(self._on_pipeline_failed)
        self.pipeline_worker.start()

    def _pipeline_preset_snapshot(self, name: str) -> dict:
        return {
            "name": name,
            "operations": [
                {
                    "display_name": item["display_name"],
                    "params": json.loads(json.dumps(item["params"], ensure_ascii=False)),
                }
                for item in self.configured_operations
            ],
        }

    def _load_preset_into_pipeline(self, preset: dict) -> None:
        operations = preset.get("operations", [])
        if not isinstance(operations, list):
            raise ValueError("Preset operations must be a list.")
        loaded_items = []
        for item in operations:
            if not isinstance(item, dict):
                raise ValueError("Each preset operation must be an object.")
            display_name = str(item.get("display_name", "")).strip()
            if not display_name:
                raise ValueError("Preset operation is missing a display_name.")
            params = item.get("params", {})
            if not isinstance(params, dict):
                raise ValueError("Preset operation params must be a JSON object.")
            loaded_items.append(
                {"display_name": display_name, "params": params, "last_status": ""}
            )
        self.configured_operations = loaded_items
        self._refresh_configured_pipeline()

    def save_pipeline_preset_named(self, preset_name: str) -> tuple[bool, str]:
        preset_name = preset_name.strip()
        if not preset_name:
            raise ValueError("Preset name is required.")
        if not self.configured_operations:
            raise ValueError(
                "Configure at least one pipeline operation before saving a preset."
            )
        presets = list(self.config.get("saved_operation_presets", []))
        preset = self._pipeline_preset_snapshot(preset_name)
        replaced = False
        for index, existing in enumerate(presets):
            if str(existing.get("name", "")).strip().lower() == preset_name.lower():
                presets[index] = preset
                replaced = True
                break
        if not replaced:
            presets.append(preset)
        presets.sort(key=lambda item: str(item.get("name", "")).lower())
        self.config["saved_operation_presets"] = presets
        self.config_manager.save(self.config)
        self.config_tab.load_presets(presets)
        action = "Updated" if replaced else "Saved"
        self._append_session_log("info", "config", f"{action} preset: {preset_name}")
        self.statusBar().showMessage(f"{action} preset '{preset_name}'.")
        return replaced, preset_name

    def save_current_pipeline_preset(self) -> None:
        if not self.configured_operations:
            QtWidgets.QMessageBox.warning(
                self,
                "Save Preset",
                "Configure at least one pipeline operation before saving a preset.",
            )
            return
        default_name = (
            f"{Path(self.current_input_path).stem} preset"
            if self.current_input_path
            else "New preset"
        )
        preset_name, ok = QtWidgets.QInputDialog.getText(
            self,
            "Save Preset",
            "Preset name:",
            text=default_name,
        )
        preset_name = preset_name.strip()
        if not ok or not preset_name:
            return
        try:
            self.save_pipeline_preset_named(preset_name)
        except Exception as exc:  # noqa: BLE001
            QtWidgets.QMessageBox.critical(self, "Save Preset", str(exc))

    def load_selected_preset(self) -> None:
        preset = self.config_tab.selected_preset()
        if preset is None:
            QtWidgets.QMessageBox.information(
                self, "Load Preset", "Select a saved preset first."
            )
            return
        try:
            self._load_preset_into_pipeline(preset)
        except Exception as exc:  # noqa: BLE001
            QtWidgets.QMessageBox.critical(self, "Load Preset", str(exc))
            return
        preset_name = str(preset.get("name", "Preset"))
        self._append_session_log("info", "config", f"Loaded preset: {preset_name}")
        self.statusBar().showMessage(f"Loaded preset '{preset_name}'.")

    def delete_selected_preset(self) -> None:
        preset = self.config_tab.selected_preset()
        if preset is None:
            QtWidgets.QMessageBox.information(
                self, "Delete Preset", "Select a saved preset first."
            )
            return
        preset_name = str(preset.get("name", "")).strip()
        if not preset_name:
            return
        presets = [
            item
            for item in self.config.get("saved_operation_presets", [])
            if str(item.get("name", "")).strip().lower() != preset_name.lower()
        ]
        self.config["saved_operation_presets"] = presets
        self.config_manager.save(self.config)
        self.config_tab.load_presets(presets)
        self._append_session_log("info", "config", f"Deleted preset: {preset_name}")
        self.statusBar().showMessage(f"Deleted preset '{preset_name}'.")

    def _on_pipeline_operation_started(self, index: int, operation) -> None:
        self.configured_operations[index]["last_status"] = "RUNNING"
        self._refresh_configured_pipeline()
        self._append_session_log(
            "info",
            "operation",
            f"Running {index + 1}. {operation.describe()}",
            {"index": index + 1, "operation": operation.name},
        )

    def _on_pipeline_operation_finished(self, index: int, operation, result) -> None:
        status = str(result.status).upper()
        self.configured_operations[index]["last_status"] = status
        level = "error" if result.status == "error" else "warning" if result.status == "warning" else "info"
        self._append_session_log(
            level,
            "operation",
            f"{status}: {operation.describe()} | {result.message}",
            {
                "index": index + 1,
                "operation": operation.name,
                "details": result.details,
            },
        )
        self._refresh_configured_pipeline()

    def _on_pipeline_finished(self, payload: dict) -> None:
        self.pipeline_tab.set_pipeline_running(False)
        saved = bool(payload.get("saved", False))
        output_path = str(payload.get("output_path", ""))
        results = payload.get("results", [])
        summary = self._summarize_results(results, saved=saved, output_path=output_path)
        self.last_run_summary = summary
        self.last_failure_payload = None
        self._append_session_log(
            "info",
            "pipeline",
            f"Pipeline finished. Saved={saved}. Results={len(results)}.",
            summary,
        )
        self.statusBar().showMessage(
            "Pipeline completed and saved." if saved else "Pipeline stopped before save."
        )
        if self.show_summary_dialog:
            self._show_run_summary("Pipeline Summary", summary)
        self.pipeline_worker = None
        if self.current_input_path:
            try:
                self.load_document(self.current_input_path, self.current_output_path)
            except Exception as exc:  # noqa: BLE001
                self._append_session_log(
                    "warning",
                    "engine",
                    f"Reload after pipeline run failed: {exc}",
                )

    def _on_pipeline_failed(self, payload: dict) -> None:
        self.pipeline_tab.set_pipeline_running(False)
        category = str(payload.get("category", "unexpected"))
        stage = str(payload.get("stage", "run"))
        message = str(payload.get("message", "Unknown failure"))
        results = payload.get("results", [])
        level = "error"
        self._append_session_log(
            level,
            category,
            f"{category.upper()} FAILURE during {stage}: {message}",
            payload,
        )
        self.statusBar().showMessage("Pipeline failed.")
        summary = self._summarize_results(
            results,
            saved=False,
            output_path=str(payload.get("output_path", "")),
            engine_failure={"category": category, "stage": stage, "message": message},
        )
        self.last_run_summary = summary
        self.last_failure_payload = payload
        if self.show_summary_dialog:
            self._show_run_summary("Pipeline Failure", summary, error=True)
        self.pipeline_worker = None
        if self.current_input_path:
            try:
                self.load_document(self.current_input_path, self.current_output_path)
            except Exception as exc:  # noqa: BLE001
                self._append_session_log(
                    "warning",
                    "engine",
                    f"Reload after pipeline failure failed: {exc}",
                )

    def _append_session_log(
        self,
        level: str,
        category: str,
        message: str,
        details: dict | None = None,
    ) -> None:
        timestamp = datetime.now()
        entry = {
            "timestamp": timestamp.isoformat(timespec="seconds"),
            "level": level,
            "category": category,
            "message": message,
            "details": details or {},
        }
        self.session_log_entries.append(entry)
        verbosity = self._current_log_verbosity()
        if verbosity == "errors_only" and level not in {"warning", "error"}:
            return
        time_text = timestamp.strftime("%H:%M:%S")
        line = f"[{time_text}] {category.upper()} {level.upper()}: {message}"
        self.pipeline_tab.log_panel.append_line(line)
        if verbosity == "detailed" and entry["details"]:
            detail_text = json.dumps(entry["details"], ensure_ascii=False, sort_keys=True)
            self.pipeline_tab.log_panel.append_line(f"  details: {detail_text}")

    def _summarize_results(
        self,
        results: list,
        saved: bool,
        output_path: str,
        engine_failure: dict | None = None,
    ) -> dict:
        ok_count = 0
        warning_count = 0
        error_count = 0
        operation_lines: list[str] = []
        for entry in results:
            result = entry.result if hasattr(entry, "result") else entry
            summary = (
                entry.operation_summary if hasattr(entry, "operation_summary") else "Operation"
            )
            status = str(result.status).lower()
            if status == "ok":
                ok_count += 1
            elif status == "warning":
                warning_count += 1
            else:
                error_count += 1
            operation_lines.append(f"{status.upper()}: {summary} | {result.message}")

        return {
            "ok_count": ok_count,
            "warning_count": warning_count,
            "error_count": error_count,
            "saved": saved,
            "output_path": output_path,
            "engine_failure": engine_failure,
            "operation_lines": operation_lines,
        }

    def _summary_text(self, summary: dict) -> str:
        lines = [
            f"Succeeded: {summary['ok_count']}",
            f"Warnings: {summary['warning_count']}",
            f"Failed: {summary['error_count']}",
            f"Saved output: {'Yes' if summary['saved'] else 'No'}",
        ]
        if summary.get("output_path"):
            lines.append(f"Output path: {summary['output_path']}")
        engine_failure = summary.get("engine_failure")
        if engine_failure:
            lines.append("")
            lines.append(
                f"Engine failure: {engine_failure['category']} during {engine_failure['stage']}"
            )
            lines.append(engine_failure["message"])
        if summary.get("operation_lines"):
            lines.append("")
            lines.append("Operation results:")
            lines.extend(summary["operation_lines"])
        return "\n".join(lines)

    def _default_log_path(self) -> str:
        configured = (
            str(self.config_tab.log_file_row.text() or self.config.get("log_file", "batch_edit.log")).strip()
            or "batch_edit.log"
        )
        path = Path(configured)
        if not path.is_absolute():
            path = Path(__file__).resolve().parents[1] / configured
        return str(path)

    def _current_log_verbosity(self) -> str:
        return str(
            self.config_tab.log_verbosity_combo.currentData()
            or self.config.get("log_verbosity", "normal")
        )

    def save_session_log(self, file_path: str | None = None) -> str | None:
        target_path = file_path
        if not target_path:
            target_path, _ = QtWidgets.QFileDialog.getSaveFileName(
                self,
                "Save session log",
                self._default_log_path(),
                "Log Files (*.log);;JSON Lines (*.jsonl);;All Files (*.*)",
            )
        if not target_path:
            return None
        path = Path(target_path)
        path.parent.mkdir(parents=True, exist_ok=True)
        if path.suffix.lower() == ".jsonl":
            lines = [json.dumps(entry, ensure_ascii=False) for entry in self.session_log_entries]
            path.write_text("\n".join(lines) + ("\n" if lines else ""), encoding="utf-8")
        else:
            text_lines = []
            for entry in self.session_log_entries:
                text_lines.append(
                    f"[{entry['timestamp']}] {entry['category'].upper()} {entry['level'].upper()}: {entry['message']}"
                )
                if entry["details"]:
                    text_lines.append(
                        json.dumps(entry["details"], ensure_ascii=False, indent=2)
                    )
            path.write_text("\n".join(text_lines), encoding="utf-8")
        self._append_session_log("info", "log", f"Saved session log to: {path}")
        return str(path)

    def _show_run_summary(self, title: str, summary: dict, error: bool = False) -> None:
        box = QtWidgets.QMessageBox(self)
        box.setWindowTitle(title)
        box.setIcon(
            QtWidgets.QMessageBox.Warning if error or summary.get("error_count") else QtWidgets.QMessageBox.Information
        )
        box.setText(self._summary_text(summary))
        save_button = box.addButton("Save Log...", QtWidgets.QMessageBox.ActionRole)
        box.addButton(QtWidgets.QMessageBox.Close)
        box.exec_()
        if box.clickedButton() == save_button:
            self.save_session_log()

    def _close_current_engine(self) -> None:
        if self.current_engine is None:
            return
        try:
            self.current_engine.close()
        finally:
            self.current_engine = None
