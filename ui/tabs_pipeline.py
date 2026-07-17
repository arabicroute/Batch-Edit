from __future__ import annotations

import json

from PyQt5 import QtWidgets

from ui.widgets_common import LogPanel, PathPickerRow


class OperationParamsDialog(QtWidgets.QDialog):
    def __init__(
        self,
        operation_name: str,
        schema: dict,
        initial_params: dict,
        helper_text: str,
        parent: QtWidgets.QWidget | None = None,
    ) -> None:
        super().__init__(parent)
        self.setWindowTitle(f"{operation_name} Parameters")
        self.resize(720, 560)

        layout = QtWidgets.QVBoxLayout(self)
        info = QtWidgets.QLabel(helper_text)
        info.setWordWrap(True)
        self.editor = QtWidgets.QPlainTextEdit()
        self.editor.setPlainText(json.dumps(initial_params, indent=2, ensure_ascii=False))
        self.schema_view = QtWidgets.QPlainTextEdit()
        self.schema_view.setReadOnly(True)
        self.schema_view.setPlainText(json.dumps(schema, indent=2, ensure_ascii=False))

        tabs = QtWidgets.QTabWidget()
        tabs.addTab(self.editor, "Parameters")
        tabs.addTab(self.schema_view, "Schema")

        buttons = QtWidgets.QDialogButtonBox(
            QtWidgets.QDialogButtonBox.Ok | QtWidgets.QDialogButtonBox.Cancel
        )
        buttons.accepted.connect(self.accept)
        buttons.rejected.connect(self.reject)

        layout.addWidget(info)
        layout.addWidget(tabs, 1)
        layout.addWidget(buttons)

    def get_params(self) -> dict:
        text = self.editor.toPlainText().strip() or "{}"
        value = json.loads(text)
        if not isinstance(value, dict):
            raise ValueError("Operation parameters must be a JSON object.")
        return value


class PipelineTab(QtWidgets.QWidget):
    def __init__(self, parent: QtWidgets.QWidget | None = None) -> None:
        super().__init__(parent)
        layout = QtWidgets.QVBoxLayout(self)

        intro = QtWidgets.QLabel(
            "Build and run an ordered list of document operations here."
        )
        intro.setWordWrap(True)

        doc_group = QtWidgets.QGroupBox("Document")
        doc_layout = QtWidgets.QVBoxLayout(doc_group)
        self.input_row = PathPickerRow("Input .docx:")
        self.output_row = PathPickerRow("Save as:")
        self.load_button = QtWidgets.QPushButton("Load Document")
        self.clear_button = QtWidgets.QPushButton("Clear")

        button_row = QtWidgets.QHBoxLayout()
        button_row.addWidget(self.load_button)
        button_row.addWidget(self.clear_button)
        button_row.addStretch(1)

        self.active_engine_value = QtWidgets.QLabel("Not loaded")
        self.engine_detail_value = QtWidgets.QLabel("Choose a document to inspect styles and operation availability.")
        self.engine_detail_value.setWordWrap(True)

        doc_layout.addWidget(self.input_row)
        doc_layout.addWidget(self.output_row)
        doc_layout.addLayout(button_row)
        doc_layout.addWidget(self.active_engine_value)
        doc_layout.addWidget(self.engine_detail_value)

        lists_layout = QtWidgets.QHBoxLayout()

        styles_group = QtWidgets.QGroupBox("Discovered Styles")
        styles_layout = QtWidgets.QVBoxLayout(styles_group)
        self.styles_summary = QtWidgets.QLabel("No document loaded.")
        self.styles_detail = QtWidgets.QLabel("")
        self.styles_detail.setWordWrap(True)
        style_filters = QtWidgets.QHBoxLayout()
        self.style_type_filter = QtWidgets.QComboBox()
        self.style_type_filter.addItem("All styles", "all")
        self.style_type_filter.addItem("Paragraph", "paragraph")
        self.style_type_filter.addItem("Character", "character")
        self.style_type_filter.addItem("Table", "table")
        self.style_type_filter.addItem("List", "list")
        self.style_type_filter.addItem("Unknown", "unknown")
        self.style_search_edit = QtWidgets.QLineEdit()
        self.style_search_edit.setPlaceholderText("Filter styles by name...")
        self.toc_style_hint = QtWidgets.QLabel("")
        self.toc_style_hint.setWordWrap(True)
        self.styles_list = QtWidgets.QListWidget()
        style_filters.addWidget(self.style_type_filter)
        style_filters.addWidget(self.style_search_edit, 1)
        styles_layout.addWidget(self.styles_summary)
        styles_layout.addWidget(self.styles_detail)
        styles_layout.addLayout(style_filters)
        styles_layout.addWidget(self.toc_style_hint)
        styles_layout.addWidget(self.styles_list, 1)

        ops_group = QtWidgets.QGroupBox("Operation Availability")
        ops_layout = QtWidgets.QVBoxLayout(ops_group)
        self.operations_summary = QtWidgets.QLabel("Load a document to evaluate registered operations.")
        self.operations_list = QtWidgets.QListWidget()
        ops_layout.addWidget(self.operations_summary)
        ops_layout.addWidget(self.operations_list, 1)

        lists_layout.addWidget(styles_group, 1)
        lists_layout.addWidget(ops_group, 1)

        pipeline_group = QtWidgets.QGroupBox("Configured Pipeline")
        pipeline_layout = QtWidgets.QVBoxLayout(pipeline_group)
        self.pipeline_summary = QtWidgets.QLabel("No operations configured.")
        self.configured_operations_list = QtWidgets.QListWidget()
        self.configured_operations_list.setSelectionMode(
            QtWidgets.QAbstractItemView.SingleSelection
        )

        controls_row = QtWidgets.QHBoxLayout()
        self.add_operation_button = QtWidgets.QPushButton("Add Operation")
        self.edit_operation_button = QtWidgets.QPushButton("Edit")
        self.remove_operation_button = QtWidgets.QPushButton("Remove")
        self.move_up_button = QtWidgets.QPushButton("Move Up")
        self.move_down_button = QtWidgets.QPushButton("Move Down")
        self.run_pipeline_button = QtWidgets.QPushButton("Run Pipeline")
        controls_row.addWidget(self.add_operation_button)
        controls_row.addWidget(self.edit_operation_button)
        controls_row.addWidget(self.remove_operation_button)
        controls_row.addWidget(self.move_up_button)
        controls_row.addWidget(self.move_down_button)
        controls_row.addStretch(1)
        controls_row.addWidget(self.run_pipeline_button)

        pipeline_layout.addWidget(self.pipeline_summary)
        pipeline_layout.addWidget(self.configured_operations_list, 1)
        pipeline_layout.addLayout(controls_row)

        self.log_panel = LogPanel()
        self.log_panel.append_line("Pipeline log initialized.")
        self._style_rows: list[dict] = []

        layout.addWidget(intro)
        layout.addWidget(doc_group)
        layout.addLayout(lists_layout, 1)
        layout.addWidget(pipeline_group, 1)
        layout.addWidget(self.log_panel, 1)

        self.set_pipeline_running(False)

    def set_document_paths(self, input_path: str, output_path: str) -> None:
        self.input_row.set_text(input_path)
        self.output_row.set_text(output_path)

    def set_engine_status(self, engine_name: str, detail: str) -> None:
        self.active_engine_value.setText(f"Active engine: {engine_name}")
        self.engine_detail_value.setText(detail)

    def set_styles(
        self,
        style_rows: list[dict],
        summary_text: str = "",
        detail_text: str = "",
        toc_hint_text: str = "",
    ) -> None:
        self._style_rows = style_rows
        self.styles_summary.setText(summary_text or f"Loaded {len(style_rows)} style(s).")
        self.styles_detail.setText(detail_text)
        self.toc_style_hint.setText(toc_hint_text)
        self.apply_style_filters()

    def set_operations(self, operation_rows: list[str]) -> None:
        self.operations_list.clear()
        self.operations_list.addItems(operation_rows)
        self.operations_summary.setText(
            f"Listed {len(operation_rows)} registered operation(s)."
        )

    def reset_document_state(self) -> None:
        self.active_engine_value.setText("Not loaded")
        self.engine_detail_value.setText(
            "Choose a document to inspect styles and operation availability."
        )
        self.style_type_filter.setCurrentIndex(0)
        self.style_search_edit.clear()
        self.styles_list.clear()
        self._style_rows = []
        self.operations_list.clear()
        self.styles_summary.setText("No document loaded.")
        self.styles_detail.setText("")
        self.toc_style_hint.setText("")
        self.operations_summary.setText(
            "Load a document to evaluate registered operations."
        )

    def set_configured_operations(self, rows: list[str]) -> None:
        self.configured_operations_list.clear()
        self.configured_operations_list.addItems(rows)
        if rows:
            self.pipeline_summary.setText(f"Configured {len(rows)} operation(s).")
        else:
            self.pipeline_summary.setText("No operations configured.")

    def selected_operation_index(self) -> int | None:
        row = self.configured_operations_list.currentRow()
        return row if row >= 0 else None

    def set_pipeline_running(self, running: bool) -> None:
        self.load_button.setEnabled(not running)
        self.clear_button.setEnabled(not running)
        self.add_operation_button.setEnabled(not running)
        self.edit_operation_button.setEnabled(not running)
        self.remove_operation_button.setEnabled(not running)
        self.move_up_button.setEnabled(not running)
        self.move_down_button.setEnabled(not running)
        self.run_pipeline_button.setEnabled(not running)

    def apply_style_filters(self) -> None:
        selected_type = self.style_type_filter.currentData()
        search_text = self.style_search_edit.text().strip().casefold()
        rows: list[str] = []
        for item in self._style_rows:
            style_type = item.get("style_type", "unknown")
            display_name = item.get("display_name", "")
            name = item.get("name", "")
            if selected_type != "all" and style_type != selected_type:
                continue
            if search_text and search_text not in name.casefold():
                continue
            rows.append(display_name)
        self.styles_list.clear()
        self.styles_list.addItems(rows)
