from __future__ import annotations

from PyQt5 import QtCore
from PyQt5 import QtWidgets

from ui.widgets_common import PathPickerRow


class ConfigTab(QtWidgets.QWidget):
    def __init__(self, parent: QtWidgets.QWidget | None = None) -> None:
        super().__init__(parent)
        layout = QtWidgets.QVBoxLayout(self)

        intro = QtWidgets.QLabel(
            "Manage app-wide defaults such as engine preference and saved presets."
        )
        intro.setWordWrap(True)

        self.com_label = QtWidgets.QLabel("Word COM: checking...")
        self.docx_label = QtWidgets.QLabel("python-docx: checking...")
        self.active_engine_label = QtWidgets.QLabel("Active engine: none")
        self.summary_label = QtWidgets.QLabel("Adjust startup defaults for document loading.")
        self.summary_label.setWordWrap(True)

        defaults_group = QtWidgets.QGroupBox("Defaults")
        defaults_form = QtWidgets.QFormLayout(defaults_group)
        self.engine_preference_combo = QtWidgets.QComboBox()
        self.engine_preference_combo.addItem("Auto (COM then python-docx)", "auto")
        self.engine_preference_combo.addItem("Force Word COM", "com")
        self.engine_preference_combo.addItem("Force python-docx", "docx")
        defaults_form.addRow("Engine preference:", self.engine_preference_combo)

        self.default_output_row = PathPickerRow("Default output folder:")
        self.default_output_row.browse_button.setText("Choose...")
        defaults_form.addRow(self.default_output_row)

        self.last_input_value = QtWidgets.QLineEdit()
        self.last_input_value.setReadOnly(True)
        defaults_form.addRow("Last input path:", self.last_input_value)

        self.error_policy_combo = QtWidgets.QComboBox()
        self.error_policy_combo.addItem("Stop on first error", "stop_on_error")
        self.error_policy_combo.addItem("Continue and report all", "continue_on_error")
        defaults_form.addRow("Error policy:", self.error_policy_combo)

        logging_group = QtWidgets.QGroupBox("Logging")
        logging_form = QtWidgets.QFormLayout(logging_group)
        self.log_verbosity_combo = QtWidgets.QComboBox()
        self.log_verbosity_combo.addItem("Normal", "normal")
        self.log_verbosity_combo.addItem("Detailed", "detailed")
        self.log_verbosity_combo.addItem("Errors only", "errors_only")
        logging_form.addRow("Verbosity:", self.log_verbosity_combo)

        self.log_file_row = PathPickerRow("Session log file:")
        self.log_file_row.browse_button.setText("Choose...")
        logging_form.addRow(self.log_file_row)

        presets_group = QtWidgets.QGroupBox("Pipeline Presets")
        presets_layout = QtWidgets.QVBoxLayout(presets_group)
        self.presets_summary = QtWidgets.QLabel("No saved presets.")
        self.presets_list = QtWidgets.QListWidget()
        self.presets_list.setSelectionMode(QtWidgets.QAbstractItemView.SingleSelection)
        self.presets_list.setAlternatingRowColors(True)

        preset_buttons = QtWidgets.QHBoxLayout()
        self.save_preset_button = QtWidgets.QPushButton("Save Current Pipeline As...")
        self.load_preset_button = QtWidgets.QPushButton("Load Selected Preset")
        self.delete_preset_button = QtWidgets.QPushButton("Delete Selected Preset")
        preset_buttons.addWidget(self.save_preset_button)
        preset_buttons.addWidget(self.load_preset_button)
        preset_buttons.addWidget(self.delete_preset_button)

        presets_layout.addWidget(self.presets_summary)
        presets_layout.addWidget(self.presets_list, 1)
        presets_layout.addLayout(preset_buttons)

        self.save_button = QtWidgets.QPushButton("Save Settings")

        layout.addWidget(intro)
        layout.addWidget(self.com_label)
        layout.addWidget(self.docx_label)
        layout.addWidget(self.active_engine_label)
        layout.addWidget(self.summary_label)
        layout.addWidget(defaults_group)
        layout.addWidget(logging_group)
        layout.addWidget(presets_group, 1)
        layout.addWidget(self.save_button)
        layout.addStretch(1)

    def update_runtime_status(self, runtime_status) -> None:
        self.com_label.setText(f"Word COM: {runtime_status.com_reason}")
        self.docx_label.setText(f"python-docx: {runtime_status.docx_reason}")

    def load_config(self, config: dict) -> None:
        preference = str(config.get("engine_preference", "auto"))
        index = self.engine_preference_combo.findData(preference)
        if index >= 0:
            self.engine_preference_combo.setCurrentIndex(index)
        self.default_output_row.set_text(str(config.get("default_output_dir", "")))
        self.last_input_value.setText(str(config.get("last_input_path", "")))
        error_index = self.error_policy_combo.findData(str(config.get("error_policy", "stop_on_error")))
        if error_index >= 0:
            self.error_policy_combo.setCurrentIndex(error_index)
        verbosity_index = self.log_verbosity_combo.findData(str(config.get("log_verbosity", "normal")))
        if verbosity_index >= 0:
            self.log_verbosity_combo.setCurrentIndex(verbosity_index)
        self.log_file_row.set_text(str(config.get("log_file", "batch_edit.log")))
        self.load_presets(config.get("saved_operation_presets", []))

    def build_config_update(self) -> dict:
        return {
            "engine_preference": self.engine_preference_combo.currentData(),
            "default_output_dir": self.default_output_row.text(),
            "error_policy": self.error_policy_combo.currentData(),
            "log_verbosity": self.log_verbosity_combo.currentData(),
            "log_file": self.log_file_row.text(),
        }

    def set_active_engine(self, engine_name: str, detail: str = "") -> None:
        text = f"Active engine: {engine_name}"
        if detail:
            text = f"{text} ({detail})"
        self.active_engine_label.setText(text)

    def set_last_input_path(self, value: str) -> None:
        self.last_input_value.setText(value)

    def load_presets(self, presets: list[dict]) -> None:
        self.presets_list.clear()
        for preset in presets:
            name = str(preset.get("name", "")).strip()
            if not name:
                continue
            item = QtWidgets.QListWidgetItem(name)
            item.setData(QtCore.Qt.UserRole, preset)
            self.presets_list.addItem(item)
        count = self.presets_list.count()
        self.presets_summary.setText(
            f"Saved presets: {count}" if count else "No saved presets."
        )

    def selected_preset(self) -> dict | None:
        item = self.presets_list.currentItem()
        if item is None:
            return None
        return item.data(QtCore.Qt.UserRole)
