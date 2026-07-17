from __future__ import annotations

from PyQt5 import QtWidgets


OPERATION_HELP = {
    "Find/Replace": {
        "summary": "Find text and replace it throughout the document.",
        "details": (
            "Supports body text in all engines, with comments and headers/footers "
            "when the current engine supports them."
        ),
        "example": "Replace 'Acme Inc.' with 'Acme Corporation' in the document body.",
    },
    "Arabic RTL Normalize": {
        "summary": "Normalize Arabic content for right-to-left reading and editing.",
        "details": (
            "Sets RTL-friendly paragraph alignment and related formatting, with a "
            "broader scope when Word COM is available."
        ),
        "example": "Right-align Arabic paragraphs and apply an Arabic-friendly font.",
    },
    "Custom TOC Builder": {
        "summary": "Insert or replace a table of contents using chosen styles.",
        "details": (
            "Builds a TOC from heading styles or custom paragraph styles, with "
            "formatting controls for the title and entries."
        ),
        "example": "Build a TOC from Heading 1 and Heading 2 at the start of the document.",
    },
}


class OperationPickerDialog(QtWidgets.QDialog):
    def __init__(
        self,
        available_operations: list[str],
        required_capabilities_by_name: dict[str, list[str]],
        parent: QtWidgets.QWidget | None = None,
    ) -> None:
        super().__init__(parent)
        self.setWindowTitle("Add Operation")
        self.resize(780, 420)
        self._required_capabilities_by_name = required_capabilities_by_name

        layout = QtWidgets.QVBoxLayout(self)
        intro = QtWidgets.QLabel(
            "Choose an operation to add to the pipeline. Each option shows what it does and a simple example."
        )
        intro.setWordWrap(True)

        splitter = QtWidgets.QSplitter()
        self.operations_list = QtWidgets.QListWidget()
        self.operations_list.addItems(available_operations)
        self.operations_list.currentTextChanged.connect(self._refresh_details)
        self.operations_list.itemDoubleClicked.connect(lambda *_: self.accept())

        details_panel = QtWidgets.QWidget()
        details_layout = QtWidgets.QVBoxLayout(details_panel)
        self.summary_label = QtWidgets.QLabel("Select an operation.")
        self.summary_label.setWordWrap(True)
        self.details_label = QtWidgets.QLabel("")
        self.details_label.setWordWrap(True)
        self.example_label = QtWidgets.QLabel("")
        self.example_label.setWordWrap(True)
        self.capabilities_label = QtWidgets.QLabel("")
        self.capabilities_label.setWordWrap(True)
        details_layout.addWidget(self.summary_label)
        details_layout.addWidget(self.details_label)
        details_layout.addWidget(self.example_label)
        details_layout.addWidget(self.capabilities_label)
        details_layout.addStretch(1)

        splitter.addWidget(self.operations_list)
        splitter.addWidget(details_panel)
        splitter.setStretchFactor(1, 1)

        buttons = QtWidgets.QDialogButtonBox(
            QtWidgets.QDialogButtonBox.Ok | QtWidgets.QDialogButtonBox.Cancel
        )
        buttons.accepted.connect(self.accept)
        buttons.rejected.connect(self.reject)

        layout.addWidget(intro)
        layout.addWidget(splitter, 1)
        layout.addWidget(buttons)

        if available_operations:
            self.operations_list.setCurrentRow(0)

    def _refresh_details(self, display_name: str) -> None:
        info = OPERATION_HELP.get(display_name, {})
        summary = info.get("summary", "No description available yet.")
        details = info.get("details", "")
        example = info.get("example", "")
        capabilities = self._required_capabilities_by_name.get(display_name, [])

        self.summary_label.setText(f"<b>{display_name}</b><br>{summary}")
        self.details_label.setText(details)
        self.example_label.setText(f"<b>Example:</b> {example}" if example else "")
        self.capabilities_label.setText(
            "Required capabilities: "
            + (", ".join(capabilities) if capabilities else "none")
        )

    def selected_operation_name(self) -> str:
        item = self.operations_list.currentItem()
        return item.text() if item is not None else ""
