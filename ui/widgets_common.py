from __future__ import annotations

from PyQt5 import QtWidgets


class PathPickerRow(QtWidgets.QWidget):
    def __init__(
        self,
        label_text: str,
        button_text: str = "Browse...",
        parent: QtWidgets.QWidget | None = None,
    ) -> None:
        super().__init__(parent)
        layout = QtWidgets.QHBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)
        self.label = QtWidgets.QLabel(label_text)
        self.path_edit = QtWidgets.QLineEdit()
        self.browse_button = QtWidgets.QPushButton(button_text)
        layout.addWidget(self.label)
        layout.addWidget(self.path_edit, 1)
        layout.addWidget(self.browse_button)

    def text(self) -> str:
        return self.path_edit.text().strip()

    def set_text(self, value: str) -> None:
        self.path_edit.setText(value)


class LogPanel(QtWidgets.QPlainTextEdit):
    def __init__(self, parent: QtWidgets.QWidget | None = None) -> None:
        super().__init__(parent)
        self.setReadOnly(True)

    def append_line(self, text: str) -> None:
        self.appendPlainText(text)
