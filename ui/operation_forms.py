from __future__ import annotations

from PyQt5 import QtCore
from PyQt5 import QtWidgets


def _tri_state_combo() -> QtWidgets.QComboBox:
    combo = QtWidgets.QComboBox()
    combo.addItem("Leave unchanged", None)
    combo.addItem("Yes", True)
    combo.addItem("No", False)
    return combo


def _set_combo_data(combo: QtWidgets.QComboBox, value) -> None:
    index = combo.findData(value)
    if index >= 0:
        combo.setCurrentIndex(index)


def _combo_value(combo: QtWidgets.QComboBox):
    return combo.currentData()


class _BaseOperationFormDialog(QtWidgets.QDialog):
    def __init__(self, title: str, parent: QtWidgets.QWidget | None = None) -> None:
        super().__init__(parent)
        self.setWindowTitle(title)
        self.resize(760, 620)
        self._layout = QtWidgets.QVBoxLayout(self)
        self.form = QtWidgets.QFormLayout()
        self.form.setFieldGrowthPolicy(QtWidgets.QFormLayout.ExpandingFieldsGrow)
        self._layout.addLayout(self.form)
        self._layout.addStretch(1)

        buttons = QtWidgets.QDialogButtonBox(
            QtWidgets.QDialogButtonBox.Ok | QtWidgets.QDialogButtonBox.Cancel
        )
        buttons.accepted.connect(self.accept)
        buttons.rejected.connect(self.reject)
        self._layout.addWidget(buttons)

    def get_params(self) -> dict:
        raise NotImplementedError


class FindReplaceFormDialog(_BaseOperationFormDialog):
    def __init__(
        self,
        initial_params: dict,
        is_docx_engine: bool,
        engine_capabilities: set[str],
        parent: QtWidgets.QWidget | None = None,
    ) -> None:
        super().__init__("Find/Replace", parent)
        self._is_docx_engine = is_docx_engine
        self._engine_capabilities = engine_capabilities

        intro = QtWidgets.QLabel(
            "Find text and replace it throughout the document. Use the options below instead of typing JSON manually."
        )
        intro.setWordWrap(True)
        self._layout.insertWidget(0, intro)

        self.find_text_edit = QtWidgets.QLineEdit()
        self.replace_text_edit = QtWidgets.QLineEdit()
        self.form.addRow("Find text:", self.find_text_edit)
        self.form.addRow("Replace with:", self.replace_text_edit)

        scope_box = QtWidgets.QGroupBox("Scope")
        scope_layout = QtWidgets.QVBoxLayout(scope_box)
        self.scope_body = QtWidgets.QCheckBox("Body")
        self.scope_comments = QtWidgets.QCheckBox("Comments")
        self.scope_headers = QtWidgets.QCheckBox("Headers / footers")
        scope_layout.addWidget(self.scope_body)
        scope_layout.addWidget(self.scope_comments)
        scope_layout.addWidget(self.scope_headers)
        self.form.addRow(scope_box)

        options_box = QtWidgets.QGroupBox("Matching options")
        options_layout = QtWidgets.QVBoxLayout(options_box)
        self.match_case_check = QtWidgets.QCheckBox("Match case")
        self.whole_word_check = QtWidgets.QCheckBox("Whole word")
        self.regex_check = QtWidgets.QCheckBox("Use regex")
        self.wildcards_check = QtWidgets.QCheckBox("Use Word wildcards")
        self.regex_check.toggled.connect(self._enforce_mutual_exclusion)
        self.wildcards_check.toggled.connect(self._enforce_mutual_exclusion)
        options_layout.addWidget(self.match_case_check)
        options_layout.addWidget(self.whole_word_check)
        options_layout.addWidget(self.regex_check)
        options_layout.addWidget(self.wildcards_check)
        self.form.addRow(options_box)

        self.match_format_group = self._build_format_group("Only replace text that already has specific formatting")
        self.target_format_group = self._build_format_group("Apply formatting to the replacement text")
        self.form.addRow(self.match_format_group)
        self.form.addRow(self.target_format_group)

        self._apply_engine_gating()
        self._load_initial(initial_params)

    def _build_format_group(self, title: str) -> QtWidgets.QGroupBox:
        group = QtWidgets.QGroupBox(title)
        group.setCheckable(True)
        group.setChecked(False)
        layout = QtWidgets.QFormLayout(group)

        font_name = QtWidgets.QLineEdit()
        size_spin = QtWidgets.QDoubleSpinBox()
        size_spin.setRange(0.0, 200.0)
        size_spin.setDecimals(1)
        size_spin.setSpecialValueText("Leave unchanged")
        size_spin.setValue(0.0)
        bold_combo = _tri_state_combo()
        italic_combo = _tri_state_combo()
        underline_combo = _tri_state_combo()
        color_edit = QtWidgets.QLineEdit()
        color_edit.setPlaceholderText("#rrggbb")
        highlight_combo = QtWidgets.QComboBox()
        highlight_combo.addItem("Leave unchanged", "")
        for value in ["yellow", "green", "turquoise", "pink", "blue", "red", "teal", "violet"]:
            highlight_combo.addItem(value.title(), value)

        layout.addRow("Font:", font_name)
        layout.addRow("Size:", size_spin)
        layout.addRow("Bold:", bold_combo)
        layout.addRow("Italic:", italic_combo)
        layout.addRow("Underline:", underline_combo)
        layout.addRow("Color:", color_edit)
        layout.addRow("Highlight:", highlight_combo)

        group._font_name = font_name  # type: ignore[attr-defined]
        group._size_spin = size_spin  # type: ignore[attr-defined]
        group._bold_combo = bold_combo  # type: ignore[attr-defined]
        group._italic_combo = italic_combo  # type: ignore[attr-defined]
        group._underline_combo = underline_combo  # type: ignore[attr-defined]
        group._color_edit = color_edit  # type: ignore[attr-defined]
        group._highlight_combo = highlight_combo  # type: ignore[attr-defined]
        return group

    def _apply_engine_gating(self) -> None:
        if self._is_docx_engine:
            self.scope_comments.setChecked(False)
            self.scope_comments.setEnabled(False)
            self.scope_headers.setChecked(False)
            self.scope_headers.setEnabled(False)
            self.regex_check.setChecked(False)
            self.regex_check.setEnabled(False)
            self.wildcards_check.setChecked(False)
            self.wildcards_check.setEnabled(False)
            self.match_format_group.setChecked(False)
            self.match_format_group.setEnabled(False)
        else:
            if "comments" not in self._engine_capabilities:
                self.scope_comments.setChecked(False)
                self.scope_comments.setEnabled(False)
            if "headers_footers" not in self._engine_capabilities:
                self.scope_headers.setChecked(False)
                self.scope_headers.setEnabled(False)
            if "find_replace_advanced" not in self._engine_capabilities:
                self.regex_check.setChecked(False)
                self.regex_check.setEnabled(False)
                self.wildcards_check.setChecked(False)
                self.wildcards_check.setEnabled(False)
                self.match_format_group.setChecked(False)
                self.match_format_group.setEnabled(False)

    def _enforce_mutual_exclusion(self) -> None:
        if self.sender() is self.regex_check and self.regex_check.isChecked():
            self.wildcards_check.setChecked(False)
        if self.sender() is self.wildcards_check and self.wildcards_check.isChecked():
            self.regex_check.setChecked(False)

    def _load_format_group(self, group: QtWidgets.QGroupBox, value: dict | None) -> None:
        if not value:
            group.setChecked(False)
            return
        group.setChecked(True)
        group._font_name.setText(str(value.get("font_name", "")))  # type: ignore[attr-defined]
        group._size_spin.setValue(float(value.get("size_pt") or 0.0))  # type: ignore[attr-defined]
        _set_combo_data(group._bold_combo, value.get("bold"))  # type: ignore[attr-defined]
        _set_combo_data(group._italic_combo, value.get("italic"))  # type: ignore[attr-defined]
        _set_combo_data(group._underline_combo, value.get("underline"))  # type: ignore[attr-defined]
        group._color_edit.setText(str(value.get("color", "")))  # type: ignore[attr-defined]
        _set_combo_data(group._highlight_combo, value.get("highlight", ""))  # type: ignore[attr-defined]

    def _build_format_value(self, group: QtWidgets.QGroupBox) -> dict | None:
        if not group.isChecked():
            return None
        value = {
            "font_name": group._font_name.text().strip() or None,  # type: ignore[attr-defined]
            "size_pt": group._size_spin.value() or None,  # type: ignore[attr-defined]
            "bold": _combo_value(group._bold_combo),  # type: ignore[attr-defined]
            "italic": _combo_value(group._italic_combo),  # type: ignore[attr-defined]
            "underline": _combo_value(group._underline_combo),  # type: ignore[attr-defined]
            "color": group._color_edit.text().strip() or None,  # type: ignore[attr-defined]
            "highlight": group._highlight_combo.currentData() or None,  # type: ignore[attr-defined]
        }
        if any(v is not None for v in value.values()):
            return value
        return {}

    def _load_initial(self, initial_params: dict) -> None:
        scope = initial_params.get("scope", {})
        self.find_text_edit.setText(str(initial_params.get("find_text", "")))
        self.replace_text_edit.setText(str(initial_params.get("replace_text", "")))
        self.scope_body.setChecked(bool(scope.get("body", True)))
        self.scope_comments.setChecked(bool(scope.get("comments", False)))
        self.scope_headers.setChecked(bool(scope.get("headers_footers", False)))
        self.match_case_check.setChecked(bool(initial_params.get("match_case", False)))
        self.whole_word_check.setChecked(bool(initial_params.get("whole_word", False)))
        if self.regex_check.isEnabled():
            self.regex_check.setChecked(bool(initial_params.get("use_regex", False)))
        if self.wildcards_check.isEnabled():
            self.wildcards_check.setChecked(bool(initial_params.get("use_wildcards", False)))
        self._load_format_group(self.match_format_group, initial_params.get("match_format"))
        self._load_format_group(self.target_format_group, initial_params.get("target_format"))
        self._apply_engine_gating()

    def get_params(self) -> dict:
        params = {
            "find_text": self.find_text_edit.text(),
            "replace_text": self.replace_text_edit.text(),
            "scope": {
                "body": self.scope_body.isChecked(),
                "comments": self.scope_comments.isChecked(),
                "headers_footers": self.scope_headers.isChecked(),
            },
            "match_case": self.match_case_check.isChecked(),
            "whole_word": self.whole_word_check.isChecked(),
            "use_regex": self.regex_check.isChecked() if self.regex_check.isEnabled() else False,
            "use_wildcards": self.wildcards_check.isChecked() if self.wildcards_check.isEnabled() else False,
        }
        match_format = self._build_format_value(self.match_format_group)
        if match_format is not None:
            params["match_format"] = match_format
        target_format = self._build_format_value(self.target_format_group)
        if target_format is not None:
            params["target_format"] = target_format
        return params


class ArabicRtlNormalizeFormDialog(_BaseOperationFormDialog):
    def __init__(
        self,
        initial_params: dict,
        is_docx_engine: bool,
        engine_capabilities: set[str],
        parent: QtWidgets.QWidget | None = None,
    ) -> None:
        super().__init__("Arabic RTL Normalize", parent)
        self._is_docx_engine = is_docx_engine
        self._engine_capabilities = engine_capabilities

        intro = QtWidgets.QLabel(
            "Normalize Arabic content for right-to-left reading by choosing scope and formatting options."
        )
        intro.setWordWrap(True)
        self._layout.insertWidget(0, intro)

        scope_box = QtWidgets.QGroupBox("Scope")
        scope_layout = QtWidgets.QVBoxLayout(scope_box)
        self.scope_body = QtWidgets.QCheckBox("Body")
        self.scope_comments = QtWidgets.QCheckBox("Comments")
        self.scope_headers = QtWidgets.QCheckBox("Headers / footers")
        scope_layout.addWidget(self.scope_body)
        scope_layout.addWidget(self.scope_comments)
        scope_layout.addWidget(self.scope_headers)
        self.form.addRow(scope_box)

        self.alignment_combo = QtWidgets.QComboBox()
        self.alignment_combo.addItem("Right", "right")
        self.alignment_combo.addItem("Center", "center")
        self.alignment_combo.addItem("Left", "left")
        self.direction_value = QtWidgets.QLabel("RTL (fixed)")
        self.font_name_edit = QtWidgets.QLineEdit()
        self.arabic_only_check = QtWidgets.QCheckBox("Only touch text that already contains Arabic characters")
        self.normalize_tables_check = QtWidgets.QCheckBox("Normalize tables")
        self.normalize_lists_check = QtWidgets.QCheckBox("Normalize lists")

        self.form.addRow("Alignment:", self.alignment_combo)
        self.form.addRow("Direction:", self.direction_value)
        self.form.addRow("Font:", self.font_name_edit)
        self.form.addRow(self.arabic_only_check)
        self.form.addRow(self.normalize_tables_check)
        self.form.addRow(self.normalize_lists_check)

        self._apply_engine_gating()
        self._load_initial(initial_params)

    def _apply_engine_gating(self) -> None:
        if self._is_docx_engine:
            self.scope_comments.setChecked(False)
            self.scope_comments.setEnabled(False)
            self.scope_headers.setChecked(False)
            self.scope_headers.setEnabled(False)
        else:
            if "comments" not in self._engine_capabilities:
                self.scope_comments.setChecked(False)
                self.scope_comments.setEnabled(False)
            if "headers_footers" not in self._engine_capabilities:
                self.scope_headers.setChecked(False)
                self.scope_headers.setEnabled(False)

    def _load_initial(self, initial_params: dict) -> None:
        scope = initial_params.get("scope", {})
        self.scope_body.setChecked(bool(scope.get("body", True)))
        self.scope_comments.setChecked(bool(scope.get("comments", False)))
        self.scope_headers.setChecked(bool(scope.get("headers_footers", False)))
        _set_combo_data(self.alignment_combo, initial_params.get("alignment", "right"))
        self.font_name_edit.setText(str(initial_params.get("font_name", "")))
        self.arabic_only_check.setChecked(bool(initial_params.get("arabic_only", True)))
        self.normalize_tables_check.setChecked(bool(initial_params.get("normalize_tables", True)))
        self.normalize_lists_check.setChecked(bool(initial_params.get("normalize_lists", True)))
        self._apply_engine_gating()

    def get_params(self) -> dict:
        return {
            "scope": {
                "body": self.scope_body.isChecked(),
                "comments": self.scope_comments.isChecked(),
                "headers_footers": self.scope_headers.isChecked(),
            },
            "alignment": self.alignment_combo.currentData(),
            "direction": "rtl",
            "font_name": self.font_name_edit.text().strip() or None,
            "arabic_only": self.arabic_only_check.isChecked(),
            "normalize_tables": self.normalize_tables_check.isChecked(),
            "normalize_lists": self.normalize_lists_check.isChecked(),
        }


class TocBuilderFormDialog(_BaseOperationFormDialog):
    def __init__(
        self,
        initial_params: dict,
        paragraph_style_names: list[str],
        parent: QtWidgets.QWidget | None = None,
    ) -> None:
        super().__init__("Custom TOC Builder", parent)
        self._paragraph_style_names = paragraph_style_names

        intro = QtWidgets.QLabel(
            "Choose the paragraph styles that should appear in the table of contents and how each one maps to a TOC level."
        )
        intro.setWordWrap(True)
        self._layout.insertWidget(0, intro)

        self.style_rows_table = QtWidgets.QTableWidget(0, 2)
        self.style_rows_table.setHorizontalHeaderLabels(["Paragraph style", "TOC level"])
        self.style_rows_table.horizontalHeader().setStretchLastSection(False)
        self.style_rows_table.horizontalHeader().setSectionResizeMode(
            0, QtWidgets.QHeaderView.Stretch
        )
        self.style_rows_table.horizontalHeader().setSectionResizeMode(
            1, QtWidgets.QHeaderView.ResizeToContents
        )
        add_style_row = QtWidgets.QHBoxLayout()
        self.add_style_button = QtWidgets.QPushButton("Add style row")
        self.remove_style_button = QtWidgets.QPushButton("Remove selected row")
        self.add_style_button.clicked.connect(self._add_style_row)
        self.remove_style_button.clicked.connect(self._remove_selected_row)
        add_style_row.addWidget(self.add_style_button)
        add_style_row.addWidget(self.remove_style_button)
        add_style_row.addStretch(1)
        self.form.addRow("Style mapping:", self.style_rows_table)
        self.form.addRow(add_style_row)

        self.title_text_edit = QtWidgets.QLineEdit()
        self.tab_leader_combo = QtWidgets.QComboBox()
        for value in ["dots", "spaces", "dashes", "lines", "heavy"]:
            self.tab_leader_combo.addItem(value.title(), value)
        self.show_page_numbers_check = QtWidgets.QCheckBox("Show page numbers")
        self.use_hyperlinks_check = QtWidgets.QCheckBox("Use hyperlinks")
        self.right_align_page_numbers_check = QtWidgets.QCheckBox("Right align page numbers")
        self.toc_font_name_edit = QtWidgets.QLineEdit()
        self.toc_font_size_spin = QtWidgets.QDoubleSpinBox()
        self.toc_font_size_spin.setRange(0.0, 200.0)
        self.toc_font_size_spin.setDecimals(1)
        self.toc_font_size_spin.setSpecialValueText("Use default")
        self.title_font_name_edit = QtWidgets.QLineEdit()
        self.title_font_size_spin = QtWidgets.QDoubleSpinBox()
        self.title_font_size_spin.setRange(0.0, 200.0)
        self.title_font_size_spin.setDecimals(1)
        self.title_font_size_spin.setSpecialValueText("Use default")
        self.title_alignment_combo = QtWidgets.QComboBox()
        self.title_alignment_combo.addItem("Left", "left")
        self.title_alignment_combo.addItem("Center", "center")
        self.title_alignment_combo.addItem("Right", "right")
        self.insertion_location_combo = QtWidgets.QComboBox()
        self.insertion_location_combo.addItem("Start of document", "start")
        self.insertion_location_combo.addItem("Current Word cursor", "cursor")
        self.insertion_location_combo.addItem("Bookmark", "bookmark")
        self.insertion_location_combo.currentIndexChanged.connect(self._refresh_bookmark_enabled)
        self.bookmark_name_edit = QtWidgets.QLineEdit()
        self.replace_existing_check = QtWidgets.QCheckBox("Replace existing TOC")

        self.form.addRow("TOC title:", self.title_text_edit)
        self.form.addRow("Tab leader:", self.tab_leader_combo)
        self.form.addRow(self.show_page_numbers_check)
        self.form.addRow(self.use_hyperlinks_check)
        self.form.addRow(self.right_align_page_numbers_check)
        self.form.addRow("TOC font:", self.toc_font_name_edit)
        self.form.addRow("TOC font size:", self.toc_font_size_spin)
        self.form.addRow("Title font:", self.title_font_name_edit)
        self.form.addRow("Title font size:", self.title_font_size_spin)
        self.form.addRow("Title alignment:", self.title_alignment_combo)
        self.form.addRow("Insert at:", self.insertion_location_combo)
        self.form.addRow("Bookmark name:", self.bookmark_name_edit)
        self.form.addRow(self.replace_existing_check)

        self._load_initial(initial_params)

    def _make_style_combo(self) -> QtWidgets.QComboBox:
        combo = QtWidgets.QComboBox()
        combo.setEditable(True)
        combo.addItems(self._paragraph_style_names)
        return combo

    def _make_level_spin(self) -> QtWidgets.QSpinBox:
        spin = QtWidgets.QSpinBox()
        spin.setRange(1, 9)
        return spin

    def _add_style_row(self, style_name: str = "", level: int = 1) -> None:
        row = self.style_rows_table.rowCount()
        self.style_rows_table.insertRow(row)
        style_combo = self._make_style_combo()
        style_combo.setCurrentText(style_name)
        level_spin = self._make_level_spin()
        level_spin.setValue(level)
        self.style_rows_table.setCellWidget(row, 0, style_combo)
        self.style_rows_table.setCellWidget(row, 1, level_spin)

    def _remove_selected_row(self) -> None:
        row = self.style_rows_table.currentRow()
        if row >= 0:
            self.style_rows_table.removeRow(row)

    def _refresh_bookmark_enabled(self) -> None:
        is_bookmark = self.insertion_location_combo.currentData() == "bookmark"
        self.bookmark_name_edit.setEnabled(is_bookmark)

    def _load_initial(self, initial_params: dict) -> None:
        style_levels = initial_params.get("style_levels", [])
        if style_levels:
            for item in style_levels:
                self._add_style_row(
                    style_name=str(item.get("style_name", "")),
                    level=int(item.get("level", 1)),
                )
        else:
            self._add_style_row()
        self.title_text_edit.setText(str(initial_params.get("title_text", "Contents")))
        _set_combo_data(self.tab_leader_combo, initial_params.get("tab_leader", "dots"))
        self.show_page_numbers_check.setChecked(bool(initial_params.get("show_page_numbers", True)))
        self.use_hyperlinks_check.setChecked(bool(initial_params.get("use_hyperlinks", True)))
        self.right_align_page_numbers_check.setChecked(
            bool(initial_params.get("right_align_page_numbers", True))
        )
        self.toc_font_name_edit.setText(str(initial_params.get("toc_font_name", "")))
        self.toc_font_size_spin.setValue(float(initial_params.get("toc_font_size") or 0.0))
        self.title_font_name_edit.setText(str(initial_params.get("title_font_name", "")))
        self.title_font_size_spin.setValue(float(initial_params.get("title_font_size") or 0.0))
        _set_combo_data(self.title_alignment_combo, initial_params.get("title_alignment", "left"))
        _set_combo_data(
            self.insertion_location_combo,
            initial_params.get("insertion_location", "start"),
        )
        self.bookmark_name_edit.setText(str(initial_params.get("bookmark_name", "")))
        self.replace_existing_check.setChecked(bool(initial_params.get("replace_existing", True)))
        self._refresh_bookmark_enabled()

    def get_params(self) -> dict:
        style_levels = []
        for row in range(self.style_rows_table.rowCount()):
            style_combo = self.style_rows_table.cellWidget(row, 0)
            level_spin = self.style_rows_table.cellWidget(row, 1)
            if not isinstance(style_combo, QtWidgets.QComboBox):
                continue
            if not isinstance(level_spin, QtWidgets.QSpinBox):
                continue
            style_name = style_combo.currentText().strip()
            if not style_name:
                continue
            style_levels.append(
                {"style_name": style_name, "level": int(level_spin.value())}
            )

        params = {
            "style_levels": style_levels,
            "title_text": self.title_text_edit.text().strip(),
            "tab_leader": self.tab_leader_combo.currentData(),
            "show_page_numbers": self.show_page_numbers_check.isChecked(),
            "use_hyperlinks": self.use_hyperlinks_check.isChecked(),
            "right_align_page_numbers": self.right_align_page_numbers_check.isChecked(),
            "title_alignment": self.title_alignment_combo.currentData(),
            "insertion_location": self.insertion_location_combo.currentData(),
            "replace_existing": self.replace_existing_check.isChecked(),
        }
        toc_font_name = self.toc_font_name_edit.text().strip()
        title_font_name = self.title_font_name_edit.text().strip()
        bookmark_name = self.bookmark_name_edit.text().strip()
        if toc_font_name:
            params["toc_font_name"] = toc_font_name
        if self.toc_font_size_spin.value():
            params["toc_font_size"] = self.toc_font_size_spin.value()
        if title_font_name:
            params["title_font_name"] = title_font_name
        if self.title_font_size_spin.value():
            params["title_font_size"] = self.title_font_size_spin.value()
        if self.insertion_location_combo.currentData() == "bookmark":
            params["bookmark_name"] = bookmark_name
        return params
