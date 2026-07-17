from __future__ import annotations

import re
from dataclasses import dataclass

from core.engines import ComEngine, DocxEngine
from core.operations import (
    Operation,
    OperationRegistry,
    OperationResult,
    OperationValidationError,
)


WORD_COLLAPSE_END = 0
WORD_ALIGN_LEFT = 0
WORD_ALIGN_CENTER = 1
WORD_ALIGN_RIGHT = 2
WORD_TAB_LEADER_MAP = {
    "spaces": 0,
    "dots": 1,
    "dashes": 2,
    "lines": 3,
    "heavy": 4,
}
HEADING_NAME_RE = re.compile(r"^heading\s+([1-9])$", re.IGNORECASE)


@dataclass(slots=True)
class TocStyleLevel:
    style_name: str
    level: int


def normalize_style_levels(raw_value) -> list[TocStyleLevel]:
    if not isinstance(raw_value, list) or not raw_value:
        raise OperationValidationError("At least one TOC style level must be selected.")
    normalized: list[TocStyleLevel] = []
    for item in raw_value:
        if not isinstance(item, dict):
            raise OperationValidationError("Each TOC style mapping must be an object.")
        style_name = str(item.get("style_name", "")).strip()
        if not style_name:
            raise OperationValidationError("Each TOC style mapping needs a style_name.")
        try:
            level = int(item.get("level"))
        except Exception as exc:  # noqa: BLE001
            raise OperationValidationError(
                f"Invalid TOC level for style '{style_name}'."
            ) from exc
        if level < 1 or level > 9:
            raise OperationValidationError("TOC levels must be between 1 and 9.")
        normalized.append(TocStyleLevel(style_name=style_name, level=level))
    return normalized


def split_heading_and_custom_styles(
    style_levels: list[TocStyleLevel],
) -> tuple[list[int], list[TocStyleLevel]]:
    heading_levels: list[int] = []
    custom_styles: list[TocStyleLevel] = []
    for style_level in style_levels:
        match = HEADING_NAME_RE.match(style_level.style_name)
        if match:
            heading_levels.append(int(match.group(1)))
        else:
            custom_styles.append(style_level)
    return heading_levels, custom_styles


def build_added_styles_value(custom_styles: list[TocStyleLevel]) -> str:
    if not custom_styles:
        return ""
    parts: list[str] = []
    for style_level in custom_styles:
        parts.append(style_level.style_name)
        parts.append(str(style_level.level))
    return ",".join(parts)


def apply_range_font(range_obj, font_name: str | None, size_pt: float | None) -> None:
    if font_name:
        range_obj.Font.Name = font_name
        try:
            range_obj.Font.NameBi = font_name
        except Exception:  # noqa: BLE001
            pass
    if size_pt is not None:
        range_obj.Font.Size = float(size_pt)


def align_paragraph(paragraph, alignment: str | None) -> None:
    alignment_map = {
        "left": WORD_ALIGN_LEFT,
        "center": WORD_ALIGN_CENTER,
        "right": WORD_ALIGN_RIGHT,
    }
    if alignment:
        paragraph.Alignment = alignment_map.get(alignment.lower(), WORD_ALIGN_LEFT)


def get_insertion_range(document, location: str, bookmark_name: str | None):
    location = location.lower()
    if location == "start":
        return document.Range(0, 0)
    if location == "cursor":
        try:
            return document.Application.Selection.Range
        except Exception as exc:  # noqa: BLE001
            raise OperationValidationError(
                "Could not resolve the current Word cursor position for TOC insertion."
            ) from exc
    if location == "bookmark":
        if not bookmark_name:
            raise OperationValidationError(
                "bookmark_name is required when insertion_location is 'bookmark'."
            )
        try:
            return document.Bookmarks(bookmark_name).Range
        except Exception as exc:  # noqa: BLE001
            raise OperationValidationError(
                f"Bookmark '{bookmark_name}' was not found."
            ) from exc
    raise OperationValidationError(
        "insertion_location must be one of: start, cursor, bookmark."
    )


@OperationRegistry.register("Custom TOC Builder")
class TocBuilderOperation(Operation):
    name = "Custom TOC Builder"
    required_capabilities = {"toc"}
    params_schema = {
        "style_levels": {"type": "array", "required": True},
        "title_text": {"type": "string", "default": "Contents"},
        "tab_leader": {"type": "string", "default": "dots"},
        "show_page_numbers": {"type": "boolean", "default": True},
        "use_hyperlinks": {"type": "boolean", "default": True},
        "right_align_page_numbers": {"type": "boolean", "default": True},
        "toc_font_name": {"type": "string", "nullable": True},
        "toc_font_size": {"type": "number", "nullable": True},
        "title_font_name": {"type": "string", "nullable": True},
        "title_font_size": {"type": "number", "nullable": True},
        "title_alignment": {"type": "string", "default": "left"},
        "insertion_location": {"type": "string", "default": "start"},
        "bookmark_name": {"type": "string", "nullable": True},
        "replace_existing": {"type": "boolean", "default": True},
    }

    def validate(self, engine) -> None:
        self.ensure_supported(engine)
        if isinstance(engine, DocxEngine):
            raise OperationValidationError(
                "TOC generation requires Microsoft Word COM and is unavailable in python-docx mode."
            )
        if not isinstance(engine, ComEngine):
            raise OperationValidationError("Unsupported engine type for TOC generation.")

        style_levels = normalize_style_levels(self.params.get("style_levels"))
        heading_levels, custom_styles = split_heading_and_custom_styles(style_levels)
        if not heading_levels and not custom_styles:
            raise OperationValidationError("At least one TOC style level must be selected.")

        tab_leader = str(self.params.get("tab_leader", "dots")).lower()
        if tab_leader not in WORD_TAB_LEADER_MAP:
            raise OperationValidationError(
                "tab_leader must be one of: spaces, dots, dashes, lines, heavy."
            )

        insertion_location = str(self.params.get("insertion_location", "start")).lower()
        if insertion_location not in {"start", "cursor", "bookmark"}:
            raise OperationValidationError(
                "insertion_location must be one of: start, cursor, bookmark."
            )
        if insertion_location == "bookmark" and not self.params.get("bookmark_name"):
            raise OperationValidationError(
                "bookmark_name is required when insertion_location is 'bookmark'."
            )

        for key in ("toc_font_size", "title_font_size"):
            value = self.params.get(key)
            if value is None:
                continue
            try:
                parsed = float(value)
            except Exception as exc:  # noqa: BLE001
                raise OperationValidationError(f"{key} must be a number.") from exc
            if parsed <= 0:
                raise OperationValidationError(f"{key} must be greater than zero.")

    def describe(self) -> str:
        style_levels = normalize_style_levels(self.params.get("style_levels"))
        styles = ", ".join(
            f"{style_level.style_name}:{style_level.level}" for style_level in style_levels
        )
        return f"{self.name}: {styles}"

    def run(self, engine) -> OperationResult:
        self.validate(engine)
        document = engine.document
        style_levels = normalize_style_levels(self.params.get("style_levels"))
        heading_levels, custom_styles = split_heading_and_custom_styles(style_levels)
        use_heading_styles = bool(heading_levels)
        upper_level = min(heading_levels) if heading_levels else 1
        lower_level = max(heading_levels) if heading_levels else 1
        added_styles = build_added_styles_value(custom_styles)
        replace_existing = bool(self.params.get("replace_existing", True))

        existing_count = int(document.TablesOfContents.Count)
        if replace_existing and existing_count:
            for index in range(existing_count, 0, -1):
                toc = document.TablesOfContents.Item(index)
                try:
                    toc.Delete()
                except Exception:  # noqa: BLE001
                    toc.Range.Delete()

        insertion_range = get_insertion_range(
            document,
            location=str(self.params.get("insertion_location", "start")),
            bookmark_name=self.params.get("bookmark_name"),
        )

        title_text = str(self.params.get("title_text", "")).strip()
        if title_text:
            start = int(insertion_range.Start)
            insertion_range.InsertBefore(title_text + "\r")
            title_range = document.Range(start, start + len(title_text))
            apply_range_font(
                title_range,
                font_name=self.params.get("title_font_name"),
                size_pt=self.params.get("title_font_size"),
            )
            try:
                align_paragraph(
                    document.Range(start, start + len(title_text) + 1).Paragraphs.Item(1),
                    self.params.get("title_alignment"),
                )
            except Exception:  # noqa: BLE001
                pass
            insertion_range = document.Range(start, start + len(title_text) + 1)
            insertion_range.Collapse(WORD_COLLAPSE_END)

        toc = document.TablesOfContents.Add(
            Range=insertion_range,
            UseHeadingStyles=use_heading_styles,
            UpperHeadingLevel=upper_level,
            LowerHeadingLevel=lower_level,
            RightAlignPageNumbers=bool(self.params.get("right_align_page_numbers", True)),
            IncludePageNumbers=bool(self.params.get("show_page_numbers", True)),
            AddedStyles=added_styles,
            UseHyperlinks=bool(self.params.get("use_hyperlinks", True)),
            UseOutlineLevels=use_heading_styles,
        )

        try:
            toc.TabLeader = WORD_TAB_LEADER_MAP[str(self.params.get("tab_leader", "dots")).lower()]
        except Exception:  # noqa: BLE001
            pass

        apply_range_font(
            toc.Range,
            font_name=self.params.get("toc_font_name"),
            size_pt=self.params.get("toc_font_size"),
        )

        try:
            document.Repaginate()
        except Exception:  # noqa: BLE001
            pass
        try:
            toc.Update()
        except Exception:  # noqa: BLE001
            document.Fields.Update()

        try:
            toc_text = str(toc.Range.Text)
        except Exception:  # noqa: BLE001
            toc_text = ""

        return OperationResult(
            status="ok",
            message=(
                f"Built TOC with {len(style_levels)} style mapping(s) via {engine.engine_name}."
            ),
            details={
                "engine": engine.engine_name,
                "style_count": len(style_levels),
                "existing_toc_replaced": bool(existing_count and replace_existing),
                "toc_count": int(document.TablesOfContents.Count),
                "toc_text_preview": toc_text[:500],
            },
        )
