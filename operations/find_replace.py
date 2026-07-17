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

try:
    from docx.enum.text import WD_COLOR_INDEX
    from docx.shared import Pt, RGBColor
except Exception:  # noqa: BLE001
    WD_COLOR_INDEX = None  # type: ignore
    Pt = None  # type: ignore
    RGBColor = None  # type: ignore


WORD_WILD = True
WORD_WRAP_STOP = 0
WORD_COLLAPSE_END = 0

DOCX_HIGHLIGHT_MAP = {
    "yellow": "YELLOW",
    "green": "BRIGHT_GREEN",
    "turquoise": "TURQUOISE",
    "pink": "PINK",
    "blue": "BLUE",
    "red": "RED",
    "teal": "TEAL",
    "violet": "VIOLET",
}


@dataclass(slots=True)
class ReplaceScope:
    body: bool = True
    comments: bool = False
    headers_footers: bool = False

    @classmethod
    def from_params(cls, params: dict) -> "ReplaceScope":
        raw_scope = params.get("scope")
        if isinstance(raw_scope, dict):
            return cls(
                body=bool(raw_scope.get("body", True)),
                comments=bool(raw_scope.get("comments", False)),
                headers_footers=bool(raw_scope.get("headers_footers", False)),
            )
        return cls()

    def any_selected(self) -> bool:
        return self.body or self.comments or self.headers_footers

    def describe(self) -> str:
        parts: list[str] = []
        if self.body:
            parts.append("body")
        if self.comments:
            parts.append("comments")
        if self.headers_footers:
            parts.append("headers/footers")
        return ", ".join(parts)


@dataclass(slots=True)
class TextFormatSpec:
    font_name: str | None = None
    size_pt: float | None = None
    bold: bool | None = None
    italic: bool | None = None
    underline: bool | None = None
    color: str | None = None
    highlight: str | None = None

    @classmethod
    def from_value(cls, value: dict | None) -> "TextFormatSpec | None":
        if not value:
            return None
        return cls(
            font_name=value.get("font_name"),
            size_pt=value.get("size_pt"),
            bold=value.get("bold"),
            italic=value.get("italic"),
            underline=value.get("underline"),
            color=value.get("color"),
            highlight=value.get("highlight"),
        )


def parse_color_value(value: str | None) -> tuple[int, int, int] | None:
    if not value:
        return None
    normalized = value.strip().lstrip("#")
    if len(normalized) != 6:
        raise ValueError("Color values must use 6-digit hex, for example '#ff0000'.")
    try:
        return (
            int(normalized[0:2], 16),
            int(normalized[2:4], 16),
            int(normalized[4:6], 16),
        )
    except ValueError as exc:
        raise ValueError(
            "Color values must use 6-digit hex, for example '#ff0000'."
        ) from exc


def build_exact_pattern(find_text: str, whole_word: bool) -> str:
    pattern = re.escape(find_text)
    if whole_word:
        pattern = rf"\b{pattern}\b"
    return pattern


def replace_text_value(
    text: str,
    find_text: str,
    replace_text: str,
    match_case: bool,
    whole_word: bool,
) -> tuple[str, int]:
    flags = 0 if match_case else re.IGNORECASE
    return re.subn(build_exact_pattern(find_text, whole_word), replace_text, text, flags=flags)


def capture_word_font(font) -> dict[str, object]:
    return {
        "Name": getattr(font, "Name", None),
        "Size": getattr(font, "Size", None),
        "Bold": getattr(font, "Bold", None),
        "Italic": getattr(font, "Italic", None),
        "Underline": getattr(font, "Underline", None),
        "Color": getattr(font, "Color", None),
        "HighlightColorIndex": getattr(font, "HighlightColorIndex", None),
    }


def apply_word_format(target, spec: TextFormatSpec) -> None:
    if spec.font_name:
        target.Name = spec.font_name
    if spec.size_pt is not None:
        target.Size = float(spec.size_pt)
    if spec.bold is not None:
        target.Bold = bool(spec.bold)
    if spec.italic is not None:
        target.Italic = bool(spec.italic)
    if spec.underline is not None:
        target.Underline = 1 if spec.underline else 0
    rgb = parse_color_value(spec.color)
    if rgb is not None:
        target.Color = rgb[0] + (rgb[1] << 8) + (rgb[2] << 16)


def reapply_word_font(target, captured: dict[str, object]) -> None:
    for key, value in captured.items():
        if value is not None:
            setattr(target, key, value)


def apply_docx_format(run, spec: TextFormatSpec | None) -> None:
    if spec is None:
        return
    font = run.font
    if spec.font_name:
        font.name = spec.font_name
    if spec.size_pt is not None and Pt is not None:
        font.size = Pt(float(spec.size_pt))
    if spec.bold is not None:
        font.bold = bool(spec.bold)
    if spec.italic is not None:
        font.italic = bool(spec.italic)
    if spec.underline is not None:
        font.underline = bool(spec.underline)
    rgb = parse_color_value(spec.color)
    if rgb is not None and RGBColor is not None:
        font.color.rgb = RGBColor(*rgb)
    if spec.highlight and WD_COLOR_INDEX is not None:
        highlight_name = DOCX_HIGHLIGHT_MAP.get(spec.highlight.strip().lower())
        if highlight_name and hasattr(WD_COLOR_INDEX, highlight_name):
            font.highlight_color = getattr(WD_COLOR_INDEX, highlight_name)


def iter_com_scope_ranges(document, scope: ReplaceScope):
    if scope.body:
        yield "Body", document.Content

    if scope.comments:
        comments = getattr(document, "Comments", None)
        if comments is not None:
            for index in range(1, comments.Count + 1):
                try:
                    comment = comments.Item(index)
                    yield f"Comment#{index}", comment.Range
                except Exception:  # noqa: BLE001
                    continue

    if scope.headers_footers:
        sections = getattr(document, "Sections", None)
        if sections is None:
            return
        for section_index in range(1, sections.Count + 1):
            section = sections.Item(section_index)
            headers = getattr(section, "Headers", None)
            footers = getattr(section, "Footers", None)
            if headers is not None:
                for item_index in range(1, headers.Count + 1):
                    try:
                        header = headers.Item(item_index)
                        yield f"Header S{section_index}#{item_index}", header.Range
                    except Exception:  # noqa: BLE001
                        continue
            if footers is not None:
                for item_index in range(1, footers.Count + 1):
                    try:
                        footer = footers.Item(item_index)
                        yield f"Footer S{section_index}#{item_index}", footer.Range
                    except Exception:  # noqa: BLE001
                        continue


@OperationRegistry.register("Find/Replace")
class FindReplaceOperation(Operation):
    name = "Find/Replace"
    required_capabilities = {"find_replace_basic"}
    params_schema = {
        "find_text": {"type": "string", "required": True},
        "replace_text": {"type": "string", "required": True},
        "scope": {"type": "object"},
        "match_case": {"type": "boolean", "default": False},
        "whole_word": {"type": "boolean", "default": False},
        "use_regex": {"type": "boolean", "default": False},
        "use_wildcards": {"type": "boolean", "default": False},
        "match_format": {"type": "object", "nullable": True},
        "target_format": {"type": "object", "nullable": True},
    }

    def validate(self, engine) -> None:
        self.ensure_supported(engine)
        find_text = str(self.params.get("find_text", ""))
        if not find_text:
            raise OperationValidationError("Find text is required.")

        scope = ReplaceScope.from_params(self.params)
        if not scope.any_selected():
            raise OperationValidationError("At least one scope must be selected.")

        use_regex = bool(self.params.get("use_regex", False))
        use_wildcards = bool(self.params.get("use_wildcards", False))
        if use_regex and use_wildcards:
            raise OperationValidationError(
                "Regex mode and wildcard mode cannot be enabled together."
            )

        match_format = TextFormatSpec.from_value(self.params.get("match_format"))
        target_format = TextFormatSpec.from_value(self.params.get("target_format"))

        if isinstance(engine, DocxEngine):
            if scope.comments or scope.headers_footers:
                raise OperationValidationError(
                    "python-docx mode supports body-only find/replace."
                )
            if use_regex or use_wildcards:
                raise OperationValidationError(
                    "python-docx mode does not support regex or wildcard find/replace."
                )
            if match_format is not None:
                raise OperationValidationError(
                    "python-docx mode does not support source formatting filters."
                )
        else:
            if scope.comments and "comments" not in engine.capabilities:
                raise OperationValidationError(
                    "The current engine does not support comment find/replace."
                )
            if scope.headers_footers and "headers_footers" not in engine.capabilities:
                raise OperationValidationError(
                    "The current engine does not support header/footer find/replace."
                )
            if (use_regex or use_wildcards or match_format is not None or target_format is not None) and (
                "find_replace_advanced" not in engine.capabilities
            ):
                raise OperationValidationError(
                    "The current engine does not support advanced find/replace features."
                )
            if use_regex:
                try:
                    re.compile(
                        find_text,
                        0 if self.params.get("match_case", False) else re.IGNORECASE,
                    )
                except re.error as exc:
                    raise OperationValidationError(
                        f"Invalid regex pattern: {exc}"
                    ) from exc

        if target_format is not None:
            parse_color_value(target_format.color)
        if match_format is not None:
            parse_color_value(match_format.color)

    def run(self, engine) -> OperationResult:
        self.validate(engine)
        if isinstance(engine, ComEngine):
            return self._run_com(engine)
        if isinstance(engine, DocxEngine):
            return self._run_docx(engine)
        raise OperationValidationError("Unsupported engine type for Find/Replace.")

    def describe(self) -> str:
        scope = ReplaceScope.from_params(self.params).describe()
        mode = "regex" if self.params.get("use_regex") else "wildcards" if self.params.get("use_wildcards") else "exact"
        return (
            f"{self.name}: '{self.params.get('find_text', '')}' -> "
            f"'{self.params.get('replace_text', '')}' [{mode}; {scope}]"
        )

    def _run_com(self, engine: ComEngine) -> OperationResult:
        scope = ReplaceScope.from_params(self.params)
        use_regex = bool(self.params.get("use_regex", False))
        use_wildcards = bool(self.params.get("use_wildcards", False))
        if use_regex:
            count = self._run_com_regex(engine, scope)
        else:
            count = self._run_com_find_replace(engine, scope, use_wildcards=use_wildcards)
        return OperationResult(
            status="ok",
            message=f"Replaced {count} match(es) via {engine.engine_name}.",
            details={
                "count": count,
                "engine": engine.engine_name,
                "scope": scope.describe(),
                "mode": "regex" if use_regex else "wildcards" if use_wildcards else "exact",
            },
        )

    def _run_docx(self, engine: DocxEngine) -> OperationResult:
        find_text = str(self.params.get("find_text", ""))
        replace_text = str(self.params.get("replace_text", ""))
        match_case = bool(self.params.get("match_case", False))
        whole_word = bool(self.params.get("whole_word", False))
        target_format = TextFormatSpec.from_value(self.params.get("target_format"))

        total = 0
        for paragraph in engine.document.paragraphs:
            for run in paragraph.runs:
                new_text, count = replace_text_value(
                    run.text,
                    find_text=find_text,
                    replace_text=replace_text,
                    match_case=match_case,
                    whole_word=whole_word,
                )
                if count:
                    run.text = new_text
                    apply_docx_format(run, target_format)
                    total += count

        return OperationResult(
            status="ok",
            message=f"Replaced {total} match(es) via {engine.engine_name}.",
            details={
                "count": total,
                "engine": engine.engine_name,
                "scope": "body",
                "mode": "exact",
            },
        )

    def _run_com_find_replace(
        self,
        engine: ComEngine,
        scope: ReplaceScope,
        use_wildcards: bool,
    ) -> int:
        find_text = str(self.params.get("find_text", ""))
        replace_text = str(self.params.get("replace_text", ""))
        match_case = bool(self.params.get("match_case", False))
        whole_word = bool(self.params.get("whole_word", False))
        match_format = TextFormatSpec.from_value(self.params.get("match_format"))
        target_format = TextFormatSpec.from_value(self.params.get("target_format"))

        total = 0
        for _, rng in iter_com_scope_ranges(engine.document, scope):
            work = rng.Duplicate
            find = work.Find
            find.ClearFormatting()
            find.Replacement.ClearFormatting()
            find.Text = find_text
            find.Replacement.Text = replace_text
            find.Forward = True
            find.Wrap = WORD_WRAP_STOP
            find.Format = bool(match_format is not None or target_format is not None)
            find.MatchCase = bool(match_case)
            find.MatchWholeWord = bool(whole_word)
            find.MatchWildcards = WORD_WILD if use_wildcards else False

            if match_format is not None:
                apply_word_format(find.Font, match_format)
            if target_format is not None:
                apply_word_format(find.Replacement.Font, target_format)

            while True:
                found = bool(find.Execute(Replace=0))
                if not found:
                    break
                saved_font = capture_word_font(work.Font)
                work.Text = replace_text
                if target_format is not None:
                    apply_word_format(work.Font, target_format)
                else:
                    reapply_word_font(work.Font, saved_font)
                total += 1
                work.Collapse(WORD_COLLAPSE_END)
        return total

    def _run_com_regex(self, engine: ComEngine, scope: ReplaceScope) -> int:
        pattern = str(self.params.get("find_text", ""))
        replacement = str(self.params.get("replace_text", ""))
        flags = 0 if self.params.get("match_case", False) else re.IGNORECASE
        compiled = re.compile(pattern, flags)
        target_format = TextFormatSpec.from_value(self.params.get("target_format"))

        total = 0
        document = engine.document
        for _, rng in iter_com_scope_ranges(document, scope):
            text = str(rng.Text)
            matches = list(compiled.finditer(text))
            for match in reversed(matches):
                start = int(rng.Start) + int(match.start())
                end = int(rng.Start) + int(match.end())
                match_range = document.Range(start, end)
                saved_font = capture_word_font(match_range.Font)
                replacement_text = match.expand(replacement)
                match_range.Text = replacement_text
                new_range = document.Range(start, start + len(replacement_text))
                if target_format is not None:
                    apply_word_format(new_range.Font, target_format)
                else:
                    reapply_word_font(new_range.Font, saved_font)
                total += 1
        return total
