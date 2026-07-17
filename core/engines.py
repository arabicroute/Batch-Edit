from __future__ import annotations

from abc import ABC, abstractmethod
from dataclasses import dataclass
from pathlib import Path
from typing import Any

try:
    import pythoncom
    import win32com.client  # type: ignore

    HAS_COM_IMPORT = True
except Exception:  # noqa: BLE001
    pythoncom = None  # type: ignore
    win32com = None  # type: ignore
    HAS_COM_IMPORT = False

try:
    from docx import Document as DocxDocument  # type: ignore

    HAS_PYDOCX = True
except Exception:  # noqa: BLE001
    DocxDocument = None  # type: ignore
    HAS_PYDOCX = False


STYLE_TYPE_MAP = {
    1: "paragraph",
    2: "character",
    3: "table",
    4: "list",
}


class EngineUnavailableError(RuntimeError):
    pass


def normalize_style_type(value: str | int | object) -> str:
    if isinstance(value, int):
        return STYLE_TYPE_MAP.get(value, "unknown")
    text = str(value).strip().lower().replace("_", " ").replace("-", " ")
    if "paragraph" in text:
        return "paragraph"
    if "character" in text:
        return "character"
    if "table" in text:
        return "table"
    if "list" in text:
        return "list"
    return "unknown"


@dataclass(slots=True)
class DocumentStyle:
    name: str
    style_type: str
    builtin: bool = False
    source: str = ""


@dataclass(slots=True)
class RuntimeStatus:
    com_available: bool
    com_reason: str
    docx_available: bool
    docx_reason: str

    def create_com_engine(self, visible: bool = False) -> "ComEngine":
        if not self.com_available:
            raise EngineUnavailableError(self.com_reason)
        return ComEngine(visible=visible)


class DocumentEngine(ABC):
    engine_name = "unknown"

    def __init__(self) -> None:
        self.document_path: str | None = None
        self.capabilities: set[str] = set()

    @abstractmethod
    def open(self, path: str) -> None:
        raise NotImplementedError

    @abstractmethod
    def save(self, path: str | None = None) -> None:
        raise NotImplementedError

    @abstractmethod
    def close(self) -> None:
        raise NotImplementedError

    @abstractmethod
    def list_styles(self) -> list[DocumentStyle]:
        raise NotImplementedError

    @abstractmethod
    def get_paragraphs(self) -> list[str]:
        raise NotImplementedError


class ComEngine(DocumentEngine):
    engine_name = "Word COM"

    def __init__(self, visible: bool = False) -> None:
        super().__init__()
        self.visible = visible
        self.word = None
        self.document = None
        self._com_initialized = False
        self.capabilities = {
            "styles",
            "find_replace_basic",
            "find_replace_advanced",
            "headers_footers",
            "comments",
            "toc",
            "rtl_basic",
            "rtl_advanced",
        }

    def open(self, path: str) -> None:
        if not HAS_COM_IMPORT:
            raise EngineUnavailableError("pywin32 is not installed.")
        normalized_path = str(Path(path).resolve())
        pythoncom.CoInitialize()
        self._com_initialized = True
        try:
            self.word = win32com.client.DispatchEx("Word.Application")
            self.word.Visible = self.visible
            self.word.DisplayAlerts = 0
            self.document = self.word.Documents.Open(normalized_path)
            self.document_path = normalized_path
        except Exception:
            self.close()
            raise

    def save(self, path: str | None = None) -> None:
        self._require_open_document()
        if path:
            self.document.SaveAs2(str(Path(path).resolve()))
            self.document_path = str(Path(path).resolve())
            return
        self.document.Save()

    def close(self) -> None:
        if self.document is not None:
            self.document.Close(SaveChanges=0)
            self.document = None
        if self.word is not None:
            self.word.Quit()
            self.word = None
        if self._com_initialized:
            pythoncom.CoUninitialize()
            self._com_initialized = False

    def list_styles(self) -> list[DocumentStyle]:
        self._require_open_document()
        styles: list[DocumentStyle] = []
        seen: set[tuple[str, str]] = set()
        for style in self.document.Styles:
            name = str(getattr(style, "NameLocal", "")).strip()
            if not name:
                continue
            style_type = normalize_style_type(getattr(style, "Type", 0))
            key = (name.casefold(), style_type)
            if key in seen:
                continue
            seen.add(key)
            styles.append(
                DocumentStyle(
                    name=name,
                    style_type=style_type,
                    builtin=bool(getattr(style, "BuiltIn", False)),
                    source=self.engine_name,
                )
            )
        return styles

    def get_paragraphs(self) -> list[str]:
        self._require_open_document()
        paragraphs: list[str] = []
        for paragraph in self.document.Paragraphs:
            text = str(paragraph.Range.Text).rstrip("\r")
            paragraphs.append(text)
        return paragraphs

    def _require_open_document(self) -> None:
        if self.document is None:
            raise RuntimeError("No Word document is currently open.")


class DocxEngine(DocumentEngine):
    engine_name = "python-docx"

    def __init__(self) -> None:
        super().__init__()
        self.document = None
        self.capabilities = {"styles", "find_replace_basic", "rtl_basic"}

    def open(self, path: str) -> None:
        if not HAS_PYDOCX:
            raise EngineUnavailableError("python-docx is not installed.")
        normalized_path = str(Path(path).resolve())
        self.document = DocxDocument(normalized_path)
        self.document_path = normalized_path

    def save(self, path: str | None = None) -> None:
        self._require_open_document()
        save_path = str(Path(path).resolve()) if path else self.document_path
        self.document.save(save_path)
        self.document_path = save_path

    def close(self) -> None:
        self.document = None

    def list_styles(self) -> list[DocumentStyle]:
        self._require_open_document()
        styles: list[DocumentStyle] = []
        seen: set[tuple[str, str]] = set()
        for style in self.document.styles:
            name = str(style.name).strip()
            if not name:
                continue
            style_type = normalize_style_type(getattr(style.type, "name", str(style.type)))
            key = (name.casefold(), style_type)
            if key in seen:
                continue
            seen.add(key)
            styles.append(
                DocumentStyle(
                    name=name,
                    style_type=style_type,
                    builtin=bool(getattr(style, "builtin", False)),
                    source=self.engine_name,
                )
            )
        return styles

    def get_paragraphs(self) -> list[str]:
        self._require_open_document()
        return [paragraph.text for paragraph in self.document.paragraphs]

    def _require_open_document(self) -> None:
        if self.document is None:
            raise RuntimeError("No DOCX document is currently open.")


def probe_com_availability() -> tuple[bool, str]:
    if not HAS_COM_IMPORT:
        return False, "COM automation unavailable: pywin32 import failed."
    word = None
    initialized = False
    try:
        pythoncom.CoInitialize()
        initialized = True
        word = win32com.client.DispatchEx("Word.Application")
        word.Visible = False
        word.DisplayAlerts = 0
        return True, "COM automation available."
    except Exception as exc:  # noqa: BLE001
        return False, f"COM automation unavailable: {exc}"
    finally:
        if word is not None:
            try:
                word.Quit()
            except Exception:  # noqa: BLE001
                pass
        if initialized:
            pythoncom.CoUninitialize()


def get_runtime_status() -> RuntimeStatus:
    com_available, com_reason = probe_com_availability()
    if HAS_PYDOCX:
        docx_available = True
        docx_reason = "python-docx available."
    else:
        docx_available = False
        docx_reason = "python-docx unavailable: import failed."
    return RuntimeStatus(
        com_available=com_available,
        com_reason=com_reason,
        docx_available=docx_available,
        docx_reason=docx_reason,
    )


def select_engine(preference: str = "auto", visible: bool = False) -> tuple[DocumentEngine, str]:
    runtime_status = get_runtime_status()
    preference = preference.lower()

    if preference not in {"auto", "com", "docx"}:
        raise ValueError("Engine preference must be 'auto', 'com', or 'docx'.")

    if preference == "com":
        return runtime_status.create_com_engine(visible=visible), runtime_status.com_reason
    if preference == "docx":
        if not runtime_status.docx_available:
            raise EngineUnavailableError(runtime_status.docx_reason)
        return DocxEngine(), runtime_status.docx_reason
    if runtime_status.com_available:
        return runtime_status.create_com_engine(visible=visible), runtime_status.com_reason
    if runtime_status.docx_available:
        return DocxEngine(), (
            f"{runtime_status.com_reason} Falling back to python-docx."
        )
    raise EngineUnavailableError(
        f"{runtime_status.com_reason} {runtime_status.docx_reason}"
    )
