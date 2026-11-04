"""Tipuri de fisiere, optiuni si planificarea conversiilor (fara dependente de Office)."""

from __future__ import annotations

import os
from collections.abc import Iterable
from dataclasses import dataclass, field
from enum import Enum
from pathlib import Path


class Kind(Enum):
    WORD = "word"
    EXCEL = "excel"
    POWERPOINT = "powerpoint"
    PDF = "pdf"


EXTENSIONS: dict[Kind, tuple[str, ...]] = {
    Kind.WORD: (".docx", ".doc", ".docm", ".rtf", ".odt"),
    Kind.EXCEL: (".xlsx", ".xls", ".xlsm", ".ods"),
    Kind.POWERPOINT: (".pptx", ".ppt", ".pptm", ".odp"),
    Kind.PDF: (".pdf",),
}
SUPPORTED: tuple[str, ...] = tuple(ext for exts in EXTENSIONS.values() for ext in exts)

LABELS = {
    Kind.WORD: "Word → PDF",
    Kind.EXCEL: "Excel → PDF",
    Kind.POWERPOINT: "PowerPoint → PDF",
    Kind.PDF: "PDF → DOCX",
}

COLLISION_SUFFIX = "convertit"


def kind_of(path: str | os.PathLike[str]) -> Kind | None:
    ext = Path(path).suffix.lower()
    for kind, exts in EXTENSIONS.items():
        if ext in exts:
            return kind
    return None


def target_extension(kind: Kind) -> str:
    return ".docx" if kind is Kind.PDF else ".pdf"


def is_candidate(path: Path) -> bool:
    """Fisier suportat, care nu e un fisier temporar de-al lui Office (~$nume.docx)."""
    return not path.name.startswith("~$") and kind_of(path) is not None


@dataclass
class ExportOptions:
    bookmarks: bool = True  # semne de carte din titluri (Word)
    pdfa: bool = False  # PDF/A-1 pentru arhivare
    for_screen: bool = False  # calitate redusa, fisier mai mic
    ocr: bool = False  # OCR pe PDF-urile scanate inainte de conversia in DOCX
    timeout: int = 180  # secunde per fisier
    pdf_mode: str = "auto"  # PDF -> DOCX: auto / layout (aspect fidel) / word (text curgator)


PDF_MODES = ("auto", "layout", "word")


@dataclass
class Job:
    src: Path
    dst: Path
    kind: Kind
    label: str = field(init=False)

    def __post_init__(self) -> None:
        self.label = LABELS[self.kind]


def collect_files(paths: Iterable[str | os.PathLike[str]]) -> list[Path]:
    """Extinde folderele recursiv si pastreaza doar fisierele suportate, fara dubluri."""
    seen: set[str] = set()
    result: list[Path] = []

    def add(p: Path) -> None:
        key = os.path.normcase(str(p))
        if key not in seen and is_candidate(p):
            seen.add(key)
            result.append(p)

    for raw in paths:
        path = Path(raw).resolve()
        if path.is_dir():
            for root, _dirs, names in os.walk(path):
                for name in sorted(names):
                    add(Path(root) / name)
        elif path.is_file():
            add(path)
    return result


def free_path(path: Path) -> Path:
    """Primul nume liber pe disc: raport.pdf, raport (2).pdf, raport (3).pdf..."""
    candidate, n = path, 2
    while candidate.exists():
        candidate = path.with_name(f"{path.stem} ({n}){path.suffix}")
        n += 1
    return candidate


def _unique(dst: Path, taken: set[str]) -> Path:
    if os.path.normcase(str(dst)) not in taken:
        return dst
    n = 1
    while True:
        suffix = COLLISION_SUFFIX if n == 1 else f"{COLLISION_SUFFIX} {n}"
        candidate = dst.with_name(f"{dst.stem} ({suffix}){dst.suffix}")
        if os.path.normcase(str(candidate)) not in taken:
            return candidate
        n += 1


def plan_jobs(files: Iterable[Path], out_dir: Path | None = None) -> list[Job]:
    """Stabileste fisierul rezultat pentru fiecare sursa.

    Un rezultat nu are voie sa suprascrie o alta sursa din lista (raport.docx + raport.pdf)
    si nici rezultatul altui fisier (raport.docx + raport.xlsx -> ambele raport.pdf).
    """
    files = list(files)
    taken = {os.path.normcase(str(f)) for f in files}
    jobs: list[Job] = []
    for src in files:
        kind = kind_of(src)
        if kind is None:
            continue
        folder = out_dir if out_dir is not None else src.parent
        dst = _unique(folder / (src.stem + target_extension(kind)), taken)
        taken.add(os.path.normcase(str(dst)))
        jobs.append(Job(src, dst, kind))
    return jobs
