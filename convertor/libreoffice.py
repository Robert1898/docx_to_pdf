"""Alternativa la Office: LibreOffice in mod headless (doar documente -> PDF)."""

from __future__ import annotations

import json
import logging
import shutil
import subprocess
import tempfile
from pathlib import Path

from .core import ExportOptions, Job, Kind
from .office import CREATE_NO_WINDOW, ConversionError, ConversionTimeout

log = logging.getLogger(__name__)

FILTERS = {
    Kind.WORD: "writer_pdf_Export",
    Kind.EXCEL: "calc_pdf_Export",
    Kind.POWERPOINT: "impress_pdf_Export",
}


def find_soffice() -> Path | None:
    candidates = [
        shutil.which("soffice"),
        r"C:\Program Files\LibreOffice\program\soffice.exe",
        r"C:\Program Files (x86)\LibreOffice\program\soffice.exe",
    ]
    for candidate in candidates:
        if candidate and Path(candidate).is_file():
            return Path(candidate)
    return None


def _filter_spec(kind: Kind, options: ExportOptions) -> str:
    params: dict[str, dict[str, object]] = {
        "ExportBookmarks": {"type": "boolean", "value": str(options.bookmarks).lower()},
    }
    if options.pdfa:
        params["SelectPdfVersion"] = {"type": "long", "value": "1"}
    if options.for_screen:
        params["ReduceImageResolution"] = {"type": "boolean", "value": "true"}
        params["MaxImageResolution"] = {"type": "long", "value": "150"}
    return f"pdf:{FILTERS[kind]}:{json.dumps(params)}"


def convert(soffice: Path, job: Job, options: ExportOptions) -> None:
    if job.kind is Kind.PDF:
        raise ConversionError("Conversia PDF → DOCX necesită Microsoft Word.")

    with tempfile.TemporaryDirectory(prefix="convertor-lo-") as tmp:
        tmp_path = Path(tmp)
        # profil separat: merge si cand utilizatorul are LibreOffice deschis
        profile = (tmp_path / "profile").as_uri()
        out_dir = tmp_path / "out"
        cmd = [
            str(soffice),
            f"-env:UserInstallation={profile}",
            "--headless",
            "--norestore",
            "--convert-to",
            _filter_spec(job.kind, options),
            "--outdir",
            str(out_dir),
            str(job.src),
        ]
        proc = subprocess.Popen(
            cmd, stdout=subprocess.PIPE, stderr=subprocess.PIPE, creationflags=CREATE_NO_WINDOW
        )
        try:
            _out, err = proc.communicate(timeout=options.timeout)
        except subprocess.TimeoutExpired as exc:
            # soffice.exe porneste soffice.bin; /T opreste tot arborele
            subprocess.run(
                ["taskkill", "/PID", str(proc.pid), "/T", "/F"],
                capture_output=True,
                creationflags=CREATE_NO_WINDOW,
            )
            proc.communicate()
            raise ConversionTimeout(
                f"LibreOffice a depășit limita de {options.timeout} s și a fost oprit."
            ) from exc

        produced = out_dir / (job.src.stem + ".pdf")
        if proc.returncode != 0 or not produced.exists():
            detail = err.decode(errors="replace").strip() or f"cod {proc.returncode}"
            raise ConversionError(f"LibreOffice nu a putut converti fișierul: {detail}")
        job.dst.parent.mkdir(parents=True, exist_ok=True)
        shutil.move(str(produced), job.dst)
