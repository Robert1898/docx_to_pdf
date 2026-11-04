from __future__ import annotations

from pathlib import Path

import pytest
from reportlab.pdfgen import canvas


@pytest.fixture(autouse=True)
def isolated_appdata(tmp_path: Path, monkeypatch: pytest.MonkeyPatch) -> None:
    """Setarile, jurnalele si marcajele testelor nu ajung in profilul real."""
    monkeypatch.setenv("APPDATA", str(tmp_path / "appdata"))
    monkeypatch.setenv("LOCALAPPDATA", str(tmp_path / "localappdata"))


def make_pdf(path: Path, pages: list[str]) -> Path:
    c = canvas.Canvas(str(path))
    for text in pages:
        c.drawString(72, 720, text)
        c.showPage()
    c.save()
    return path


@pytest.fixture
def pdf_factory(tmp_path: Path):
    def factory(name: str, pages: list[str]) -> Path:
        return make_pdf(tmp_path / name, pages)

    return factory
