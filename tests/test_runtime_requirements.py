"""Dependentele aplicatiei extrase din requirements.lock, pentru SBOM."""

from __future__ import annotations

import importlib.util
import sys
from pathlib import Path

import pytest

ROOT = Path(__file__).resolve().parents[1]
_spec = importlib.util.spec_from_file_location("runtime_requirements",
                                               ROOT / "tools" / "runtime_requirements.py")
assert _spec and _spec.loader
rr = importlib.util.module_from_spec(_spec)
sys.modules["runtime_requirements"] = rr  # dataclass-urile isi cauta modulul aici
_spec.loader.exec_module(rr)

LOCK = """\
#
# pip-compile --generate-hashes requirements-dev.txt
#
charset-normalizer==3.5.3 \\
    --hash=sha256:aaa
    # via reportlab
Pillow==13.0.0 \\
    --hash=sha256:bbb \\
    --hash=sha256:ccc
    # via
    #   -r requirements.txt
    #   reportlab
pluggy==1.7.0 \\
    --hash=sha256:ddd
    # via
    #   pytest
    #   pytest-cov
pytest==9.2.0 \\
    --hash=sha256:eee
    # via -r requirements-dev.txt
reportlab==5.1.0 \\
    --hash=sha256:fff
    # via -r requirements.txt
typing_extensions==4.16.0 \\
    --hash=sha256:ggg
    # via charset-normalizer

# The following packages are considered to be unsafe in a requirements file:
setuptools==84.2.1 \\
    --hash=sha256:hhh
    # via pyinstaller
"""


def test_only_the_application_dependencies_are_kept() -> None:
    blocks = rr.runtime_blocks(rr.parse(LOCK))
    # direct (requirements.txt), prin reportlab, si mai departe prin charset-normalizer
    assert [b.name for b in blocks] == ["charset-normalizer", "pillow", "reportlab",
                                        "typing-extensions"]
    text = rr.render(blocks)
    assert "Pillow==13.0.0 \\\n    --hash=sha256:bbb \\\n    --hash=sha256:ccc\n" in text
    for dev in ("pytest", "pluggy", "setuptools"):
        assert f"\n{dev}==" not in text


def test_real_lock_keeps_runtime_and_drops_dev_tools() -> None:
    names = {b.name for b in rr.runtime_blocks(rr.parse(
        (ROOT / "requirements.lock").read_text(encoding="utf-8")))}
    assert {"pywin32", "pypdf", "pypdfium2", "reportlab", "pillow", "ttkbootstrap",
            "tkinterdnd2", "pycryptodome"} <= names
    assert not names & {"pytest", "pytest-cov", "ruff", "pyinstaller", "pyright",
                        "cyclonedx-bom", "setuptools"}


def test_every_kept_package_has_hashes() -> None:
    for block in rr.runtime_blocks(rr.parse(
            (ROOT / "requirements.lock").read_text(encoding="utf-8"))):
        assert any("--hash=sha256:" in line for line in block.lines), block.name


def test_cli_writes_the_file(tmp_path: Path) -> None:
    lock = tmp_path / "requirements.lock"
    lock.write_text(LOCK, encoding="utf-8")
    out = tmp_path / "runtime.lock"
    assert rr.main([str(lock), "-o", str(out)]) == 0
    assert "reportlab==5.1.0" in out.read_text(encoding="utf-8")


def test_cli_refuses_a_lock_without_application_dependencies(tmp_path: Path,
                                                             capsys: pytest.CaptureFixture) -> None:
    lock = tmp_path / "requirements.lock"
    lock.write_text("pytest==9.2.0 \\\n    --hash=sha256:eee\n    # via -r requirements-dev.txt\n",
                    encoding="utf-8")
    assert rr.main([str(lock)]) == 1
    assert "Nicio dependenta" in capsys.readouterr().err
