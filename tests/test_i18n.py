from __future__ import annotations

import ast
import re
from pathlib import Path

import pytest

from convertor import i18n
from convertor.engine import Status
from convertor.i18n import _
from convertor.i18n_en import EN
from convertor.office import OFFICE_DIED

SOURCE = Path(__file__).resolve().parents[1] / "convertor"
PLACEHOLDER = re.compile(r"\{[^}]*\}")


def translatable() -> set[str]:
    found = {s.value for s in Status} | {OFFICE_DIED}
    for path in SOURCE.rglob("*.py"):
        for node in ast.walk(ast.parse(path.read_text(encoding="utf-8"))):
            if (isinstance(node, ast.Call) and isinstance(node.func, ast.Name)
                    and node.func.id == "_" and node.args
                    and isinstance(node.args[0], ast.Constant)):
                found.add(node.args[0].value)
    return found


def test_every_text_has_an_english_translation() -> None:
    missing = sorted(translatable() - EN.keys())
    assert not missing, missing


def test_no_stale_translations() -> None:
    stale = sorted(EN.keys() - translatable())
    assert not stale, stale


def test_placeholders_match() -> None:
    wrong = [k for k, v in EN.items()
             if sorted(PLACEHOLDER.findall(k)) != sorted(PLACEHOLDER.findall(v))]
    assert not wrong, wrong


def test_switching_language() -> None:
    assert i18n.set_language("en") == "en"
    assert _("Setări") == "Settings"
    assert Status.DONE.text == "Done"
    assert _("text necunoscut") == "text necunoscut"
    assert i18n.set_language("ro") == "ro"
    assert _("Setări") == "Setări"
    assert i18n.set_language("klingon") == "ro"


def test_system_language(monkeypatch: pytest.MonkeyPatch) -> None:
    monkeypatch.setattr(i18n, "system_language", lambda: "en")
    assert i18n.set_language("system") == "en"
