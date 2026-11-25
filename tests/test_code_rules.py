"""Reguli pentru codul aplicatiei, verificate pe surse."""

from __future__ import annotations

import re
from pathlib import Path

SOURCE = Path(__file__).resolve().parents[1] / "convertor"
BROAD = re.compile(r"except\s*(\(.*\b(Exception|BaseException)\b.*\)|(Exception|BaseException)\b)"
                   r"|except\s*:|suppress\(\s*(Exception|BaseException)\s*\)")


def test_every_broad_except_explains_why() -> None:
    """Un `except Exception` (sau `suppress(Exception)`) e permis doar cu motivul scris pe
    aceeasi linie sau pe linia de deasupra; altfel trebuie prinse exceptiile concrete."""
    missing = []
    for path in sorted(SOURCE.rglob("*.py")):
        lines = path.read_text(encoding="utf-8").splitlines()
        for number, line in enumerate(lines, 1):
            if not BROAD.search(line):
                continue
            above = lines[number - 2].strip() if number > 1 else ""
            if "#" not in line and not above.startswith("#"):
                missing.append(f"{path.relative_to(SOURCE.parent)}:{number}: {line.strip()}")
    assert not missing, "\n".join(missing)


def test_rule_recognizes_broad_handlers() -> None:
    for line in ("except Exception:", "except Exception as exc:", "except BaseException:",
                 "except (OSError, Exception) as exc:", "except:",
                 "with contextlib.suppress(Exception):"):
        assert BROAD.search(line), line
    for line in ("except OSError:", "except (ImportError, tk.TclError) as exc:",
                 "except pywintypes.error as exc:", "with contextlib.suppress(FileNotFoundError):"):
        assert not BROAD.search(line), line
