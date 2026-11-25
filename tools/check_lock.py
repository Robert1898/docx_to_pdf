"""Verifica daca requirements.lock e sincronizat cu requirements.txt si requirements-dev.txt.

Regenereaza lock-ul intr-un fisier temporar, pornind de la cel din repo si fara --upgrade:
versiunile fixate raman, se schimba doar ce cer fisierele requirements (pachet nou sau scos,
versiune exclusa de un interval sau fixata altfel). Apoi il compara cu cel din repo, fara
antetul de comentarii. Daca difera, afiseaza diferenta si comanda de rulat, cod de iesire 1.

    python tools/check_lock.py

Trebuie rulat pe Windows cu Python 3.12, ca lock-ul (marcajele de platforma conteaza).
"""

from __future__ import annotations

import argparse
import difflib
import os
import shutil
import subprocess
import sys
import tempfile
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
LOCK = "requirements.lock"
SOURCE = "requirements-dev.txt"  # include si requirements.txt (-r)
OPTIONS = ["--generate-hashes", "--allow-unsafe", "--strip-extras", "--no-emit-index-url"]
FIX_COMMAND = ("python -m piptools compile " + " ".join(OPTIONS)
               + f" --output-file {LOCK} {SOURCE}")


def body(text: str) -> list[str]:
    """Lock-ul fara antetul de comentarii (comanda si versiunea de Python din el)."""
    lines = text.splitlines()
    start = 0
    while start < len(lines) and (lines[start].startswith("#") or not lines[start].strip()):
        start += 1
    return lines[start:]


def differences(current: str, regenerated: str) -> list[str]:
    return list(difflib.unified_diff(body(current), body(regenerated),
                                     f"{LOCK} (din repo)", f"{LOCK} (regenerat)", lineterm=""))


def regenerate(root: Path) -> str:
    """Lock-ul regenerat din fisierele requirements, pornind de la cel existent."""
    handle, temp = tempfile.mkstemp(prefix=".lock-check-", suffix=".txt", dir=root)
    os.close(handle)
    try:
        shutil.copyfile(root / LOCK, temp)  # pip-compile pastreaza versiunile din fisierul tinta
        subprocess.run([sys.executable, "-m", "piptools", "compile", "--quiet", *OPTIONS,
                        "--output-file", temp, SOURCE], cwd=root, check=True)
        return Path(temp).read_text(encoding="utf-8")
    finally:
        os.remove(temp)


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description="Verifica requirements.lock.")
    parser.add_argument("--root", type=Path, default=ROOT)
    args = parser.parse_args(argv)
    current = (args.root / LOCK).read_text(encoding="utf-8")
    diff = differences(current, regenerate(args.root))
    if not diff:
        print(f"{LOCK} e sincronizat cu {SOURCE} si requirements.txt.")
        return 0
    print("\n".join(diff))
    message = (f"{LOCK} nu e sincronizat cu fisierele requirements. Regenereaza-l pe Windows "
               f"cu Python 3.12 si pune-l in commit: {FIX_COMMAND}")
    print(f"\n{message}", file=sys.stderr)
    if os.environ.get("GITHUB_ACTIONS") == "true":
        print(f"::error file={LOCK}::{message}")  # vizibil si in rezumatul rularii
    return 1


if __name__ == "__main__":
    sys.exit(main())
