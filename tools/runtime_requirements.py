"""Din requirements.lock (pip-compile, cu hash-uri) pastreaza doar dependentele aplicatiei.

Lock-ul e generat din requirements-dev.txt, deci contine si uneltele de dezvoltare (pytest,
PyInstaller, ruff...). Pentru SBOM conteaza doar ce ajunge in aplicatie: pachetele cerute
de requirements.txt si, pe rand, dependentele lor, dupa adnotarile „# via” ale lui pip-compile.

    python tools/runtime_requirements.py requirements.lock -o runtime.lock

Rezultatul se instaleaza cu `pip install --require-hashes --no-deps -r runtime.lock`.
"""

from __future__ import annotations

import argparse
import re
import sys
from dataclasses import dataclass, field
from pathlib import Path

RUNTIME_SOURCE = "-r requirements.txt"
_PACKAGE = re.compile(r"^([A-Za-z0-9][A-Za-z0-9._-]*)==\S+")


def _normalize(name: str) -> str:
    return re.sub(r"[-_.]+", "-", name).lower()


@dataclass
class Block:
    name: str
    lines: list[str] = field(default_factory=list)  # pachetul si hash-urile lui
    via: set[str] = field(default_factory=set)  # cine il cere (pachete sau „-r fisier”)


def parse(text: str) -> list[Block]:
    blocks: list[Block] = []
    in_via = False
    for line in text.splitlines():
        match = _PACKAGE.match(line)
        if match:
            blocks.append(Block(_normalize(match.group(1)), [line]))
            in_via = False
            continue
        if not blocks:
            continue
        block = blocks[-1]
        stripped = line.strip()
        if stripped.startswith("--hash="):
            block.lines.append(line)
        elif stripped.startswith("# via"):
            rest = stripped.removeprefix("# via").strip()
            if rest:
                block.via.add(rest if rest.startswith("-r ") else _normalize(rest))
            in_via = not rest
        elif in_via and stripped.startswith("#   "):
            source = stripped.removeprefix("#").strip()
            block.via.add(source if source.startswith("-r ") else _normalize(source))
        else:
            in_via = False
    return blocks


def runtime_blocks(blocks: list[Block]) -> list[Block]:
    """Pachetele cerute de requirements.txt si, recursiv, dependentele lor."""
    runtime = {b.name for b in blocks if RUNTIME_SOURCE in b.via}
    changed = True
    while changed:
        changed = False
        for block in blocks:
            if block.name not in runtime and block.via & runtime:
                runtime.add(block.name)
                changed = True
    return [b for b in blocks if b.name in runtime]


def render(blocks: list[Block]) -> str:
    header = "# Dependentele aplicatiei, extrase din requirements.lock (pentru SBOM)\n"
    return header + "".join("\n".join(b.lines) + "\n" for b in blocks)


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(
        description="Pastreaza din requirements.lock doar dependentele aplicatiei.")
    parser.add_argument("lock", type=Path)
    parser.add_argument("-o", "--output", type=Path)
    args = parser.parse_args(argv)
    blocks = runtime_blocks(parse(args.lock.read_text(encoding="utf-8")))
    if not blocks:
        print(f"Nicio dependenta „{RUNTIME_SOURCE}” in {args.lock}", file=sys.stderr)
        return 1
    text = render(blocks)
    if args.output:
        args.output.write_text(text, encoding="utf-8", newline="\n")
    else:
        sys.stdout.write(text)
    return 0


if __name__ == "__main__":
    sys.exit(main())
