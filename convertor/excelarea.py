"""Zonele de tiparit dintr-o foaie Excel, calculate din continutul real (fara Excel).

Primeste dreptunghiuri cu continut (celule cu valori, celule unite, grafice/imagini) si le
grupeaza in zone. Continutul aflat la o distanta mare de restul (de ex. o celula uitata la
randul 5000) devine o zona separata, tiparita pe pagina ei, in loc sa se tipareasca miile
de randuri goale dintre ele.
"""

from __future__ import annotations

from collections import Counter
from collections.abc import Iterable, Sequence

Rect = tuple[int, int, int, int]  # rand sus, coloana stanga, rand jos, coloana dreapta

ROW_GAP = 100  # peste atatea randuri goale, continutul de dedesubt e o zona separata
COL_GAP = 30
MAX_AREAS = 20
HEADER_SEARCH_ROWS = 5  # antetul e cautat doar in primele randuri ale zonei
LABEL_SEARCH_COLS = 3
FILLED = 0.6  # cat de plin trebuie sa fie un rand / o coloana ca sa conteze


def has_content(value: object) -> bool:
    """Valoare care se vede tiparita: nu None, nu text gol sau doar spatii.

    Formulele care intorc "" ajung aici ca text gol, deci sunt ignorate."""
    if value is None:
        return False
    if isinstance(value, str):
        return bool(value.strip())
    return True


def cells_with_content(values: object, first_row: int, first_col: int) -> list[Rect]:
    """Celulele cu continut din rezultatul Range.Value2 (scalar sau tuplu de tupluri)."""
    if not isinstance(values, tuple):
        values = ((values,),)
    found = []
    for i, row in enumerate(values):
        for j, value in enumerate(row):
            if has_content(value):
                r, c = first_row + i, first_col + j
                found.append((r, c, r, c))
    return found


def _merge_spans(spans: Iterable[tuple[int, int]], gap: int) -> list[tuple[int, int]]:
    merged: list[list[int]] = []
    for start, end in sorted(spans):
        if merged and start <= merged[-1][1] + gap + 1:
            merged[-1][1] = max(merged[-1][1], end)
        else:
            merged.append([start, end])
    return [(a, b) for a, b in merged]


def bounding_box(rects: Sequence[Rect]) -> Rect:
    return (min(r[0] for r in rects), min(r[1] for r in rects),
            max(r[2] for r in rects), max(r[3] for r in rects))


def content_areas(rects: Sequence[Rect], row_gap: int = ROW_GAP, col_gap: int = COL_GAP,
                  max_areas: int = MAX_AREAS) -> list[Rect]:
    """Zonele de tiparit, de sus in jos si de la stanga la dreapta."""
    if not rects:
        return []
    areas: list[Rect] = []
    for row_start, row_end in _merge_spans(((r[0], r[2]) for r in rects), row_gap):
        band = [r for r in rects if row_start <= r[0] and r[2] <= row_end]
        for col_start, col_end in _merge_spans(((r[1], r[3]) for r in band), col_gap):
            block = [r for r in band if col_start <= r[1] and r[3] <= col_end]
            areas.append(bounding_box(block))
    if len(areas) > max_areas:  # continut foarte imprastiat: o singura zona
        return [bounding_box(rects)]
    return areas


def _inside(cell: Rect, area: Rect) -> bool:
    return area[0] <= cell[0] <= area[2] and area[1] <= cell[1] <= area[3]


def header_rows(cells: Sequence[Rect], area: Rect) -> tuple[int, int] | None:
    """Randurile de repetat sus pe fiecare pagina: de la inceputul zonei pana la antet.

    Antetul e primul rand (din primele cateva) plin cam cat randurile de date; asa un titlu
    pe un singur rand deasupra tabelului nu e confundat cu antetul, dar se repeta impreuna
    cu el. None daca zona nu pare un tabel."""
    counts = Counter(c[0] for c in cells if _inside(c, area))
    if not counts:
        return None
    fullest = max(counts.values())
    if fullest < 2:
        return None
    last = min(area[0] + HEADER_SEARCH_ROWS - 1, area[2])
    for row in range(area[0], last + 1):
        if counts.get(row, 0) >= max(2, FILLED * fullest):
            return area[0], row
    return None


def label_column(cells: Sequence[Rect], area: Rect, below_row: int) -> int | None:
    """Coloana de repetat in stanga pe fiecare pagina (de ex. CNP, nume, cod): prima dintre
    primele cateva coloane completata pe majoritatea randurilor de date."""
    data = [c for c in cells if _inside(c, area) and c[0] > below_row]
    rows = {c[0] for c in data}
    if len(rows) < 2:
        return None
    per_column = Counter(c[1] for c in data)
    for col in range(area[1], min(area[1] + LABEL_SEARCH_COLS - 1, area[3]) + 1):
        if per_column.get(col, 0) >= FILLED * len(rows):
            return col
    return None


def clamp(rect: Rect, limit: Rect) -> Rect:
    """Limiteaza un dreptunghi la altul (de ex. la zona folosita a foii)."""
    return (max(rect[0], limit[0]), max(rect[1], limit[1]),
            min(rect[2], limit[2]), min(rect[3], limit[3]))
