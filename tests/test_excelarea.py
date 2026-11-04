from __future__ import annotations

from convertor import excelarea
from convertor.excelarea import (
    cells_with_content,
    clamp,
    content_areas,
    has_content,
    header_rows,
    label_column,
)


def cell(r: int, c: int) -> excelarea.Rect:
    return (r, c, r, c)


def test_has_content() -> None:
    assert has_content("text") and has_content(0) and has_content(0.0) and has_content(False)
    assert not has_content(None)
    assert not has_content("")  # formula care intoarce ""
    assert not has_content("   ")  # doar spatii


def test_cells_with_content_reads_value2_shapes() -> None:
    values = (("Produs", None, ""), (None, 5, "  "))
    assert cells_with_content(values, 3, 2) == [cell(3, 2), cell(4, 3)]
    assert cells_with_content("singur", 1, 1) == [cell(1, 1)]  # Value2 pe o celula: scalar
    assert cells_with_content(None, 1, 1) == []


def test_compact_table_is_one_area() -> None:
    rects = [cell(r, c) for r in range(1, 22) for c in range(1, 7)]
    assert content_areas(rects) == [(1, 1, 21, 6)]


def test_far_away_cell_becomes_its_own_area() -> None:
    rects = [cell(r, c) for r in range(1, 12) for c in range(1, 5)] + [cell(5000, 2)]
    assert content_areas(rects) == [(1, 1, 11, 4), (5000, 2, 5000, 2)]


def test_far_right_cell_becomes_its_own_area() -> None:
    rects = [cell(r, c) for r in range(1, 12) for c in range(1, 5)] + [cell(3, 200)]
    assert content_areas(rects) == [(1, 1, 11, 4), (3, 200, 3, 200)]


def test_small_gaps_stay_in_one_area() -> None:
    rects = [cell(1, 1), cell(60, 1), cell(60, 20)]
    assert content_areas(rects) == [(1, 1, 60, 20)]


def test_merged_title_widens_the_area() -> None:
    rects = [cell(r, c) for r in range(2, 10) for c in range(1, 5)] + [(1, 1, 1, 8)]
    assert content_areas(rects) == [(1, 1, 9, 8)]


def test_very_scattered_content_falls_back_to_one_area() -> None:
    rects = [cell(r * 500, 1) for r in range(1, 40)]
    assert content_areas(rects) == [(500, 1, 19500, 1)]


def test_empty_sheet_has_no_areas() -> None:
    assert content_areas([]) == []


def table(top: int, rows: int, cols: int, first_col_filled: bool = True) -> list:
    """Antet pe randul `top`, apoi `rows` randuri de date pe `cols` coloane."""
    cells = [cell(top, c) for c in range(1, cols + 1)]
    for r in range(top + 1, top + 1 + rows):
        start = 1 if first_col_filled else 2
        cells += [cell(r, c) for c in range(start, cols + 1)]
    return cells


def test_header_is_first_row_of_a_plain_table() -> None:
    cells = table(1, 20, 10)
    assert header_rows(cells, (1, 1, 21, 10)) == (1, 1)


def test_title_above_header_is_repeated_together_with_it() -> None:
    cells = [cell(1, 1), *table(3, 20, 10)]  # titlu pe randul 1, rand gol, antet pe 3
    assert header_rows(cells, (1, 1, 23, 10)) == (1, 3)


def test_no_header_for_single_column_lists() -> None:
    cells = [cell(r, 1) for r in range(1, 30)]
    assert header_rows(cells, (1, 1, 29, 1)) is None


def test_label_column_is_first_filled_column() -> None:
    cells = table(1, 20, 10)
    assert label_column(cells, (1, 1, 21, 10), below_row=1) == 1


def test_label_column_skips_mostly_empty_first_column() -> None:
    cells = table(1, 20, 10, first_col_filled=False)
    assert label_column(cells, (1, 1, 21, 10), below_row=1) == 2


def test_clamp_keeps_result_inside_used_range() -> None:
    assert clamp((1, 1, 1_048_576, 6), (1, 1, 21, 6)) == (1, 1, 21, 6)
