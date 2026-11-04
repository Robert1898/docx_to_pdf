from __future__ import annotations

from pathlib import Path

import pytest

from convertor import cli


def test_parser_options() -> None:
    args = cli.build_parser().parse_args(["a.docx", "-o", "out", "--pdfa", "--screen",
                                          "--no-bookmarks", "--timeout", "60"])
    assert args.inputs == [Path("a.docx")]
    assert args.out == Path("out")
    assert (args.pdfa, args.screen, args.no_bookmarks, args.timeout) == (True, True, True, 60)


def test_no_inputs_prints_help(capsys: pytest.CaptureFixture[str]) -> None:
    assert cli.main([]) == 2
    assert "usage" in capsys.readouterr().out


def test_no_supported_files(tmp_path: Path, capsys: pytest.CaptureFixture[str]) -> None:
    (tmp_path / "note.txt").write_text("x")
    assert cli.main([str(tmp_path)]) == 2
    assert "Nu am găsit" in capsys.readouterr().err


def test_timeout_too_small(tmp_path: Path) -> None:
    assert cli.main([str(tmp_path), "--timeout", "1"]) == 2


def test_existing_results_are_skipped_by_default(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch, capsys: pytest.CaptureFixture[str]
) -> None:
    (tmp_path / "a.docx").write_bytes(b"x")
    (tmp_path / "a.pdf").write_bytes(b"old")
    seen = {}

    def fake_run(self, jobs, skip=()):
        seen["jobs"], seen["skip"] = list(jobs), list(skip)
        return []

    monkeypatch.setattr(cli.Engine, "run", fake_run)
    assert cli.main([str(tmp_path / "a.docx")]) == 0
    assert [j.src.name for j in seen["skip"]] == ["a.docx"]
    assert (tmp_path / "a.pdf").read_bytes() == b"old"
