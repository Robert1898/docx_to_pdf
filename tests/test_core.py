from __future__ import annotations

from pathlib import Path

from convertor.core import Kind, collect_files, free_path, kind_of, plan_jobs, target_extension


def touch(path: Path) -> Path:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_bytes(b"x")
    return path


def test_kind_of_is_case_insensitive() -> None:
    assert kind_of("a.DOCX") is Kind.WORD
    assert kind_of("b.xlsm") is Kind.EXCEL
    assert kind_of("c.Pptx") is Kind.POWERPOINT
    assert kind_of("d.pdf") is Kind.PDF
    assert kind_of("e.txt") is None


def test_target_extension() -> None:
    assert target_extension(Kind.PDF) == ".docx"
    assert target_extension(Kind.WORD) == ".pdf"
    assert target_extension(Kind.EXCEL) == ".pdf"


def test_collect_files_recurses_filters_and_dedupes(tmp_path: Path) -> None:
    a = touch(tmp_path / "a.docx")
    touch(tmp_path / "~$a.docx")  # fisier temporar Office
    touch(tmp_path / "note.txt")
    b = touch(tmp_path / "sub" / "b.pdf")
    files = collect_files([tmp_path, a])
    assert files == [a.resolve(), b.resolve()]


def test_plan_jobs_next_to_source(tmp_path: Path) -> None:
    src = touch(tmp_path / "raport.docx")
    [job] = plan_jobs([src])
    assert job.dst == tmp_path / "raport.pdf"
    assert job.label == "Word → PDF"


def test_plan_jobs_never_overwrites_another_source(tmp_path: Path) -> None:
    docx = touch(tmp_path / "raport.docx")
    pdf = touch(tmp_path / "raport.pdf")
    jobs = {j.src.name: j.dst.name for j in plan_jobs([docx, pdf])}
    assert jobs == {
        "raport.docx": "raport (convertit).pdf",
        "raport.pdf": "raport (convertit).docx",
    }


def test_plan_jobs_two_sources_same_target(tmp_path: Path) -> None:
    docx = touch(tmp_path / "x.docx")
    xlsx = touch(tmp_path / "x.xlsx")
    pptx = touch(tmp_path / "x.pptx")
    names = [j.dst.name for j in plan_jobs([docx, xlsx, pptx])]
    assert names == ["x.pdf", "x (convertit).pdf", "x (convertit 2).pdf"]


def test_plan_jobs_out_dir_collisions_across_folders(tmp_path: Path) -> None:
    a = touch(tmp_path / "one" / "doc.docx")
    b = touch(tmp_path / "two" / "doc.docx")
    out = tmp_path / "out"
    names = [j.dst for j in plan_jobs([a, b], out)]
    assert names == [out / "doc.pdf", out / "doc (convertit).pdf"]


def test_free_path(tmp_path: Path) -> None:
    target = tmp_path / "f.pdf"
    assert free_path(target) == target
    touch(target)
    touch(tmp_path / "f (2).pdf")
    assert free_path(target) == tmp_path / "f (3).pdf"
