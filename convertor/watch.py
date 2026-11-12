"""Folder urmarit: fisierele puse intr-un folder sunt convertite automat.

Un fisier e luat abia cand marimea si data lui nu s-au mai schimbat intre doua verificari
(adica s-a terminat de copiat). Fisierele convertite sunt tinute minte (cale + data + marime),
ca sa nu fie convertite din nou la fiecare verificare sau dupa repornire. Rezultatele merg in
folderul de iesire, fara sa suprascrie nimic.
"""

from __future__ import annotations

import json
import logging
import os
import threading
from collections.abc import Callable
from pathlib import Path

from .core import Job, collect_files, free_path, plan_jobs
from .paths import state_dir

log = logging.getLogger(__name__)

Fingerprint = tuple[float, int]  # data modificarii, marime


def _key(path: Path) -> str:
    return os.path.normcase(str(path))


def _fingerprint(path: Path) -> Fingerprint | None:
    try:
        stat = path.stat()
    except OSError:
        return None
    return stat.st_mtime, stat.st_size


def default_output(folder_in: Path) -> Path:
    return folder_in / "Convertite"


class WatchState:
    """Ce s-a convertit deja (pastrat pe disc intre porniri)."""

    def __init__(self, path: Path | None = None) -> None:
        self.path = path or state_dir() / "watch_done.json"
        try:
            data = json.loads(self.path.read_text(encoding="utf-8"))
            self.done: dict[str, list] = data if isinstance(data, dict) else {}
        except (OSError, ValueError):
            self.done = {}

    def is_done(self, path: Path, fingerprint: Fingerprint) -> bool:
        return self.done.get(_key(path)) == list(fingerprint)

    def mark(self, path: Path, fingerprint: Fingerprint) -> None:
        self.done[_key(path)] = list(fingerprint)

    def save(self) -> None:
        # uitam fisierele care nu mai exista, ca lista sa nu creasca la nesfarsit
        self.done = {k: v for k, v in self.done.items() if Path(k).exists()}
        tmp = self.path.with_suffix(".tmp")
        try:
            tmp.write_text(json.dumps(self.done), encoding="utf-8")
            tmp.replace(self.path)
        except OSError as exc:
            log.warning("Nu am putut salva starea folderului urmarit: %s", exc)


def ready_files(folder_in: Path, folder_out: Path, state: WatchState,
                previous: dict[str, Fingerprint]) -> tuple[list[Path], dict[str, Fingerprint]]:
    """Fisierele gata de convertit si amprentele vazute acum (pentru verificarea urmatoare).

    Un fisier e gata daca are aceeasi amprenta ca la verificarea anterioara (nu se mai scrie
    in el) si nu a fost deja convertit in forma asta. Folderul de iesire e ignorat."""
    out_key = _key(folder_out)
    seen: dict[str, Fingerprint] = {}
    ready = []
    for path in collect_files([folder_in]):
        if _key(path).startswith(out_key + os.sep) or _key(path.parent) == out_key:
            continue
        fingerprint = _fingerprint(path)
        if fingerprint is None:
            continue
        seen[_key(path)] = fingerprint
        if previous.get(_key(path)) == fingerprint and not state.is_done(path, fingerprint):
            ready.append(path)
    return ready, seen


def plan_watch_jobs(files: list[Path], folder_out: Path, pdf_target: str) -> list[Job]:
    jobs = plan_jobs(files, folder_out, pdf_target)
    for job in jobs:
        job.dst = free_path(job.dst)  # niciodata nu suprascriem un rezultat existent
    return jobs


class FolderWatcher(threading.Thread):
    """Verifica folderul periodic si trimite fisierele noi la conversie."""

    def __init__(self, folder_in: Path, folder_out: Path,
                 convert: Callable[[list[Job]], list], interval: float = 5.0,
                 pdf_target: Callable[[], str] = lambda: "docx") -> None:
        super().__init__(name="folder-watcher", daemon=True)
        self.folder_in, self.folder_out = folder_in, folder_out
        self.convert = convert
        self.interval = interval
        self.pdf_target = pdf_target
        self.state = WatchState()
        self._stop_event = threading.Event()

    def run(self) -> None:
        log.info("Urmaresc %s -> %s", self.folder_in, self.folder_out)
        previous: dict[str, Fingerprint] = {}
        while not self._stop_event.wait(self.interval):
            if not self.folder_in.is_dir():
                continue
            try:
                ready, previous = ready_files(self.folder_in, self.folder_out, self.state,
                                              previous)
                if not ready:
                    continue
                self.folder_out.mkdir(parents=True, exist_ok=True)
                jobs = plan_watch_jobs(ready, self.folder_out, self.pdf_target())
                self.convert(jobs)
                for path in ready:
                    fingerprint = _fingerprint(path)
                    if fingerprint is not None:
                        self.state.mark(path, fingerprint)
                self.state.save()
            except Exception:  # folderul urmarit nu are voie sa opreasca aplicatia
                log.exception("Eroare in folderul urmarit")

    def stop(self) -> None:
        self._stop_event.set()
