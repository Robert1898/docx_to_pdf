"""Folder urmarit: fisierele puse intr-un folder sunt convertite automat.

Un fisier e luat abia cand marimea si data lui nu s-au mai schimbat intre doua verificari
(adica s-a terminat de copiat). Rezultatul fiecarei conversii e tinut minte (cale + data +
marime), ca un fisier sa nu fie luat din nou la fiecare verificare sau dupa repornire:

* convertit cu succes: nu mai e atins (decat daca fisierul se schimba);
* esuat: e reincercat de cel mult MAX_RETRIES ori, cu pauze tot mai lungi intre incercari,
  apoi e marcat ca esuat definitiv si notat in jurnal;
* anulat (de ex. la inchiderea aplicatiei): e reluat la urmatoarea verificare, fara sa conteze
  ca incercare.

Rezultatele merg in folderul de iesire, fara sa suprascrie nimic.
"""

from __future__ import annotations

import json
import logging
import os
import threading
import time
from collections.abc import Callable, Sequence
from dataclasses import dataclass
from pathlib import Path
from typing import TYPE_CHECKING, Any, cast

from .core import Job, collect_files, free_path, plan_jobs
from .paths import state_dir

if TYPE_CHECKING:
    from .engine import Result

log = logging.getLogger(__name__)

Fingerprint = tuple[float, int]  # data modificarii, marime

MAX_RETRIES = 3  # reincercari dupa prima conversie esuata (deci cel mult 4 incercari)
RETRY_DELAY = 60.0  # secunde pana la reincercarea n (n * RETRY_DELAY)
STATE_VERSION = 2

DONE, RETRY, FAILED = "done", "retry", "failed"


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


@dataclass
class Entry:
    """Ce stim despre un fisier, pentru amprenta lui de acum."""

    fingerprint: Fingerprint
    status: str  # DONE / RETRY / FAILED
    attempts: int = 0  # conversii esuate
    last: float = 0.0  # momentul ultimei incercari (time.time())
    error: str = ""

    def to_json(self) -> dict:
        return {"fp": list(self.fingerprint), "status": self.status, "attempts": self.attempts,
                "last": self.last, "error": self.error}

    @classmethod
    def from_json(cls, data: object) -> Entry | None:
        # formatul vechi (pana la 2.3.0): doar amprenta fisierelor convertite
        if isinstance(data, list):
            data = {"fp": data, "status": DONE}
        if not isinstance(data, dict):
            return None
        data = cast("dict[str, Any]", data)  # valori din JSON: verificate de conversiile de jos
        try:
            mtime, size = data["fp"]
            entry = cls((float(mtime), int(size)), str(data.get("status", DONE)),
                        int(data.get("attempts", 0)), float(data.get("last", 0.0)),
                        str(data.get("error", "")))
        except (KeyError, TypeError, ValueError):
            return None
        return entry if entry.status in (DONE, RETRY, FAILED) else None


class WatchState:
    """Ce s-a convertit sau a esuat deja (pastrat pe disc intre porniri)."""

    def __init__(self, path: Path | None = None, retry_delay: float = RETRY_DELAY) -> None:
        self.path = path or state_dir() / "watch_done.json"
        self.retry_delay = retry_delay
        self.entries: dict[str, Entry] = {}
        try:
            data = json.loads(self.path.read_text(encoding="utf-8"))
        except (OSError, ValueError):
            return
        if isinstance(data, dict) and data.get("version") == STATE_VERSION:
            files = data.get("files", {})
        else:
            files = data  # format vechi: {cale: [data, marime]}
        if not isinstance(files, dict):
            return
        for key, value in files.items():
            entry = Entry.from_json(value)
            if entry is not None:
                self.entries[str(key)] = entry

    def _entry(self, path: Path, fingerprint: Fingerprint) -> Entry | None:
        """Intrarea fisierului, doar daca e pentru aceeasi amprenta (altfel e alt continut)."""
        entry = self.entries.get(_key(path))
        return entry if entry is not None and entry.fingerprint == fingerprint else None

    def should_convert(self, path: Path, fingerprint: Fingerprint,
                       now: float | None = None) -> bool:
        entry = self._entry(path, fingerprint)
        if entry is None:
            return True
        if entry.status != RETRY:
            return False  # convertit sau esuat definitiv
        now = time.time() if now is None else now
        return now - entry.last >= self.retry_delay * entry.attempts

    def is_done(self, path: Path, fingerprint: Fingerprint) -> bool:
        entry = self._entry(path, fingerprint)
        return entry is not None and entry.status == DONE

    def status(self, path: Path, fingerprint: Fingerprint) -> str | None:
        entry = self._entry(path, fingerprint)
        return entry.status if entry else None

    def mark_done(self, path: Path, fingerprint: Fingerprint) -> None:
        self.entries[_key(path)] = Entry(fingerprint, DONE, last=time.time())

    def mark_failed(self, path: Path, fingerprint: Fingerprint, error: str,
                    now: float | None = None) -> bool:
        """Noteaza o conversie esuata. True daca fisierul e abandonat definitiv acum."""
        previous = self._entry(path, fingerprint)
        attempts = (previous.attempts if previous else 0) + 1
        status = FAILED if attempts > MAX_RETRIES else RETRY
        self.entries[_key(path)] = Entry(fingerprint, status, attempts,
                                         time.time() if now is None else now, error)
        return status == FAILED

    def save(self) -> None:
        # uitam fisierele care nu mai exista, ca lista sa nu creasca la nesfarsit
        self.entries = {k: v for k, v in self.entries.items() if Path(k).exists()}
        data = {"version": STATE_VERSION,
                "files": {k: v.to_json() for k, v in self.entries.items()}}
        tmp = self.path.with_suffix(".tmp")
        try:
            tmp.write_text(json.dumps(data, ensure_ascii=False), encoding="utf-8")
            tmp.replace(self.path)
        except OSError as exc:
            log.warning("Nu am putut salva starea folderului urmarit: %s", exc)


def ready_files(folder_in: Path, folder_out: Path, state: WatchState,
                previous: dict[str, Fingerprint],
                now: float | None = None) -> tuple[list[Path], dict[str, Fingerprint]]:
    """Fisierele gata de convertit si amprentele vazute acum (pentru verificarea urmatoare).

    Un fisier e gata daca are aceeasi amprenta ca la verificarea anterioara (nu se mai scrie
    in el) si nu a fost deja convertit in forma asta (sau i-a venit randul la reincercare).
    Folderul de iesire e ignorat."""
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
        if previous.get(_key(path)) == fingerprint and state.should_convert(path, fingerprint,
                                                                            now):
            ready.append(path)
    return ready, seen


def plan_watch_jobs(files: list[Path], folder_out: Path, pdf_target: str) -> list[Job]:
    jobs = plan_jobs(files, folder_out, pdf_target)
    for job in jobs:
        job.dst = free_path(job.dst)  # niciodata nu suprascriem un rezultat existent
    return jobs


def record_results(state: WatchState, jobs: Sequence[Job], results: Sequence[Result] | None,
                   fingerprints: dict[str, Fingerprint], error: str = "",
                   now: float | None = None) -> list[tuple[Path, str]]:
    """Noteaza in stare rezultatul fiecarei conversii. Intoarce fisierele abandonate acum
    (cale, ultima eroare).

    Doar fisierele convertite cu succes sunt marcate ca facute. Amprenta folosita e cea de
    dinainte de conversie: un fisier schimbat intre timp e luat din nou. `results` None
    inseamna ca tot lotul a esuat (cu `error`)."""
    by_source = {_key(r.job.src): r for r in results or ()}
    gave_up = []
    for job in jobs:
        fingerprint = fingerprints.get(_key(job.src))
        if fingerprint is None:
            continue
        result = by_source.get(_key(job.src))
        status = result.status.name if result is not None else "FAILED"
        if status in ("DONE", "SKIPPED"):
            state.mark_done(job.src, fingerprint)
        elif status == "CANCELLED":
            continue  # nu e vina fisierului: il reluam fara sa numaram o incercare
        else:
            message = (result.message if result is not None else "") or error or status
            if state.mark_failed(job.src, fingerprint, message, now):
                log.warning("Folder urmarit: renunt la %s dupa %d incercari: %s",
                            job.src, MAX_RETRIES + 1, message)
                gave_up.append((job.src, message))
            else:
                log.info("Folder urmarit: %s a esuat (%s), il reincerc", job.src, message)
    return gave_up


Converter = Callable[[list[Job]], Sequence["Result"]]


class FolderWatcher(threading.Thread):
    """Verifica folderul periodic si trimite fisierele noi la conversie.

    `convert` primeste joburile si intoarce rezultatele lor (ca Engine.run); `on_give_up`
    e apelat (din firul de urmarire) pentru fiecare fisier abandonat dupa toate incercarile."""

    def __init__(self, folder_in: Path, folder_out: Path, convert: Converter,
                 interval: float = 5.0, pdf_target: Callable[[], str] = lambda: "docx",
                 on_give_up: Callable[[Path, str], None] | None = None,
                 retry_delay: float = RETRY_DELAY) -> None:
        super().__init__(name="folder-watcher", daemon=True)
        self.folder_in, self.folder_out = folder_in, folder_out
        self.convert = convert
        self.interval = interval
        self.pdf_target = pdf_target
        self.on_give_up = on_give_up
        self.state = WatchState(retry_delay=retry_delay)
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
                if ready:
                    self._convert(ready, previous)
            except Exception:  # folderul urmarit nu are voie sa opreasca aplicatia
                log.exception("Eroare in folderul urmarit")

    def _convert(self, ready: list[Path], fingerprints: dict[str, Fingerprint]) -> None:
        self.folder_out.mkdir(parents=True, exist_ok=True)
        jobs = plan_watch_jobs(ready, self.folder_out, self.pdf_target())
        results: Sequence[Result] | None = None
        error = ""
        try:
            results = self.convert(jobs)
        except Exception as exc:  # intentionat larg: orice eroare a lotului conteaza ca esec
            log.exception("Conversie esuata in folderul urmarit")
            error = str(exc) or exc.__class__.__name__
        gave_up = record_results(self.state, jobs, results, fingerprints, error)
        self.state.save()
        for path, message in gave_up:
            if self.on_give_up is not None:
                self.on_give_up(path, message)

    def stop(self) -> None:
        self._stop_event.set()
