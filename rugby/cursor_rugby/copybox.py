"""Passaggio 2. Copia la cartella della partita sul box e confronta gli sha256."""

from __future__ import annotations

import shutil
from pathlib import Path

from rugby.cursor_rugby.errors import StepError
from rugby.cursor_rugby.log import write_log
from rugby.cursor_rugby.xlsxutil import sha256_file

_STEP = "2"


def copy_to_box(source: Path, box_root: Path, partita: str, log_path: Path) -> Path:
    source = Path(source)
    if not source.is_dir():
        _fail(log_path, f"cartella di origine assente: {source}")
    files = [path for path in source.rglob("*") if path.is_file()]
    if not files:
        _fail(log_path, "cartella di origine vuota")
    empties = [str(path.relative_to(source)) for path in files if path.stat().st_size == 0]
    if empties:
        _fail(log_path, "file da 0 byte: " + ", ".join(empties))

    dest = Path(box_root) / partita
    extras = _extras(source, dest)
    if extras:
        _fail(log_path, "sul box ci sono file non presenti nell'origine: " + ", ".join(extras))

    dest.mkdir(parents=True, exist_ok=True)
    for path in files:
        relative = path.relative_to(source)
        target = dest / relative
        if target.is_file() and sha256_file(target) != sha256_file(path):
            _fail(log_path, f"file diverso già presente, non sovrascrivo: {relative}")
        if target.is_file() and sha256_file(target) == sha256_file(path):
            continue
        target.parent.mkdir(parents=True, exist_ok=True)
        shutil.copy2(path, target)

    mismatched = []
    for path in files:
        relative = path.relative_to(source)
        if sha256_file(path) != sha256_file(dest / relative):
            mismatched.append(str(relative))
    if mismatched:
        _fail(log_path, "sha256 diverso dopo la copia: " + ", ".join(mismatched))
    write_log(log_path, _STEP, "ok", cartella=str(dest), file=len(files))
    return dest


def _extras(source: Path, dest: Path) -> list[str]:
    if not dest.exists():
        return []
    source_names = {path.relative_to(source).as_posix() for path in source.rglob("*") if path.is_file()}
    return sorted(
        path.relative_to(dest).as_posix()
        for path in dest.rglob("*")
        if path.is_file() and path.name != "cursor-rugby.log" and path.relative_to(dest).as_posix() not in source_names
    )


def _fail(log_path: Path, message: str) -> None:
    write_log(log_path, _STEP, "errore", messaggio=message)
    raise StepError(_STEP, message)
