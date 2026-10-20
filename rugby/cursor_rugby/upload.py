"""Passaggio 9. Pubblica l'xlsx solo se, riletto, coincide e supera tre controlli."""

from __future__ import annotations

import shutil
from pathlib import Path

from rugby.cursor_rugby.errors import StepError
from rugby.cursor_rugby.log import write_log
from rugby.cursor_rugby.xlsxutil import repeat_checks, sha256_file

_STEP = "9"


def publish(local: Path, template: Path, destination: Path, log_path: Path) -> Path:
    local = Path(local)
    template = Path(template)
    destination = Path(destination)
    if not local.is_file() or local.stat().st_size == 0:
        _fail(log_path, f"file locale assente o vuoto: {local}")
    destination.parent.mkdir(parents=True, exist_ok=True)
    backup = destination.with_name(destination.name + ".bak")
    had_destination = destination.is_file()
    if had_destination:
        shutil.copy2(destination, backup)
        if sha256_file(destination) == sha256_file(local):
            try:
                repeat_checks(destination, template)
            except StepError:
                _fail(log_path, "il file già presente non supera i tre controlli")
            write_log(log_path, _STEP, "ok", azione="gia-uguale", file=str(destination))
            return destination

    shutil.copy2(local, destination)
    if sha256_file(destination) != sha256_file(local):
        _restore(destination, backup, had_destination)
        _fail(log_path, "il file riletto non coincide con quello scritto")
    try:
        repeat_checks(destination, template)
    except StepError as exc:
        _restore(destination, backup, had_destination)
        write_log(log_path, _STEP, "errore", messaggio=exc.message)
        raise
    write_log(log_path, _STEP, "ok", file=str(destination), backup=str(backup) if had_destination else "")
    return destination


def _restore(destination: Path, backup: Path, had_destination: bool) -> None:
    if had_destination and backup.is_file():
        shutil.copy2(backup, destination)
    elif destination.exists() and not had_destination:
        destination.unlink()


def _fail(log_path: Path, message: str) -> None:
    write_log(log_path, _STEP, "errore", messaggio=message)
    raise StepError(_STEP, message)
