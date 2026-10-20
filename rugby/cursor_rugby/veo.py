"""Passaggio 1. Veo si scarica solo dal Clubhouse, con il login di Riccardo.

Senza l'export già fatto il passaggio si ferma e non crea una cartella a metà.
Con l'export controlla i nomi, i byte e la lista delle clip.
"""

from __future__ import annotations

import re
import shutil
from pathlib import Path

from rugby.cursor_rugby.errors import StepError
from rugby.cursor_rugby.log import write_log
from rugby.cursor_rugby.names import folder_name, parse_date
from rugby.cursor_rugby.xlsxutil import sha256_file

_STEP = "1"
CLIP_PREFIX = "Fase_Gioco.mp4"


_CLIP_NAME = re.compile(r"^(\d{2,}) (\d{6})_-_.+\.mp4$", re.IGNORECASE)


def clip_clock(name: str) -> str | None:
    match = _CLIP_NAME.match(name)
    if match is None:
        return None
    return match.group(2)


def clip_pattern_ok(name: str) -> bool:
    stem, _, ext = name.rpartition(".")
    if ext.lower() != "mp4" or not stem.endswith("_-_Fase_Gioco"):
        return False
    number, space, clock = stem[: -len("_-_Fase_Gioco")].partition(" ")
    return bool(space) and number.isdigit() and len(number) == 2 and len(clock) == 6 and clock.isdigit()


def prepare_export(
    export_dir: Path,
    dest_root: Path,
    when: str,
    opponent: str,
    lista: Path,
    log_path: Path,
) -> Path:
    expected = _read_lista(lista)
    source = Path(export_dir)
    if not source.is_dir():
        _fail(log_path, f"export assente: {source}")
    clips = sorted(path for path in source.iterdir() if path.is_file() and clip_pattern_ok(path.name))
    videos = sorted(
        path
        for path in source.iterdir()
        if path.is_file() and path.suffix.lower() == ".mp4" and not clip_pattern_ok(path.name)
    )
    empties = [path.name for path in clips + videos if path.stat().st_size == 0]
    if empties:
        _fail(log_path, "file da 0 byte: " + ", ".join(empties))
    if len(videos) != 1:
        _fail(log_path, f"serve un solo video intero, trovati {len(videos)}")
    found = {path.name for path in clips}
    if found != set(expected):
        missing = [name for name in expected if name not in found]
        extra = sorted(found - set(expected))
        _fail(log_path, f"la lista Veo non coincide. mancano {missing or '—'}; in più {extra or '—'}")

    match = dest_root / folder_name(parse_date(when), opponent) / "Fasi gioco"
    incoming = [videos[0], *clips]
    if match.is_dir() and _same_bytes(incoming, match):
        write_log(log_path, _STEP, "ok", azione="gia-presente", cartella=str(match), clip=len(clips))
        return match

    temporary = match.parent / (match.name + ".in-corso")
    if temporary.exists():
        shutil.rmtree(temporary)
    temporary.mkdir(parents=True)
    try:
        for path in incoming:
            target = temporary / path.name
            shutil.copy2(path, target)
            if sha256_file(target) != sha256_file(path):
                raise StepError(_STEP, f"copia non verificata: {path.name}")
        if match.exists():
            shutil.rmtree(match)
        match.parent.mkdir(parents=True, exist_ok=True)
        temporary.rename(match)
    except Exception:
        if temporary.exists():
            shutil.rmtree(temporary, ignore_errors=True)
        raise
    write_log(log_path, _STEP, "ok", cartella=str(match), clip=len(clips), video=videos[0].name)
    return match


def refuse_remote_download(log_path: Path, url: str) -> None:
    _fail(
        log_path,
        "il passaggio 1 resta manuale: Veo permette il download solo dal Clubhouse, "
        "con il login di Riccardo (Download full game e Download all saved clips). "
        f"URL non scaricato: {url}. Esporta la partita e rilancia con --export e --lista.",
    )


def _read_lista(path: Path) -> list[str]:
    if not path.is_file():
        raise StepError(_STEP, f"lista Veo assente: {path}")
    names = [line.strip() for line in path.read_text(encoding="utf-8").splitlines() if line.strip()]
    if not names:
        raise StepError(_STEP, "lista Veo vuota")
    bad = [name for name in names if not clip_pattern_ok(name)]
    if bad:
        raise StepError(_STEP, "nomi clip non nel formato NN hhmmss_-_Fase_Gioco.mp4: " + ", ".join(bad))
    if len(names) != len(set(names)):
        raise StepError(_STEP, "lista Veo con nomi duplicati")
    return names


def _same_bytes(sources: list[Path], dest: Path) -> bool:
    for source in sources:
        target = dest / source.name
        if not target.is_file() or target.stat().st_size == 0:
            return False
        if sha256_file(source) != sha256_file(target):
            return False
    return True


def _fail(log_path: Path, message: str) -> None:
    write_log(log_path, _STEP, "errore", messaggio=message)
    raise StepError(_STEP, message)
