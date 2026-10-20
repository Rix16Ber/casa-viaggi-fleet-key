"""Passaggio 4. Rinomina le clip abbinandole all'orario di inizio del CSV di Coach."""

from __future__ import annotations

import csv
import shlex
from pathlib import Path

from rugby.cursor_rugby.clock import clock_key
from rugby.cursor_rugby.errors import StepError
from rugby.cursor_rugby.log import write_log
from rugby.cursor_rugby.veo import clip_clock

_STEP = "4"


def rename_clips(clip_dir: Path, table: Path, log_path: Path) -> Path:
    clip_dir = Path(clip_dir)
    try:
        rows = _read_table(table)
    except StepError as exc:
        write_log(log_path, _STEP, "errore", messaggio=exc.message)
        raise
    clips = [path for path in clip_dir.iterdir() if path.is_file() and clip_clock(path.name)]
    by_clock = {}
    for path in clips:
        clock = clip_clock(path.name)
        if clock in by_clock:
            _fail(log_path, f"due clip con lo stesso orario {clock}")
        by_clock[clock] = path

    map_path = clip_dir / "rinomina-vecchio-nuovo.tsv"
    undo_path = clip_dir / "annulla-rinomina.sh"
    if map_path.is_file() and _already_renamed(clip_dir, map_path, rows):
        write_log(log_path, _STEP, "ok", azione="gia-rinominato", righe=len(rows))
        return map_path

    pairs = []
    for row in rows:
        source = by_clock.get(row["key"])
        if source is None:
            _fail(log_path, f"nessuna clip con orario di inizio {row['inizio']}")
        pairs.append((source, row["nuovo"]))
    used = {source.name for source, _ in pairs}
    orphans = sorted(path.name for path in clips if path.name not in used)
    if orphans:
        _fail(log_path, "clip senza riga nel CSV: " + ", ".join(orphans))
    targets = [target for _, target in pairs]
    if len(targets) != len(set(targets)):
        _fail(log_path, "nomi nuovi duplicati nel CSV")

    staged = []
    try:
        for source, target_name in pairs:
            target = clip_dir / target_name
            if target.exists() and target != source:
                raise StepError(_STEP, f"nome già occupato: {target_name}")
            temporary = clip_dir / f".rinomina-{source.name}"
            source.rename(temporary)
            staged.append((temporary, target, source.name, target_name))
        lines = ["vecchio\tnuovo"]
        undo = ["#!/bin/sh", "cd \"$(dirname \"$0\")\"", "set -e"]
        for temporary, target, old_name, new_name in staged:
            temporary.rename(target)
            lines.append(f"{old_name}\t{new_name}")
            undo.append(f"mv -n -- {shlex.quote(new_name)} {shlex.quote(old_name)}")
        map_path.write_text("\n".join(lines) + "\n", encoding="utf-8")
        undo_path.write_text("\n".join(undo) + "\n", encoding="utf-8")
        undo_path.chmod(0o755)
    except Exception:
        for temporary, target, old_name, _new_name in reversed(staged):
            if temporary.exists() and not (clip_dir / old_name).exists():
                temporary.rename(clip_dir / old_name)
            elif target.exists() and target.name != old_name and not (clip_dir / old_name).exists():
                target.rename(clip_dir / old_name)
        write_log(log_path, _STEP, "errore", messaggio="rinomina annullata")
        raise
    write_log(log_path, _STEP, "ok", righe=len(pairs), mappa=str(map_path))
    return map_path


def _read_table(path: Path) -> list[dict[str, str]]:
    if not path.is_file():
        raise StepError(_STEP, f"CSV assente: {path}")
    with path.open(encoding="utf-8", newline="") as handle:
        reader = csv.DictReader(handle)
        if reader.fieldnames is None or "inizio" not in reader.fieldnames or "nuovo_nome" not in reader.fieldnames:
            raise StepError(_STEP, "il CSV deve avere le colonne inizio e nuovo_nome")
        rows = []
        for raw in reader:
            inizio = (raw.get("inizio") or "").strip()
            nuovo = (raw.get("nuovo_nome") or "").strip()
            if not inizio or not nuovo:
                raise StepError(_STEP, "riga senza inizio o senza nuovo_nome")
            rows.append({"inizio": inizio, "key": clock_key(inizio), "nuovo": nuovo})
    if not rows:
        raise StepError(_STEP, "CSV senza righe")
    keys = [row["key"] for row in rows]
    if len(keys) != len(set(keys)):
        raise StepError(_STEP, "orari di inizio duplicati nel CSV")
    return rows


def _already_renamed(clip_dir: Path, map_path: Path, rows: list[dict[str, str]]) -> bool:
    mapping = {}
    for line in map_path.read_text(encoding="utf-8").splitlines()[1:]:
        if not line.strip():
            continue
        old, _, new = line.partition("\t")
        mapping[old] = new
    if len(mapping) != len(rows):
        return False
    expected = {row["nuovo"] for row in rows}
    if set(mapping.values()) != expected:
        return False
    return all((clip_dir / name).is_file() for name in expected)


def _fail(log_path: Path, message: str) -> None:
    write_log(log_path, _STEP, "errore", messaggio=message)
    raise StepError(_STEP, message)
