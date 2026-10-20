"""Passaggio 4. Rinomina le clip con il CSV di Coach.

Il CSV non si cambia: nome attuale, nome nuovo, orario video Veo, nota Veo.
La mappa vecchio→nuovo e annulla-rinomina.sh restano quelle dello script di Padova,
con la cartella e il numero di righe che arrivano dal comando.
"""

from __future__ import annotations

import csv
from pathlib import Path

from rugby.cursor_rugby.clock import clock_key
from rugby.cursor_rugby.errors import StepError
from rugby.cursor_rugby.log import write_log
from rugby.cursor_rugby.veo import clip_clock

_STEP = "4"
_COLUMNS = ("nome attuale", "nome nuovo", "orario video Veo", "nota Veo")


def rename_clips(clip_dir: Path, table: Path, log_path: Path, dry_run: bool = False) -> tuple[int, int]:
    clip_dir = Path(clip_dir)
    if not clip_dir.is_dir():
        _fail(log_path, f"NO_DIR: {clip_dir}")
    try:
        rows = _read_table(table)
    except StepError as exc:
        write_log(log_path, _STEP, "errore", messaggio=exc.message)
        raise

    clips = [path for path in clip_dir.iterdir() if path.is_file() and clip_clock(path.name)]
    by_name = {}
    for path in clips:
        if path.name in by_name:
            _fail(log_path, f"due clip con lo stesso nome {path.name}")
        by_name[path.name] = path

    map_path = clip_dir / "rinomina-vecchio-nuovo.tsv"
    undo_path = clip_dir / "annulla-rinomina.sh"
    if not dry_run and map_path.is_file() and _already_renamed(clip_dir, map_path, rows):
        write_log(log_path, _STEP, "ok", azione="gia-rinominato", righe=len(rows), orfani=0)
        return len(rows), 0

    missing = []
    occupied = []
    pairs = []
    for row in rows:
        source = by_name.get(row["attuale"])
        if source is None:
            missing.append(row["attuale"])
            continue
        if clip_clock(source.name) != row["key"]:
            _fail(log_path, f"orario diverso dal nome: {row['attuale']} / {row['orario']}")
        target = clip_dir / row["nuovo"]
        if target.exists() and target != source:
            occupied.append(row["nuovo"])
        pairs.append((source, row["nuovo"]))
    if missing or occupied:
        details = [f"MANCA: {name}" for name in missing] + [f"ESISTE GIA: {name}" for name in occupied]
        _fail(log_path, "PRECHECK FALLITO, nessuna rinomina. " + "; ".join(details))

    used = {source.name for source, _ in pairs}
    orphans = sorted(path.name for path in clips if path.name not in used)
    if orphans:
        _fail(log_path, "clip senza riga nel CSV: " + ", ".join(orphans))
    targets = [target for _, target in pairs]
    if len(targets) != len(set(targets)):
        _fail(log_path, "nomi nuovi duplicati nel CSV")

    if dry_run:
        write_log(log_path, _STEP, "ok", azione="dry-run", righe=len(pairs), orfani=0)
        return len(pairs), 0

    staged = []
    try:
        for source, target_name in pairs:
            target = clip_dir / target_name
            temporary = clip_dir / f".rinomina-{source.name}"
            source.rename(temporary)
            staged.append((temporary, target, source.name, target_name))
        lines = ["vecchio\tnuovo"]
        for temporary, target, old_name, new_name in staged:
            temporary.rename(target)
            lines.append(f"{old_name}\t{new_name}")
        map_path.write_text("\n".join(lines) + "\n", encoding="utf-8")
        _write_undo(undo_path, map_path.name)
    except Exception:
        for temporary, target, old_name, _new_name in reversed(staged):
            if temporary.exists() and not (clip_dir / old_name).exists():
                temporary.rename(clip_dir / old_name)
            elif target.exists() and target.name != old_name and not (clip_dir / old_name).exists():
                target.rename(clip_dir / old_name)
        write_log(log_path, _STEP, "errore", messaggio="rinomina annullata")
        raise
    write_log(log_path, _STEP, "ok", righe=len(pairs), orfani=0, mappa=str(map_path))
    return len(pairs), 0


def _read_table(path: Path) -> list[dict[str, str]]:
    if not path.is_file():
        raise StepError(_STEP, f"CSV assente: {path}")
    with path.open(encoding="utf-8", newline="") as handle:
        reader = csv.DictReader(handle)
        if reader.fieldnames is None or any(column not in reader.fieldnames for column in _COLUMNS):
            raise StepError(_STEP, "il CSV deve avere le colonne: " + ", ".join(_COLUMNS))
        rows = []
        for raw in reader:
            attuale = (raw.get("nome attuale") or "").strip()
            nuovo = (raw.get("nome nuovo") or "").strip()
            orario = (raw.get("orario video Veo") or "").strip()
            if not attuale or not nuovo or not orario:
                raise StepError(_STEP, "riga senza nome attuale, nome nuovo o orario")
            if clip_clock(attuale) is None:
                raise StepError(_STEP, f"nome attuale non è una clip: {attuale}")
            rows.append(
                {
                    "attuale": attuale,
                    "nuovo": nuovo,
                    "orario": orario,
                    "key": clock_key(orario),
                    "nota": (raw.get("nota Veo") or "").strip(),
                }
            )
    if not rows:
        raise StepError(_STEP, "CSV senza righe")
    keys = [row["key"] for row in rows]
    if len(keys) != len(set(keys)):
        raise StepError(_STEP, "orari di inizio duplicati nel CSV")
    names = [row["attuale"] for row in rows]
    if len(names) != len(set(names)):
        raise StepError(_STEP, "nomi attuali duplicati nel CSV")
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


def _write_undo(path: Path, map_name: str) -> None:
    path.write_text(
        "\n".join(
            [
                "#!/bin/bash",
                "export LC_ALL=en_US.UTF-8",
                'cd "$(dirname "$0")" || exit 2',
                "while IFS=$'\\t' read -r a b; do [ \"$a\" = vecchio ] && continue",
                '  if [ -f "$b" ] && [ ! -e "$a" ]; then mv -n "$b" "$a" && echo "ripristinato: $a"; else echo "SALTATO: $b"; fi',
                f"done < {map_name}",
                "",
            ]
        ),
        encoding="utf-8",
    )
    path.chmod(0o755)


def _fail(log_path: Path, message: str) -> None:
    write_log(log_path, _STEP, "errore", messaggio=message)
    raise StepError(_STEP, message)
