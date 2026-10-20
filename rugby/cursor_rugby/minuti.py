"""Passaggio 11. Somma i minuti dei giocatori e controlla che il totale torni."""

from __future__ import annotations

import csv
import re
from pathlib import Path

from rugby.cursor_rugby.errors import StepError
from rugby.cursor_rugby.log import write_log

_STEP = "11"
_SUB = re.compile(r"^(?P<name>.+?)\s*\((?P<minute>\d+)\s*[’'`']\s*(?P<sub>[^)]+)\)$")


def summarize(sources: list[Path], dest: Path, log_path: Path, match_minutes: int = 80) -> Path:
    if not sources:
        _fail(log_path, "nessun file partita")
    totals: dict[str, int] = {}
    match_totals = []
    for source in sources:
        played = _read_source(Path(source), match_minutes)
        match_totals.append(sum(played.values()))
        for name, minutes in played.items():
            totals[name] = totals.get(name, 0) + minutes
    grand = sum(totals.values())
    if grand != sum(match_totals):
        _fail(log_path, f"totale {grand} diverso dalla somma delle partite {sum(match_totals)}")
    lines = ["giocatore\tminuti"]
    for name in sorted(totals):
        lines.append(f"{name}\t{totals[name]}")
    text = "\n".join(lines) + "\n"
    dest = Path(dest)
    if dest.is_file() and dest.read_text(encoding="utf-8") == text:
        write_log(log_path, _STEP, "ok", azione="gia-presente", giocatori=len(totals), minuti=grand)
        return dest
    dest.parent.mkdir(parents=True, exist_ok=True)
    temporary = dest.with_suffix(dest.suffix + ".tmp")
    temporary.write_text(text, encoding="utf-8")
    temporary.replace(dest)
    write_log(log_path, _STEP, "ok", giocatori=len(totals), minuti=grand, file=str(dest))
    return dest


def parse_formation(line: str, match_minutes: int = 80) -> dict[str, int]:
    body = line.split(":", 1)[-1]
    body = re.split(r"\.\s*All\b", body, maxsplit=1)[0]
    body = body.strip().strip(".")
    played: dict[str, int] = {}
    starters = 0
    for token in _split_players(body):
        starters += 1
        match = _SUB.match(token)
        if match:
            name = _clean(match.group("name"))
            minute = int(match.group("minute"))
            sub = _clean(match.group("sub"))
            if minute > match_minutes:
                raise StepError(_STEP, f"cambio oltre la fine partita: {token}")
            _add(played, name, minute)
            _add(played, sub, match_minutes - minute)
        else:
            _add(played, _clean(token), match_minutes)
    if starters != 15:
        raise StepError(_STEP, f"la formazione ha {starters} titolari, non 15")
    if sum(played.values()) != 15 * match_minutes:
        raise StepError(_STEP, "la somma dei minuti non è 15 per la durata della partita")
    return played


def _read_source(path: Path, match_minutes: int) -> dict[str, int]:
    if not path.is_file():
        raise StepError(_STEP, f"file assente: {path}")
    if path.suffix.lower() == ".csv":
        return _read_csv(path, match_minutes)
    text = path.read_text(encoding="utf-8")
    for line in text.splitlines():
        if line.lower().startswith("montebelluna") and ":" in line:
            return parse_formation(line, match_minutes)
    raise StepError(_STEP, f"nessuna formazione del Montebelluna in {path.name}")


def _read_csv(path: Path, match_minutes: int) -> dict[str, int]:
    with path.open(encoding="utf-8", newline="") as handle:
        reader = csv.DictReader(handle)
        if reader.fieldnames is None or not {"giocatore", "minuti"} <= set(reader.fieldnames):
            raise StepError(_STEP, "il CSV minuti deve avere giocatore e minuti")
        played: dict[str, int] = {}
        for raw in reader:
            name = _clean(raw.get("giocatore") or "")
            if not name:
                raise StepError(_STEP, "riga minuti senza giocatore")
            _add(played, name, int(raw["minuti"]))
    if sum(played.values()) != 15 * match_minutes:
        raise StepError(_STEP, f"i minuti di {path.name} non sommano a {15 * match_minutes}")
    return played


def _split_players(body: str) -> list[str]:
    parts: list[str] = []
    buf = []
    depth = 0
    for char in body:
        if char == "(":
            depth += 1
        elif char == ")":
            depth = max(0, depth - 1)
        if char == "," and depth == 0:
            parts.append("".join(buf).strip())
            buf = []
        else:
            buf.append(char)
    tail = "".join(buf).strip()
    if tail:
        parts.append(tail)
    return [part for part in parts if part]


def _clean(name: str) -> str:
    return re.sub(r"\s+", " ", name).strip(" .")


def _add(played: dict[str, int], name: str, minutes: int) -> None:
    if not name:
        raise StepError(_STEP, "nome giocatore vuoto")
    played[name] = played.get(name, 0) + minutes


def _fail(log_path: Path, message: str) -> None:
    write_log(log_path, _STEP, "errore", messaggio=message)
    raise StepError(_STEP, message)
