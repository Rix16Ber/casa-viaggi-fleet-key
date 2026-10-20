"""Passaggio 10. Testo a tabulazioni per il foglio Partite, con lo stesso numero di colonne."""

from __future__ import annotations

from pathlib import Path

from rugby.cursor_rugby.errors import StepError
from rugby.cursor_rugby.log import write_log
from rugby.cursor_rugby.xlsxutil import Workbook, parse_rows, row_map

_STEP = "10"
COLONNE = (
    "data",
    "avversario",
    "punti_montebelluna",
    "punti_avversario",
    "mete_montebelluna",
    "mete_avversario",
    "trasf_ok_montebelluna",
    "trasf_ok_avversario",
    "trasf_no_montebelluna",
    "trasf_no_avversario",
    "piazzati_ok_montebelluna",
    "piazzati_ok_avversario",
    "piazzati_no_montebelluna",
    "piazzati_no_avversario",
    "touche_lanciate_montebelluna",
    "touche_vinte_montebelluna",
    "touche_perse_montebelluna",
    "touche_rubate_montebelluna",
    "touche_lanciate_avversario",
    "touche_vinte_avversario",
    "touche_perse_avversario",
    "touche_rubate_avversario",
    "mischie_montebelluna",
    "mischie_vinte_montebelluna",
    "mischie_avversario",
    "mischie_vinte_avversario",
    "minuti",
)


def write_partite_row(
    workbook: Path,
    dest: Path,
    when: str,
    opponent: str,
    log_path: Path,
    intestazione: Path | None = None,
) -> Path:
    header = _header(intestazione)
    if len(header) != len(COLONNE):
        _fail(
            log_path,
            f"il foglio Partite ha {len(header)} colonne, il testo ne prepara {len(COLONNE)}",
        )
    if list(header) != list(COLONNE):
        _fail(log_path, "l'intestazione del foglio Partite non coincide con le colonne attese")
    values = _values(Path(workbook), when, opponent)
    if len(values) != len(header):
        _fail(log_path, "riga e intestazione hanno un numero di colonne diverso")
    line = "\t".join(header) + "\n" + "\t".join(values) + "\n"
    dest = Path(dest)
    if dest.is_file() and dest.read_text(encoding="utf-8") == line:
        write_log(log_path, _STEP, "ok", azione="gia-presente", file=str(dest), colonne=len(header))
        return dest
    dest.parent.mkdir(parents=True, exist_ok=True)
    temporary = dest.with_suffix(dest.suffix + ".tmp")
    temporary.write_text(line, encoding="utf-8")
    written = temporary.read_text(encoding="utf-8").splitlines()
    if any(len(row.split("\t")) != len(header) for row in written):
        temporary.unlink(missing_ok=True)
        _fail(log_path, "controllo colonne fallito dopo la scrittura")
    temporary.replace(dest)
    write_log(log_path, _STEP, "ok", file=str(dest), colonne=len(header))
    return dest


def _header(path: Path | None) -> tuple[str, ...]:
    if path is None:
        return COLONNE
    if not path.is_file():
        raise StepError(_STEP, f"intestazione assente: {path}")
    first = path.read_text(encoding="utf-8").splitlines()[0]
    return tuple(column.strip() for column in first.split("\t"))


def _values(workbook: Path, when: str, opponent: str) -> list[str]:
    book = Workbook(workbook)
    strings = book.strings()
    summary = {row["row"]: row_map(row, strings) for row in parse_rows(book.sheet_xml("Riassunto"))}
    monte = summary.get(9, {})
    other = summary.get(10, {})
    touche_m = summary.get(14, {})
    touche_o = summary.get(15, {})
    scrum_m = summary.get(19, {})
    scrum_o = summary.get(20, {})
    monte_points = _points(monte)
    other_points = _points(other)
    minutes = _minutes(book)
    return [
        when,
        opponent,
        str(monte_points),
        str(other_points),
        _num(monte, "C"),
        _num(other, "C"),
        _num(monte, "D"),
        _num(other, "D"),
        _num(monte, "E"),
        _num(other, "E"),
        _num(monte, "F"),
        _num(other, "F"),
        _num(monte, "G"),
        _num(other, "G"),
        _num(touche_m, "C"),
        _num(touche_m, "D"),
        _num(touche_m, "E"),
        _num(touche_m, "F"),
        _num(touche_o, "C"),
        _num(touche_o, "D"),
        _num(touche_o, "E"),
        _num(touche_o, "F"),
        _num(scrum_m, "C"),
        _num(scrum_m, "D"),
        _num(scrum_o, "C"),
        _num(scrum_o, "D"),
        minutes,
    ]


def _points(row: dict[str, str]) -> int:
    return int(_num(row, "C")) * 5 + int(_num(row, "D")) * 2 + int(_num(row, "F")) * 3


def _num(row: dict[str, str], column: str) -> str:
    text = (row.get(column) or "").strip()
    if not text:
        return "0"
    try:
        return str(int(float(text)))
    except ValueError as exc:
        raise StepError(_STEP, f"numero non valido in colonna {column}: {text}") from exc


def _minutes(book: Workbook) -> str:
    total = 0
    for row in parse_rows(book.sheet_xml("Fasi PI")):
        for cell in row["cells"]:
            if cell["column"] == "F" and cell["formula"] is None and cell["value"] and not cell["shared"]:
                if row["row"] == 1:
                    continue
                total += int(float(cell["value"]))
    return str(round(total / 60, 1))


def _fail(log_path: Path, message: str) -> None:
    write_log(log_path, _STEP, "errore", messaggio=message)
    raise StepError(_STEP, message)
