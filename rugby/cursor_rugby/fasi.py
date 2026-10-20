"""Passaggio 6. Scrive orari, inizio e fine, da/dove e nota.

Metri ed Esito restano vuoti: li compila Coach Rugby.
"""

from __future__ import annotations

import csv
from pathlib import Path

from rugby.cursor_rugby.clock import format_clock, game_minute, parse_clock, parse_half
from rugby.cursor_rugby.errors import StepError
from rugby.cursor_rugby.log import write_log
from rugby.cursor_rugby.xlsxutil import (
    Workbook,
    formula_cell,
    number_cell,
    parse_rows,
    render_row,
    replace_sheet_data,
    set_formula,
    shared_cell,
    style_of,
    update_dimension,
)

_STEP = "6"
_COLUMNS = ("inizio", "tempo", "comincia", "da", "palla", "finisce", "dove")


def fill_fasi(
    workbook: Path,
    table: Path,
    durations: dict[str, int],
    log_path: Path,
    template: Path,
    ripresa_minuto: int = 40,
) -> None:
    rows = _read_rows(table, durations, ripresa_minuto)
    _check_minutes(rows)
    book = Workbook(workbook)
    sheet = book.sheet_xml("Fasi PI")
    parsed = parse_rows(sheet)
    header = next(row for row in parsed if row["row"] == 1)
    sample = next((row for row in parsed if row["row"] == 2), header)
    styles = {column: style_of(sample, column) for column in "ABCDEFGHIJKLMN"}
    header_styles = {column: style_of(header, column) for column in "ABCDEFGHIJKLM"}
    judgments = _judgments(parsed)

    built = [_header_row(header, header_styles, book, any(row["nota"] for row in rows))]
    for index, row in enumerate(rows, start=2):
        built.append(_data_row(index, row, styles, book, judgments.get(row["n"])))
    sum_row = len(rows) + 2
    built.append(
        render_row(
            sum_row,
            ' x14ac:dyDescent="0.2"',
            [("F", formula_cell("F", sum_row, f"SUM(F2:F{sum_row - 1})", str(sum(row["durata"] for row in rows)), styles["F"]))],
        )
    )
    updated = replace_sheet_data(sheet, "".join(built))
    last_column = "N" if any(row["nota"] for row in rows) else "M"
    updated = update_dimension(updated, f"A1:{last_column}{sum_row}")
    book.set_sheet_xml("Fasi PI", updated)
    summary = set_formula(book.sheet_xml("Riassunto"), "F6", f"'Fasi PI'!F{sum_row}/60")
    book.set_sheet_xml("Riassunto", summary)
    _retarget_sum_chain(book, f"F{sum_row}")
    book.save(workbook)
    from rugby.cursor_rugby.xlsxutil import three_checks

    three_checks(workbook, template)
    write_log(log_path, _STEP, "ok", fasi=len(rows), somma_secondi=sum(row["durata"] for row in rows))


def _read_rows(table: Path, durations: dict[str, int], ripresa_minuto: int) -> list[dict]:
    if not table.is_file():
        raise StepError(_STEP, f"CSV assente: {table}")
    with table.open(encoding="utf-8", newline="") as handle:
        reader = csv.DictReader(handle)
        missing = [name for name in _COLUMNS if reader.fieldnames is None or name not in reader.fieldnames]
        if missing:
            raise StepError(_STEP, "colonne mancanti nel CSV: " + ", ".join(missing))
        raw_rows = list(reader)
    if not raw_rows:
        raise StepError(_STEP, "CSV fasi vuoto")
    second_start = None
    prepared = []
    for raw in raw_rows:
        half = parse_half(raw["tempo"])
        start = parse_clock(raw["inizio"])
        if half == 2 and second_start is None:
            second_start = start
        prepared.append((raw, half, start))
    rows = []
    for number, (raw, half, start) in enumerate(prepared, start=1):
        key = f"{start // 3600:02d}{(start % 3600) // 60:02d}{start % 60:02d}"
        duration = durations.get(key)
        if duration is None and (raw.get("durata_s") or "").strip():
            duration = int(raw["durata_s"])
        if duration is None:
            raise StepError(_STEP, f"durata assente per l'orario {raw['inizio']}")
        end = start + duration
        if end <= start:
            raise StepError(_STEP, f"la fine non è dopo l'inizio: {raw['inizio']}")
        minuto = game_minute(raw["inizio"], half, second_start, ripresa_minuto)
        rows.append(
            {
                "n": number,
                "tempo": "1°" if half == 1 else "2°",
                "inizio": format_clock(start),
                "fine": format_clock(end),
                "minuto": minuto,
                "durata": duration,
                "comincia": raw["comincia"].strip(),
                "da": raw["da"].strip(),
                "palla": raw["palla"].strip(),
                "finisce": raw["finisce"].strip(),
                "dove": raw["dove"].strip(),
                "nota": (raw.get("nota") or "").strip(),
            }
        )
    return rows


def _check_minutes(rows: list[dict]) -> None:
    previous = -1
    for row in rows:
        if row["minuto"] < previous:
            raise StepError(_STEP, f"minuti non crescenti alla fase {row['n']}")
        previous = row["minuto"]


def _judgments(parsed: list[dict]) -> dict[int, dict[str, str]]:
    kept = {}
    for row in parsed:
        if row["row"] < 2:
            continue
        values = {}
        for cell in row["cells"]:
            if cell["column"] in {"L", "M"} and (cell["inner"] or 't="' in cell["attrs"]):
                values[cell["column"]] = f"<c{cell['attrs']}>{cell['inner']}</c>" if cell["inner"] else f"<c{cell['attrs']}/>"
        number = next((cell["value"] for cell in row["cells"] if cell["column"] == "A" and not cell["shared"]), None)
        if number and str(number).isdigit() and values:
            kept[int(number)] = values
    return kept


def _header_row(header: dict, styles: dict[str, str | None], book: Workbook, with_note: bool) -> str:
    cells = []
    labels = {
        "A": "N",
        "B": "Tempo",
        "C": "Inizio",
        "D": "Fine",
        "E": "Minuto",
        "F": "Durata (s)",
        "G": "Comincia con",
        "H": "Da dove",
        "I": "Palla a",
        "J": "Finisce con",
        "K": "Dove",
        "L": "Esito",
        "M": "Metri",
    }
    existing = {cell["column"]: cell["inner"] for cell in header["cells"]}
    for column, label in labels.items():
        if column in existing and existing[column]:
            attrs_inner = existing[column]
            # keep the original header cell, retargeted to row 1
            ref = f'{column}1'
            cell_xml = f'<c r="{ref}"' + _attrs_without_ref(header, column) + f">{attrs_inner}</c>"
            cells.append((column, cell_xml))
        else:
            cells.append((column, shared_cell(column, 1, book.intern(label), styles.get(column))))
    if with_note:
        cells.append(("N", shared_cell("N", 1, book.intern("Nota"), styles.get("N") or styles.get("M"))))
    return render_row(1, header["attrs"], cells)


def _attrs_without_ref(row: dict, column: str) -> str:
    for cell in row["cells"]:
        if cell["column"] == column:
            return re_sub_ref(cell["attrs"])
    return ""


def re_sub_ref(attrs: str) -> str:
    import re

    return re.sub(r'\s*r="[A-Z]+\d+"', "", attrs)


def _data_row(number: int, row: dict, styles: dict[str, str | None], book: Workbook, judgment: dict[str, str] | None) -> str:
    cells = [
        ("A", number_cell("A", number, str(row["n"]), styles["A"])),
        ("B", shared_cell("B", number, book.intern(row["tempo"]), styles["B"])),
        ("C", shared_cell("C", number, book.intern(row["inizio"]), styles["C"])),
        ("D", shared_cell("D", number, book.intern(row["fine"]), styles["D"])),
        ("E", shared_cell("E", number, book.intern(f"{row['minuto']}'"), styles["E"])),
        ("F", number_cell("F", number, str(row["durata"]), styles["F"])),
        ("G", shared_cell("G", number, book.intern(row["comincia"]), styles["G"])),
        ("H", shared_cell("H", number, book.intern(row["da"]), styles["H"])),
        ("I", shared_cell("I", number, book.intern(row["palla"]), styles["I"])),
        ("J", shared_cell("J", number, book.intern(row["finisce"]), styles["J"])),
        ("K", shared_cell("K", number, book.intern(row["dove"]), styles["K"])),
    ]
    if judgment:
        for column in ("L", "M"):
            previous = judgment.get(column)
            if previous:
                cells.append((column, _retarget(column, number, previous)))
    if row["nota"]:
        cells.append(("N", shared_cell("N", number, book.intern(row["nota"]), styles.get("N"))))
    return render_row(number, ' spans="1:13" ht="16" x14ac:dyDescent="0.2"', cells)


def _retarget(column: str, number: int, cell_xml: str) -> str:
    import re

    return re.sub(r'\br="[A-Z]+\d+"', f'r="{column}{number}"', cell_xml, count=1)


def _retarget_sum_chain(book: Workbook, new_ref: str) -> None:
    import re

    raw = book.parts.get("xl/calcChain.xml")
    if raw is None:
        return
    text = raw.decode("utf-8")
    updated = re.sub(r'<c r="F\d+" i="4"', f'<c r="{new_ref}" i="4"', text, count=1)
    book.parts["xl/calcChain.xml"] = updated.encode("utf-8")
