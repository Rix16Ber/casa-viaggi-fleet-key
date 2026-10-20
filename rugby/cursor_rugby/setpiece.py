"""Passaggio 7. Scrive solo gli eventi veri di touche e mischia, con la squadra."""

from __future__ import annotations

import csv
import re
from pathlib import Path

from rugby.cursor_rugby.errors import StepError
from rugby.cursor_rugby.log import write_log
from rugby.cursor_rugby.xlsxutil import (
    Workbook,
    parse_rows,
    render_row,
    replace_sheet_data,
    shared_cell,
    style_of,
    update_dimension,
)

_STEP = "7"
_REAL = {"touche", "mischia"}


def fill_set_piece(workbook: Path, table: Path, log_path: Path, template: Path) -> None:
    events = _read_events(table)
    book = Workbook(workbook)
    sheet = book.sheet_xml("Set piece")
    parsed = parse_rows(sheet)
    strings = book.strings()
    preamble, mid, sample = _split(parsed, strings)
    styles = {column: style_of(sample, column) for column in "ABCDE"}
    number = 1
    touche = [row for row in events if row["tipo"] == "touche"]
    scrum = [row for row in events if row["tipo"] == "mischia"]
    rendered = []
    mapping = {}
    for row in preamble:
        mapping[row["row"]] = number
        rendered.append(_renumber(row, number))
        number += 1
    first_touche = number
    for event in touche:
        rendered.append(_event_row(number, event, styles, book))
        number += 1
    last_touche = number - 1
    for row in mid:
        mapping[row["row"]] = number
        rendered.append(_renumber(row, number))
        number += 1
    for event in scrum:
        rendered.append(_event_row(number, event, styles, book))
        number += 1
    updated = replace_sheet_data(sheet, "".join(rendered))
    updated = _remap_merges(updated, mapping)
    updated = update_dimension(updated, f"A1:E{number - 1}")
    book.set_sheet_xml("Set piece", updated)
    book.save(workbook)
    from rugby.cursor_rugby.xlsxutil import three_checks

    three_checks(workbook, template)
    write_log(
        log_path,
        _STEP,
        "ok",
        touche=len(touche),
        mischie=len(scrum),
        prima=first_touche,
        ultima_touche=last_touche,
    )


def _read_events(table: Path) -> list[dict[str, str]]:
    if not table.is_file():
        raise StepError(_STEP, f"CSV assente: {table}")
    with table.open(encoding="utf-8", newline="") as handle:
        reader = csv.DictReader(handle)
        if reader.fieldnames is None or not {"tipo", "orario", "squadra"} <= set(reader.fieldnames):
            raise StepError(_STEP, "il CSV deve avere tipo, orario e squadra")
        events = []
        for raw in reader:
            kind = (raw.get("tipo") or "").strip().lower()
            if kind not in _REAL:
                continue
            team = (raw.get("squadra") or "").strip()
            if not team:
                raise StepError(_STEP, f"evento {raw.get('tipo')} alle {raw.get('orario')} senza squadra")
            clock = (raw.get("orario") or "").strip()
            if not clock:
                raise StepError(_STEP, "evento senza orario")
            events.append(
                {
                    "tipo": kind,
                    "etichetta": "Touche" if kind == "touche" else "Mischia",
                    "orario": clock,
                    "squadra": team,
                    "nota": (raw.get("nota") or "").strip(),
                    "esito": (raw.get("esito") or "").strip(),
                }
            )
    if not events:
        raise StepError(_STEP, "nessun evento vero nel CSV")
    return events


def _split(parsed: list[dict], strings: list[str | None]) -> tuple[list[dict], list[dict], dict]:
    from rugby.cursor_rugby.xlsxutil import cell_text

    tipo_rows = []
    for row in parsed:
        for cell in row["cells"]:
            if cell["column"] == "A" and cell_text(cell, strings) == "Tipo":
                tipo_rows.append(row["row"])
    if len(tipo_rows) < 2:
        raise StepError(_STEP, "nel foglio Set piece non ci sono le due intestazioni Tipo")
    first, second = tipo_rows[0], tipo_rows[1]
    preamble = [row for row in parsed if row["row"] <= first]
    mid = [row for row in parsed if first < row["row"] <= second and _is_structural(row, strings)]
    sample = next((row for row in parsed if first < row["row"] < second), preamble[-1])
    return preamble, mid, sample


def _is_structural(row: dict, strings: list[str | None]) -> bool:
    from rugby.cursor_rugby.xlsxutil import cell_text

    text = ""
    for cell in row["cells"]:
        if cell["column"] == "A":
            text = cell_text(cell, strings) or ""
    if text in {"Tipo", "MISCHIE", "TOUCHE"} or text.startswith("MISCHIE") or text.startswith("SET PIECE"):
        return True
    # section titles are long notes or headings stored in A
    return text.startswith("MISCHIE —") or text.startswith("MISCHIE -") or "clip Scrum" in text


def _event_row(number: int, event: dict[str, str], styles: dict[str, str | None], book: Workbook) -> str:
    cells = [
        ("A", shared_cell("A", number, book.intern(event["etichetta"]), styles["A"])),
        ("B", shared_cell("B", number, book.intern(event["orario"]), styles["B"])),
        ("C", shared_cell("C", number, book.intern(event["squadra"]), styles["C"])),
        ("D", shared_cell("D", number, book.intern(event["nota"]), styles["D"])),
        ("E", shared_cell("E", number, book.intern(event["esito"]), styles["E"])),
    ]
    return render_row(number, ' x14ac:dyDescent="0.2"', cells)


def _renumber(row: dict, number: int) -> str:
    cells = []
    for cell in row["cells"]:
        attrs = re.sub(r'\br="[A-Z]+\d+"', f'r="{cell["column"]}{number}"', cell["attrs"], count=1)
        if cell["inner"]:
            cells.append((cell["column"], f"<c{attrs}>{cell['inner']}</c>"))
        else:
            cells.append((cell["column"], f"<c{attrs}/>"))
    return render_row(number, row["attrs"], cells)


def _remap_merges(sheet_xml: str, mapping: dict[int, int]) -> str:
    def swap(match: re.Match[str]) -> str:
        ref = match.group(1)
        start, _, end = ref.partition(":")
        start_row = int(re.search(r"\d+", start).group(0))
        end_row = int(re.search(r"\d+", end).group(0)) if end else start_row
        if start_row not in mapping or end_row not in mapping:
            return ""
        start_col = re.match(r"[A-Z]+", start).group(0)
        end_col = re.match(r"[A-Z]+", end).group(0) if end else start_col
        return f'<mergeCell ref="{start_col}{mapping[start_row]}:{end_col}{mapping[end_row]}"/>'

    return re.sub(r'<mergeCell ref="([^"]+)"\s*/>', swap, sheet_xml)
