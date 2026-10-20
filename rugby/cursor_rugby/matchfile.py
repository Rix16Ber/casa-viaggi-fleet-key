"""Passaggio 5. Crea il file partita dal modello, senza toccare i grafici."""

from __future__ import annotations

import shutil
from datetime import date
from pathlib import Path

from rugby.cursor_rugby.errors import StepError
from rugby.cursor_rugby.log import write_log
from rugby.cursor_rugby.names import workbook_name
from rugby.cursor_rugby.xlsxutil import (
    Workbook,
    blank_rows,
    cell_text,
    chart_hashes,
    parse_rows,
    row_map,
    set_shared_cell,
    three_checks,
)

_STEP = "5"
_OPPONENT_CELLS = ("G5", "B10", "B15", "B20")
_CLEAR_SUMMARY = ("D5", "F5", "C9", "D9", "E9", "F9", "G9", "C10", "D10", "E10", "F10", "G10")


def create_match_file(template: Path, dest_dir: Path, opponent: str, when: date, log_path: Path) -> Path:
    template = Path(template)
    if not template.is_file():
        _fail(log_path, f"modello assente: {template}")
    dest = Path(dest_dir) / workbook_name(when, opponent)
    title = f"Montebelluna – {opponent.strip()} · {when:%d/%m/%Y}"
    if dest.is_file():
        if _already(dest, template, title):
            write_log(log_path, _STEP, "ok", azione="gia-presente", file=str(dest))
            return dest
        _fail(log_path, f"esiste già un file diverso, non lo sovrascrivo: {dest.name}")

    dest.parent.mkdir(parents=True, exist_ok=True)
    shutil.copy2(template, dest)
    book = Workbook(dest)
    summary = book.sheet_xml("Riassunto")
    summary = set_shared_cell(summary, "B1", book.intern(title))
    for ref in _OPPONENT_CELLS:
        summary = set_shared_cell(summary, ref, book.intern(opponent.strip()))
    for ref in _CLEAR_SUMMARY:
        summary = _blank_if_present(summary, ref)
    book.set_sheet_xml("Riassunto", summary)
    book.set_sheet_xml("Set piece", _blank_named_rows(book, "Set piece", {"Touche", "Mischia"}))
    book.set_sheet_xml("Eventi", _blank_numbered_rows(book, "Eventi"))
    book.set_sheet_xml("Fasi PI", _blank_phase_rows(book))
    book.save(dest)
    three_checks(dest, template)
    write_log(log_path, _STEP, "ok", file=str(dest))
    return dest


def _already(path: Path, template: Path, title: str) -> bool:
    if chart_hashes(path) != chart_hashes(template):
        return False
    book = Workbook(path)
    strings = book.strings()
    for row in parse_rows(book.sheet_xml("Riassunto")):
        if row["row"] == 1:
            return row_map(row, strings).get("B") == title
    return False


def _blank_if_present(sheet_xml: str, ref: str) -> str:
    from rugby.cursor_rugby.xlsxutil import clear_value

    if f'r="{ref}"' not in sheet_xml:
        return sheet_xml
    try:
        return clear_value(sheet_xml, ref)
    except StepError:
        return sheet_xml


def _blank_named_rows(book: Workbook, sheet: str, labels: set[str]) -> str:
    xml = book.sheet_xml(sheet)
    strings = book.strings()
    rows = {
        row["row"]
        for row in parse_rows(xml)
        if row_map(row, strings).get("A") in labels
    }
    return blank_rows(xml, rows)


def _blank_numbered_rows(book: Workbook, sheet: str) -> str:
    xml = book.sheet_xml(sheet)
    strings = book.strings()
    rows = set()
    for row in parse_rows(xml):
        marker = row_map(row, strings).get("A", "")
        if marker.isdigit():
            rows.add(row["row"])
    return blank_rows(xml, rows)


def _blank_phase_rows(book: Workbook) -> str:
    xml = book.sheet_xml("Fasi PI")
    rows = set()
    for row in parse_rows(xml):
        if row["row"] == 1 or any(cell["formula"] for cell in row["cells"]):
            continue
        rows.add(row["row"])
    return blank_rows(xml, rows)


def _fail(log_path: Path, message: str) -> None:
    write_log(log_path, _STEP, "errore", messaggio=message)
    raise StepError(_STEP, message)


def title_of(path: Path) -> str | None:
    book = Workbook(path)
    strings = book.strings()
    for row in parse_rows(book.sheet_xml("Riassunto")):
        if row["row"] == 1:
            for cell in row["cells"]:
                if cell["column"] == "B":
                    return cell_text(cell, strings)
    return None
