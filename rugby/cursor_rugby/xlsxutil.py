"""Lettura e modifica dei fogli xlsx via XML. Non usa openpyxl.

I grafici del modello sono un disegno più un'immagine sul foglio Riassunto.
Quei file nello zip non vengono riscritti.
"""

from __future__ import annotations

import hashlib
import re
import zipfile
from io import BytesIO
from pathlib import Path
from xml.sax.saxutils import escape, unescape

from rugby.cursor_rugby.errors import StepError

_STEP = "9"
_NS_T = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"
_CELL_RE = re.compile(r"<c\b([^>]*?)(?:/>|>(.*?)</c>)", re.S)
_ROW_RE = re.compile(r"<row\b([^>]*)>(.*?)</row>", re.S)
_SI_RE = re.compile(r"<si\b[^>]*>.*?</si>", re.S)
_SHEET_DATA_RE = re.compile(r"(<sheetData\b[^>]*>)(.*?)(</sheetData>)", re.S)
_CHART_PARTS = (
    "xl/drawings/drawing1.xml",
    "xl/media/image1.png",
    "xl/drawings/_rels/drawing1.xml.rels",
    "xl/worksheets/_rels/sheet1.xml.rels",
)
_SHEET_NAMES = ("Riassunto", "Set piece", "Eventi", "Fasi PI")


class Workbook:
    def __init__(self, path: Path):
        self.path = Path(path)
        if not self.path.is_file() or self.path.stat().st_size == 0:
            raise StepError(_STEP, f"file vuoto o assente: {self.path}")
        try:
            with zipfile.ZipFile(self.path) as archive:
                self.parts = {name: archive.read(name) for name in archive.namelist()}
        except zipfile.BadZipFile as exc:
            raise StepError(_STEP, f"il file non si apre: {self.path}") from exc

    def sheet_names(self) -> list[str]:
        xml = self.parts["xl/workbook.xml"].decode("utf-8")
        return re.findall(r'<sheet\b[^>]*\bname="([^"]+)"', xml)

    def sheet_xml(self, name: str) -> str:
        target = self._sheet_target(name)
        return self.parts[target].decode("utf-8")

    def set_sheet_xml(self, name: str, xml: str) -> None:
        self.parts[self._sheet_target(name)] = xml.encode("utf-8")

    def _sheet_target(self, name: str) -> str:
        workbook = self.parts["xl/workbook.xml"].decode("utf-8")
        rels = self.parts["xl/_rels/workbook.xml.rels"].decode("utf-8")
        match = re.search(rf'<sheet\b[^>]*\bname="{re.escape(name)}"[^>]*\br:id="([^"]+)"', workbook)
        if match is None:
            match = re.search(rf'<sheet\b[^>]*\br:id="([^"]+)"[^>]*\bname="{re.escape(name)}"', workbook)
        if match is None:
            raise StepError(_STEP, f"foglio assente: {name}")
        rel = re.search(
            rf'<Relationship\b[^>]*\bId="{re.escape(match.group(1))}"[^>]*\bTarget="([^"]+)"',
            rels,
        )
        if rel is None:
            raise StepError(_STEP, f"relazione assente per il foglio {name}")
        target = rel.group(1)
        if not target.startswith("xl/"):
            target = "xl/" + target.lstrip("/")
        return target

    def strings(self) -> list[str | None]:
        xml = self.parts.get("xl/sharedStrings.xml")
        if xml is None:
            return []
        values: list[str | None] = []
        for raw in _SI_RE.findall(xml.decode("utf-8")):
            text = re.search(r"<t(?:\s[^>]*)?>(.*?)</t>", raw, re.S)
            if text is None or "<r>" in raw:
                values.append(None)
            else:
                values.append(unescape(text.group(1)))
        return values

    def intern(self, text: str) -> int:
        xml = self.parts.get("xl/sharedStrings.xml")
        if xml is None:
            xml = (
                '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                '<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
                'count="0" uniqueCount="0"></sst>'
            ).encode("utf-8")
            self.parts["xl/sharedStrings.xml"] = xml
        current = self.strings()
        if text in current:
            return current.index(text)
        space = ' xml:space="preserve"' if text[:1].isspace() or text[-1:].isspace() else ""
        piece = f"<si><t{space}>{escape(text)}</t></si>"
        body = xml.decode("utf-8")
        if "</sst>" not in body:
            raise StepError(_STEP, "sharedStrings.xml senza chiusura")
        body = body.replace("</sst>", piece + "</sst>", 1)
        unique = len(current) + 1
        count_match = re.search(r'\bcount="(\d+)"', body)
        count = int(count_match.group(1)) + 1 if count_match else unique
        body = re.sub(r'\buniqueCount="\d+"', f'uniqueCount="{unique}"', body, count=1)
        body = re.sub(r'\bcount="\d+"', f'count="{count}"', body, count=1)
        self.parts["xl/sharedStrings.xml"] = body.encode("utf-8")
        return unique - 1

    def save(self, path: Path) -> None:
        path = Path(path)
        path.parent.mkdir(parents=True, exist_ok=True)
        buffer = BytesIO()
        with zipfile.ZipFile(buffer, "w", compression=zipfile.ZIP_DEFLATED) as archive:
            for name, payload in self.parts.items():
                archive.writestr(name, payload)
        temporary = path.with_suffix(path.suffix + ".tmp")
        temporary.write_bytes(buffer.getvalue())
        temporary.replace(path)
        self.path = path


def chart_hashes(path: Path) -> dict[str, str]:
    with zipfile.ZipFile(path) as archive:
        names = set(archive.namelist())
        found = {}
        for part in _CHART_PARTS:
            if part not in names:
                raise StepError(_STEP, f"grafico assente: {part}")
            found[part] = hashlib.sha256(archive.read(part)).hexdigest()
        return found


def three_checks(path: Path, template: Path) -> list[str]:
    """Tre controlli: apertura, byte dei grafici, formule ancora presenti."""
    notes = []
    book = Workbook(path)
    names = book.sheet_names()
    if tuple(names) != _SHEET_NAMES:
        raise StepError(_STEP, f"fogli diversi dal modello: {names}")
    if path.stat().st_size == 0:
        raise StepError(_STEP, "file da 0 byte")
    notes.append("apertura ok")

    current = chart_hashes(path)
    expected = chart_hashes(template)
    if current != expected:
        raise StepError(_STEP, "i byte di disegno o immagine non coincidono col modello")
    notes.append("grafici invariati")

    summary = book.sheet_xml("Riassunto")
    phases = book.sheet_xml("Fasi PI")
    if "'Fasi PI'!F" not in summary or "<f>SUM(F2:F" not in phases:
        raise StepError(_STEP, "le formule di minuti o della somma fasi non ci sono più")
    notes.append("formule presenti")
    return notes


def repeat_checks(path: Path, template: Path, times: int = 3) -> None:
    if times != 3:
        raise StepError(_STEP, "i controlli del caricamento sono tre")
    for _ in range(times):
        three_checks(path, template)


def parse_rows(sheet_xml: str) -> list[dict]:
    rows = []
    for attrs, inner in _ROW_RE.findall(sheet_xml):
        number = re.search(r'\br="(\d+)"', attrs)
        if number is None:
            raise StepError(_STEP, "riga senza numero")
        cells = []
        for cell_attrs, cell_inner in _CELL_RE.findall(inner):
            ref = re.search(r'\br="([A-Z]+)(\d+)"', cell_attrs)
            if ref is None:
                raise StepError(_STEP, "cella senza riferimento")
            value = re.search(r"<v>(.*?)</v>", cell_inner or "", re.S)
            formula = re.search(r"<f[^>]*>(.*?)</f>", cell_inner or "", re.S)
            cells.append(
                {
                    "column": ref.group(1),
                    "attrs": cell_attrs,
                    "inner": cell_inner or "",
                    "value": unescape(value.group(1)) if value else None,
                    "formula": formula.group(1) if formula else None,
                    "shared": 't="s"' in cell_attrs,
                }
            )
        rows.append({"row": int(number.group(1)), "attrs": attrs, "cells": cells, "inner": inner})
    return rows


def cell_text(cell: dict, strings: list[str | None]) -> str | None:
    if cell["value"] is None:
        return None
    if cell["shared"]:
        index = int(cell["value"])
        if index >= len(strings):
            return None
        return strings[index]
    return cell["value"]


def row_map(row: dict, strings: list[str | None]) -> dict[str, str]:
    found = {}
    for cell in row["cells"]:
        text = cell_text(cell, strings)
        if text is not None:
            found[cell["column"]] = text
    return found


def replace_sheet_data(sheet_xml: str, rows_xml: str) -> str:
    match = _SHEET_DATA_RE.search(sheet_xml)
    if match is None:
        raise StepError(_STEP, "sheetData assente")
    return sheet_xml[: match.start(2)] + rows_xml + sheet_xml[match.end(2) :]


def render_row(row_number: int, extra_attrs: str, cells: list[tuple[str, str]]) -> str:
    """cells: (column, inner-or-self xml starting at the c tag)."""
    body = "".join(cells_xml for _, cells_xml in cells)
    attrs = re.sub(r'\br="\d+"', f'r="{row_number}"', extra_attrs, count=1)
    if 'r="' not in attrs:
        attrs = f' r="{row_number}"' + attrs
    return f"<row{attrs}>{body}</row>"


def shared_cell(column: str, row_number: int, index: int, style: str | None) -> str:
    style_attr = f' s="{style}"' if style else ""
    return f'<c r="{column}{row_number}"{style_attr} t="s"><v>{index}</v></c>'


def number_cell(column: str, row_number: int, number: str, style: str | None) -> str:
    style_attr = f' s="{style}"' if style else ""
    return f'<c r="{column}{row_number}"{style_attr}><v>{number}</v></c>'


def formula_cell(column: str, row_number: int, formula: str, cached: str | None, style: str | None) -> str:
    style_attr = f' s="{style}"' if style else ""
    cached_xml = f"<v>{cached}</v>" if cached is not None else ""
    return f'<c r="{column}{row_number}"{style_attr}><f>{escape(formula)}</f>{cached_xml}</c>'


def style_of(row: dict, column: str) -> str | None:
    for cell in row["cells"]:
        if cell["column"] == column:
            found = re.search(r'\bs="(\d+)"', cell["attrs"])
            return found.group(1) if found else None
    return None


def set_shared_cell(sheet_xml: str, ref: str, index: int) -> str:
    pattern = re.compile(rf'<c r="{re.escape(ref)}"([^>]*?)(?:/>|>(.*?)</c>)', re.S)

    def swap(match: re.Match[str]) -> str:
        attrs = re.sub(r'\s+t="[^"]*"', "", match.group(1))
        return f'<c r="{ref}"{attrs} t="s"><v>{index}</v></c>'

    updated, count = pattern.subn(swap, sheet_xml, count=1)
    if count != 1:
        raise StepError(_STEP, f"cella assente: {ref}")
    return updated


def clear_value(sheet_xml: str, ref: str) -> str:
    pattern = re.compile(rf'<c r="{re.escape(ref)}"([^>]*?)(?:/>|>(.*?)</c>)', re.S)

    def swap(match: re.Match[str]) -> str:
        attrs = re.sub(r'\s+t="[^"]*"', "", match.group(1))
        return f'<c r="{ref}"{attrs}/>'

    updated, count = pattern.subn(swap, sheet_xml, count=1)
    if count != 1:
        raise StepError(_STEP, f"cella assente: {ref}")
    return updated


def set_formula(sheet_xml: str, ref: str, formula: str) -> str:
    pattern = re.compile(rf'(<c r="{re.escape(ref)}"[^>]*>)(.*?)(</c>)', re.S)

    def swap(match: re.Match[str]) -> str:
        inner = re.sub(r"<f[^>]*>.*?</f>", f"<f>{escape(formula)}</f>", match.group(2), count=1, flags=re.S)
        if "<f" not in inner:
            raise StepError(_STEP, f"formula assente in {ref}")
        return match.group(1) + inner + match.group(3)

    updated, count = pattern.subn(swap, sheet_xml, count=1)
    if count != 1:
        raise StepError(_STEP, f"cella formula assente: {ref}")
    return updated


def update_dimension(sheet_xml: str, ref: str) -> str:
    if "<dimension " not in sheet_xml:
        return sheet_xml
    return re.sub(r'<dimension ref="[^"]+"', f'<dimension ref="{ref}"', sheet_xml, count=1)


def update_calc_chain(book: Workbook, old_ref: str, new_ref: str) -> None:
    raw = book.parts.get("xl/calcChain.xml")
    if raw is None or old_ref == new_ref:
        return
    book.parts["xl/calcChain.xml"] = raw.replace(f'r="{old_ref}"'.encode("utf-8"), f'r="{new_ref}"'.encode("utf-8"))


def blank_rows(sheet_xml: str, row_numbers: set[int]) -> str:
    """Svuota le celle di quelle righe. Le formule restano."""

    def swap(match: re.Match[str]) -> str:
        attrs, inner = match.group(1), match.group(2)
        number = int(re.search(r'\br="(\d+)"', attrs).group(1))
        if number not in row_numbers:
            return match.group(0)
        cells = []
        for cell_attrs, cell_inner in _CELL_RE.findall(inner):
            if cell_inner and "<f" in cell_inner:
                cells.append(f"<c{cell_attrs}>{cell_inner}</c>")
                continue
            ref = re.search(r'\br="[^"]+"', cell_attrs)
            style = re.search(r'\bs="(\d+)"', cell_attrs)
            style_attr = f' s="{style.group(1)}"' if style else ""
            cells.append(f"<c {ref.group(0)}{style_attr}/>")
        return f"<row{attrs}>{''.join(cells)}</row>"

    return _ROW_RE.sub(swap, sheet_xml)


def sha256_file(path: Path) -> str:
    digest = hashlib.sha256()
    with Path(path).open("rb") as handle:
        for chunk in iter(lambda: handle.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()

