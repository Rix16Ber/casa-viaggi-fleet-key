"""Foglio minimo con gli stessi nomi, le formule e un grafico a immagine."""

from __future__ import annotations

import base64
import zipfile
from pathlib import Path
from xml.sax.saxutils import escape

PNG = base64.b64decode(
    "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg=="
)


def build_template(path: Path) -> Path:
    strings: list[str] = []

    def intern(text: str) -> int:
        if text not in strings:
            strings.append(text)
        return strings.index(text)

    def cell(ref: str, text: str | None = None, number: str | None = None, formula: str | None = None) -> str:
        if formula is not None:
            cached = f"<v>{number}</v>" if number is not None else ""
            return f'<c r="{ref}"><f>{escape(formula)}</f>{cached}</c>'
        if text is not None:
            return f'<c r="{ref}" t="s"><v>{intern(text)}</v></c>'
        if number is not None:
            return f'<c r="{ref}"><v>{number}</v></c>'
        return f'<c r="{ref}"/>'

    def row(number: int, cells: list[str]) -> str:
        return f'<row r="{number}">{"".join(cells)}</row>'

    summary = [
        row(1, [cell("B1", "Montebelluna – Paese · 03/10/2026")]),
        row(5, [cell("B5", "Montebelluna"), cell("D5", number="68"), cell("F5", number="0"), cell("G5", "Paese")]),
        row(6, [cell("F6", number="35.6", formula="'Fasi PI'!F50/60")]),
        row(9, [cell("B9", "Montebelluna"), cell("C9", number="0"), cell("D9", number="0"), cell("E9", number="0"), cell("F9", number="0"), cell("G9", number="0")]),
        row(10, [cell("B10", "Paese"), cell("C10", number="10"), cell("D10", number="9"), cell("E10", number="1"), cell("F10", number="0"), cell("G10", number="0")]),
        row(14, [cell("B14", "Montebelluna"), cell("C14", number="14"), cell("D14", number="6"), cell("E14", number="3"), cell("F14", number="5")]),
        row(15, [cell("B15", "Paese"), cell("C15", number="4"), cell("D15", number="2"), cell("E15", number="0"), cell("F15", number="2")]),
        row(19, [cell("B19", "Montebelluna"), cell("C19", number="5"), cell("D19", number="4")]),
        row(20, [cell("B20", "Paese"), cell("C20", number="8"), cell("D20", number="5")]),
    ]
    set_piece = [
        row(1, [cell("A1", "SET PIECE — titolo")]),
        row(2, [cell("A2", "nota di testa")]),
        row(4, [cell("A4", "Tipo"), cell("B4", "Orario"), cell("C4", "Chi"), cell("D4", "Nota"), cell("E4", "Esito")]),
        row(5, [cell("A5", "Touche"), cell("B5", "02:52"), cell("C5", "Monte"), cell("D5", "Persa"), cell("E5", "persa")]),
        row(6, [cell("A6", "MISCHIE — clip Scrum")]),
        row(7, [cell("A7", "Tipo"), cell("B7", "Orario"), cell("C7", "Chi"), cell("D7", "Nota"), cell("E7", "Esito")]),
        row(8, [cell("A8", "Mischia"), cell("B8", "10:49"), cell("C8", "Paese"), cell("D8", "Vinta"), cell("E8", "vinta")]),
    ]
    eventi = [
        row(1, [cell("A1", "EVENTI")]),
        row(3, [cell("A3", "#"), cell("B3", "Orario")]),
        row(4, [cell("A4", "1"), cell("B4", "03:38"), cell("C4", "Paese"), cell("D4", "Meta")]),
    ]
    fasi = [
        row(1, [cell(f"{column}1", label) for column, label in zip("ABCDEFGHIJKLM", ["N", "Tempo", "Inizio", "Fine", "Minuto", "Durata (s)", "Comincia con", "Da dove", "Palla a", "Finisce con", "Dove", "Esito", "Metri"])]),
        row(
            2,
            [
                cell("A2", number="1"),
                cell("B2", "1°"),
                cell("C2", "01:08"),
                cell("D2", "02:20"),
                cell("E2", "0'"),
                cell("F2", number="72"),
                cell("G2", "Calcio"),
                cell("H2", "metà"),
                cell("I2", "Paese"),
                cell("J2", "Touche"),
                cell("K2", "metà"),
                cell("L2", "Positiva"),
                cell("M2", "Medio"),
            ],
        ),
        row(50, [cell("F50", number="72", formula="SUM(F2:F49)")]),
    ]
    sheets = {
        "sheet1.xml": _worksheet("".join(summary), "B1:G20"),
        "sheet2.xml": _worksheet("".join(set_piece), "A1:E8", merges=["A1:E1", "A2:E2", "A6:E6"]),
        "sheet3.xml": _worksheet("".join(eventi), "A1:D4"),
        "sheet4.xml": _worksheet("".join(fasi), "A1:M50"),
    }
    path.parent.mkdir(parents=True, exist_ok=True)
    with zipfile.ZipFile(path, "w") as archive:
        archive.writestr("[Content_Types].xml", _content_types())
        archive.writestr("_rels/.rels", _rels([("rId1", "xl/workbook.xml", "officeDocument")]))
        archive.writestr("xl/workbook.xml", _workbook())
        archive.writestr(
            "xl/_rels/workbook.xml.rels",
            _rels(
                [
                    ("rId1", "worksheets/sheet1.xml", "worksheet"),
                    ("rId2", "worksheets/sheet2.xml", "worksheet"),
                    ("rId3", "worksheets/sheet3.xml", "worksheet"),
                    ("rId4", "worksheets/sheet4.xml", "worksheet"),
                    ("rId5", "sharedStrings.xml", "sharedStrings"),
                ]
            ),
        )
        for name, xml in sheets.items():
            archive.writestr(f"xl/worksheets/{name}", xml)
        archive.writestr("xl/sharedStrings.xml", _sst(strings))
        archive.writestr("xl/drawings/drawing1.xml", "<xdr:wsDr xmlns:xdr=\"http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing\"/>")
        archive.writestr(
            "xl/drawings/_rels/drawing1.xml.rels",
            _rels([("rId1", "../media/image1.png", "image")]),
        )
        archive.writestr(
            "xl/worksheets/_rels/sheet1.xml.rels",
            _rels([("rId1", "../drawings/drawing1.xml", "drawing")]),
        )
        archive.writestr("xl/media/image1.png", PNG)
        archive.writestr(
            "xl/calcChain.xml",
            '<calcChain xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
            '<c r="F50" i="4" l="1"/><c r="F6" i="1" s="1"/></calcChain>',
        )
    return path


def _worksheet(rows: str, dimension: str, merges: list[str] | None = None) -> str:
    merge_xml = ""
    if merges:
        merge_xml = f'<mergeCells count="{len(merges)}">' + "".join(f'<mergeCell ref="{ref}"/>' for ref in merges) + "</mergeCells>"
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
        'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
        f'<dimension ref="{dimension}"/>'
        f"<sheetData>{rows}</sheetData>{merge_xml}</worksheet>"
    )


def _workbook() -> str:
    sheets = "".join(
        f'<sheet name="{name}" sheetId="{index}" r:id="rId{index}"/>'
        for index, name in enumerate(("Riassunto", "Set piece", "Eventi", "Fasi PI"), start=1)
    )
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
        'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
        f"<sheets>{sheets}</sheets></workbook>"
    )


def _sst(strings: list[str]) -> str:
    body = "".join(f"<si><t>{escape(text)}</t></si>" for text in strings)
    count = len(strings)
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        f'<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="{count}" uniqueCount="{count}">'
        f"{body}</sst>"
    )


def _content_types() -> str:
    overrides = [
        "/xl/workbook.xml",
        "/xl/worksheets/sheet1.xml",
        "/xl/worksheets/sheet2.xml",
        "/xl/worksheets/sheet3.xml",
        "/xl/worksheets/sheet4.xml",
        "/xl/sharedStrings.xml",
        "/xl/drawings/drawing1.xml",
        "/xl/calcChain.xml",
    ]
    body = '<Default Extension="png" ContentType="image/png"/>'
    body += "".join(
        f'<Override PartName="{name}" ContentType="application/xml"/>' for name in overrides
    )
    return f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">{body}</Types>'


def _rels(items: list[tuple[str, str, str]]) -> str:
    types = {
        "officeDocument": "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument",
        "worksheet": "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet",
        "sharedStrings": "http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings",
        "drawing": "http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing",
        "image": "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image",
    }
    body = "".join(
        f'<Relationship Id="{rel_id}" Type="{types[kind]}" Target="{target}"/>'
        for rel_id, target, kind in items
    )
    return f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">{body}</Relationships>'
