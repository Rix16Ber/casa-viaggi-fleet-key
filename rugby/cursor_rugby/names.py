"""Nomi di cartella e di file usati da Montebelluna."""

from __future__ import annotations

import re
from datetime import date


def parse_date(value: str) -> date:
    text = value.strip()
    for fmt in ("%Y-%m-%d", "%d-%m-%Y", "%d/%m/%Y"):
        try:
            return datetime_date(text, fmt)
        except ValueError:
            continue
    raise ValueError(f"data non riconosciuta: {value}")


def datetime_date(text: str, fmt: str) -> date:
    from datetime import datetime

    return datetime.strptime(text, fmt).date()


def folder_name(when: date, opponent: str) -> str:
    return f"{when:%y%m%d} -Montebelluna - {opponent.strip()}"


def workbook_name(when: date, opponent: str) -> str:
    slug = re.sub(r"\s+", "-", opponent.strip())
    return f"{slug}-Montebelluna-{when:%d-%m-%Y}-Riassunto.xlsx"


def partita_dir(root, when: date, opponent: str):
    from pathlib import Path

    return Path(root) / folder_name(when, opponent)
