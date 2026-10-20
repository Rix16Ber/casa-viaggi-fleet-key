"""Orari delle clip e minuto di gioco. Non decide metri né esito."""

from __future__ import annotations

import re

from rugby.cursor_rugby.errors import StepError

_STEP = "6"


def parse_clock(value: str) -> int:
    text = value.strip()
    if re.fullmatch(r"\d{6}", text):
        hours, minutes, seconds = int(text[0:2]), int(text[2:4]), int(text[4:6])
        if minutes > 59 or seconds > 59:
            raise StepError(_STEP, f"orario non valido: {value}")
        return hours * 3600 + minutes * 60 + seconds
    parts = text.split(":")
    try:
        numbers = [int(part) for part in parts]
    except ValueError as exc:
        raise StepError(_STEP, f"orario non valido: {value}") from exc
    if len(numbers) == 2 and numbers[1] < 60:
        return numbers[0] * 60 + numbers[1]
    if len(numbers) == 3 and numbers[1] < 60 and numbers[2] < 60:
        return numbers[0] * 3600 + numbers[1] * 60 + numbers[2]
    raise StepError(_STEP, f"orario non valido: {value}")


def format_clock(total_seconds: int) -> str:
    if total_seconds < 0:
        raise StepError(_STEP, "durata negativa")
    return f"{total_seconds // 60:02d}:{total_seconds % 60:02d}"


def clock_key(value: str) -> str:
    total = parse_clock(value)
    hours, rest = divmod(total, 3600)
    minutes, seconds = divmod(rest, 60)
    return f"{hours:02d}{minutes:02d}{seconds:02d}"


def parse_half(value: str) -> int:
    text = value.strip().lower().replace("°", "").replace("º", "")
    if text in {"1", "1o", "primo", "pt"}:
        return 1
    if text in {"2", "2o", "secondo", "st"}:
        return 2
    raise StepError(_STEP, f"tempo non valido: {value}")


def game_minute(start: str, half: int, second_half_start: int | None, ripresa_minuto: int) -> int:
    seconds = parse_clock(start)
    if half == 1:
        return max(0, seconds // 60 - 1)
    if second_half_start is None:
        raise StepError(_STEP, "manca il primo orario del secondo tempo")
    if seconds < second_half_start:
        raise StepError(_STEP, f"orario {start} è prima dell'inizio del secondo tempo")
    return ripresa_minuto + (seconds - second_half_start) // 60
