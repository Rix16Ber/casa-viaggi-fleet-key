"""Passaggio 1 quando la partita è pubblica su app.veo.co.

Legge l'elenco del club e scarica il follow-cam e le clip che la pagina già mostra.
Il panorama resta dov'è. Senza file completi la cartella non viene pubblicata.
"""

from __future__ import annotations

import csv
import json
import re
import shutil
import urllib.request
from datetime import date, datetime
from pathlib import Path

from rugby.cursor_rugby.clock import format_clock
from rugby.cursor_rugby.errors import StepError
from rugby.cursor_rugby.log import write_log
from rugby.cursor_rugby.names import folder_name, parse_date
from rugby.cursor_rugby.xlsxutil import sha256_file

_STEP = "1"
_API = "https://app.veo.co/api/app"
_USER_AGENT = "Mozilla/5.0 cursor-rugby"


def download_public_match(
    source: str,
    dest_root: Path,
    when: str,
    opponent: str,
    log_path: Path,
    client: object | None = None,
) -> Path:
    day = parse_date(when)
    name = opponent.strip()
    if not name:
        _fail(log_path, "manca l'avversario")
    veo = client or HttpVeo()
    match = _resolve(veo, source, day, name, log_path)
    video = _follow_cam(veo, match["slug"], log_path)
    clips, skipped = _clips(veo, match["slug"], name, log_path)
    plan = [video, *clips]
    lengths = {}
    for item in plan:
        size = int(veo.content_length(item["url"]))
        if size <= 0:
            _fail(log_path, f"file senza dimensione: {item['file']}")
        item["bytes"] = size
        lengths[item["file"]] = size

    dest = Path(dest_root) / folder_name(day, name)
    if dest.is_dir() and _already(dest, match["slug"], lengths):
        write_log(log_path, _STEP, "ok", azione="gia-presente", cartella=str(dest), clip=len(clips))
        return dest
    if dest.exists():
        _fail(log_path, f"esiste già una cartella diversa, non la sovrascrivo: {dest.name}")

    temporary = dest.parent / (dest.name + ".in-corso")
    if temporary.exists():
        shutil.rmtree(temporary)
    temporary.mkdir(parents=True)
    try:
        stored = []
        for item in plan:
            target = temporary / item["file"]
            target.parent.mkdir(parents=True, exist_ok=True)
            write_log(log_path, _STEP, "avvio", file=item["file"], bytes=item["bytes"])
            veo.download(item["url"], target)
            if target.stat().st_size != item["bytes"] or target.stat().st_size == 0:
                _fail(log_path, f"copia incompleta: {item['file']}")
            item["sha256"] = sha256_file(target)
            stored.append(item)
        _write_table(temporary / "veo-highlight.csv", stored[1:])
        manifest = {
            "slug": match["slug"],
            "titolo": match.get("title") or "",
            "data": f"{day:%Y-%m-%d}",
            "avversario": name,
            "video": stored[0]["file"],
            "file": {item["file"]: {"bytes": item["bytes"], "sha256": item["sha256"]} for item in stored},
            "saltati": skipped,
        }
        (temporary / "veo-pubblico.json").write_text(
            json.dumps(manifest, ensure_ascii=False, indent=2) + "\n",
            encoding="utf-8",
        )
        if not _already(temporary, match["slug"], lengths):
            _fail(log_path, "il controllo della cartella non torna")
        dest.parent.mkdir(parents=True, exist_ok=True)
        temporary.rename(dest)
    except StepError:
        if temporary.exists():
            shutil.rmtree(temporary, ignore_errors=True)
        raise
    except Exception as exc:
        if temporary.exists():
            shutil.rmtree(temporary, ignore_errors=True)
        _fail(log_path, f"Veo non risponde: {exc.__class__.__name__}")
    write_log(
        log_path,
        _STEP,
        "ok",
        cartella=str(dest),
        clip=len(clips),
        video=video["file"],
        saltati=len(skipped),
    )
    return dest


def list_recordings(source: str, log_path: Path, client: object | None = None) -> list[dict]:
    club, _match = _parse_source(source)
    if not club:
        _fail(log_path, f"serve il link del club: {source}")
    veo = client or HttpVeo()
    rows = []
    for item in veo.recordings(club):
        rows.append(
            {
                "data": _start_date(item.get("start")).isoformat(),
                "titolo": item.get("title") or "",
                "slug": item.get("slug") or "",
                "squadra": item.get("team") or "",
            }
        )
    if not rows:
        _fail(log_path, "elenco Veo vuoto")
    write_log(log_path, _STEP, "ok", azione="elenco", partite=len(rows))
    return rows


class HttpVeo:
    def recordings(self, club: str) -> list[dict]:
        data = self._json(f"{_API}/clubs/{club}/recordings/")
        if not isinstance(data, list):
            raise StepError(_STEP, "elenco Veo illeggibile")
        return data

    def match(self, slug: str) -> dict:
        data = self._json(f"{_API}/matches/{slug}/")
        if not isinstance(data, dict) or not data.get("slug"):
            raise StepError(_STEP, f"partita Veo illeggibile: {slug}")
        return data

    def videos(self, slug: str) -> list[dict]:
        data = self._json(f"{_API}/matches/{slug}/videos/")
        if not isinstance(data, list):
            raise StepError(_STEP, "video Veo illeggibili")
        return data

    def highlights(self, slug: str) -> list[dict]:
        data = self._json(f"{_API}/matches/{slug}/highlights/")
        if not isinstance(data, list):
            raise StepError(_STEP, "clip Veo illeggibili")
        return data

    def content_length(self, url: str) -> int:
        request = urllib.request.Request(url, method="HEAD", headers={"User-Agent": _USER_AGENT})
        with urllib.request.urlopen(request, timeout=60) as response:
            return int(response.headers.get("Content-Length") or 0)

    def download(self, url: str, dest: Path) -> None:
        request = urllib.request.Request(url, headers={"User-Agent": _USER_AGENT})
        partial = dest.with_suffix(dest.suffix + ".parziale")
        with urllib.request.urlopen(request, timeout=300) as response, partial.open("wb") as handle:
            while True:
                chunk = response.read(1024 * 1024)
                if not chunk:
                    break
                handle.write(chunk)
        partial.replace(dest)

    def _json(self, url: str) -> object:
        request = urllib.request.Request(
            url,
            headers={"Accept": "application/json", "User-Agent": _USER_AGENT},
        )
        with urllib.request.urlopen(request, timeout=60) as response:
            return json.load(response)


def _resolve(veo: object, source: str, day: date, opponent: str, log_path: Path) -> dict:
    club, match_slug = _parse_source(source)
    if match_slug:
        match = veo.match(match_slug)
        _check_identity(match, day, opponent, log_path)
        return match
    if not club:
        _fail(log_path, f"link Veo non riconosciuto: {source}")
    chosen = []
    for item in veo.recordings(club):
        if _start_date(item.get("start")) != day:
            continue
        match = veo.match(item["slug"])
        if _opponent_ok(match, opponent):
            chosen.append(match)
    if len(chosen) != 1:
        titles = [item.get("title") or item.get("slug") or "?" for item in chosen] or ["nessuna"]
        _fail(log_path, f"partite del {day:%d/%m/%Y} per {opponent}: {', '.join(titles)}")
    return chosen[0]


def _check_identity(match: dict, day: date, opponent: str, log_path: Path) -> None:
    if _start_date(match.get("start")) != day or not _opponent_ok(match, opponent):
        _fail(
            log_path,
            f"la partita Veo è «{match.get('title')}» del {_start_date(match.get('start')):%d/%m/%Y}",
        )


def _follow_cam(veo: object, slug: str, log_path: Path) -> dict:
    found = [
        item
        for item in veo.videos(slug)
        if item.get("render_type") == "standard" and item.get("mime_type") == "video/mp4" and item.get("url")
    ]
    if len(found) != 1:
        _fail(log_path, f"follow-cam pubblico assente, trovati {len(found)}")
    return {"file": "video-intero.mp4", "url": found[0]["url"]}


def _clips(veo: object, slug: str, opponent: str, log_path: Path) -> tuple[list[dict], list[str]]:
    rows = []
    skipped = []
    for item in veo.highlights(slug):
        videos = [video for video in item.get("videos") or [] if video.get("url")]
        if not videos:
            if int(item.get("duration") or 0) > 0:
                _fail(log_path, f"clip senza video: {item.get('id')}")
            skipped.append(str(item.get("id") or ""))
            continue
        rows.append(item)
    rows.sort(key=lambda item: (int(item["start"]), str(item.get("id"))))
    if not rows:
        _fail(log_path, "nessuna clip pubblica")
    width = max(2, len(str(len(rows))))
    named = []
    seen = set()
    for index, item in enumerate(rows, start=1):
        clock = _hhmmss(int(item["start"]))
        kind = _kind(item)
        file_name = f"{index:0{width}d} {clock}_-_{kind}.mp4"
        if file_name in seen:
            _fail(log_path, f"nome clip duplicato: {file_name}")
        seen.add(file_name)
        named.append(
            {
                "file": f"Clip Veo/{file_name}",
                "url": item["videos"][0]["url"],
                "inizio": format_clock(int(item["start"])),
                "durata_s": int(item["duration"]),
                "tipo": kind,
                "squadra": _side(item.get("team_association"), opponent),
                "commento": _clean(item.get("comment")),
            }
        )
    return named, skipped


def _already(folder: Path, slug: str, lengths: dict[str, int]) -> bool:
    manifest_path = folder / "veo-pubblico.json"
    if not manifest_path.is_file():
        return False
    manifest = json.loads(manifest_path.read_text(encoding="utf-8"))
    if manifest.get("slug") != slug:
        return False
    recorded = manifest.get("file") or {}
    if set(recorded) != set(lengths):
        return False
    for name, size in lengths.items():
        path = folder / name
        info = recorded[name]
        if not path.is_file() or path.stat().st_size != size or info.get("bytes") != size:
            return False
        if sha256_file(path) != info.get("sha256"):
            return False
    return True


def _write_table(path: Path, clips: list[dict]) -> None:
    with path.open("w", encoding="utf-8", newline="") as handle:
        writer = csv.DictWriter(
            handle,
            fieldnames=["inizio", "durata_s", "tipo", "squadra", "commento", "file"],
            extrasaction="ignore",
        )
        writer.writeheader()
        writer.writerows(clips)


def _parse_source(source: str) -> tuple[str, str]:
    text = source.strip().rstrip("/")
    club = re.search(r"app\.veo\.co/clubs/([a-z0-9-]+)(?:/|$)", text)
    match = re.search(r"app\.veo\.co/matches/([a-z0-9-]+)(?:/|$)", text)
    return (club.group(1) if club else "", match.group(1) if match else "")


def _start_date(value: object) -> date:
    if not isinstance(value, str) or not value:
        raise StepError(_STEP, "data Veo assente")
    try:
        return datetime.fromisoformat(value).date()
    except ValueError as exc:
        raise StepError(_STEP, f"data Veo non valida: {value}") from exc


def _opponent_ok(match: dict, opponent: str) -> bool:
    wanted = _fold(opponent)
    fields = (
        match.get("opponent_club_name"),
        match.get("opponent_team_name"),
        match.get("title"),
    )
    return any(wanted and wanted in _fold(field) for field in fields)


def _kind(item: dict) -> str:
    tags = item.get("tags") or []
    raw = ""
    if tags and tags[0].get("slug"):
        raw = str(tags[0]["slug"])
    elif item.get("type"):
        raw = str(item["type"])
    cleaned = re.sub(r"[^a-z0-9-]+", "", raw.lower().replace("_", ""))
    if not cleaned:
        raise StepError(_STEP, f"clip senza tipo: {item.get('id')}")
    return cleaned


def _side(association: object, opponent: str) -> str:
    if association == "own":
        return "Montebelluna"
    if association == "opponent":
        return opponent
    return ""


def _clean(value: object) -> str:
    if not isinstance(value, str):
        return ""
    return value.replace("\u200b", "").strip()


def _hhmmss(total: int) -> str:
    hours, rest = divmod(total, 3600)
    minutes, seconds = divmod(rest, 60)
    return f"{hours:02d}{minutes:02d}{seconds:02d}"


def _fold(value: object) -> str:
    return re.sub(r"\s+", " ", str(value or "")).casefold().strip()


def _fail(log_path: Path, message: str) -> None:
    write_log(log_path, _STEP, "errore", messaggio=message)
    raise StepError(_STEP, message)
