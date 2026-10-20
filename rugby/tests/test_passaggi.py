import tempfile
import unittest
from pathlib import Path

from rugby.cursor_rugby.clock import format_clock, game_minute, parse_clock
from rugby.cursor_rugby.copybox import copy_to_box
from rugby.cursor_rugby.errors import StepError
from rugby.cursor_rugby.fasi import fill_fasi
from rugby.cursor_rugby.matchfile import create_match_file
from rugby.cursor_rugby.minuti import parse_formation, summarize
from rugby.cursor_rugby.names import parse_date
from rugby.cursor_rugby.partite import COLONNE, write_partite_row
from rugby.cursor_rugby.rename_clips import rename_clips
from rugby.cursor_rugby.setpiece import fill_set_piece
from rugby.cursor_rugby.upload import publish
from rugby.cursor_rugby.veo import prepare_export, refuse_remote_download
from rugby.cursor_rugby.veo_public import download_public_match, list_recordings
from rugby.cursor_rugby.xlsxutil import Workbook, chart_hashes, parse_rows, row_map
from rugby.tests.factory import build_template

FORMATION = (
    "Montebelluna rugby 1977 : Giocatore A, Giocatore B, Giocatore C, Giocatore D, "
    "Giocatore E (78' Giocatore F), Giocatore G, Giocatore H (20' Giocatore I), "
    "Giocatore J (66' Giocatore K), Giocatore L, Giocatore M (78' Giocatore N), "
    "Giocatore O, Giocatore P, Giocatore Q (57' Giocatore R), "
    "Giocatore S (57' Giocatore T), Giocatore U (49' Giocatore V). All. Allenatore A, Allenatore B."
)


class ClockTest(unittest.TestCase):
    def test_fine_e_minuto(self):
        self.assertEqual(format_clock(parse_clock("01:08") + 72), "02:20")
        self.assertEqual(game_minute("01:08", 1, None, 40), 0)
        self.assertEqual(game_minute("05:08", 1, None, 40), 4)
        origin = parse_clock("43:33")
        self.assertEqual(game_minute("43:33", 2, origin, 40), 40)
        self.assertEqual(game_minute("45:19", 2, origin, 40), 41)


class VeoCopyRenameTest(unittest.TestCase):
    def test_export_copia_e_rinomina(self):
        with tempfile.TemporaryDirectory() as raw:
            root = Path(raw)
            export = root / "export"
            export.mkdir()
            (export / "partita-intera.mp4").write_bytes(b"video")
            (export / "01 000108_-_Fase_Gioco.mp4").write_bytes(b"clip-a")
            (export / "02 000251_-_Fase_Gioco.mp4").write_bytes(b"clip-b")
            lista = root / "lista.txt"
            lista.write_text(
                "01 000108_-_Fase_Gioco.mp4\n02 000251_-_Fase_Gioco.mp4\n",
                encoding="utf-8",
            )
            log = root / "log.txt"
            match = prepare_export(export, root / "od", "2026-10-18", "CUS Padova", lista, log)
            again = prepare_export(export, root / "od", "2026-10-18", "CUS Padova", lista, log)
            self.assertEqual(match, again)
            box = copy_to_box(match.parent, root / "box", "261018 -Montebelluna - CUS Padova", log)
            copy_to_box(match.parent, root / "box", "261018 -Montebelluna - CUS Padova", log)
            table = root / "rinomina.csv"
            table.write_text(
                "nome attuale,nome nuovo,orario video Veo,nota Veo\n"
                "01 000108_-_Fase_Gioco.mp4,01 - Calcio.mp4,01:08,Calcio\n"
                "02 000251_-_Fase_Gioco.mp4,02 - Mischia.mp4,02:51,Mischia\n",
                encoding="utf-8",
            )
            rename_clips(box / "Fasi gioco", table, log)
            self.assertTrue((box / "Fasi gioco" / "01 - Calcio.mp4").is_file())
            rename_clips(box / "Fasi gioco", table, log)
            mapping = (box / "Fasi gioco" / "rinomina-vecchio-nuovo.tsv").read_text(encoding="utf-8")
            self.assertIn("01 - Calcio.mp4", mapping)
            undo = (box / "Fasi gioco" / "annulla-rinomina.sh").read_text(encoding="utf-8")
            self.assertIn("rinomina-vecchio-nuovo.tsv", undo)

    def test_file_vuoto_e_url(self):
        with tempfile.TemporaryDirectory() as raw:
            root = Path(raw)
            export = root / "export"
            export.mkdir()
            (export / "partita.mp4").write_bytes(b"video")
            (export / "01 000108_-_Fase_Gioco.mp4").write_bytes(b"")
            lista = root / "lista.txt"
            lista.write_text("01 000108_-_Fase_Gioco.mp4\n", encoding="utf-8")
            with self.assertRaises(StepError):
                prepare_export(export, root / "od", "18-10-2026", "Paese", lista, root / "log.txt")
            self.assertFalse((root / "od").exists())
            with self.assertRaises(StepError) as raised:
                refuse_remote_download(root / "log.txt", "https://app.veo.co/matches/esempio")
            self.assertIn("manuale", str(raised.exception))

    def test_rinomina_da_orario_anche_senza_fase_gioco(self):
        with tempfile.TemporaryDirectory() as raw:
            root = Path(raw)
            clips = root / "clip"
            clips.mkdir()
            (clips / "01 000449_-_Touche.mp4").write_bytes(b"touche")
            (clips / "video-intero.mp4").write_bytes(b"video")
            table = root / "tabella-rinomina.csv"
            table.write_text(
                "nome attuale,nome nuovo,orario video Veo,nota Veo\n"
                "01 000449_-_Touche.mp4,01 - Touche.mp4,04:49,Touche\n",
                encoding="utf-8",
            )
            log = root / "log.txt"
            matched, orphans = rename_clips(clips, table, log, dry_run=True)
            self.assertEqual((matched, orphans), (1, 0))
            self.assertTrue((clips / "01 000449_-_Touche.mp4").is_file())
            rename_clips(clips, table, log)
            self.assertTrue((clips / "01 - Touche.mp4").is_file())
            self.assertTrue((clips / "video-intero.mp4").is_file())
            rename_clips(clips, table, log)
            with self.assertRaises(StepError):
                rename_clips(clips, root / "manca.csv", log)
            self.assertIn("CSV assente", log.read_text(encoding="utf-8"))


class WorkbookTest(unittest.TestCase):
    def test_file_fasi_setpiece_e_controlli(self):
        with tempfile.TemporaryDirectory() as raw:
            root = Path(raw)
            template = build_template(root / "modello.xlsx")
            before = chart_hashes(template)
            created = create_match_file(template, root / "out", "CUS Padova", parse_date("2026-10-18"), root / "log.txt")
            create_match_file(template, root / "out", "CUS Padova", parse_date("2026-10-18"), root / "log.txt")
            self.assertEqual(chart_hashes(created), before)
            book = Workbook(created)
            strings = book.strings()
            summary = {row["row"]: row_map(row, strings) for row in parse_rows(book.sheet_xml("Riassunto"))}
            self.assertEqual(summary[1]["B"], "Montebelluna – CUS Padova · 18/10/2026")
            self.assertEqual(summary[5]["G"], "CUS Padova")
            self.assertNotIn("D", summary[5])

            fasi = root / "fasi.csv"
            fasi.write_text(
                "inizio,tempo,comincia,da,palla,finisce,dove,nota\n"
                "01:08,1°,Calcio,metà,Paese,Touche,metà,prima\n"
                "43:33,2°,Calcio,metà,Monte,Touche,loro 22,\n",
                encoding="utf-8",
            )
            durate = root / "durate.csv"
            durate.write_text("inizio,durata_s\n01:08,72\n43:33,71\n", encoding="utf-8")
            # sul modello, così Esito e Metri della fase 1 restano
            fill_fasi(template, fasi, {"000108": 72, "004333": 71}, root / "log.txt", template)
            filled = Workbook(template)
            phase_rows = parse_rows(filled.sheet_xml("Fasi PI"))
            texts = [row_map(row, filled.strings()) for row in phase_rows]
            self.assertEqual(texts[1]["D"], "02:20")
            self.assertEqual(texts[1]["E"], "0'")
            self.assertEqual(texts[1]["L"], "Positiva")
            self.assertEqual(texts[1]["M"], "Medio")
            self.assertNotIn("L", texts[2])
            self.assertEqual(texts[2]["E"], "40'")
            self.assertEqual(texts[1]["N"], "prima")
            formulas = [cell["formula"] for row in phase_rows for cell in row["cells"] if cell["formula"]]
            self.assertIn("SUM(F2:F3)", formulas)

            piece = root / "set.csv"
            piece.write_text(
                "tipo,orario,squadra,nota,esito\n"
                "commento,00:01,,,\n"
                "Touche,02:52,Monte,Persa,persa\n"
                "Touche,07:19,,Rubata,persa\n",
                encoding="utf-8",
            )
            with self.assertRaises(StepError):
                fill_set_piece(created, piece, root / "log.txt", template)
            piece.write_text(
                "tipo,orario,squadra,nota,esito\n"
                "commento,00:01,,,\n"
                "Touche,02:52,Monte,Persa,persa\n"
                "Mischia,10:49,CUS Padova,Vinta,vinta\n",
                encoding="utf-8",
            )
            fill_set_piece(created, piece, root / "log.txt", template)
            piece_rows = [
                row_map(row, Workbook(created).strings())
                for row in parse_rows(Workbook(created).sheet_xml("Set piece"))
            ]
            labels = [row.get("A") for row in piece_rows]
            self.assertEqual(labels.count("Touche"), 1)
            self.assertEqual(labels.count("Mischia"), 1)
            self.assertIn("MISCHIE — clip Scrum", labels)
            self.assertTrue(all(row.get("C") for row in piece_rows if row.get("A") in {"Touche", "Mischia"}))
            self.assertEqual(chart_hashes(created), before)

            published = publish(created, template, root / "pubblico.xlsx", root / "log.txt")
            publish(created, template, published, root / "log.txt")
            broken = root / "rotto.xlsx"
            build_template(broken)
            # togli la formula e prova a pubblicarlo sopra una copia buona
            good = root / "buono.xlsx"
            good.write_bytes(template.read_bytes())
            stripped = Workbook(broken)
            xml = stripped.sheet_xml("Riassunto").replace("<f>'Fasi PI'!F50/60</f>", "")
            stripped.set_sheet_xml("Riassunto", xml)
            stripped.save(broken)
            with self.assertRaises(StepError):
                publish(broken, template, good, root / "log.txt")
            self.assertEqual(good.read_bytes(), template.read_bytes())

            row = write_partite_row(created, root / "partite.tsv", "18/10/2026", "CUS Padova", root / "log.txt")
            header, data = row.read_text(encoding="utf-8").splitlines()
            self.assertEqual(len(header.split("\t")), len(COLONNE))
            self.assertEqual(len(data.split("\t")), len(header.split("\t")))
            bad = root / "intestazione.tsv"
            bad.write_text("solo\tuna\n", encoding="utf-8")
            with self.assertRaises(StepError):
                write_partite_row(created, root / "no.tsv", "18/10/2026", "CUS Padova", root / "log.txt", bad)


class RealTemplateTest(unittest.TestCase):
    def test_modello_paese(self):
        template = Path("/tmp/monte/paese.xlsx")
        if not template.is_file():
            self.skipTest("modello reale non presente")
        with tempfile.TemporaryDirectory() as raw:
            root = Path(raw)
            before = chart_hashes(template)
            created = create_match_file(template, root, "CUS Padova", parse_date("2026-10-18"), root / "log.txt")
            self.assertEqual(chart_hashes(created), before)
            book = Workbook(created)
            self.assertEqual(book.sheet_names(), ["Riassunto", "Set piece", "Eventi", "Fasi PI"])
            summary = book.sheet_xml("Riassunto")
            self.assertIn("'Fasi PI'!F", summary)
            self.assertIn("drawing1.xml", book.parts["xl/worksheets/_rels/sheet1.xml.rels"].decode("utf-8"))


class MinutiTest(unittest.TestCase):
    def test_somma_partite(self):
        played = parse_formation(FORMATION)
        self.assertEqual(sum(played.values()), 15 * 80)
        self.assertEqual(played["Giocatore E"], 78)
        self.assertEqual(played["Giocatore F"], 2)
        with tempfile.TemporaryDirectory() as raw:
            root = Path(raw)
            first = root / "formazione.txt"
            first.write_text(FORMATION + "\n", encoding="utf-8")
            second = root / "minuti.csv"
            roster = [f"Giocatore {index}" for index in range(1, 16)]
            roster[0] = "Giocatore A"
            lines = ["giocatore,minuti"] + [f"{name},80" for name in roster]
            second.write_text("\n".join(lines) + "\n", encoding="utf-8")
            dest = summarize([first, second], root / "stagione.tsv", root / "log.txt")
            summarize([first, second], dest, root / "log.txt")
            rows = dict(line.split("\t") for line in dest.read_text(encoding="utf-8").splitlines()[1:])
            self.assertEqual(rows["Giocatore A"], "160")
            self.assertEqual(sum(int(value) for value in rows.values()), 15 * 80 * 2)


class _VeoFinto:
    def __init__(self, body_size: int | None = None):
        self.body_size = body_size
        self.downloads = 0

    def recordings(self, club: str) -> list[dict]:
        self.club = club
        return [
            {
                "slug": "20261018-montebelluna-cus-padova",
                "title": "Montebelluna - CUS Padova",
                "start": "2026-10-18T12:00:00+02:00",
                "team": "1° Squadra",
            },
            {
                "slug": "20261003-partita",
                "title": "Montebelluna Rugby",
                "start": "2026-10-03T12:00:00+02:00",
                "team": "1° Squadra",
            },
        ]

    def match(self, slug: str) -> dict:
        if slug.startswith("20261003"):
            return {
                "slug": slug,
                "title": "Montebelluna Rugby",
                "opponent_club_name": "Paese",
                "start": "2026-10-03T12:00:00+02:00",
            }
        return {
            "slug": slug,
            "title": "Montebelluna - CUS Padova",
            "opponent_club_name": "CUS Padova",
            "start": "2026-10-18T12:00:00+02:00",
        }

    def videos(self, slug: str) -> list[dict]:
        return [
            {"render_type": "panorama", "mime_type": "video/mp4", "url": "https://cdn.test/panorama.mp4"},
            {"render_type": "standard", "mime_type": "video/mp4", "url": "https://cdn.test/full.mp4"},
        ]

    def highlights(self, slug: str) -> list[dict]:
        return [
            {
                "id": "dopo",
                "start": 618,
                "duration": 20,
                "type": "lineout",
                "team_association": "opponent",
                "comment": "\u200bVinta",
                "tags": [{"slug": "lineout"}],
                "videos": [{"url": "https://cdn.test/clip-b.mp4"}],
            },
            {
                "id": "prima",
                "start": 11,
                "duration": 15,
                "type": "try_",
                "team_association": "own",
                "comment": "\u200b",
                "tags": [{"slug": "try"}],
                "videos": [{"url": "https://cdn.test/clip-a.mp4"}],
            },
            {
                "id": "vuoto",
                "start": 0,
                "duration": 0,
                "type": "half_start",
                "team_association": None,
                "comment": None,
                "tags": [{"slug": "half-start"}],
                "videos": [],
            },
        ]

    def content_length(self, url: str) -> int:
        if self.body_size is not None and url.endswith("full.mp4"):
            return self.body_size
        return len(self._body(url))

    def download(self, url: str, dest: Path) -> None:
        self.downloads += 1
        dest.write_bytes(self._body(url))

    def _body(self, url: str) -> bytes:
        return {
            "https://cdn.test/full.mp4": b"VIDEO",
            "https://cdn.test/clip-a.mp4": b"A",
            "https://cdn.test/clip-b.mp4": b"BB",
        }[url]


class VeoPublicTest(unittest.TestCase):
    def test_scarica_il_follow_cam_e_le_clip(self):
        with tempfile.TemporaryDirectory() as raw:
            root = Path(raw)
            client = _VeoFinto()
            source = "https://app.veo.co/clubs/montebelluna-rugby-club-1977/recordings/"
            folder = download_public_match(source, root, "2026-10-18", "CUS Padova", root / "log.txt", client)
            again = download_public_match(source, root, "18/10/2026", "CUS Padova", root / "log.txt", client)
            self.assertEqual(folder, again)
            self.assertEqual(client.downloads, 3)
            self.assertEqual((folder / "video-intero.mp4").read_bytes(), b"VIDEO")
            names = sorted(path.name for path in (folder / "Clip Veo").iterdir())
            self.assertEqual(names, ["01 000011_-_try.mp4", "02 001018_-_lineout.mp4"])
            table = (folder / "veo-highlight.csv").read_text(encoding="utf-8")
            self.assertIn("Montebelluna", table)
            self.assertIn("CUS Padova", table)
            self.assertIn("Vinta", table)
            self.assertNotIn("\u200b", table)
            self.assertFalse((folder / "panorama.mp4").exists())
            rows = list_recordings(source, root / "log.txt", client)
            self.assertEqual(rows[0]["slug"], "20261018-montebelluna-cus-padova")

    def test_si_ferma_senza_cartella(self):
        with tempfile.TemporaryDirectory() as raw:
            root = Path(raw)
            client = _VeoFinto(body_size=99)
            with self.assertRaises(StepError):
                download_public_match(
                    "https://app.veo.co/clubs/montebelluna-rugby-club-1977/recordings/",
                    root,
                    "2026-10-18",
                    "CUS Padova",
                    root / "log.txt",
                    client,
                )
            self.assertFalse(list(root.glob("*Montebelluna*")))
            with self.assertRaises(StepError):
                download_public_match(
                    "https://app.veo.co/matches/20261018-montebelluna-cus-padova/",
                    root,
                    "2026-10-18",
                    "Paese",
                    root / "log.txt",
                    client,
                )
            self.assertFalse(list(root.glob("*Montebelluna*")))


if __name__ == "__main__":
    unittest.main()
