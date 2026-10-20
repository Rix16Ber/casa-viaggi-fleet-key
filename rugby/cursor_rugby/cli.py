"""Azioni ripetitive delle partite del Montebelluna. Il giudizio resta a Coach Rugby."""

from __future__ import annotations

import argparse
import sys
from pathlib import Path

from rugby.cursor_rugby.copybox import copy_to_box
from rugby.cursor_rugby.errors import StepError
from rugby.cursor_rugby.fasi import fill_fasi
from rugby.cursor_rugby.matchfile import create_match_file
from rugby.cursor_rugby.minuti import summarize
from rugby.cursor_rugby.names import folder_name, parse_date
from rugby.cursor_rugby.partite import write_partite_row
from rugby.cursor_rugby.rename_clips import rename_clips
from rugby.cursor_rugby.setpiece import fill_set_piece
from rugby.cursor_rugby.upload import publish
from rugby.cursor_rugby.veo import prepare_export, refuse_remote_download
from rugby.cursor_rugby.veo_public import download_public_match, list_recordings


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(prog="cursor-rugby")
    parser.add_argument("--log", type=Path, default=Path("cursor-rugby.log"))
    sub = parser.add_subparsers(dest="command", required=True)

    step1 = sub.add_parser("step1", help="prepara la cartella dal link pubblico o da un export")
    step1.add_argument("--club")
    step1.add_argument("--export", type=Path)
    step1.add_argument("--url")
    step1.add_argument("--dest", type=Path, required=True)
    step1.add_argument("--data", required=True)
    step1.add_argument("--avversario", required=True)
    step1.add_argument("--lista", type=Path)

    listed = sub.add_parser("elenco", help="elenca le partite pubbliche del club")
    listed.add_argument("--club", required=True)

    step2 = sub.add_parser("step2", help="copia la cartella sul box e confronta gli sha256")
    step2.add_argument("--origine", type=Path, required=True)
    step2.add_argument("--box", type=Path, default=Path("/workspace/rugby-video"))
    step2.add_argument("--data", required=True)
    step2.add_argument("--avversario", required=True)

    step4 = sub.add_parser("step4", help="rinomina le clip con il CSV di Coach")
    step4.add_argument("--clip", type=Path, required=True)
    step4.add_argument("--csv", type=Path, required=True)
    step4.add_argument("--dry-run", action="store_true")

    step5 = sub.add_parser("step5", help="crea il riassunto dal modello")
    step5.add_argument("--modello", type=Path, required=True)
    step5.add_argument("--dest", type=Path, required=True)
    step5.add_argument("--data", required=True)
    step5.add_argument("--avversario", required=True)

    step6 = sub.add_parser("step6", help="compila Fasi PI, senza metri né esito")
    step6.add_argument("--file", type=Path, required=True)
    step6.add_argument("--csv", type=Path, required=True)
    step6.add_argument("--durate", type=Path, required=True)
    step6.add_argument("--modello", type=Path, required=True)
    step6.add_argument("--minuto-ripresa", type=int, default=40)

    step7 = sub.add_parser("step7", help="compila Set piece dagli eventi veri")
    step7.add_argument("--file", type=Path, required=True)
    step7.add_argument("--csv", type=Path, required=True)
    step7.add_argument("--modello", type=Path, required=True)

    step9 = sub.add_parser("step9", help="carica l'xlsx solo dopo backup, rilettura e tre controlli")
    step9.add_argument("--file", type=Path, required=True)
    step9.add_argument("--modello", type=Path, required=True)
    step9.add_argument("--dest", type=Path, required=True)

    step10 = sub.add_parser("step10", help="scrive la riga del foglio Partite")
    step10.add_argument("--file", type=Path, required=True)
    step10.add_argument("--dest", type=Path, required=True)
    step10.add_argument("--data", required=True)
    step10.add_argument("--avversario", required=True)
    step10.add_argument("--intestazione", type=Path)

    step11 = sub.add_parser("step11", help="somma i minuti di fine stagione")
    step11.add_argument("--sorgente", type=Path, nargs="+", required=True)
    step11.add_argument("--dest", type=Path, required=True)
    step11.add_argument("--durata", type=int, default=80)

    check = sub.add_parser("pre-consegna", help="controllo pre-consegna, sola lettura")
    check.add_argument("file", type=Path)
    check.add_argument("conteggi", type=Path)
    check.add_argument("badge", nargs="?", type=Path)

    args = parser.parse_args(argv)
    try:
        return _run(args) or 0
    except StepError as exc:
        print(exc, file=sys.stderr)
        return 1


def _run(args: argparse.Namespace) -> int:
    command = args.command
    if command == "elenco":
        for row in list_recordings(args.club, args.log):
            print(f"{row['data']}\t{row['titolo']}\t{row['slug']}")
    elif command == "step1":
        public = args.club or (args.url if not args.export else "")
        if public:
            folder = download_public_match(public, args.dest, args.data, args.avversario, args.log)
            print(folder)
        elif args.url:
            refuse_remote_download(args.log, args.url)
        elif args.export is None or args.lista is None:
            raise StepError("1", "servono --club oppure --export e --lista.")
        else:
            prepare_export(args.export, args.dest, args.data, args.avversario, args.lista, args.log)
    elif command == "step2":
        copy_to_box(args.origine, args.box, folder_name(parse_date(args.data), args.avversario), args.log)
    elif command == "step4":
        matched, orphans = rename_clips(args.clip, args.csv, args.log, dry_run=args.dry_run)
        print(f"ABBINAMENTI: {matched}")
        print(f"ORFANI: {orphans}")
    elif command == "step5":
        create_match_file(args.modello, args.dest, args.avversario, parse_date(args.data), args.log)
    elif command == "step6":
        fill_fasi(args.file, args.csv, _durations(args.durate), args.log, args.modello, args.minuto_ripresa)
    elif command == "step7":
        fill_set_piece(args.file, args.csv, args.log, args.modello)
    elif command == "step9":
        publish(args.file, args.modello, args.dest, args.log)
    elif command == "step10":
        write_partite_row(args.file, args.dest, args.data, args.avversario, args.log, args.intestazione)
    elif command == "step11":
        summarize(args.sorgente, args.dest, args.log, args.durata)
    elif command == "pre-consegna":
        from rugby.check_partita import main as check_main

        argv = [str(args.file), str(args.conteggi)]
        if args.badge is not None:
            argv.append(str(args.badge))
        return check_main(argv, args.log)
    return 0


def _durations(path: Path) -> dict[str, int]:
    import csv

    from rugby.cursor_rugby.clock import clock_key

    found = {}
    with path.open(encoding="utf-8", newline="") as handle:
        for raw in csv.DictReader(handle):
            found[clock_key(raw["inizio"])] = int(raw["durata_s"])
    return found


if __name__ == "__main__":
    raise SystemExit(main())
