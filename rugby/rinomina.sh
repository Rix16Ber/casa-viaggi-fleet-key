#!/bin/bash
# Generalizzazione di rinomina.sh di Padova.
# Stesso precheck, stessa mappa vecchio→nuovo, stesso annulla-rinomina.sh.
# La cartella e il CSV arrivano dal comando. --dry-run non sposta i file.
export LC_ALL=en_US.UTF-8
DIR=""
CSV=""
DRY=""
while [ $# -gt 0 ]; do
  case "$1" in
    --dry-run) DRY=1; shift ;;
    --cartella) DIR="${2:-}"; shift 2 ;;
    --csv) CSV="${2:-}"; shift 2 ;;
    *) echo "uso: rinomina.sh --cartella DIR --csv FILE [--dry-run]"; exit 2 ;;
  esac
done
if [ -z "$DIR" ] || [ -z "$CSV" ]; then
  echo "uso: rinomina.sh --cartella DIR --csv FILE [--dry-run]"
  exit 2
fi
ROOT="$(cd "$(dirname "$0")/.." && pwd)"
cd "$ROOT" || { echo "NO_DIR"; exit 2; }
exec python3 -m rugby.cursor_rugby --log /tmp/cursor-rugby-rinomina.log step4 --clip "$DIR" --csv "$CSV" ${DRY:+--dry-run}
