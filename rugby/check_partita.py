#!/usr/bin/env python3
"""Controllo file partita (riscaricato da OneDrive). Uso: check-partita.py file.xlsx veo-conteggi.csv [badge.csv]
badge.csv opzionale: tipo,orario,squadra (PMA/CUS) dal badge Veo."""
import json
import sys
import csv
import re
import zipfile
import openpyxl
from datetime import datetime, timezone
from pathlib import Path
from xml.etree import ElementTree as ET

_STEP = "pre-consegna"
_NS = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"
_REL = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"
_SET_ALIAS = {
    "tipo": "tipo",
    "orario veo": "ora",
    "orario": "ora",
    "orario video": "ora",
    "chi": "chi",
    "chi rimessa": "chi",
    "squadra": "chi",
    "nota veo": "nota",
    "nota": "nota",
    "esito": "esito",
    "esito conteggio": "conteggio",
}
_FASE_ALIAS = {
    "n": "N",
    "inizio": "Inizio",
    "fine": "Fine",
    "da dove": "Da dove",
    "dove": "Dove",
    "nota veo": "Nota Veo",
    "nota": "Nota Veo",
    "metri": "Metri",
    "esito": "Esito",
    "start set piece": "Start set piece",
    "end set piece": "End set piece",
    "comincia con": "Start set piece",
    "finisce con": "End set piece",
}


def main(argv=None, log_path=None) -> int:
    if argv is None:
        argv = sys.argv[1:]
    saved = sys.argv
    sys.argv = ["check-partita.py", *argv]
    try:
        return _controlla(Path(log_path) if log_path else Path("cursor-rugby.log"))
    finally:
        sys.argv = saved


def _controlla(log_path: Path) -> int:
    f,veo=sys.argv[1],sys.argv[2]; badgef=sys.argv[3] if len(sys.argv)>3 else None
    wb=openpyxl.load_workbook(f,data_only=True)
    res=[]
    def rep(stato,nome,det=''): res.append((stato,nome,det)); print(f'[{stato}] {nome}'+(f'\n      '+det.replace('\n','\n      ') if det else ''))
    T={'Monte':'PMA','Padova':'CUS'}
    # ---- Set piece (colonne dal nome dell'intestazione, anche se ripetuta)
    sp=wb['Set piece']; rows=_righe_set_piece(sp)
    V={x['tipo']:x for x in csv.DictReader(open(veo))}
    def n(x): return int(x) if x not in (None,'') else None
    print('== 1) Totali file vs Veo')
    M=[x for x in rows if x['tipo']=='Mischia']; resets={'PMA':0,'CUS':0}; reset_list=[]
    for x in M:
        k=len(re.findall(r'Reset \d+:\d+',x['nota'] or '')); resets[T.get(x['chi'],'?')]=resets.get(T.get(x['chi'],'?'),0)+k
        reset_list+= [(t,x['ora']) for t in re.findall(r'Reset (\d+:\d+)',x['nota'] or '')]
    v=V.get('Scrum')
    if v:
        righe={s:sum(1 for x in M if T.get(x['chi'])==s) for s in ('PMA','CUS')}
        tot=len(M)+sum(resets.values())
        det=f"righe file {len(M)} (PMA {righe['PMA']}, CUS {righe['CUS']}) + reset in nota {sum(resets.values())} (PMA {resets['PMA']}, CUS {resets['CUS']}) = {tot} (PMA {righe['PMA']+resets['PMA']}, CUS {righe['CUS']+resets['CUS']}); Veo {v['totale']} (PMA {v['PMA']}, CUS {v['CUS']}, reset {v['reset']})"
        ok= tot==n(v['totale']) and (not v['PMA'] or righe['PMA']+resets['PMA']==n(v['PMA'])) and (not v['CUS'] or righe['CUS']+resets['CUS']==n(v['CUS'])) and (not v['reset'] or sum(resets.values())==n(v['reset']))
        rep('PASS' if ok else 'FAIL','Mischie',det)
    L=[x for x in rows if x['tipo']=='Touche']; v=V.get('Lineout')
    if v:
        det=f"file {len(L)} (PMA {sum(T.get(x['chi'])=='PMA' for x in L)}, CUS {sum(T.get(x['chi'])=='CUS' for x in L)}); Veo {v['totale']}; differenza {n(v['totale'])-len(L)}"
        rep('PASS' if len(L)==n(v['totale']) else 'FAIL','Touche',det+("\nIl file non dice quali touche Veo mancano: servono gli orari Veo per il confronto riga per riga." if len(L)!=n(v['totale']) else ''))
    ev=wb['Eventi']; E=[(ev.cell(r,2).value,ev.cell(r,3).value,ev.cell(r,4).value) for r in range(4,ev.max_row+1) if ev.cell(r,4).value]
    for tipo,lab in (('Try','Meta'),('Conversion','Trasformazione')):
        v=V.get(tipo)
        if not v: continue
        xs=[e for e in E if e[2]==lab]
        rif=n(v.get('riusciti')) if v.get('riusciti') else n(v['totale'])
        cosa='riuscite su Veo' if v.get('riusciti') else 'Veo'
        det=f"file {len(xs)} ("+', '.join(f'{o} {T.get(s,s)}' for o,s,_ in xs)+f"); {cosa} {rif}"+(f" (tentativi Veo {v['totale']})" if v.get('riusciti') else '')+f"; differenza {rif-len(xs)}"
        rep('PASS' if len(xs)==rif else 'FAIL',f'{lab} (Veo {tipo})',det)
    print('== 2) Celle vuote')
    vu=[f"riga {x['r']} {x['tipo']} {x['ora']}: "+'/'.join(k for k in ('chi','nota') if not x[k]) for x in rows if not x['chi'] or not x['nota']]
    rep('FAIL' if vu else 'PASS','Set piece Squadra/Nota Veo','\n'.join(vu))
    da=[f"{x['tipo']} {x['ora']} ({x['chi']})" for x in rows if not x['esito']]
    rep('INFO',f'Set piece Esito da compilare a Riccardo: {len(da)}','\n'.join(da))
    fg=_foglio_fasi(wb)
    if fg is None:
        rep('FAIL','Foglio fasi','mancano Fasi Gioco e Fasi PI')
        F=[]
    else:
        F=_righe_fasi(fg)
        if F is None:
            rep('FAIL','Foglio fasi','manca la colonna N')
            F=[]
    vu=[f"fase {x['N']} ({x['Inizio']}): "+'/'.join(k for k in ('Nota Veo','Metri','Esito') if x.get(k) in (None,'')) for x in F if any(x.get(k) in (None,'') for k in ('Nota Veo','Metri','Esito'))]
    rep('FAIL' if vu else 'PASS',f'Fasi Gioco Nota/Metri/Esito ({len(F)} fasi)','\n'.join(vu))
    print('== 3) Regole')
    def sec(t):
        a=[int(v) for v in str(t).split(':')]; return a[0]*60+a[1]
    mete=[(sec(o),T.get(sq,sq)) for o,sq,ti in E if ti=='Meta' and o]
    def pos(line,x,inizio):
        """distanza dalla linea di meta CUS; linee senza squadra = nostra metà (10 = 10 m dalla metà campo)"""
        s=(line or '').strip().upper().replace(' ','')
        if s.startswith('META'):
            if 'CUS' in s: return 0
            if inizio: return 100 if 'LIBERAZIONE' in str(x.get('Start set piece')).upper() else 0
            try:
                chi=[t for o,t in mete if sec(x['Inizio'])<=o<=sec(x['Fine'])+5]
            except Exception: chi=[]
            if chi: return 100 if chi[0]=='CUS' else 0
            txt=(str(x.get('Start set piece'))+' '+str(x.get('End set piece'))).upper()
            return 100 if 'CUS' in txt else 0
        m=re.match(r'(\d+)(CUS|MONTE|PMA)?$',s)
        if not m: return None
        d=int(m.group(1)); cus=m.group(2)=='CUS'
        if d==10: return 40 if cus else 60
        if d==50: return 50
        return d if cus else 100-d
    nostre=[f"fase {x['N']}: {x['Da dove']} → {x['Dove']}" for x in F if any(re.fullmatch(r'\s*\d+\s*',str(x.get(k) or '')) for k in ('Da dove','Dove'))]
    incr=[]
    for a,b in zip(F,F[1:]):
        da,db=str(a.get('Dove') or ''),str(b.get('Da dove') or '')
        if re.sub(r'\D','',da)==re.sub(r'\D','',db) and da and db and ('CUS' in da)!=('CUS' in db) and re.sub(r'\D','',da)!='50':
            incr.append(f"fase {a['N']} finisce «{da}» / fase {b['N']} parte «{db}»")
    rep('INFO',f'Linee senza squadra lette come nostre: {len(nostre)} fasi')
    rep('WARN' if incr else 'PASS','Linee incrociate con fasi vicine (stessa linea, squadra diversa)','\n'.join(incr))
    bad=[f"fase {x['N']} {x['Inizio']}: «{x['End set piece']}» Esito={x['Esito']}" for x in F if re.match(r'calcio(?!.*contro)',str(x.get('End set piece') or ''),re.I) and 'CUS' not in str(x.get('End set piece')).upper() and str(x.get('Esito')).lower()!='positiva']
    rep('FAIL' if bad else 'PASS','Calci senza CUS = nostri e Esito Positivo','\n'.join(bad))
    sc=[]
    for x in rows:
        d=str(x['esito'] or '').lower()
        if d=='calcio contro' or d=='calcio favore' or d=='calcio a favore':
            ec=x.get('conteggio'); att='persa' if 'contro' in d else 'vinta'
            if ec!=att: sc.append(f"{x['tipo']} {x['ora']}: D={x['esito']} conteggio={ec} atteso {att}")
    for x in F:
        e=str(x.get('End set piece') or '').lower()
        if re.match(r'calcio\s*contro',e) and str(x['Esito']).lower()!='negativa': sc.append(f"fase {x['N']}: «{x['End set piece']}» Esito={x['Esito']} atteso Negativa")
    rep('FAIL' if sc else 'PASS','Calcio contro → Negativa/persa, calcio favore → Positiva/vinta','\n'.join(sc))
    if badgef:
        B={(b['tipo'],b['orario']):b['squadra'] for b in csv.DictReader(open(badgef))}
        bb=[f"{x['tipo']} {x['ora']}: file {x['chi']} / badge {B[(x['tipo'],x['ora'])]}" for x in rows if (x['tipo'],x['ora']) in B and T.get(x['chi'])!=B[(x['tipo'],x['ora'])]]
        miss=[f"{x['tipo']} {x['ora']}" for x in rows if (x['tipo'],x['ora']) not in B]
        rep('FAIL' if bb else 'PASS',f'Squadra set piece = badge Veo ({len(B)} badge)','\n'.join(bb+(['senza badge: '+', '.join(miss)] if miss else [])))
    mm=[]
    for x in F:
        p0,p1=pos(x.get('Da dove'),x,True),pos(x.get('Dove'),x,False)
        if p0 is None or p1 is None: mm.append(f"fase {x['N']}: non leggibile «{x.get('Da dove')} → {x.get('Dove')}» Metri={x['Metri']}"); continue
        att='Positiva' if p1<p0 else 'Negativa' if p1>p0 else 'Medio'
        if str(x['Metri']).lower().rstrip('a').rstrip('o')!=att.lower().rstrip('a').rstrip('o'):
            mm.append(f"fase {x['N']}: {x.get('Da dove')} → {x.get('Dove')} atteso {att}, file {x['Metri']}")
    rep('FAIL' if mm else 'PASS','Metri coerenti con Da dove → Dove','\n'.join(mm))
    print('== 4) calcChain')
    chain=_problemi_calc_chain(f)
    if chain:
        rep('FAIL','calcChain','Excel chiede di riparare il file quando calcChain indica una cella senza formula o fuori dal foglio. Su Padova è successo: un file da riparare non si consegna.\n'+'\n'.join(chain))
    else:
        rep('PASS','calcChain','nessun riferimento orfano')
    fails = sum(s == "FAIL" for s, _, _ in res)
    warns = sum(s == "WARN" for s, _, _ in res)
    print('== ESITO:', 'FAIL' if fails else 'PASS', f"({fails} FAIL, {warns} WARN)")
    _log(log_path, "errore" if fails else "ok", fail=fails, warn=warns, file=str(f))
    return 1 if fails else 0


def _norm_header(value: object) -> str:
    return re.sub(r"\s+", " ", str(value or "")).strip().casefold()


def _foglio_fasi(wb):
    if "Fasi Gioco" in wb.sheetnames:
        return wb["Fasi Gioco"]
    if "Fasi PI" in wb.sheetnames:
        return wb["Fasi PI"]
    return None


def _righe_set_piece(sp) -> list[dict]:
    rows = []
    columns: dict[str, int] = {}
    for r in range(1, sp.max_row + 1):
        values = [sp.cell(r, c).value for c in range(1, sp.max_column + 1)]
        if any(_norm_header(value) == "tipo" for value in values):
            columns = {}
            for index, value in enumerate(values, start=1):
                key = _SET_ALIAS.get(_norm_header(value))
                if key:
                    columns[key] = index
            continue
        if "tipo" not in columns:
            continue
        tipo = sp.cell(r, columns["tipo"]).value
        if tipo not in ("Touche", "Mischia"):
            continue

        def take(name: str, row: int = r) -> object:
            column = columns.get(name)
            return sp.cell(row, column).value if column else None

        rows.append(
            dict(
                r=r,
                tipo=tipo,
                ora=take("ora"),
                chi=take("chi"),
                esito=take("esito"),
                nota=take("nota"),
                conteggio=take("conteggio"),
            )
        )
    return rows


def _righe_fasi(fg) -> list[dict] | None:
    header_row = None
    headers: list = []
    for r in range(1, min(fg.max_row, 8) + 1):
        values = [fg.cell(r, c).value for c in range(1, fg.max_column + 1)]
        norms = {_norm_header(value) for value in values}
        if "da dove" in norms or ("n" in norms and "inizio" in norms):
            header_row = r
            headers = values
            break
    if header_row is None:
        return None
    present = {_norm_header(value) for value in headers}
    index: dict[str, int] = {}
    for position, value in enumerate(headers):
        key = _norm_header(value)
        if key == "nota" and "nota veo" in present:
            continue
        canon = _FASE_ALIAS.get(key)
        if canon:
            index[canon] = position
    if "N" not in index:
        return None
    rows = []
    width = fg.max_column
    for r in range(header_row + 1, fg.max_row + 1):
        values = [fg.cell(r, c).value for c in range(1, width + 1)]
        if index["N"] >= len(values) or values[index["N"]] is None:
            continue
        rows.append({canon: values[position] if position < len(values) else None for canon, position in index.items()})
    return rows


def _problemi_calc_chain(path: str) -> list[str]:
    problems = []
    with zipfile.ZipFile(path) as book:
        if "xl/calcChain.xml" not in book.namelist():
            return problems
        workbook = ET.fromstring(book.read("xl/workbook.xml"))
        rels = ET.fromstring(book.read("xl/_rels/workbook.xml.rels"))
        targets = {rel.attrib["Id"]: rel.attrib["Target"] for rel in rels}
        by_id = {}
        for sheet in workbook.findall(f"{_NS}sheets/{_NS}sheet"):
            target = targets.get(sheet.attrib.get(_REL + "id"), "")
            if target.startswith("/"):
                target = target[1:]
            elif not target.startswith("xl/"):
                target = "xl/" + target
            by_id[sheet.attrib.get("sheetId")] = (sheet.attrib.get("name"), target)
        cache: dict[str, tuple[str, str | None, set[str]] | None] = {}

        def info(sheet_id: str):
            if sheet_id not in cache:
                found = by_id.get(sheet_id)
                if found is None or found[1] not in book.namelist():
                    cache[sheet_id] = None
                else:
                    root = ET.fromstring(book.read(found[1]))
                    dimension = root.find(f"{_NS}dimension")
                    ref = dimension.attrib.get("ref") if dimension is not None else None
                    formulas = {
                        cell.attrib["r"]
                        for cell in root.iter(f"{_NS}c")
                        if cell.attrib.get("r") and cell.find(f"{_NS}f") is not None
                    }
                    cache[sheet_id] = (found[0], ref, formulas)
            return cache[sheet_id]

        last = None
        for cell in ET.fromstring(book.read("xl/calcChain.xml")).findall(f"{_NS}c"):
            ref = cell.attrib.get("r", "")
            sheet_id = cell.attrib.get("i") or last
            if cell.attrib.get("i"):
                last = cell.attrib["i"]
            found = info(sheet_id) if sheet_id else None
            if found is None:
                problems.append(f"{ref}: foglio {sheet_id} assente")
                continue
            name, dimension, formulas = found
            if dimension and not _dentro_il_range(ref, dimension):
                problems.append(f"{name}!{ref}: fuori dal range {dimension}")
            elif ref not in formulas:
                problems.append(f"{name}!{ref}: senza formula")
    return problems


def _dentro_il_range(ref: str, dimension: str) -> bool:
    def parse(cell: str) -> tuple[int, int]:
        match = re.match(r"([A-Z]+)(\d+)", cell.upper())
        if match is None:
            return 0, 0
        column = 0
        for char in match.group(1):
            column = column * 26 + ord(char) - 64
        return column, int(match.group(2))

    start, _, end = dimension.partition(":")
    left, top = parse(start)
    right, bottom = parse(end or start)
    column, row = parse(ref)
    return min(left, right) <= column <= max(left, right) and min(top, bottom) <= row <= max(top, bottom)


def _log(path: Path, status: str, **fields: object) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    record = {
        "ts": datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ"),
        "step": _STEP,
        "status": status,
    }
    record.update(fields)
    with path.open("a", encoding="utf-8") as handle:
        handle.write(json.dumps(record, ensure_ascii=False) + "\n")


if __name__ == "__main__":
    sys.exit(main())
