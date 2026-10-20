# Domande

Le legge dr eggbot. Non sono per Riccardo.

## 2026-10-20 — controllo pre-consegna

Il passaggio non è stato aggiunto. In questo turno non sono arrivati `check-partita.py` né il CSV dei conteggi Veo di esempio.

Per integrarlo senza cambiare la logica dei controlli servono quei file, così come li ha scritti Coach Rugby. Il conteggio di esempio non deve contenere nomi di giocatori: il repository è pubblico.

Integrazione prevista, a file arrivato:

- passaggio separato in `rugby/`
- si può rilanciare, scrive un log
- exit code diverso da 0 se c'è un FAIL
- sull'xlsx nessuna scrittura con openpyxl (la lettura va bene)
- test solo con fixture finte

Uso indicato: `python3 check-partita.py file.xlsx conteggi.csv [badge.csv]`.

Esito che Coach ha già visto su Padova, da confermare sullo script vero: 0 FAIL, 1 WARN, linee incrociate nelle fasi 20/21, 23/24 e 39/40.
