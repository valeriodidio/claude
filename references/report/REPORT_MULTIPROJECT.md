# Gestione multi-progetto — Backend condiviso

Regole e checklist per aggiungere un nuovo report senza rompere quelli esistenti.

---

## Il problema: main.py è condiviso

Tutti i report girano sullo **stesso processo FastAPI** (`YepponReportsAPI`).
Il file `main.py` è **uno solo** sul server e registra i router di tutti i report.

Ogni progetto di report ha il suo repository locale con il suo `main.py`, ma
sul server ne esiste uno solo. Se carichi in staging il `main.py` del progetto
nuovo, **sovrascrivi** quello del server e perdi i router dei report precedenti.

```
Server (un solo processo FastAPI)
├── app/main.py             ← UNICO, contiene TUTTI i router
├── app/routers/
│   ├── turnover.py         ← report 1
│   ├── pl_prodotti.py      ← report 2
│   └── nuovo_report.py     ← report 3 (nuovo)
└── app/services/
    ├── turnover_query.py
    ├── pl_prodotti_query.py
    └── nuovo_report_query.py
```

---

## Regola principale

> **Non caricare mai `main.py` in staging senza prima leggere quello attuale sul server.**

---

## Workflow corretto per aggiungere un nuovo report

### 1. Leggi il main.py attuale dal server

Via RDP o PowerShell:
```powershell
Get-Content D:\admin_yeppon_python\apps\yeppon_reports\python_backend\app\main.py
```

Oppure scaricalo via FileZilla prima di procedere.

### 2. Aggiungi il nuovo router al main.py esistente

Parti dal `main.py` del server (non da quello del progetto locale) e aggiungi:

```python
# riga import — aggiungi in coda agli import esistenti
from app.routers import turnover, pl_prodotti, nuovo_report

# riga include_router — aggiungi in coda ai router esistenti
app.include_router(nuovo_report.router, prefix="/api/reports/nuovo_report", tags=["nuovo_report"])
```

### 3. Carica in staging solo i file nuovi/modificati

```
staging\
├── asp_nuovo\
│   └── nuovo_report\
│       ├── nuovo_report.asp
│       └── download-nuovo-report.xlsx.asp
└── python_backend\
    └── app\
        ├── main.py                          ← il main.py fuso (tutti i router)
        ├── routers\
        │   └── nuovo_report.py              ← solo il nuovo
        └── services\
            └── nuovo_report_query.py        ← solo il nuovo
```

I file dei report **già esistenti** (`turnover.py`, `pl_prodotti.py`, ecc.)
**non vanno ricaricati** se non sono stati modificati.

### 4. Deploy e verifica

```powershell
cd D:\admin_yeppon_python\apps\deploy
.\deploy.ps1
```

Dopo il restart verifica che tutti i report rispondano:
```powershell
Invoke-WebRequest -Uri "http://127.0.0.1:8000/api/reports/health"       -UseBasicParsing
Invoke-WebRequest -Uri "http://127.0.0.1:8000/api/reports/turnover/filters"    -UseBasicParsing
Invoke-WebRequest -Uri "http://127.0.0.1:8000/api/reports/pl_prodotti/filters" -UseBasicParsing
Invoke-WebRequest -Uri "http://127.0.0.1:8000/api/reports/nuovo_report/filters" -UseBasicParsing
```

---

## Se hai già sovrascritto main.py per errore

Il deploy.ps1 fa un backup automatico prima di ogni deploy. Trovi il main.py
precedente in:

```
D:\admin_yeppon_python\deploy\backups\<timestamp>\python\app\main.py
```

Lista backup disponibili:
```powershell
Get-ChildItem D:\admin_yeppon_python\deploy\backups\ | Sort-Object LastWriteTime -Descending
```

Procedura di recupero:
1. Leggi il backup per vedere quali router c'erano
2. Fondi il backup con il main.py nuovo (aggiungi i router mancanti)
3. Carica il main.py fuso in staging e rideploya

---

## Stato router attivi sul server (aggiorna ad ogni nuovo report)

| Report | Router | Prefix |
|--------|--------|--------|
| Fatturato (Turnover) | `turnover` | `/api/reports/turnover` |
| P&L Prodotti | `pl_prodotti` | `/api/reports/pl_prodotti` |

> Aggiorna questa tabella ogni volta che aggiungi un report.

---

## Checklist nuovo report

- [ ] Letto il `main.py` attuale dal server prima di modificarlo
- [ ] Aggiunto `from app.routers import nuovo_report` agli import esistenti
- [ ] Aggiunto `app.include_router(...)` in coda ai router esistenti
- [ ] Caricato in staging solo i file nuovi + il `main.py` fuso
- [ ] Verificato dopo il deploy che tutti i report precedenti rispondono ancora
- [ ] Aggiornata la tabella "Stato router attivi" in questo file
