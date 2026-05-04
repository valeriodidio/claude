# Yeppon Reports API

Backend FastAPI per la reportistica dell'admin. Espone endpoint JSON che l'admin Classic ASP consuma via `fetch()` oppure tramite proxy `MSXML2.ServerXMLHTTP` (per l'export Excel).

Architettura:

```
Browser utente
    |
    v
IIS (Classic ASP admin)  <-- autenticazione utente esistente
    |
    | fetch AJAX  (dati JSON)
    | proxy ASP  (download Excel)
    v
FastAPI (localhost:8000) <-- SOLO localhost, non esposto a internet
    |
    v
MySQL (yeppon_stats.actual_turnover)
```

Il servizio FastAPI **non deve mai essere esposto a internet**. Gira sul
loopback (127.0.0.1) e viene chiamato:
- dal browser tramite reverse proxy IIS su `/api/reports/*`
- dall'ASP tramite `MSXML2.ServerXMLHTTP` su `http://127.0.0.1:8000`

## 1. Prerequisiti

- Windows Server con IIS (quello che gi\u00e0 ospita Classic ASP)
- Python 3.11+ (installabile da python.org, spuntare "Add Python to PATH")
- Pacchetto [NSSM](https://nssm.cc/download) per eseguire il servizio come Windows Service
- Il server deve poter raggiungere il MySQL `yeppon_stats`

## 2. Installazione

Dal prompt Amministratore, nella cartella dove copi il progetto (es. `C:\apps\yeppon_reports\`):

```bat
cd C:\apps\yeppon_reports\python_backend
python -m venv venv
venv\Scripts\activate
pip install -r requirements.txt
copy .env.example .env
notepad .env
```

Compila `.env`:

```ini
DB_HOST=your-mysql-host
DB_PORT=3306
DB_USER=reports_reader
DB_PASSWORD=...
DB_NAME=yeppon_stats

APP_HOST=127.0.0.1
APP_PORT=8000

INTERNAL_TOKEN=<un valore random lungo, es. generato con "python -c 'import secrets; print(secrets.token_hex(32))'">
```

**Importante:** lo stesso `INTERNAL_TOKEN` va messo in:
- `.env` del backend
- `asp_nuovo\actual_turnover_general.asp` (variabile `INTERNAL_TOKEN`)
- `asp_nuovo\download-turnover-xlsx.asp` (variabile `INTERNAL_TOKEN`)

Meglio ancora: spostalo in un include ASP protetto (`/int/reports_config.asp`) e includilo da entrambi i file. Cos\u00ec ruoti il token cambiando un solo file.

## 3. Test manuale

Con il `venv` attivo:

```bat
python -m uvicorn app.main:app --host 127.0.0.1 --port 8000
```

Verifica:

```bat
curl http://127.0.0.1:8000/api/reports/health
# {"status":"ok"}
```

Con token:

```bat
curl -H "X-Internal-Token: <il tuo token>" ^
  "http://127.0.0.1:8000/api/reports/turnover/summary?dal=2026-03-01&al=2026-03-31"
```

Documentazione interattiva (Swagger UI, **solo per test** — non esporla):
http://127.0.0.1:8000/api/reports/docs

## 4. Installazione come Windows Service (NSSM)

```bat
nssm install YeppunReportsAPI "C:\apps\yeppon_reports\python_backend\venv\Scripts\python.exe"
nssm set YeppunReportsAPI AppParameters "-m uvicorn app.main:app --host 127.0.0.1 --port 8000"
nssm set YeppunReportsAPI AppDirectory "C:\apps\yeppon_reports\python_backend"
nssm set YeppunReportsAPI AppStdout "C:\apps\yeppon_reports\logs\stdout.log"
nssm set YeppunReportsAPI AppStderr "C:\apps\yeppon_reports\logs\stderr.log"
nssm set YeppunReportsAPI Start SERVICE_AUTO_START
nssm start YeppunReportsAPI
```

Controllo:

```bat
sc query YeppunReportsAPI
nssm status YeppunReportsAPI
```

Log: `C:\apps\yeppon_reports\logs\stdout.log` e `stderr.log`.

## 5. Reverse proxy IIS

Serve per far raggiungere a un browser (autenticato sull'admin) gli endpoint Python senza esporli direttamente. Il browser chiama `/api/reports/...` sullo stesso host dell'admin; IIS gira le chiamate a `localhost:8000`.

**Prerequisiti IIS** (una volta sola):
- Installa [URL Rewrite](https://www.iis.net/downloads/microsoft/url-rewrite)
- Installa [Application Request Routing (ARR)](https://www.iis.net/downloads/microsoft/application-request-routing)
- In IIS Manager \u2192 server root \u2192 "Application Request Routing Cache" \u2192 Server Proxy Settings \u2192 **Enable proxy** \u2705

Poi nel sito dell'admin, aggiungi al `web.config` (nella root o in `/api/`):

```xml
<configuration>
  <system.webServer>
    <rewrite>
      <rules>
        <rule name="Proxy to Python Reports API" stopProcessing="true">
          <match url="^api/reports/(.*)" />
          <action type="Rewrite" url="http://127.0.0.1:8000/api/reports/{R:1}" />
          <serverVariables>
            <set name="HTTP_X_FORWARDED_HOST" value="{HTTP_HOST}" />
          </serverVariables>
        </rule>
      </rules>
    </rewrite>
  </system.webServer>
</configuration>
```

Verifica dal browser (dopo login admin):

```
https://admin.tuosito.it/api/reports/health
```

## 6. Deploy delle pagine ASP

Copia i file da `asp_nuovo\` nella cartella corrispondente dell'admin:

| Origine                                        | Destinazione                                   |
|------------------------------------------------|------------------------------------------------|
| `asp_nuovo\actual_turnover_general.asp`         | `C:\inetpub\wwwroot\admin\actual_turnover_general.asp` (o dove sta il vecchio) |
| `asp_nuovo\download-turnover-xlsx.asp`          | stessa cartella della pagina sopra              |

**Tieni da parte il file originale** (`actual_turnover_general.asp`) rinominato in `.asp.old`, cos\u00ec puoi fare rollback in 5 secondi.

## 7. Utente MySQL dedicato

Crea un utente **solo lettura** sulla tabella `actual_turnover`:

```sql
CREATE USER 'reports_reader'@'%' IDENTIFIED BY '<password lunga random>';
GRANT SELECT ON yeppon_stats.actual_turnover TO 'reports_reader'@'%';
FLUSH PRIVILEGES;
```

Se l'app Python gira sullo stesso host del DB, sostituisci `%` con `localhost`.

## 8. Troubleshooting

**Il servizio non parte:** controlla `stderr.log` in `C:\apps\yeppon_reports\logs\`. Spesso \u00e8 un errore di `.env` (tipicamente password DB o host sbagliato).

**CORS errors dal browser:** se l'admin gira su un dominio diverso dal proxy IIS (non dovrebbe), imposta `ALLOWED_ORIGINS=https://admin.tuosito.it` in `.env`.

**401 Unauthorized sugli endpoint:** `INTERNAL_TOKEN` non coincide tra `.env` e i file ASP.

**Pagina bianca con browser DevTools "Blocked by CORS":** sicuro che chiami via `/api/reports/...` (stesso dominio) e non `http://127.0.0.1:8000/...`?

**Il proxy IIS restituisce 502:** il servizio Python \u00e8 gi\u00f9. `nssm status YeppunReportsAPI`.

**Query molto lenta:** controlla indici su `actual_turnover`. Consigliati:
```sql
CREATE INDEX idx_at_dataordine ON actual_turnover(dataordine);
CREATE INDEX idx_at_macrocat   ON actual_turnover(macrocat);
CREATE INDEX idx_at_nomemktp   ON actual_turnover(nomemktp);
CREATE INDEX idx_at_codice     ON actual_turnover(codice);
```

## 9. Struttura file

```
python_backend/
  app/
    __init__.py
    main.py              # FastAPI app + middleware + router registration
    config.py            # settings (.env via pydantic-settings)
    db.py                # SQLAlchemy engine singleton
    routers/
      __init__.py
      turnover.py        # endpoint /api/reports/turnover/*
    services/
      __init__.py
      turnover_query.py  # query MySQL + aggregazioni pandas
      excel_export.py    # generazione .xlsx con openpyxl
  requirements.txt
  .env.example
  README.md  (questo file)
```

## 10. Endpoint disponibili

| Metodo | Endpoint                                        | Descrizione |
|--------|-------------------------------------------------|-------------|
| GET    | `/api/reports/health`                           | Ping        |
| GET    | `/api/reports/turnover/filters`                 | Opzioni statiche (marketplace, fornitori, macrocat, user_type) |
| GET    | `/api/reports/turnover/filters/dependent`       | Opzioni in cascata (cate2, cate3, marca) sui filtri attuali |
| GET    | `/api/reports/turnover/summary`                 | Tabella principale: riga per macrocat + totale |
| GET    | `/api/reports/turnover/drilldown/{level}`       | Drill inline; level \u2208 {cate2, cate3, marca, fornitore} |
| GET    | `/api/reports/turnover/marketplace-breakdown`   | Ripartizione per marketplace |
| GET    | `/api/reports/turnover/trend/{dimension}`       | Trend: day, hour, weekday, month, province, marketplace |
| GET    | `/api/reports/turnover/product-trend?codice=..` | Trend giornaliero di un singolo prodotto |
| GET    | `/api/reports/turnover/export.xlsx`             | Export Excel multi-sheet |

Query parameters comuni (tutti gli endpoint report):

- `dal`, `al` (obbligatori): `YYYY-MM-DD`
- `marketplace` (multi): `?marketplace=Amazon&marketplace=eBay`
- `fornitore`, `macrocat`, `cate2`, `cate3`, `marca`, `codice`, `user_type` (opzionali)
