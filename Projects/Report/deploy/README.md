# Deploy automatico Yeppon Reports

Script per automatizzare il deploy degli aggiornamenti al sistema di report
(ASP admin + backend Python FastAPI) sul server di produzione.

## Idea

Uno script PowerShell sul server che:

1. Legge una cartella di **staging** dove tu depositi via FileZilla i file modificati
2. Fa **backup** dei file di produzione che sta per sovrascrivere
3. **Copia** i file nelle loro posizioni corrette
4. **Patcha il token** leggendo il valore vero da `.env` (cosi' non lo tieni sul tuo PC)
5. Fa **pip install** solo se `requirements.txt` e' cambiato
6. Fa **restart** del servizio NSSM solo se i file Python sono cambiati
7. Fa uno **smoke test** health check
8. **Archivia** lo staging in una cartella separata (staging pulita dopo ogni deploy)

Tempo tipico: 10-30 secondi.

## Installazione una tantum

Dal tuo PC carica (con FileZilla o copia RDP) sul server tutta la cartella
`deploy\` dentro `D:\admin_yeppon_python\apps\yeppon_reports\`:

```
[PC locale]                               [server]
deploy\deploy.ps1            --->   D:\admin_yeppon_python\apps\deploy\deploy.ps1
deploy\watcher.ps1           --->   D:\admin_yeppon_python\apps\deploy\watcher.ps1
deploy\install-watcher.ps1   --->   D:\admin_yeppon_python\apps\deploy\install-watcher.ps1
deploy\README.md             --->   D:\admin_yeppon_python\apps\deploy\README.md
```

Crea le cartelle di staging e backup (solo la prima volta):

```powershell
New-Item -ItemType Directory -Path "D:\admin_yeppon_python\staging" -Force
New-Item -ItemType Directory -Path "D:\admin_yeppon_python\deploy\backups"        -Force
New-Item -ItemType Directory -Path "D:\admin_yeppon_python\deploy\archived"       -Force
```

(La cartella `D:\admin_yeppon_python\staging\` e' la staging dove caricherai i file
ogni volta. `D:\admin_yeppon_python\deploy\backups\` e `D:\admin_yeppon_python\deploy\archived\` li crea automaticamente
lo script, ma meglio farli da subito.)

## Workflow tipico

### 1. Sul tuo PC: modifichi i file

Apri `actual_turnover_general.asp` o i file Python, modifichi, salvi.

### 2. Via FileZilla: carica solo quello che hai cambiato

La cartella di staging sul server deve rispecchiare la struttura di produzione.
Lo script copia tutto quello che trova sotto `asp_nuovo\` mantenendo i
sottopath (finisce sotto `D:\www\yeppon.it\admin\`). Stessa cosa per i file
Python sotto `python_backend\` (finiscono sotto `D:\admin_yeppon_python\apps\yeppon_reports\python_backend\`).

```
D:\admin_yeppon_python\staging\
├── asp_nuovo\
│   └── actualturnover\                     (mantieni il nome della sottocartella in admin\)
│       ├── actual_turnover_general.asp     (se lo hai cambiato)
│       └── download-turnover-xlsx.asp      (se lo hai cambiato)
└── python_backend\
    ├── app\
    │   ├── main.py                         (se lo hai cambiato)
    │   ├── config.py
    │   ├── routers\
    │   │   └── turnover.py
    │   └── services\
    │       ├── turnover_query.py
    │       └── excel_export.py
    └── requirements.txt                    (se lo hai cambiato)
```

Mapping staging -> produzione:

| Staging | Produzione |
|---|---|
| `asp_nuovo\actualturnover\foo.asp` | `D:\www\yeppon.it\admin\actualturnover\foo.asp` |
| `asp_nuovo\spedizioni\bar.asp` | `D:\www\yeppon.it\admin\spedizioni\bar.asp` |
| `python_backend\app\routers\turnover.py` | `D:\admin_yeppon_python\apps\yeppon_reports\python_backend\app\routers\turnover.py` |

Attenzione: se metti un `.asp` **direttamente sotto `asp_nuovo\`** (senza
sottocartella), finisce in `D:\www\yeppon.it\admin\` nudo. Quasi sempre non
e' quello che vuoi: crea sempre la sottocartella corretta
(`actualturnover\`, `spedizioni\`, ecc.).

**Non devi caricare tutto.** Carica solo i file modificati, mantenendo
la struttura delle cartelle.

**Non caricare:** `.env`, `venv\`, `__pycache__\`, `*.pyc`, file `.bak` / `.old`.
(Lo script li ignorerebbe comunque, ma meglio tenerli fuori.)

### 3. Sul server: RDP + PowerShell Admin

```powershell
cd D:\admin_yeppon_python\apps\yeppon_reports
.\deploy.ps1
```

Lo script:

- Elenca cosa sta per fare
- Chiede conferma `(y/N)`
- Procede: backup -> copia -> patch token -> pip (se serve) -> restart (se serve) -> health check
- Alla fine stampa il **tag del backup** utile per rollback

### 4. Verifica

Apri l'admin nel browser e controlla che tutto funzioni.

## Comandi utili

### Preview senza scrivere niente (dry run)

```powershell
.\deploy.ps1 -DryRun
```

Mostra tutto quello che farebbe, senza toccare niente. Utile la prima volta
o quando non sei sicuro di cosa hai caricato in staging.

### Deploy senza chiedere conferma

```powershell
.\deploy.ps1 -AutoConfirm
```

### Rollback all'ultimo backup

```powershell
.\deploy.ps1 -Rollback
```

Ripristina i file di PRIMA dell'ultimo deploy. Restart automatico del servizio
se avevi cambiato file Python.

### Rollback a un backup specifico

```powershell
.\deploy.ps1 -Rollback -BackupTag "20260421_143022"
```

I tag dei backup disponibili li trovi in `D:\admin_yeppon_python\deploy\backups\`.

### Deploy Python senza restart automatico

Se stai facendo piu' deploy a catena e vuoi riavviare una volta sola alla fine:

```powershell
.\deploy.ps1 -SkipRestart
# ... altri deploy ...
.\deploy.ps1              # ultimo deploy, restart automatico
```

O forzare il restart manuale:

```powershell
& D:\admin_yeppon_python\bin\nssm.exe restart YepponReportsAPI
```

## Esempi di casi d'uso

### Caso 1 — cambiato solo il colore di un'intestazione nell'ASP

Modifica CSS dentro `actual_turnover_general.asp`, lo carichi in
`D:\admin_yeppon_python\staging\asp_nuovo\actualturnover\actual_turnover_general.asp`, lanci:

```powershell
.\deploy.ps1 -AutoConfirm
```

Output atteso: "Nessun restart necessario (solo ASP cambiato)". **5 secondi**.

### Caso 2 — aggiunto un nuovo endpoint a turnover.py

Modifichi `routers/turnover.py` e `services/turnover_query.py`. Li carichi
in staging:

```
D:\admin_yeppon_python\staging\python_backend\app\routers\turnover.py
D:\admin_yeppon_python\staging\python_backend\app\services\turnover_query.py
```

Lanci `.\deploy.ps1`. Lo script copia, restart NSSM, health check. **15 secondi**.

### Caso 3 — aggiunta una nuova dipendenza Python

Aggiorni `requirements.txt` localmente (aggiungi la riga `nomepacchetto==1.2.3`),
carichi in staging:

```
D:\admin_yeppon_python\staging\python_backend\requirements.txt
D:\admin_yeppon_python\staging\python_backend\app\services\qualcosa.py
```

Lo script nota che `requirements.txt` e' cambiato, attiva il venv, fa `pip install`,
poi restart. **30-60 secondi** (dipende dalla libreria).

### Caso 4 — deploy andato male, torno indietro

```powershell
.\deploy.ps1 -Rollback
```

Ripristina backup + restart se serve + health check. **10 secondi**.

## Come funziona il patch del token

Il file `actual_turnover_general.asp` del tuo PC locale contiene questa riga:

```asp
INTERNAL_TOKEN = "change_me_long_random_string"
```

E' un **placeholder**. Lo script, quando copia l'ASP sul server, lo sostituisce
con il valore vero letto da `D:\admin_yeppon_python\apps\yeppon_reports\python_backend\.env`.

Vantaggi:
- il token vero non sta mai sul PC di sviluppo
- se ruoti il token in `.env`, al prossimo deploy l'ASP viene patchato col nuovo
- lo stesso file puoi metterlo in git senza leak

Se il token nel `.env` cambia ma tu non ridepoli l'ASP, l'admin continua a
usare quello vecchio -> 401 dall'API. In quel caso: ricarica l'ASP in staging
(senza modifiche) e rilancia `.\deploy.ps1`, verra' ripatchato.

## Come sono organizzati backup e archivi

Dopo ogni deploy:

```
D:\admin_yeppon_python\deploy\
├── yeppon_reports\         (staging, sempre vuota dopo il deploy)
├── backups\
│   ├── 20260421_143022\    (file di produzione PRIMA di questo deploy)
│   │   ├── asp\
│   │   │   └── actual_turnover_general.asp
│   │   └── python\
│   │       └── app\
│   │           └── main.py
│   ├── 20260420_120015\
│   └── ...                 (ultimi 10 conservati, piu' vecchi cancellati)
└── archived\
    ├── 20260421_143022\    (file che avevi caricato in staging in quel deploy)
    │   ├── asp_nuovo\
    │   └── python_backend\
    ├── 20260420_120015\
    └── ...
```

- `backups\` serve per **rollback** (i vecchi file di produzione)
- `archived\` serve per **traccia/audit** (cosa hai caricato in quel deploy)

Lo script mantiene gli ultimi 10 di entrambi, i piu' vecchi li cancella
automaticamente.

## Troubleshooting

### "cartella staging non trovata: D:\admin_yeppon_python\staging"

Crea la cartella: `New-Item -ItemType Directory -Path "D:\admin_yeppon_python\staging" -Force`

### "INTERNAL_TOKEN non definito in .env"

Verifica che `D:\admin_yeppon_python\apps\yeppon_reports\python_backend\.env` abbia una riga
tipo `INTERNAL_TOKEN=qualcosaqui` (senza spazi prima del segno =).

### "pip install fallito"

Probabile problema di rete o pacchetto non trovato. Lo script si ferma qui
e NON fa il restart. Puoi riprovare a mano:

```powershell
cd D:\admin_yeppon_python\apps\yeppon_reports\python_backend
venv\Scripts\activate
pip install -r requirements.txt
.\nssm.exe restart YepponReportsAPI
```

### "servizio NON risponde dopo restart"

Lo script esce con codice 2 e ti dice di guardare `stderr.log`. Opzioni:

1. Diagnostica: `Get-Content D:\admin_yeppon_python\apps\yeppon_reports\logs\stderr.log -Tail 30`
2. Rollback: `.\deploy.ps1 -Rollback`

### Il servizio gira, ma l'admin mostra 401

Il token nell'ASP non combacia con quello nel `.env`. Rilancia il deploy
anche senza modifiche: l'ASP viene ri-patchato col token corrente.

### Voglio vedere cosa cambierebbe senza fare niente

```powershell
.\deploy.ps1 -DryRun
```

### Voglio cancellare tutti i backup

```powershell
Remove-Item D:\admin_yeppon_python\deploy\backups\* -Recurse -Force
Remove-Item D:\admin_yeppon_python\deploy\archived\* -Recurse -Force
```

(Meglio non farlo: occupano poco e ti salvano la vita in caso di rollback.)

## Auto-deploy (watcher)

Opzionale. Se lo attivi, carichi i file via FileZilla e basta: 30 secondi
dopo aver finito il deploy parte da solo, senza RDP.

### Come funziona

Un secondo servizio Windows chiamato `YepponReportsDeployWatcher` gira
in background e tiene d'occhio la cartella `D:\admin_yeppon_python\staging\`.

- Ogni modifica (create / change / rename / delete) resetta un timer
- Quando passano **30 secondi** senza altre modifiche, il watcher lancia
  `deploy.ps1 -AutoConfirm`
- Se intanto carichi altri file, il timer si resetta (niente deploy
  sovrapposti)
- Se la staging e' vuota (es. dopo un'archiviazione), il watcher non
  fa niente

E' un servizio Windows NSSM come il backend Python: si avvia con Windows,
si riavvia da solo se crasha, e logga tutto.

### Installazione (una tantum)

```powershell
cd D:\admin_yeppon_python\apps\deploy
.\install-watcher.ps1
```

Lo script crea il servizio, lo configura e lo avvia. Output atteso alla
fine: `servizio in esecuzione.`

### Verifica che funzioni

Prova rapida senza toccare la produzione: crea e cancella un file vuoto in
staging, poi guarda il log:

```powershell
New-Item D:\admin_yeppon_python\staging\test.txt -ItemType File -Force
Remove-Item D:\admin_yeppon_python\staging\test.txt -Force
Get-Content D:\admin_yeppon_python\deploy\watcher.log -Tail 20 -Wait
```

Dovresti vedere righe tipo:

```
2026-04-22 10:30:01 | INFO  | silenzio per 30s ma staging vuota, reset
```

(vuota perche' hai cancellato il file subito dopo — niente deploy parte)

### Flusso di lavoro con il watcher attivo

1. Sul tuo PC modifichi i file
2. Via FileZilla carichi in `D:\admin_yeppon_python\staging\asp_nuovo\...` o
   `\python_backend\...`
3. **Aspetti 30 secondi.** Niente altro da fare.
4. Il deploy parte da solo, logga su `D:\admin_yeppon_python\deploy\watcher.log`
5. Se era un file Python, il servizio `YepponReportsAPI` viene riavviato
6. Verifichi l'admin nel browser

### Cosa fare se vai di fretta e carichi piu' file

Il watcher aspetta 30s dall'**ultima** modifica. Quindi se carichi 5 file
in 10 secondi, parte 30s dopo l'ultimo. Nessun rischio di deploy parziale.

Se pero' il collegamento FileZilla e' lento e un upload solo ci mette piu'
di 30s, devi:

- Aumentare `QuietSeconds` in `watcher.ps1` (poi restart del servizio), oppure
- Stoppare il watcher mentre carichi:

```powershell
& D:\admin_yeppon_python\bin\nssm.exe stop YepponReportsDeployWatcher
# carica tutto via FileZilla
& D:\admin_yeppon_python\bin\nssm.exe start YepponReportsDeployWatcher
# 30s dopo, il deploy parte
```

### Comandi utili

```powershell
# stato
& D:\admin_yeppon_python\bin\nssm.exe status YepponReportsDeployWatcher

# stop/start/restart
& D:\admin_yeppon_python\bin\nssm.exe stop    YepponReportsDeployWatcher
& D:\admin_yeppon_python\bin\nssm.exe start   YepponReportsDeployWatcher
& D:\admin_yeppon_python\bin\nssm.exe restart YepponReportsDeployWatcher

# log del watcher (righe gia' formattate)
Get-Content D:\admin_yeppon_python\deploy\watcher.log -Tail 50 -Wait

# log servizio NSSM (crash, stdout raw)
Get-Content D:\admin_yeppon_python\deploy\watcher-service-stderr.log -Tail 50 -Wait

# disinstallare
cd D:\admin_yeppon_python\apps\deploy
.\install-watcher.ps1 -Uninstall
```

### Troubleshooting watcher

**"Il servizio parte e si ferma subito"**

Guarda `D:\admin_yeppon_python\deploy\watcher-service-stderr.log`. Tipicamente: path sbagliato
in `watcher.ps1` o NSSM lanciato senza privilegi admin.

**"Carico un file ma il deploy non parte"**

Verifica prima di tutto che il servizio giri:

```powershell
& D:\admin_yeppon_python\bin\nssm.exe status YepponReportsDeployWatcher
```

Deve dire `SERVICE_RUNNING`. Poi guarda il log del watcher:

```powershell
Get-Content D:\admin_yeppon_python\deploy\watcher.log -Tail 30
```

Se non vedi nessun evento registrato, FileSystemWatcher non ha intercettato
l'upload (puo' succedere con alcuni protocolli di rete esotici — non dovrebbe
con FileZilla/SFTP standard, ma controlla). In extremis puoi sempre lanciare
`deploy.ps1` a mano, il watcher non toglie niente.

**"Il deploy parte ma fallisce"**

Il watcher intercetta il fallimento e lo logga, ma non ci sono ritentativi
automatici. Il fallimento finisce in `D:\admin_yeppon_python\deploy\watcher.log` con `[deploy]`
davanti a ogni riga di output. Se il problema e' un file caricato male,
ricarica e il watcher ripartira' (30s dopo).

**"Voglio il watcher ma con un'attesa diversa"**

Modifica `QuietSeconds` in cima a `D:\admin_yeppon_python\apps\deploy\watcher.ps1`
e restart del servizio. Suggerimento: 30s va bene per upload FTP standard,
60-90s se la connessione e' lenta.

## Configurazione

I path sono tutti in cima al file `deploy.ps1` (hashtable `$CFG`) e, se
usi il watcher, anche in `watcher.ps1` e `install-watcher.ps1` (stessa
struttura `$CFG`). Se cambi qualcosa sul server (es. sposti il backend
in un altro path), modifica quelle sezioni in testa ai file.
