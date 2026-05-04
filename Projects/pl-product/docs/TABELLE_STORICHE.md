# Tabelle storiche P&L e Resi

Documentazione delle due tabelle su `yeppon_stats` che alimentano il report
P&L prodotti e il report dettaglio resi.

---

## Indice

1. [Architettura generale](#architettura-generale)
2. [`yeppon_stats.pl_prodotti` — snapshot multi-periodo](#yeppon_statspl_prodotti)
3. [`yeppon_stats.pl_prodotti_corrieri` — dettaglio corriere](#yeppon_statspl_prodotti_corrieri)
4. [`yeppon_stats.pl_prodotti_corrieri_dettaglio` — dettaglio per ordine](#yeppon_statspl_prodotti_corrieri_dettaglio)
5. [`yeppon_stats.pl_prodotti_marketplace` — breakdown marketplace](#yeppon_statspl_prodotti_marketplace)
6. [`yeppon_stats.resi_impatto_economico` — storico cumulativo](#yeppon_statsresi_impatto_economico)
7. [Workflow di popolamento](#workflow-di-popolamento)
8. [Query tipo per il report](#query-tipo-per-il-report)
9. [Manutenzione](#manutenzione)

---

## Architettura generale

Le due tabelle hanno **strategie di storicizzazione diverse** perché contengono
dati di natura diversa.

| Tabella | Natura del dato | Strategia |
|---|---|---|
| `pl_prodotti` | Aggregato calcolato su una **finestra mobile** di N giorni | Snapshot: una "fotografia" per ogni `(data_snapshot, periodo_giorni)` |
| `pl_prodotti_corrieri` | Aggregato spedizioni per (prodotto, corriere) sulla stessa finestra | Snapshot per `(data_snapshot, periodo_giorni, id_p, corriere)` |
| `pl_prodotti_corrieri_dettaglio` | Dettaglio spedizioni per **singolo ordine** (drill-down) | Snapshot per `(data_snapshot, periodo_giorni, id_p, id_ordine)` |
| `pl_prodotti_marketplace` | Vendite/margine per (prodotto, marketplace) sulla stessa finestra | Snapshot per `(data_snapshot, periodo_giorni, id_p, marketplace)` |
| `resi_impatto_economico` | **Eventi atomici** (singoli RMA con la propria `data_rma`) | UPSERT continuo, niente snapshot. Si filtra per `data_rma` |

Lo script che le popola:

- `export_pl_prodotti.py` → scrive su `pl_prodotti`, `pl_prodotti_corrieri`,
  `pl_prodotti_corrieri_dettaglio` **e `pl_prodotti_marketplace`** in un unico
  run, salva 6 snapshot in un colpo solo (uno per ciascun
  `PERIODO ∈ {30, 60, 90, 120, 180, 360}`).
- `export_dettaglio_resi_completo.py` → scrive su `resi_impatto_economico`,
  gira ogni giorno, fa UPSERT della finestra `PERIODO=180gg`.

---

## `yeppon_stats.pl_prodotti`

### Cos'è

Una riga = il **P&L di un prodotto** calcolato su una specifica finestra
temporale (`periodo_giorni`) a una specifica data di calcolo
(`data_snapshot`).

Esempio:

| data_snapshot | id_p | periodo_giorni | tot_fatturato | margine_effettivo |
|---|---|---|---|---|
| 2026-04-28 | 12345 | 30  | 1.200,00 |   180,00 |
| 2026-04-28 | 12345 | 90  | 4.500,00 |   720,00 |
| 2026-04-28 | 12345 | 360 | 18.000,00 | 2.900,00 |
| 2026-04-29 | 12345 | 30  | 1.350,00 |   210,00 |
| ...        | ...   | ... | ...      | ...      |

### Chiave logica

```
(data_snapshot, id_p, periodo_giorni)   ← UNIQUE
```

Significa: **per ogni giorno** in cui gira il batch, **per ogni prodotto** e
**per ogni periodo di analisi**, c'è esattamente una riga.

### Schema

```sql
CREATE TABLE IF NOT EXISTS yeppon_stats.pl_prodotti (
  id            BIGINT UNSIGNED AUTO_INCREMENT PRIMARY KEY,
  id_p          INT UNSIGNED NOT NULL,
  codice        VARCHAR(50)  NOT NULL DEFAULT '',
  nome          VARCHAR(255) NOT NULL DEFAULT '',
  marca         VARCHAR(100) NOT NULL DEFAULT '',
  disp_fornitore INT NOT NULL DEFAULT 0,
  status_prodotto TINYINT NOT NULL DEFAULT 0,
  bloccato      TINYINT NOT NULL DEFAULT 0,
  bloccato_fino DATE DEFAULT NULL,
  num_ordini    INT UNSIGNED NOT NULL DEFAULT 0,
  tot_pezzi     INT UNSIGNED NOT NULL DEFAULT 0,
  tot_fatturato DOUBLE NOT NULL DEFAULT 0,
  tot_margine   DOUBLE NOT NULL DEFAULT 0,
  perc_margine  DOUBLE NOT NULL DEFAULT 0,
  tot_sped_fatt DOUBLE NOT NULL DEFAULT 0,
  tot_sped_costi DOUBLE NOT NULL DEFAULT 0,
  delta_sped    DOUBLE NOT NULL DEFAULT 0,
  perc_delta_sped DOUBLE NOT NULL DEFAULT 0,
  qty_resi      INT UNSIGNED NOT NULL DEFAULT 0,
  perc_resi     DOUBLE NOT NULL DEFAULT 0,
  imp_eco_resi  DOUBLE NOT NULL DEFAULT 0,
  imp_Danni     DOUBLE NOT NULL DEFAULT 0,
  imp_Danno_rientrato DOUBLE NOT NULL DEFAULT 0,
  imp_Difettosi DOUBLE NOT NULL DEFAULT 0,
  imp_Giacenza  DOUBLE NOT NULL DEFAULT 0,
  imp_Mancato_Ritiro DOUBLE NOT NULL DEFAULT 0,
  imp_Prodotto_non_conforme DOUBLE NOT NULL DEFAULT 0,
  imp_Recesso   DOUBLE NOT NULL DEFAULT 0,
  imp_Reclamo_contestazioni DOUBLE NOT NULL DEFAULT 0,
  imp_Smarrimenti DOUBLE NOT NULL DEFAULT 0,
  margine_effettivo DOUBLE NOT NULL DEFAULT 0,
  perc_margine_eff  DOUBLE NOT NULL DEFAULT 0,
  periodo_giorni INT UNSIGNED NOT NULL DEFAULT 180,
  data_snapshot DATE NOT NULL DEFAULT (CURRENT_DATE),
  data_calcolo  DATETIME NOT NULL DEFAULT CURRENT_TIMESTAMP,
  UNIQUE KEY uk_snap_prod_periodo (data_snapshot, id_p, periodo_giorni),
  INDEX idx_snapshot_periodo (data_snapshot, periodo_giorni),
  INDEX idx_periodo_snapshot (periodo_giorni, data_snapshot),
  INDEX idx_id_p (id_p),
  INDEX idx_codice (codice)
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4;
```

### Significato delle colonne principali

| Colonna | Cosa rappresenta |
|---|---|
| `data_snapshot` | Giorno in cui è stato calcolato lo snapshot (= `CURDATE()` al momento del run) |
| `periodo_giorni` | Finestra di analisi: 30, 60, 90, 120, 180 o 360 giorni |
| `data_calcolo` | Timestamp esatto del run (utile per audit) |
| `tot_fatturato` | Vendite (`venduto_tot + spese_sped_grat`) nel periodo |
| `tot_margine` | Margine prima di spedizioni e resi |
| `tot_sped_fatt` / `tot_sped_costi` / `delta_sped` | Incasso, costo e delta spedizioni ricalcolati con pesi corretti |
| `qty_resi` | Pezzi resi nel periodo (somma `quantita_rma`) |
| `imp_eco_resi` | Impatto economico totale resi (somma di tutte le `imp_*`) |
| `imp_Danni`, `imp_Difettosi`, ... | Impatto resi suddiviso per tipologia RMA |
| `margine_effettivo` | `tot_margine + delta_sped - imp_eco_resi` |
| `perc_*` | Percentuali derivate (calcolate dal Python, non sommabili tra periodi/prodotti) |

### Comportamento del batch

A ogni run:
1. Per ciascun `periodo ∈ PERIODI`:
   - Cancella le righe esistenti per `(CURDATE(), periodo)` → idempotenza giornaliera.
   - Inserisce una riga per ogni prodotto con vendite nel periodo.
2. Le righe di **giorni precedenti** non vengono toccate → storico permanente.

> **Volume stimato:** ~50k prodotti × 6 periodi × 365 giorni ≈ 110M righe/anno
> a regime con snapshot giornaliero. Considerare snapshot settimanale se non
> serve la granularità giornaliera.

---

## `yeppon_stats.pl_prodotti_corrieri`

### Cos'è

Una riga = il **delta spedizioni di un prodotto su uno specifico corriere**
sulla stessa finestra temporale di `pl_prodotti`. Serve per analizzare a
posteriori quale corriere genera marginalità (positiva o negativa) prodotto
per prodotto.

Esempio: il prodotto `12345` nel periodo a 90gg ha viaggiato con BRT, GLS e
SDA. Una riga per ciascun corriere.

| data_snapshot | id_p | periodo_giorni | corriere | num_ordini | tot_sped_fatt | tot_sped_costi | delta_sped |
|---|---|---|---|---|---|---|---|
| 2026-04-28 | 12345 | 90 | BRT | 120 | 480,00 | 510,00 | -30,00 |
| 2026-04-28 | 12345 | 90 | GLS |  45 | 180,00 | 165,00 | +15,00 |
| 2026-04-28 | 12345 | 90 | SDA |  20 |  90,00 |  82,00 | +8,00  |

### Chiave logica

```
(data_snapshot, periodo_giorni, id_p, corriere)   ← UNIQUE
```

### Schema

```sql
CREATE TABLE IF NOT EXISTS yeppon_stats.pl_prodotti_corrieri (
  id              BIGINT UNSIGNED AUTO_INCREMENT PRIMARY KEY,
  data_snapshot   DATE         NOT NULL,
  periodo_giorni  INT UNSIGNED NOT NULL,
  id_p            INT UNSIGNED NOT NULL,
  corriere        VARCHAR(50)  NOT NULL DEFAULT '',
  num_ordini      INT UNSIGNED NOT NULL DEFAULT 0,
  tot_pezzi       INT UNSIGNED NOT NULL DEFAULT 0,
  tot_sped_fatt   DOUBLE       NOT NULL DEFAULT 0,
  tot_sped_costi  DOUBLE       NOT NULL DEFAULT 0,
  delta_sped      DOUBLE       NOT NULL DEFAULT 0,
  perc_delta_sped DOUBLE       NOT NULL DEFAULT 0,
  data_calcolo    DATETIME     NOT NULL DEFAULT CURRENT_TIMESTAMP,
  UNIQUE KEY uk_snap_prod_per_corr (data_snapshot, periodo_giorni, id_p, corriere),
  INDEX idx_id_p (id_p),
  INDEX idx_corriere (corriere),
  INDEX idx_snap_per (data_snapshot, periodo_giorni)
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4;
```

### Significato delle colonne principali

| Colonna | Cosa rappresenta |
|---|---|
| `data_snapshot` | Stesso valore di `pl_prodotti` (giorno del run) |
| `periodo_giorni` | Stesso valore di `pl_prodotti` (30/60/90/120/180/360) |
| `corriere` | Nome del corriere: `TNT`, `DHL`, `BRT`, `SDA`, `GLS`, `Fastest`, `Zust`, `Arcese`, `Milkman`, `Fercam`, `Spring` |
| `num_ordini` | Ordini distinti del prodotto spediti con quel corriere |
| `tot_pezzi` | Pezzi venduti (somma quantita per gli ordini in oggetto) |
| `tot_sped_fatt` | Incasso spedizione attribuito a quel corriere (proporzionato per peso) |
| `tot_sped_costi` | Costo spedizione **reale** dalla fattura del corriere (proporzionato per peso) |
| `delta_sped` | `tot_sped_fatt - tot_sped_costi` (positivo = guadagno spedizione) |
| `perc_delta_sped` | `delta_sped / tot_sped_fatt * 100` |

### Cosa NON c'è

- **Ordini senza fattura corriere caricata** sono **esclusi** dalla tabella
  (no fallback tariffario). Quindi la somma di `num_ordini` su tutti i
  corrieri di un prodotto può essere **inferiore** al `num_ordini` di
  `pl_prodotti` per lo stesso `id_p`.
- Ordini Digitec / Asendia (incasso e costo spedizione = 0) sono esclusi.

### Comportamento del batch

A ogni run, dentro lo stesso loop di `pl_prodotti`:
1. Calcola le aggregazioni per `(id_p, corriere)` usando il dettaglio
   ordine × prodotto.
2. Cancella le righe esistenti per `(CURDATE(), periodo)`.
3. Inserisce le nuove righe.

> **Volume stimato:** ~3–5 corrieri attivi per prodotto × 50k prodotti ×
> 6 periodi × 365 giorni ≈ 350–500M righe/anno a regime con snapshot
> giornaliero. Considerare snapshot settimanale o mensile.

---

## `yeppon_stats.pl_prodotti_corrieri_dettaglio`

### Cos'è

Tabella di **drill-down**: una riga per ogni `(id_p, id_ordine)` per ciascuno
snapshot/periodo, con il corriere usato e i numeri di spedizione di quel
singolo ordine. Permette di scendere dal dato aggregato di
`pl_prodotti_corrieri` fino al dettaglio dell'ordine specifico.

Il volume è contenuto (~50k righe per snapshot/periodo) perché ci sono
solo gli ordini con fattura corriere reale e con prodotti tracciati.

### Chiave logica

```
(data_snapshot, periodo_giorni, id_p, id_ordine)   ← UNIQUE
```

### Schema

```sql
CREATE TABLE IF NOT EXISTS yeppon_stats.pl_prodotti_corrieri_dettaglio (
  id              BIGINT UNSIGNED AUTO_INCREMENT PRIMARY KEY,
  data_snapshot   DATE         NOT NULL,
  periodo_giorni  INT UNSIGNED NOT NULL,
  id_p            INT UNSIGNED NOT NULL,
  id_ordine       BIGINT UNSIGNED NOT NULL,
  corriere        VARCHAR(50)  NOT NULL DEFAULT '',
  quantita        INT UNSIGNED NOT NULL DEFAULT 0,
  sped_fatt       DOUBLE       NOT NULL DEFAULT 0,
  sped_costi      DOUBLE       NOT NULL DEFAULT 0,
  delta_sped      DOUBLE       NOT NULL DEFAULT 0,
  data_calcolo    DATETIME     NOT NULL DEFAULT CURRENT_TIMESTAMP,
  UNIQUE KEY uk_snap_per_prod_ord (data_snapshot, periodo_giorni, id_p, id_ordine),
  INDEX idx_id_p (id_p),
  INDEX idx_id_ordine (id_ordine),
  INDEX idx_corriere (corriere),
  INDEX idx_snap_per (data_snapshot, periodo_giorni)
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4;
```

### Significato delle colonne

| Colonna | Cosa rappresenta |
|---|---|
| `data_snapshot` | Giorno del run (= `CURDATE()`) |
| `periodo_giorni` | Finestra di analisi: 30/60/90/120/180/360 |
| `id_p` | Prodotto |
| `id_ordine` | Ordine specifico in cui il prodotto è stato venduto |
| `corriere` | Corriere effettivo della fattura |
| `quantita` | Pezzi del prodotto in quell'ordine |
| `sped_fatt` | Quota incasso spedizione attribuita al prodotto in quell'ordine |
| `sped_costi` | Quota costo spedizione (da fattura corriere) |
| `delta_sped` | `sped_fatt - sped_costi` |
| `data_calcolo` | Timestamp esatto del run |

### Comportamento del batch

A ogni run di `export_pl_prodotti.py`, dentro il loop sui periodi:
1. `DELETE` selettivo per `(CURDATE(), periodo)`.
2. INSERT batch del dettaglio per ordine (solo ordini con fattura
   corriere reale, no fallback tariffario).

### Quando usarla

- Drill-down dal report aggregato: "questo prodotto ha delta negativo su
  BRT, quali ordini specificamente?".
- Verifica spot dell'attribuzione costo/incasso per ordine.
- Analisi storica fine-grained ordine per ordine.

---

## `yeppon_stats.pl_prodotti_marketplace`

### Cos'è

Una riga = vendite di un **prodotto su uno specifico marketplace** calcolate
sulla stessa finestra mobile delle altre tabelle PL
(`periodo_giorni ∈ {30, 60, 90, 120, 180, 360}`).

Serve per il **breakdown per marketplace** del report PL
(`get_marketplace_breakdown` in [app/services/pl_prodotti_query.py](../app/services/pl_prodotti_query.py)),
che può funzionare per singolo prodotto, per marca o aggregato.

### Schema

```sql
CREATE TABLE IF NOT EXISTS yeppon_stats.pl_prodotti_marketplace (
    id              BIGINT UNSIGNED AUTO_INCREMENT PRIMARY KEY,
    data_snapshot   DATE            NOT NULL,
    periodo_giorni  INT UNSIGNED    NOT NULL,
    id_p            INT UNSIGNED    NOT NULL,
    marketplace     VARCHAR(50)     NOT NULL DEFAULT '',
    num_ordini      INT UNSIGNED    NOT NULL DEFAULT 0,
    tot_pezzi       INT UNSIGNED    NOT NULL DEFAULT 0,
    tot_fatturato   DOUBLE          NOT NULL DEFAULT 0,
    tot_margine     DOUBLE          NOT NULL DEFAULT 0,
    perc_margine    DOUBLE          NOT NULL DEFAULT 0,
    tot_sped_fatt   DOUBLE          NOT NULL DEFAULT 0,
    tot_sped_costi  DOUBLE          NOT NULL DEFAULT 0,
    delta_sped      DOUBLE          NOT NULL DEFAULT 0,
    perc_delta_sped DOUBLE          NOT NULL DEFAULT 0,
    qty_resi        INT UNSIGNED    NOT NULL DEFAULT 0,
    perc_resi       DOUBLE          NOT NULL DEFAULT 0,
    imp_eco_resi    DOUBLE          NOT NULL DEFAULT 0,
    margine_effettivo DOUBLE        NOT NULL DEFAULT 0,
    perc_margine_eff  DOUBLE        NOT NULL DEFAULT 0,
    data_calcolo    DATETIME        NOT NULL DEFAULT CURRENT_TIMESTAMP,

    UNIQUE KEY uk_snap_per_prod_mktp (data_snapshot, periodo_giorni, id_p, marketplace),
    INDEX idx_id_p          (id_p),
    INDEX idx_marketplace   (marketplace),
    INDEX idx_snap_per      (data_snapshot, periodo_giorni)
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4;
```

### Significato delle colonne

| Colonna | Cosa rappresenta |
|---|---|
| `data_snapshot` | Data in cui è stato calcolato lo snapshot |
| `periodo_giorni` | Finestra temporale dell'aggregato (30/60/90/120/180/360) |
| `id_p` | ID prodotto |
| `marketplace` | Nome del marketplace (es. "yeppon", "amazon DE", "leroy merlin ES", "stockly") |
| `num_ordini` | Numero ordini distinti su quel marketplace per quel prodotto |
| `tot_pezzi` | Totale pezzi venduti |
| `tot_fatturato` | Fatturato lordo (somma `venduto_tot`) |
| `tot_margine` | Margine totale (`utile - comm_mktp - comm_pagamento`) |
| `perc_margine` | `tot_margine / tot_fatturato * 100` |
| `tot_sped_fatt` | Totale incasso spedizione attribuito a (id_p, marketplace) |
| `tot_sped_costi` | Totale costo spedizione attribuito a (id_p, marketplace) |
| `delta_sped` | `tot_sped_fatt - tot_sped_costi` (positivo = guadagno spedizione) |
| `perc_delta_sped` | `delta_sped / tot_sped_fatt * 100` |
| `qty_resi` | Quantità totale di pezzi resi |
| `perc_resi` | `qty_resi / tot_pezzi * 100` |
| `imp_eco_resi` | Impatto economico totale dei resi (somma di `perdita_netta` per tutti i tipi RMA) |
| `margine_effettivo` | `tot_margine + delta_sped - imp_eco_resi` |
| `perc_margine_eff` | `margine_effettivo / tot_fatturato * 100` |
| `data_calcolo` | Timestamp tecnico (default `CURRENT_TIMESTAMP`) |

### Stato attuale

✅ **La tabella viene popolata** da `export_pl_prodotti.py` tramite
`_save_snapshot_marketplace()`, chiamata dentro `_run_period` dopo la
costruzione dei DataFrame per marketplace (così include anche le metriche
di delta spedizioni, resi e margine effettivo calcolate da `compute()`).
Strategia DELETE+INSERT come per le altre tabelle snapshot: idempotente
per re-run nello stesso giorno, isolata per `(data_snapshot, periodo_giorni)`.

### Migrazione: nuove colonne sped/resi/margine effettivo

Se hai già la tabella con il solo subset (`num_ordini`, `tot_pezzi`,
`tot_fatturato`, `tot_margine`, `perc_margine`), aggiungi le nuove
colonne:

```sql
ALTER TABLE yeppon_stats.pl_prodotti_marketplace
  ADD COLUMN tot_sped_fatt   DOUBLE NOT NULL DEFAULT 0 AFTER perc_margine,
  ADD COLUMN tot_sped_costi  DOUBLE NOT NULL DEFAULT 0 AFTER tot_sped_fatt,
  ADD COLUMN delta_sped      DOUBLE NOT NULL DEFAULT 0 AFTER tot_sped_costi,
  ADD COLUMN perc_delta_sped DOUBLE NOT NULL DEFAULT 0 AFTER delta_sped,
  ADD COLUMN qty_resi        INT UNSIGNED NOT NULL DEFAULT 0 AFTER perc_delta_sped,
  ADD COLUMN perc_resi       DOUBLE NOT NULL DEFAULT 0 AFTER qty_resi,
  ADD COLUMN imp_eco_resi    DOUBLE NOT NULL DEFAULT 0 AFTER perc_resi,
  ADD COLUMN margine_effettivo DOUBLE NOT NULL DEFAULT 0 AFTER imp_eco_resi,
  ADD COLUMN perc_margine_eff  DOUBLE NOT NULL DEFAULT 0 AFTER margine_effettivo;
```

Dopo l'ALTER, rilanciare `export_pl_prodotti.py` per ripopolare lo
snapshot del giorno con i nuovi campi.

### Pattern di lettura tipico

```sql
-- Top 5 marketplace per fatturato di un prodotto, ultimo snapshot 180gg
SELECT marketplace, num_ordini, tot_pezzi, tot_fatturato, tot_margine
FROM yeppon_stats.pl_prodotti_marketplace
WHERE id_p = :id_p
  AND periodo_giorni = 180
  AND data_snapshot = (
    SELECT MAX(data_snapshot)
    FROM yeppon_stats.pl_prodotti_marketplace
    WHERE periodo_giorni = 180
  )
ORDER BY tot_fatturato DESC
LIMIT 5;
```

---

## `yeppon_stats.resi_impatto_economico`

### Cos'è

Una riga = un **prodotto in un RMA** (singolo evento di reso) con tutti i
dettagli di calcolo perdita/recupero.

A differenza di `pl_prodotti` non è uno snapshot: è una **tabella di eventi
storici** che cresce nel tempo, e ogni RMA mantiene la sua `data_rma` fissa.

### Chiave logica

```
(id_rma, id_p)   ← UNIQUE
```

Un prodotto può comparire una sola volta per RMA (un RMA può contenere più
prodotti distinti).

### Schema

```sql
CREATE TABLE IF NOT EXISTS yeppon_stats.resi_impatto_economico (
  id_rma                BIGINT NOT NULL,
  id_ordine             BIGINT NOT NULL,
  id_p                  INT UNSIGNED NOT NULL,
  tipo_rma              VARCHAR(50) NOT NULL DEFAULT '',
  marketplace           VARCHAR(50) NOT NULL DEFAULT '',
  quantita_rma          INT NOT NULL DEFAULT 0,
  costo_unitario        DOUBLE NOT NULL DEFAULT 0,
  prezzo_vendita_unit   DOUBLE NOT NULL DEFAULT 0,
  valore_rimborso       DOUBLE NOT NULL DEFAULT 0,
  perc_rimborso         DOUBLE NOT NULL DEFAULT 0,
  costo_perdita         DOUBLE NOT NULL DEFAULT 0,
  valore_claims         DOUBLE NOT NULL DEFAULT 0,
  valore_recuperato     DOUBLE NOT NULL DEFAULT 0,
  perdita_netta         DOUBLE NOT NULL DEFAULT 0,
  costo_sped_rientro    DOUBLE NOT NULL DEFAULT 0,
  ha_ingresso           TINYINT NOT NULL DEFAULT 0,
  ha_ndc                TINYINT NOT NULL DEFAULT 0,
  nota_recovery         VARCHAR(500) NOT NULL DEFAULT '',
  data_rma              DATETIME NULL,
  data_aggiornamento    DATETIME NOT NULL DEFAULT CURRENT_TIMESTAMP
                        ON UPDATE CURRENT_TIMESTAMP,
  UNIQUE KEY uk_rma_prodotto (id_rma, id_p),
  INDEX idx_data_rma (data_rma),
  INDEX idx_id_p (id_p),
  INDEX idx_id_ordine (id_ordine),
  INDEX idx_tipo_rma (tipo_rma)
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4;
```

### Significato delle colonne principali

| Colonna | Cosa rappresenta |
|---|---|
| `data_rma` | Data di creazione dell'RMA (campo discriminante per filtrare per periodo) |
| `tipo_rma` | "Danni", "Difettosi", "Giacenza", "Recesso", "Mancato Ritiro", ecc. |
| `valore_rimborso` | Importo rimborsato al cliente |
| `valore_claims` | Importo recuperato da claims al corriere |
| `costo_perdita` | Costo lordo della perdita (proporzionale al rimborso sul costo di acquisto). NON include la spedizione di rientro |
| `valore_recuperato` | Recupero stimato via rivendita / allocazione / SEDE |
| `perdita_netta` | `costo_perdita + costo_sped_rientro - valore_claims - valore_recuperato` |
| `costo_sped_rientro` | Costo di spedizione di rientro (= `spese_sped_corriere`, fallback a `spese_sped_grat + spese_sped_cliente` se 0). Valorizzato una sola volta per `id_rma`, solo per Recesso/Giacenza/Danni/Prodotto non conforme con `pixmania != 52` (Stockly escluso) e solo se almeno una riga dell'RMA ha ingresso. Sommato direttamente in `perdita_netta` |
| `nota_recovery` | Spiegazione testuale della logica di recovery applicata |
| `ha_ingresso` | 1 se il prodotto è rientrato in magazzino |
| `ha_ndc` | 1 se per Giacenza/Recesso esiste una NDC associata |
| `data_aggiornamento` | Quando lo script ha (re)scritto la riga |

### Comportamento del batch

A ogni run:
1. Calcola la finestra `PERIODO=180gg` di RMA.
2. Per ogni `(id_rma, id_p)`:
   - Se nuovo → INSERT.
   - Se esistente → UPDATE con i valori ricalcolati (utile quando uno
     stato cambia: es. un'allocazione tardiva, un claim risolto, ecc.).
3. Gli RMA **fuori finestra** (più vecchi di 180gg) **non vengono toccati**:
   restano con i valori dell'ultimo update utile.

> Conseguenza pratica: uno stesso RMA può essere aggiornato più volte
> finché rientra nella finestra dei 180gg, poi si "congela".

---

## Workflow di popolamento

```
┌────────────────────────────────────┐
│ export_dettaglio_resi_completo.py  │  → UPSERT in resi_impatto_economico
└────────────────┬───────────────────┘
                 │ (deve girare PRIMA del PL,
                 │  così pl_prodotti legge dati freschi)
                 ▼
┌────────────────────────────────────┐
│       export_pl_prodotti.py        │  → 6 snapshot in pl_prodotti
└────────────────────────────────────┘
                 │
                 ▼
        ┌─────────────────┐
        │ Report / API    │
        └─────────────────┘
```

### Cadenza consigliata

| Job | Quando | Durata stimata |
|---|---|---|
| `export_dettaglio_resi_completo.py` | 1× al giorno (notte) | ~5–15 min |
| `export_pl_prodotti.py` | 1× al giorno (dopo i resi) | ~6× il tempo del run a 180gg (loop su 6 periodi) |

### Backfill iniziale (una tantum)

1. **Resi**: lancia lo script con `PERIODO = 720` per popolare due anni di
   storico, poi rimetti `PERIODO = 180`.
2. **PL**: nessun backfill possibile (snapshot prospettici), parti da oggi.

---

## Query tipo per il report

### 1. P&L di tutti i prodotti — periodo X — alla data Y

```sql
SELECT *
FROM yeppon_stats.pl_prodotti
WHERE data_snapshot = :data_snap        -- es. '2026-04-28'
  AND periodo_giorni = :periodo         -- es. 90
ORDER BY tot_fatturato DESC;
```

### 2. Evoluzione di un prodotto nel tempo (a parità di periodo)

```sql
SELECT data_snapshot, tot_fatturato, margine_effettivo, perc_margine_eff
FROM yeppon_stats.pl_prodotti
WHERE id_p = :id_p
  AND periodo_giorni = 90
  AND data_snapshot BETWEEN :from AND :to
ORDER BY data_snapshot;
```

### 3. Dropdown: snapshot disponibili

```sql
SELECT DISTINCT data_snapshot, periodo_giorni
FROM yeppon_stats.pl_prodotti
ORDER BY data_snapshot DESC, periodo_giorni;
```

### 3-bis. Dettaglio corrieri di un prodotto

```sql
SELECT corriere, num_ordini, tot_pezzi,
       tot_sped_fatt, tot_sped_costi,
       delta_sped, perc_delta_sped
FROM yeppon_stats.pl_prodotti_corrieri
WHERE data_snapshot = :data_snap
  AND periodo_giorni = :periodo
  AND id_p = :id_p
ORDER BY tot_sped_costi DESC;
```

### 3-ter. Top corrieri per delta negativo (perdita su spedizione)

```sql
SELECT corriere,
       SUM(num_ordini)     AS num_ordini,
       SUM(tot_sped_fatt)  AS incasso,
       SUM(tot_sped_costi) AS costo,
       SUM(delta_sped)     AS delta
FROM yeppon_stats.pl_prodotti_corrieri
WHERE data_snapshot = :data_snap
  AND periodo_giorni = :periodo
GROUP BY corriere
ORDER BY delta ASC;       -- più negativi in cima
```

### 3-quater. Drill-down sugli ordini di un prodotto

```sql
SELECT id_ordine, corriere, quantita,
       sped_fatt, sped_costi, delta_sped
FROM yeppon_stats.pl_prodotti_corrieri_dettaglio
WHERE data_snapshot = :data_snap
  AND periodo_giorni = :periodo
  AND id_p = :id_p
ORDER BY delta_sped ASC;  -- ordini più in perdita in cima
```

### 4. Dettaglio resi in un range di date (intervallo arbitrario)

```sql
SELECT *
FROM yeppon_stats.resi_impatto_economico
WHERE data_rma >= :from_date
  AND data_rma <  :to_date_exclusive    -- intervallo half-open
ORDER BY data_rma DESC;
```

### 5. Aggregato resi per tipologia, range arbitrario

```sql
SELECT tipo_rma,
       COUNT(*)                AS n_righe,
       SUM(quantita_rma)       AS pezzi,
       SUM(costo_perdita)      AS perdita_lorda,
       SUM(valore_claims)      AS claims,
       SUM(valore_recuperato)  AS recupero,
       SUM(perdita_netta)      AS perdita_netta
FROM yeppon_stats.resi_impatto_economico
WHERE data_rma BETWEEN :from AND :to
GROUP BY tipo_rma
ORDER BY perdita_netta DESC;
```

### 6. Resi di uno specifico prodotto in un range

```sql
SELECT id_rma, id_ordine, tipo_rma, data_rma,
       quantita_rma, costo_perdita, valore_claims,
       valore_recuperato, perdita_netta, nota_recovery
FROM yeppon_stats.resi_impatto_economico
WHERE id_p = :id_p
  AND data_rma BETWEEN :from AND :to
ORDER BY data_rma DESC;
```

### ⚠️ Attenzione alle aggregazioni

Le colonne **percentuali** (`perc_margine`, `perc_resi`, `perc_delta_sped`,
`perc_margine_eff`) **non sono sommabili**: vanno **ricalcolate** nel report
dal numeratore e denominatore originali.

Esempio corretto per la `perc_margine` di una marca:
```sql
SELECT marca,
       SUM(tot_margine) / NULLIF(SUM(tot_fatturato), 0) * 100 AS perc_margine
FROM yeppon_stats.pl_prodotti
WHERE data_snapshot = :d AND periodo_giorni = :p
GROUP BY marca;
```

---

## Manutenzione

### Re-run nello stesso giorno

Entrambi gli script sono **idempotenti**: rilanciarli lo stesso giorno
sovrascrive i dati di quel giorno senza duplicare nulla.

- `pl_prodotti`: `DELETE` selettivo + INSERT per `(CURDATE(), periodo)`.
- `pl_prodotti_corrieri`: `DELETE` selettivo + INSERT per `(CURDATE(), periodo)`.
- `pl_prodotti_corrieri_dettaglio`: `DELETE` selettivo + INSERT per `(CURDATE(), periodo)`.
- `resi_impatto_economico`: `UPSERT` su `(id_rma, id_p)`.

### Pulizia storico vecchio

Se in futuro vuoi alleggerire `pl_prodotti`, puoi cancellare snapshot vecchi
mantenendo una granularità ridotta (es. solo il primo del mese):

```sql
-- Esempio: tieni solo gli snapshot mensili oltre 90 giorni fa
DELETE FROM yeppon_stats.pl_prodotti
WHERE data_snapshot < CURRENT_DATE - INTERVAL 90 DAY
  AND DAY(data_snapshot) <> 1;
```

Per `resi_impatto_economico` di norma **non** si cancella nulla: è la fonte
di verità storica dei resi.

### Verifica integrità

```sql
-- pl_prodotti: ci dovrebbero essere 6 righe per prodotto/snapshot
SELECT data_snapshot, COUNT(DISTINCT periodo_giorni) AS n_periodi
FROM yeppon_stats.pl_prodotti
GROUP BY data_snapshot
ORDER BY data_snapshot DESC
LIMIT 10;

-- pl_prodotti_corrieri: confronto numero corrieri attivi per snapshot/periodo
SELECT data_snapshot, periodo_giorni,
       COUNT(DISTINCT corriere)       AS n_corrieri,
       COUNT(DISTINCT id_p)           AS n_prodotti,
       COUNT(*)                       AS n_righe
FROM yeppon_stats.pl_prodotti_corrieri
GROUP BY data_snapshot, periodo_giorni
ORDER BY data_snapshot DESC, periodo_giorni;

-- resi_impatto_economico: nessun duplicato per chiave logica
SELECT id_rma, id_p, COUNT(*) c
FROM yeppon_stats.resi_impatto_economico
GROUP BY id_rma, id_p
HAVING c > 1;   -- deve essere vuoto
```

### Cosa fare se la UNIQUE KEY su `resi_impatto_economico` non si crea

L'errore è dovuto a duplicati esistenti. Soluzione più semplice:

```sql
TRUNCATE TABLE yeppon_stats.resi_impatto_economico;
ALTER TABLE yeppon_stats.resi_impatto_economico
  ADD UNIQUE KEY uk_rma_prodotto (id_rma, id_p);
-- poi rilancia export_dettaglio_resi_completo.py con PERIODO=720 per backfill
```

### Migrazione: aggiunta colonna `costo_sped_rientro`

```sql
ALTER TABLE yeppon_stats.resi_impatto_economico
  ADD COLUMN costo_sped_rientro DECIMAL(10,2) NOT NULL DEFAULT 0 AFTER perdita_netta;
```

Dopo l'ALTER, rilanciare `export_dettaglio_resi_completo.py` per popolare la colonna sullo storico esistente (UPSERT aggiorna le righe).
