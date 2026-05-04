# Deploy P&L Prodotti — Passaggi produzione

## File modificati da copiare sul server
```
app/services/pl_prodotti_query.py
app/routers/pl_prodotti.py
app/static/test_pl_prodotti.html
```

---

## 1. Migrazione DB (MySQL produzione)

Eseguire UNA SOLA VOLTA su `yeppon_stats`:

```sql
ALTER TABLE yeppon_stats.pl_prodotti_marketplace
  ADD COLUMN tot_sped_fatt     DOUBLE       NOT NULL DEFAULT 0 AFTER perc_margine,
  ADD COLUMN tot_sped_costi    DOUBLE       NOT NULL DEFAULT 0 AFTER tot_sped_fatt,
  ADD COLUMN delta_sped        DOUBLE       NOT NULL DEFAULT 0 AFTER tot_sped_costi,
  ADD COLUMN perc_delta_sped   DOUBLE       NOT NULL DEFAULT 0 AFTER delta_sped,
  ADD COLUMN qty_resi          INT UNSIGNED NOT NULL DEFAULT 0 AFTER perc_delta_sped,
  ADD COLUMN perc_resi         DOUBLE       NOT NULL DEFAULT 0 AFTER qty_resi,
  ADD COLUMN imp_eco_resi      DOUBLE       NOT NULL DEFAULT 0 AFTER perc_resi,
  ADD COLUMN margine_effettivo DOUBLE       NOT NULL DEFAULT 0 AFTER imp_eco_resi,
  ADD COLUMN perc_margine_eff  DOUBLE       NOT NULL DEFAULT 0 AFTER margine_effettivo;
```

> Se la tabella non esiste ancora in produzione, viene creata automaticamente
> dal primo run di `export_pl_prodotti.py` con il nuovo schema.

---

## 2. Deploy file

Copiare i 3 file sul server (o usare `deploy_to_prod.bat` dopo aver impostato `PROD_PATH`).

---

## 3. Riavvio FastAPI

```bash
# Linux/systemd
sudo systemctl restart pl-prodotti

# oppure manualmente
pkill -f "uvicorn app.main"
uvicorn app.main:app --host 0.0.0.0 --port 8000
```

---

## 4. Popolare pl_prodotti_marketplace (nuove colonne)

Eseguire `export_pl_prodotti.py` sul server batch. Alla prossima esecuzione
scriverà automaticamente anche le nuove colonne (`delta_sped`, `qty_resi`, ecc.).

> Fino a quel momento la sezione "Distribuzione Vendite per Marketplace"
> mostrerà i dati base (fatturato, margine) ma con Δ Sped. e Resi a zero.

---

## Cosa è cambiato (riepilogo)

| Cosa | Dettaglio |
|------|-----------|
| Periodo default | 180 → **90 giorni** |
| Nuovi endpoint | `/corrieri_summary`, `/marketplace_breakdown` |
| Sezione marketplace | Tabella con 13 colonne + grafico pezzi per marketplace |
| Drilldown Δ Sped. | Grafico a barre verticale affianco alla tabella corrieri |
| Drilldown Resi | Grafico a torta per tipologia affianco alla tabella |
| Colonne ordinabili | Click sulle intestazioni della tabella principale |
| Grafici sidebar | Nascosti quando si filtra un singolo prodotto |
