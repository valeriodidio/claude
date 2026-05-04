-- Schema MySQL per dev locale del report P&L Prodotti.
-- Riproduce le QUATTRO tabelle pre-calcolate dagli script batch:
--   yeppon_stats.pl_prodotti                       (export_pl_prodotti.py)
--   yeppon_stats.pl_prodotti_corrieri              (export_pl_prodotti.py)
--   yeppon_stats.pl_prodotti_corrieri_dettaglio    (export_pl_prodotti.py)
--   yeppon_stats.resi_impatto_economico            (export_dettaglio_resi_completo.py)
-- Crea inoltre l'utente read-only `reports_reader`.

CREATE DATABASE IF NOT EXISTS smart2        DEFAULT CHARACTER SET utf8mb4 COLLATE utf8mb4_unicode_ci;
CREATE DATABASE IF NOT EXISTS yeppon_stats  DEFAULT CHARACTER SET utf8mb4 COLLATE utf8mb4_unicode_ci;

USE yeppon_stats;

-- ── Tabella P&L per prodotto (snapshot multi-periodo) ───────────────────────
-- Una riga = P&L di un prodotto calcolato su una finestra (periodo_giorni)
-- a una specifica data di calcolo (data_snapshot).
-- Chiave logica: (data_snapshot, id_p, periodo_giorni) → UNIQUE
CREATE TABLE IF NOT EXISTS pl_prodotti (
    id                          BIGINT UNSIGNED AUTO_INCREMENT PRIMARY KEY,
    id_p                        INT UNSIGNED    NOT NULL,
    codice                      VARCHAR(50)     NOT NULL DEFAULT '',
    nome                        VARCHAR(255)    NOT NULL DEFAULT '',
    marca                       VARCHAR(100)    NOT NULL DEFAULT '',
    disp_fornitore              INT             NOT NULL DEFAULT 0,
    status_prodotto             TINYINT         NOT NULL DEFAULT 0,
    bloccato                    TINYINT         NOT NULL DEFAULT 0,
    bloccato_fino               DATE                     DEFAULT NULL,
    num_ordini                  INT UNSIGNED    NOT NULL DEFAULT 0,
    tot_pezzi                   INT UNSIGNED    NOT NULL DEFAULT 0,
    tot_fatturato               DOUBLE          NOT NULL DEFAULT 0,
    tot_margine                 DOUBLE          NOT NULL DEFAULT 0,
    perc_margine                DOUBLE          NOT NULL DEFAULT 0,
    tot_sped_fatt               DOUBLE          NOT NULL DEFAULT 0,
    tot_sped_costi              DOUBLE          NOT NULL DEFAULT 0,
    delta_sped                  DOUBLE          NOT NULL DEFAULT 0,
    perc_delta_sped             DOUBLE          NOT NULL DEFAULT 0,
    qty_resi                    INT UNSIGNED    NOT NULL DEFAULT 0,
    perc_resi                   DOUBLE          NOT NULL DEFAULT 0,
    imp_eco_resi                DOUBLE          NOT NULL DEFAULT 0,
    imp_Danni                   DOUBLE          NOT NULL DEFAULT 0,
    imp_Danno_rientrato         DOUBLE          NOT NULL DEFAULT 0,
    imp_Difettosi               DOUBLE          NOT NULL DEFAULT 0,
    imp_Giacenza                DOUBLE          NOT NULL DEFAULT 0,
    imp_Mancato_Ritiro          DOUBLE          NOT NULL DEFAULT 0,
    imp_Prodotto_non_conforme   DOUBLE          NOT NULL DEFAULT 0,
    imp_Recesso                 DOUBLE          NOT NULL DEFAULT 0,
    imp_Reclamo_contestazioni   DOUBLE          NOT NULL DEFAULT 0,
    imp_Smarrimenti             DOUBLE          NOT NULL DEFAULT 0,
    margine_effettivo           DOUBLE          NOT NULL DEFAULT 0,
    perc_margine_eff            DOUBLE          NOT NULL DEFAULT 0,
    categoria                   VARCHAR(100)    NOT NULL DEFAULT '',
    categoria2                  VARCHAR(100)    NOT NULL DEFAULT '',
    categoria3                  VARCHAR(100)    NOT NULL DEFAULT '',
    sender                      VARCHAR(100)    NOT NULL DEFAULT '',
    fornitore                   VARCHAR(100)    NOT NULL DEFAULT '',
    periodo_giorni              INT UNSIGNED    NOT NULL DEFAULT 180,
    data_snapshot               DATE            NOT NULL DEFAULT (CURRENT_DATE),
    data_calcolo                DATETIME        NOT NULL DEFAULT CURRENT_TIMESTAMP,

    UNIQUE KEY uk_snap_prod_periodo (data_snapshot, id_p, periodo_giorni),
    INDEX idx_snapshot_periodo  (data_snapshot, periodo_giorni),
    INDEX idx_periodo_snapshot  (periodo_giorni, data_snapshot),
    INDEX idx_id_p              (id_p),
    INDEX idx_codice            (codice),
    INDEX idx_marca             (marca),
    INDEX idx_status            (status_prodotto),
    INDEX idx_bloccato          (bloccato),
    INDEX idx_categoria         (categoria),
    INDEX idx_sender            (sender),
    INDEX idx_fornitore         (fornitore)
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4;


-- ── Dettaglio spedizioni per (prodotto, corriere) ───────────────────────────
-- Una riga = delta spedizioni di un prodotto su uno specifico corriere
-- per una finestra (periodo_giorni) a una specifica data_snapshot.
-- Solo ordini con fattura corriere reale (no fallback tariffario).
-- Chiave logica: (data_snapshot, periodo_giorni, id_p, corriere) → UNIQUE
CREATE TABLE IF NOT EXISTS pl_prodotti_corrieri (
    id              BIGINT UNSIGNED AUTO_INCREMENT PRIMARY KEY,
    data_snapshot   DATE            NOT NULL,
    periodo_giorni  INT UNSIGNED    NOT NULL,
    id_p            INT UNSIGNED    NOT NULL,
    corriere        VARCHAR(50)     NOT NULL DEFAULT '',
    num_ordini      INT UNSIGNED    NOT NULL DEFAULT 0,
    tot_pezzi       INT UNSIGNED    NOT NULL DEFAULT 0,
    tot_sped_fatt   DOUBLE          NOT NULL DEFAULT 0,
    tot_sped_costi  DOUBLE          NOT NULL DEFAULT 0,
    delta_sped      DOUBLE          NOT NULL DEFAULT 0,
    perc_delta_sped DOUBLE          NOT NULL DEFAULT 0,
    data_calcolo    DATETIME        NOT NULL DEFAULT CURRENT_TIMESTAMP,

    UNIQUE KEY uk_snap_prod_per_corr (data_snapshot, periodo_giorni, id_p, corriere),
    INDEX idx_id_p      (id_p),
    INDEX idx_corriere  (corriere),
    INDEX idx_snap_per  (data_snapshot, periodo_giorni)
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4;


-- ── Dettaglio spedizioni per (prodotto, ordine) — drill-down ───────────────
-- Una riga = un ordine specifico con la quota spedizione attribuita al prodotto.
-- Solo ordini con fattura corriere reale (no fallback tariffario).
-- Chiave logica: (data_snapshot, periodo_giorni, id_p, id_ordine) → UNIQUE
CREATE TABLE IF NOT EXISTS pl_prodotti_corrieri_dettaglio (
    id              BIGINT UNSIGNED AUTO_INCREMENT PRIMARY KEY,
    data_snapshot   DATE            NOT NULL,
    periodo_giorni  INT UNSIGNED    NOT NULL,
    id_p            INT UNSIGNED    NOT NULL,
    id_ordine       BIGINT UNSIGNED NOT NULL,
    corriere        VARCHAR(50)     NOT NULL DEFAULT '',
    quantita        INT UNSIGNED    NOT NULL DEFAULT 0,
    sped_fatt       DOUBLE          NOT NULL DEFAULT 0,
    sped_costi      DOUBLE          NOT NULL DEFAULT 0,
    delta_sped      DOUBLE          NOT NULL DEFAULT 0,
    data_calcolo    DATETIME        NOT NULL DEFAULT CURRENT_TIMESTAMP,

    UNIQUE KEY uk_snap_per_prod_ord (data_snapshot, periodo_giorni, id_p, id_ordine),
    INDEX idx_id_p      (id_p),
    INDEX idx_id_ordine (id_ordine),
    INDEX idx_corriere  (corriere),
    INDEX idx_snap_per  (data_snapshot, periodo_giorni)
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4;


-- ── Impatto economico resi (storico cumulativo, eventi atomici) ─────────────
-- Una riga = un prodotto in un RMA (singolo evento di reso).
-- NON è uno snapshot: è una tabella di eventi storici che cresce nel tempo.
-- Chiave logica: (id_rma, id_p) → UNIQUE
CREATE TABLE IF NOT EXISTS resi_impatto_economico (
    id_rma                  BIGINT          NOT NULL,
    id_ordine               BIGINT          NOT NULL,
    id_p                    INT UNSIGNED    NOT NULL,
    tipo_rma                VARCHAR(50)     NOT NULL DEFAULT '',
    marketplace             VARCHAR(50)     NOT NULL DEFAULT '',
    quantita_rma            INT             NOT NULL DEFAULT 0,
    costo_unitario          DOUBLE          NOT NULL DEFAULT 0,
    prezzo_vendita_unit     DOUBLE          NOT NULL DEFAULT 0,
    valore_rimborso         DOUBLE          NOT NULL DEFAULT 0,
    perc_rimborso           DOUBLE          NOT NULL DEFAULT 0,
    costo_perdita           DOUBLE          NOT NULL DEFAULT 0,
    valore_claims           DOUBLE          NOT NULL DEFAULT 0,
    valore_recuperato       DOUBLE          NOT NULL DEFAULT 0,
    perdita_netta           DOUBLE          NOT NULL DEFAULT 0,
    costo_sped_rientro      DOUBLE          NOT NULL DEFAULT 0,
    ha_ingresso             TINYINT         NOT NULL DEFAULT 0,
    ha_ndc                  TINYINT         NOT NULL DEFAULT 0,
    nota_recovery           VARCHAR(500)    NOT NULL DEFAULT '',
    data_rma                DATETIME                 DEFAULT NULL,
    data_calcolo            DATETIME        NOT NULL DEFAULT CURRENT_TIMESTAMP,
    data_aggiornamento      DATETIME        NOT NULL DEFAULT CURRENT_TIMESTAMP
                            ON UPDATE CURRENT_TIMESTAMP,

    UNIQUE KEY uk_rma_prodotto  (id_rma, id_p),
    INDEX idx_data_rma          (data_rma),
    INDEX idx_id_p              (id_p),
    INDEX idx_id_ordine         (id_ordine),
    INDEX idx_tipo_rma          (tipo_rma),
    INDEX idx_marketplace       (marketplace)
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4;


-- ── Dettaglio vendite per (prodotto, marketplace) ───────────────────────────
-- Una riga = vendite di un prodotto su un marketplace specifico
-- per una finestra (periodo_giorni) a una specifica data_snapshot.
-- Chiave logica: (data_snapshot, periodo_giorni, id_p, marketplace) → UNIQUE
CREATE TABLE IF NOT EXISTS pl_prodotti_marketplace (
    id                BIGINT UNSIGNED AUTO_INCREMENT PRIMARY KEY,
    data_snapshot     DATE            NOT NULL,
    periodo_giorni    INT UNSIGNED    NOT NULL,
    id_p              INT UNSIGNED    NOT NULL,
    marketplace       VARCHAR(50)     NOT NULL DEFAULT '',
    num_ordini        INT UNSIGNED    NOT NULL DEFAULT 0,
    tot_pezzi         INT UNSIGNED    NOT NULL DEFAULT 0,
    tot_fatturato     DOUBLE          NOT NULL DEFAULT 0,
    tot_margine       DOUBLE          NOT NULL DEFAULT 0,
    perc_margine      DOUBLE          NOT NULL DEFAULT 0,
    tot_sped_fatt     DOUBLE          NOT NULL DEFAULT 0,
    tot_sped_costi    DOUBLE          NOT NULL DEFAULT 0,
    delta_sped        DOUBLE          NOT NULL DEFAULT 0,
    perc_delta_sped   DOUBLE          NOT NULL DEFAULT 0,
    qty_resi          INT UNSIGNED    NOT NULL DEFAULT 0,
    perc_resi         DOUBLE          NOT NULL DEFAULT 0,
    imp_eco_resi      DOUBLE          NOT NULL DEFAULT 0,
    margine_effettivo DOUBLE          NOT NULL DEFAULT 0,
    perc_margine_eff  DOUBLE          NOT NULL DEFAULT 0,
    data_calcolo      DATETIME        NOT NULL DEFAULT CURRENT_TIMESTAMP,

    UNIQUE KEY uk_snap_per_prod_mktp (data_snapshot, periodo_giorni, id_p, marketplace),
    INDEX idx_id_p          (id_p),
    INDEX idx_marketplace   (marketplace),
    INDEX idx_snap_per      (data_snapshot, periodo_giorni)
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4;


-- ── Utente read-only usato dall'app ─────────────────────────────────────────
CREATE USER IF NOT EXISTS 'reports_reader'@'%' IDENTIFIED BY 'dev_password_change_me';
GRANT SELECT ON yeppon_stats.pl_prodotti                       TO 'reports_reader'@'%';
GRANT SELECT ON yeppon_stats.pl_prodotti_corrieri              TO 'reports_reader'@'%';
GRANT SELECT ON yeppon_stats.pl_prodotti_corrieri_dettaglio    TO 'reports_reader'@'%';
GRANT SELECT ON yeppon_stats.resi_impatto_economico            TO 'reports_reader'@'%';
GRANT SELECT ON yeppon_stats.pl_prodotti_marketplace           TO 'reports_reader'@'%';
FLUSH PRIVILEGES;
