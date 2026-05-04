"""Logica SQL del report P&L Prodotti.

Le tabelle sorgenti sono pre-calcolate dagli script batch:
  - yeppon_stats.pl_prodotti              (export_pl_prodotti.py)
  - yeppon_stats.resi_impatto_economico   (export_dettaglio_resi_completo.py)

Struttura snapshot di pl_prodotti:
  - Chiave logica: (data_snapshot, id_p, periodo_giorni)
  - periodo_giorni ∈ {30, 60, 90, 120, 180, 360}
  - Per ogni query è SEMPRE necessario specificare data_snapshot + periodo_giorni.
    Se data_snapshot è None, viene risolta automaticamente all'ultimo disponibile.

Qui dentro NON c'è logica HTTP: solo costruzione query e formattazione record.
"""
import logging
from dataclasses import dataclass, field
from datetime import date, datetime
from typing import List, Optional

import pandas as pd
from sqlalchemy import text
from sqlalchemy.engine import Engine

log = logging.getLogger(__name__)

# ── Tabelle ─────────────────────────────────────────────────────────────────
TABLE_PL = "yeppon_stats.pl_prodotti"
TABLE_RESI = "yeppon_stats.resi_impatto_economico"
TABLE_CORRIERI = "yeppon_stats.pl_prodotti_corrieri"
TABLE_CORRIERI_DET = "yeppon_stats.pl_prodotti_corrieri_dettaglio"
TABLE_MARKETPLACE = "yeppon_stats.pl_prodotti_marketplace"

# Periodi supportati
PERIODI_VALIDI = {30, 60, 90, 120, 180, 360}

# Tipi RMA "noti" (mantieniamo l'ordine usato nel report Excel)
TIPI_RMA = [
    "Danni",
    "Danno rientrato",
    "Difettosi",
    "Giacenza",
    "Mancato Ritiro",
    "Prodotto non conforme",
    "Recesso",
    "Reclamo e contestazioni",
    "Smarrimenti",
]

# Colonne "imp_*" del PL (una per tipo RMA)
IMP_COLS = {
    "Danni": "imp_Danni",
    "Danno rientrato": "imp_Danno_rientrato",
    "Difettosi": "imp_Difettosi",
    "Giacenza": "imp_Giacenza",
    "Mancato Ritiro": "imp_Mancato_Ritiro",
    "Prodotto non conforme": "imp_Prodotto_non_conforme",
    "Recesso": "imp_Recesso",
    "Reclamo e contestazioni": "imp_Reclamo_contestazioni",
    "Smarrimenti": "imp_Smarrimenti",
}

# Whitelist dei campi su cui consentire ORDER BY (anti SQL injection)
SORTABLE_FIELDS = {
    "id_p", "codice", "nome", "marca",
    "num_ordini", "tot_pezzi",
    "tot_fatturato", "tot_margine", "perc_margine",
    "tot_sped_fatt", "tot_sped_costi", "delta_sped", "perc_delta_sped",
    "qty_resi", "perc_resi", "imp_eco_resi",
    "margine_effettivo", "perc_margine_eff",
    "data_snapshot",
}


# ── Filtri ──────────────────────────────────────────────────────────────────
@dataclass
class Filters:
    # ── Dimensioni snapshot (obbligatorie per pl_prodotti) ───────────────────
    # data_snapshot=None → usa l'ultimo snapshot disponibile per il periodo
    data_snapshot: Optional[date] = None
    periodo_giorni: int = 90          # 30 / 60 / 90 / 120 / 180 / 360

    # ── Filtri su prodotto ───────────────────────────────────────────────────
    marca: List[str] = field(default_factory=list)
    status_prodotto: Optional[int] = None    # 1=attivo / 0=disattivo
    bloccato: Optional[int] = None           # 1=sì / 0=no
    margine_negativo: bool = False           # solo prodotti con margine_eff < 0
    solo_con_resi: bool = False              # solo prodotti con qty_resi > 0
    search: Optional[str] = None            # cerca su codice / nome / id_p
    min_fatturato: Optional[float] = None
    min_ordini: Optional[int] = None

    # ── Filtri per resi (usati in get_resi e get_resi_global) ───────────────
    marketplace: List[str] = field(default_factory=list)
    tipo_rma: List[str] = field(default_factory=list)
    resi_dal: Optional[date] = None
    resi_al: Optional[date] = None

    # ── Paginazione / ordinamento ────────────────────────────────────────────
    sort_by: str = "tot_fatturato"
    sort_dir: str = "desc"
    page: int = 1
    page_size: int = 100



# ── Helpers serializzazione ──────────────────────────────────────────────────
def _row_to_dict(row) -> dict:
    """Converte un row SQLAlchemy Mapping in dizionario JSON-serializzabile."""
    if row is None:
        return {}
    out = {}
    for k, v in dict(row).items():
        if isinstance(v, datetime):
            out[k] = v.isoformat()
        elif isinstance(v, date):
            out[k] = v.isoformat()
        elif hasattr(v, '__float__'):          # Decimal, numpy float, ecc.
            out[k] = float(v)
        else:
            out[k] = v
    return out


def _filters_to_dict(f: Filters, snap: Optional[date] = None) -> dict:
    """Serializza i filtri attivi nella risposta API."""
    effective_snap = snap or f.data_snapshot
    return {
        "data_snapshot":    effective_snap.isoformat() if effective_snap else None,
        "periodo_giorni":   f.periodo_giorni,
        "marca":            f.marca,
        "status_prodotto":  f.status_prodotto,
        "bloccato":         f.bloccato,
        "margine_negativo": f.margine_negativo,
        "solo_con_resi":    f.solo_con_resi,
        "search":           f.search,
        "min_fatturato":    f.min_fatturato,
        "min_ordini":       f.min_ordini,
        "marketplace":      f.marketplace,
        "tipo_rma":         f.tipo_rma,
        "resi_dal":         f.resi_dal.isoformat() if f.resi_dal else None,
        "resi_al":          f.resi_al.isoformat()  if f.resi_al  else None,
        "sort_by":          f.sort_by,
        "sort_dir":         f.sort_dir,
    }


# ── Risoluzione snapshot ─────────────────────────────────────────────────────
def resolve_snapshot(engine: Engine, f: Filters, table: str = TABLE_PL) -> date:
    """Restituisce data_snapshot effettiva: quella richiesta oppure l'ultima disponibile.

    `table` consente di risolvere lo snapshot dalla tabella effettivamente
    interrogata: pl_prodotti, pl_prodotti_corrieri o pl_prodotti_corrieri_dettaglio.
    Le tre tabelle possono avere snapshot diversi (un batch può aggiornarne una
    e non l'altra), quindi ognuna va risolta indipendentemente — altrimenti
    si rischia di filtrare per una data che esiste solo in pl_prodotti e
    ottenere 0 righe sulle altre.
    """
    if f.data_snapshot is not None:
        return f.data_snapshot
    sql = f"SELECT MAX(data_snapshot) FROM {table} WHERE periodo_giorni = :p"
    try:
        with engine.connect() as conn:
            result = conn.execute(text(sql), {"p": f.periodo_giorni}).scalar()
    except Exception:
        log.exception("resolve_snapshot fallita su %s (periodo=%s)", table, f.periodo_giorni)
        result = None
    if result is None:
        # Tabella vuota o non esistente: usa oggi come fallback
        return date.today()
    return result


# ── Costruzione WHERE per pl_prodotti ────────────────────────────────────────
def _where(f: Filters, resolved_snapshot: date) -> tuple[str, dict]:
    """Costruisce WHERE su pl_prodotti. data_snapshot e periodo_giorni sono SEMPRE presenti."""
    clauses: list[str] = [
        "data_snapshot = :data_snapshot",
        "periodo_giorni = :periodo_giorni",
    ]
    params: dict = {
        "data_snapshot": resolved_snapshot,
        "periodo_giorni": int(f.periodo_giorni),
    }

    if f.marca:
        clauses.append("marca IN :marca")
        params["marca"] = tuple(f.marca)

    if f.status_prodotto is not None:
        clauses.append("status_prodotto = :status_prodotto")
        params["status_prodotto"] = int(f.status_prodotto)

    if f.bloccato is not None:
        clauses.append("bloccato = :bloccato")
        params["bloccato"] = int(f.bloccato)

    if f.margine_negativo:
        clauses.append("margine_effettivo < 0")

    if f.solo_con_resi:
        clauses.append("qty_resi > 0")

    if f.search:
        clauses.append(
            "(codice LIKE :search OR nome LIKE :search OR CAST(id_p AS CHAR) = :search_exact)"
        )
        params["search"] = f"%{f.search}%"
        params["search_exact"] = f.search

    if f.min_fatturato is not None:
        clauses.append("tot_fatturato >= :min_fatturato")
        params["min_fatturato"] = float(f.min_fatturato)

    if f.min_ordini is not None:
        clauses.append("num_ordini >= :min_ordini")
        params["min_ordini"] = int(f.min_ordini)

    where = " WHERE " + " AND ".join(clauses)
    return where, params


# ── Costruzione WHERE per resi_impatto_economico ─────────────────────────────
def _where_resi(f: Filters, extra_clauses: list = None, extra_params: dict = None) -> tuple[str, dict]:
    """Costruisce WHERE su resi_impatto_economico."""
    clauses: list[str] = list(extra_clauses or [])
    params: dict = dict(extra_params or {})

    if f.marketplace:
        clauses.append("marketplace IN :marketplace")
        params["marketplace"] = tuple(f.marketplace)

    if f.tipo_rma:
        clauses.append("tipo_rma IN :tipo_rma")
        params["tipo_rma"] = tuple(f.tipo_rma)

    if f.resi_dal:
        clauses.append("data_rma >= :resi_dal")
        params["resi_dal"] = f.resi_dal

    if f.resi_al:
        clauses.append("data_rma < :resi_al_excl")
        # intervallo half-open: al giorno incluso → aggiungiamo 1 giorno
        from datetime import timedelta
        params["resi_al_excl"] = f.resi_al + timedelta(days=1)

    where = (" WHERE " + " AND ".join(clauses)) if clauses else ""
    return where, params


def _order_by(f: Filters) -> str:
    field_ = f.sort_by if f.sort_by in SORTABLE_FIELDS else "tot_fatturato"
    direction = "ASC" if f.sort_dir.lower() == "asc" else "DESC"
    return f" ORDER BY {field_} {direction}, id_p ASC"


# ── Filtri/options per la UI ─────────────────────────────────────────────────
def get_filter_options(engine: Engine) -> dict:
    """Restituisce gli elenchi distinti per popolare i filtri dell'admin."""

    # Snapshot disponibili (ultimi 60 per non appesantire, ordinati dal più recente)
    sql_snapshots = f"""
        SELECT DISTINCT data_snapshot
        FROM {TABLE_PL}
        ORDER BY data_snapshot DESC
        LIMIT 60
    """
    # Periodi disponibili
    sql_periodi = f"""
        SELECT DISTINCT periodo_giorni
        FROM {TABLE_PL}
        ORDER BY periodo_giorni
    """
    # Marche (dall'ultimo snapshot disponibile)
    sql_marche = f"""
        SELECT DISTINCT marca
        FROM {TABLE_PL}
        WHERE data_snapshot = (SELECT MAX(data_snapshot) FROM {TABLE_PL})
          AND marca <> ''
        ORDER BY marca
    """
    sql_mktp = f"SELECT DISTINCT marketplace FROM {TABLE_RESI} WHERE marketplace <> '' ORDER BY marketplace"

    # Metadati ultimo snapshot
    sql_meta = f"""
        SELECT
            MAX(data_snapshot)  AS ultimo_snapshot,
            MIN(data_snapshot)  AS primo_snapshot,
            COUNT(DISTINCT data_snapshot) AS n_snapshot,
            COUNT(DISTINCT periodo_giorni) AS n_periodi
        FROM {TABLE_PL}
    """

    with engine.connect() as conn:
        snapshots = [r[0].isoformat() if hasattr(r[0], 'isoformat') else str(r[0])
                     for r in conn.execute(text(sql_snapshots)).fetchall() if r[0]]
        periodi = [int(r[0]) for r in conn.execute(text(sql_periodi)).fetchall() if r[0]]
        marche = [r[0] for r in conn.execute(text(sql_marche)).fetchall() if r[0]]
        try:
            mktp = [r[0] for r in conn.execute(text(sql_mktp)).fetchall() if r[0]]
        except Exception:
            mktp = []
        meta_row = conn.execute(text(sql_meta)).fetchone()

    # Se non ci sono periodi nel DB usa i valori standard
    if not periodi:
        periodi = sorted(PERIODI_VALIDI)

    meta = {}
    if meta_row:
        meta = {
            "ultimo_snapshot": meta_row[0].isoformat() if meta_row[0] else None,
            "primo_snapshot": meta_row[1].isoformat() if meta_row[1] else None,
            "n_snapshot": int(meta_row[2]) if meta_row[2] else 0,
            "n_periodi": int(meta_row[3]) if meta_row[3] else 0,
        }

    return {
        "snapshots": snapshots,
        "periodi_giorni": periodi,
        "marca": marche,
        "marketplace": mktp,
        "tipi_rma": TIPI_RMA,
        "status_prodotto": [
            {"value": 1, "label": "Attivo"},
            {"value": 0, "label": "Disattivo"},
        ],
        "bloccato": [
            {"value": 0, "label": "No"},
            {"value": 1, "label": "Sì"},
        ],
        "sortable_fields": sorted(SORTABLE_FIELDS),
        "meta": meta,
    }


# ── Summary KPI ──────────────────────────────────────────────────────────────
def get_summary(engine: Engine, f: Filters) -> dict:
    """Aggrega KPI complessivi sui prodotti che rispettano i filtri."""
    snap = resolve_snapshot(engine, f)
    where, params = _where(f, snap)
    imp_sum = ", ".join(f"SUM({c}) AS {c}" for c in IMP_COLS.values())
    sql = f"""
        SELECT
            COUNT(*)                          AS n_prodotti,
            SUM(num_ordini)                   AS num_ordini,
            SUM(tot_pezzi)                    AS tot_pezzi,
            SUM(tot_fatturato)                AS tot_fatturato,
            SUM(tot_margine)                  AS tot_margine,
            SUM(tot_sped_fatt)                AS tot_sped_fatt,
            SUM(tot_sped_costi)               AS tot_sped_costi,
            SUM(delta_sped)                   AS delta_sped,
            SUM(qty_resi)                     AS qty_resi,
            SUM(imp_eco_resi)                 AS imp_eco_resi,
            SUM(margine_effettivo)            AS margine_effettivo,
            {imp_sum}
        FROM {TABLE_PL}
        {where}
    """
    with engine.connect() as conn:
        row = conn.execute(text(sql), params).mappings().fetchone()

    out = dict(row) if row else {}
    for k, v in list(out.items()):
        if v is None:
            out[k] = 0
        elif isinstance(v, (int, float)):
            out[k] = float(v)

    fatt = out.get("tot_fatturato", 0) or 0
    pezzi = out.get("tot_pezzi", 0) or 0
    sped_fatt = out.get("tot_sped_fatt", 0) or 0
    out["perc_margine"] = round(out["tot_margine"] / fatt * 100, 2) if fatt else 0
    out["perc_margine_eff"] = round(out["margine_effettivo"] / fatt * 100, 2) if fatt else 0
    out["perc_delta_sped"] = round(out["delta_sped"] / sped_fatt * 100, 2) if sped_fatt else 0
    out["perc_resi"] = round(out["qty_resi"] / pezzi * 100, 2) if pezzi else 0

    return {
        "kpi": out,
        "data_snapshot": snap.isoformat(),
        "periodo_giorni": f.periodo_giorni,
        "filters": _filters_to_dict(f, snap),
    }


# ── Lista paginata ───────────────────────────────────────────────────────────
def get_list(engine: Engine, f: Filters) -> dict:
    snap = resolve_snapshot(engine, f)
    where, params = _where(f, snap)
    order = _order_by(f)

    sql_count = f"SELECT COUNT(*) FROM {TABLE_PL}{where}"

    page = max(1, int(f.page))
    page_size = max(1, min(int(f.page_size), 1000))
    offset = (page - 1) * page_size

    sql_data = f"""
        SELECT
            id_p, codice, nome, marca, disp_fornitore,
            status_prodotto, bloccato, bloccato_fino,
            num_ordini, tot_pezzi,
            tot_fatturato, tot_margine, perc_margine,
            tot_sped_fatt, tot_sped_costi, delta_sped, perc_delta_sped,
            qty_resi, perc_resi, imp_eco_resi,
            imp_Danni, imp_Danno_rientrato, imp_Difettosi, imp_Giacenza,
            imp_Mancato_Ritiro, imp_Prodotto_non_conforme, imp_Recesso,
            imp_Reclamo_contestazioni, imp_Smarrimenti,
            margine_effettivo, perc_margine_eff,
            periodo_giorni, data_snapshot, data_calcolo
        FROM {TABLE_PL}
        {where}
        {order}
        LIMIT :limit OFFSET :offset
    """
    p_data = dict(params, limit=page_size, offset=offset)
    with engine.connect() as conn:
        total = conn.execute(text(sql_count), params).scalar() or 0
        rows = conn.execute(text(sql_data), p_data).mappings().fetchall()

    return {
        "page": page,
        "page_size": page_size,
        "total": int(total),
        "n_pages": (int(total) + page_size - 1) // page_size,
        "data_snapshot": snap.isoformat(),
        "periodo_giorni": f.periodo_giorni,
        "rows": [_row_to_dict(r) for r in rows],
        "filters": _filters_to_dict(f, snap),
    }


# ── Detail singolo prodotto ──────────────────────────────────────────────────
def get_detail(engine: Engine, id_p: int, f: Filters) -> Optional[dict]:
    snap = resolve_snapshot(engine, f)
    sql = f"""
        SELECT * FROM {TABLE_PL}
        WHERE id_p = :id_p
          AND data_snapshot = :data_snapshot
          AND periodo_giorni = :periodo_giorni
        LIMIT 1
    """
    with engine.connect() as conn:
        row = conn.execute(text(sql), {
            "id_p": int(id_p),
            "data_snapshot": snap,
            "periodo_giorni": f.periodo_giorni,
        }).mappings().fetchone()
    if not row:
        return None

    detail = _row_to_dict(row)
    detail["resi_per_tipo"] = [
        {
            "tipo_rma": tipo,
            "perdita_netta": float(detail.get(IMP_COLS[tipo], 0) or 0),
        }
        for tipo in TIPI_RMA
    ]

    # Evoluzione storica per questo prodotto (tutti gli snapshot, stesso periodo)
    sql_evo = f"""
        SELECT data_snapshot, tot_fatturato, tot_margine, margine_effettivo,
               perc_margine_eff, qty_resi, imp_eco_resi
        FROM {TABLE_PL}
        WHERE id_p = :id_p AND periodo_giorni = :periodo_giorni
        ORDER BY data_snapshot DESC
        LIMIT 90
    """
    with engine.connect() as conn:
        evo = conn.execute(text(sql_evo), {
            "id_p": int(id_p),
            "periodo_giorni": f.periodo_giorni,
        }).mappings().fetchall()
    detail["evoluzione"] = [_row_to_dict(r) for r in evo]

    return detail


# ── Resi dettaglio per prodotto ──────────────────────────────────────────────
def get_resi(engine: Engine, id_p: int, f: Filters) -> dict:
    """Ritorna l'elenco completo dei resi per un singolo prodotto."""
    where, params = _where_resi(
        f,
        extra_clauses=["id_p = :id_p"],
        extra_params={"id_p": int(id_p)},
    )

    sql_rows = f"""
        SELECT
            id_rma, id_ordine, tipo_rma, marketplace,
            quantita_rma, costo_unitario, prezzo_vendita_unit,
            valore_rimborso, perc_rimborso, costo_perdita,
            COALESCE(costo_sped_rientro, 0) AS costo_sped_rientro,
            valore_claims, valore_recuperato, perdita_netta,
            ha_ingresso, ha_ndc, nota_recovery, data_rma
        FROM {TABLE_RESI}
        {where}
        ORDER BY data_rma DESC, id_rma DESC
    """
    sql_agg = f"""
        SELECT
            tipo_rma,
            COUNT(*)                  AS n_resi,
            SUM(quantita_rma)         AS qty,
            SUM(costo_perdita)        AS costo_perdita,
            SUM(COALESCE(costo_sped_rientro, 0)) AS costo_sped_rientro,
            SUM(valore_claims)        AS claims,
            SUM(valore_recuperato)    AS recuperato,
            SUM(perdita_netta)        AS perdita_netta
        FROM {TABLE_RESI}
        {where}
        GROUP BY tipo_rma
        ORDER BY perdita_netta DESC
    """
    with engine.connect() as conn:
        rows = conn.execute(text(sql_rows), params).mappings().fetchall()
        agg = conn.execute(text(sql_agg), params).mappings().fetchall()

    return {
        "id_p": int(id_p),
        "rows": [_row_to_dict(r) for r in rows],
        "aggregato_per_tipo": [_row_to_dict(r) for r in agg],
        "filters": _filters_to_dict(f),
    }


# ── Corrieri spedizione per prodotto ────────────────────────────────────────
def get_corrieri(engine: Engine, id_p: int, f: Filters) -> dict:
    """
    Dettaglio spedizioni per corriere di un singolo prodotto.
    Legge da pl_prodotti_corrieri (solo ordini con fattura reale; no fallback).
    Lo snapshot è risolto sulla tabella corrieri stessa (può differire da pl_prodotti).
    Restituisce rows vuoto se la tabella non esiste ancora; logga l'errore vero.
    """
    # IMPORTANTE: snapshot risolto da TABLE_CORRIERI, non da TABLE_PL.
    # Le due tabelle possono avere date di calcolo diverse.
    snap = resolve_snapshot(engine, f, TABLE_CORRIERI)
    sql = f"""
        SELECT corriere, num_ordini, tot_pezzi,
               tot_sped_fatt, tot_sped_costi,
               delta_sped, perc_delta_sped
        FROM {TABLE_CORRIERI}
        WHERE id_p         = :id_p
          AND data_snapshot  = :data_snapshot
          AND periodo_giorni = :periodo_giorni
        ORDER BY tot_sped_costi DESC
    """
    rows = []
    try:
        with engine.connect() as conn:
            rows = conn.execute(text(sql), {
                "id_p": int(id_p),
                "data_snapshot": snap,
                "periodo_giorni": f.periodo_giorni,
            }).mappings().fetchall()
    except Exception:
        # Logga il traceback completo invece di nasconderlo dietro un return vuoto
        log.exception(
            "get_corrieri fallita id_p=%s snap=%s periodo=%s",
            id_p, snap, f.periodo_giorni,
        )

    log.info(
        "get_corrieri id_p=%s snap=%s periodo=%s -> %d righe",
        id_p, snap, f.periodo_giorni, len(rows),
    )
    print(
        f"[DEBUG-PL] get_corrieri id_p={id_p} snap={snap} periodo={f.periodo_giorni} -> {len(rows)} righe",
        flush=True,
    )

    return {
        "id_p": int(id_p),
        "data_snapshot": snap.isoformat(),
        "periodo_giorni": f.periodo_giorni,
        "rows": [_row_to_dict(r) for r in rows],
    }


# ── Dettaglio spedizioni per ordine (drill-down) ────────────────────────────
def get_corrieri_dettaglio(engine: Engine, id_p: int, f: Filters) -> dict:
    """
    Elenco dei singoli ordini con la quota spedizione attribuita al prodotto.
    Legge da pl_prodotti_corrieri_dettaglio.
    Lo snapshot è risolto sulla tabella stessa (può differire dalle altre).
    Ordini ordinati per delta_sped ASC (più in perdita in cima).
    Restituisce rows vuoto se la tabella non esiste ancora; logga l'errore vero.
    """
    # IMPORTANTE: snapshot risolto da TABLE_CORRIERI_DET, non da TABLE_PL.
    snap = resolve_snapshot(engine, f, TABLE_CORRIERI_DET)
    sql = f"""
        SELECT id_ordine, corriere, quantita,
               sped_fatt, sped_costi, delta_sped
        FROM {TABLE_CORRIERI_DET}
        WHERE id_p         = :id_p
          AND data_snapshot  = :data_snapshot
          AND periodo_giorni = :periodo_giorni
        ORDER BY delta_sped ASC
    """
    rows = []
    try:
        with engine.connect() as conn:
            rows = conn.execute(text(sql), {
                "id_p": int(id_p),
                "data_snapshot": snap,
                "periodo_giorni": f.periodo_giorni,
            }).mappings().fetchall()
    except Exception:
        log.exception(
            "get_corrieri_dettaglio fallita id_p=%s snap=%s periodo=%s",
            id_p, snap, f.periodo_giorni,
        )

    log.info(
        "get_corrieri_dettaglio id_p=%s snap=%s periodo=%s -> %d righe",
        id_p, snap, f.periodo_giorni, len(rows),
    )
    print(
        f"[DEBUG-PL] get_corrieri_dettaglio id_p={id_p} snap={snap} periodo={f.periodo_giorni} -> {len(rows)} righe",
        flush=True,
    )

    return {
        "id_p": int(id_p),
        "data_snapshot": snap.isoformat(),
        "periodo_giorni": f.periodo_giorni,
        "rows": [_row_to_dict(r) for r in rows],
    }


# ── Resi globali (aggregato su tutti i prodotti) ─────────────────────────────
def get_resi_global(engine: Engine, f: Filters) -> dict:
    """Aggregato resi senza filtro su id_p: utile per report resi stand-alone."""
    where, params = _where_resi(f)

    sql_agg_tipo = f"""
        SELECT
            tipo_rma,
            COUNT(*)                  AS n_resi,
            SUM(quantita_rma)         AS qty,
            SUM(costo_perdita)        AS costo_perdita,
            SUM(valore_claims)        AS claims,
            SUM(valore_recuperato)    AS recuperato,
            SUM(perdita_netta)        AS perdita_netta
        FROM {TABLE_RESI}
        {where}
        GROUP BY tipo_rma
        ORDER BY perdita_netta DESC
    """
    sql_agg_mktp = f"""
        SELECT
            marketplace,
            COUNT(*)                  AS n_resi,
            SUM(quantita_rma)         AS qty,
            SUM(perdita_netta)        AS perdita_netta
        FROM {TABLE_RESI}
        {where}
        GROUP BY marketplace
        ORDER BY perdita_netta DESC
    """
    sql_totali = f"""
        SELECT
            COUNT(*)                  AS n_resi,
            SUM(quantita_rma)         AS qty_totale,
            SUM(costo_perdita)        AS perdita_lorda,
            SUM(valore_claims)        AS claims_totali,
            SUM(valore_recuperato)    AS recupero_totale,
            SUM(perdita_netta)        AS perdita_netta_totale
        FROM {TABLE_RESI}
        {where}
    """
    with engine.connect() as conn:
        per_tipo = conn.execute(text(sql_agg_tipo), params).mappings().fetchall()
        per_mktp = conn.execute(text(sql_agg_mktp), params).mappings().fetchall()
        totali = conn.execute(text(sql_totali), params).mappings().fetchone()

    return {
        "totali": _row_to_dict(totali) if totali else {},
        "per_tipo": [_row_to_dict(r) for r in per_tipo],
        "per_marketplace": [_row_to_dict(r) for r in per_mktp],
        "filters": _filters_to_dict(f),
    }


# ── Top-N drilldowns ─────────────────────────────────────────────────────────
def get_top(engine: Engine, f: Filters, metric: str, limit: int = 50, asc: bool = False) -> dict:
    """Top/bottom-N prodotti per una metrica."""
    if metric not in SORTABLE_FIELDS:
        raise ValueError(f"Metrica non valida: {metric}")
    snap = resolve_snapshot(engine, f)
    where, params = _where(f, snap)
    direction = "ASC" if asc else "DESC"
    sql = f"""
        SELECT id_p, codice, nome, marca,
               num_ordini, tot_pezzi, tot_fatturato,
               tot_margine, perc_margine, margine_effettivo, perc_margine_eff,
               qty_resi, perc_resi, imp_eco_resi, delta_sped
        FROM {TABLE_PL}
        {where}
        ORDER BY {metric} {direction}, id_p ASC
        LIMIT :limit
    """
    p = dict(params, limit=int(limit))
    with engine.connect() as conn:
        rows = conn.execute(text(sql), p).mappings().fetchall()
    return {
        "metric": metric,
        "direction": direction,
        "data_snapshot": snap.isoformat(),
        "periodo_giorni": f.periodo_giorni,
        "rows": [_row_to_dict(r) for r in rows],
    }


# ── Corrieri summary (aggregato su tutti i prodotti) ────────────────────────
def get_corrieri_summary(engine: Engine, f: Filters) -> dict:
    """Delta spedizioni aggregato per corriere su tutto il catalogo."""
    snap = resolve_snapshot(engine, f, TABLE_CORRIERI)
    sql = f"""
        SELECT
            corriere,
            SUM(num_ordini)     AS num_ordini,
            SUM(tot_pezzi)      AS tot_pezzi,
            SUM(tot_sped_fatt)  AS tot_sped_fatt,
            SUM(tot_sped_costi) AS tot_sped_costi,
            SUM(delta_sped)     AS delta_sped
        FROM {TABLE_CORRIERI}
        WHERE data_snapshot  = :snap
          AND periodo_giorni = :periodo
        GROUP BY corriere
        ORDER BY delta_sped ASC
    """
    rows = []
    try:
        with engine.connect() as conn:
            rows = conn.execute(
                text(sql), {"snap": snap, "periodo": int(f.periodo_giorni)}
            ).mappings().fetchall()
    except Exception:
        log.exception("get_corrieri_summary snap=%s periodo=%s", snap, f.periodo_giorni)
    return {
        "data_snapshot":  snap.isoformat(),
        "periodo_giorni": f.periodo_giorni,
        "rows": [_row_to_dict(r) for r in rows],
    }



# ── Marketplace breakdown (per prodotto, per marca, o aggregato) ────────────
def get_marketplace_breakdown(
    engine: Engine,
    f: Filters,
    id_p: Optional[int] = None,
    marca: Optional[str] = None,
) -> dict:
    """Breakdown vendite per marketplace.

    - id_p specificato  → singolo prodotto
    - marca specificata → tutti i prodotti di quella marca
    - nessuno dei due   → aggregato su tutti i prodotti del filtro corrente
    """
    snap = resolve_snapshot(engine, f, TABLE_MARKETPLACE)

    if id_p is not None:
        # Query diretta su pl_prodotti_marketplace per singolo prodotto
        sql = f"""
            SELECT
                marketplace,
                SUM(num_ordini)        AS num_ordini,
                SUM(tot_pezzi)         AS tot_pezzi,
                SUM(tot_fatturato)     AS tot_fatturato,
                SUM(tot_margine)       AS tot_margine,
                SUM(tot_sped_fatt)     AS tot_sped_fatt,
                SUM(tot_sped_costi)    AS tot_sped_costi,
                SUM(delta_sped)        AS delta_sped,
                SUM(qty_resi)          AS qty_resi,
                SUM(imp_eco_resi)      AS imp_eco_resi,
                SUM(margine_effettivo) AS margine_effettivo
            FROM {TABLE_MARKETPLACE}
            WHERE data_snapshot  = :snap
              AND periodo_giorni = :periodo
              AND id_p           = :id_p
            GROUP BY marketplace
            ORDER BY tot_fatturato DESC
        """
        params: dict = {"snap": snap, "periodo": int(f.periodo_giorni), "id_p": int(id_p)}
    elif marca is not None:
        # Join con pl_prodotti per filtrare per marca
        sql = f"""
            SELECT
                m.marketplace,
                SUM(m.num_ordini)        AS num_ordini,
                SUM(m.tot_pezzi)         AS tot_pezzi,
                SUM(m.tot_fatturato)     AS tot_fatturato,
                SUM(m.tot_margine)       AS tot_margine,
                SUM(m.tot_sped_fatt)     AS tot_sped_fatt,
                SUM(m.tot_sped_costi)    AS tot_sped_costi,
                SUM(m.delta_sped)        AS delta_sped,
                SUM(m.qty_resi)          AS qty_resi,
                SUM(m.imp_eco_resi)      AS imp_eco_resi,
                SUM(m.margine_effettivo) AS margine_effettivo
            FROM {TABLE_MARKETPLACE} m
            JOIN {TABLE_PL} p
              ON p.id_p           = m.id_p
             AND p.data_snapshot  = m.data_snapshot
             AND p.periodo_giorni = m.periodo_giorni
            WHERE m.data_snapshot  = :snap
              AND m.periodo_giorni = :periodo
              AND p.marca          = :marca
            GROUP BY m.marketplace
            ORDER BY tot_fatturato DESC
        """
        params = {"snap": snap, "periodo": int(f.periodo_giorni), "marca": marca}
    else:
        # Aggregato su tutti i prodotti del periodo
        sql = f"""
            SELECT
                marketplace,
                SUM(num_ordini)        AS num_ordini,
                SUM(tot_pezzi)         AS tot_pezzi,
                SUM(tot_fatturato)     AS tot_fatturato,
                SUM(tot_margine)       AS tot_margine,
                SUM(tot_sped_fatt)     AS tot_sped_fatt,
                SUM(tot_sped_costi)    AS tot_sped_costi,
                SUM(delta_sped)        AS delta_sped,
                SUM(qty_resi)          AS qty_resi,
                SUM(imp_eco_resi)      AS imp_eco_resi,
                SUM(margine_effettivo) AS margine_effettivo
            FROM {TABLE_MARKETPLACE}
            WHERE data_snapshot  = :snap
              AND periodo_giorni = :periodo
            GROUP BY marketplace
            ORDER BY tot_fatturato DESC
        """
        params = {"snap": snap, "periodo": int(f.periodo_giorni)}

    rows = []
    try:
        with engine.connect() as conn:
            rows = conn.execute(text(sql), params).mappings().fetchall()
    except Exception:
        log.exception("get_marketplace_breakdown snap=%s periodo=%s id_p=%s marca=%s",
                      snap, f.periodo_giorni, id_p, marca)
        # Probabilmente schema vecchio (colonne mancanti): ritorna vuoto senza 500
        return {
            "data_snapshot":  snap.isoformat(),
            "periodo_giorni": f.periodo_giorni,
            "rows": [],
            "error": "schema_mismatch",
        }

    # Calcola totali per percentuali (float per evitare Decimal/float mismatch)
    tot_pezzi     = float(sum((r["tot_pezzi"]     or 0) for r in rows))
    tot_fatturato = float(sum((r["tot_fatturato"] or 0) for r in rows))

    result_rows = []
    for r in rows:
        try:
            d = dict(r)
            # Serializza date/Decimal
            for k, v in list(d.items()):
                if isinstance(v, (date, datetime)):
                    d[k] = v.isoformat()
                elif hasattr(v, '__float__') and not isinstance(v, (int, bool)):
                    d[k] = float(v)
            d["perc_pezzi"]       = round(100 * (d.get("tot_pezzi",    0) or 0) / tot_pezzi,     1) if tot_pezzi     else 0.0
            d["perc_fatturato"]   = round(100 * (d.get("tot_fatturato", 0) or 0) / tot_fatturato, 1) if tot_fatturato else 0.0
            fat = d.get("tot_fatturato") or 0
            d["perc_margine"]     = round(100 * (d.get("tot_margine",      0) or 0) / fat, 1) if fat else 0.0
            d["perc_margine_eff"] = round(100 * (d.get("margine_effettivo", 0) or 0) / fat, 1) if fat else 0.0
            result_rows.append(d)
        except Exception:
            log.exception("get_marketplace_breakdown: errore su riga %s", dict(r) if r else r)

    return {
        "data_snapshot":  snap.isoformat(),
        "periodo_giorni": f.periodo_giorni,
        "rows": result_rows,
    }
