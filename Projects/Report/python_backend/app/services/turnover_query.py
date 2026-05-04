"""
Query builder e aggregazioni per la tabella yeppon_stats.actual_turnover.

Schema colonne usate:
    marca              varchar(50)
    quantita           int
    turnover           double       -- fatturato netto della riga
    prezzo_acquisto    double       -- costo per pezzo (pu\u00f2 essere 0/NULL)
    macrocat           varchar(50)  -- macro categoria gi\u00e0 classificata
    cate1              varchar(100)
    cate2              varchar(100)
    cate3              varchar(100)
    nomemktp           varchar(50)  -- nome del marketplace (Yeppon, Amazon, eBay, ...)
    dataordine         timestamp
    provincia          varchar(20)
    codice             varchar(30)
    id_p               int
    nome               varchar(300)
    mktpordinato       varchar(50)
    order_id           int
    fornitore          varchar(50)
    ean                varchar(50)
    user_type          varchar(30)
    id                 bigint
    cod_fornitore      varchar(50)
    tipo_spedizione    varchar(150)
    importo_spe        double
"""
from __future__ import annotations

from dataclasses import dataclass, field
from datetime import date
from typing import Any, List, Optional

import math

import numpy as np
import pandas as pd
from sqlalchemy import text
from sqlalchemy.engine import Engine


TABLE = "actual_turnover"

# ---------------------------------------------------------------------------
# H&A: gli ordini dell'utente idUtente=2884093 vengono classificati come
# marketplace "H&A" anzichè usare il loro nomemktp originale.
# _TABLE_EXPR è una subquery che espone la stessa interfaccia di actual_turnover
# ma con nomemktp già riscritto: basta sostituire TABLE con _TABLE_EXPR nelle query.
# ---------------------------------------------------------------------------
_HA_UTENTE_ID = 2884093

_TABLE_EXPR = f"""(
    SELECT
        t.id, t.marca, t.quantita, t.turnover, t.prezzo_acquisto,
        t.macrocat, t.cate1, t.cate2, t.cate3,
        CASE WHEN oc.idUtente = {_HA_UTENTE_ID} THEN 'H&A'
             ELSE t.nomemktp END                                      AS nomemktp,
        t.dataordine, t.provincia, t.codice, t.id_p, t.nome,
        t.mktpordinato, t.order_id, t.fornitore, t.ean,
        t.user_type, t.cod_fornitore, t.tipo_spedizione, t.importo_spe
    FROM {TABLE} t
    LEFT JOIN smart2.ordini_cliente oc ON oc.id = t.order_id
) AS {TABLE}"""


@dataclass
class Filters:
    """Parametri di filtro normalizzati dall'API."""
    dal: date
    al: date
    marketplace: List[str] = field(default_factory=list)
    fornitore: List[str] = field(default_factory=list)
    macrocat: List[str] = field(default_factory=list)
    cate2: List[str] = field(default_factory=list)
    cate3: List[str] = field(default_factory=list)
    marca: List[str] = field(default_factory=list)
    user_type: List[str] = field(default_factory=list)
    codice: Optional[str] = None


def _build_where(f: Filters) -> tuple[str, dict[str, Any]]:
    """
    Costruisce la clausola WHERE in modo sicuro con parametri nominati.
    Ritorna (where_sql, params).
    """
    clauses = ["dataordine >= :dal", "dataordine < :al_plus1"]
    # al_plus1 = giorno dopo, per prendere l'intera giornata "al" compresa
    params: dict[str, Any] = {
        "dal": f.dal.isoformat(),
        "al_plus1": (pd.Timestamp(f.al) + pd.Timedelta(days=1)).date().isoformat(),
    }

    def add_in(col: str, values: List[str], prefix: str):
        if not values:
            return
        placeholders = ",".join(f":{prefix}_{i}" for i in range(len(values)))
        clauses.append(f"{col} IN ({placeholders})")
        for i, v in enumerate(values):
            params[f"{prefix}_{i}"] = v

    add_in("nomemktp",  f.marketplace, "mktp")
    add_in("fornitore", f.fornitore,   "frn")
    add_in("macrocat",  f.macrocat,    "mc")
    add_in("cate2",     f.cate2,       "c2")
    add_in("cate3",     f.cate3,       "c3")
    add_in("marca",     f.marca,       "mk")
    add_in("user_type", f.user_type,   "ut")

    if f.codice:
        clauses.append("codice = :codice")
        params["codice"] = f.codice

    return " AND ".join(clauses), params


# ---------------------------------------------------------------------------
# Query base (aggregata) + calcolo margine con pandas
# ---------------------------------------------------------------------------

def _aggregate(engine: Engine, group_by: List[str], f: Filters,
               extra_select: str = "") -> pd.DataFrame:
    """
    Esegue l'aggregazione base sulle colonne group_by.
    Ritorna DataFrame con colonne:
        <group_by...>, qta, turnover, costo, turnover_val, costo_val
    Dove *_val sono gli stessi valori ma calcolati SOLO sulle righe
    dove prezzo_acquisto > 0 (per margine "onesto").
    """
    where, params = _build_where(f)

    group_sql = ", ".join(group_by) if group_by else ""
    select_groups = ", ".join(group_by) + "," if group_by else ""

    sql = f"""
        SELECT
            {select_groups}
            COALESCE(SUM(quantita), 0)                                AS qta,
            COALESCE(SUM(turnover), 0)                                AS turnover,
            COALESCE(SUM(prezzo_acquisto), 0)              AS costo,
            COALESCE(SUM(CASE WHEN prezzo_acquisto > 0
                              THEN turnover ELSE 0 END), 0)           AS turnover_val,
            COALESCE(SUM(CASE WHEN prezzo_acquisto > 0
                              THEN prezzo_acquisto
                              ELSE 0 END), 0)                         AS costo_val
            {extra_select}
        FROM {_TABLE_EXPR}
        WHERE {where}
        {"GROUP BY " + group_sql if group_by else ""}
    """

    with engine.connect() as conn:
        df = pd.read_sql(text(sql), conn, params=params)

    return df


def _safe_float(x) -> float:
    """Converte x in float, mappando None/NaN/Inf a 0.0 per uso nei calcoli.
    Serve a evitare che dati sporchi nel DB (es. turnover = NaN in un
    DOUBLE) propaghino NaN nel risultato e facciano esplodere json.dumps
    (Starlette usa allow_nan=False)."""
    if x is None:
        return 0.0
    try:
        f = float(x)
    except (TypeError, ValueError):
        return 0.0
    if math.isnan(f) or math.isinf(f):
        return 0.0
    return f


def _json_safe(df: pd.DataFrame) -> pd.DataFrame:
    """Sostituisce NaN/+Inf/-Inf con None sulle colonne float.
    Starlette.JSONResponse chiama json.dumps con allow_nan=False, quindi
    qualsiasi NaN/Inf rimasto nel DataFrame farebbe esplodere la risposta
    con ValueError: "Out of range float values are not JSON compliant".
    Fix: per ogni colonna float con valori non-finiti, cast a object e
    sostituisci con None."""
    df = df.replace([np.inf, -np.inf], np.nan)
    for col in df.columns:
        if df[col].dtype.kind == "f" and df[col].isna().any():
            df[col] = df[col].astype(object).where(df[col].notna(), None)
    return df


def _add_metrics(df: pd.DataFrame) -> pd.DataFrame:
    """
    Aggiunge prezzo_medio e margine_pct.
    - prezzo_medio = turnover / qta (None se qta=0)
    - margine_pct = (turnover_val - costo_val) / turnover_val * 100
                    (None se turnover_val=0, cio\u00e8 nessuna riga con costo valorizzato)

    Tutti i calcoli passano da _safe_float per evitare NaN/Inf propagati
    da dati sporchi nel DB.
    """
    df = df.copy()
    df["prezzo_medio"] = df.apply(
        lambda r: round(_safe_float(r["turnover"]) / _safe_float(r["qta"]), 2)
                  if _safe_float(r["qta"]) > 0 else None, axis=1
    )
    df["margine_pct"] = df.apply(
        lambda r: round(
            (_safe_float(r["turnover_val"]) - _safe_float(r["costo_val"]))
            / _safe_float(r["turnover_val"]) * 100, 2
        ) if _safe_float(r["turnover_val"]) > 0 else None, axis=1
    )
    df["qta"] = df["qta"].fillna(0).astype(int)
    df["turnover"] = df["turnover"].apply(_safe_float).round(2)
    df["costo"]    = df["costo"].apply(_safe_float).round(2)
    # colonne tecniche che non esponiamo fuori
    df = df.drop(columns=["turnover_val", "costo_val"], errors="ignore")
    # Ultima rete di sicurezza: converti NaN/Inf residui in None.
    df = _json_safe(df)
    return df


# ---------------------------------------------------------------------------
# Endpoint-level
# ---------------------------------------------------------------------------

def get_summary(engine: Engine, f: Filters) -> dict:
    """
    Tabella principale: riga per macrocat + totale complessivo.
    """
    df = _aggregate(engine, ["macrocat"], f)
    df["macrocat"] = df["macrocat"].fillna("(non classificata)")

    # totale generale
    tot = _aggregate(engine, [], f).iloc[0].to_dict()

    df = _add_metrics(df)
    df = df.sort_values("turnover", ascending=False)

    total_row = _add_metrics(pd.DataFrame([tot])).iloc[0].to_dict()

    return {
        "total": total_row,
        "rows": df.to_dict(orient="records"),
    }


def get_drilldown(engine: Engine, level: str, f: Filters) -> dict:
    """
    Drill-down per cate2, cate3, marca o fornitore dentro ai filtri correnti.
    level \u2208 {'cate2','cate3','marca','fornitore'}
    """
    allowed = {"cate2", "cate3", "marca", "fornitore"}
    if level not in allowed:
        raise ValueError(f"level non valido: {level}")

    df = _aggregate(engine, [level], f)
    df[level] = df[level].fillna("(non valorizzato)")
    df = _add_metrics(df)
    df = df.sort_values("turnover", ascending=False)
    return {"level": level, "rows": df.to_dict(orient="records")}


def _pivot_by_marketplace(engine: Engine, group_cols: List[str], f: Filters,
                          mktp_whitelist: Optional[List[str]] = None) -> dict:
    """
    Aggrega per (group_cols + nomemktp) e restituisce una struttura "pivot":
        marketplaces: ordinati per fatturato totale DESC
        rows: lista, una per combinazione di group_cols, con:
            - le group_cols come campi
            - cells: { marketplace_name: {qta, turnover, costo, prezzo_medio, margine_pct} }
            - total:  { qta, turnover, costo, prezzo_medio, margine_pct }  -> somma di riga
        total_row: { cells: {...}, total: {...} }  -> somma complessiva
    """
    where, params = _build_where(f)

    group_sql = ", ".join(group_cols)
    select_groups = group_sql + ","

    sql = f"""
        SELECT
            {select_groups}
            COALESCE(nomemktp, '(n/d)')                               AS _mktp,
            COALESCE(SUM(quantita), 0)                                AS qta,
            COALESCE(SUM(turnover), 0)                                AS turnover,
            COALESCE(SUM(prezzo_acquisto), 0)              AS costo,
            COALESCE(SUM(CASE WHEN prezzo_acquisto > 0
                              THEN turnover ELSE 0 END), 0)           AS turnover_val,
            COALESCE(SUM(CASE WHEN prezzo_acquisto > 0
                              THEN prezzo_acquisto
                              ELSE 0 END), 0)                         AS costo_val
        FROM {_TABLE_EXPR}
        WHERE {where}
        GROUP BY {group_sql}, COALESCE(nomemktp,'(n/d)')
    """

    with engine.connect() as conn:
        df = pd.read_sql(text(sql), conn, params=params)

    # ordine marketplace: per fatturato totale nel sottoinsieme richiesto
    if mktp_whitelist:
        df = df[df["_mktp"].isin(mktp_whitelist)]
    mktp_order = (
        df.groupby("_mktp")["turnover"].sum().sort_values(ascending=False).index.tolist()
    )

    def _row_metrics(qta, turnover, costo, turnover_val, costo_val):
        qta = int(qta or 0)
        turnover = round(_safe_float(turnover), 2)
        costo = round(_safe_float(costo), 2)
        prezzo_medio = round(turnover / qta, 2) if qta > 0 else None
        tv = _safe_float(turnover_val)
        cv = _safe_float(costo_val)
        margine_pct = round((tv - cv) / tv * 100, 2) if tv > 0 else None
        return {
            "qta": qta,
            "turnover": turnover,
            "costo": costo,
            "prezzo_medio": prezzo_medio,
            "margine_pct": margine_pct,
        }

    # build rows: una per tupla group_cols
    rows_out: list[dict] = []
    if not df.empty:
        for grp_values, grp_df in df.groupby(group_cols, dropna=False):
            if not isinstance(grp_values, tuple):
                grp_values = (grp_values,)
            row: dict = {}
            for col, val in zip(group_cols, grp_values):
                row[col] = "(non classificata)" if pd.isna(val) else val

            cells: dict = {}
            for _, r in grp_df.iterrows():
                cells[r["_mktp"]] = _row_metrics(
                    r["qta"], r["turnover"], r["costo"],
                    r["turnover_val"], r["costo_val"],
                )
            row["cells"] = cells
            row["total"] = _row_metrics(
                grp_df["qta"].sum(), grp_df["turnover"].sum(), grp_df["costo"].sum(),
                grp_df["turnover_val"].sum(), grp_df["costo_val"].sum(),
            )
            rows_out.append(row)

    # totale colonne (sum di tutte le righe)
    total_cells: dict = {}
    for m in mktp_order:
        sub = df[df["_mktp"] == m]
        total_cells[m] = _row_metrics(
            sub["qta"].sum(), sub["turnover"].sum(), sub["costo"].sum(),
            sub["turnover_val"].sum(), sub["costo_val"].sum(),
        )
    total_all = _row_metrics(
        df["qta"].sum() if not df.empty else 0,
        df["turnover"].sum() if not df.empty else 0,
        df["costo"].sum() if not df.empty else 0,
        df["turnover_val"].sum() if not df.empty else 0,
        df["costo_val"].sum() if not df.empty else 0,
    )

    # ordina rows per totale riga turnover DESC
    rows_out.sort(key=lambda r: r["total"]["turnover"], reverse=True)

    return {
        "marketplaces": mktp_order,
        "rows": rows_out,
        "total_row": {"cells": total_cells, "total": total_all},
    }


def get_summary_by_marketplace(engine: Engine, f: Filters) -> dict:
    """Tabella Riepilogo in modalit\u00e0 pivot: una riga per macrocat, colonne per marketplace."""
    return _pivot_by_marketplace(engine, ["macrocat"], f)


def get_drilldown_by_marketplace(engine: Engine, level: str, f: Filters) -> dict:
    """Drill-down in modalit\u00e0 pivot. level \u2208 {cate2, cate3, marca, fornitore}."""
    allowed = {"cate2", "cate3", "marca", "fornitore"}
    if level not in allowed:
        raise ValueError(f"level non valido: {level}")
    res = _pivot_by_marketplace(engine, [level], f)
    res["level"] = level
    return res


def get_product_list_by_marketplace(engine: Engine, f: Filters,
                                     page: int = 1, page_size: int = 500,
                                     sort_by: str = "turnover",
                                     sort_dir: str = "desc") -> dict:
    """
    Lista prodotti paginata in modalit\u00e0 pivot per marketplace.
    Ritorna base list (come get_product_list) con in pi\u00f9:
        - marketplaces: [...]
        - per ogni row: row["cells"] = { mktp: {qta,turnover,costo,prezzo_medio,margine_pct} }
    """
    # 1. usa la paginazione "piatta" standard per decidere quali codici finiscono in pagina
    base = get_product_list(engine, f, page=page, page_size=page_size,
                            sort_by=sort_by, sort_dir=sort_dir)
    if not base["rows"]:
        base["marketplaces"] = []
        return base

    codici_pagina = [r["codice"] for r in base["rows"]]

    # 2. query aggregata per (codice, nomemktp) solo per i codici della pagina
    where, params = _build_where(f)
    placeholders = ",".join(f":cod_{i}" for i in range(len(codici_pagina)))
    p2 = dict(params)
    for i, c in enumerate(codici_pagina):
        p2[f"cod_{i}"] = c

    sql = f"""
        SELECT
            COALESCE(codice,'(s/c)')                                  AS codice,
            COALESCE(nomemktp,'(n/d)')                                AS _mktp,
            COALESCE(SUM(quantita), 0)                                AS qta,
            COALESCE(SUM(turnover), 0)                                AS turnover,
            COALESCE(SUM(prezzo_acquisto), 0)              AS costo,
            COALESCE(SUM(CASE WHEN prezzo_acquisto > 0
                              THEN turnover ELSE 0 END), 0)           AS turnover_val,
            COALESCE(SUM(CASE WHEN prezzo_acquisto > 0
                              THEN prezzo_acquisto
                              ELSE 0 END), 0)                         AS costo_val
        FROM {_TABLE_EXPR}
        WHERE {where}
          AND COALESCE(codice,'(s/c)') IN ({placeholders})
        GROUP BY COALESCE(codice,'(s/c)'), COALESCE(nomemktp,'(n/d)')
    """
    with engine.connect() as conn:
        df = pd.read_sql(text(sql), conn, params=p2)

    # ordine marketplace globale per fatturato nella pagina
    mktp_order = (
        df.groupby("_mktp")["turnover"].sum()
          .sort_values(ascending=False).index.tolist()
        if not df.empty else []
    )

    # mappa codice -> cells
    cells_by_codice: dict[str, dict] = {c: {} for c in codici_pagina}
    for _, r in df.iterrows():
        qta = int(r["qta"] or 0)
        turn = _safe_float(r["turnover"])
        prezzo_medio = round(turn / qta, 2) if qta > 0 else None
        tval = _safe_float(r["turnover_val"])
        cval = _safe_float(r["costo_val"])
        margine = round((tval - cval) / tval * 100, 2) if tval > 0 else None
        cells_by_codice.setdefault(r["codice"], {})[r["_mktp"]] = {
            "qta": qta,
            "turnover": round(turn, 2),
            "costo": round(_safe_float(r["costo"]), 2),
            "prezzo_medio": prezzo_medio,
            "margine_pct": margine,
        }

    # aggancia cells a ogni row della pagina, preservando l'ordine
    for row in base["rows"]:
        row["cells"] = cells_by_codice.get(row["codice"], {})

    base["marketplaces"] = mktp_order
    return base


def get_marketplace_breakdown(engine: Engine, f: Filters) -> dict:
    """
    Ripartizione per marketplace (nomemktp) con i filtri correnti.
    """
    df = _aggregate(engine, ["nomemktp"], f)
    df["nomemktp"] = df["nomemktp"].fillna("(non valorizzato)")
    df = _add_metrics(df)
    df = df.sort_values("turnover", ascending=False)
    return {"rows": df.to_dict(orient="records")}


def get_trend(engine: Engine, dimension: str, f: Filters) -> dict:
    """
    Serie temporale / geografica.
    dimension \u2208 {'day','week','hour','weekday','month','province','marketplace'}
    Ritorna: { rows: [ { bucket, qta, turnover, prezzo_medio, margine_pct }, ... ] }
    """
    if dimension == "day":
        group_expr = "DATE(dataordine) AS bucket"
        order = "bucket"
    elif dimension == "week":
        # Lunedi' della settimana ISO come data; cosi' la UI puo' formattarla facilmente.
        # MySQL WEEKDAY: 0=Lun..6=Dom, quindi DATE_SUB(data, WEEKDAY(data) GIORNI) = lunedi' della settimana.
        group_expr = "DATE(DATE_SUB(dataordine, INTERVAL WEEKDAY(dataordine) DAY)) AS bucket"
        order = "bucket"
    elif dimension == "hour":
        group_expr = "HOUR(dataordine) AS bucket"
        order = "bucket"
    elif dimension == "weekday":
        # 1=Lun ... 7=Dom (MySQL WEEKDAY: 0=Lun..6=Dom; shift +1)
        group_expr = "(WEEKDAY(dataordine) + 1) AS bucket"
        order = "bucket"
    elif dimension == "month":
        group_expr = "DATE_FORMAT(dataordine, '%Y-%m') AS bucket"
        order = "bucket"
    elif dimension == "province":
        group_expr = "COALESCE(provincia,'(n/d)') AS bucket"
        order = "turnover DESC"
    elif dimension == "marketplace":
        group_expr = "COALESCE(nomemktp,'(n/d)') AS bucket"
        order = "turnover DESC"
    else:
        raise ValueError(f"dimension non valida: {dimension}")

    where, params = _build_where(f)

    sql = f"""
        SELECT
            {group_expr},
            COALESCE(SUM(quantita), 0)                                AS qta,
            COALESCE(SUM(turnover), 0)                                AS turnover,
            COALESCE(SUM(prezzo_acquisto), 0)              AS costo,
            COALESCE(SUM(CASE WHEN prezzo_acquisto > 0
                              THEN turnover ELSE 0 END), 0)           AS turnover_val,
            COALESCE(SUM(CASE WHEN prezzo_acquisto > 0
                              THEN prezzo_acquisto
                              ELSE 0 END), 0)                         AS costo_val
        FROM {_TABLE_EXPR}
        WHERE {where}
        GROUP BY bucket
        ORDER BY {order}
    """

    with engine.connect() as conn:
        df = pd.read_sql(text(sql), conn, params=params)

    df = _add_metrics(df)
    # cast ISO string per i date
    if dimension in ("day", "week"):
        df["bucket"] = df["bucket"].astype(str)
    return {"dimension": dimension, "rows": df.to_dict(orient="records")}


def get_product_list(engine: Engine, f: Filters,
                     page: int = 1, page_size: int = 500,
                     sort_by: str = "turnover", sort_dir: str = "desc") -> dict:
    """
    Elenco paginato dei prodotti (uno per codice) con tutte le metriche di vendita
    nel periodo/filtri selezionati.
    """
    if page < 1:
        page = 1
    if page_size < 1 or page_size > 2000:
        page_size = 500

    sort_cols = {
        "codice": "codice",
        "nome": "MAX(nome)",
        "marca": "MAX(marca)",
        "qta": "SUM(quantita)",
        "turnover": "SUM(turnover)",
        "costo": "SUM(prezzo_acquisto)",
        "ordini": "COUNT(DISTINCT order_id)",
        "last_order": "MAX(dataordine)",
    }
    sort_expr = sort_cols.get(sort_by, "SUM(turnover)")
    sort_dir_sql = "ASC" if str(sort_dir).lower() == "asc" else "DESC"

    where, params = _build_where(f)

    count_sql = f"SELECT COUNT(DISTINCT COALESCE(codice,'(s/c)')) FROM {_TABLE_EXPR} WHERE {where}"

    data_sql = f"""
        SELECT
            COALESCE(codice, '(s/c)')                                 AS codice,
            MAX(nome)                                                 AS nome,
            MAX(marca)                                                AS marca,
            MAX(macrocat)                                             AS macrocat,
            MAX(cate2)                                                AS cate2,
            MAX(cate3)                                                AS cate3,
            MAX(ean)                                                  AS ean,
            GROUP_CONCAT(DISTINCT fornitore ORDER BY fornitore SEPARATOR ', ')    AS fornitori,
            GROUP_CONCAT(DISTINCT nomemktp ORDER BY nomemktp SEPARATOR ', ')      AS marketplaces,
            COALESCE(SUM(quantita), 0)                                AS qta,
            COUNT(DISTINCT order_id)                                  AS ordini,
            COUNT(DISTINCT DATE(dataordine))                          AS giorni_attivita,
            COALESCE(SUM(turnover), 0)                                AS turnover,
            COALESCE(SUM(prezzo_acquisto), 0)              AS costo,
            COALESCE(SUM(CASE WHEN prezzo_acquisto > 0
                              THEN turnover ELSE 0 END), 0)           AS turnover_val,
            COALESCE(SUM(CASE WHEN prezzo_acquisto > 0
                              THEN prezzo_acquisto
                              ELSE 0 END), 0)                         AS costo_val,
            MIN(dataordine)                                           AS first_order,
            MAX(dataordine)                                           AS last_order
        FROM {_TABLE_EXPR}
        WHERE {where}
        GROUP BY COALESCE(codice,'(s/c)')
        ORDER BY {sort_expr} {sort_dir_sql}
        LIMIT :_limit OFFSET :_offset
    """

    q_params = dict(params)
    q_params["_limit"] = int(page_size)
    q_params["_offset"] = int((page - 1) * page_size)

    with engine.connect() as conn:
        total_rows = conn.execute(text(count_sql), params).scalar() or 0
        df = pd.read_sql(text(data_sql), conn, params=q_params)

    if df.empty:
        return {
            "total_rows": int(total_rows),
            "page": page, "page_size": page_size,
            "pages": 0, "rows": [],
            "sort_by": sort_by, "sort_dir": sort_dir_sql.lower(),
        }

    df["prezzo_medio"] = df.apply(
        lambda r: round(_safe_float(r["turnover"]) / _safe_float(r["qta"]), 2)
                  if _safe_float(r["qta"]) > 0 else None, axis=1
    )
    df["margine_pct"] = df.apply(
        lambda r: round(
            (_safe_float(r["turnover_val"]) - _safe_float(r["costo_val"]))
            / _safe_float(r["turnover_val"]) * 100, 2
        ) if _safe_float(r["turnover_val"]) > 0 else None, axis=1
    )
    df["qta"] = df["qta"].fillna(0).astype(int)
    df["ordini"] = df["ordini"].fillna(0).astype(int)
    df["giorni_attivita"] = df["giorni_attivita"].fillna(0).astype(int)
    df["turnover"] = df["turnover"].apply(_safe_float).round(2)
    df["costo"]    = df["costo"].apply(_safe_float).round(2)
    df["first_order"] = df["first_order"].astype(str)
    df["last_order"] = df["last_order"].astype(str)
    df = df.drop(columns=["turnover_val", "costo_val"], errors="ignore")
    # Ultima rete di sicurezza NaN/Inf -> None per JSON-compliance
    df = _json_safe(df)

    pages = (int(total_rows) + page_size - 1) // page_size
    return {
        "total_rows": int(total_rows),
        "page": page,
        "page_size": page_size,
        "pages": pages,
        "sort_by": sort_by,
        "sort_dir": sort_dir_sql.lower(),
        "rows": df.to_dict(orient="records"),
    }


def get_product_trend(engine: Engine, f: Filters) -> dict:
    """
    Trend giornaliero del singolo prodotto (quando codice \u00e8 specificato nei filtri).
    """
    if not f.codice:
        raise ValueError("codice prodotto richiesto per product-trend")

    where, params = _build_where(f)

    sql = f"""
        SELECT
            DATE(dataordine)                           AS data,
            MONTH(dataordine)                          AS mese,
            COALESCE(SUM(quantita), 0)                 AS qta,
            COALESCE(SUM(turnover), 0)                 AS turnover
        FROM {_TABLE_EXPR}
        WHERE {where}
        GROUP BY DATE(dataordine), MONTH(dataordine)
        ORDER BY data
    """
    with engine.connect() as conn:
        df = pd.read_sql(text(sql), conn, params=params)

    if df.empty:
        return {"codice": f.codice, "rows": [], "quarters": {}}

    df["data"] = df["data"].astype(str)
    df["qta"] = df["qta"].astype(int)
    df["turnover"] = df["turnover"].astype(float).round(2)
    df["prezzo_medio"] = df.apply(
        lambda r: round(r["turnover"] / r["qta"], 2) if r["qta"] > 0 else 0, axis=1
    )

    # raggruppa per trimestre per la UI che mostrava Q1/Q2/Q3/Q4
    quarters = {
        "Q1": df[df["mese"].between(1, 3)].to_dict(orient="records"),
        "Q2": df[df["mese"].between(4, 6)].to_dict(orient="records"),
        "Q3": df[df["mese"].between(7, 9)].to_dict(orient="records"),
        "Q4": df[df["mese"].between(10, 12)].to_dict(orient="records"),
    }
    return {
        "codice": f.codice,
        "rows": df.drop(columns=["mese"]).to_dict(orient="records"),
        "quarters": quarters,
    }


# ---------------------------------------------------------------------------
# Opzioni filtri
# ---------------------------------------------------------------------------

def get_filter_options(engine: Engine) -> dict:
    """
    Ritorna la lista dei valori distinti per popolare le dropdown.
    Calcolato sugli ultimi 365 giorni per evitare query enormi.
    """
    result: dict[str, list[str]] = {
        "marketplace": [],
        "fornitore": [],
        "macrocat": [],
        "user_type": [],
    }

    queries = {
        "marketplace": f"SELECT DISTINCT nomemktp FROM {_TABLE_EXPR} "
                       "WHERE dataordine >= (CURRENT_DATE - INTERVAL 365 DAY) "
                       "AND nomemktp IS NOT NULL AND nomemktp <> '' ORDER BY nomemktp",
        "fornitore":   "SELECT DISTINCT fornitore FROM actual_turnover "
                       "WHERE dataordine >= (CURRENT_DATE - INTERVAL 365 DAY) "
                       "AND fornitore IS NOT NULL AND fornitore <> '' ORDER BY fornitore",
        "macrocat":    "SELECT DISTINCT macrocat FROM actual_turnover "
                       "WHERE dataordine >= (CURRENT_DATE - INTERVAL 365 DAY) "
                       "AND macrocat IS NOT NULL AND macrocat <> '' ORDER BY macrocat",
        "user_type":   "SELECT DISTINCT user_type FROM actual_turnover "
                       "WHERE dataordine >= (CURRENT_DATE - INTERVAL 365 DAY) "
                       "AND user_type IS NOT NULL AND user_type <> '' ORDER BY user_type",
    }
    with engine.connect() as conn:
        for key, q in queries.items():
            rows = conn.execute(text(q)).all()
            result[key] = [r[0] for r in rows]
    return result


def get_dependent_options(engine: Engine, f: Filters) -> dict:
    """
    Opzioni dipendenti dai filtri attuali: cate2, cate3, marca.
    Serve a popolare i dropdown in cascata (es. se scegli macrocat=GED,
    la lista cate2 mostra solo quelle della macro).
    """
    where, params = _build_where(f)

    qs = {
        "cate2": f"SELECT DISTINCT cate2 FROM {TABLE} WHERE {where} "
                 f"AND cate2 IS NOT NULL AND cate2 <> '' ORDER BY cate2",
        "cate3": f"SELECT DISTINCT cate3 FROM {TABLE} WHERE {where} "
                 f"AND cate3 IS NOT NULL AND cate3 <> '' ORDER BY cate3",
        "marca": f"SELECT DISTINCT marca FROM {TABLE} WHERE {where} "
                 f"AND marca IS NOT NULL AND marca <> '' ORDER BY marca",
    }
    out: dict[str, list[str]] = {}
    with engine.connect() as conn:
        for key, q in qs.items():
            rows = conn.execute(text(q), params).all()
            out[key] = [r[0] for r in rows]
    return out
