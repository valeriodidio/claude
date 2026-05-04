"""Endpoint HTTP del report P&L Prodotti.

I router non contengono logica SQL: si limitano a leggere querystring,
costruire i Filters, chiamare i service e restituire JSON (o Excel).

Filtri principali:
  data_snapshot  → data dello snapshot da mostrare (default: ultima disponibile)
  periodo_giorni → finestra di analisi: 30/60/90/120/180/360 (default: 180)
  marca[]        → filtra per marca
  status_prodotto → 1=attivo / 0=disattivo
  bloccato       → 1=sì / 0=no
  margine_negativo → true = solo prodotti in perdita
  solo_con_resi  → true = solo prodotti con resi
  search         → testo libero su codice/nome/id_p
  min_fatturato  → fatturato minimo
  min_ordini     → numero ordini minimo
  marketplace[]  → filtra resi per marketplace
  tipo_rma[]     → filtra resi per tipo
  resi_dal/resi_al → filtra resi per data
"""
from datetime import date, datetime
from typing import List, Optional

from fastapi import APIRouter, Depends, HTTPException, Query
from fastapi.responses import StreamingResponse

from app.config import get_settings
from app.db import get_engine
from app.services import excel_export
from app.services import pl_prodotti_query as q
from app.services.pl_prodotti_query import Filters

router = APIRouter()


# ── Dependency: filtri comuni ────────────────────────────────────────────────
def _filters(
    # Dimensioni snapshot
    data_snapshot: Optional[date] = Query(None, description="Data snapshot (default: ultima disponibile)"),
    periodo_giorni: int = Query(90, description="Finestra analisi: 30/60/90/120/180/360"),
    # Filtri prodotto
    marca: Optional[List[str]] = Query(None),
    status_prodotto: Optional[int] = Query(None, ge=0, le=1, description="1=attivo, 0=disattivo"),
    bloccato: Optional[int] = Query(None, ge=0, le=1, description="1=bloccato, 0=non bloccato"),
    margine_negativo: bool = Query(False, description="Solo prodotti con margine effettivo < 0"),
    solo_con_resi: bool = Query(False, description="Solo prodotti con almeno un reso"),
    search: Optional[str] = Query(None, max_length=80, description="Cerca su codice/nome/id_p"),
    min_fatturato: Optional[float] = Query(None, ge=0),
    min_ordini: Optional[int] = Query(None, ge=0),
    # Filtri categoria/sender/fornitore
    categoria:  Optional[List[str]] = Query(None),
    categoria2: Optional[List[str]] = Query(None),
    categoria3: Optional[List[str]] = Query(None),
    sender:     Optional[List[str]] = Query(None),
    fornitore:  Optional[List[str]] = Query(None),
    # Filtri resi
    marketplace: Optional[List[str]] = Query(None),
    tipo_rma: Optional[List[str]] = Query(None),
    resi_dal: Optional[date] = Query(None, description="Data inizio resi (inclusa)"),
    resi_al: Optional[date] = Query(None, description="Data fine resi (inclusa)"),
    # Paginazione/ordinamento
    sort_by: str = Query("tot_fatturato"),
    sort_dir: str = Query("desc"),
    page: int = Query(1, ge=1),
    page_size: int = Query(100, ge=1, le=99999),
) -> Filters:
    periodo = periodo_giorni if periodo_giorni in q.PERIODI_VALIDI else 90
    return Filters(
        data_snapshot=data_snapshot,
        periodo_giorni=periodo,
        marca=marca or [],
        status_prodotto=status_prodotto,
        bloccato=bloccato,
        margine_negativo=margine_negativo,
        solo_con_resi=solo_con_resi,
        search=(search or "").strip() or None,
        min_fatturato=min_fatturato,
        min_ordini=min_ordini,
        categoria=categoria or [],
        categoria2=categoria2 or [],
        categoria3=categoria3 or [],
        sender=sender or [],
        fornitore=fornitore or [],
        marketplace=marketplace or [],
        tipo_rma=tipo_rma or [],
        resi_dal=resi_dal,
        resi_al=resi_al,
        sort_by=sort_by,
        sort_dir=sort_dir,
        page=page,
        page_size=page_size,
    )


# ── Endpoint: opzioni filtri ─────────────────────────────────────────────────
@router.get("/filters")
def filters_options():
    """Snapshot disponibili, periodi, marche, marketplace, tipi RMA."""
    return q.get_filter_options(get_engine())


# ── Endpoint: KPI aggregati ──────────────────────────────────────────────────
@router.get("/summary")
def summary(f: Filters = Depends(_filters)):
    """KPI aggregati su tutto il dataset filtrato (no paginazione)."""
    return q.get_summary(get_engine(), f)


# ── Endpoint: lista paginata ─────────────────────────────────────────────────
@router.get("/list")
def list_products(f: Filters = Depends(_filters)):
    """Elenco prodotti con paginazione e ordinamento."""
    return q.get_list(get_engine(), f)


# ── Endpoint: detail singolo prodotto ────────────────────────────────────────
@router.get("/detail/{id_p}")
def detail(
    id_p: int,
    periodo_giorni: int = Query(90),
    data_snapshot: Optional[date] = Query(None),
):
    f = Filters(periodo_giorni=periodo_giorni, data_snapshot=data_snapshot)
    data = q.get_detail(get_engine(), id_p, f)
    if not data:
        raise HTTPException(
            404,
            f"Prodotto {id_p} non trovato per snapshot={data_snapshot}, periodo={periodo_giorni}",
        )
    return data


# ── Endpoint: resi per prodotto ───────────────────────────────────────────────
@router.get("/resi/{id_p}")
def resi(id_p: int, f: Filters = Depends(_filters)):
    """Dettaglio per RMA + aggregato per tipo di reso di un prodotto."""
    return q.get_resi(get_engine(), id_p, f)


# ── Endpoint: corrieri spedizione per prodotto ───────────────────────────────
@router.get("/corrieri/{id_p}")
def corrieri(
    id_p: int,
    periodo_giorni: int = Query(90),
    data_snapshot: Optional[date] = Query(None),
):
    """Dettaglio spedizioni per corriere di un singolo prodotto."""
    f = Filters(periodo_giorni=periodo_giorni, data_snapshot=data_snapshot)
    return q.get_corrieri(get_engine(), id_p, f)


# ── Endpoint: dettaglio ordini spedizione per prodotto ───────────────────────
@router.get("/corrieri_dettaglio/{id_p}")
def corrieri_dettaglio(
    id_p: int,
    periodo_giorni: int = Query(90),
    data_snapshot: Optional[date] = Query(None),
):
    """Drill-down spedizioni per singolo ordine di un prodotto."""
    f = Filters(periodo_giorni=periodo_giorni, data_snapshot=data_snapshot)
    return q.get_corrieri_dettaglio(get_engine(), id_p, f)


# ── Endpoint: resi globali (tutti i prodotti) ────────────────────────────────
@router.get("/resi")
def resi_global(f: Filters = Depends(_filters)):
    """Aggregato resi su tutti i prodotti, filtrabile per tipo/marketplace/data."""
    return q.get_resi_global(get_engine(), f)


# ── Endpoint: export Excel ───────────────────────────────────────────────────
@router.get("/export.xlsx")
def export_xlsx(f: Filters = Depends(_filters)):
    """Esporta la lista prodotti (con tutti i filtri attivi) in formato Excel."""
    import pandas as pd

    # Prende tutto senza paginazione
    f_all = Filters(
        **{k: v for k, v in vars(f).items() if k not in ("page", "page_size")},
        page=1,
        page_size=99999,
    )
    result = q.get_list(get_engine(), f_all)
    rows = result.get("rows", [])

    df = pd.DataFrame(rows)
    xlsx_bytes = excel_export.df_to_excel_bytes(df, periodo_giorni=f.periodo_giorni)

    from datetime import datetime
    fname = f"pl_prodotti_{datetime.now():%Y%m%d_%H%M}.xlsx"
    return StreamingResponse(
        iter([xlsx_bytes]),
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": f'attachment; filename="{fname}"'},
    )


# ── Endpoint: top-N per metrica ──────────────────────────────────────────────
@router.get("/top/{metric}")
def top(
    metric: str,
    asc: bool = Query(False, description="Se true, restituisce i peggiori"),
    limit: int = Query(50, ge=1, le=500),
    f: Filters = Depends(_filters),
):
    try:
        return q.get_top(get_engine(), f, metric=metric, limit=limit, asc=asc)
    except ValueError as e:
        raise HTTPException(400, str(e))


# ── Endpoint: delta spedizioni aggregato per corriere ────────────────────────
@router.get("/corrieri_summary")
def corrieri_summary(
    periodo_giorni: int            = Query(90),
    data_snapshot:  Optional[date] = Query(None),
):
    """Delta spedizioni aggregato per corriere su tutto il catalogo."""
    f = Filters(periodo_giorni=periodo_giorni, data_snapshot=data_snapshot)
    return q.get_corrieri_summary(get_engine(), f)



# ── Endpoint: breakdown vendite per marketplace ──────────────────────────────
@router.get("/marketplace_breakdown")
def marketplace_breakdown(
    id_p:           Optional[int]  = Query(None, description="Singolo prodotto (opzionale)"),
    marca:          Optional[str]  = Query(None, description="Filtra per marca"),
    periodo_giorni: int            = Query(90),
    data_snapshot:  Optional[date] = Query(None),
):
    """Breakdown vendite per marketplace di un singolo prodotto, marca o tutto il catalogo."""
    f = Filters(periodo_giorni=periodo_giorni, data_snapshot=data_snapshot)
    return q.get_marketplace_breakdown(get_engine(), f, id_p=id_p, marca=marca)
