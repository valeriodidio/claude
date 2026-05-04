"""
Endpoint HTTP del report Actual Turnover.
Prefisso applicato nel main.py: /api/reports/turnover
"""
from __future__ import annotations

from datetime import date, timedelta
from typing import List, Optional

from fastapi import APIRouter, Depends, HTTPException, Query
from fastapi.responses import Response

from app.db import get_engine
from app.services import turnover_query as q
from app.services.excel_export import (
    build_xlsx,
    build_xlsx_split,
    build_xlsx_product_list,
    build_xlsx_product_list_split,
)
from app.services.turnover_query import Filters

router = APIRouter()


# ---------------------------------------------------------------------------
# Dependency: filtri comuni da query string
# ---------------------------------------------------------------------------

def _filters(
    dal: date = Query(..., description="Data inizio (YYYY-MM-DD)"),
    al: date = Query(..., description="Data fine (YYYY-MM-DD, inclusiva)"),
    marketplace: Optional[List[str]] = Query(None, description="Multi-select"),
    fornitore:   Optional[List[str]] = Query(None, description="Multi-select"),
    macrocat:    Optional[List[str]] = Query(None, description="Multi-select"),
    cate2:       Optional[List[str]] = Query(None, description="Multi-select"),
    cate3:       Optional[List[str]] = Query(None, description="Multi-select"),
    marca:       Optional[List[str]] = Query(None, description="Multi-select"),
    user_type:   Optional[List[str]] = Query(None, description="Multi-select"),
    codice:      Optional[str] = Query(None),
) -> Filters:
    if al < dal:
        raise HTTPException(400, "al non pu\u00f2 essere minore di dal")
    # hard cap per evitare query su intervalli enormi
    if (al - dal).days > 400:
        raise HTTPException(400, "intervallo massimo supportato: 400 giorni")

    def _clean(xs: Optional[List[str]]) -> List[str]:
        return [x for x in (xs or []) if x is not None and x != ""]

    return Filters(
        dal=dal,
        al=al,
        marketplace=_clean(marketplace),
        fornitore=_clean(fornitore),
        macrocat=_clean(macrocat),
        cate2=_clean(cate2),
        cate3=_clean(cate3),
        marca=_clean(marca),
        user_type=_clean(user_type),
        codice=codice,
    )


# ---------------------------------------------------------------------------
# Filtri
# ---------------------------------------------------------------------------

@router.get("/filters")
def filters_options():
    """Opzioni statiche per i dropdown (calcolate sugli ultimi 365gg)."""
    return q.get_filter_options(get_engine())


@router.get("/filters/dependent")
def dependent_options(f: Filters = Depends(_filters)):
    """Opzioni dipendenti dai filtri attuali (cate2/cate3/marca)."""
    return q.get_dependent_options(get_engine(), f)


# ---------------------------------------------------------------------------
# Report
# ---------------------------------------------------------------------------

@router.get("/summary")
def summary(f: Filters = Depends(_filters)):
    """Tabella principale per macrocat + totale complessivo."""
    return q.get_summary(get_engine(), f)


@router.get("/drilldown/{level}")
def drilldown(level: str, f: Filters = Depends(_filters)):
    """Drill inline: level \u2208 {cate2, cate3, marca, fornitore}."""
    try:
        return q.get_drilldown(get_engine(), level, f)
    except ValueError as e:
        raise HTTPException(400, str(e))


@router.get("/marketplace-breakdown")
def marketplace_breakdown(f: Filters = Depends(_filters)):
    """Ripartizione dinamica per marketplace con i filtri attuali."""
    return q.get_marketplace_breakdown(get_engine(), f)


# --- Split per marketplace (pivot) -----------------------------------------

@router.get("/summary-by-marketplace")
def summary_by_marketplace(f: Filters = Depends(_filters)):
    """Riepilogo pivot: righe=macrocat, colonne=marketplace."""
    return q.get_summary_by_marketplace(get_engine(), f)


@router.get("/drilldown-by-marketplace/{level}")
def drilldown_by_marketplace(level: str, f: Filters = Depends(_filters)):
    """Drill-down pivot. level \u2208 {cate2, cate3, marca, fornitore}."""
    try:
        return q.get_drilldown_by_marketplace(get_engine(), level, f)
    except ValueError as e:
        raise HTTPException(400, str(e))


@router.get("/product-list-by-marketplace")
def product_list_by_marketplace(
    page: int = Query(1, ge=1),
    page_size: int = Query(500, ge=1, le=2000),
    sort_by: str = Query("turnover"),
    sort_dir: str = Query("desc"),
    f: Filters = Depends(_filters),
):
    """Lista prodotti pivot per marketplace."""
    return q.get_product_list_by_marketplace(
        get_engine(), f,
        page=page, page_size=page_size,
        sort_by=sort_by, sort_dir=sort_dir,
    )


@router.get("/trend/{dimension}")
def trend(dimension: str, f: Filters = Depends(_filters)):
    """
    dimension:
      - day         (serie giornaliera)
      - week        (serie settimanale; bucket = lunedi' della settimana ISO)
      - hour        (0..23)
      - weekday     (1=Lun..7=Dom)
      - month       (YYYY-MM)
      - province    (top province)
      - marketplace (alias di /marketplace-breakdown per UI uniforme)
    """
    try:
        return q.get_trend(get_engine(), dimension, f)
    except ValueError as e:
        raise HTTPException(400, str(e))


@router.get("/product-trend")
def product_trend(f: Filters = Depends(_filters)):
    """Trend giornaliero di un singolo prodotto (richiede ?codice=...)."""
    if not f.codice:
        raise HTTPException(400, "parametro 'codice' obbligatorio")
    return q.get_product_trend(get_engine(), f)


@router.get("/product-list")
def product_list(
    page: int = Query(1, ge=1),
    page_size: int = Query(500, ge=1, le=2000),
    sort_by: str = Query("turnover"),
    sort_dir: str = Query("desc"),
    f: Filters = Depends(_filters),
):
    """Elenco prodotti aggregato per codice, paginato, con i filtri correnti."""
    return q.get_product_list(
        get_engine(), f,
        page=page, page_size=page_size,
        sort_by=sort_by, sort_dir=sort_dir,
    )


# ---------------------------------------------------------------------------
# Export Excel
# ---------------------------------------------------------------------------

@router.get("/export.xlsx")
def export_xlsx(
    view: str = Query("flat", description="flat|split: riflette la vista UI"),
    metric: str = Query("turnover", description="Metrica compatta quando view=split"),
    expanded: Optional[List[str]] = Query(
        None,
        description="Marketplace con colonna espansa (solo view=split). "
                    "Usa '__TOTAL__' per espandere la colonna Totale.",
    ),
    f: Filters = Depends(_filters),
):
    """Export Excel della vista corrente.

    - view=flat  -> workbook multi-sheet standard (riepilogo, marketplace, ecc.)
    - view=split -> singolo foglio pivot con stessa struttura della UI,
                    rispettando la metrica compatta e le colonne espanse.
    """
    if view == "split":
        data = build_xlsx_split(
            get_engine(), f,
            expanded_mktps=set(expanded or []),
            metric=metric,
        )
        filename = (
            f"actual_turnover_split_{f.dal.isoformat()}_{f.al.isoformat()}.xlsx"
        )
    else:
        data = build_xlsx(get_engine(), f)
        filename = f"actual_turnover_{f.dal.isoformat()}_{f.al.isoformat()}.xlsx"

    return Response(
        content=data,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": f'attachment; filename="{filename}"'},
    )


@router.get("/product-list.xlsx")
def export_product_list_xlsx(
    view: str = Query("flat", description="flat|split"),
    metric: str = Query("turnover"),
    expanded: Optional[List[str]] = Query(None),
    sort_by: str = Query("turnover"),
    sort_dir: str = Query("desc"),
    f: Filters = Depends(_filters),
):
    """Export Excel del tab 'Dettaglio prodotti' corrente.

    I filtri di contesto (macrocat/cate2/cate3/marca) viaggiano come filtri
    normali (il frontend li manda come valori singoli al posto del multi-select).
    """
    # ctx per l'intestazione: primo valore se e' un singleton
    def _single(xs):
        return xs[0] if xs and len(xs) == 1 else None

    ctx = {
        "macrocat": _single(f.macrocat),
        "cate2":    _single(f.cate2),
        "cate3":    _single(f.cate3),
        "marca":    _single(f.marca),
    }

    if view == "split":
        data = build_xlsx_product_list_split(
            get_engine(), f,
            ctx=ctx,
            expanded_mktps=set(expanded or []),
            metric=metric,
            sort_by=sort_by, sort_dir=sort_dir,
        )
        filename = (
            f"actual_turnover_products_split_"
            f"{f.dal.isoformat()}_{f.al.isoformat()}.xlsx"
        )
    else:
        data = build_xlsx_product_list(
            get_engine(), f,
            ctx=ctx,
            sort_by=sort_by, sort_dir=sort_dir,
        )
        filename = (
            f"actual_turnover_products_"
            f"{f.dal.isoformat()}_{f.al.isoformat()}.xlsx"
        )

    return Response(
        content=data,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": f'attachment; filename="{filename}"'},
    )
