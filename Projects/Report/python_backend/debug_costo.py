import sys; sys.path.insert(0, '.')
from datetime import date
from app.db import get_engine
from app.services.turnover_query import _aggregate, Filters
from sqlalchemy import text

eng = get_engine()
# 1) conta righe e somme grezze sul DB che usa l'app
with eng.connect() as c:
    r = c.execute(text("""
        SELECT
          COUNT(*)                                  AS n,
          ROUND(SUM(turnover),2)                    AS fatturato,
          ROUND(SUM(prezzo_acquisto),2)             AS costo_no_qta,
          ROUND(SUM(prezzo_acquisto*quantita),2)    AS costo_x_qta
        FROM actual_turnover
        WHERE dataordine >= '2026-04-01' AND dataordine < '2026-04-23'
    """)).mappings().one()
    print("RAW DB:", dict(r))

# 2) chiama la stessa funzione che chiama l'API
f = Filters(dal=date(2026,4,1), al=date(2026,4,22))
df = _aggregate(eng, ["macrocat"], f)
print("\n_aggregate() output:")
print(df[["macrocat","qta","turnover","costo"]].to_string(index=False))
print("\nTOTALE: turnover=%.2f  costo=%.2f" % (df["turnover"].sum(), df["costo"].sum()))