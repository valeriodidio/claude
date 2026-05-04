"""Importa dati da produzione al MySQL locale per avere dati realistici in dev.

Strategie per tabella:
  pl_prodotti            → importa gli ultimi N snapshot completi (default 1).
                           Ogni snapshot = tutti i prodotti × 6 periodi per una data.
                           NON ha senso limitare per numero di righe.
  resi_impatto_economico → importa gli ultimi N resi per data_rma (default 5000).

Uso:
    python dev/seed_from_prod.py
    python dev/seed_from_prod.py --snapshots 3        # ultimi 3 giorni di PL
    python dev/seed_from_prod.py --resi-rows 10000    # più resi
    python dev/seed_from_prod.py --no-truncate        # append senza svuotare

Legge:
    dev/.env.prod  → credenziali produzione (PROD_DB_*)
    .env           → credenziali locali root (DB_*)
"""
import argparse
import sys
from pathlib import Path

import pandas as pd
from dotenv import dotenv_values
from sqlalchemy import create_engine, text

ROOT = Path(__file__).resolve().parent.parent
PROD_ENV = ROOT / "dev" / ".env.prod"
LOCAL_ENV = ROOT / ".env"


def _url(env: dict, prefix: str) -> str:
    user = env.get(f"{prefix}_USER", "")
    pw   = env.get(f"{prefix}_PASSWORD", "")
    host = env.get(f"{prefix}_HOST", "127.0.0.1")
    port = env.get(f"{prefix}_PORT", "3306")
    name = env.get(f"{prefix}_NAME", "smart2")
    return f"mysql+pymysql://{user}:{pw}@{host}:{port}/{name}?charset=utf8mb4"


def seed_pl_prodotti(eng_prod, eng_local, n_snapshots: int, no_truncate: bool):
    """
    Importa gli ultimi N snapshot completi di pl_prodotti.
    Ogni snapshot contiene tutti i prodotti × tutti i periodi per una data_snapshot.
    """
    print(f"\n→ yeppon_stats.pl_prodotti (ultimi {n_snapshots} snapshot)...")

    # Trova le ultime N date di snapshot disponibili in prod
    sql_dates = text("""
        SELECT DISTINCT data_snapshot
        FROM yeppon_stats.pl_prodotti
        ORDER BY data_snapshot DESC
        LIMIT :n
    """)
    with eng_prod.connect() as c:
        dates = [row[0] for row in c.execute(sql_dates, {"n": n_snapshots}).fetchall()]

    if not dates:
        print("   (tabella vuota in produzione, salto)")
        return

    print(f"   Snapshot da importare: {[str(d) for d in dates]}")

    # Scarica tutte le righe per quelle date
    placeholders = ", ".join(f"'{d}'" for d in dates)
    sql_fetch = text(f"""
        SELECT * FROM yeppon_stats.pl_prodotti
        WHERE data_snapshot IN ({placeholders})
    """)
    print("   Scarico da produzione...", end=" ", flush=True)
    with eng_prod.connect() as c:
        df = pd.read_sql(sql_fetch, c)
    print(f"{len(df):,} righe")

    if df.empty:
        print("   (nessuna riga, salto INSERT)")
        return

    if "id" in df.columns:
        df = df.drop(columns=["id"])

    with eng_local.connect() as c:
        if not no_truncate:
            print("   Svuoto tabella locale...", end=" ", flush=True)
            c.execute(text("SET FOREIGN_KEY_CHECKS=0"))
            c.execute(text("TRUNCATE TABLE yeppon_stats.pl_prodotti"))
            c.execute(text("SET FOREIGN_KEY_CHECKS=1"))
            c.commit()
            print("OK")

    print(f"   Inserisco {len(df):,} righe...", end=" ", flush=True)
    df.to_sql(
        "pl_prodotti", eng_local,
        schema="yeppon_stats", if_exists="append",
        index=False, method="multi", chunksize=500,
    )
    print("OK")


def seed_pl_prodotti_corrieri(eng_prod, eng_local, n_snapshots: int, no_truncate: bool):
    """
    Importa gli ultimi N snapshot di pl_prodotti_corrieri.
    Stessa logica di seed_pl_prodotti: per data_snapshot.
    """
    print(f"\n→ yeppon_stats.pl_prodotti_corrieri (ultimi {n_snapshots} snapshot)...")

    sql_dates = text("""
        SELECT DISTINCT data_snapshot
        FROM yeppon_stats.pl_prodotti_corrieri
        ORDER BY data_snapshot DESC
        LIMIT :n
    """)
    with eng_prod.connect() as c:
        dates = [row[0] for row in c.execute(sql_dates, {"n": n_snapshots}).fetchall()]

    if not dates:
        print("   (tabella vuota in produzione, salto)")
        return

    print(f"   Snapshot da importare: {[str(d) for d in dates]}")

    placeholders = ", ".join(f"'{d}'" for d in dates)
    sql_fetch = text(f"""
        SELECT * FROM yeppon_stats.pl_prodotti_corrieri
        WHERE data_snapshot IN ({placeholders})
    """)
    print("   Scarico da produzione...", end=" ", flush=True)
    with eng_prod.connect() as c:
        df = pd.read_sql(sql_fetch, c)
    print(f"{len(df):,} righe")

    if df.empty:
        print("   (nessuna riga, salto INSERT)")
        return

    if "id" in df.columns:
        df = df.drop(columns=["id"])

    with eng_local.connect() as c:
        if not no_truncate:
            print("   Svuoto tabella locale...", end=" ", flush=True)
            c.execute(text("SET FOREIGN_KEY_CHECKS=0"))
            c.execute(text("TRUNCATE TABLE yeppon_stats.pl_prodotti_corrieri"))
            c.execute(text("SET FOREIGN_KEY_CHECKS=1"))
            c.commit()
            print("OK")

    print(f"   Inserisco {len(df):,} righe...", end=" ", flush=True)
    df.to_sql(
        "pl_prodotti_corrieri", eng_local,
        schema="yeppon_stats", if_exists="append",
        index=False, method="multi", chunksize=500,
    )
    print("OK")


def seed_pl_prodotti_corrieri_dettaglio(eng_prod, eng_local, n_snapshots: int, no_truncate: bool):
    """
    Importa gli ultimi N snapshot di pl_prodotti_corrieri_dettaglio.
    Stessa logica di seed_pl_prodotti: per data_snapshot.
    """
    print(f"\n→ yeppon_stats.pl_prodotti_corrieri_dettaglio (ultimi {n_snapshots} snapshot)...")

    sql_dates = text("""
        SELECT DISTINCT data_snapshot
        FROM yeppon_stats.pl_prodotti_corrieri_dettaglio
        ORDER BY data_snapshot DESC
        LIMIT :n
    """)
    with eng_prod.connect() as c:
        dates = [row[0] for row in c.execute(sql_dates, {"n": n_snapshots}).fetchall()]

    if not dates:
        print("   (tabella vuota in produzione, salto)")
        return

    print(f"   Snapshot da importare: {[str(d) for d in dates]}")

    placeholders = ", ".join(f"'{d}'" for d in dates)
    sql_fetch = text(f"""
        SELECT * FROM yeppon_stats.pl_prodotti_corrieri_dettaglio
        WHERE data_snapshot IN ({placeholders})
    """)
    print("   Scarico da produzione...", end=" ", flush=True)
    with eng_prod.connect() as c:
        df = pd.read_sql(sql_fetch, c)
    print(f"{len(df):,} righe")

    if df.empty:
        print("   (nessuna riga, salto INSERT)")
        return

    if "id" in df.columns:
        df = df.drop(columns=["id"])

    with eng_local.connect() as c:
        if not no_truncate:
            print("   Svuoto tabella locale...", end=" ", flush=True)
            c.execute(text("SET FOREIGN_KEY_CHECKS=0"))
            c.execute(text("TRUNCATE TABLE yeppon_stats.pl_prodotti_corrieri_dettaglio"))
            c.execute(text("SET FOREIGN_KEY_CHECKS=1"))
            c.commit()
            print("OK")

    print(f"   Inserisco {len(df):,} righe...", end=" ", flush=True)
    df.to_sql(
        "pl_prodotti_corrieri_dettaglio", eng_local,
        schema="yeppon_stats", if_exists="append",
        index=False, method="multi", chunksize=500,
    )
    print("OK")


def seed_pl_prodotti_marketplace(eng_prod, eng_local, n_snapshots: int, no_truncate: bool):
    """
    Importa gli ultimi N snapshot di pl_prodotti_marketplace.
    La tabella potrebbe non esistere ancora in produzione: in quel caso salta silenziosamente.
    """
    print(f"\n→ yeppon_stats.pl_prodotti_marketplace (ultimi {n_snapshots} snapshot)...")

    # Verifica che la tabella esista in produzione
    check_sql = text("""
        SELECT COUNT(*) FROM information_schema.tables
        WHERE table_schema = 'yeppon_stats'
          AND table_name   = 'pl_prodotti_marketplace'
    """)
    try:
        with eng_prod.connect() as c:
            exists = c.execute(check_sql).scalar()
    except Exception as e:
        print(f"   (impossibile verificare tabella: {e}, salto)")
        return

    if not exists:
        print("   (tabella non ancora creata in produzione — da popolare con export_pl_prodotti.py)")
        return

    sql_dates = text("""
        SELECT DISTINCT data_snapshot
        FROM yeppon_stats.pl_prodotti_marketplace
        ORDER BY data_snapshot DESC
        LIMIT :n
    """)
    with eng_prod.connect() as c:
        dates = [row[0] for row in c.execute(sql_dates, {"n": n_snapshots}).fetchall()]

    if not dates:
        print("   (tabella vuota in produzione, salto)")
        return

    print(f"   Snapshot da importare: {[str(d) for d in dates]}")

    placeholders = ", ".join(f"'{d}'" for d in dates)
    sql_fetch = text(f"""
        SELECT * FROM yeppon_stats.pl_prodotti_marketplace
        WHERE data_snapshot IN ({placeholders})
    """)
    print("   Scarico da produzione...", end=" ", flush=True)
    with eng_prod.connect() as c:
        df = pd.read_sql(sql_fetch, c)
    print(f"{len(df):,} righe")

    if df.empty:
        print("   (nessuna riga, salto INSERT)")
        return

    if "id" in df.columns:
        df = df.drop(columns=["id"])

    with eng_local.connect() as c:
        if not no_truncate:
            print("   Svuoto tabella locale...", end=" ", flush=True)
            c.execute(text("SET FOREIGN_KEY_CHECKS=0"))
            c.execute(text("TRUNCATE TABLE yeppon_stats.pl_prodotti_marketplace"))
            c.execute(text("SET FOREIGN_KEY_CHECKS=1"))
            c.commit()
            print("OK")

    print(f"   Inserisco {len(df):,} righe...", end=" ", flush=True)
    df.to_sql(
        "pl_prodotti_marketplace", eng_local,
        schema="yeppon_stats", if_exists="append",
        index=False, method="multi", chunksize=500,
    )
    print("OK")


def seed_resi(eng_prod, eng_local, n_rows: int, no_truncate: bool):
    """
    Importa i resi da resi_impatto_economico.
    n_rows=0 → importa tutto (nessun LIMIT).
    """
    label = "tutte le righe" if n_rows == 0 else f"ultime {n_rows:,} righe per data_rma"
    print(f"\n→ yeppon_stats.resi_impatto_economico ({label})...")

    if n_rows == 0:
        sql_fetch = text("""
            SELECT * FROM yeppon_stats.resi_impatto_economico
            ORDER BY data_rma DESC, id_rma DESC
        """)
        params = {}
    else:
        sql_fetch = text("""
            SELECT * FROM yeppon_stats.resi_impatto_economico
            ORDER BY data_rma DESC, id_rma DESC
            LIMIT :n
        """)
        params = {"n": n_rows}

    print("   Scarico da produzione...", end=" ", flush=True)
    with eng_prod.connect() as c:
        df = pd.read_sql(sql_fetch, c, params=params)
    print(f"{len(df):,} righe")

    if df.empty:
        print("   (nessuna riga, salto INSERT)")
        return

    with eng_local.connect() as c:
        if not no_truncate:
            print("   Svuoto tabella locale...", end=" ", flush=True)
            c.execute(text("SET FOREIGN_KEY_CHECKS=0"))
            c.execute(text("TRUNCATE TABLE yeppon_stats.resi_impatto_economico"))
            c.execute(text("SET FOREIGN_KEY_CHECKS=1"))
            c.commit()
            print("OK")

    # Rimuovi colonne presenti in produzione ma non nello schema locale
    cols_to_drop = [c for c in ("id", "data_calcolo") if c in df.columns]
    if cols_to_drop:
        df = df.drop(columns=cols_to_drop)

    print(f"   Inserisco {len(df):,} righe...", end=" ", flush=True)
    df.to_sql(
        "resi_impatto_economico", eng_local,
        schema="yeppon_stats", if_exists="append",
        index=False, method="multi", chunksize=500,
    )
    print("OK")


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument(
        "--snapshots", type=int, default=1,
        help="Numero di snapshot giornalieri PL da importare (default 1 = solo l'ultimo)",
    )
    parser.add_argument(
        "--resi-rows", type=int, default=0,
        help="Numero massimo di resi da importare (default 0 = tutti)",
    )
    parser.add_argument(
        "--no-truncate", action="store_true",
        help="Non svuotare le tabelle locali prima del seed",
    )
    parser.add_argument(
        "--only", choices=["pl", "corrieri", "corrieri_det", "resi", "marketplace"],
        help="Importa solo una delle tabelle",
    )
    args = parser.parse_args()

    if not PROD_ENV.exists():
        print(f"ERRORE: {PROD_ENV} non trovato. Copia da dev/.env.prod.example.")
        sys.exit(1)
    if not LOCAL_ENV.exists():
        print(f"ERRORE: {LOCAL_ENV} non trovato. Copia da .env.example.")
        sys.exit(1)

    prod_env  = dotenv_values(PROD_ENV)
    local_env = dotenv_values(LOCAL_ENV)

    # Connessione produzione (read-only)
    eng_prod = create_engine(
        _url(prod_env, "PROD_DB"),
        connect_args={"connect_timeout": 60},
    )
    # Connessione locale root (per TRUNCATE + INSERT)
    local_env_root = dict(local_env)
    local_env_root["DB_USER"]     = "root"
    local_env_root["DB_PASSWORD"] = local_env.get("DB_ROOT_PASSWORD", "root_dev")
    local_env_root["DB_NAME"]     = "yeppon_stats"
    eng_local = create_engine(
        _url(local_env_root, "DB"),
        connect_args={"connect_timeout": 60},
    )

    if args.only not in ("resi", "corrieri", "corrieri_det", "marketplace"):
        seed_pl_prodotti(eng_prod, eng_local, args.snapshots, args.no_truncate)

    if args.only not in ("pl", "resi", "corrieri_det", "marketplace"):
        seed_pl_prodotti_corrieri(eng_prod, eng_local, args.snapshots, args.no_truncate)

    if args.only not in ("pl", "resi", "corrieri", "marketplace"):
        seed_pl_prodotti_corrieri_dettaglio(eng_prod, eng_local, args.snapshots, args.no_truncate)

    if args.only not in ("pl", "corrieri", "corrieri_det", "marketplace"):
        seed_resi(eng_prod, eng_local, args.resi_rows, args.no_truncate)

    if args.only not in ("pl", "corrieri", "corrieri_det", "resi"):
        seed_pl_prodotti_marketplace(eng_prod, eng_local, args.snapshots, args.no_truncate)

    print("\n=== SEED COMPLETATO ===\n")


if __name__ == "__main__":
    main()
