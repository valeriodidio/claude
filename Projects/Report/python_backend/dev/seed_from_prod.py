"""
seed_from_prod.py
-----------------
Importa le ultime N righe da actual_turnover del DB di produzione
nel MySQL locale (Docker), più i record ordini_cliente corrispondenti
(necessari per il riconoscimento degli ordini H&A).

Uso:
    python dev/seed_from_prod.py
    python dev/seed_from_prod.py --rows 50000   # quante righe importare
    python dev/seed_from_prod.py --no-truncate  # append invece di svuotare prima

Legge le credenziali di produzione da dev/.env.prod
(non committare mai questo file nel repo).
"""

import argparse
import os
import sys
import time
from pathlib import Path

# ---------------------------------------------------------------------------
# Dipendenze: pymysql e python-dotenv devono essere nel venv
# ---------------------------------------------------------------------------
try:
    import pymysql
    from dotenv import dotenv_values
except ImportError as e:
    print(f"\nERRORE import: {e}")
    print("Assicurati di aver attivato il venv e installato le dipendenze:")
    print("    venv\\Scripts\\activate")
    print("    pip install -r requirements.txt")
    sys.exit(1)

# ---------------------------------------------------------------------------
# Configurazione
# ---------------------------------------------------------------------------
SCRIPT_DIR = Path(__file__).resolve().parent
ENV_PROD    = SCRIPT_DIR / ".env.prod"
ENV_LOCAL   = SCRIPT_DIR.parent / ".env"

BATCH_SIZE  = 500          # righe per INSERT batch

# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

def load_env(path: Path) -> dict:
    if not path.exists():
        return {}
    return dotenv_values(str(path))


def connect(host, port, user, password, db, label="DB") -> pymysql.Connection:
    print(f"   Connessione a {label} ({host}:{port}/{db})...", end=" ", flush=True)
    try:
        conn = pymysql.connect(
            host=host,
            port=int(port),
            user=user,
            password=password,
            database=db,
            charset="utf8mb4",
            connect_timeout=10,
        )
        print("OK")
        return conn
    except Exception as e:
        print(f"ERRORE\n   {e}")
        sys.exit(1)


def fetch_columns(conn, table: str) -> list[str]:
    with conn.cursor() as cur:
        cur.execute(f"SHOW COLUMNS FROM `{table}`")
        return [row[0] for row in cur.fetchall()]


def main():
    parser = argparse.ArgumentParser(description="Seed DB locale da produzione")
    parser.add_argument("--rows",        type=int,  default=20000, help="Numero di righe da importare (default: 20000)")
    parser.add_argument("--no-truncate", action="store_true",      help="Append invece di svuotare la tabella locale prima")
    args = parser.parse_args()

    # ------------------------------------------------------------------
    # 1. Carica config
    # ------------------------------------------------------------------
    prod_env  = load_env(ENV_PROD)
    local_env = load_env(ENV_LOCAL)

    if not prod_env:
        print(f"\nERRORE: file non trovato: {ENV_PROD}")
        print(f"Copialo dall'esempio: {ENV_PROD.with_suffix('.example')}")
        print("e compila i valori di produzione.")
        sys.exit(1)

    missing = [k for k in ("PROD_DB_HOST", "PROD_DB_USER", "PROD_DB_PASSWORD", "PROD_DB_NAME")
               if not prod_env.get(k)]
    if missing:
        print(f"\nERRORE: campi mancanti in {ENV_PROD}: {', '.join(missing)}")
        sys.exit(1)

    # ------------------------------------------------------------------
    # 2. Connessioni
    # ------------------------------------------------------------------
    print()
    print("=== SEED da produzione ===")

    prod_conn = connect(
        host     = prod_env["PROD_DB_HOST"],
        port     = prod_env.get("PROD_DB_PORT", 3306),
        user     = prod_env["PROD_DB_USER"],
        password = prod_env["PROD_DB_PASSWORD"],
        db       = prod_env.get("PROD_DB_NAME", "yeppon_stats"),
        label    = "PRODUZIONE",
    )

    # Usa le credenziali root per il DB locale: reports_reader ha solo SELECT
    # e non puo' fare TRUNCATE/INSERT. Le credenziali root sono nel .env locale
    # (LOCAL_DB_ROOT_USER / LOCAL_DB_ROOT_PASSWORD) e valgono solo per Docker dev.
    local_conn = connect(
        host     = local_env.get("DB_HOST", "127.0.0.1"),
        port     = local_env.get("DB_PORT", 3306),
        user     = local_env.get("LOCAL_DB_ROOT_USER", "root"),
        password = local_env.get("LOCAL_DB_ROOT_PASSWORD", "root_dev"),
        db       = local_env.get("DB_NAME", "yeppon_stats"),
        label    = "LOCALE (Docker, root)",
    )

    # ------------------------------------------------------------------
    # 2b. Assicura che lo schema locale sia allineato a init.sql.
    # Se il volume Docker e' stato creato con una versione vecchia dell'init.sql
    # (senza il database `smart2` o senza i GRANT per usr_report), MySQL non
    # rilancia mai l'init: e' responsabilita' del seed metterlo a posto.
    # Gli statement sono idempotenti: non fanno nulla se lo schema e' gia' ok.
    # ------------------------------------------------------------------
    usr_report_pwd = local_env.get("DB_PASSWORD", "dev_password_change_me")
    with local_conn.cursor() as cur:
        cur.execute("CREATE DATABASE IF NOT EXISTS smart2 "
                    "DEFAULT CHARACTER SET utf8mb4 COLLATE utf8mb4_unicode_ci")
        cur.execute("""
            CREATE TABLE IF NOT EXISTS smart2.ordini_cliente (
                id        INT PRIMARY KEY,
                idUtente  INT NOT NULL,
                INDEX idx_id_utente (idUtente)
            )
        """)
        cur.execute(
            f"CREATE USER IF NOT EXISTS 'usr_report'@'%' IDENTIFIED BY '{usr_report_pwd}'"
        )
        cur.execute("GRANT SELECT ON yeppon_stats.actual_turnover TO 'usr_report'@'%'")
        cur.execute("GRANT SELECT ON smart2.ordini_cliente         TO 'usr_report'@'%'")
        cur.execute("FLUSH PRIVILEGES")
    local_conn.commit()

    # ------------------------------------------------------------------
    # 3. Lettura colonne actual_turnover (usa quelle che esistono su entrambi i lati)
    # ------------------------------------------------------------------
    prod_cols  = fetch_columns(prod_conn, "actual_turnover")
    local_cols = fetch_columns(local_conn, "actual_turnover")
    cols = [c for c in prod_cols if c != "id" and c in local_cols]

    cols_sql     = ", ".join(f"`{c}`" for c in cols)
    placeholders = ", ".join(["%s"] * len(cols))

    # ------------------------------------------------------------------
    # 4. Fetch actual_turnover da produzione
    # ------------------------------------------------------------------
    print(f"   Leggo le ultime {args.rows:,} righe da actual_turnover...", end=" ", flush=True)
    t0 = time.time()
    with prod_conn.cursor() as cur:
        cur.execute(
            f"SELECT {cols_sql} FROM actual_turnover "
            f"ORDER BY dataordine DESC, id DESC "
            f"LIMIT %s",
            (args.rows,)
        )
        rows = cur.fetchall()
    print(f"{len(rows):,} righe lette in {time.time()-t0:.1f}s")

    if not rows:
        print("   Nessuna riga trovata in produzione. Controlla le credenziali.")
        prod_conn.close()
        local_conn.close()
        sys.exit(1)

    # Ricava i order_id univoci per il seed di ordini_cliente
    order_id_idx = cols.index("order_id") if "order_id" in cols else None
    order_ids: list[int] = []
    if order_id_idx is not None:
        order_ids = list({r[order_id_idx] for r in rows if r[order_id_idx] is not None})

    # ------------------------------------------------------------------
    # 4b. Fetch ordini_cliente corrispondenti (per H&A join)
    # ------------------------------------------------------------------
    oc_rows: list[tuple] = []
    if order_ids:
        print(f"   Leggo ordini_cliente per {len(order_ids):,} order_id...", end=" ", flush=True)
        t1 = time.time()
        # verifica che la tabella esista in produzione
        try:
            with prod_conn.cursor() as _cur:
                _cur.execute("SELECT 1 FROM smart2.ordini_cliente LIMIT 1")
            table_exists = True
        except Exception:
            table_exists = False

        if table_exists:
            CHUNK = 1000
            for i in range(0, len(order_ids), CHUNK):
                chunk = order_ids[i : i + CHUNK]
                ph = ", ".join(["%s"] * len(chunk))
                with prod_conn.cursor() as cur:
                    cur.execute(
                        f"SELECT id, idUtente FROM smart2.ordini_cliente WHERE id IN ({ph})",
                        chunk,
                    )
                    oc_rows.extend(cur.fetchall())
            print(f"{len(oc_rows):,} righe in {time.time()-t1:.1f}s")
        else:
            print("tabella ordini_cliente non trovata in produzione, skip.")

    prod_conn.close()

    # ------------------------------------------------------------------
    # 5. Import nel DB locale
    # ------------------------------------------------------------------
    with local_conn.cursor() as cur:
        if not args.no_truncate:
            print("   Svuoto tabelle locali...", end=" ", flush=True)
            cur.execute("SET FOREIGN_KEY_CHECKS=0")
            cur.execute("TRUNCATE TABLE yeppon_stats.actual_turnover")
            cur.execute("TRUNCATE TABLE smart2.ordini_cliente")
            cur.execute("SET FOREIGN_KEY_CHECKS=1")
            local_conn.commit()
            print("OK")

        # -- actual_turnover --
        total    = len(rows)
        inserted = 0
        print(f"   Inserisco {total:,} righe in actual_turnover (batch {BATCH_SIZE})...")
        t0 = time.time()
        for i in range(0, total, BATCH_SIZE):
            batch = rows[i : i + BATCH_SIZE]
            cur.executemany(
                f"INSERT INTO yeppon_stats.actual_turnover ({cols_sql}) VALUES ({placeholders})",
                batch,
            )
            local_conn.commit()
            inserted += len(batch)
            pct = inserted / total * 100
            bar = "#" * int(pct / 5)
            print(f"\r   [{bar:<20}] {pct:5.1f}%  {inserted:,}/{total:,}  ({time.time()-t0:.0f}s)", end="", flush=True)
        print()

        # -- ordini_cliente --
        if oc_rows:
            print(f"   Inserisco {len(oc_rows):,} righe in ordini_cliente...", end=" ", flush=True)
            t1 = time.time()
            for i in range(0, len(oc_rows), BATCH_SIZE):
                batch = oc_rows[i : i + BATCH_SIZE]
                cur.executemany(
                    "INSERT IGNORE INTO smart2.ordini_cliente (id, idUtente) VALUES (%s, %s)",
                    batch,
                )
            local_conn.commit()
            print(f"OK ({time.time()-t1:.1f}s)")

    local_conn.close()
    print(f"\n=== SEED OK: {inserted:,} righe actual_turnover, {len(oc_rows):,} ordini_cliente ===\n")


if __name__ == "__main__":
    main()
