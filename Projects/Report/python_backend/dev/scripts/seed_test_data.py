"""
Popola la tabella actual_turnover con dati finti ma plausibili per testare il report.

Uso:
    python dev/scripts/seed_test_data.py               # 20.000 righe su 90 giorni
    python dev/scripts/seed_test_data.py --rows 50000 --days 180

Richiede che il container MySQL sia su (docker compose up -d).
"""
from __future__ import annotations

import argparse
import os
import random
import sys
from datetime import datetime, timedelta

from sqlalchemy import create_engine, text


# ---- Cataloghi fittizi --------------------------------------------------------

MARKETPLACES = [
    "Yeppon", "Amazon", "eBay", "ManoMano", "CDiscount",
    "Shopify", "Privalia", "Kelkoo",
]

# distribuzione: Yeppon e Amazon dominano
MKTP_WEIGHTS = [35, 25, 12, 8, 6, 6, 5, 3]

MACROCATS = [
    ("PED", ["Elettrodomestici piccoli", "Clima"],
            {"Elettrodomestici piccoli": ["Macchine caff\u00e8", "Stiro", "Aspirapolvere"],
             "Clima": ["Ventilatori", "Condizionatori portatili"]}),
    ("GED", ["Grandi elettrodomestici"],
            {"Grandi elettrodomestici": ["Lavatrici", "Frigoriferi", "Lavastoviglie"]}),
    ("IT", ["Informatica", "Accessori PC"],
           {"Informatica": ["Notebook", "Desktop", "Monitor"],
            "Accessori PC": ["Tastiere", "Mouse", "Cuffie"]}),
    ("Telefonia", ["Smartphone", "Accessori telefonia"],
                 {"Smartphone": ["Android", "iPhone"],
                  "Accessori telefonia": ["Cover", "Caricatori", "Powerbank"]}),
    ("TV", ["Televisori", "Audio video"],
           {"Televisori": ["OLED", "QLED", "LED"],
            "Audio video": ["Soundbar", "Home theatre"]}),
    ("Foto", ["Fotografia"], {"Fotografia": ["Reflex", "Mirrorless", "Obiettivi"]}),
    ("Videogiochi", ["Console e giochi"],
                   {"Console e giochi": ["PlayStation", "Xbox", "Nintendo"]}),
    ("GiocattoliBaby", ["Giocattoli", "Prima infanzia"],
                      {"Giocattoli": ["LEGO", "Puzzle", "Bambole"],
                       "Prima infanzia": ["Passeggini", "Seggiolini auto"]}),
    ("CasaFaiDaTe", ["Casa", "Fai da te"],
                   {"Casa": ["Arredo giardino", "Cucina"],
                    "Fai da te": ["Utensili elettrici", "Utensili manuali"]}),
    ("AutoBiciMotoNautica", ["Auto", "Bici", "Moto"],
                           {"Auto": ["Accessori auto"], "Bici": ["MTB", "E-bike"],
                            "Moto": ["Caschi", "Abbigliamento moto"]}),
    ("ModaCosmetica", ["Moda", "Cosmetica"],
                     {"Moda": ["Abbigliamento", "Calzature"],
                      "Cosmetica": ["Profumi", "Cura capelli"]}),
]

BRANDS = {
    "PED":       ["DeLonghi", "Philips", "Rowenta", "Braun", "Imetec"],
    "GED":       ["Samsung", "LG", "Bosch", "Whirlpool", "Candy", "Beko"],
    "IT":        ["HP", "Dell", "Lenovo", "Asus", "Acer", "Apple"],
    "Telefonia": ["Apple", "Samsung", "Xiaomi", "Oppo", "Motorola"],
    "TV":        ["Samsung", "LG", "Sony", "Philips", "Panasonic"],
    "Foto":      ["Canon", "Nikon", "Sony", "Fujifilm"],
    "Videogiochi": ["Sony", "Microsoft", "Nintendo", "EA"],
    "GiocattoliBaby": ["LEGO", "Chicco", "Hasbro", "Fisher-Price"],
    "CasaFaiDaTe": ["Bosch", "Black&Decker", "Makita", "Ikea"],
    "AutoBiciMotoNautica": ["Bosch", "Michelin", "Pirelli", "AGV", "Shoei"],
    "ModaCosmetica": ["Nike", "Adidas", "L'Or\u00e9al", "Chanel"],
}

FORNITORI = [f"Fornitore_{i:02d}" for i in range(1, 21)]

PROVINCE = ["MI", "RM", "TO", "NA", "PA", "BO", "FI", "VE", "BA", "CA",
            "GE", "PD", "BS", "VR", "CT", "BG", "MB", "SA", "PR", "MO",
            "TV", "PE", "RE", "LE", "AN", "LI", "CZ", "LT", "TN", "UD"]

USER_TYPES = ["Privato", "Business", "Nuovo", "Fedele"]

SHIPPING_TYPES = ["Standard", "Express", "Ritiro negozio", "Corriere pesante"]


def build_engine(host: str, port: int) -> "sqlalchemy.Engine":
    url = f"mysql+pymysql://root:root_dev@{host}:{port}/yeppon_stats?charset=utf8mb4"
    return create_engine(url, pool_pre_ping=True, future=True)


def random_price(base_min: float, base_max: float) -> tuple[float, float]:
    """Ritorna (turnover_riga, prezzo_acquisto_per_pezzo)."""
    price = round(random.uniform(base_min, base_max), 2)
    # margine random 8-35%, oppure 0 (dropship) nel 15% dei casi
    if random.random() < 0.15:
        acquisto = 0.0
    else:
        margine = random.uniform(0.08, 0.35)
        acquisto = round(price * (1 - margine), 2)
    return price, acquisto


def pick_weighted(items, weights):
    return random.choices(items, weights=weights, k=1)[0]


def generate_rows(n: int, days: int):
    today = datetime.now().replace(hour=0, minute=0, second=0, microsecond=0)
    # preparo cataloghi aggregati
    catalog = []
    for macro, cate2_list, cate3_map in MACROCATS:
        for c2 in cate2_list:
            for c3 in cate3_map.get(c2, [c2]):
                for brand in BRANDS[macro]:
                    for i in range(1, 4):   # 3 prodotti per triplette
                        codice = f"{macro[:3].upper()}-{brand[:3].upper()}-{c3[:3].upper()}-{i:03d}"
                        catalog.append({
                            "macrocat": macro,
                            "cate1":    macro,
                            "cate2":    c2,
                            "cate3":    c3,
                            "marca":    brand,
                            "codice":   codice,
                            "id_p":     abs(hash(codice)) % 900000 + 100,
                            "nome":     f"{brand} {c3} modello {i}",
                            "ean":      "".join(str(random.randint(0, 9)) for _ in range(13)),
                        })

    for i in range(n):
        prod = random.choice(catalog)

        # base price diversa per categoria
        if prod["macrocat"] in ("GED", "TV", "IT"):
            price_range = (200, 1500)
        elif prod["macrocat"] == "Telefonia":
            price_range = (80, 1200)
        elif prod["macrocat"] in ("Foto", "Videogiochi"):
            price_range = (30, 900)
        else:
            price_range = (10, 400)

        quantita = random.choices([1, 1, 1, 1, 2, 2, 3, 5], k=1)[0]
        prezzo_unit, acquisto_unit = random_price(*price_range)
        turnover = round(prezzo_unit * quantita, 2)

        # distribuzione temporale: pi\u00f9 ordini recenti + pattern settimanale
        offset_days = random.choices(
            range(days), weights=[1 + i // 10 for i in range(days)], k=1
        )[0]
        order_date = today - timedelta(days=offset_days)
        order_date = order_date.replace(
            hour=random.randint(8, 22),
            minute=random.randint(0, 59),
            second=random.randint(0, 59),
        )

        yield {
            "marca":           prod["marca"],
            "quantita":        quantita,
            "turnover":        turnover,
            "prezzo_acquisto": acquisto_unit,
            "macrocat":        prod["macrocat"],
            "cate1":           prod["cate1"],
            "cate2":           prod["cate2"],
            "cate3":           prod["cate3"],
            "nomemktp":        pick_weighted(MARKETPLACES, MKTP_WEIGHTS),
            "dataordine":      order_date.strftime("%Y-%m-%d %H:%M:%S"),
            "provincia":       random.choice(PROVINCE),
            "codice":          prod["codice"],
            "id_p":            prod["id_p"],
            "nome":            prod["nome"],
            "mktpordinato":    pick_weighted(MARKETPLACES, MKTP_WEIGHTS),
            "order_id":        random.randint(1000000, 9999999),
            "fornitore":       random.choice(FORNITORI),
            "ean":             prod["ean"],
            "user_type":       random.choice(USER_TYPES),
            "cod_fornitore":   f"FRN{random.randint(100, 999)}",
            "tipo_spedizione": random.choice(SHIPPING_TYPES),
            "importo_spe":     round(random.uniform(0, 15), 2),
        }


INSERT_SQL = """
INSERT INTO actual_turnover (
    marca, quantita, turnover, prezzo_acquisto, macrocat,
    cate1, cate2, cate3, nomemktp, dataordine, provincia,
    codice, id_p, nome, mktpordinato, order_id, fornitore,
    ean, user_type, cod_fornitore, tipo_spedizione, importo_spe
) VALUES (
    :marca, :quantita, :turnover, :prezzo_acquisto, :macrocat,
    :cate1, :cate2, :cate3, :nomemktp, :dataordine, :provincia,
    :codice, :id_p, :nome, :mktpordinato, :order_id, :fornitore,
    :ean, :user_type, :cod_fornitore, :tipo_spedizione, :importo_spe
)
"""


def main():
    p = argparse.ArgumentParser()
    p.add_argument("--rows", type=int, default=20000)
    p.add_argument("--days", type=int, default=90)
    p.add_argument("--host", default=os.getenv("DB_HOST", "127.0.0.1"))
    p.add_argument("--port", type=int, default=int(os.getenv("DB_PORT", "3306")))
    p.add_argument("--truncate", action="store_true",
                   help="svuota la tabella prima di inserire")
    args = p.parse_args()

    random.seed(42)  # riproducibile

    engine = build_engine(args.host, args.port)
    print(f"Connessione a {args.host}:{args.port} ...")

    with engine.begin() as conn:
        if args.truncate:
            print("Svuoto la tabella...")
            conn.execute(text("TRUNCATE TABLE actual_turnover"))

        batch: list[dict] = []
        BATCH_SIZE = 1000
        print(f"Inserisco {args.rows} righe su {args.days} giorni...")
        for i, row in enumerate(generate_rows(args.rows, args.days), 1):
            batch.append(row)
            if len(batch) >= BATCH_SIZE:
                conn.execute(text(INSERT_SQL), batch)
                batch.clear()
                if i % 5000 == 0:
                    print(f"  {i} / {args.rows}")
        if batch:
            conn.execute(text(INSERT_SQL), batch)

    with engine.connect() as conn:
        tot = conn.execute(text("SELECT COUNT(*) FROM actual_turnover")).scalar()
        ft = conn.execute(text("SELECT ROUND(SUM(turnover),2) FROM actual_turnover")).scalar()
        print(f"\nFatto. Righe in tabella: {tot}. Fatturato totale: {ft} \u20ac")


if __name__ == "__main__":
    main()
