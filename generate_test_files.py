"""
One-off script to generate two synthetic test files (test800lignes.xlsx and
test850lignes.xlsx) with the same column structure as fichier 1.xlsx, built to
exercise the comparator: shared/unchanged rows, modified rows, deleted rows
(present only in the 800 file) and new rows (present only in the 850 file).
"""
import random
from datetime import date, timedelta

import pandas as pd

random.seed(42)

SRC_COLUMNS = list(pd.read_excel("fichier 1.xlsx", engine="openpyxl").columns)
STAKE_COL = SRC_COLUMNS[5]  # "Stake for Thales (M€)" — reuse exact original bytes

COUNTRIES = ["France", "Angleterre", "Espagne", "Japon", "Allemagne", "Italie",
             "USA", "Canada", "Belgique", "Pays-Bas", "Suisse", "Portugal"]
ENTITIES = ["Thales USA", "TAS France", "Thales LAS France", "Thales CDI France",
            "Thales AVS France", "Thales DIS France", "Thales Six GTS France",
            "Thales Nederland", "Thales UK", "Thales Deutschland"]
GBU_BL = ["IFE", "OEN", "LAS", "CDI", "AVS", "DIS", "GTS", "SIX"]
CUSTOMERS = ["Air France", "Unseenlab", "Etat", "Visa", "Lufthansa", "SNCF",
             "British Airways", "Deutsche Bahn", "Ministere de la Defense",
             "Renault", "Airbus", "EDF", "Orange", "NATO", "Emirates"]
SALES_STAGES = ["ABC", "DEF", "GHI", "JKL", "MNO", "PQR", "STU"]
COMPETITORS = [None, None, None, "Leonardo", "Raytheon", "Lockheed Martin",
               "BAE Systems", "Airbus Defence"]

BASE_DATE = date(2022, 1, 1)


def random_row(opp_id: int) -> dict:
    return {
        "Opportunity": opp_id,
        "Country": random.choice(COUNTRIES),
        "Opportunity Name": f"Opportunity {opp_id}",
        "Operating Unit/Legal Entity": random.choice(ENTITIES),
        "GBU/BL": random.choice(GBU_BL),
        STAKE_COL: random.randint(1, 50),
        "Direct customer": random.choice(CUSTOMERS),
        "End user": random.choice(CUSTOMERS),
        "Booking date": BASE_DATE + timedelta(days=random.randint(0, 1460)),
        "Sales Stage": random.choice(SALES_STAGES),
        "Products": f"Produit {random.randint(1, 30)}",
        "Competitors": random.choice(COMPETITORS),
    }


def modify_row(row: dict) -> dict:
    """Return a copy of row with 1-3 fields changed (simulates a quarter update)."""
    new_row = dict(row)
    fields = random.sample(
        [STAKE_COL, "Sales Stage", "Products", "Competitors", "Country", "End user"],
        k=random.randint(1, 3),
    )
    for f in fields:
        if f == STAKE_COL:
            new_row[f] = max(1, row[f] + random.choice([-5, -2, -1, 1, 2, 5, 10]))
        elif f == "Sales Stage":
            new_row[f] = random.choice(SALES_STAGES)
        elif f == "Products":
            new_row[f] = f"Produit {random.randint(1, 30)}"
        elif f == "Competitors":
            new_row[f] = random.choice(COMPETITORS)
        elif f == "Country":
            new_row[f] = random.choice(COUNTRIES)
        elif f == "End user":
            new_row[f] = random.choice(CUSTOMERS)
    return new_row


# ── Build the shared population ────────────────────────────────────────────
COMMON_N = 750       # present in both files
ONLY_800_N = 50      # present only in the 800-row file  -> "deleted" in comparison
ONLY_850_N = 100     # present only in the 850-row file  -> "new" in comparison

common_ids = list(range(1, COMMON_N + 1))
only_800_ids = list(range(COMMON_N + 1, COMMON_N + 1 + ONLY_800_N))          # 751-800
only_850_ids = list(range(COMMON_N + ONLY_800_N + 1, COMMON_N + ONLY_800_N + 1 + ONLY_850_N))  # 801-900

common_rows = {i: random_row(i) for i in common_ids}

# 800-row file: common rows (baseline) + rows only present in this file
rows_800 = [common_rows[i] for i in common_ids] + [random_row(i) for i in only_800_ids]

# 850-row file: common rows (~20% modified) + new rows only present in this file
MODIFIED_RATIO = 0.20
modified_ids = set(random.sample(common_ids, int(COMMON_N * MODIFIED_RATIO)))

rows_850 = []
for i in common_ids:
    base = common_rows[i]
    rows_850.append(modify_row(base) if i in modified_ids else dict(base))
rows_850 += [random_row(i) for i in only_850_ids]

df_800 = pd.DataFrame(rows_800, columns=SRC_COLUMNS)
df_850 = pd.DataFrame(rows_850, columns=SRC_COLUMNS)

assert len(df_800) == 800, len(df_800)
assert len(df_850) == 850, len(df_850)

df_800.to_excel("test800lignes.xlsx", index=False, engine="openpyxl")
df_850.to_excel("test850lignes.xlsx", index=False, engine="openpyxl")

print(f"test800lignes.xlsx : {len(df_800)} lignes")
print(f"test850lignes.xlsx : {len(df_850)} lignes")
print(f"  - communes            : {COMMON_N} (dont {len(modified_ids)} modifiees)")
print(f"  - supprimees (800->850): {ONLY_800_N}")
print(f"  - nouvelles (800->850) : {ONLY_850_N}")
