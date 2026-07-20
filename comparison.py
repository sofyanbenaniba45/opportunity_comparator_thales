"""
Logique de comparaison entre deux DataFrames d'opportunités commerciales.
"""
import pandas as pd
import numpy as np


KEY_COL = "Opportunity Name"


def normalize_key(series: pd.Series) -> pd.Series:
    """Normalise les clés pour la comparaison (strip + lowercase)."""
    return series.astype(str).str.strip().str.lower()


def values_differ(v1, v2) -> bool:
    """Retourne True si deux valeurs sont différentes, en gérant NaN/None."""
    # Les deux sont NaN/None/vide → identiques
    def is_empty(v):
        if v is None:
            return True
        if isinstance(v, float) and np.isnan(v):
            return True
        if str(v).strip() in ("", "nan", "None"):
            return True
        return False

    if is_empty(v1) and is_empty(v2):
        return False
    if is_empty(v1) or is_empty(v2):
        return True
    return str(v1).strip() != str(v2).strip()


def compare_dataframes(df_prev: pd.DataFrame, df_curr: pd.DataFrame) -> dict:
    """
    Compare deux DataFrames sur la base de la colonne KEY_COL.

    Retourne un dict avec :
      - 'new'      : DataFrame des lignes nouvelles (dans curr, pas dans prev)
      - 'deleted'  : DataFrame des lignes supprimées (dans prev, pas dans curr)
      - 'modified' : DataFrame des lignes modifiées avec colonnes supplémentaires
      - 'unchanged': DataFrame des lignes inchangées
      - 'diff_map' : dict {opportunity_name_normalized: set(colonnes_modifiées)}
    """
    # Copie pour ne pas altérer les originaux
    prev = df_prev.copy()
    curr = df_curr.copy()

    # Clés normalisées pour la comparaison
    prev["__key__"] = normalize_key(prev[KEY_COL])
    curr["__key__"] = normalize_key(curr[KEY_COL])

    prev_keys = set(prev["__key__"])
    curr_keys = set(curr["__key__"])

    new_keys = curr_keys - prev_keys
    deleted_keys = prev_keys - curr_keys
    common_keys = prev_keys & curr_keys

    # Indexer par clé normalisée pour lookup rapide
    prev_indexed = prev.set_index("__key__")
    curr_indexed = curr.set_index("__key__")

    columns = [c for c in df_curr.columns if c != "__key__"]

    modified_rows = []
    unchanged_rows = []
    diff_map = {}  # key → set of column names that differ

    for key in common_keys:
        row_prev = prev_indexed.loc[key]
        row_curr = curr_indexed.loc[key]

        changed_cols = set()
        for col in columns:
            v_prev = row_prev[col] if col in row_prev.index else None
            v_curr = row_curr[col] if col in row_curr.index else None
            if values_differ(v_prev, v_curr):
                changed_cols.add(col)

        if changed_cols:
            diff_map[key] = changed_cols
            row_data = row_curr[columns].copy()
            row_data = row_data.to_frame().T
            row_data["__status__"] = "modified"
            modified_rows.append(row_data)
        else:
            row_data = row_curr[columns].copy()
            row_data = row_data.to_frame().T
            row_data["__status__"] = "unchanged"
            unchanged_rows.append(row_data)

    def build_df(keys, source_indexed, status):
        if not keys:
            return pd.DataFrame(columns=columns + ["__status__"])
        rows = source_indexed.loc[list(keys)].reindex(columns=columns).copy()
        rows["__status__"] = status
        return rows.reset_index(drop=True)

    df_new = build_df(new_keys, curr_indexed, "new")
    df_deleted = build_df(deleted_keys, prev_indexed, "deleted")

    df_modified = (
        pd.concat(modified_rows, ignore_index=True) if modified_rows
        else pd.DataFrame(columns=columns + ["__status__"])
    )
    df_unchanged = (
        pd.concat(unchanged_rows, ignore_index=True) if unchanged_rows
        else pd.DataFrame(columns=columns + ["__status__"])
    )

    return {
        "new": df_new,
        "deleted": df_deleted,
        "modified": df_modified,
        "unchanged": df_unchanged,
        "diff_map": diff_map,
        "columns": columns,
        "prev_indexed": prev_indexed,  # pour récupérer les anciennes valeurs
    }
