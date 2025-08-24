# remove_add_algo.py
import re
import pandas as pd
from typing import List, Tuple

def _norm(s) -> str:
    s = "" if pd.isna(s) else str(s)
    s = s.strip()
    # drop trailing colon(s)
    s = re.sub(r":+\s*$", "", s)
    # collapse internal whitespace
    s = re.sub(r"\s{2,}", " ", s)
    return s

def _flatten_label_value(df: pd.DataFrame, pairs: List[Tuple[int, int]]) -> pd.DataFrame:
    """
    Turn a grid like your de.xlsx (label/value in neighboring columns)
    into a 2-column table: Field, Value.
    pairs = [(0,1), (2,3)] means: col0->label, col1->value, col2->label, col3->value
    """
    rows = []
    n = len(df)
    for a, b in pairs:
        if a >= df.shape[1]:
            continue
        # value col b is optional
        has_b = b < df.shape[1]
        for i in range(n):
            lab = _norm(df.iat[i, a]) if a < df.shape[1] else ""
            val = _norm(df.iat[i, b]) if has_b else ""
            if lab and lab.lower() != "nan":
                rows.append((lab, val))
    out = pd.DataFrame(rows, columns=["Field", "Value"])
    # remove blank/duplicate rows while keeping first occurrence
    out = out[out["Field"].astype(str).str.len() > 0].copy()
    out = out.drop_duplicates(subset=["Field"], keep="first").reset_index(drop=True)
    return out

def transform_remove_add(df_in: pd.DataFrame) -> pd.DataFrame:
    """
    MAIN: your Remove & Add algorithm.
    Currently:
      - Detects label/value layout like de.xlsx
      - Flattens to two columns: Field, Value
    Customize the 'EXPECTED_COLS' section below if you need a different final schema.
    """
    df = df_in.copy()

    # If it's an Excel-exported “report” grid like your de.xlsx (5 columns),
    # pair (0,1) and (2,3). Otherwise, try a best-effort guess:
    if df.shape[1] >= 4:
        out = _flatten_label_value(df, pairs=[(0, 1), (2, 3)])
    elif df.shape[1] == 2:
        out = df.copy()
        out.columns = ["Field", "Value"]
        out["Field"] = out["Field"].map(_norm)
        out["Value"] = out["Value"].map(_norm)
    else:
        # Fallback: take first column as Field, second as Value (if present)
        cols = list(df.columns)
        if len(cols) == 1:
            out = pd.DataFrame({"Field": df.iloc[:, 0].map(_norm), "Value": ""})
        else:
            out = pd.DataFrame({"Field": df.iloc[:, 0].map(_norm), "Value": df.iloc[:, 1].map(_norm)})

    return out
