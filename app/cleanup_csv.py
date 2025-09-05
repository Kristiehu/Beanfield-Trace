
# cleanup_csv.py — generic work-order CSV cleaner (normalize + integrity checks)
from __future__ import annotations
import io, re
from dataclasses import dataclass
from typing import Dict, Any, Optional
import pandas as pd

@dataclass
class CsvCleanResult:
    cleaned_df: pd.DataFrame
    integrity: Dict[str, Any]

def run_clean_csv(raw_csv_bytes: bytes, encoding: str = "utf-8") -> CsvCleanResult:
    # Read everything as string to avoid dtype loss, then coerce selected columns later
    df = pd.read_csv(io.BytesIO(raw_csv_bytes), dtype=str, keep_default_na=False, encoding=encoding)

    # Strip whitespace in headers
    df.columns = [c.strip() for c in df.columns]

    # Trim cells & normalize internal spaces
    for col in df.columns:
        df[col] = df[col].astype(str).str.replace(r"\s+", " ", regex=True).str.strip()

    # Drop fully-empty rows
    df = df[~(df.apply(lambda r: "".join(r.values.astype(str)).strip() == "", axis=1))].copy()

    # Drop duplicate rows (exact match across all columns)
    before = len(df)
    df = df.drop_duplicates().reset_index(drop=True)
    dups_dropped = before - len(df)

    # Heuristic coercions (if present)
    num_cols = ["PMID", "Fibre Count", "Fibres", "End to End Length(m)"]
    for nc in num_cols:
        if nc in df.columns:
            df[nc] = pd.to_numeric(df[nc].str.extract(r"([\d.]+)")[0], errors="coerce")

    # Integrity report
    integrity = {
        "rows": int(len(df)),
        "cols": int(df.shape[1]),
        "duplicates_removed": int(dups_dropped),
        "empty_cells": int((df == "").sum().sum()),
        "columns": list(df.columns),
    }

    return CsvCleanResult(df, integrity)

def export_csv_bytes(result: CsvCleanResult, basename: str = "work_order") -> Dict[str, bytes]:
    artifacts: Dict[str, bytes] = {}
    # Cleaned CSV
    bio = io.StringIO()
    result.cleaned_df.to_csv(bio, index=False)
    artifacts[f"{basename}.clean.csv"] = bio.getvalue().encode("utf-8")
    # Integrity report (txt)
    rep = io.StringIO()
    rep.write("CSV Integrity Report\n")
    for k, v in result.integrity.items():
        rep.write(f"{k}: {v}\n")
    artifacts[f"{basename}.integrity.txt"] = rep.getvalue().encode("utf-8")
    return artifacts
