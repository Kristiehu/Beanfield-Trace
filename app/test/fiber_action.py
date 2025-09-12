# fiber_action.py (updated)
from __future__ import annotations

import io
import json
import re
from typing import Dict, IO, Iterable, List, Optional, Set, Tuple, Union

import pandas as pd

# --------------------------- Regexes ---------------------------
_ID = r"[A-Za-z0-9#._/-]+"  # supports 40084B, 5116-1117803, 50986_A, etc.

# Matches: "Splice 41121 [73-84] - 41122 [73-84]" or "BREAK ..."
PAIR_LINE_RE = re.compile(
    rf'(?P<verb>Splice|BREAK)\s+'
    rf'(?P<A>{_ID})\s*\[(?P<Ar>[^\]]+)\]\s*-\s*'
    rf'(?P<B>{_ID})\s*\[(?P<Br>[^\]]+)\]',
    re.IGNORECASE,
)

# Optional tail like " M#248 - M#248" after a closing bracket/paren
LOC_TAIL_RE = re.compile(r"\]\s*([^\[\n]*?)$")  # capture suffix after last ']'
BOX_TAIL_RE = re.compile(r"\b([A-Za-z0-9#._/-]{2,})\s*-\s*([A-Za-z0-9#._/-]{2,})\s*$")

# --------------------------- JSON parsing ---------------------------

def _extract_connection_blocks(b: Optional[bytes]) -> List[str]:
    """Return every 'Connections' string blob from the WO JSON export."""
    if not b:
        return []
    try:
        data = json.loads(b.decode("utf-8", "ignore"))
    except Exception:
        return []

    blocks: List[str] = []
    # Primary layout seen in examples
    try:
        groups = data.get("Report: Splice Details", [])
        if groups:
            # Usually a list whose first element has key "" with a list of dicts
            for blk in groups[0].get("", []):
                s = blk.get("Connections")
                if isinstance(s, str):
                    blocks.append(s)
    except Exception:
        pass

    # Fallback: walk the whole JSON and pull any long strings that mention CA and PMID
    if not blocks:
        def walk(obj):
            if isinstance(obj, dict):
                for v in obj.values():
                    yield from walk(v)
            elif isinstance(obj, list):
                for v in obj:
                    yield from walk(v)
            elif isinstance(obj, str):
                if "PMID" in obj and "CA" in obj:
                    yield obj
        blocks = list(walk(data))

    return blocks


_HEADER_RE = re.compile(r"^\s*([A-Za-z0-9#._/-]{2,})\s*:\s*OSP Splice Box\b.*$", re.IGNORECASE)
_CA_PMD_RE = re.compile(r"^\s*CA\d+\s*[^A-Za-z0-9]*PMID[^A-Za-z0-9]*\s*([0-9A-Z]+)\b")
_PMID_PAIR_RE = re.compile(
    r"^\s*PMID[:]\s*([0-9A-Z]+)\b.*?--\s*Splice\s*--\s*PMID[:]\s*([0-9A-Z]+)\b",
    re.IGNORECASE,
)

def _ca_order_and_pairs(b: Optional[bytes]) -> Tuple[Dict[str, Dict[str, int]], Dict[frozenset, Tuple[str, str]]]:
    """
    Parse JSON 'Connections' strings and build:
      - ca_order_by_box: { box_name -> { pmid -> order_index } } using CA1, CA2, ...
      - pair_orient: { frozenset({A,B}) -> (A,B) } using 'PMID: A -- Splice -- PMID: B' (first seen wins)
    """
    ca_order_by_box: Dict[str, Dict[str, int]] = {}
    pair_orient: Dict[frozenset, Tuple[str, str]] = {}

    for blob in _extract_connection_blocks(b):
        current_box: Optional[str] = None
        order_map: Dict[str, int] = {}

        for line in blob.splitlines():
            m_head = _HEADER_RE.match(line)
            if m_head:
                # Commit previous box (if any)
                if current_box and order_map:
                    ca_order_by_box[current_box] = dict(order_map)
                current_box = m_head.group(1).strip()
                order_map = {}
                continue

            m_ca = _CA_PMD_RE.match(line)
            if m_ca and current_box:
                pmid = m_ca.group(1).strip()
                if pmid not in order_map:
                    order_map[pmid] = len(order_map)  # 0-based CA ordering
                continue

            m_pair = _PMID_PAIR_RE.match(line)
            if m_pair:
                A, B = m_pair.group(1), m_pair.group(2)
                key = frozenset((A, B))
                if key not in pair_orient:
                    pair_orient[key] = (A, B)

        # flush last
        if current_box and order_map and current_box not in ca_order_by_box:
            ca_order_by_box[current_box] = dict(order_map)

    return ca_order_by_box, pair_orient


# --------------------------- Core logic ---------------------------

def _determine_desired_order(
    desc: str,
    A: str, Ar: str, B: str, Br: str,
    ca_order_by_box: Dict[str, Dict[str, int]],
    pair_orient: Dict[frozenset, Tuple[str, str]],
) -> Tuple[str, str]:
    """
    Choose (left_id, right_id) for the pair inside `desc`:
      1) If description ends with 'BoxA - BoxB' and BoxA == BoxB and both IDs appear
         in that box's CA list, respect the CA order (lower index comes first).
      2) Else, if we have an explicit 'PMID: A -- Splice -- PMID: B' from JSON, use it.
      3) Otherwise, keep original (A,B).
    """
    # Try to extract a " ... M#248 - M#248" tail
    tail = (LOC_TAIL_RE.search(desc) or (None,)).group(1) if LOC_TAIL_RE.search(desc) else ""
    box_hint: Optional[str] = None
    if tail:
        mbox = BOX_TAIL_RE.search(tail)
        if mbox:
            boxA, boxB = mbox.group(1).strip(), mbox.group(2).strip()
            if boxA == boxB:
                box_hint = boxA

    if box_hint and box_hint in ca_order_by_box:
        idx = ca_order_by_box[box_hint]
        if A in idx and B in idx:
            return (A, B) if idx[A] <= idx[B] else (B, A)

    key = frozenset((A, B))
    if key in pair_orient:
        natural = pair_orient[key]
        # Ensure we return the explicit (A,B) ordering from JSON
        return natural

    return (A, B)


def _reorder_pairs_in_df(
    df: pd.DataFrame,
    ca_order_by_box: Dict[str, Dict[str, int]],
    pair_orient: Dict[frozenset, Tuple[str, str]],
) -> pd.DataFrame:
    """Swap A/B inside text to match desired orientation; preserve ranges and tails."""
    if df.empty:
        return df

    out = df.copy()
    for idx, row in out.iterrows():
        desc = str(row.get("Description", "") or "")
        m = PAIR_LINE_RE.search(desc)
        if not m:
            continue
        verb, A, Ar, B, Br = m.group("verb", "A", "Ar", "B", "Br")
        left, right = _determine_desired_order(desc, A, Ar, B, Br, ca_order_by_box, pair_orient)
        if (A, B) == (left, right):
            continue  # already correct
        # rewrite keeping original ranges bound to their original IDs
        swapped = f"{verb} {left}[{Ar if left==A else Br}] - {right}[{Br if right==B else Ar}]"
        new_desc = desc[:m.start()] + swapped + desc[m.end():]
        out.at[idx, "Description"] = new_desc
    return out


def _read_actions_table(upload: Union[str, IO[bytes], bytes]) -> pd.DataFrame:
    """Read the WO CSV (with or without header)."""
    if isinstance(upload, bytes):
        buf = io.BytesIO(upload)
    elif hasattr(upload, "read"):
        upload.seek(0)
        buf = upload
    else:
        buf = upload  # path or file-like acceptable by pandas
    df = pd.read_csv(buf, header=None, dtype=str, keep_default_na=False)
    # Try to normalize to columns: Action, Description, SAP (if present), others untouched
    cols = list(df.columns)
    name_map = {}
    if len(cols) >= 1: name_map[cols[0]] = "Action"
    if len(cols) >= 2: name_map[cols[1]] = "Description"
    if len(cols) >= 3: name_map[cols[2]] = "SAP"
    df = df.rename(columns=name_map)
    return df


def _filter_action_rows(df: pd.DataFrame) -> pd.DataFrame:
    """Keep only rows relevant to splicing/break actions."""
    if df.empty:
        return df
    # Heuristics: rows whose Description contains "Splice" or "BREAK"
    mask = df["Description"].str.contains(r"\b(Splice|BREAK)\b", case=False, regex=True)
    return df.loc[mask].reset_index(drop=True)


def compute_fibre_actions(
    csv_source: Union[str, IO[bytes], bytes],
    json_source: Optional[Union[str, IO[bytes], bytes]] = None,
    *,
    enrich_from_json: bool = True,
) -> pd.DataFrame:
    """
    Read the WO CSV and (optionally) JSON, and reorder splice pairs so the left→right
    matches the natural CA order inside the splice box block (CA1, CA2, CA3…).
    Falls back to the explicit 'PMID: A -- Splice -- PMID: B' lines if needed.

    Returns a DataFrame with at least ['Action','Description', ...] columns.
    """
    df = _read_actions_table(csv_source)
    if "Description" not in df.columns:
        return pd.DataFrame(columns=["Action", "Description"])

    df = _filter_action_rows(df)

    ca_order_by_box: Dict[str, Dict[str, int]] = {}
    pair_orient: Dict[frozenset, Tuple[str, str]] = {}
    if enrich_from_json and json_source is not None:
        if isinstance(json_source, bytes):
            jbytes = json_source
        elif hasattr(json_source, "read"):
            json_source.seek(0)
            jbytes = json_source.read()
        else:
            with open(json_source, "rb") as fh:
                jbytes = fh.read()
        ca_order_by_box, pair_orient = _ca_order_and_pairs(jbytes)

    if ca_order_by_box or pair_orient:
        df = _reorder_pairs_in_df(df, ca_order_by_box, pair_orient)

    # Ensure consistent column order in output
    desired_cols = [c for c in ["Action", "Description", "SAP"] if c in df.columns]
    other_cols = [c for c in df.columns if c not in desired_cols]
    return df[desired_cols + other_cols].reset_index(drop=True)


# --------------------------- Excel export ---------------------------

def _autowidth_xlsxwriter(writer: pd.ExcelWriter, sheet_name: str, df: pd.DataFrame) -> None:
    ws = writer.sheets[sheet_name]
    # Compute max width per column (limit to avoid massive columns)
    for i, col in enumerate(df.columns):
        series = df[col].astype(str)
        width = min(max(series.map(len).max() + 2, len(str(col)) + 2), 80)
        ws.set_column(i, i, width)

def actions_to_workbook_bytes(df: pd.DataFrame, *, sheet_name: str = "Fibre Action") -> bytes:
    """Export the DataFrame to a single-sheet .xlsx with auto-width columns."""
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        df.to_excel(writer, index=False, sheet_name=sheet_name)
        _autowidth_xlsxwriter(writer, sheet_name, df)
    return output.getvalue()
