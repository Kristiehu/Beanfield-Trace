# fiber_action.py
from __future__ import annotations

import io
import json
import re
from typing import Dict, IO, List, Optional, Set, Tuple, Union

import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill, Border, Side
from openpyxl.utils import get_column_letter


# =========================
# Pair/Location parsing
# =========================

# Match a pair like:
#   "Splice 41121[73 - 84] - 41122[73 - 84]"
#   "Splice 41121 [73-84] 41122 [73-84]"
PAIR_RE = re.compile(
    r'(?P<verb>Splice|BREAK)\s+'
    r'(?P<A>[A-Za-z0-9#._/-]+)\s*\[\s*(?P<Ar>[^\]]+?)\s*\]\s*'
    r'(?:-|\s+)?\s*'
    r'(?P<B>[A-Za-z0-9#._/-]+)\s*\[\s*(?P<Br>[^\]]+?)\s*\]',
    re.IGNORECASE,
)

# Find all "... TOKEN - TOKEN ..." pairs near the tail; we will keep the LAST pair only.
TAIL_PAIR_RE = re.compile(r'([A-Za-z0-9#._/-]{2,})\s*-\s*([A-Za-z0-9#._/-]{2,})')

# JSON header for a splice box block, e.g. "M#248 : OSP Splice Box - 400D5 : M#248"
BOX_HEADER_RE = re.compile(r'^\s*([A-Za-z0-9#._/-]{2,})\s*:\s*OSP Splice Box\b.*$', re.IGNORECASE)
# CA lines like "CA2: PMID: 41121 …" (tolerate odd tokens between 'PMID' and the number)
CA_PMD_RE = re.compile(r'^\s*CA\d+.*?PMID[^\d]*([0-9]+[0-9A-Z]*)\b', re.IGNORECASE)
# Lines like "PMID: 41121 … -- Splice -- PMID: 41122 …"
PAIR_ORIENT_RE = re.compile(
    r'^\s*PMID[:]\s*([0-9A-Z]+)\b.*?--\s*Splice\s*--\s*PMID[:]\s*([0-9A-Z]+)\b',
    re.IGNORECASE,
)


# =========================
# CSV reading
# =========================
def _read_actions_csv(upload: Union[str, IO[bytes], bytes]) -> pd.DataFrame:
    """
    Read the WO CSV and return Action/Description/SAP table (starting at its header if present).
    """
    if isinstance(upload, bytes):
        text = upload.decode("utf-8", "ignore")
    elif hasattr(upload, "read"):
        upload.seek(0)
        text = upload.read().decode("utf-8", "ignore")
    else:
        text = open(upload, "r", encoding="utf-8").read()

    lines = text.splitlines()
    hdr_idx = None
    for i, line in enumerate(lines):
        if line.strip().lower().startswith("action,description"):
            hdr_idx = i
            break

# Only treat these as valid location tokens in tails
LOC_TOKEN_RE = re.compile(r'^(?:M#\d+[A-Z]?|D#\d+[A-Z]?|PA\d+|BFMA\d+)$', re.IGNORECASE)

def _pick_location_tail(s: str) -> tuple[str, str]:
    """Return the last 'LocA - LocB' pair where both sides look like real locations."""
    pairs = TAIL_PAIR_RE.findall(s)
    for left, right in reversed(pairs):
        if LOC_TOKEN_RE.match(left) and LOC_TOKEN_RE.match(right):
            return left, right
    return "", ""


    if hdr_idx is not None:
        df = pd.read_csv(io.StringIO("\n".join(lines[hdr_idx:])),
                         dtype=str, keep_default_na=False)
    else:
        df = pd.read_csv(io.StringIO(text), dtype=str, keep_default_na=False)

    # normalize columns
    cols = [c.strip() for c in df.columns]
    rename = {}
    for c in df.columns:
        lc = c.strip().lower()
        if lc == "action": rename[c] = "Action"
        if lc == "description": rename[c] = "Description"
        if lc == "sap": rename[c] = "SAP"
    df = df.rename(columns=rename)
    for col in ("Action", "Description"):
        if col not in df.columns:
            df[col] = ""
    if "SAP" not in df.columns:
        df["SAP"] = ""
    return df[["Action", "Description", "SAP"]].copy()


# =========================
# JSON → CA order maps
# =========================
def _extract_connection_blobs(json_source: Optional[Union[str, IO[bytes], bytes]]) -> List[str]:
    if json_source is None:
        return []
    if isinstance(json_source, bytes):
        raw = json_source.decode("utf-8", "ignore")
    elif hasattr(json_source, "read"):
        json_source.seek(0)
        raw = json_source.read().decode("utf-8", "ignore")
    else:
        raw = open(json_source, "r", encoding="utf-8").read()
    try:
        data = json.loads(raw)
    except Exception:
        return []
    blobs: List[str] = []
    try:
        for blk in data.get("Report: Splice Details", [])[0].get("", []):
            s = blk.get("Connections")
            if isinstance(s, str) and s.strip():
                blobs.append(s)
    except Exception:
        pass
    return blobs


def _parse_json_for_order(json_source: Optional[Union[str, IO[bytes], bytes]]
                          ) -> Tuple[Dict[str, Dict[str, int]], Dict[str, Set[str]], Dict[frozenset, Tuple[str, str]]]:
    """
    Returns:
      ca_order_by_box: { box_name -> { pmid -> rank } }  # from CA1, CA2, ...
      boxes_by_pmid:   { pmid -> set(box_name) }
      pair_orient:     { frozenset({A,B}) -> (A,B) }     # explicit JSON "PMID A -- Splice -- PMID B"
    """
    ca_order_by_box: Dict[str, Dict[str, int]] = {}
    boxes_by_pmid: Dict[str, Set[str]] = {}
    pair_orient: Dict[frozenset, Tuple[str, str]] = {}

    for blob in _extract_connection_blobs(json_source):
        current_box: Optional[str] = None
        order_map: Dict[str, int] = {}
        for line in blob.splitlines():
            m = BOX_HEADER_RE.match(line)
            if m:
                if current_box and order_map and current_box not in ca_order_by_box:
                    ca_order_by_box[current_box] = dict(order_map)
                current_box = m.group(1).strip()
                order_map = {}
                continue
            m = CA_PMD_RE.match(line)
            if m and current_box:
                pmid = m.group(1).strip()
                if pmid not in order_map:
                    order_map[pmid] = len(order_map)
                    boxes_by_pmid.setdefault(pmid, set()).add(current_box)
                continue
            m = PAIR_ORIENT_RE.match(line)
            if m:
                a, b = m.group(1), m.group(2)
                pair_orient.setdefault(frozenset((a, b))), (a, b)
                pair_orient[frozenset((a, b))] = (a, b)
        if current_box and order_map and current_box not in ca_order_by_box:
            ca_order_by_box[current_box] = dict(order_map)

    return ca_order_by_box, boxes_by_pmid, pair_orient


def _find_common_box(
    A: str,
    B: str,
    ca_order_by_box: Dict[str, Dict[str, int]],
    boxes_by_pmid: Dict[str, Set[str]],
    loc_hint: Optional[str],
) -> Optional[str]:
    """
    Prefer a box that contains BOTH PMIDs.
    If loc_hint is provided and valid and contains both, use it.
    If multiple common boxes, prefer one where A<=B by CA order.
    """
    if loc_hint and loc_hint in ca_order_by_box:
        om = ca_order_by_box[loc_hint]
        if A in om and B in om:
            return loc_hint

    common = boxes_by_pmid.get(A, set()) & boxes_by_pmid.get(B, set())
    if not common:
        return None
    if len(common) == 1:
        return next(iter(common))
    # choose the first where A comes before (or equal) B
    for bx in common:
        om = ca_order_by_box.get(bx, {})
        if A in om and B in om and om[A] <= om[B]:
            return bx
    # otherwise take a stable one
    return sorted(common)[0]


# =========================
# Normalization / Color
# =========================
def _normalize_desc(action: str, desc: str,
                    ca_order_by_box: Dict[str, Dict[str, int]],
                    boxes_by_pmid: Dict[str, Set[str]],
                    pair_orient: Dict[frozenset, Tuple[str, str]]) -> str:
    """
    Parse any messy description, enforce pair order by CA list, and emit:
      "Splice 41121 [73-84] 41122 [73-84] M#248 - M#248"
      "BREAK  35289 [3-4]  35116 [3-4]  M#1189A - PA27014"
    """
    s = str(desc or "")
    m = PAIR_RE.search(s)
    if not m:
        return s  # leave non-pair lines untouched

    verb = m.group("verb")
    A, Ar, B, Br = m.group("A"), m.group("Ar"), m.group("B"), m.group("Br")
    # Decide verb from action too (e.g., "Remove (E)" → BREAK)
    if "remove" in (action or "").lower() or "break" in (action or "").lower():
        verb = "BREAK"
    else:
        verb = "Splice"

    # find last "TOKEN - TOKEN" tail pair → location A/B
    locA = locB = ""
    tail_pairs = TAIL_PAIR_RE.findall(s)
    if tail_pairs:


        # find real location tail (ignore 'Box - 400D5' etc.)
        locA, locB = _pick_location_tail(s)
    loc_hint = locA if (locA == locB and locA) else None

    # Order by CA list within chosen box
    box = _find_common_box(A, B, ca_order_by_box, boxes_by_pmid, loc_hint)
    if box:
        om = ca_order_by_box.get(box, {})
        if A in om and B in om and om[A] > om[B]:
            A, B, Ar, Br = B, A, Br, Ar
    else:
        # fallback 1: explicit JSON orientation, if seen
        key = frozenset((A, B))
        if key in pair_orient:
            a, b = pair_orient[key]
            if {A, B} == {a, b} and A == b:
                A, B, Ar, Br = B, A, Br, Ar
        else:
            # fallback 2: if only one PMID is known anywhere, keep the known one on the left
            a_boxes = boxes_by_pmid.get(A, set())
            b_boxes = boxes_by_pmid.get(B, set())
            if a_boxes and not b_boxes:
                pass  # A already on the left
            elif b_boxes and not a_boxes:
                A, B, Ar, Br = B, A, Br, Ar

    # Build normalized output; keep location pair if available
    core = f"{verb} {A} [{Ar}] {B} [{Br}]"
    if locA and locB:
        core += f" {locA} - {locB}"
    return core



def _assign_color(action: str, description: str) -> str:
    """
    RGB hex without '#':
      BREAK -> red,  Splice -> orange.
      (You can expand rules later for equipment/cableinfo rows if those are included.)
    """
    a = (action or "").lower()
    d = (description or "").lower()
    if "break" in d or "remove" in a:
        return "FF6666"  # red
    return "FFA64D"      # orange (Splice default)


# =========================
# Public API
# =========================
def compute_fibre_actions(
    csv_source: Union[str, IO[bytes], bytes],
    json_source: Optional[Union[str, IO[bytes], bytes]] = None,
    **_compat_kwargs,  # absorb legacy kwargs like enrich_from_json
) -> pd.DataFrame:
    """
    Build Fibre Action table:
      - Parse CSV Action/Description/SAP.
      - Reorder splice/break pairs using JSON CA order.
      - Normalize Description to: "<Splice|BREAK> A [ra] B [rb] LeftLoc - RightLoc".
      - Add Color column for Excel styling.
    """
    df = _read_actions_csv(csv_source)
    if df.empty:
        return df

    ca_order_by_box, boxes_by_pmid, pair_orient = _parse_json_for_order(json_source)

    # Only rows that contain "Splice" or "BREAK" are reformatted; others pass through
    mask = df["Description"].str.contains(r"\b(Splice|BREAK|Remove)\b", case=False, regex=True)
    df = df.loc[mask].reset_index(drop=True)

    df["Description"] = [
        _normalize_desc(a, d, ca_order_by_box, boxes_by_pmid, pair_orient)
        for a, d in zip(df["Action"], df["Description"])
    ]
    df["Color"] = [_assign_color(a, d) for a, d in zip(df["Action"], df["Description"])]

    return df[["Action", "Description", "SAP", "Color"]]


# =========================
# Excel export (color rows)
# =========================
def actions_to_workbook_bytes(df: pd.DataFrame, *, sheet_name: str = "Fibre Action") -> bytes:
    wb = Workbook()
    ws = wb.active
    ws.title = sheet_name

    headers = ["Action", "Description", "SAP"]
    ws.append(headers)

    # Styles
    bold = Font(bold=True)
    center = Alignment(horizontal="center", vertical="center", wrap_text=True)
    top_left = Alignment(vertical="top", wrap_text=True)
    box = Border(left=Side(style="thin"), right=Side(style="thin"),
                 top=Side(style="thin"), bottom=Side(style="thin"))
    header_fill = PatternFill("solid", fgColor="DDDDDD")

    # Header cells
    for cell in ws[1]:
        cell.font = bold
        cell.alignment = center
        cell.border = box
        cell.fill = header_fill

    # Data rows with color
    for _, r in df.iterrows():
        ws.append([r.get("Action", ""), r.get("Description", ""), r.get("SAP", "")])
        row_idx = ws.max_row
        fill_clr = (r.get("Color") or "FFFFFF").upper().replace("#", "")
        row_fill = PatternFill("solid", fgColor=fill_clr)
        for cell in ws[row_idx]:
            cell.border = box
            cell.alignment = top_left
            cell.fill = row_fill

    # Auto-width
    for i, col in enumerate(headers, start=1):
        series = [str(col)] + [str(v) for v in df[col]]
        width = min(max(len(x) for x in series) + 2, 100)
        ws.column_dimensions[get_column_letter(i)].width = width

    out = io.BytesIO()
    wb.save(out)
    return out.getvalue()
