# filename: fiber_trace.py
from __future__ import annotations
from pathlib import Path
from typing import Dict, Iterable, List, Optional, Tuple
import json, re
import pandas as pd

# =========================
# Schema
# =========================
COLUMNS = [
    "Detail Item",
    "Report: Splice DetailsConnections",
    "# of Fbrs",
    "WO Action#",
    "Length",
    "OTDR Length",      
    "Meter Marks",
    "Eq Location",
    "EQ Type",
    "Activity",
    "Tray",
    "Slot",
    "Map It",
]

# =========================
# Utilities
# =========================
_RX_PMID         = re.compile(r"\bPMID\s*:\s*(\d+)", re.I)
_RX_APTUM_F      = re.compile(r"\bAptum\s*ID\s*:\s*\.?F(\d+)", re.I)
_RX_LEN_CALC     = re.compile(r"(?:Calculated|Cable Length Calculated)\s*:\s*([0-9]+)", re.I)
_RX_LEN_OTDR     = re.compile(r"(?:OTDR|Cable Length\s*:?OTDR)\s*[: ]\s*([0-9]+)", re.I)
_RX_RANGE        = re.compile(r"\[([0-9]+)\s*-\s*([0-9]+)\]")
_RX_FX           = re.compile(r"\.F(\d+)", re.I)

def _compact_numbers(nums: Iterable[int]) -> str:
    """Return '3,5-8,11' from [3,5,6,7,8,11]."""
    s = sorted(set(nums))
    if not s:
        return ""
    out: List[str] = []
    start = prev = s[0]
    for n in s[1:]:
        if n == prev + 1:
            prev = n
            continue
        # flush
        if start == prev:
            out.append(f"{start}")
        else:
            out.append(f"{start}-{prev}")
        start = prev = n
    # flush last
    if start == prev:
        out.append(f"{start}")
    else:
        out.append(f"{start}-{prev}")
    return ",".join(out)

def _extract_pmids(line: str) -> List[int]:
    return [int(m.group(1)) for m in _RX_PMID.finditer(line or "")]

def _extract_aptum_f(line: str) -> List[int]:
    return [int(m.group(1)) for m in _RX_APTUM_F.finditer(line or "")]

def _extract_lengths_blob(cable_lengths: List[dict]) -> Dict[int, Tuple[Optional[int], Optional[int]]]:
    """
    Turn:
      [{"PMID: 34377": "Calculated: 265m"}, {"PMID: 35844": "OTDR: 125m, Calculated: 106m"}]
    into:
      {34377: (265, None), 35844: (106, 125)}
    """
    out: Dict[int, Tuple[Optional[int], Optional[int]]] = {}
    for item in cable_lengths or []:
        for k, v in item.items():
            pmid_match = _RX_PMID.search(k)
            if not pmid_match:
                continue
            pmid = int(pmid_match.group(1))
            calc = None
            otdr = None
            # normalize value to plain "words" so "m" suffix doesn't matter
            text = str(v)
            m_calc = _RX_LEN_CALC.search(text)
            if m_calc:
                calc = int(m_calc.group(1))
            m_otdr = _RX_LEN_OTDR.search(text)
            if m_otdr:
                otdr = int(m_otdr.group(1))
            out[pmid] = (calc, otdr)
    return out

def _build_equipment_location_line(location: str, typ: str) -> Tuple[List[str], Optional[str]]:
    value = f"{location}, {typ}".strip().strip(",")
    row = [
        "Equipment Location",
        value,
        "", "", "", "", "", ""  # (# of Fbrs, WO Action#, Length, ~OTDR Length, Meter Marks, Google Maps)
    ]
    # Google Maps URL can be attached by your upstream (if you have lat/lon),
    # here we keep it blank and let the caller fill if available.
    return row, None

def _build_equipment_line(equipment: str) -> List[str]:
    return ["Equipment", equipment, "", "", "", "", "", ""]

def _build_cable_attach_lines(cables: List[dict]) -> List[List[str]]:
    rows: List[List[str]] = []
    for idx, item in enumerate(cables or [], start=1):
        # allow {"CA1": "..."} style
        if isinstance(item, dict) and item:
            label, text = list(item.items())[0]
            value = f"{label}: {text}"
        else:
            value = str(item)
        rows.append(["Cable[attach]", value, "", "", "", "", "", ""])
    return rows

def _group_existing_splices(splice_lines: List[str]) -> Optional[Tuple[str, int]]:
    """
    Given many line items like:
      "PMID: 34377, Aptum ID: .F243 -- Splice -- PMID: 35844, Aptum ID: .F387"
    return:
      ("PMID: 34377, Aptum ID: .F[243-246] -- Existing Splice(s) -- PMID: 35844, Aptum ID: .F[387-390]", 4)
    Assumptions:
      - All these lines belong to the current equipment block.
      - PMIDs on both sides are consistent (if multiple distinct pairs exist, we keep the first consistent pair).
    """
    if not splice_lines:
        return None
    # find the dominant pair (left PMID, right PMID)
    pairs: Dict[Tuple[int, int], List[Tuple[int, int]]] = {}
    for line in splice_lines:
        pmids = _extract_pmids(line)
        if len(pmids) < 2:
            # we require Left/Right PMIDs
            continue
        left_pmid, right_pmid = pmids[0], pmids[-1]
        left_fs  = _extract_aptum_f(line)
        # heuristic: first half of Fs belong to left; second half belong to right
        # but many lines format as exactly two .F numbers – one per side – which is ideal
        if len(left_fs) >= 2:
            mid = len(left_fs) // 2
            lfs = [left_fs[0]]
            rfs = [left_fs[-1]]
        elif len(left_fs) == 1:
            lfs, rfs = [left_fs[0]], []
        else:
            lfs, rfs = [], []

        # Alternatively, a more robust parse for exactly two .F values:
        both = list(_RX_FX.finditer(line))
        if len(both) >= 2:
            lfs = [int(both[0].group(1))]
            rfs = [int(both[-1].group(1))]

        pairs.setdefault((left_pmid, right_pmid), [])
        for L in lfs or [None]:
            for R in rfs or [None]:
                if L is not None and R is not None:
                    pairs[(left_pmid, right_pmid)].append((L, R))

    if not pairs:
        return None

    # pick the pair with most observations
    (lp, rp), fibers = max(pairs.items(), key=lambda kv: len(kv[1]))
    left_numbers  = [a for a, _ in fibers]
    right_numbers = [b for _, b in fibers]

    left_comp  = _compact_numbers(left_numbers)
    right_comp = _compact_numbers(right_numbers)
    fbrs_count = len(set(left_numbers))  # # of fibers from left side (should match right if 1-to-1)

    left_part  = f"PMID: {lp}, Aptum ID: .F[{left_comp}]"  if left_comp  else f"PMID: {lp}"
    right_part = f"PMID: {rp}, Aptum ID: .F[{right_comp}]" if right_comp else f"PMID: {rp}"
    text = f"{left_part} -- Existing Splice(s) -- {right_part}"
    return text, fbrs_count

def _build_cableinfo_lines(cable_lengths_map: Dict[int, Tuple[Optional[int], Optional[int]]]) -> List[List[str]]:
    rows: List[List[str]] = []
    # deterministic order by PMID
    for pmid in sorted(cable_lengths_map.keys()):
        calc, otdr = cable_lengths_map[pmid]
        line = f"PMID: {pmid}: " + \
               (f"OTDR: {otdr}m, " if otdr is not None else "") + \
               (f"Calculated: {calc}m" if calc is not None else "")
        rows.append(["CableInfo", line, "", "", 
                     (calc if calc is not None else ""), 
                     (otdr if otdr is not None else ""),
                     "", ""])
    return rows

# =========================
# Transformer (one location)
# =========================
def build_location_block(location_payload: dict) -> Tuple[List[List[str]], List[Optional[str]]]:
    """
    Input example:
    {
      "Location": "Montreal, PA21483",
      "Type": "Beanfield Manhole/Handwell with Splice Box",
      "Equipment": "21483-450D-1 : OSP Splice Box - 600D : PA21483, 450-1 FOSC",
      "Cables": [{"CA1": "PMID: 33557, 288F, 1250 Rene Levesque, 450-2"}, {"CA2": "PMID: 33805, 288F, PA15393"}],
      "Splice": [
          "PMID: 34377, Aptum ID: .F243 -- Splice -- PMID: 35844, Aptum ID: .F387",
          "PMID: 34377, Aptum ID: .F244 -- Splice -- PMID: 35844, Aptum ID: .F388",
          "PMID: 34377, Aptum ID: .F245 -- Splice -- PMID: 35844, Aptum ID: .F389",
          "PMID: 34377, Aptum ID: .F246 -- Splice -- PMID: 35844, Aptum ID: .F390"
      ],
      "Cable Lengths": [
         {"PMID: 34377": "Calculated: 265m"},
         {"PMID: 35844": "OTDR: 125m, Calculated: 106m"}
      ]
    }
    """
    rows: List[List[str]] = []
    urls: List[Optional[str]] = []   # parallel list for Google Maps url (set only on location row)

    # 1) Equipment Location
    loc = str(location_payload.get("Location", "")).strip()
    typ = str(location_payload.get("Type", "")).strip()
    row_loc, link = _build_equipment_location_line(loc, typ)
    rows.append(row_loc)
    urls.append(link)

    # 2) Equipment
    eq = str(location_payload.get("Equipment", "")).strip()
    if eq:
        rows.append(_build_equipment_line(eq)); urls.append(None)

    # 3) Cable[attach] list
    cables = location_payload.get("Cables") or []
    for r in _build_cable_attach_lines(cables):
        rows.append(r); urls.append(None)

    # 4) Existing Splice grouping (within this equipment only)
    splice_val = location_payload.get("Splice")
    splice_lines: List[str] = []
    if isinstance(splice_val, list):
        splice_lines = [str(x) for x in splice_val]
    elif isinstance(splice_val, str) and splice_val.strip():
        # If you still sometimes store a single delimited string
        splice_lines = [s for s in re.split(r"\n|\r\n", splice_val) if s.strip()]

    grouped = _group_existing_splices(splice_lines)
    if grouped:
        text, cnt = grouped
        rows.append(["Existing Splice", text, int(cnt), "", "", "", "", ""])
        urls.append(None)

    # 5) CableInfo (with numeric Length / ~OTDR Length)
    cable_lengths_map = _extract_lengths_blob(location_payload.get("Cable Lengths") or [])
    for r in _build_cableinfo_lines(cable_lengths_map):
        rows.append(r); urls.append(None)

    return rows, urls

# =========================
# High-level builders
# =========================
def build_overview_df(json_path: Path) -> Tuple[pd.DataFrame, List[Optional[str]]]:
    """
    Read your work-order JSON (already pre-cleaned), iterate locations, emit a single DF.
    Expected top-level schema: { "locations": [ <location_payload>, ... ] }
    If your file is raw, adapt here to reach that shape.
    """
    data = json.loads(Path(json_path).read_text(encoding="utf-8"))

    # Be flexible: accept either {"locations": [...]} or a flat single location object
    locs = data.get("locations")
    if isinstance(locs, list):
        location_list = locs
    else:
        # fallback: if the file *is already* a single location payload
        location_list = [data]

    all_rows: List[List[str]] = []
    all_urls: List[Optional[str]] = []
    for block in location_list:
        rows, urls = build_location_block(block)
        all_rows.extend(rows)
        all_urls.extend(urls)

    df = pd.DataFrame(all_rows, columns=COLUMNS)
    return df, all_urls

# =========================
# Excel export with first-column coloring
# =========================
def write_styled_excel(df: pd.DataFrame, urls: List[Optional[str]], out_path: Path, sheet: str = "Sheet1") -> None:
    """
    Writes the dataframe to XLSX using xlsxwriter and applies color to the *first column*
    based on the “Detail Item” value.
      - Equipment / Equipment Location: dark green
      - Cable[attach]: light green
      - Break Splice: red
      - Splice required: orange
      - CableInfo: blue
      - Existing Splice: (treat as 'Equipment' tone? Spec didn’t mandate; we leave default (black) or choose dark green)
    """
    out_path = Path(out_path)
    with pd.ExcelWriter(out_path, engine="xlsxwriter") as xw:
        df.to_excel(xw, index=False, sheet_name=sheet)
        wb  = xw.book
        ws  = xw.sheets[sheet]

        # Formats
        fmt_loc   = wb.add_format({"font_color": "#0B6623", "bold": True})  # dark green
        fmt_equip = wb.add_format({"font_color": "#0B6623", "bold": False})
        fmt_cab   = wb.add_format({"font_color": "#32CD32"})                # light/grass green
        fmt_break = wb.add_format({"font_color": "#FF0000"})                 # red
        fmt_req   = wb.add_format({"font_color": "#FF8C00"})                 # orange
        fmt_info  = wb.add_format({"font_color": "#0000FF"})                 # blue

        # Column widths (auto-ish)
        for ci, col in enumerate(df.columns):
            width = max(10, min(80, int(df[col].astype(str).map(len).max() + 2)))
            ws.set_column(ci, ci, width)

        # Apply first-column styling and optional link in "Google Maps"
        detail_col = 0
        link_col   = df.columns.get_loc("Google Maps") if "Google Maps" in df.columns else None

        for r in range(len(df)):
            label = str(df.iat[r, detail_col])
            if   label == "Equipment Location":
                ws.write(r + 1, detail_col, label, fmt_loc)
            elif label == "Equipment":
                ws.write(r + 1, detail_col, label, fmt_equip)
            elif label == "Cable[attach]":
                ws.write(r + 1, detail_col, label, fmt_cab)
            elif label == "Break Splice":
                ws.write(r + 1, detail_col, label, fmt_break)
            elif label == "Splice required":
                ws.write(r + 1, detail_col, label, fmt_req)
            elif label == "CableInfo":
                ws.write(r + 1, detail_col, label, fmt_info)
            else:
                # Existing Splice or any other; leave default
                pass

            # Optional Google Maps link on the same row
            if link_col is not None:
                url = urls[r] if r < len(urls) else None
                if label == "Equipment Location" and url:
                    ws.write_url(r + 1, link_col, url, string="Google Maps")
                else:
                    ws.write(r + 1, link_col, "")

def generate_xlsx(json_path: str | Path, out_xlsx: str | Path) -> pd.DataFrame:
    """Public entrypoint for the Streamlit button: build the DF and write the xlsx."""
    df, urls = build_overview_df(Path(json_path))
    write_styled_excel(df, urls, Path(out_xlsx), sheet="Fibre Trace")
    return df
