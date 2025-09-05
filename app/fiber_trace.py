
from __future__ import annotations
from pathlib import Path
import re, json
import pandas as pd
from typing import List, Tuple, Optional

# Exact schema to match reference
COLUMNS = [
    "Detail Item",
    "Report: Splice DetailsConnections",
    "# of Fbrs",
    "WO Action#",
    "Length",
    "~OTDR Length",
    "Meter Marks",
    "Eq Location",
    "EQ Type",
    "Activity",
    "Tray",
    "Slot",
    "Map It",
]

# Token decoding used in QSP-style exports
TOKENS = [
    ("<COMMA>", ", "),
    ("<COLON>", ":"),
    ("<OPEN>", "("),
    ("<CLOSE>", ")"),
    ("<AND>", "&"),
]

def decode_tokens(s: str) -> str:
    if not isinstance(s, str):
        return s
    out = s
    for src, dst in TOKENS:
        out = out.replace(src, dst)
    return re.sub(r"\s{2,}", " ", out).strip()

def split_sections(connections: str) -> list[str]:
    txt = connections.replace("\r\n", "\n").replace("\r", "\n")
    return [p.strip() for p in re.split(r"\n\.\n", txt) if p.strip()]

def extract_latlon(text: str):
    t = decode_tokens(text)
    m = re.search(r"([+-]?\d+\.\d+)\s*,\s*([+-]?\d+\.\d+)", t)
    return (float(m.group(1)), float(m.group(2))) if m else (None, None)

def guess_eq_location(desc: str) -> str:
    d = desc.lower()
    if "bfmh" in d or "beanfield manhole" in d: return "BFMH"
    if "hmh" in d or "hydro manhole" in d: return "HMH"
    return ""

def guess_eq_type(equip_line: str) -> str:
    e = equip_line.lower()
    if "osp splice box" in e or "generic osp splice box" in e: return "FOSC"
    return ""

def parse_existing_splice_fibers(line: str):
    line = decode_tokens(line)
    total, found = 0, False
    for a,b in re.findall(r"\[([0-9]+)\s*-\s*([0-9]+)\]", line):
        found = True
        a,b = int(a), int(b)
        if b >= a: total += (b - a + 1)
    return total if found else None

def parse_cableinfo_metrics(text: str):
    s = decode_tokens(text)
    # Length
    length = ""
    m_len = re.search(r"Cable Length\s*:\s*([0-9]+)", s, flags=re.I)
    if m_len:
        length = m_len.group(1)
    # ~OTDR Length
    otdr = ""
    m_otdr = re.search(r"OTDR\s*;\s*([0-9]+)", s, flags=re.I)
    if m_otdr:
        otdr = m_otdr.group(1)
    # Meter Marks
    mm = ""
    m_mm = re.search(r"OTDR\s*-\s*([0-9]+)", s, flags=re.I)
    if m_mm:
        mm = m_mm.group(1)
    return length, otdr, mm

def find_location_line(lines: list[str]) -> int:
    # Prefer a line with lat/lon; else anything with 'address'
    for i, ln in enumerate(lines):
        if re.search(r"[+-]?\d+\.\d+\s*,\s*[+-]?\d+\.\d+", decode_tokens(ln)):
            return i
    for i, ln in enumerate(lines):
        if "address" in ln.lower():
            return i
    return 0

def section_to_rows(section_text: str):
    rows: list[dict] = []
    lines = [ln for ln in section_text.splitlines() if ln.strip() and ln.strip() != "."]
    if not lines: 
        return rows, None

    # Location line + hyperlink URL
    loc_idx = find_location_line(lines)
    loc_line = decode_tokens(lines[loc_idx])
    lat, lon = extract_latlon(loc_line)
    map_url = f"https://maps.google.com/?q={lat},{lon}" if lat is not None and lon is not None else ""
    eq_loc  = guess_eq_location(loc_line)

    # 1) Equipment Location row (bold+red later; hyperlink only here)
    rows.append({
        "Detail Item": "Equipment Location",
        "Report: Splice DetailsConnections": loc_line,  # full address (no dots)
        "# of Fbrs": "",
        "WO Action#": "",
        "Length": "",
        "~OTDR Length": "",
        "Meter Marks": "",
        "Eq Location": eq_loc,
        "EQ Type": "",
        "Activity": "",
        "Tray": "",
        "Slot": "",
        "Map It": "Google Maps" if map_url else "",
    })

    # 2) Equipment row
    equip_line = ""
    if loc_idx + 1 < len(lines):
        equip_line = decode_tokens(lines[loc_idx + 1])
    if not equip_line or ("splice box" not in equip_line.lower() and not re.search(r"\\b\\d{3}D\\b", equip_line)):
        for ln in lines:
            if "splice box" in ln.lower() or re.search(r"\\b\\d{3}D\\b", ln):
                equip_line = decode_tokens(ln); break
    eq_type = guess_eq_type(equip_line)

    if equip_line:
        rows.append({
            "Detail Item": "Equipment",
            "Report: Splice DetailsConnections": equip_line,
            "# of Fbrs": "",
            "WO Action#": "",
            "Length": "",
            "~OTDR Length": "",
            "Meter Marks": "",
            "Eq Location": eq_loc,
            "EQ Type": eq_type,
            "Activity": "",
            "Tray": "",
            "Slot": "",
            "Map It": "",  # no link for non-location rows
        })

    # 3) Remaining lines
    for idx, ln in enumerate(lines):
        if idx in (loc_idx, loc_idx + 1): 
            continue
        text = decode_tokens(ln.strip())
        if not text or text == ".": 
            continue
        if "presented by qsp" in text.lower(): 
            continue

        if re.match(r"^(CA\\d+|CA[A-Za-z0-9\\-]+)\\s*[:;]", text) or text.startswith("CA"):
            d = {k: "" for k in COLUMNS}
            d.update({
                "Detail Item": "Cable[attach]",
                "Report: Splice DetailsConnections": text,
            })
            rows.append(d)
            continue

        if text.startswith("PMID:") and "Splice" in text:
            cnt = parse_existing_splice_fibers(text)
            d = {k: "" for k in COLUMNS}
            d.update({
                "Detail Item": "Existing Splice",
                "Report: Splice DetailsConnections": text,
                "# of Fbrs": str(cnt) if cnt is not None else "",
            })
            rows.append(d)
            continue

        if text.startswith("PMID:") and "Cable Length" in text:
            length, otdr, mm = parse_cableinfo_metrics(text)
            d = {k: "" for k in COLUMNS}
            d.update({
                "Detail Item": "CableInfo",
                "Report: Splice DetailsConnections": text,
                "Length": length,
                "~OTDR Length": otdr,
                "Meter Marks": mm,
            })
            rows.append(d)
            continue

        # Fallback informational line
        d = {k: "" for k in COLUMNS}
        d.update({
            "Detail Item": "Info",
            "Report: Splice DetailsConnections": text,
        })
        rows.append(d)

    return rows, (map_url or None)

def build_overview_df(json_path: Path):
    data = json.loads(json_path.read_text(encoding="utf-8"))
    connections = data["Report: Splice Details"][0][""][0]["Connections"]
    sections = split_sections(connections)

    all_rows: list[dict] = []
    urls: list[Optional[str]] = []
    for sec in sections:
        if len(sec.strip()) < 20: 
            continue
        rows, url = section_to_rows(sec)
        all_rows.extend(rows)
        for i, _ in enumerate(rows):
            urls.append(url if i == 0 else None)  # link only on first row of each cluster

    df = pd.DataFrame(all_rows, columns=COLUMNS).fillna("")
    return df, urls

def write_styled_excel(df: pd.DataFrame, urls, out_xlsx: Path, sheet="Sheet1"):
    with pd.ExcelWriter(out_xlsx, engine="xlsxwriter") as writer:
        df.to_excel(writer, index=False, sheet_name=sheet)
        wb  = writer.book
        ws  = writer.sheets[sheet]

        # Column widths to match reference
        widths = {"A":16,"B":80,"C":10,"D":12,"E":8,"F":14,"G":12,"H":12,"I":10,"J":10,"K":8,"L":8,"M":14}
        for col, w in widths.items():
            ws.set_column(f"{col}:{col}", w)

        # Freeze header
        ws.freeze_panes(1, 0)

        # Excel table styling
        nrows, ncols = df.shape
        headers = [{"header": h} for h in df.columns]
        ws.add_table(0, 0, nrows, ncols - 1, {
            "columns": headers,
            "style": "Table Style Medium 9",
            "banded_rows": True,
        })

        # Red+bold format for 'Equipment Location' rows
        fmt_loc = wb.add_format({"bold": True, "font_color": "white", "bg_color": "#D32F2F"})
        col_map = df.columns.get_loc("Map It")
        for r in range(nrows):
            is_loc = (df.iat[r, 0] == "Equipment Location")
            if is_loc:
                ws.set_row(r + 1, None, fmt_loc)      # +1 to skip header row
                if urls[r]:
                    ws.write_url(r + 1, col_map, urls[r], string="Google Maps")
            else:
                # ensure non-location rows have no link
                ws.write(r + 1, col_map, "")

def generate_xlsx(json_path: str | Path, out_xlsx: str | Path) -> pd.DataFrame:
    """Public entry: build dataframe and write only the Excel file (no CSV/KML)."""
    df, urls = build_overview_df(Path(json_path))
    write_styled_excel(df, urls, Path(out_xlsx), sheet="Sheet1")
    return df
