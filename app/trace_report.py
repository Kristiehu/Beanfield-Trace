#!/usr/bin/env python3
"""
Generate a 2‑sheet Excel report from:
  1) Work Order CSV (like WO24218.csv)
  2) Circuit JSON (like WO24218.json)
and a few manual “metadata” fields (same as your app UI).

Output:
  - Sheet "Summary": header grid matching your screenshot + Action/Description/SAP
  - Sheet "Details": rows grouped like “Equipment Location / Equipment / Cable[attach] / Existing Splice / CableInfo … Google Maps”
    with a Google Maps hyperlink when (lat, lon) are found.

Usage (CLI example):
  python generate_trace_report.py \
      --wo_csv WO24218.csv \
      --circuit_json WO24218.json \
      --out_xlsx OUT.xlsx \
      --designer_name "abcd" \
      --designer_email "abcd@beanfield.com" \
      --designer_phone "4164164164" \
      --fibers 2 \
      --order_id "ORDER-267175" \
      --client_name "ABC In" \
      --build_type "Order" \
      --a_end "BFMH-0021" \
      --z_end "SITE5761" \
      --route_type "Primary" \
      --circuit_version "1" \
      --circuit_id "CK37105" \
      --jira "ORDER-267175" \
      --service_type "IRU Fbr" \
      --device_type "N/A" \
      --circuit_type "NEW"
"""

import argparse
import json
import re
from pathlib import Path

import pandas as pd
import io, os, json


# ---------- Helpers

def _safe_read_csv(path: Path) -> pd.DataFrame:
    return pd.read_csv(path, on_bad_lines="skip", encoding_errors="ignore")

def _kv_from_wo(df: pd.DataFrame) -> dict:
    """
    Your WO CSV looks like:
        Work Order:,WO24218,
        Phase:,New,
        ...
        ORDER Number:,ORDER-267175,
        ,,
        Action,Description,SAP
        1: Remove (E),...
    Pandas will have the first line as header; we normalize.
    """
    # Normalize columns
    df = df.copy()
    df.columns = [str(c).strip() for c in df.columns]
    # All key:value rows live in the first two columns
    c0, c1 = df.columns[:2]
    # stop before the Action header
    up_to = df.index[df[c0].astype(str).str.strip().str.lower() == "action"]
    end = int(up_to[0]) if len(up_to) else len(df)

    kv = {}
    for _, row in df.iloc[:end].iterrows():
        k = str(row[c0]).strip()
        v = str(row[c1]).strip()
        if k and k.lower() != "nan":
            if v.lower() in ("nan", ""):
                continue
            kv[k.rstrip(":")] = v
    return kv

def _actions_from_wo(df: pd.DataFrame) -> pd.DataFrame:
    # Find the "Action" header row
    df = df.copy()
    df.columns = [str(c).strip() for c in df.columns]
    c0 = df.columns[0]
    start_idx = df.index[df[c0].astype(str).str.strip().str.lower() == "action"]
    if not len(start_idx):
        return pd.DataFrame(columns=["Action", "Description", "SAP"])
    start = int(start_idx[0]) + 1

    # Slice & standardize columns
    sub = df.iloc[start:].reset_index(drop=True)
    cols = list(sub.columns) + ["", ""]
    # Map first 3 columns to Action/Description/SAP
    out = sub.iloc[:, :3].copy()
    out.columns = ["Action", "Description", "SAP"]
    # Drop completely empty rows
    out = out.dropna(how="all")
    return out

def _count_breaks(actions_df: pd.DataFrame) -> int:
    if actions_df.empty:
        return 0
    txt = actions_df["Description"].astype(str).str.upper()
    # robust: count rows that *contain* 'BREAK'
    return int((txt.str.contains(r"\bBREAK\b", regex=True)).sum())

def _count_splices(actions_df: pd.DataFrame) -> int:
    if actions_df.empty:
        return 0
    txt = actions_df["Description"].astype(str).str.upper()
    return int((txt.str.contains(r"\bSPLICE\b", regex=True)).sum())

def _normalize_json_block_text(block: str) -> str:
    """
    Your JSON stores long text with tokens like <COLON>, <COMMA>.
    Normalize them to real punctuation and unify newlines.
    """
    s = block
    s = s.replace("<COLON>", ":").replace("<COMMA>", ",")
    # Some dumps use '\n.' lines; keep real newlines
    s = s.replace("\\n", "\n")
    return s

def _extract_latlon(s: str):
    """
    Try a few patterns to find latitude/longitude in a line.
    Returns (lat, lon) as strings or (None, None).
    """
    # e.g., "Lat: 43.64, Lon: -79.38"
    m = re.search(r"lat[:=]\s*([\-+]?\d+(?:\.\d+)?)\D+lon[:=]\s*([\-+]?\d+(?:\.\d+)?)", s, re.I)
    if m:
        return m.group(1), m.group(2)
    # e.g., "... 43.64, -79.38 ..."
    m = re.search(r"([\-+]?\d+(?:\.\d+)?)\s*,\s*([\-+]?\d+(?:\.\d+)?)", s)
    if m:
        # crude sanity: lat in [-90,90], lon in [-180,180]
        lat, lon = float(m.group(1)), float(m.group(2))
        if -90 <= lat <= 90 and -180 <= lon <= 180:
            return str(lat), str(lon)
    return None, None
    

def _details_from_json(json_src) -> pd.DataFrame:
    """
    Accepts a filesystem path, a Streamlit UploadedFile, bytes, or any file-like object.
    """
    # Path-like
    if isinstance(json_src, (str, os.PathLike, Path)):
        with open(json_src, "r", encoding="utf-8", errors="ignore") as f:
            j = json.load(f)

    # Streamlit UploadedFile (preferred)
    elif hasattr(json_src, "getvalue"):
        raw = json_src.getvalue()                       # bytes
        j = json.loads(raw.decode("utf-8", errors="ignore"))

    # Generic file-like object
    elif hasattr(json_src, "read"):
        raw = json_src.read()
        if isinstance(raw, (bytes, bytearray)):
            raw = raw.decode("utf-8", errors="ignore")
        j = json.loads(raw)

    # Raw bytes
    elif isinstance(json_src, (bytes, bytearray)):
        j = json.loads(bytes(json_src).decode("utf-8", errors="ignore"))

    else:
        raise TypeError(f"Unsupported json_src type: {type(json_src)}")

#  Parse JSON into a details table
def build_details_df_from_payload(payload: dict) -> pd.DataFrame:
    """
    Parses payload["Report: Splice Details"] into a simple dataframe.
    Returns columns that are safe for .copy() and downstream use.
    """
    text_blob = None
    try:
        text_blob = payload["Report: Splice Details"][0][""][0]
    except Exception:
        # fallback: find the first very long string in the payload
        def find_first_long_string(obj):
            if isinstance(obj, str) and len(obj) > 1000: return obj
            if isinstance(obj, dict):
                for v in obj.values():
                    r = find_first_long_string(v)
                    if r: return r
            if isinstance(obj, list):
                for v in obj:
                    r = find_first_long_string(v)
                    if r: return r
            return None
        text_blob = find_first_long_string(payload)

    if not text_blob:
        return pd.DataFrame(columns=[
            "segment_index","city","site","address","latitude","longitude","box_descriptor","raw"
        ])

    segments = re.split(r"\n\.\n\.\n", text_blob.strip())
    header_re = re.compile(
        r"""^\s*
        (?P<city>[^,]+)\s*,\s*
        (?P<site>[^,]+)\s*,\s*
        Address:(?P<address>[^,]*)
        \(\)\s*:\s*
        (?P<lat>-?\d+(?:\.\d+)?)\s*,\s*(?P<lon>-?\d+(?:\.\d+)?)
        \s*:\s*:\s*
        (?P<box>.+?)\s*$""",
        re.VERBOSE
    )

    rows = []
    for idx, seg in enumerate(segments):
        seg = seg.strip()
        lines = [l for l in seg.splitlines() if l.strip()]
        city=site=address=lat=lon=box=None
        for line in lines:
            m = header_re.match(line)
            if m:
                city = m.group("city")
                site = m.group("site")
                address = m.group("address")
                lat = float(m.group("lat"))
                lon = float(m.group("lon"))
                box = m.group("box")
                break
        rows.append(dict(
            segment_index=idx,
            city=city, site=site, address=address,
            latitude=lat, longitude=lon,
            box_descriptor=box, raw=seg
        ))
    return pd.DataFrame(rows)

def build_workbook(
    out_xlsx: Path,
    meta: dict,
    wo_kv: dict,
    actions_df: pd.DataFrame,
    details_df: pd.DataFrame,
    end_to_end_m: str = "",
    end_to_end_otdr_m: str = "",
):
    """
    Write Excel with two sheets using xlsxwriter formatting and a Google Maps hyperlink column.
    """
    with pd.ExcelWriter(out_xlsx, engine="xlsxwriter") as xl:
        # -------- Sheet 1: Summary
        ws1_name = "Summary"
        # Header grid matching your screenshot
        # Left side (labels / values), Right side (counts & A/Z ends)
        summary_rows = [
            ("Order Number:", meta["order_id"], "Number of Fibre Breaks:", str(_count_breaks(actions_df))),
            ("Work Order Number:", wo_kv.get("Work Order", ""), "Number of Fibre Splices", str(_count_splices(actions_df))),
            ("Order A to Z:", f"{meta['a_end']}_{meta['z_end']}", "End to End Length(m)", end_to_end_m),
            ("Designer:", meta["designer_name"], "End to End ~ OTDR(m)", end_to_end_otdr_m),
            ("Contact Number:", meta["designer_phone"], "A END:", meta["a_end"]),
            ("Date (dd/mm/yyyy):", wo_kv.get("Created On", ""), "Z END:", meta["z_end"]),
            ("Details:", wo_kv.get("Details", ""), "", ""),
            ("ORDER Number:", meta["order_id"], "", ""),
        ]

        # Create a sheet & write the grid
        book = xl.book
        ws1 = xl.book.add_worksheet(ws1_name)
        header_fmt = book.add_format({"bold": True, "align": "left"})
        border_fmt = book.add_format({"border": 1})
        ws1.set_column("A:A", 22)
        ws1.set_column("B:B", 40)
        ws1.set_column("C:C", 28)
        ws1.set_column("D:D", 22)

        r = 0
        for (l1, v1, l2, v2) in summary_rows:
            ws1.write(r, 0, l1, header_fmt)
            ws1.write(r, 1, v1)
            if l2:
                ws1.write(r, 2, l2, header_fmt)
                ws1.write(r, 3, v2)
            r += 1

        # Blank row
        r += 1

        # Action table header
        ws1.write(r, 0, "Action", header_fmt)
        ws1.write(r, 1, "Description", header_fmt)
        ws1.write(r, 2, "SAP", header_fmt)
        r += 1

        # Action rows
        if not actions_df.empty:
            for _, row in actions_df.iterrows():
                ws1.write(r, 0, str(row.get("Action", "")))
                ws1.write(r, 1, str(row.get("Description", "")))
                ws1.write(r, 2, str(row.get("SAP", "")))
                r += 1

        # Add a light border around the action table
        ws1.conditional_format( (len(summary_rows)+2, 0, r-1, 2),
                                {"type": "no_blanks", "format": border_fmt})

        # -------- Sheet 2: Details
        ws2_name = "Details"
        details_cols = [
            "Detail Item", "Report: Splice DetailsConnections", "# of Fbrs", "WO Action#",
            "Length", "~OTDR Length", "Meter Marks", "Eq Location", "EQ Type", "Activity",
            "Tray", "Slot", "Map It"
        ]
        det = details_df.reindex(columns=details_cols, fill_value="")

        det.to_excel(xl, sheet_name=ws2_name, index=False)
        ws2 = xl.sheets[ws2_name]
        ws2.set_column("A:A", 12)
        ws2.set_column("B:B", 120)
        ws2.set_column("M:M", 20)

        # Turn "Map It" strings into actual hyperlinks
        for i, url in enumerate(det["Map It"].tolist(), start=1):  # +1 for header row
            if isinstance(url, str) and url.startswith("http"):
                ws2.write_url(i, details_cols.index("Map It"), url, string="Google Maps")


def main():
    p = argparse.ArgumentParser()
    p.add_argument("--wo_csv", required=True, type=Path)
    p.add_argument("--circuit_json", required=True, type=Path)
    p.add_argument("--out_xlsx", required=True, type=Path)

    # Manual fields (left panel)
    p.add_argument("--designer_name", required=True)
    p.add_argument("--designer_email", required=True)
    p.add_argument("--designer_phone", required=True)
    p.add_argument("--fibers", required=True, type=int)
    p.add_argument("--order_id", required=True)
    p.add_argument("--client_name", required=True)
    p.add_argument("--build_type", required=True)
    p.add_argument("--a_end", required=True)
    p.add_argument("--z_end", required=True)
    p.add_argument("--route_type", required=True)
    p.add_argument("--circuit_version", required=True)
    p.add_argument("--circuit_id", required=True)
    p.add_argument("--jira", required=True)
    p.add_argument("--service_type", required=True)
    p.add_argument("--device_type", required=True)
    p.add_argument("--circuit_type", required=True)

    # Optional: if you already know these totals; otherwise leave blank
    p.add_argument("--end_to_end_m", default="")
    p.add_argument("--end_to_end_otdr_m", default="")

    args = p.parse_args()

    # Read WO CSV
    wo_df = _safe_read_csv(args.wo_csv)
    wo_kv = _kv_from_wo(wo_df)
    actions_df = _actions_from_wo(wo_df)

    # Parse JSON into a details table
    details_df = _details_from_json(args.circuit_json)

    meta = dict(
        designer_name=args.designer_name,
        designer_email=args.designer_email,
        designer_phone=args.designer_phone,
        fibers=args.fibers,
        order_id=args.order_id,
        client_name=args.client_name,
        build_type=args.build_type,
        a_end=args.a_end,
        z_end=args.z_end,
        route_type=args.route_type,
        circuit_version=args.circuit_version,
        circuit_id=args.circuit_id,
        jira=args.jira,
        service_type=args.service_type,
        device_type=args.device_type,
        circuit_type=args.circuit_type,
    )

    build_workbook(
        out_xlsx=args.out_xlsx,
        meta=meta,
        wo_kv=wo_kv,
        actions_df=actions_df,
        details_df=details_df,
        end_to_end_m=args.end_to_end_m,
        end_to_end_otdr_m=args.end_to_end_otdr_m,
    )
    print(f"✅ Wrote {args.out_xlsx}")


if __name__ == "__main__":
    main()
