# activities.py
from __future__ import annotations
import io, json
from typing import Any, Dict, Iterable, List, Optional, Tuple
import pandas as pd
from PIL import Image, ImageDraw, ImageFont

# ---------- public ----------
def build_trace_booklet(csv_bytes: bytes,
                        json_bytes: bytes,
                        templates: Dict[str, bytes] | Dict[str, str],
                        fallback_pages: int = 8) -> io.BytesIO:
    """
    Build an XLSX booklet with a Cover page + one sheet per 'page/connection'.
    - Robustly discovers 'pages' even when JSON lacks 'Report: Splice Details'/'Connections'.
    - If nothing is discoverable, creates `fallback_pages` empty pages so the workbook still builds.
    - `templates`: mapping like {"400D5": "ref/400D.png", "450D6": "ref/450D.png"} (path or raw PNG/JPG bytes).
    Returns BytesIO of the xlsx.
    """
    # CSV not required today (kept for future enrichment)
    try:
        _ = pd.read_csv(io.BytesIO(csv_bytes), low_memory=False)
    except Exception:
        pass

    try:
        payload = json.loads(json_bytes.decode("utf-8", errors="ignore"))
    except Exception as e:
        raise ValueError(f"Unable to parse JSON: {e}")

    pages = _extract_pages(payload)
    if not pages:  # still nothing? make N empty pages to satisfy UX
        pages = [[] for _ in range(max(1, fallback_pages))]

    # build tables and select template for each page
    sheets: List[Tuple[str, pd.DataFrame, Optional[Tuple[str, bytes]]]] = []
    total = len(pages)
    for i, items in enumerate(pages, 1):
        df = _table_for_items(items)
        tpl_key = _guess_template_key(df, items)
        tpl_bytes = _load_template_bytes(templates, tpl_key)
        sheets.append((f"Page {i} of {total}", df, (tpl_key, tpl_bytes) if tpl_bytes else None))

    # ---- write workbook ----
    out = io.BytesIO()
    with pd.ExcelWriter(out, engine="xlsxwriter") as xw:
        wb = xw.book
        hdr = wb.add_format({"bold": True, "valign": "vcenter"})
        h1  = wb.add_format({"bold": True, "font_size": 18})
        normal = wb.add_format({"valign": "vcenter"})

        # Cover
        ws_cover = wb.add_worksheet("Cover")
        ws_cover.write(0, 0, "Trace Booklet", h1)
        ws_cover.write(2, 0, "Pages", hdr)
        for idx, (name, df, tpl) in enumerate(sheets, start=1):
            ws_cover.write_url(2 + idx, 0, f"internal:'{name}'!A13", string=name)
            ws_cover.write(2 + idx, 1, f"Rows: {len(df)}")
            if tpl and tpl[0]:
                ws_cover.write(2 + idx, 2, f"Template: {tpl[0]}")
        ws_cover.set_column(0, 2, 30)

        # show one copy of each known template on the cover (optional)
        col = 5
        for key, val in templates.items():
            try:
                img_bytes = val if isinstance(val, bytes) else open(val, "rb").read()
                ws_cover.insert_image(1, col, "tpl.png",
                    {"image_data": io.BytesIO(img_bytes), "x_scale": 0.3, "y_scale": 0.3})
                ws_cover.write(0, col, key, hdr)
                col += 3
            except Exception:
                pass

        # Page sheets
        for name, df, tpl in sheets:
            if df.empty:
                df = pd.DataFrame({"Item": [], "Detail": [], "lat": [], "lon": []})
            df.to_excel(xw, sheet_name=name, index=False, startrow=12)
            ws = xw.sheets[name]
            for cx, colname in enumerate(df.columns):
                ws.write(12, cx, colname, hdr)
            _auto_width(ws, df, start_row=12)

            # template image at top
            if tpl:
                title, img = tpl
                if img:
                    ws.insert_image(0, 0, "template.png",
                                    {"image_data": io.BytesIO(img), "x_scale": 0.6, "y_scale": 0.6})
                ws.write(0, 6, "Template:", hdr)
                ws.write(0, 7, title or "—", normal)

            # per-row Google Map thumbnail + text link
            lat_i = _find_col(df, ["lat", "latitude", "Lat"])
            lon_i = _find_col(df, ["lon", "longitude", "Lon"])
            if lat_i is not None and lon_i is not None:
                for r in range(len(df)):
                    try:
                        lat = float(df.iloc[r, lat_i]); lon = float(df.iloc[r, lon_i])
                    except Exception:
                        continue
                    url = f"https://www.google.com/maps?q={lat},{lon}"
                    thumb = _make_map_thumb(f"{lat:.6f},{lon:.6f}")
                    ws.insert_image(13 + r, len(df.columns), "map.png",
                        {"image_data": io.BytesIO(thumb), "x_scale": 0.6, "y_scale": 0.6, "url": url})
                    ws.write_url(13 + r, len(df.columns) + 1, url, string="Google Map")

    out.seek(0)
    return out

# ---------- page discovery ----------
def _extract_pages(obj: Any) -> List[List[Dict[str, Any]]]:
    """
    Try multiple strategies to discover 'pages' (each page is a list of item dicts):
    1) Look for keys containing both 'splice' and 'detail' → normalize any embedded 'Connections'.
    2) Recursively collect every list of dict items that look like connection items (Label/Detail/etc.)
       and treat *siblings* or *top-level arrays* as pages.
    3) If a single long list of items is found, split it into <=8 equal-ish pages.
    """
    # 1) Preferred: named sections (Report: Splice Details, etc.)
    sections = _find_key_sections(obj, must_have_all=["splice", "detail"])
    pages: List[List[Dict[str, Any]]] = []
    for sec in sections:
        pages.extend(_pages_from_section(sec))

    if pages:
        return [p for p in pages if p]  # non-empty only

    # 2) Heuristic discovery of item-lists anywhere
    candidate_lists = _collect_item_lists(obj)
    if len(candidate_lists) > 1:
        return [lst for lst in candidate_lists if lst]

    if len(candidate_lists) == 1:
        # 3) One long list → split up to 8 pages
        items = candidate_lists[0]
        if not items:
            return []
        k = min(8, max(1, (len(items) + 7) // 8))  # ~chunks aiming for ≤8 pages
        size = max(1, (len(items) + k - 1) // k)
        return [items[i:i+size] for i in range(0, len(items), size)]

    return []

def _pages_from_section(section: Any) -> List[List[Dict[str, Any]]]:
    """
    Normalize common shapes inside a 'Splice Details' section to a list-of-pages.
    Accepts:
      - [{ "": [ {"Connections":[...]}, ... ]}, {"Connections":[...]}]
      - {"Connections":[...]}, {"Connections":{...}}
      - Arbitrary dict/list mixing — we dig until item lists are found.
    """
    pages: List[List[Dict[str, Any]]] = []
    if isinstance(section, list):
        for block in section:
            pages.extend(_pages_from_section(block))
        return pages

    if isinstance(section, dict):
        # embedded list under empty key (very common)
        if "" in section and isinstance(section[""], list):
            for inner in section[""]:
                pages.extend(_pages_from_section(inner))

        # direct connections
        if "Connections" in section:
            items = section["Connections"]
            items = items if isinstance(items, list) else [items]
            pages.append([it for it in items if isinstance(it, dict)])
            # keep walking other values too (more pages possible)
        for v in section.values():
            pages.extend(_pages_from_section(v))

    # scalars ignored
    return pages

def _collect_item_lists(obj: Any) -> List[List[Dict[str, Any]]]:
    """
    Recursively gather lists that look like lists of item dicts
    (dict containing any of: Label/Part/ID + Detail/Description/Text).
    """
    out: List[List[Dict[str, Any]]] = []
    if isinstance(obj, list):
        if obj and all(isinstance(el, dict) for el in obj) and any(_looks_like_item_dict(el) for el in obj):
            out.append([el for el in obj if isinstance(el, dict)])
        else:
            for el in obj:
                out.extend(_collect_item_lists(el))
    elif isinstance(obj, dict):
        for v in obj.values():
            out.extend(_collect_item_lists(v))
    return out

def _looks_like_item_dict(d: Dict[str, Any]) -> bool:
    if not isinstance(d, dict):
        return False
    keys = {k.lower() for k in d.keys()}
    has_labelish = any(k in keys for k in ["label","part","id","item"])
    has_detailish = any(k in keys for k in ["detail","description","text"])
    return has_labelish or has_detailish

def _find_key_sections(obj: Any, must_have_all: List[str]) -> List[Any]:
    """Find dict values whose key contains ALL tokens in `must_have_all` (case-insensitive)."""
    hits: List[Any] = []
    if isinstance(obj, dict):
        for k, v in obj.items():
            key = str(k).lower()
            if all(tok in key for tok in must_have_all):
                hits.append(v)
            hits.extend(_find_key_sections(v, must_have_all))
    elif isinstance(obj, list):
        for el in obj:
            hits.extend(_find_key_sections(el, must_have_all))
    return hits

# ---------- tables / template detection / utils ----------
def _table_for_items(items: List[Dict[str, Any]]) -> pd.DataFrame:
    rows = []
    for it in items:
        if not isinstance(it, dict):
            continue
        rows.append({
            "Item": str(it.get("Label") or it.get("Part") or it.get("ID") or ""),
            "Detail": str(it.get("Detail") or it.get("Description") or it.get("Text") or ""),
            "lat": _try_float(it.get("lat") or it.get("Lat") or it.get("Latitude")),
            "lon": _try_float(it.get("lon") or it.get("Lon") or it.get("Longitude")),
        })
    df = pd.DataFrame(rows)
    for c in ["Item", "Detail", "lat", "lon"]:
        if c not in df.columns: df[c] = None
    return df[["Item", "Detail", "lat", "lon"]]

def _guess_template_key(df: pd.DataFrame, items: List[Dict[str, Any]]) -> Optional[str]:
    text = " ".join(map(str, df["Detail"].dropna().tolist())).upper()
    for key in ["600D","450D6","450D","450B","400B6","400A8","400D5","400D","200D","200B","PLP","FEC","WM ACCUM"]:
        if key in text: return key
    # last chance: look at any string fields from items
    for it in items:
        for v in it.values():
            if isinstance(v, str):
                up = v.upper()
                for key in ["600D","450D6","450D","450B","400B6","400A8","400D5","400D","200D","200B","PLP","FEC","WM ACCUM"]:
                    if key in up: return key
    return None

def _load_template_bytes(templates: Dict[str, bytes] | Dict[str, str], key: Optional[str]) -> Optional[bytes]:
    if not key or key not in templates: return None
    val = templates[key]
    if isinstance(val, bytes): return val
    with open(val, "rb") as f: return f.read()

def _find_col(df: pd.DataFrame, candidates: Iterable[str]) -> Optional[int]:
    lower = {c.lower(): i for i, c in enumerate(df.columns)}
    for cand in candidates:
        if cand.lower() in lower: return lower[cand.lower()]
    return None

def _try_float(x):
    try: return float(x)
    except Exception: return None

def _auto_width(ws, df: pd.DataFrame, start_row: int = 0):
    for colx, col in enumerate(df.columns):
        maxlen = len(str(col))
        for val in df[col].astype(str).head(200):
            maxlen = max(maxlen, len(val))
        ws.set_column(colx, colx, min(60, max(10, int(maxlen * 1.1))))

def _make_map_thumb(text: str) -> bytes:
    W, H = 180, 120
    img = Image.new("RGB", (W, H), (230, 235, 240))
    d = ImageDraw.Draw(img)
    d.ellipse((78, 30, 102, 54), fill=(200, 50, 60))
    d.rectangle((88, 54, 92, 85), fill=(200, 50, 60))
    try: font = ImageFont.load_default()
    except Exception: font = None
    d.text((8, 95), f"{text}", fill=(0, 0, 0), font=font)
    bio = io.BytesIO(); img.save(bio, format="PNG"); bio.seek(0); return bio.getvalue()
