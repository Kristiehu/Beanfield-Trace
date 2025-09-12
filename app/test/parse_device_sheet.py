# -*- coding: utf-8 -*-
"""
parse_device_sheet.py
Usage:
  python parse_device_sheet.py --json WO24218.json --out devices.csv --append-details
Reads a JSON with structure like:
{"Report: Splice Details":[{"":[{"Connections":"..."}]}]}
Parses the "Connections" text and produces a table:
Device Name | Device Type | Type | Lat / Long | UG | AR | Activity | Dev Order
"""
import re, json, argparse, pandas as pd

HEADER_RE = re.compile(r'^[^,]+,\s*([^,]+),\s*Address:.*?:\s*(-?\d{1,2}\.\d+),\s*(-?\d{1,3}\.\d+)\s*:\s*:\s*(.+)$')
DEV_FULL_RE = re.compile(r'^\s*(.+?)\s*:\s*OSP Splice Box\s*-\s*([A-Za-z0-9]+)\s*:\s*(.+)$')
DEV1_RE     = re.compile(r'^\s*(.+?)\s*:\s*OSP Splice Box\s*-\s*([A-Za-z0-9]+)\s*:')
GEN_FULL_RE = re.compile(r'^\s*(.+?)\s*:\s*Generic OSP Splice Box(?:\s*-\s*([A-Za-z0-9]+))?\s*:\s*(.+)$')
GEN_RE      = re.compile(r'^\s*(.+?)\s*:\s*Generic OSP Splice Box(?:\s*-\s*([A-Za-z0-9]+))?\s*:')
SS_FULL_RE  = re.compile(r'^\s*(.+?)\s*:\s*SS\b.*?:\s*(.+)$')
SS_RE       = re.compile(r'^\s*(.+?)\s*:\s*SS\b.*:')
_COORD_RE = re.compile(r"\b-?\d{1,3}\.\d+\s*,\s*-?\d{1,3}\.\d+\b")

def clean_text(t: str) -> str:
    t=(t or '')
    t=t.replace('<COMMA>',',').replace('<COLON>',':').replace('<OPEN>','(').replace('<CLOSE>',')').replace('<AND>','&')
    t=re.sub(r'\s+', ' ', t)
    return t.strip()

def classify_site(site_type: str):
    s=(site_type or '').lower()
    if 'beanfield manhole/handwell' in s: return 'BFMH',100,0
    if 'hydro manhole' in s: return 'THESMH',100,0
    if 'utility pole' in s or 'pole' in s: return 'Pole',0,100
    return 'Unknown',0,0

# parse_device_sheet.py
def _looks_like_connection_blob(s: str) -> bool:
    """
    Heuristics to accept either the explicit 'Connections' text OR the raw
    device/connection blob commonly found under ["Report: Splice Details"][0][""][0].
    """
    if not s:
        return False
    s = s.strip()

    # Strong indicators
    if "Address:" in s:
        return True
    if "Splice" in s or "FOSC" in s or "PMID" in s:
        return True
    if _COORD_RE.search(s):  # looks like '43.1, -79.2'
        return True

    # Gentle fallback: moderately sized text with structure
    if len(s) >= 40 and s.count(":") >= 2:
        return True

    return False

def gather_connections(obj):
    """
    Return a list of candidate 'Connections' text blobs from the circuit JSON.
    Supports two common shapes:
      A) {"Report: Splice Details":[{"":[{"Connections":"...text..."}]}]}
      B) {"Report: Splice Details":[{"":["...text..."]}]}
    Falls back to strings that look like connection/device blobs.
    """
    out = []

    def rec(x):
        if isinstance(x, dict):
            for k, v in x.items():
                # explicit field
                if k == "Connections" and isinstance(v, str) and v.strip():
                    out.append(v)
                    continue
                # common containers under the target report
                if k in ("Report: Splice Details", ""):
                    rec(v)
                    continue
                rec(v)
        elif isinstance(x, list):
            for it in x:
                rec(it)
        elif isinstance(x, str):
            if _looks_like_connection_blob(x):
                out.append(x)

    rec(obj)

    # Deduplicate while preserving order
    seen = set()
    deduped = []
    for s in out:
        t = s.strip()
        if t and t not in seen:
            seen.add(t)
            deduped.append(t)
    return deduped

def parse_device_table(connections_text: str, append_details=False) -> pd.DataFrame:
    lat=lon=None; site_type=None
    rows=[]
    for raw in connections_text.splitlines():
        line=raw.strip()
        if not line or line.startswith('CA') or line.startswith('PMID') or line.startswith('.') or line.startswith('Presented by'):
            continue
        # Header?
        m=HEADER_RE.match(line)
        if m:
            lat=float(m.group(2)); lon=float(m.group(3)); site_type=m.group(4)
            continue
        # Device Patterns
        matched=False
        for (regex, detail_group, devtype_group, dtype_default, box_kind) in [
            (DEV_FULL_RE, 3, 2, None, 'FOSC'),
            (DEV1_RE,     None, 2, None, 'FOSC'),
            (GEN_FULL_RE, 3, 2, 'Unknown', 'FOSC'),
            (GEN_RE,      None, 2, 'Unknown', 'FOSC'),
            (SS_FULL_RE,  2, None, 'SS', 'SS/Coil'),
            (SS_RE,       None, None, 'SS', 'SS/Coil'),
        ]:
            m=regex.match(line)
            if not m:
                continue
            name=clean_text(m.group(1))
            devtype=clean_text(m.group(devtype_group)) if (devtype_group and m.group(devtype_group)) else dtype_default or 'Unknown'
            if append_details and detail_group and m.group(detail_group):
                details=clean_text(m.group(detail_group))
                name=f"{name}: {details}"
            typ,UG,AR = classify_site(site_type)
            rows.append([name, devtype, typ, f"{lat},{lon}", UG, AR, "", None])
            matched=True
            break
        # If no regex matched, ignore the line
    for i,row in enumerate(rows, start=1):
        row[-1]=i
    return pd.DataFrame(rows, columns=['Device Name','Device Type','Type','Lat / Long','UG','AR','Activity','Dev Order'])


# helper for connection errors ---
import re
from typing import Any, Dict, List, Tuple

def _normalize_key(k: str) -> str:
    return re.sub(r'[^a-z0-9]+', ' ', str(k).lower()).strip()

def _find_key_ci(d: Dict[str, Any], needle_substr: str) -> str | None:
    ns = _normalize_key(needle_substr)
    for k in d.keys():
        if ns in _normalize_key(k):
            return k
    return None

def _as_list(x: Any) -> List[Any]:
    if x is None:
        return []
    if isinstance(x, list):
        return x
    if isinstance(x, dict):
        for key_guess in ("items", "rows", "data", "list", "values"):
            v = x.get(key_guess)
            if isinstance(v, list):
                return v
    return [x]

def _extract_connections_from_record(rec: Dict[str, Any]) -> Any:
    if not isinstance(rec, dict):
        return None
    for key_guess in ("Connections", "Connection", "connections", "connection", "links", "splices"):
        if key_guess in rec:
            return rec[key_guess]
        k_ci = _find_key_ci(rec, key_guess)
        if k_ci is not None:
            return rec[k_ci]
    return None

def extract_splice_connections(payload: Dict[str, Any]) -> Tuple[List[Dict[str, Any]], str]:
    """
    Find 'Splice Details' → connections even if the schema varies.
    Returns (connections_list, warn_msg). warn_msg == '' when all good.
    """
    if not isinstance(payload, dict):
        return [], "JSON payload is not an object."

    # Some dumps have a top-level 'Report' container; some don't.
    top = payload
    k_report = _find_key_ci(payload, "report")
    if k_report and isinstance(payload.get(k_report), dict):
        top = payload[k_report]

    if not isinstance(top, dict):
        return [], "JSON 'Report' is not an object."

    # Find the section that looks like 'Splice Details' (or just 'Splice')
    k_splice = _find_key_ci(top, "splice details") or _find_key_ci(top, "splice")
    if not k_splice:
        return [], "No 'Splice Details' section found in JSON."

    section = top[k_splice]

    # Try direct connections at section level
    connections: List[Dict[str, Any]] = []
    if isinstance(section, dict):
        found = _extract_connections_from_record(section)
        if found is not None:
            connections.extend(_as_list(found))

    # Try rows/items list under the section
    for rec in _as_list(section):
        if isinstance(rec, dict):
            found = _extract_connections_from_record(rec)
            if found is not None:
                connections.extend(_as_list(found))

    # Last resort: deep scan for *connection* keys anywhere under section
    if not connections:
        stack = [section]
        while stack:
            node = stack.pop()
            if isinstance(node, dict):
                for k, v in node.items():
                    if "connection" in _normalize_key(k):
                        connections.extend(_as_list(v))
                    elif isinstance(v, (dict, list)):
                        stack.append(v)
            elif isinstance(node, list):
                stack.extend(node)

    if connections:
        # Normalize to list of dicts
        norm = []
        for c in connections:
            if isinstance(c, dict):
                norm.append(c)
            else:
                norm.append({"value": c})
        return norm, ""

    return [], "No 'Connections' strings found under the 'Splice Details' section."

# ---



def main():
    ap=argparse.ArgumentParser()
    ap.add_argument('--json', required=True, help='Input JSON file')
    ap.add_argument('--out', default='Activity Overview Map.csv', help='Output CSV path')
    ap.add_argument('--append-details', action='store_true', help='Append trailing details after device name (to match some samples)')
    args=ap.parse_args()
    with open(args.json, 'r', encoding='utf-8') as f:
        data=json.load(f)
    # There may be multiple "Connections"; join them (often duplicates)
    conns=gather_connections(data)
    if not conns:
        raise SystemExit("No 'Connections' strings found under 'Report: Splice Details'.")
    # Use the first unique string
    conn_text=conns[0]
    df=parse_device_table(conn_text, append_details=args.append_details)
    df.to_csv(args.out, index=False, encoding='utf-8-sig')
    print(f"Wrote {len(df)} rows to {args.out}")

if __name__=='__main__':
    main()
