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

def gather_connections(obj):
    out=[]
    def rec(x):
        if isinstance(x, dict):
            for k,v in x.items():
                if k=='Connections' and isinstance(v,str):
                    out.append(v)
                else:
                    rec(v)
        elif isinstance(x, list):
            for it in x:
                rec(it)
    rec(obj)
    return out

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
