import json, re, pandas as pd, os, math
from pathlib import Path
from xml.sax.saxutils import escape
from kml_helper import make_kml_header, make_kml_footer, pick_style_id

from xml.sax.saxutils import escape
def _esc(x):  # safe XML text
    return escape("" if x is None else str(x))


json_path = Path("data/WO24218.json")
csv_path = Path("data/WO24218.csv")
out_path = Path("output/WO24218.kml")


def _esc(x):
    return escape("" if x is None else str(x))

def to_kml(title, placemarks):
    lines = []
    lines.append("<?xml version='1.0' encoding='UTF-8'?>")
    lines.append("<kml xmlns='http://www.opengis.net/kml/2.2'>")
    lines.append("<Document>")
    lines.append(f"  <name>{_esc(title)}</name>")

    for pm in placemarks:
        pm_name = _esc(pm.get("name", ""))
        lat = pm.get("lat")
        lon = pm.get("lon")
        desc = pm.get("description", "") or ""

        lines.append("  <Placemark>")
        lines.append(f"    <name>{pm_name}</name>")
        # Keep description inside CDATA so HTML is safe
        safe_desc = desc.replace("]]>", "]]]><![CDATA[>")
        lines.append(f"    <description><![CDATA[{safe_desc}]]></description>")
        if lat is not None and lon is not None:
            lines.append("    <Point>")
            lines.append(f"      <coordinates>{lon},{lat},0</coordinates>")
            lines.append("    </Point>")
        lines.append("  </Placemark>")

    lines.append("</Document>")
    lines.append("</kml>")
    return "\n".join(lines)


def read_json_texts(p):
    texts = []
    if not p.exists():
        return texts
    with open(p, "r", encoding="utf-8") as f:
        data = json.load(f)
    # Walk the JSON and collect any long strings that might contain address/lat/lon patterns
    def walk(x):
        if isinstance(x, dict):
            for v in x.values():
                walk(v)
        elif isinstance(x, list):
            for v in x:
                walk(v)
        elif isinstance(x, str):
            if "Address:" in x and (":" in x or "," in x):
                texts.append(x)
    walk(data)
    return texts

def clean_tokens(s: str) -> str:
    rep = {
        "<COMMA>": ",",
        "<AND>": "&",
        "<OPEN>": "(",
        "<CLOSE>": ")",
        "\u00A0": " ",
    }
    for k, v in rep.items():
        s = s.replace(k, v)
    return s

# regex for lines that look like:
# City,Site, Address:... ,() : 43.644719, -79.385046 :  : Feature description
header_re = re.compile(
    r"^\s*(?P<city>[^,\n]+)\s*,\s*(?P<site>[^,\n]+)\s*,\s*Address:(?P<addr>.*?),\(\)\s*:\s*(?P<lat>-?\d+\.\d+)\s*,\s*(?P<lon>-?\d+\.\d+)\s*:\s*:\s*(?P<feat>[^\n]+)\s*$",
    re.MULTILINE
)

def prox_key(lat, lon, name):
    # round to ~6 decimals to avoid dup jitter
    return (round(float(lat), 6), round(float(lon), 6), (name or "").strip().lower())

points = []

# Parse JSON
json_points = []
texts = read_json_texts(json_path)
for t in texts:
    t = clean_tokens(t)
    for m in header_re.finditer(t):
        d = m.groupdict()
        item = {
            "source": "json",
            "name": d["site"].strip(),
            "city": d["city"].strip(),
            "address": d["addr"].strip(),
            "feature": d["feat"].strip(),
            "lat": float(d["lat"]),
            "lon": float(d["lon"]),
        }
        json_points.append(item)

# Deduplicate JSON points
seen = set()
dedup_json = []
for it in json_points:
    key = prox_key(it["lat"], it["lon"], it["name"])
    if key in seen:
        continue
    seen.add(key)
    dedup_json.append(it)

# Parse CSV in a flexible way
csv_points = []
if csv_path.exists():
    df = pd.read_csv(csv_path)
    # Identify lat/lon columns
    cols = {c.lower(): c for c in df.columns}
    latcol = next((cols[c] for c in cols if "lat" in c and not any(b in c for b in ["plate", "relat"])), None)
    loncol = next((cols[c] for c in cols if any(k in c for k in ["lon", "lng"]) and "along" not in c), None)
    # fallbacks common in GIS exports
    if latcol is None and "y" in cols: latcol = cols["y"]
    if loncol is None and "x" in cols: loncol = cols["x"]
    # pick a name-ish column
    namecol = None
    for candidate in ["name", "site", "id", "title", "label", "asset", "node", "location", "address", "desc"]:
        if candidate in cols:
            namecol = cols[candidate]
            break
            
    # build points if we have coordinates
    if latcol is not None and loncol is not None:
        for _, row in df.iterrows():
            try:
                lat = float(row[latcol])
                lon = float(row[loncol])
            except Exception:
                continue
            if not (math.isfinite(lat) and math.isfinite(lon)):
                continue
            name = str(row[namecol]) if namecol is not None and not pd.isna(row[namecol]) else ""
            addr = str(row["Address"]) if "Address" in df.columns and not pd.isna(row["Address"]) else ""
            feat = ""
            csv_points.append({
                "source": "csv",
                "name": name.strip() or "CSV Item",
                "city": "",
                "address": addr.strip(),
                "feature": feat,
                "lat": lat,
                "lon": lon,
            })

# Merge and de-duplicate
all_points = []
seen = set()
for it in (dedup_json + csv_points):
    key = prox_key(it["lat"], it["lon"], it["name"])
    if key in seen:
        continue
    seen.add(key)
    all_points.append(it)

# Minimal KML templates
kml_header = """<?xml version="1.0" encoding="UTF-8"?>
<kml xmlns="http://www.opengis.net/kml/2.2">
<Document>
  <name>WO24218</name>
  <Style id="pointStyle">
    <IconStyle>
      <scale>1.1</scale>
      <Icon>
        <href>http://maps.google.com/mapfiles/kml/paddle/red-circle.png</href>
      </Icon>
    </IconStyle>
  </Style>
"""

kml_footer = """</Document>
</kml>
"""

def escape(s):
    if s is None:
        return ""
    # basic XML escape
    return (str(s)
            .replace("&", "&amp;")
            .replace("<", "&lt;")
            .replace(">", "&gt;"))

placemarks = []
for pt in all_points:
    name = escape(pt["name"] or f"{pt['feature'] or 'Point'}")
    desc_lines = []
    if pt["feature"]:
        desc_lines.append(f"<b>Feature</b>: {escape(pt['feature'])}")
        
    if pt["address"]:
        desc_lines.append(f"<b>Address</b>: {escape(pt['address'])}")
    if pt["city"]:
        desc_lines.append(f"<b>City</b>: {escape(pt['city'])}")
    desc_lines.append(f"<b>Source</b>: {pt['source'].upper()}")
    desc = "<br/>".join(desc_lines)
    placemark = f"""  <Placemark>
    <name>{name}</name>
    <styleUrl>#pointStyle</styleUrl>
    <description><![CDATA[{desc}]]></description>
    <Point><coordinates>{pt['lon']},{pt['lat']},0</coordinates></Point>
  </Placemark>"""
    placemarks.append(placemark)

kml_text = kml_header + "\n".join(placemarks) + "\n" + kml_footer
with open(out_path, "w", encoding="utf-8") as f:
    f.write(kml_text)

len_json = len(dedup_json)
len_csv = len(csv_points)
len_all = len(all_points)


print(f"Parsed {len_json} unique points from JSON, {len_csv} from CSV; wrote {len_all} unique placemarks to {out_path.name}.")

