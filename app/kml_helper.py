"""
Dynamic KML header/footer generator that adapts to your JSON
and produces colorized styles for Primary/Diverse/Triverse
routes + facility/plant icons.

Usage
-----
from dynamic_kml_header_footer import make_kml_header, make_kml_footer, pick_style_id

kml_header = make_kml_header(payload_json, defaults={
    "order_number": "ORDER-267175",
    "circuit_id": "CK37105",
    "client_name": "ABC In",
    "a_end": "BFMH-0021",
    "z_end": "SITE5761",
    "route_type": "Primary",  # Primary | Diverse | Triverse
    "service_type": "IRU Fbr",
})

# later when creating Placemarks
style_id = pick_style_id(facility="BFMH", ar=True, is_no_wo=False, route_type="Primary")
# -> use f"<styleUrl>#{style_id}</styleUrl>" in the Placemark

kml_footer = make_kml_footer()
"""

from __future__ import annotations
from xml.sax.saxutils import escape

# --- Color + icon presets ---
# KML color is AABBGGRR (little-endian)
ROUTE_LABEL_COLORS = {
    "Primary": "ffffff00",   # yellow labels/line
    "Diverse": "ff00a5ff",   # orange-ish
    "Triverse": "ff800080",  # purple
}
GRAY = "ff888888"
WHITE = "ffffffff"

ICON_BY_KEY = {
    ("BFMH", "UG"):   "http://maps.google.com/mapfiles/kml/shapes/square.png",
    ("THESMH", "UG"): "http://maps.google.com/mapfiles/kml/shapes/square.png",
    ("Pole", "AR"):   "http://maps.google.com/mapfiles/kml/shapes/placemark_circle.png",
}
DEFAULT_ICON = "http://maps.google.com/mapfiles/kml/shapes/placemark_circle.png"


def _get_meta(payload: dict, defaults: dict) -> dict:
    """Pulls useful bits from payload JSON, falling back to defaults/empty.
    You can extend the keys below if your JSON carries them under other paths.
    """
    meta = dict(defaults or {})
    # Common/expected keys if your pipeline already builds a bundle
    # like: {"metadata": {...}}
    md = (payload or {}).get("metadata", {}) if isinstance(payload, dict) else {}

    def pick(*keys, default=None):
        for k in keys:
            if k in md and md[k]:
                return md[k]
        return default

    meta.setdefault("order_number", pick("order_number", "WO"))
    meta.setdefault("circuit_id", pick("circuit_id", ""))
    meta.setdefault("client_name", pick("client_name", ""))
    meta.setdefault("a_end", pick("a_end", ""))
    meta.setdefault("z_end", pick("z_end", ""))
    meta.setdefault("service_type", pick("service_type", ""))
    meta.setdefault("route_type", pick("route_type", "Primary"))

    # Compose a readable <Document><name>
    parts = [
        meta.get("order_number"),
        meta.get("circuit_id"),
        meta.get("client_name"),
        f"A_{meta.get('a_end','')}" if meta.get("a_end") else None,
        f"Z_{meta.get('z_end','')}" if meta.get("z_end") else None,
        meta.get("route_type"),
        meta.get("service_type"),
    ]
    doc_name = "_".join([str(p).replace(" ", "") for p in parts if p]) or "WO"
    meta["doc_name"] = doc_name
    return meta


def _style_block(style_id: str, icon_href: str | None, label_color: str, line_color: str | None = None) -> str:
    """Builds a single <Style> block."""
    icon_xml = (
        f"""
  <IconStyle>
    <scale>1.1</scale>
    <Icon><href>{escape(icon_href or DEFAULT_ICON)}</href></Icon>
  </IconStyle>"""
        if icon_href else ""
    )
    line_xml = (
        f"""
  <LineStyle>
    <color>{line_color}</color>
    <width>4</width>
  </LineStyle>"""
        if line_color else ""
    )
    return f"""
  <Style id="{escape(style_id)}">
  <LabelStyle><color>{label_color}</color><scale>1.15</scale></LabelStyle>{icon_xml}{line_xml}
  </Style>"""


def _all_styles(route_type: str) -> str:
    """Returns all <Style> blocks needed for the map, colored by route_type."""
    label = ROUTE_LABEL_COLORS.get(route_type, WHITE)

    # Point styles by facility/plant (UG/AR) + a gray variant for No WO Activity
    blocks = []
    for fac in ["BFMH", "THESMH", "Pole", "Other"]:
        for plant in ["UG", "AR"]:
            icon = ICON_BY_KEY.get((fac, plant), DEFAULT_ICON)
            sid = f"pt_{fac}_{plant}"
            blocks.append(_style_block(sid, icon, label))
            blocks.append(_style_block(sid + "_muted", icon, GRAY))  # No WO Activity

    # Default fallback
    blocks.append(_style_block("pt_default", DEFAULT_ICON, label))
    blocks.append(_style_block("pt_default_muted", DEFAULT_ICON, GRAY))

    # Route line style (color matches labels)
    blocks.append(_style_block(f"route_{route_type.lower()}", None, label, line_color=label))
    return "\n".join(blocks)


def make_kml_header(payload_json: dict | None, defaults: dict | None = None) -> str:
    """Create a KML header string (<kml><Document>…styles…)
    using values discovered in payload_json (or provided via defaults).
    """
    meta = _get_meta(payload_json or {}, defaults or {})
    name = escape(meta.get("doc_name", "WO"))
    route_type = meta.get("route_type", "Primary")
    styles = _all_styles(route_type)

    return f"""<?xml version=\"1.0\" encoding=\"UTF-8\"?>
<kml xmlns=\"http://www.opengis.net/kml/2.2\"> 
<Document>
  <name>{name}</name>
{styles}
"""


def make_kml_footer() -> str:
    return """</Document>\n</kml>\n"""


def pick_style_id(*, facility: str | None, ar: bool | None, is_no_wo: bool, route_type: str) -> str:
    """Return the appropriate style id for a Placemark.
    facility: 'BFMH' | 'THESMH' | 'Pole' | other
    ar: aerial? True -> 'AR', False/None -> 'UG'
    is_no_wo: if True, returns the muted (gray) variant
    route_type is unused in the id (color comes from header styles), but kept for clarity.
    """
    fac = (facility or "Other").strip().upper()
    if fac not in {"BFMH", "THESMH", "POLE"}:
        fac = "Other"
    plant = "AR" if ar else "UG"
    base = f"pt_{fac}_{plant}" if fac != "Other" else "pt_default"
    return base + ("_muted" if is_no_wo else "")

