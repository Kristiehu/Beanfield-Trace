from __future__ import annotations
import re, io, json
from dataclasses import dataclass
from typing import Dict, Any, List, Optional, Tuple
import pandas as pd

@dataclass
class CleanResult:
    order_id: Optional[str]
    clean_json: Dict[str, Any]
    nodes_df: pd.DataFrame
    cables_df: pd.DataFrame
    events_df: pd.DataFrame
    path_list: List[str]
    integrity: Dict[str, Any]

def run_clean(raw_json_bytes: bytes) -> CleanResult:
    raw = json.loads(raw_json_bytes.decode("utf-8", errors="ignore"))
    order_id = _guess_order_id(raw)
    blocks = _find_connections_recursive(raw)
    segments: List[str] = []
    for blk in blocks:
        norm = _normalize_tokens(blk)
        segments.extend(_segment(norm))

    parsed = _parse_segments_with_context(segments)

    nodes_df  = pd.DataFrame(parsed["nodes"]) if parsed["nodes"] else pd.DataFrame(columns=["node_id","name","type","lat","lon","address"])
    cables_df = pd.DataFrame(parsed["cables"]) if parsed["cables"] else pd.DataFrame(columns=["pmid","aptum_id","fibres","from_node","from_box","to_node","to_box","length_m","length_raw","otdr_m"])
    events_df = pd.DataFrame(parsed["events"]) if parsed["events"] else pd.DataFrame(columns=["type","pmid_left","pmid_right","aptum_left","aptum_right"])

    integrity = _integrity_report(parsed["nodes"], parsed["cables"], parsed["events"])

    clean = {
        "order_id": order_id,
        "a_end": parsed["a_end"],
        "z_end": parsed["z_end"],
        "nodes": parsed["nodes"],
        "cables": parsed["cables"],
        "events": parsed["events"],
        "path": parsed["path"],
        "integrity": integrity,
    }

    return CleanResult(order_id, clean, nodes_df, cables_df, events_df, parsed["path"], integrity)

# ---------------- core helpers ----------------

def _guess_order_id(raw: Dict[str, Any]) -> Optional[str]:
    try_text = json.dumps(raw)[:5000]
    m = re.search(r"(CHGMGT-\d{4,6}|ORDER-\d+)", try_text, re.I)
    return m.group(1) if m else None

def _find_connections_recursive(obj) -> List[str]:
    out: List[str] = []
    if isinstance(obj, dict):
        for k, v in obj.items():
            if k == "Connections" and isinstance(v, str):
                out.append(v)
            else:
                out.extend(_find_connections_recursive(v))
    elif isinstance(obj, list):
        for it in obj:
            out.extend(_find_connections_recursive(it))
    return out

def _normalize_tokens(s: str) -> str:
    s = s.replace("<OPEN>", "").replace("<CLOSE>", "")
    s = s.replace("<COMMA>", ",").replace("<COLON>", ":")
    s = s.replace("\r", "")
    return s

def _segment(text: str) -> List[str]:
    # split on blank lines or dotted separators, but keep order
    parts = re.split(r"(?:\n\s*\.\s*\n)|(?:\n{2,})", text)
    parts = [p.strip() for p in parts if p.strip()]
    return parts

RE_NODE = re.compile(
    r"^(?P<city>[A-Za-z]+)\s*,\s*(?P<name>[^,:]+)\s*,\s*Address:.*?:\s*\(?(?P<lat>-?\d+(?:\.\d+)?)\s*,\s*(?P<lon>-?\d+(?:\.\d+)?)\)?\s*:?\s*:?\s*(?P<stype>.+)$"
)

RE_BOX = re.compile(
    r"^(?P<box_id>[^:\n]+)\s*:\s*OSP Splice Box\s*-\s*(?P<range>[^:\n]+)\s*:"
)

RE_CA = re.compile(
    r"^CA(?P<idx>\d+)\s*:\s*PMID\s*:\s*(?P<pmid>\d+)\s*,\s*(?P<fibres>\d+)\s*(?:R?F?)\s*,\s*(?P<to_node>[^,\n]+)(?:\s*,\s*(?P<to_box>[^,\n]+))?"
)

RE_EVENT = re.compile(
    r"PMID:\s*(?P<l>\d+)\s*,.*?--\s*Splice\s*--\s*PMID:\s*(?P<r>\d+)\s*,", re.I|re.S
)

RE_LEN = re.compile(
    r"PMID[: ]\s*(?P<pmid>\d+).*?(?:Cable Length(?: Calculated)?:\s*(?P<len>[\d.]+)\s*m)", re.I
)
RE_OTDR = re.compile(r"OTDR\s*(?P<otdr>[\d.]+)\s*m", re.I)

def _parse_segments_with_context(segments: List[str]) -> Dict[str, Any]:
    nodes: List[Dict[str, Any]] = []
    boxes: List[Dict[str, Any]] = []
    cables: List[Dict[str, Any]] = []
    events: List[Dict[str, Any]] = []
    a_end = {}
    z_end = {}

    current_node: Optional[str] = None
    current_box: Optional[str] = None

    for seg in segments:
        lines = [l.strip() for l in seg.splitlines() if l.strip() and l.strip() != "."]
        if not lines:
            continue

        # node header (may be first line)
        header = lines[0]
        mnode = RE_NODE.match(header)
        if mnode:
            name = mnode.group("name").strip()
            node_id = name if name.startswith("PA") else f"PA{name}"
            try:
                lat = float(mnode.group("lat")); lon = float(mnode.group("lon"))
            except Exception:
                lat = lon = None
            nodes.append({
                "node_id": node_id,
                "name": name,
                "type": mnode.group("stype").strip(),
                "lat": lat, "lon": lon, "address": ""
            })
            current_node = node_id
            current_box = None

            if re.search(r"\bVault\b", seg, re.I) and not a_end:
                a_end = {"site": name}
            z_end = {"site": name}

            lines = lines[1:]  # continue parsing rest of the segment

        # parse boxes + CA lines regardless of header presence (node context carries over)
        for l in lines:
            mb = RE_BOX.match(l)
            if mb and current_node:
                current_box = mb.group("box_id").strip()
                boxes.append({
                    "node_id": current_node,
                    "box_id": current_box,
                    "model": "OSP Splice Box",
                    "range": mb.group("range").strip()
                })
                continue

            mca = RE_CA.match(l)
            if mca:
                pmid = mca.group("pmid")
                fibres = int(mca.group("fibres"))
                to_node = (mca.group("to_node") or "").strip().replace(" ", "")
                to_box  = (mca.group("to_box") or "").strip()
                if to_node and not to_node.startswith("PA"):
                    if re.fullmatch(r"\d{3,6}", to_node):
                        to_node = "PA" + to_node
                cables.append({
                    "pmid": pmid, "aptum_id": None, "fibres": fibres,
                    "from_node": current_node, "from_box": current_box,
                    "to_node": to_node or None, "to_box": to_box or None,
                    "length_m": None, "length_raw": None, "otdr_m": None
                })

        # splice events + lengths
        me = RE_EVENT.search(seg)
        if me:
            events.append({
                "type": "splice",
                "pmid_left": me.group("l"),
                "pmid_right": me.group("r"),
                "aptum_left": None, "aptum_right": None
            })

        for l in lines:
            mlen = RE_LEN.search(l)
            if mlen:
                pmid = mlen.group("pmid")
                length = float(mlen.group("len"))
                for c in reversed(cables):
                    if c["pmid"] == pmid and c["length_m"] is None:
                        c["length_m"] = length
                        c["length_raw"] = l
                        motdr = RE_OTDR.search(l)
                        if motdr:
                            c["otdr_m"] = float(motdr.group("otdr"))
                        break

    # build simple path
    path = []
    for c in cables:
        if c["from_node"]:
            hop = f"{c['from_node']}/{c['from_box'] or ''}".rstrip("/")
            if not path or path[-1] != hop:
                path.append(hop)
        if c["to_node"]:
            hop = f"{c['to_node']}/{c['to_box'] or ''}".rstrip("/")
            if not path or path[-1] != hop:
                path.append(hop)

    return {
        "nodes": nodes, "boxes": boxes, "cables": cables, "events": events,
        "a_end": a_end, "z_end": z_end, "path": path
    }

def _integrity_report(nodes, cables, events) -> Dict[str, Any]:
    node_ids = {n["node_id"] for n in nodes}
    orphan_from = [c for c in cables if not c["from_node"]]
    orphan_to   = [c for c in cables if c["to_node"] and c["to_node"] not in node_ids]
    missing_len = [c["pmid"] for c in cables if c["length_m"] is None and c["otdr_m"] is None]
    return {
        "counts": {"nodes": len(nodes), "cables": len(cables), "events": len(events)},
        "orphans_from": len(orphan_from),
        "orphans_to": len(orphan_to),
        "cables_missing_length": len(set(missing_len)),
    }

def export_bytes(result: CleanResult) -> Dict[str, bytes]:
    artifacts: Dict[str, bytes] = {}

    def to_csv_bytes(df: pd.DataFrame) -> bytes:
        bio = io.StringIO(); df.to_csv(bio, index=False); return bio.getvalue().encode("utf-8")

    artifacts[_fname(result.order_id, "clean.json")] = json.dumps(result.clean_json, ensure_ascii=False, indent=2).encode("utf-8")
    artifacts[_fname(result.order_id, "nodes.csv")]  = to_csv_bytes(result.nodes_df)
    artifacts[_fname(result.order_id, "cables.csv")] = to_csv_bytes(result.cables_df)
    artifacts[_fname(result.order_id, "events.csv")] = to_csv_bytes(result.events_df)

    printable = build_print_table(result)
    artifacts[_fname(result.order_id, "trace_table.csv")] = to_csv_bytes(printable)

    return artifacts

def build_print_table(result: CleanResult) -> pd.DataFrame:
    rows = [{"#": i, "Hop": hop} for i, hop in enumerate(result.path_list, start=1)]
    df = pd.DataFrame(rows, columns=["#","Hop"])
    return df

def _fname(order_id: Optional[str], suffix: str) -> str:
    prefix = (order_id or "order").replace("/", "_")
    return f"{prefix}.{suffix}"
