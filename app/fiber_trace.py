# fiber_trace.py — Build Fibre Trace and merge actions (from fibre_action.compute_fibre_actions)
from __future__ import annotations
from pathlib import Path
from typing import Dict, List, Optional, Tuple, Union, IO
import io, json, re
import pandas as pd

# Public columns (sheet schema)
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
    "Google Map",
]

# ---------- helpers ----------
def _decode(s: str) -> str:
    if not isinstance(s, str): return ""
    return (s.replace("<COMMA>", ",").replace("<COLON>", ":")
              .replace("<OPEN>", "(").replace("<CLOSE>", ")")
              .replace("<AND>", "&").replace("\\n", "\n").strip())

_HDR  = re.compile(r"^\s*([^,]+)\s*,\s*([^,]+)\s*,\s*Address:(.*?)\(\)\s*:\s*(-?\d+(?:\.\d+)?)\s*,\s*(-?\d+(?:\.\d+)?)\s*:\s*:\s*(.+?)\s*$")
_PMID = re.compile(r"\b(?:n?PMID)\s*[: ]\s*([A-Za-z0-9_-]+)", re.I)
_FNUM = re.compile(r"\.F\s*([0-9]+)")
_CALC = re.compile(r"(?:Calculated|Cable\s*Length\s*Calculated)\s*:\s*([0-9]+)", re.I)
_OTDR = re.compile(r"OTDR\s*[:; -]\s*([0-9]+)", re.I)

def _pmids(s: str) -> List[str]:
    return [m.group(1).upper() for m in _PMID.finditer(s or "")]

def _compact(nums: List[int]) -> str:
    s = sorted({int(n) for n in nums if str(n).isdigit()})
    if not s: return ""
    out=[]; a=b=s[0]
    for n in s[1:]:
        if n==b+1: b=n
        else: out.append((a,b)); a=b=n
    out.append((a,b))
    return ",".join(str(x) if x==y else f"{x}-{y}" for x,y in out)

def _get_connection_blobs(payload: dict) -> List[str]:
    blobs=[]
    try:
        arr = payload.get("Report: Splice Details", [])
        if isinstance(arr, list) and arr:
            inner = arr[0].get("", [])
            if isinstance(inner, list):
                for it in inner:
                    s = it.get("Connections")
                    if isinstance(s, str) and s.strip():
                        blobs.append(s)
    except Exception:
        pass
    seen=set(); out=[]
    for s in blobs:
        if s not in seen:
            out.append(s); seen.add(s)
    return out

def _split(blob: str) -> List[List[str]]:
    s = _decode(blob)
    parts = re.split(r"\n\.\n", s.strip())
    sections=[]
    for p in parts:
        if not p.strip(): 
            continue
        lines=[ln for ln in p.strip().splitlines() if ln.strip()]
        while lines and lines[0].strip()==".": lines.pop(0)
        if lines: sections.append(lines)
    return sections

def _parse_section(lines: List[str]):
    header = _decode(lines[0] if lines else "")
    gmap_url=None; eq_loc=None; eq_type=None
    locline = header
    m=_HDR.match(header)
    if m:
        city, site, _addr, lat, lon, typ = m.groups()
        locline  = f"{city}, {site}, {typ}".strip().strip(",")
        gmap_url = f"https://www.google.com/maps?q={lat},{lon}"
        eq_loc   = f"{city}, {site}".strip().strip(",")
        eq_type  = typ.strip()

    equip   = _decode(lines[1]) if len(lines)>1 else ""
    cables: List[str] = []
    splices: List[str] = []  # JSON splice lines => Existing Splice grouping
    lengths: Dict[str, Tuple[Optional[int], Optional[int]]] = {}
    pmid_sig=set()

    for raw in lines[2:]:
        t=_decode(raw); up=t.upper()
        if t.upper().startswith("CA"):
            cables.append(t); pmid_sig.update(_pmids(t))
        elif "-- SPLICE --" in up:
            splices.append(t); pmid_sig.update(_pmids(t))
        elif ("CABLE LENGTH" in up) or ("CALCULATED" in up) or ("OTDR" in up):
            ids=_pmids(t)
            if not ids: continue
            pmid=ids[0]; calc=None; otdr=None
            mC=_CALC.search(t); mO=_OTDR.search(t)
            if mC: calc=int(mC.group(1))
            if mO: otdr=int(mO.group(1))
            prev=lengths.get(pmid,(None,None))
            lengths[pmid]=(calc if calc is not None else prev[0],
                           otdr if otdr is not None else prev[1])
            pmid_sig.add(pmid)

    return locline, equip, cables, splices, lengths, pmid_sig, gmap_url, eq_loc, eq_type

def _group_splices(splices: List[str]) -> Optional[Tuple[str,int]]:
    if not splices: return None
    pairs={}
    for s in splices:
        pmids=_pmids(s)
        fnums=[int(x) for x in _FNUM.findall(s)]
        if len(pmids)>=2 and len(fnums)>=2:
            lp, rp = pmids[0], pmids[-1]
            L, R   = fnums[0], fnums[-1]
            d=pairs.setdefault((lp,rp), {"L":[], "R":[]})
            d["L"].append(L); d["R"].append(R)
    if not pairs: return None
    (lp,rp), LR = max(pairs.items(), key=lambda kv: len(kv[1]["L"]))
    Lc=_compact(LR["L"]); Rc=_compact(LR["R"]); cnt=len(set(LR["L"]))
    text=f"PMID: {lp}, Aptum ID: .F[{Lc}] -- Existing Splice(s) -- PMID: {rp}, Aptum ID: .F[{Rc}]"
    return text, cnt

# ---------- core build ----------
def build_fiber_trace(payload: dict) -> Tuple[pd.DataFrame, List[Optional[str]], List[Tuple[int,int]], List[dict], str, str]:
    rows: List[List[object]] = []
    gmaps: List[Optional[str]] = []
    ranges: List[Tuple[int,int]] = []
    meta  : List[dict] = []

    agg_map={}; order=[]
    for blob in _get_connection_blobs(payload):
        for lines in _split(blob):
            locline,equip,cables,splices,lengths,pmid_sig,gmap_url,eq_loc,eq_type=_parse_section(lines)
            key=(locline,equip)
            if key not in agg_map:
                agg_map[key]={"gmap":gmap_url,"eq_loc":eq_loc,"eq_type":eq_type,
                              "cables":set(),"splices":[],"lengths":{},"pmids":set()}
                order.append(key)
            A=agg_map[key]
            A["cables"].update(cables)
            A["splices"].extend(splices)
            A["pmids"].update(pmid_sig)
            for pmid,(calc,otdr) in lengths.items():
                prev=A["lengths"].get(pmid,(None,None))
                A["lengths"][pmid]=(calc if calc is not None else prev[0],
                                    otdr if otdr is not None else prev[1])

    # A/Z derived from first/last equipment-location in JSON order (fallback to header line)
    a_loc = (agg_map[order[0]]["eq_loc"] or order[0][0]) if order else ""
    z_loc = (agg_map[order[-1]]["eq_loc"] or order[-1][0]) if order else ""

    # Add A/Z header rows
    rows.append(["A Location", a_loc or "", "", "", "", "", "", "", "", ""]); gmaps.append(None)
    rows.append(["Z Location", z_loc or "", "", "", "", "", "", "", "", ""]); gmaps.append(None)

    for key in order:
        locline,equip = key
        A=agg_map[key]
        start=len(rows)
        rows.append(["Equipment Location", locline, "", "", "", "", "", A["eq_loc"] or "", A["eq_type"] or "", "Google Map" if A["gmap"] else ""])
        gmaps.append(A["gmap"])
        if equip:
            rows.append(["Equipment", equip, "", "", "", "", "", "", "", ""]); gmaps.append(None)
        for c in sorted(A["cables"]):
            rows.append(["Cable[attach]", c, "", "", "", "", "", "", "", ""]); gmaps.append(None)
        grp=_group_splices(A["splices"])
        if grp:
            text,cnt=grp
            rows.append(["Existing Splice", text, int(cnt), "", "", "", "", "", "", ""]); gmaps.append(None)
        for pmid in sorted(A["lengths"].keys()):
            calc,otdr=A["lengths"][pmid]
            line=f"PMID: {pmid}: " + (f"OTDR: {otdr}m, " if otdr is not None else "") + (f"Calculated: {calc}m" if calc is not None else "")
            rows.append(["CableInfo", line, "", "", calc if calc is not None else "", otdr if otdr is not None else "", "", "", "", ""]); gmaps.append(None)
        end=len(rows)-1
        ranges.append((start,end))
        meta.append({"start":start,"end":end,"locline":locline,"equipment":equip,"pmids":set(A["pmids"])})

    df = pd.DataFrame(rows, columns=COLUMNS).fillna("")
    for col in ("# of Fbrs","Length","OTDR Length"): df[col]=pd.to_numeric(df[col], errors="coerce")
    return df, gmaps, ranges, meta, a_loc, z_loc

# ---------- actions (from fibre_action) ----------
_ACTION_PMIDS = re.compile(r'\b([A-Za-z0-9#._/-]+)\s*\[[^\]]+\]')
def _pmids_loose_for_action(desc: str) -> List[str]:
    # Extract IDs like "41121 [73-84]" → "41121" (also keeps alnum tokens like 40084B)
    out=[]
    for m in _ACTION_PMIDS.finditer(str(desc) or ""):
        tok = m.group(1)
        out.append(tok.upper())
    # also consider explicit "PMID: X" if present
    out.extend(_pmids(desc))
    return list(dict.fromkeys(out))  # de-dup preserve order

def _pick_col(df: pd.DataFrame, candidates: List[str]) -> Optional[str]:
    for c in candidates:
        if c in df.columns: return c
    norm = {re.sub(r"\s+", "", c).lower(): c for c in df.columns}
    for want in candidates:
        key = re.sub(r"\s+", "", want).lower()
        if key in norm: return norm[key]
    return None

_RANGE = re.compile(r"\[([0-9]+)(?:\s*-\s*([0-9]+))?\]")

def _fibercount(s: str) -> Optional[int]:
    rngs=[(int(a), int(b) if b else int(a)) for a,b in _RANGE.findall(str(s) or "")]
    if not rngs: return None
    return max((b-a+1) for a,b in rngs)

def _label_from_actiondesc(action: str, desc: str) -> Optional[str]:
    a=(action or "").lower(); d=(desc or "").lower()
    if ("remove" in a) or ("break" in d): return "Break Splice"
    if ("add" in a) or ("splice" in d):   return "Splice Required"
    return None

def _classify_and_filter(acts: pd.DataFrame) -> pd.DataFrame:
    a_col=_pick_col(acts, ["Action"]); d_col=_pick_col(acts, ["Description"])
    if not a_col or not d_col: return pd.DataFrame(columns=["Action","Description","SAP"])
    out=acts[[a_col,d_col] + ([c for c in acts.columns if c.lower()=="sap"] if "SAP" in acts.columns else [])].copy()
    out.columns=["Action","Description"] + (["SAP"] if out.shape[1]==3 else [])
    # keep only Remove (E) and Add / Splice
    mask = out["Action"].str.contains(r"remove", case=False, na=False) | \
           out["Action"].str.contains(r"add", case=False, na=False) | \
           out["Description"].str.contains(r"(?:break|splice)", case=False, na=False)
    return out[mask].reset_index(drop=True)

def insert_actions(df: pd.DataFrame, gmaps: List[Optional[str]],
                   ranges: List[Tuple[int,int]], meta: List[dict],
                   actions_df: Optional[pd.DataFrame]) -> Tuple[pd.DataFrame, List[Optional[str]], Dict[str,int]]:
    if actions_df is None or actions_df.empty: 
        return df, gmaps, {"inserted":0,"adds":0,"removes":0}

    acts = _classify_and_filter(actions_df)
    if acts.empty:
        return df, gmaps, {"inserted":0,"adds":0,"removes":0}

    out=df.copy(); gm=gmaps[:]
    addmap: Dict[int, List[List[object]]] = {}
    stats={"inserted":0,"adds":0,"removes":0}

    for _, row in acts.iterrows():
        action=str(row.get("Action",""))
        desc  =str(row.get("Description",""))
        label = _label_from_actiondesc(action, desc)
        if not label: 
            continue

        pmidset=set(_pmids_loose_for_action(desc))
        if not pmidset:
            continue

        # find best equipment-block by PMID overlap
        best=None; best_ov=0
        for bi, m in enumerate(meta):
            ov=len(pmidset & m["pmids"])
            if ov>best_ov: best_ov=ov; best=bi
        if best is None or best_ov==0:
            continue

        fbrs_val = _fibercount(desc) or ""
        new_row=[label, desc,
                 fbrs_val,
                 action,  # put the full "WO Action#" string
                 "", "", "", "", "", ""]
        addmap.setdefault(best, []).append(new_row)

        stats["inserted"]+=1
        if "remove" in action.lower(): stats["removes"]+=1
        elif "add" in action.lower():  stats["adds"]+=1

    # insert from bottom to top so indices don’t shift
    for bi in sorted(addmap.keys(), reverse=True):
        start,end=ranges[bi]
        block=out.iloc[start:end+1]
        insert_at=start+1
        cab_rows=block[block["Detail Item"]=="Cable[attach]"]
        if not cab_rows.empty:
            insert_at=cab_rows.index[-1]+1
        else:
            eq_rows=block[block["Detail Item"]=="Equipment"]
            if not eq_rows.empty:
                insert_at=eq_rows.index[-1]+1

        add_df=pd.DataFrame(addmap[bi], columns=out.columns)
        out=pd.concat([out.iloc[:insert_at], add_df, out.iloc[insert_at:]], ignore_index=True)
        gm = gm[:insert_at] + [None]*len(add_df) + gm[insert_at:]

    return out, gm, stats

# ---------- writer (fills + Google Map hyperlinks) ----------
def write_xlsx_bytes(df: pd.DataFrame, gmaps: List[Optional[str]], sheet: str = "Fibre Trace") -> bytes:
    bio = io.BytesIO()
    with pd.ExcelWriter(bio, engine="xlsxwriter") as xw:
        df.to_excel(xw, index=False, sheet_name=sheet)
        wb=xw.book; ws=xw.sheets[sheet]

        # autosize
        df_str=df.astype(str)
        for ci,col in enumerate(df.columns):
            width=min(80, max(10, df_str[col].map(len).max()+2))
            ws.set_column(ci, ci, width)

        # first-column fills
        fmt_loc   = wb.add_format({"bg_color":"#CDEFD1","bold":True})  # gentle green
        fmt_equip = wb.add_format({"bg_color":"#CDEFD1"})
        fmt_cab   = wb.add_format({"bg_color":"#B8F1B0"})
        fmt_break = wb.add_format({"bg_color":"#FF0000"})   # Break: red
        fmt_req   = wb.add_format({"bg_color":"#FFA500"})   # Splice/Add: orange
        fmt_info  = wb.add_format({"bg_color":"#ADD8E6"})
        link_fmt  = wb.add_format({"font_color":"blue","underline":1})

        c_first = 0
        c_gmap  = df.columns.get_loc("Google Map")

        for r in range(len(df)):
            label=str(df.iat[r, c_first])
            if   label=="Equipment Location": fmt=fmt_loc
            elif label=="Equipment":          fmt=fmt_equip
            elif label=="Cable[attach]":      fmt=fmt_cab
            elif label=="Break Splice":       fmt=fmt_break
            elif label in ("Splice Required","Splice required"): fmt=fmt_req
            elif label=="CableInfo":          fmt=fmt_info
            elif label in ("A Location","Z Location"): fmt=fmt_loc
            else: fmt=None
            if fmt is not None:
                ws.write(r+1, c_first, label, fmt)

            url = gmaps[r] if r < len(gmaps) else None
            if url:
                ws.write_url(r+1, c_gmap, url, link_fmt, string="Google Map")
    return bio.getvalue()

# ---------- top-level helpers used by Streamlit ----------
def build_trace_and_actions_from_sources(*, json_source: Union[str, IO[bytes], bytes],
                                         csv_source: Optional[Union[str, IO[bytes], bytes]] = None):
    """Build the trace from JSON and (optionally) merge actions parsed by fibre_action against the trace.
       Returns: df, gmaps, stats, a_loc, z_loc
    """
    # Load JSON from path/bytes/file-like
    if isinstance(json_source, (str, Path)):
        payload = json.loads(Path(json_source).read_text(encoding="utf-8"))
    elif hasattr(json_source, "read"):
        payload = json.loads(json_source.read().decode("utf-8"))
        json_source.seek(0)
    elif isinstance(json_source, (bytes, bytearray)):
        payload = json.loads(json_source.decode("utf-8"))
    else:
        raise TypeError("Unsupported json_source type")

    base_df, gmaps, ranges, meta, a_loc, z_loc = build_fiber_trace(payload)

    if csv_source is not None:
        # import fibre_action locally
        import importlib.util, sys as _sys
        spec = importlib.util.spec_from_file_location("fiber_action", str(Path(__file__).with_name("fiber_action.py")))
        fa = importlib.util.module_from_spec(spec)
        _sys.modules["fiber_action"] = fa
        spec.loader.exec_module(fa)

        actions_df = fa.compute_fibre_actions(csv_source, json_source=None)
        out_df, out_gmaps, stats = insert_actions(base_df, gmaps, ranges, meta, actions_df)
    else:
        out_df, out_gmaps, stats = base_df, gmaps, {"inserted":0,"adds":0,"removes":0}

    return out_df, out_gmaps, stats, a_loc, z_loc

# --- Compatibility wrapper for legacy callers -----------------------------
def build_trace_df_and_stats(json_bytes: bytes, csv_bytes: bytes | None):
    """
    Legacy adapter that returns (df, stats) only.
    Internally uses build_trace_and_actions_from_sources (which also computes gmaps/AZ).
    """
    df, _gmaps, stats, _a_loc, _z_loc = build_trace_and_actions_from_sources(
        json_source=json_bytes,
        csv_source=csv_bytes
    )
    return df, stats
