# app.py
import re, json, io, math, os, tempfile
from io import BytesIO
import pandas as pd
from datetime import date, datetime
import streamlit as st
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill, Border, Side
from openpyxl.utils import get_column_letter
from pathlib import Path
from trace_report import _kv_from_wo, _actions_from_wo, _details_from_json # 'upload' helpers 
from helper import (
    require_inputs,
    placemarks_from_wo_df,
    placemarks_from_payload,
    dedupe_placemarks,
    read_uploaded_table,
)

# --------------------------- Helpers ---------------------------
# parses lines like: "... Address: ... () : 43.644719, -79.385046 :  : something"
_COORD_RE = re.compile(r":\s*(-?\d+(?:\.\d+)?)\s*,\s*(-?\d+(?:\.\d+)?)\s*:\s*:\s*(.*)$")

# =======================================================================
# --------------------------- Streamlit UI ---------------------------
ROUTE_TYPES   = ["Primary", "Diverse", "Triverse"]
CIRCUIT_TYPES = ["NEW", "Existing"]
st.set_page_config(page_title="Trace Builder", layout="wide")
from helper import get_work_order_number

work_order_title = get_work_order_number(up_csv) if 'up_csv' in globals() else "Work Order"
st.title(f"Trace Builder - [{work_order_title}]")
st.caption("Upload WO.csv + WO.json, choose options, then export the Excel report (Cover Page, Fibre Trace, Activity Overview Map, Error Report) and a colorized KML/KMZ.")

# =======================================================================
# ---------------------------- sidebar inputs -------------------------
st.sidebar.header("Project metadata")
st.sidebar.markdown("Fill in the project metadata to generate a complete workbook.")
with st.sidebar:
    st.header("Project metadata")
    work_order    = st.text_input("Work Order #", "WO24218")
    designer_name  = st.text_input("Designer Name", "abcd")
    designer_email = st.text_input("Designer Email", "abcd@beanfield.com")
    designer_phone = st.text_input("Designer Phone #", "4164164164")
    fibers_assigned = st.number_input("# of Fibers Assigned", min_value=1, step=1, value=2)

    order_number   = st.text_input("Order Number / Circuit ID", "ORDER-267175")
    circuit_id     = st.text_input("Circuit ID", "CK37105")
    client_name    = st.text_input("Client Name", "ABC In")
    service_type   = st.text_input("Service Type", "IRU Fbr")
    device_type    = st.text_input("Device Type", "N/A")

    route_type     = st.selectbox("Route Type", ROUTE_TYPES, index=0)
    circuit_type   = st.selectbox("Select Circuit Type", CIRCUIT_TYPES, index=0)

    a_end          = st.text_input("A End (auto)", "BFMH-0021")
    z_end          = st.text_input("Z End (auto)", "SITE5761")
    circuit_version= st.text_input("Circuit Version", "1")

    template_xlsm  = st.file_uploader("Macro template (.xlsm, optional)", type=["xlsm"])
    want_xlsm      = st.checkbox("Export as .xlsm (requires template)", value=bool(template_xlsm))

# =======================================================================   
# ---------------------------- UPLOADS ---------------------------------
st.subheader("0) Before Upload Files...")
st.markdown("Create your circuit as per normal, and save the Fibre Trace as a JSON File")
st.subheader("1) Upload inputs")
col_csv, col_json = st.columns(2, gap="large")

with col_csv:
    with st.container(border=True):
        st.caption("Upload Work Order CSV file")
        up_csv = st.file_uploader(" ", type=["csv"], key="wo_csv")
        run_csv = st.button("Clean up CSV", type="primary", use_container_width=True, key="btn_cleanup_csv")

with col_json:
    with st.container(border=True):
        st.caption("Upload Work Order JSON file")
        up_json = st.file_uploader("  ", type=["json"], key="wo_json")
        run_json = st.button("Clean up JSON (trace)", type="primary", use_container_width=True, key="btn_cleanup_json")

ready = (up_csv is not None) and (up_json is not None)
if not ready:
    st.info("Upload both files to enable generation.")
else:
    wo_df = read_uploaded_table(up_csv)
    # use functions from this script:
    wo_kv = _kv_from_wo(wo_df)
    actions_df = _actions_from_wo(wo_df)
    details_df = _details_from_json(up_json)
    # Save latest into session for reuse by all panels
    if up_csv:  st.session_state["wo_csv_file"]  = up_csv
    if up_json: st.session_state["wo_json_file"] = up_json

    def _get_csv_bytes() -> bytes | None:
        f = st.session_state.get("wo_csv_file")
        if not f: return None
        try: f.seek(0)
        except Exception: pass
        return f.read()

    def _get_json_bytes() -> bytes | None:
        f = st.session_state.get("wo_json_file")
        if not f: return None
        try: f.seek(0)
        except Exception: pass
        return f.read()

    # Helper to require inputs (so each panel can call it)
    def require_inputs(need_csv: bool, need_json: bool) -> tuple[bytes | None, bytes | None, bool]:
        csv_b  = _get_csv_bytes()  if need_csv  else None
        json_b = _get_json_bytes() if need_json else None
        ok = True
        if need_csv and not csv_b:
            st.warning("Please upload the Work Order CSV above.")
            ok = False
        if need_json and not json_b:
            st.warning("Please upload the Work Order JSON above.")
            ok = False
        return csv_b, json_b, ok

    # Utility: make a temp file from uploaded bytes and yield the path
    from contextlib import contextmanager
    @contextmanager
    def as_tempfile(data: bytes, suffix: str):
        tf = tempfile.NamedTemporaryFile(delete=False, suffix=suffix)
        try:
            tf.write(data)
            tf.flush(); tf.close()
            yield Path(tf.name)
        finally:
            try: os.unlink(tf.name)
            except Exception: pass

# =======================================================================
# ------------------------------ Cleanup --------------------------------
from cleanup_json import run_clean_json, export_bytes as export_json_bytes, build_print_table
from cleanup_csv import run_clean_csv, export_csv_bytes

# --- CSV ---
if run_csv:
    if not up_csv:
        st.error("Please upload a CSV first.")
    else:
        try:
            res_csv = run_clean_csv(up_csv.getvalue())
            st.success("CSV cleanup completed.")
            c1, c2, c3 = st.columns(3)
            c1.metric("Rows", res_csv.integrity["rows"])
            c2.metric("Columns", res_csv.integrity["cols"])
            c3.metric("Duplicates removed", res_csv.integrity["duplicates_removed"])

            with st.expander("Preview • Cleaned CSV (top 200)", expanded=True):
                st.dataframe(res_csv.cleaned_df.head(200), use_container_width=True)

            arts = export_csv_bytes(res_csv, basename=(up_csv.name.rsplit(".",1)[0] or "work_order"))
            
        except Exception as e:
            st.exception(e)

# --- JSON ---
if run_json:
    if not up_json:
        st.error("Please upload a JSON first.")
    else:
        try:
            res = run_clean_json(up_json.getvalue())
            st.success("JSON cleanup completed.")
            c1, c2, c3 = st.columns(3)
            c1.metric("Nodes", res.integrity["counts"]["nodes"])
            c2.metric("Cables", res.integrity["counts"]["cables"])
            c3.metric("Splice Events", res.integrity["counts"]["events"])

            with st.expander("A→Z Trace (printable preview)", expanded=True):
                printable_df = build_print_table(res)
                st.dataframe(printable_df, use_container_width=True, hide_index=True)

            tabs = st.tabs(["Nodes", "Cables", "Events", "Clean JSON"])
            with tabs[0]:
                st.dataframe(res.nodes_df.head(200), use_container_width=True)
            with tabs[1]:
                st.dataframe(res.cables_df.head(200), use_container_width=True)
            with tabs[2]:
                st.dataframe(res.events_df.head(200), use_container_width=True)
            with tabs[3]:
                st.code(json.dumps(res.clean_json, ensure_ascii=False, indent=2), language="json")

            arts = export_json_bytes(res)
        except Exception as e:
            st.exception(e)

# =======================================================================
# ------------------------------ Generate ------------------------------
st.subheader("2) Generate outputs")

# --------------------- Fibre Action Button -----------------------------
# Generate an Excel workbook with a single "Fibre Action" sheet.
# Requires: the `compute_fibre_actions` function from fiber_action.py.
# Overview:
# 1. Read the uploaded CSV and JSON files.
# 2. Use `compute_fibre_actions` to create a DataFrame representing the fibre actions.
# 3. Export the DataFrame to an Excel file and offer it for download.
from fiber_action import compute_fibre_actions, actions_to_workbook_bytes

def _style_add_remove(row: pd.Series) -> list[str]:
    a = str(row.get("Action", "")).lower()
    if "add" in a:
        return ["background-color: #009688; color: #FFFFFF"] * len(row)  # blue-green
    if "remove" in a:
        return ["background-color: #8e24aa; color: #FFFFFF"] * len(row)  # purple
    return [""] * len(row)

with st.container(border=True):
    st.markdown("### Fibre Action")

    if st.button("Generate", type="primary", key="btn_fa", disabled=not ready):
        csv_b, json_b, ok = require_inputs(need_csv=True, need_json=True)
        if ok:
            try:
                df = compute_fibre_actions(csv_b, json_b)
                if df.empty:
                    st.warning("No splice/break actions detected.")
                else:
                    st.success(f"Built {len(df)} Fibre Action rows (CA-ordered, normalized).")
                    st.session_state["fa_df"] = df

                    # === Preview with row colors (Add = dark green, Remove = red) ===
                    preview = df[["Action", "Description"]].copy()
                    styled = preview.style.apply(_style_add_remove, axis=1)
                    st.dataframe(styled, use_container_width=True, hide_index=True)

                    # Download: keeps your Excel coloring from fiber_action.py
                    xlsx_bytes = actions_to_workbook_bytes(df)
                    st.download_button(
                        label="Download Fibre Action.xlsx",
                        data=xlsx_bytes,
                        file_name="Fibre_Action.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        key="dl_fa",
                    )
            except Exception as e:
                st.error(f"Failed to build Fibre Action: {e}")
# ------------------------------------------------------------------------

# ------------------------------------------------------------------------


# --------------------- Fibre Trace Button -------------------------------
# Generate an Excel workbook with Cover Page + Fibre Trace sheet.
# Requires: the `build_fiber_trace` function from fiber_trace.py.
# Overview:
# 1. Read the uploaded JSON file.
# 2. Use `build_fiber_trace` to create a DataFrame representing the fibre trace.
#3. Export the DataFrame to an Excel file and offer it for download.
from fiber_trace import generate_xlsx 

with st.container(border=True):
    st.markdown("Fibre Trace")

    if st.button("Generate", type="primary", key="btn_fibre_trace", disabled=not ready):
        try:
            # Save uploaded JSON to a temp file
            up_json.seek(0)
            with tempfile.NamedTemporaryFile(suffix=".json", delete=False) as tmp:
                tmp.write(up_json.read())
                json_path = Path(tmp.name)

            # Output path (temp)
            out_xlsx = json_path.with_name("fibre_trace.xlsx")

            # Generate Excel only (no CSV/KML)
            df_trace = generate_xlsx(json_path, out_xlsx)

            # Download Excel
            with open(out_xlsx, "rb") as fh:
                st.download_button(
                    "Download Fibre Trace (Excel)",
                    fh,
                    file_name="fibre_trace.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    key="dl_fibre_trace_xlsx",
                )

            # Preview (optional)
            st.success(f"Fibre Trace generated: {len(df_trace)} rows.")
            st.dataframe(df_trace, use_container_width=True, hide_index=True)

        except Exception as e:
            st.error(f"Failed to generate Fibre Trace: {e}")
#-------------------------------------------------------------------------

# --------------------- Activity Overview Map Button ---------------------
# Parse the 'Connections' strings from the JSON, then parse each line into a table.
# Requires: the `parse_device_table` function from parse_device_sheet.py.
# Overview:
# 1. Extract 'Connections' strings from the JSON robustly.
# 2. For each 'Connections' string, parse it into a DataFrame using `parse_device_table`.
# 3. Combine all DataFrames, deduplicate, and offer CSV download and preview.

from parse_device_sheet import parse_device_table  

with st.container(border=True):
    st.markdown("Activity Overview Map")

    if st.button("Generate", type="primary", key="btn_activity_overview", disabled=not ready):
        try:
            # 1) Inputs
            _out = require_inputs(up_csv, up_json)
            meta = {}
            if isinstance(_out, tuple):
                if len(_out) < 2:
                    raise ValueError("require_inputs returned fewer than 2 items")
                wo_df, payload = _out[0], _out[1]
                if len(_out) > 2:
                    meta = _out[2] or {}
            elif isinstance(_out, dict):
                wo_df   = _out["wo_df"]
                payload = _out["payload"]
                meta    = _out.get("meta", {}) or {}
            else:
                raise TypeError(f"Unexpected return type from require_inputs: {type(_out).__name__}")

            # 2) Normalize payload (dict), in case we got a JSON string/bytes
            if isinstance(payload, (bytes, bytearray)):
                try:
                    payload = payload.decode("utf-8", "replace")
                except Exception:
                    pass
            if isinstance(payload, str):
                try:
                    payload = json.loads(payload)
                except Exception:
                    # leave as string if not valid JSON
                    pass

            # 3) Helpers to gather 'Connections' robustly
            _COORD_RE = re.compile(r"\b-?\d{1,3}\.\d+\s*,\s*-?\d{1,3}\.\d+\b")

            def _looks_like_connection_blob(s: str) -> bool:
                if not s:
                    return False
                s = s.strip()
                return (
                    "Address:" in s
                    or "Splice" in s
                    or "FOSC" in s
                    or "PMID" in s
                    or _COORD_RE.search(s) is not None
                    or (len(s) >= 40 and s.count(":") >= 2)
                )

            def _dedupe(seq):
                seen = set()
                out = []
                for x in seq:
                    t = (x or "").strip()
                    if t and t not in seen:
                        seen.add(t); out.append(t)
                return out

            def _gather_connections_any(obj):
                """Global scan: collect any dict key named 'Connections' anywhere."""
                out = []
                def rec(x):
                    if isinstance(x, dict):
                        for k, v in x.items():
                            if isinstance(k, str) and k.strip().lower() == "connections":
                                if isinstance(v, str) and v.strip():
                                    out.append(v)
                                elif isinstance(v, list):
                                    for it in v:
                                        if isinstance(it, str) and it.strip():
                                            out.append(it)
                                        else:
                                            rec(it)
                                else:
                                    rec(v)
                            else:
                                rec(v)
                    elif isinstance(x, list):
                        for it in x:
                            rec(it)
                rec(obj)
                return _dedupe(out)

            def _gather_connections_loose(obj):
                """Very loose: collect long-ish strings that look like device/connection blobs."""
                out = []
                def rec(x):
                    if isinstance(x, dict):
                        for _, v in x.items():
                            rec(v)
                    elif isinstance(x, list):
                        for it in x:
                            rec(it)
                    elif isinstance(x, str):
                        if _looks_like_connection_blob(x):
                            out.append(x)
                rec(obj)
                return _dedupe(out)

            # 4) Try fast path first (the exact shape you showed), then global, then loose
            conns = []
            try:
                if isinstance(payload, dict):
                    r = payload.get("Report: Splice Details")
                    if isinstance(r, list) and r:
                        inner = r[0]
                        if isinstance(inner, dict) and "" in inner:
                            items = inner[""]
                            if isinstance(items, list):
                                conns = [d["Connections"] for d in items
                                         if isinstance(d, dict) and isinstance(d.get("Connections"), str)]
                                conns = _dedupe(conns)
            except Exception:
                pass

            if not conns and isinstance(payload, (dict, list)):
                conns = _gather_connections_any(payload)

            if not conns and isinstance(payload, (dict, list)):
                conns = _gather_connections_loose(payload)

            if not conns:
                st.error("No 'Connections' strings found in the JSON under 'Report: Splice Details'.")
                with st.expander("Debug details"):
                    st.write("payload type:", type(payload).__name__)
                    if isinstance(payload, dict):
                        st.write("Top-level keys:", list(payload.keys())[:30])
                        st.json(payload.get("Report: Splice Details", {}))
                st.stop()

            # # one tab per connection (# of conns found preview in tabs) -- IGNORE: don't need this function for now --
            # # One or more 'Connections' blobs found. Each blob may contain multiple lines/devices.   
            # # Optional: heads-up on what we found
            # st.info(f"Found {len(conns)} 'Connections' blobs in the JSON File.")

            # frames = []   # list of tuples: (idx, df_or_none, err_or_none)
            # for idx, blob in enumerate(conns):
            #     try:
            #         cand = parse_device_table(blob.strip(), append_details=True)
            #         if cand is not None and not cand.empty:
            #             cand = cand.copy()
            #             cand.insert(0, "connections_blob_index", idx)  # keep provenance
            #             frames.append((idx, cand, None))
            #         else:
            #             frames.append((idx, None, "Empty table"))
            #     except Exception as e:
            #         frames.append((idx, None, str(e)))

            # # Build tab titles with row counts
            # titles = [
            #     f"Conn {idx+1} ({0 if df is None else len(df)})"
            #     for idx, df, err in frames
            # ]
            # tabs = st.tabs(titles)

            # # Render one tab per connection
            # for (tab, (idx, df_i, err)) in zip(tabs, frames):
            #     with tab:
            #         st.caption(f"Blob #{idx+1}")
            #         if err:
            #             st.error(f"Parse error: {err}")
            #         elif df_i is None or df_i.empty:
            #             st.warning("No rows in this blob.")
            #         else:
            #             # Per-tab download
            #             csv_bytes_i = df_i.to_csv(index=False, encoding="utf-8-sig").encode("utf-8-sig")
            #             st.download_button(
            #                 f"Download CSV for Conn {idx+1}",
            #                 data=csv_bytes_i,
            #                 file_name=f"activity_overview_conn_{idx+1}.csv",
            #                 mime="text/csv",
            #                 key=f"dl_activity_overview_conn_{idx+1}",
            #             )
            #             # Hide the provenance column in the preview (still in the CSV)
            #             show_df = df_i.drop(columns=["connections_blob_index"], errors="ignore")
            #             st.dataframe(show_df, use_container_width=True, hide_index=True)

            # # Combined view/export
            # valid_dfs = [df for _, df, err in frames if df is not None and not df.empty]
            # if valid_dfs:
            #     combined = pd.concat(valid_dfs, ignore_index=True)
            #     before = len(combined)
            #     combined = combined.drop_duplicates()
            #     deduped = before - len(combined)

            #     st.divider()
            #     st.subheader("All connections (combined)")
            #     if deduped > 0:
            #         st.caption(f"Removed {deduped} duplicate rows when combining.")

            #     csv_all = combined.to_csv(index=False, encoding="utf-8-sig").encode("utf-8-sig")
            #     st.download_button(
            #         "Download ALL connections CSV",
            #         data=csv_all,
            #         file_name="activity_overview_all.csv",
            #         mime="text/csv",
            #         key="dl_activity_overview_all",
            #     )
            #     st.dataframe(combined, use_container_width=True, hide_index=True)
            # else:
            #     st.warning("No rows were produced from any 'Connections' blob.")
            

            # 5) Parse table
            if parse_device_table is None:
                raise RuntimeError("`parse_device_table` not found in parse_device_sheet.py. Please export it.")

            df = None
            errors = []
            for idx, blob in enumerate(conns[:5]):  # try first few candidates in case some are malformed
                try:
                    cand = parse_device_table(blob.strip(), append_details=True)
                    if cand is not None and not cand.empty:
                        df = cand
                        break
                except Exception as e:
                    errors.append(f"blob#{idx}: {e!s}")

            if df is None or df.empty:
                st.warning("Activity Overview Map produced an empty table from all candidate 'Connections' blobs.")
                if errors:
                    with st.expander("Parser error details"):
                        for msg in errors:
                            st.write(msg)
                st.stop()

            # 6) Offer CSV + preview
            csv_bytes = df.to_csv(index=False, encoding="utf-8-sig").encode("utf-8-sig")
            st.success(f"Generated Activity Overview table with {len(df)} rows.")
            st.download_button(
                "Download Activity Overview CSV",
                data=csv_bytes,
                file_name="activity_overview.csv",
                mime="text/csv",
                key="dl_activity_overview_csv",
            )
            st.dataframe(df, use_container_width=True, hide_index=True)

        except Exception as e:
            st.error(f"Failed to generate Activity Overview Map: {e}")
# ------------------------------------------------------------------------  

# --------------------- Generate KML Button ------------------------------
# Generate a colorized KML from the WO.csv and the JSON.
# Requires: the `to_kml` function from _to_kml.py.
# Overview:
# 1. Read the uploaded CSV and JSON files.
# 2. Extract placemarks from both sources.
# 3. Deduplicate placemarks.
# 4. Generate a KML string and offer it for download.

from _to_kml import to_kml 
with st.container(border=True):
    st.markdown ("KML")
    if st.button("Generate", type="primary", key="btn_generate_kml", disabled=not ready):
        if not up_csv or not up_json:
            st.error("Please provide both WO.csv and the circuit JSON.")
            st.stop()

        wo_df = pd.read_csv(up_csv)
        payload = json.load(up_json)

        pms = placemarks_from_wo_df(wo_df) + placemarks_from_payload(payload)
        pms = dedupe_placemarks(pms)

        kml_str = to_kml(title=order_number or "Trace", placemarks=pms)  # ✅ correct signature
        kml_bytes = kml_str.encode("utf-8")

        st.success(f"KML created with {len(pms)} placemarks.")
        st.download_button(
            "Download KML",
            data=kml_bytes,
            file_name=f"{order_number}_Trace.kml",
            mime="application/vnd.google-earth.kml+xml",
            key="dl_kml",
        )
# ------------------------------------------------------------------------        