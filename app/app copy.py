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

# --------------------------- Streamlit UI ---------------------------
ROUTE_TYPES   = ["Primary", "Diverse", "Triverse"]
CIRCUIT_TYPES = ["NEW", "Existing"]
st.set_page_config(page_title="Trace Builder", layout="wide")
from helper import get_work_order_number

work_order_title = get_work_order_number(up_csv) if 'up_csv' in globals() else "Work Order"
st.title(f"Trace Builder - [{work_order_title}]")
st.caption("Upload WO.csv + WO.json, choose options, then export the Excel report (Cover Page, Fibre Trace, Activity Overview Map, Error Report) and a colorized KML/KMZ.")

# --------------------- sidebar inputs ---------------------------------------
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
    
# --------------------------------- UPLOADS ----------------------------------
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

from cleanup_json import run_clean_json, export_bytes as export_json_bytes, build_print_table
from cleanup_csv import run_clean_csv, export_csv_bytes
# ---------------- CSV CLEANUP ----------------
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

# ---------------- JSON CLEANUP ----------------
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


# --------------------------------- Generate ----------------------------------

st.subheader("2) Generate outputs")

# --------------------- Fibre Action Button ------------------------------
from fiber_action import (
    read_uploaded_table,
    transform_fibre_action_summary_grid,
    transform_fibre_action_actions,
    fibre_action_excel_bytes,
    simplify_description,
)
from helper import extract_meta_from_label_value_df
with st.container(border=True):
    st.markdown("Fibre Action")

    btn_disabled = (up_csv is None)
    if st.button("Generate", type="primary", key="btn_fibre_action", disabled=btn_disabled):
        try:
            with st.spinner("Building Fibre Action…"):
                # 1) Read WO CSV
                wo_df = read_uploaded_table(up_csv)

                # 2) Build Summary grid (L1,V1,L2,V2)
                # Pull values from the CSV (label:value layout)
                pref = extract_meta_from_label_value_df(wo_df)

                # (optional) keep sidebar/session in sync so fields flip from grey→white
                st.session_state["wo_number"]   = pref.get("wo_id") or st.session_state.get("wo_number", "")
                st.session_state["designer_name"]= pref.get("designer_name") or st.session_state.get("designer_name", "")
                st.session_state["order_number"] = pref.get("order_id") or st.session_state.get("order_number", "")
                st.session_state["date"]         = pref.get("date_ddmmyyyy") or st.session_state.get("date", "")

                # Build the meta used by Summary
                try:
                    pref = extract_meta_from_label_value_df(wo_df)  # returns dict with keys like below in most setups
                except Exception:
                    pref = {}

                meta = {
                    "order_id":       st.session_state["order_number"],               # from CSV ORDER Number
                    "wo_id":          st.session_state["wo_number"],                  # from CSV Work Order
                    "designer_name":  st.session_state["designer_name"],              # from CSV Created By
                    "designer_phone": st.session_state.get("designer_phone", ""),
                    "date":           st.session_state["date"],                       # from CSV Created On (dd/mm/yyyy)
                    "a_end":          st.session_state.get("a_end", ""),
                    "z_end":          st.session_state.get("z_end", ""),
                    "details":        "OSP DF",
                }

                summary_df = transform_fibre_action_summary_grid(wo_df, meta)

                # 3) Fibre Action list
                actions_df = transform_fibre_action_actions(wo_df)
                if "Description" in actions_df.columns:
                    actions_df["Description"] = actions_df["Description"].astype(str).map(simplify_description)
                else:
                    actions_df.insert(1, "Description", "—")
                if "SAP" not in actions_df.columns:
                    actions_df["SAP"] = ""
                actions_df = actions_df[["Action", "Description", "SAP"]].fillna("")

                # 4) Preview
                t1, t2 = st.tabs(["Summary", "Fibre Action"])
                with t1:
                    st.dataframe(summary_df, use_container_width=True, hide_index=True)
                with t2:
                    st.dataframe(actions_df, use_container_width=True, hide_index=True)

                # 5) Excel + download
                xbytes = fibre_action_excel_bytes(summary_df, actions_df, title="fibre_action")
            st.success(f"Fibre Action generated: {len(actions_df)} actions.")
            st.download_button(
                "Download fibre_action.xlsx",
                data=xbytes,
                file_name="fibre_action.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                key="dl_fibre_action_xlsx",
            )

        except Exception as e:
            st.exception(e)
            st.error(f"Failed to generate Fibre Action: {e}")
# ------------------------------------------------------------------------

# --------------------- Fibre Trace Button -------------------------------
from fiber_trace_v0 import generate_xlsx 
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
from parse_device_sheet import main as parse_device_main  # for activity overview parsing
with st.container(border=True):
    st.markdown("Activity Overview Map")
    # The button that triggers the generation
    if st.button("Generate", type="primary", key="btn_activity_overview", disabled=not ready):
        # Make sure inputs are present and readable
        wo_df, payload = require_inputs(up_csv, up_json)

        # Try to use the exact logic from parse_device_sheet.py
        try:
            from parse_device_sheet import gather_connections, parse_device_table  # exact logic

        except Exception:
            gather_connections = None
            parse_device_table = None

        # Fallback gatherer if functions weren't exported in the module
        def _fallback_gather_connections(obj):
            out = []
            def rec(x):
                if isinstance(x, dict):
                    for k, v in x.items():
                        if k == "Connections" and isinstance(v, str):
                            out.append(v)
                        else:
                            rec(v)
                elif isinstance(x, list):
                    for it in x:
                        rec(it)
            rec(obj)
            return out

        # Decide which gatherer we’ll use
        _gather = gather_connections or _fallback_gather_connections

        try:
            conns = _gather(payload)
            if not conns:
                st.error("No 'Connections' strings found in the JSON under 'Report: Splice Details'.")
                st.stop()

            # Use the first unique string (how the script behaves)
            conn_text = conns[0]

            # If parse_device_table is available, use it; otherwise fail loudly (you want exact logic)
            if parse_device_table is None:
                raise RuntimeError("`parse_device_table` not found in parse_device_sheet.py. Please export it.")

            # Build the table (append_details=True keeps richer rows if your script supports it)
            df = parse_device_table(conn_text, append_details=True)

            if df is None or df.empty:
                st.warning("Activity Overview Map produced an empty table.")
                st.stop()

            # Prepare CSV bytes and a preview
            csv_bytes = df.to_csv(index=False, encoding="utf-8-sig").encode("utf-8-sig")

            st.success(f"Generated Activity Overview table with {len(df)} rows.")
            st.download_button(
                "Download Activity Overview CSV",
                data=csv_bytes,
                file_name="activity_overview.csv",
                mime="text/csv",
                key="dl_activity_overview_csv",
            )

            # Preview (responsive, no index)
            st.dataframe(df, use_container_width=True, hide_index=True)

        except Exception as e:
            st.error(f"Failed to generate Activity Overview Map: {e}")    
    # -----------------------------------------------------------------
# ------------------------------------------------------------------------  

# --------------------- Generate KML Button ------------------------------
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