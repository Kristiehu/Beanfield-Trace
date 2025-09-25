# 📡 Fiber Trace Automation Toolkit

- A collection of Python scripts and helpers for processing fibre network data, generating reports, and exporting GIS-ready files (CSV, Excel, KML).  
- This project streamlines common workflows around work orders, splice reports, and capacity planning.

---

## 📂 Repository Structure
- `README.md` — Project documentation
- `trace_report.py` — Input Files Upload & Parse Logic
  
### Main Logic
- [x] `fiber_action.py` — Function 01 -  Fiber actions (Remove & Add)
- [x] `fibre_trace.py` — Function 02 -  **Core Fibre Trace Algorithms** *
- [x] `parse_device_sheet.py` — Function 03 - Activity Overview Map
- [x] `_to_kml.py` — Function 04 - KML Map Generation
- [ ] `activities.py` - Under development
- [ ] `cover_page.py` - Under development


### Helper
- `kml_helper.py` — Helper (KML Gneneration)
- `run_clean_csv.py`, `run_clean_json.py` - File Cleanup & Preview

### Input
- `WO[Number].csv` — Example work order CSV
- `WO[Number].json` — Example work order JSON


---

## 🛠️ Installation

Clone the repo and install dependencies:

```bash
git clone https://github.com/Kristiehu/Beanfield-Trace.git
cd trace
pip install -r requirements.txt
```

## ⚙️ App Setup
```bash
cd app
streamlit run app.py
```
Now the [Streamlit](https://github.com/streamlit/streamlit) App will be popup for you to explore 🚀!

### Additional Usage
1. Parse and Generate Reports
```bash
python trace_report.py --wo WO24218.json --csv WO24218.csv --out report.xlsx
```

3. Export to KML
```bash
python _to_kml.py --wo WO24218.json --csv WO24218.csv --out WO24218.kml
```

4. Run Remove & Add Algorithm
```bash
python remove_add_algo.py --in changes.csv --out output.csv
```

![myImage](https://media.giphy.com/media/XRB1uf2F9bGOA/giphy.gif)
