import io, re, pandas as pd
from fiber_action import compute_fibre_actions

SAMPLE_CSV = """Action,Description,SAP
1: Add,Splice 41122 [73 - 84] 41121 [73 - 84] M#248 - M#248,
"""

SAMPLE_JSON = b'''
{"Report: Splice Details":[{"":[
  {"Connections":"PMID: 41121, Aptum: APT; .F79 -- Splice -- PMID: 41122, Aptum: APT; .F79"}
]}]}
'''

def test_pairs_follow_json_natural_order():
    df = compute_fibre_actions(
        csv_source=io.StringIO(SAMPLE_CSV),
        json_source=io.BytesIO(SAMPLE_JSON),
        enrich_from_json=True,
        sort_by_action_number=False
    )
    desc = df.loc[0, "Description"]
    assert re.search(r"^Splice\s+41121\s*\[\s*73\s*-\s*84\s*\].*41122\s*\[\s*73\s*-\s*84\s*\]", desc)
