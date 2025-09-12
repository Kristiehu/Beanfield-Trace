# tests/test_gather_connections.py
import json
from parse_device_sheet import gather_connections

SHAPE_A = {
    "Report: Splice Details": [
        {"": [
            {"Connections": "City,Site, Address: X,() : 43.1, -79.2 :  : OSP Splice Box - ABC\n."}
        ]}
    ]
}

SHAPE_B = {
    "Report: Splice Details": [
        {"": [
            "City,Site, Address: X,() : 43.1, -79.2 :  : OSP Splice Box - ABC\n.\n."
        ]}
    ]
}

def test_gather_connections_shape_a():
    out = gather_connections(SHAPE_A)
    assert out and "Address:" in out[0]

def test_gather_connections_shape_b():
    out = gather_connections(SHAPE_B)
    assert out and "Address:" in out[0]
