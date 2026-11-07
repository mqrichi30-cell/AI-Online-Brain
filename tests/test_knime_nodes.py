"""Corre el script del nodo 02 tal cual con un `knime.scripting.io` simulado."""
import json
import runpy
import sys
import types
from pathlib import Path

import pandas as pd

ROOT = Path(__file__).resolve().parents[1]


class _Table:
    def __init__(self, df):
        self.df = df

    def to_pandas(self):
        return self.df

    @classmethod
    def from_pandas(cls, df):
        return cls(df)


def fake_knio(inputs, flow_vars):
    io = types.ModuleType("knime.scripting.io")
    io.input_tables = [_Table(d) for d in inputs]
    io.output_tables = [None, None]
    io.flow_variables = dict(flow_vars)
    io.Table = _Table
    pkg, sub = types.ModuleType("knime"), types.ModuleType("knime.scripting")
    pkg.scripting, sub.io = sub, io
    sys.modules.update({"knime": pkg, "knime.scripting": sub, "knime.scripting.io": io})
    return io


def run_match(tmp_path, nova, sap):
    cfg = tmp_path / "config.json"
    cfg.write_text(json.dumps({"state_db": str(tmp_path / "s.sqlite"), "sap": {"po_column": "BSTKD"},
                               "matching": {"fire_once_per_day": True}}))
    io = fake_knio([pd.DataFrame({"po": nova}), pd.DataFrame({"BSTKD": sap})],
                   {"lib_dir": str(ROOT / "knime" / "lib"), "config_path": str(cfg)})
    runpy.run_path(str(ROOT / "knime" / "nodes" / "02_match_trigger.py"))
    return io


def test_node_fires_once_per_day(tmp_path):
    io = run_match(tmp_path, ["1", "1", "2"], ["0001", "2"])
    assert io.flow_variables["trigger_fired"] == 1 and io.flow_variables["PortChoice"] == "top"
    summary = io.output_tables[0].df.iloc[0]
    assert summary["nova_unique_pos"] == 2 and summary["nova_rows"] == 3
    io = run_match(tmp_path, ["1", "2"], ["1", "2"])
    assert io.flow_variables["trigger_fired"] == 0
    assert "Ya se disparó hoy" in io.flow_variables["trigger_reason"]


def test_node_not_fired(tmp_path):
    io = run_match(tmp_path, ["1", "2", "3"], ["1", "2"])
    assert io.flow_variables["PortChoice"] == "bottom"
    detail = io.output_tables[1].df
    assert detail.loc[detail.po == "3", "estatus"].item() == "SOLO_NOVA"


def test_node_autodetects_po_columns_and_uses_data_area(tmp_path):
    import shutil
    data = tmp_path / "data"
    shutil.copytree(ROOT / "knime" / "lib", data / "lib")
    (data / "config.json").write_text(json.dumps({"state_db": "state.sqlite", "sap": {},
                                                  "matching": {"fire_once_per_day": False}}))
    io = fake_knio([pd.DataFrame({"PO Number": ["1", "2"], "DC": ["a", "b"]}),
                    pd.DataFrame({"Material": ["x", "y"], "BSTKD": ["0002", "1"]})], {})
    io.get_workflow_data_area_dir = lambda: str(data)
    runpy.run_path(str(ROOT / "knime" / "nodes" / "02_match_trigger.py"))
    s = io.output_tables[0].df.iloc[0]
    assert (s["nova_po_column"], s["sap_po_column"]) == ("PO Number", "BSTKD")
    assert io.flow_variables["trigger_fired"] == 1
    assert (data / "state.sqlite").exists()


def test_build_workflow_package(tmp_path):
    import zipfile
    import xml.dom.minidom
    sys.path.insert(0, str(ROOT / "knime" / "workflows"))
    import build_workflow
    wf = build_workflow.build(tmp_path / "build")
    for f in wf.rglob("*.xml"):
        xml.dom.minidom.parse(str(f))
    xml.dom.minidom.parse(str(wf / "workflow.knime"))
    knwf = build_workflow.pack(wf, tmp_path / "x.knwf")
    names = zipfile.ZipFile(knwf).namelist()
    assert "NOVA_SAP_Trigger/workflow.knime" in names
    assert "NOVA_SAP_Trigger/data/lib/nova_sap/po_matching.py" in names
    assert "NOVA_SAP_Trigger/Python Script (#3)/settings.xml" in names
