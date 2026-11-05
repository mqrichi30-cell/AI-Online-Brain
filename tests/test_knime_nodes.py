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
