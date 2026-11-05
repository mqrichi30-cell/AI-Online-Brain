"""Genera el workflow de KNIME listo para importar (.knwf).

    python knime/workflows/build_workflow.py            -> dist/NOVA_SAP_Trigger.knwf

El workflow es autocontenido: la librería nova_sap, el config.json y dos Excel de
ejemplo viajan dentro de la carpeta data/ del workflow, así que se importa y corre.
"""
from __future__ import annotations

import json
import shutil
import sys
import zipfile
from pathlib import Path
from xml.sax.saxutils import escape

ROOT = Path(__file__).resolve().parents[2]
WF_NAME = "NOVA_SAP_Trigger"
KNIME_VERSION = "5.2.0"  # versión mínima declarada; probado en 5.12
HEADER = ('<?xml version="1.0" encoding="UTF-8"?>\n<config xmlns="http://www.knime.org/2008/09/XMLConfig" '
          'xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xsi:schemaLocation="http://www.knime.org/2008/09/'
          'XMLConfig http://www.knime.org/XMLConfig_2008_09.xsd" key="{key}">\n')


# --------------------------------------------------------------------------- XML mini-builder
def e(key, value, typ="xstring"):
    if value is None:
        return f'<entry key="{escape(key)}" type="{typ}" isnull="true" value=""/>'
    if isinstance(value, bool):
        typ, value = "xboolean", str(value).lower()
    elif isinstance(value, int) and typ == "xstring":
        typ = "xint"
    return f'<entry key="{escape(key)}" type="{typ}" value="{escape(str(value), {chr(34): "&quot;", chr(10): "%%00010", chr(13): "%%00013"})}"/>'


def c(key, *children):
    if not children:
        return f'<config key="{escape(key)}"/>'
    return f'<config key="{escape(key)}">' + "".join(children) + "</config>"


def doc(key, *children) -> str:
    return HEADER.format(key=key) + "".join(children) + "\n</config>\n"


def annotation(text):
    return c("nodeAnnotation", e("text", text), e("contentType", "text/plain"), e("bgcolor", 16777215),
             e("x-coordinate", 0), e("y-coordinate", 0), e("width", 0), e("height", 0),
             e("alignment", "CENTER"), e("borderSize", 0), e("borderColor", 16777215), e("defFontSize", -1),
             e("annotation-version", 20151123), c("styles"))


def node_settings(*, name, factory, bundle_name, bundle_sym, feature_name, feature_sym, model, label,
                  creation_config="", out_ports=1, variables=""):
    ports = "".join(c(f"port_{i}", e("index", i), e("port_dir_location", None)) for i in range(1, out_ports + 1))
    return doc(
        "settings.xml",
        e("node_file", "settings.xml"), c("flow_stack"),
        c("internal_node_subsettings", e("memory_policy", "CacheSmallInMemory")),
        model, variables, annotation(label), e("customDescription", None), e("state", "IDLE"),
        e("factory", factory), e("node-name", name), e("node-bundle-name", bundle_name),
        e("node-bundle-symbolic-name", bundle_sym), e("node-bundle-vendor", "KNIME AG, Zurich, Switzerland"),
        e("node-bundle-version", KNIME_VERSION), e("node-feature-name", feature_name),
        e("node-feature-symbolic-name", feature_sym), e("node-feature-vendor", "KNIME AG, Zurich, Switzerland"),
        e("node-feature-version", KNIME_VERSION), c("factory_settings"), creation_config,
        e("name", name), e("hasContent", False), e("isInactive", False), c("ports", ports),
        c("filestores", e("file_store_location", None), e("file_store_id", None)),
    )


def table_ports(n):
    return "".join(c(f"port_{i}", e("object_class", "org.knime.core.node.BufferedDataTable")) for i in range(n))


# --------------------------------------------------------------------------- nodos
def excel_reader(file_in_data_area: str, label: str) -> str:
    chooser = c(
        "file_selection",
        c("file_system_chooser__Internals", e("has_fs_port", False), e("overwritten_by_variable", False),
          e("convenience_fs_category", "RELATIVE"), e("relative_to", "knime.workflow.data"),
          e("mountpoint", "LOCAL"), e("spaceId", ""), e("spaceName", ""), e("custom_url_timeout", 1000),
          e("connected_fs", True)),
        c("path", e("location_present", True), e("file_system_type", "RELATIVE"),
          e("file_system_specifier", "knime.workflow.data"), e("path", file_in_data_area)),
        c("filter_mode_Internals", e("SettingsModelID", "SMID_FilterMode"), e("EnabledStatus", True)),
        c("filter_mode", e("filter_mode", "FILE"), e("include_subfolders", False),
          c("filter_options", e("filter_files_extension", False), e("files_extension_expression", ""),
            e("files_extension_case_sensitive", False), e("filter_files_name", False),
            e("files_name_expression", "*"), e("files_name_case_sensitive", False),
            e("files_name_filter_type", "WILDCARD"), e("include_hidden_files", False),
            e("include_special_files", True), e("filter_folders_name", False),
            e("folders_name_expression", "*"), e("folders_name_case_sensitive", False),
            e("folders_name_filter_type", "WILDCARD"), e("include_hidden_folders", False),
            e("follow_links", True))),
    )
    model = c(
        "model",
        c("settings", c("file_selection_Internals", e("SettingsModelID", "SMID_ReaderFileChooser"),
                        e("EnabledStatus", True)), chooser,
          e("sheet_selection", "FIRST"), e("sheet_name", "default"), e("sheet_index", 0),
          e("table_contains_column_names", True), e("column_names_row_number", 1, "xlong"),
          e("row_id", "GENERATE"), e("row_id_column", "A"), e("sheet_area", "ENTIRE"),
          e("read_from_column", "A"), e("read_to_column", ""), e("read_from_row", "1"), e("read_to_row", ""),
          e("empty_col_header_prefix", "empty_"), e("column_name_mode", "EXCEL_COL_NAME")),
        c("advanced_settings", e("fail_on_differing_specs", False), e("append_path_column_Internals", False),
          e("path_column_name_Internals", "Path"), e("skip_empty_cols", False), e("skip_empty_rows", True),
          e("limit_data_rows_scanned", True), e("max_data_rows_scanned", 10000, "xlong"),
          e("save_table_spec_config_Internals", False), e("check_table_spec", False),
          e("use_15_digits_precision", True), e("skip_hidden_columns", True), e("skip_hidden_rows", True),
          e("replace_empty_strings_with_missings", True), e("reevaluate_formulas", False),
          e("formula_error_handling", "PATTERN"), e("formula_error_pattern", "#XL_EVAL_ERROR#")),
        c("encryption", c("password_Internals", e("SettingsModelID", "SMID_authentication"), e("EnabledStatus", True)),
          c("password", e("credentials", None), e("username", ""), e("selectedType", "NONE"))),
    )
    return node_settings(
        name="Excel Reader", factory="org.knime.ext.poi3.node.io.filehandling.excel.reader.ExcelTableReaderNodeFactory",
        bundle_name="KNIME POI (Excel) Plug-in", bundle_sym="org.knime.ext.poi3", feature_name="KNIME Excel Support",
        feature_sym="org.knime.features.ext.poi.feature.group", model=model, label=label,
        creation_config=c("node_creation_config", c("File System Connection"), c("Data Table")))


def python_script(script: str, label: str, n_in: int, n_out: int) -> str:
    model = c("model", e("script", script), e("python3_command", ""))
    creation = c("node_creation_config", c("Input object (pickled)"), c("Input table", table_ports(n_in)),
                 c("Output table", table_ports(n_out)), c("Output image"), c("Output object (pickled)"),
                 c("Python environment"))
    return node_settings(
        name="Python Script", factory="org.knime.python3.scripting.nodes2.script.PythonScriptNodeFactory",
        bundle_name="KNIME Python 3 - Scripting Nodes", bundle_sym="org.knime.python3.scripting.nodes",
        feature_name="KNIME Python Integration", feature_sym="org.knime.features.python3.scripting.feature.group",
        model=model, label=label, creation_config=creation, out_ports=n_out)


def if_switch(label: str) -> str:
    model = c("model", e("PortChoice", "top"), e("activate_all_outputs_during_configure", False))
    variables = c("variables", e("version", "V_2019_09_13"),
                  c("tree", c("PortChoice", e("used_variable", "PortChoice"), e("exposed_variable", None))))
    return node_settings(
        name="IF Switch", factory="org.knime.base.node.switches.manualif.ManualIfNodeFactory",
        bundle_name="KNIME Base Nodes", bundle_sym="org.knime.base", feature_name="KNIME Core",
        feature_sym="org.knime.features.base.feature.group", model=model, label=label, out_ports=2,
        variables=variables)


# --------------------------------------------------------------------------- workflow
NODES = [
    # id, carpeta, x, y, settings
    (1, "Excel Reader (#1)", 80, 120),
    (2, "Excel Reader (#2)", 80, 300),
    (3, "Python Script (#3)", 360, 200),
    (4, "IF Switch (#4)", 620, 200),
]
CONNECTIONS = [  # (src, src_port, dst, dst_port)  — puerto 0 = variables
    (1, 1, 3, 1), (2, 1, 3, 2), (3, 1, 4, 1),
]


def workflow_knime() -> str:
    nodes = "".join(
        c(f"node_{nid}", e("id", nid), e("node_settings_file", f"{folder}/settings.xml"),
          e("node_is_meta", False), e("node_type", "NativeNode"),
          e("ui_classname", "org.knime.core.node.workflow.NodeUIInformation"),
          c("ui_settings", c("extrainfo.node.bounds", e("array-size", 4), e("0", x), e("1", y),
                             e("2", 80), e("3", 80))))
        for nid, folder, x, y in NODES)
    conns = "".join(
        c(f"connection_{i}", e("sourceID", s), e("destID", d), e("sourcePort", sp), e("destPort", dp))
        for i, (s, sp, d, dp) in enumerate(CONNECTIONS))
    note = ("NOVA vs SAP - banderazo\n\n"
            "1) Doble clic en 'Excel Reader' de NOVA y de SAP -> elige tus archivos (vienen con ejemplos).\n"
            "2) Execute all. Clic en 'Python Script' -> salida 1 = resumen, salida 2 = detalle por PO.\n"
            "3) IF Switch: salida de arriba = hubo banderazo (aqui van los reportes); abajo = no.\n"
            "La columna de PO se detecta sola (PO Number, BSTKD, ...).")
    ann = c("annotation_0", e("text", note), e("contentType", "text/plain"), e("bgcolor", 16775885),
            e("x-coordinate", 40), e("y-coordinate", -90), e("width", 700), e("height", 150),
            e("alignment", "LEFT"), e("borderSize", 10), e("borderColor", 16766976), e("defFontSize", 12),
            e("annotation-version", 20151123), c("styles"))
    return doc("workflow.knime", e("created_by", f"{KNIME_VERSION}.v000000000000"),
               e("created_by_nightly", False), e("version", "5.1.0"), e("name", None),
               e("customDescription", None), e("state", "IDLE"), c("workflow_credentials"),
               c("annotations", ann), c("nodes", nodes), c("connections", conns))


def make_sample_excels(data: Path) -> None:
    from openpyxl import Workbook

    wb = Workbook()
    ws = wb.active
    ws.append(["PO Number", "DC", "Status"])
    for po, dc in [("4519001001", "6006"), ("4519001002", "6006"), ("4519001002", "7038"), ("4519001003", "7038")]:
        ws.append([int(po), dc, "Open"])
    wb.save(data / "ejemplo_nova.xlsx")
    wb = Workbook()
    ws = wb.active
    ws.append(["Sales Document", "Sold-to", "BSTKD"])
    for doc_no, po in [(1001, "004519001001"), (1002, "004519001002"), (1003, "004519001003"), (1004, "004519001003")]:
        ws.append([doc_no, "WALMART", po])
    wb.save(data / "ejemplo_sap.xlsx")


def build(out_dir: Path) -> Path:
    wf = out_dir / WF_NAME
    if wf.exists():
        shutil.rmtree(wf)
    data = wf / "data"
    shutil.copytree(ROOT / "knime" / "lib", data / "lib", ignore=shutil.ignore_patterns("__pycache__"))
    cfg = json.loads((ROOT / "knime" / "config" / "config.example.json").read_text(encoding="utf-8"))
    cfg["state_db"] = "state.sqlite"
    cfg["sap"]["po_column"] = ""
    cfg["matching"]["fire_once_per_day"] = False
    (data / "config.json").write_text(json.dumps(cfg, indent=2, ensure_ascii=False), encoding="utf-8")
    make_sample_excels(data)

    script = (ROOT / "knime" / "nodes" / "02_match_trigger.py").read_text(encoding="utf-8")
    settings = {
        1: excel_reader("ejemplo_nova.xlsx", "NOVA\n(elige tu archivo)"),
        2: excel_reader("ejemplo_sap.xlsx", "SAP\n(elige tu archivo)"),
        3: python_script(script, "Match NOVA vs SAP\n(banderazo)", 2, 2),
        4: if_switch("Hubo banderazo?\narriba = SI"),
    }
    for nid, folder, *_ in NODES:
        (wf / folder).mkdir(parents=True)
        (wf / folder / "settings.xml").write_text(settings[nid], encoding="utf-8")
    (wf / "workflow.knime").write_text(workflow_knime(), encoding="utf-8")
    return wf


def pack(wf: Path, target: Path) -> Path:
    with zipfile.ZipFile(target, "w", zipfile.ZIP_DEFLATED) as z:
        for p in sorted(wf.rglob("*")):
            if p.is_file() and "__pycache__" not in p.parts and p.name not in ("state.sqlite", ".knimeLock"):
                z.write(p, Path(wf.name) / p.relative_to(wf))
    return target


if __name__ == "__main__":
    out = Path(sys.argv[1]) if len(sys.argv) > 1 else ROOT / "dist"
    out.mkdir(parents=True, exist_ok=True)
    wf = build(out / "build")
    print(pack(wf, out / f"{WF_NAME}.knwf"))
