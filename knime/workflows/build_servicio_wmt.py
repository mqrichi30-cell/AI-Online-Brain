"""Genera WMT_Reportes_Servicio.knwf: servicio REST para KNIME Hub.

    Container Input (JSON) "solicitud"  ->  Python Script (motor wmt_reportes)
        ->  String to JSON  ->  Container Output (JSON) "respuesta"

Power Automate lo llama con HTTP (Basic auth = contrasena de aplicacion del Hub).
"""
from __future__ import annotations

import json
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent))
import build_workflow as bw  # noqa: E402  (mismo mini-constructor XML del otro .knwf)

ROOT = Path(__file__).resolve().parents[2]
WF_NAME = "WMT_Reportes_Servicio"

EJEMPLO = {
    "reporte": "Pricing Report", "turno": "AM", "fecha": "2026-11-09",
    "asunto_formato": "WM Turn Pricing {FECHA:MM.dd.yyyy}", "machote_eml_base64": "", "modo_prueba": True,
}

SCRIPT_MAIN = '''

# ============================== nodo KNIME ==============================
import json as _json
import knime.scripting.io as knio

_df = knio.input_tables[0].to_pandas()
_v = _df.iloc[0, 0] if len(_df) else {}
if isinstance(_v, (dict, list)):
    _sol = _v
else:
    try:
        _sol = _json.loads(str(_v))
    except ValueError:
        import ast as _ast
        _sol = _ast.literal_eval(str(_v))   # KNIME a veces entrega el JSON como repr de Python
# TODO (lunes): aqui entran las tablas de SAP de los nodos DB -> datos={"Pivot": df, ...}
_resp = responder(_sol, None)
knio.output_tables[0] = knio.Table.from_pandas(pd.DataFrame({"respuesta": [_json.dumps(_resp, ensure_ascii=False)]}))
'''


def json_input(param: str, default: dict) -> str:
    model = bw.c("model", bw.e("parameterName", param), bw.e("json", json.dumps(default, ensure_ascii=False)),
                 bw.e("description", "Peticion de Power Automate: reporte, turno, fecha, asunto_formato, machote_eml_base64"),
                 bw.e("useFullyQualifiedName", False))
    return bw.node_settings(
        name="Container Input (JSON)", factory="org.knime.json.node.input.JSONInputNodeFactory",
        bundle_name="KNIME JSON-Processing", bundle_sym="org.knime.json", feature_name="KNIME JSON-Processing",
        feature_sym="org.knime.features.json.feature.group", model=model, label="Peticion de\nPower Automate")


def string_to_json(col: str) -> str:
    model = bw.c("model", bw.e("input.column", col), bw.e("remove.input.column", True),
                 bw.e("new.column.name", col), bw.e("fail.on.error", True), bw.e("allow.comments", False))
    return bw.node_settings(
        name="String to JSON", factory="org.knime.json.node.fromstring.StringToJSONNodeFactory",
        bundle_name="KNIME JSON-Processing", bundle_sym="org.knime.json", feature_name="KNIME JSON-Processing",
        feature_sym="org.knime.features.json.feature.group", model=model, label="a JSON")


def json_output(param: str, col: str) -> str:
    model = bw.c("model", bw.e("parameterName", param), bw.e("useFullyQualifiedName", False),
                 bw.e("jsonColumnName", col), bw.e("keepOneRowTablesSimple", True),
                 bw.e("exampleJson", json.dumps({"ok": True, "error": "", "correos": []})),
                 bw.e("description", "Correos listos: asunto, html, to, resumen, adjunto_nombre, adjunto_base64"))
    return bw.node_settings(
        name="Container Output (JSON)", factory="org.knime.json.node.output.JSONOutputNodeFactory",
        bundle_name="KNIME JSON-Processing", bundle_sym="org.knime.json", feature_name="KNIME JSON-Processing",
        feature_sym="org.knime.features.json.feature.group", model=model, label="Respuesta a\nPower Automate",
        out_ports=0)


NODES = [(1, "Container Input (JSON) (#1)", 60, 160), (2, "Python Script (#2)", 280, 160),
         (3, "String to JSON (#3)", 500, 160), (4, "Container Output (JSON) (#4)", 720, 160)]
CONNS = [(1, 1, 2, 1), (2, 1, 3, 1), (3, 1, 4, 1)]


def workflow_knime() -> str:
    nodes = "".join(
        bw.c(f"node_{nid}", bw.e("id", nid), bw.e("node_settings_file", f"{folder}/settings.xml"),
             bw.e("node_is_meta", False), bw.e("node_type", "NativeNode"),
             bw.e("ui_classname", "org.knime.core.node.workflow.NodeUIInformation"),
             bw.c("ui_settings", bw.c("extrainfo.node.bounds", bw.e("array-size", 4), bw.e("0", x), bw.e("1", y),
                                      bw.e("2", 80), bw.e("3", 80))))
        for nid, folder, x, y in NODES)
    conns = "".join(bw.c(f"connection_{i}", bw.e("sourceID", s), bw.e("destID", d), bw.e("sourcePort", sp),
                         bw.e("destPort", dp)) for i, (s, sp, d, dp) in enumerate(CONNS))
    note = ("WMT Reportes - servicio REST (KNIME Hub)\n"
            "Power Automate manda: reporte, turno, fecha, asunto_formato, machote_eml_base64.\n"
            "Regresa: correos con asunto exacto, cuerpo HTML y Excel adjunto (base64).\n"
            "Hoy usa DATOS DE PRUEBA. Pendiente: conectar nodos DB de SAP (VBAK/VBAP/VBEP/LIKP) al Python Script.\n"
            "Deploy: Hub -> este workflow -> Deploy -> Service. Copia la URL de 'Open API docs' al flujo de Power Automate.")
    ann = bw.c("annotation_0", bw.e("text", note), bw.e("contentType", "text/plain"), bw.e("bgcolor", 16775885),
               bw.e("x-coordinate", 40), bw.e("y-coordinate", -60), bw.e("width", 780), bw.e("height", 130),
               bw.e("alignment", "LEFT"), bw.e("borderSize", 10), bw.e("borderColor", 16766976), bw.e("defFontSize", 12),
               bw.e("annotation-version", 20151123), bw.c("styles"))
    return bw.doc("workflow.knime", bw.e("created_by", f"{bw.KNIME_VERSION}.v000000000000"),
                  bw.e("created_by_nightly", False), bw.e("version", "5.1.0"), bw.e("name", None),
                  bw.e("customDescription", None), bw.e("state", "IDLE"), bw.c("workflow_credentials"),
                  bw.c("annotations", ann), bw.c("nodes", nodes), bw.c("connections", conns))


def build(out_dir: Path) -> Path:
    import shutil
    wf = out_dir / WF_NAME
    if wf.exists():
        shutil.rmtree(wf)
    motor = (ROOT / "knime" / "lib" / "nova_sap" / "wmt_reportes.py").read_text(encoding="utf-8")
    settings = {
        1: json_input("solicitud", EJEMPLO),
        2: bw.python_script(motor + SCRIPT_MAIN, "Motor de reportes\n(SAP -> correos)", 1, 1),
        3: string_to_json("respuesta"),
        4: json_output("respuesta", "respuesta"),
    }
    for nid, folder, *_ in NODES:
        (wf / folder).mkdir(parents=True)
        (wf / folder / "settings.xml").write_text(settings[nid], encoding="utf-8")
    (wf / "workflow.knime").write_text(workflow_knime(), encoding="utf-8")
    return wf


if __name__ == "__main__":
    out = Path(sys.argv[1]) if len(sys.argv) > 1 else ROOT / "dist"
    out.mkdir(parents=True, exist_ok=True)
    wf = build(out / "build")
    print(bw.pack(wf, out / f"{WF_NAME}.knwf"))
