#!/usr/bin/env python3
"""
Arma el .zip importable de Power Automate a partir de los archivos sueltos
de flow/.

Un paquete de importacion legacy ("Import Package (Legacy)") necesita esta
estructura exacta:

    manifest.json
    Microsoft.Flow/flows/manifest.json
    Microsoft.Flow/flows/<flowResourceId>/definition.json
    Microsoft.Flow/flows/<flowResourceId>/apisMap.json
    Microsoft.Flow/flows/<flowResourceId>/connectionsMap.json

Uso:
    python3 tools/build_flow_zip.py [-o dist/Wakefern11sAutopushV15.zip]
"""

from __future__ import annotations

import argparse
import json
import os
import zipfile

FLOW_RESOURCE_ID = "6d2947b7-73a5-4d97-9cb2-0c991ffa6c32"
FLOW_DIR = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "flow")


def compact(path: str) -> str:
    """Reserializa el JSON compacto, como lo produce la exportacion real."""
    with open(path, encoding="utf-8") as fh:
        return json.dumps(json.load(fh), separators=(",", ":"))


def main() -> None:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument(
        "-o",
        "--output",
        default="dist/Wakefern11sblockAutopushV15WHOLESALE2ALERTS.zip",
    )
    args = parser.parse_args()

    base = f"Microsoft.Flow/flows/{FLOW_RESOURCE_ID}"
    entries = {
        "manifest.json": compact(os.path.join(FLOW_DIR, "manifest.json")),
        "Microsoft.Flow/flows/manifest.json": compact(
            os.path.join(FLOW_DIR, "flows-manifest.json")
        ),
        f"{base}/definition.json": compact(os.path.join(FLOW_DIR, "definition.json")),
        f"{base}/apisMap.json": compact(os.path.join(FLOW_DIR, "apisMap.json")),
        f"{base}/connectionsMap.json": compact(
            os.path.join(FLOW_DIR, "connectionsMap.json")
        ),
    }

    outdir = os.path.dirname(args.output)
    if outdir:
        os.makedirs(outdir, exist_ok=True)

    with zipfile.ZipFile(args.output, "w", zipfile.ZIP_DEFLATED) as zf:
        for name, content in entries.items():
            zf.writestr(name, content)

    print(f"Paquete generado: {args.output}")
    for name in entries:
        print(f"   {name}")


if __name__ == "__main__":
    main()
