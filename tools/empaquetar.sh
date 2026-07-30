#!/usr/bin/env bash
# Empaqueta la solucion tal y como la exporta Power Platform.
#
# Power Platform rechaza el paquete si contiene entradas de directorio o
# carpetas que no reconoce, y el error que muestra es solo "An unknown error
# occurred". De ahi que aqui se listen los miembros de forma explicita y se
# use -D para no escribir entradas de directorio.
set -euo pipefail
cd "$(dirname "$0")/.."
SRC=solutions/AWGAutoconsolidaciones
VER=$(grep -o '<Version>[^<]*</Version>' "$SRC/solution.xml" | head -1 | tr -d '<>/' | sed 's/Version//g')
OUT="dist/AWGAutoconsolidaciones_${VER//./_}.zip"

extra=$(find "$SRC" -maxdepth 1 -mindepth 1 \
  ! -name Workflows ! -name customizations.xml ! -name solution.xml \
  ! -name '\[Content_Types\].xml')
if [ -n "$extra" ]; then
  echo "Hay ficheros que no pertenecen a la solucion:" >&2
  echo "$extra" >&2
  exit 1
fi

mkdir -p dist && rm -f "$OUT"
( cd "$SRC" && zip -q -r -X -D "../../$OUT" \
    customizations.xml solution.xml Workflows '[Content_Types].xml' )
echo "$OUT"
unzip -l "$OUT" | tail -3
