"""Lógica del trigger NOVA vs SAP.

Regla de negocio:
  1. Se cuentan POs ÚNICAS (una PO duplicada cuenta una sola vez) en NOVA y en SAP.
  2. El trigger ("banderazo") se dispara cuando:
       a) #POs únicas NOVA == #POs únicas SAP, y
       b) TODAS las POs de NOVA existen en SAP.
     (a + b implica que ambos conjuntos son idénticos.)
  3. Por seguridad NO se dispara si NOVA viene vacío (normalmente significa que
     la extracción falló o el portal no cargó), salvo que se permita explícitamente.
"""
from __future__ import annotations

import math
import re
from dataclasses import dataclass, field
from typing import Iterable

_SCI_NOTATION = re.compile(r"^\d+(\.\d+)?[eE]\+?\d+$")
_FLOAT_INT = re.compile(r"^(\d+)\.0+$")


def normalize_po(value, strip_leading_zeros: bool = True) -> str | None:
    """Normaliza un número de PO para que NOVA y SAP sean comparables.

    - Quita espacios y caracteres no alfanuméricos de los extremos.
    - Convierte floats de Excel (1234567890.0) y notación científica a entero.
    - Opcionalmente quita ceros a la izquierda (SAP suele rellenar con ceros).
    Devuelve None para vacíos / NaN.
    """
    if value is None:
        return None
    if isinstance(value, float):
        if math.isnan(value):
            return None
        if value.is_integer():
            value = int(value)
    s = str(value).strip()
    if not s or s.lower() in {"nan", "none", "null"}:
        return None
    if _SCI_NOTATION.match(s):
        s = str(int(float(s)))
    m = _FLOAT_INT.match(s)
    if m:
        s = m.group(1)
    s = re.sub(r"\s+", "", s).upper()
    s = s.strip("-_./#")
    if strip_leading_zeros and s.isdigit():
        s = s.lstrip("0") or "0"
    return s or None


@dataclass
class TriggerResult:
    fired: bool
    reason: str
    nova_rows: int
    sap_rows: int
    nova_unique: int
    sap_unique: int
    nova_duplicates: dict[str, int] = field(default_factory=dict)
    sap_duplicates: dict[str, int] = field(default_factory=dict)
    missing_in_sap: list[str] = field(default_factory=list)
    extra_in_sap: list[str] = field(default_factory=list)

    def summary(self) -> dict:
        return {
            "trigger_fired": self.fired,
            "reason": self.reason,
            "nova_rows": self.nova_rows,
            "sap_rows": self.sap_rows,
            "nova_unique_pos": self.nova_unique,
            "sap_unique_pos": self.sap_unique,
            "nova_duplicate_pos": len(self.nova_duplicates),
            "sap_duplicate_pos": len(self.sap_duplicates),
            "missing_in_sap_count": len(self.missing_in_sap),
            "extra_in_sap_count": len(self.extra_in_sap),
        }


def _count(pos: Iterable, strip_zeros: bool) -> tuple[int, dict[str, int]]:
    counts: dict[str, int] = {}
    rows = 0
    for raw in pos:
        po = normalize_po(raw, strip_zeros)
        if po is None:
            continue
        rows += 1
        counts[po] = counts.get(po, 0) + 1
    return rows, counts


def evaluate_trigger(
    nova_pos: Iterable,
    sap_pos: Iterable,
    strip_leading_zeros: bool = True,
    allow_empty: bool = False,
) -> TriggerResult:
    nova_rows, nova_counts = _count(nova_pos, strip_leading_zeros)
    sap_rows, sap_counts = _count(sap_pos, strip_leading_zeros)

    nova_set, sap_set = set(nova_counts), set(sap_counts)
    missing = sorted(nova_set - sap_set)
    extra = sorted(sap_set - nova_set)

    if not nova_set and not allow_empty:
        fired, reason = False, "NOVA sin POs: posible falla de extracción, no se dispara"
    elif len(nova_set) != len(sap_set):
        fired = False
        reason = f"Conteo distinto: NOVA={len(nova_set)} vs SAP={len(sap_set)}"
    elif missing:
        fired = False
        reason = f"{len(missing)} PO(s) de NOVA no están en SAP"
    else:
        fired, reason = True, "Match completo: mismo conteo y todas las POs de NOVA están en SAP"

    return TriggerResult(
        fired=fired,
        reason=reason,
        nova_rows=nova_rows,
        sap_rows=sap_rows,
        nova_unique=len(nova_set),
        sap_unique=len(sap_set),
        nova_duplicates={k: v for k, v in nova_counts.items() if v > 1},
        sap_duplicates={k: v for k, v in sap_counts.items() if v > 1},
        missing_in_sap=missing,
        extra_in_sap=extra,
    )


def build_detail_rows(nova_pos: Iterable, sap_pos: Iterable, strip_leading_zeros: bool = True) -> list[dict]:
    """Tabla de detalle PO por PO (para auditoría / reporte de diferencias)."""
    _, n = _count(nova_pos, strip_leading_zeros)
    _, s = _count(sap_pos, strip_leading_zeros)
    rows = []
    for po in sorted(set(n) | set(s)):
        status = "MATCH" if po in n and po in s else ("SOLO_NOVA" if po in n else "SOLO_SAP")
        rows.append({"po": po, "veces_nova": n.get(po, 0), "veces_sap": s.get(po, 0), "estatus": status})
    return rows
