#!/usr/bin/env python3
"""Prototipo del clasificador local (Consolidation / Cancellation / No Match).

Replica en Python lo que despues iran a ser expresiones de Power Automate,
para poder medir el acierto contra correos reales antes de tocar el flujo.
"""
import re

# Marcadores donde empieza el historial citado. Todo lo posterior se descarta.
CORTES = [
    r'\nFrom:\s',
    r'\n-----\s*Original Message\s*-----',
    r'\nOn\s.{0,80}\swrote:',
    r'\n_{10,}',
    r'FOR ROUTING USE WORKFLOW=',
    r'\nSent:\s',
]

# Ruido de cabecera que se quita antes de analizar.
RUIDO = [r'^\s*\[EXTERNAL\]\s*', r'^\s*Business Use\s*', r'^\s*WARNING:.*$']

CONSOLIDACION = [
    'consolidate', 'consolidated', 'consolidation',
    'ship together', 'ship with', 'ride with', 'riding with',
    'combine', 'pairing', 'load with', 'attach to',
]
CANCELACION = [
    'cancel', 'cancellation', 'cancelled', 'canceled',
    'delete', 'void the order', 'void po',
]
# Frases que niegan la peticion aunque aparezca el verbo.
NEGACIONES = [
    'do not consolidate', "don't consolidate", 'no longer need',
    'do not ship with', 'cancel the consolidation', 'cancel the ship with',
]

RE_PO = re.compile(r'\b(?:po|p\.o\.|purchase order)s?\b[^\n]{0,40}?(\d{4,8})', re.I)
# Los SO de AWG son numeros de 10 digitos y a menudo llegan en lista, uno por
# linea, lejos de la palabra "SO". Se detectan por forma, excluyendo telefonos.
RE_SO = re.compile(r'(?<![\d.-])(\d{10})(?![\d.-])')
RE_TEL = re.compile(r'(?:phone|tel|cell|ext)', re.I)


def recortar(cuerpo: str) -> str:
    """Deja solo el mensaje mas reciente, sin historial citado."""
    texto = cuerpo.replace('\r\n', '\n')
    corte = len(texto)
    for patron in CORTES:
        m = re.search(patron, texto, re.I)
        if m:
            corte = min(corte, m.start())
    texto = texto[:corte]
    for patron in RUIDO:
        texto = re.sub(patron, '', texto, flags=re.I | re.M)
    return texto.strip()


def identificadores(texto: str):
    pos = RE_PO.findall(texto)
    sos = [n for linea in texto.split('\n') if not RE_TEL.search(linea)
           for n in RE_SO.findall(linea)]
    return pos, sos


def clasificar(asunto: str, cuerpo: str):
    """Devuelve (tipo, motivo). El asunto cuenta: a veces el PO solo esta ahi."""
    reciente = recortar(cuerpo)
    ambito = f'{asunto}\n{reciente}'.lower()

    for frase in NEGACIONES:
        if frase in ambito:
            return 'No Match', f'negacion detectada: "{frase}"'

    pos, sos = identificadores(f'{asunto}\n{reciente}')
    tiene_id = bool(pos or sos)

    hit_can = next((p for p in CANCELACION if p in ambito), None)
    hit_con = next((p for p in CONSOLIDACION if p in ambito), None)

    # Cancelacion gana: "cancel the ride-with" es una cancelacion.
    if hit_can:
        if not tiene_id:
            return 'No Match', f'"{hit_can}" pero sin PO ni SO'
        return 'Cancellation', f'"{hit_can}" + ids {pos or sos}'
    if hit_con:
        if not tiene_id:
            return 'No Match', f'"{hit_con}" pero sin PO ni SO'
        return 'Consolidation', f'"{hit_con}" + ids {pos or sos}'
    return 'No Match', 'ninguna senal de consolidacion ni cancelacion'
