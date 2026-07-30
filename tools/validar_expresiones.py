#!/usr/bin/env python3
"""Valida la sintaxis de las expresiones de Power Automate.

Power Automate solo parsea las expresiones al guardar o al activar el flujo,
asi que un literal mal escrito sobrevive a la importacion y solo se descubre
cuando ya esta en el entorno. Esto lo detecta antes.

Comprueba, por expresion:
  - las comillas simples dentro de un literal van duplicadas
  - los parentesis quedan equilibrados
  - fuera de los literales solo aparece sintaxis de expresion

La tercera es la que atrapa el caso traicionero: un atributo HTML escrito con
comilla simple dentro de un literal cierra y reabre la cadena, de modo que el
recuento de comillas cuadra pero el texto del atributo queda suelto como si
fuera codigo.
"""
import json, glob, sys, string

FUERA_OK = set(string.ascii_letters + string.digits + "_(),.?[]@{}$+-*/ \t\r\n")


def revisar(expr):
    """Devuelve None si la expresion esta bien, o un mensaje con el fallo."""
    i = prof = 0
    dentro = False
    sueltos = []
    while i < len(expr):
        c = expr[i]
        if dentro:
            if c == "'":
                if expr[i + 1:i + 2] == "'":
                    i += 2
                    continue
                dentro = False
            i += 1
            continue
        if c == "'":
            dentro = True
        elif c == '(':
            prof += 1
        elif c == ')':
            prof -= 1
            if prof < 0:
                return f'parentesis de cierre de mas en la posicion {i}'
        elif c not in FUERA_OK:
            sueltos.append((i, c))
        i += 1
    if dentro:
        return 'literal sin cerrar: comilla simple sin escapar'
    if prof:
        return f'faltan {prof} parentesis de cierre'
    if sueltos:
        pos, car = sueltos[0]
        return (f'"{car}" suelto fuera de literal en la posicion {pos} '
                f'({len(sueltos)} en total) -> ...{expr[max(0, pos - 40):pos + 40]}...')
    return None


def expresiones(accion):
    ent = accion.get('inputs')
    if isinstance(ent, str):
        yield ent
    elif isinstance(ent, dict):
        for v in (ent.get('parameters') or {}).values():
            if isinstance(v, str):
                yield v


def recorrer(acciones, ruta, fallos):
    for nombre, accion in acciones.items():
        for v in expresiones(accion):
            if v.startswith('@'):
                fallo = revisar(v[1:])
                if fallo:
                    fallos.append((ruta + nombre, fallo, v))
        for rama in ('actions', 'else'):
            sub = accion.get(rama) or {}
            sub = sub.get('actions', {}) if rama == 'else' else sub
            if isinstance(sub, dict) and sub:
                recorrer(sub, f'{ruta}{nombre}/', fallos)


fallos = []
for fichero in sorted(glob.glob('solutions/AWGAutoconsolidaciones/Workflows/*.json')):
    d = json.load(open(fichero, encoding='utf-8'))
    recorrer(d['properties']['definition']['actions'],
             fichero.split('/')[-1][:14] + ' :: ', fallos)

for nombre, fallo, expr in fallos:
    print(f'FALLO {nombre}\n      {fallo}\n')
print(f'{len(fallos)} expresiones invalidas' if fallos
      else 'Todas las expresiones son sintacticamente validas')
sys.exit(1 if fallos else 0)
