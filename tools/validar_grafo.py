#!/usr/bin/env python3
"""Comprueba que cada referencia outputs()/body() sea alcanzable por runAfter.

Logic Apps exige que una accion solo referencie la salida de otra si esa otra
esta en su camino de runAfter, o dentro de un scope que lo este. Una referencia
a una accion que quedo en una rama paralela pasa cualquier validador de
sintaxis y solo la rechaza el servidor al guardar, con este mensaje:

  "cannot reference action X. Action X must either be in 'runAfter' path or
   within a scope action on the 'runAfter' path"
"""
import json, glob, re, sys


def alcanzables(nombre, acciones, heredadas):
    """Acciones cuyas salidas puede usar `nombre`: las heredadas del scope padre
    mas todo lo que le precede por runAfter entre sus hermanas."""
    vistas, pila = set(heredadas), list((acciones[nombre].get('runAfter') or {}))
    while pila:
        act = pila.pop()
        if act in vistas or act not in acciones:
            continue
        vistas.add(act)
        pila.extend(acciones[act].get('runAfter') or {})
        # una accion de scope tambien expone lo que contiene
        for rama in ('actions', 'else'):
            sub = acciones[act].get(rama) or {}
            sub = sub.get('actions', {}) if rama == 'else' else sub
            if isinstance(sub, dict):
                vistas.update(sub)
    return vistas


REF = re.compile(r"(?:outputs|body)\('([^']+)'\)")


def recorrer(acciones, heredadas, ruta, fallos):
    for nombre, accion in acciones.items():
        propias = alcanzables(nombre, acciones, heredadas) | {nombre}
        for ref in set(REF.findall(json.dumps(accion.get('inputs')))):
            if ref not in propias and not ref.startswith('When_'):
                fallos.append((ruta + nombre, ref))
        for rama in ('actions', 'else'):
            sub = accion.get(rama) or {}
            sub = sub.get('actions', {}) if rama == 'else' else sub
            if isinstance(sub, dict) and sub:
                recorrer(sub, propias, f'{ruta}{nombre}/', fallos)


fallos = []
for f in sorted(glob.glob('solutions/AWGAutoconsolidaciones/Workflows/*.json')):
    d = json.load(open(f, encoding='utf-8'))
    recorrer(d['properties']['definition']['actions'], set(),
             f.split('/')[-1][:14] + ' :: ', fallos)

for accion, ref in fallos:
    print(f'FALLO {accion}\n      no alcanza a {ref} por runAfter\n')
print(f'{len(fallos)} referencias inalcanzables' if fallos
      else 'Todas las referencias son alcanzables por runAfter')
sys.exit(1 if fallos else 0)
