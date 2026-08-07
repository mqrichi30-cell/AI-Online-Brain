# Portafolio 90 días — versión legible para sala grande

Rediseño de la diapositiva del PDCA que resultó ilegible en pantalla de televisor:
el original condensaba las 12 automatizaciones en una sola lámina con cuerpo de
texto a **7,5–8,5 pt**.

## Qué cambió

| | Original | Versión legible |
|---|---|---|
| Diapositivas | 1 | 7 + anexo |
| Tamaño mínimo del contenido | 7,5 pt | 15 pt |
| Cuerpo principal (beneficios) | 8 pt | 23 pt |
| Nombres de proyecto | 8,5 pt | 19–34 pt |
| Estado | texto de color, 8,5 pt | pastilla sólida con texto oscuro, 14–20 pt |

Se conservan el fondo, la paleta (`E8EFF9` / `9FB4D2` / `1B3157` / `F3B65F` /
`7CE7AE` / `39C6F5` / `FF8C7C`), Calibri y el texto original palabra por palabra.

## Estructura

1. **Portada / resumen** — 12 automatizaciones, contadores por estado y el Top 5 con estado y fecha.
2. **–6. Detalle por proyecto** (uno por lámina) — beneficios, impactos OTSR, owners, partners, fecha.
3. **Las otras 7 automatizaciones** — 3 en progreso y 4 completadas.
4. **Anexo** — la tabla completa original, intacta, como material de respaldo.

Cada lámina lleva notas del orador.

## Versión de 2 láminas (la vigente)

`Portafolio_90dias_2laminas.pptx` — el formato pedido para el PDCA: una lámina
**In Progress** y una **Completed**.

- **In Progress**: banda destacada del **SWAT Team PGP** (ciclo de 6 semanas, duplas,
  mentor semanal rotativo, kanban, sync semanal) y 9 tarjetas — 6 en curso,
  1 en discovery y 2 en on hold (`PGP Ion` y `ZE Display Auto-Release`), con
  follow-up a finales de agosto.
- **Completed**: las 5 automatizaciones vivas, en lista de ancho completo a 24 pt.

La mecánica completa del SWAT va en las notas del orador, no en la lámina.

### Supuestos sobre los datos

- La fila `ZE Category Emails & Display Auto-Release` del original se parte en dos:
  **ZE Category Emails** (Cristhofer, completado) y **ZE Display Auto-Release**
  (OMA / Mariela, on hold, help needed from COTS).
- `COTS` se escribe tal como se indicó; el original usa `COTC` para DSD KNIME.
- `EDI Orders Validation` se mantiene en Discovery.
- Las tarjetas no llevan fecha: no caben junto al owner a tamaño legible.
- En la tarjeta de ZE aparece solo OMA (Mariela); K. Aguilar y B. Miller están en las notas.

## Regenerar

```bash
npm install pptxgenjs
node src/build2.js               # -> Portafolio_90dias_2laminas.pptx

# versión larga de 7 láminas + anexo
node src/build.js                # genera new.pptx (láminas 1–7)
python3 src/merge_appendix.py    # añade el anexo -> Portafolio_90dias_v3_legible.pptx
```

Los scripts leen `bg.png` desde su propio directorio. `merge_appendix.py` espera
`new.pptx` y `original.pptx` en el directorio de trabajo;
`src/Portafolio_90dias_v2_11_original.pptx` es la fuente del anexo.
