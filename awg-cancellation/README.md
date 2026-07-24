# AWG – Cancellation (flujo 00 + nuevo flujo RPA)

Objetivo:
1. **Modificar el flujo 00** (`GENERAL - 00 Email Router / Process Classifier`) para
   que OSSGenAI pueda identificar cuándo un correo es una **Cancellation**.
2. **Crear un flujo nuevo** que lea la respuesta, identifique las órdenes a cancelar
   y las envíe a un **RPA de cancelaciones** (Subject + contacto exactos).

---

## Parte 1 — Cambio en el flujo 00 (LISTO)

El 00 responde el hilo a `ossgenai.im@pg.com` con un prompt clasificador
(acción **`GENERAL - 00 HTTP Reply Original to OSSGenAI SameThread`**, campo `Body`,
dentro del `comment`). Hoy permite 3 valores de `ProcessType`:
`Consolidation`, `Wakefern Appointments`, `No Match`.

**Cambios aplicados** (ver `classifier-00-NEW-comment.html` y
`GENERAL-00-MODIFIED-definition.json`):

1. Se añadió una **regla de clasificación de Cancellation**:
   > **Cancellation rule:** Classify as Cancellation only when the latest actual
   > customer email contains BOTH: 1) an explicit customer request to cancel, void,
   > or delete one or more orders (cancel, please cancel, cancellation, void, delete
   > order, remove the order); and 2) at least one explicit PO or order number tied
   > to PO/order context (PO, PO#, PO #, purchase order, order, SO, sales order)…
   > A cancellation request is never a Consolidation. If either requirement is
   > missing, classify as No Match.

2. Se amplió la lista de valores de salida (de **tres** a **cuatro**):
   `Consolidation, Cancellation, Wakefern Appointments, No Match`.

Todo lo demás del prompt (READ SCOPE, exclusiones, formato de tabla de 1 columna /
1 fila, el `AWG-FLOW-ID`/`AWG-STEP` oculto) se mantiene **idéntico**.

### Cómo aplicarlo (solución **managed**)

> ⚠️ La solución exportada es **managed** (`Managed=1`, v1.0.0.1, publisher
> `awgnacso`). Una solución managed **no se edita en su entorno destino**. Aplica el
> cambio en el **entorno de origen / versión unmanaged** y reexporta, o:

- **Opción recomendada:** en tu entorno DEV, abre el flujo 00 (versión unmanaged),
  entra a la acción `GENERAL - 00 HTTP Reply Original to OSSGenAI SameThread` y
  reemplaza el texto del prompt por el de `classifier-00-NEW-comment.html`
  (es solo el `comment`; conserva el resto del JSON del `Body`). Sube la versión y
  reexporta la solución (managed) para PROD.
- **Alternativa:** importa `GENERAL-00-MODIFIED-definition.json` como definición del
  flujo (mismo flujo) si trabajas por definición.

> Nota de formato: dentro del flujo, los saltos del prompt van como `\n` **literal**
> (backslash-n), no como salto de línea real. El archivo `.json` ya lo trae así.

---

## Parte 2 — Nuevo flujo de Cancellation (PENDIENTE de datos)

Diseño propuesto, **consistente con la coreografía actual** (00 clasifica →
OSSGenAI responde `ProcessType` → un flujo downstream actúa):

```
[00] clasifica -> ProcessType = Cancellation
        │
        ▼
[Cancellation-A]  (trigger: correo de ossgenai.im@pg.com con la tabla ProcessType)
        │   - detecta ProcessType = Cancellation
        │   - responde el hilo a OSSGenAI con un prompt de EXTRACCIÓN
        │     (WORKFLOW=na_om_data_agent) exigiendo una "Formatted Table"
        │     EXACTA con las órdenes a cancelar
        ▼
[Cancellation-B]  (trigger: respuesta de OSSGenAI con la Formatted Table)
        │   - parsea la tabla con el MISMO método del flujo 02
        │     (Formatted Table -> </table> -> skip(split('<tr'),2) -> celdas </td>)
        │   - arma el correo al RPA de cancelaciones
        ▼
   Correo al RPA  (Subject EXACTO + contacto EXACTO)
```

Esto reutiliza el parser de tablas ya probado en `AWG-CON-02` y el patrón de
prompt/AWG-STEP del 00, para que la lectura sea siempre idéntica.

### Necesito para construirlo

1. **Subject EXACTO** del correo al RPA de cancelaciones.
2. **Contacto (email)** del RPA.
3. **Identificador** que el RPA necesita por orden: `PO #`, `SO #`, `Delivery #` (o varios).
4. **Formato que espera el RPA** en el cuerpo (¿tabla?, ¿bloque?, ¿un PO por línea?).
5. (Opcional) ¿Extracción vía OSSGenAI como arriba (recomendado), o el flujo nuevo
   debe parsear directamente el correo del cliente sin pasar de nuevo por OSSGenAI?
