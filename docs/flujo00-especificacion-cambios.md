# Flujo `GENERAL - 00 Email Router / Process Classifier` — cambios a aplicar

**Entorno:** Personal Productivity (default)
**Solución:** AWGAutoconsolidaciones
**Flujo:** GENERAL - 00 Email Router / Process Classifier

## Problema

El flujo etiqueta todos los correos entrantes y luego no hace nada mas.
No crea el registro de tracking ni responde a la IA.

Dos causas independientes:

1. La comprobacion de "ya enrutado" compara tambien por asunto. Como los
   asuntos se repiten, casi todo correo se da por ya enrutado y se descarta.
2. El marcado de leido/categorias se ejecuta ANTES de esa comprobacion, asi
   que etiqueta incluso los correos que va a descartar.

---

## Cambio 1 — Filtro de la consulta a SharePoint

**Accion:** `GENERAL - 00 Check Already Routed By ConversationId`
(es una accion de SharePoint "Get items")

**Campo:** `Filter Query` (esta en parametros avanzados; puede aparecer como
"Filter Query" o "Consulta de filtro")

**Valor actual** (contiene tres condiciones unidas por `or`):

```
OriginalMessageId eq '@{replace(outputs('Compose_Original_Immutable_ID'),'''','''''')}' or OriginalConversationId eq '@{replace(outputs('Compose_Original_Conversation_ID'),'''','''''')}' or OriginalSubject eq '@{replace(coalesce(triggerOutputs()?['body/subject'], ''),'''','''''')}'
```

**Valor nuevo** (se elimina unicamente el tercer `or ... OriginalSubject ...`):

```
OriginalMessageId eq '@{replace(outputs('Compose_Original_Immutable_ID'),'''','''''')}' or OriginalConversationId eq '@{replace(outputs('Compose_Original_Conversation_ID'),'''','''''')}'
```

Las dos primeras condiciones se dejan exactamente igual, sin retocar comillas
ni espacios.

---

## Cambio 2 — Reubicar el marcado de leido/categorias

**Accion:** `GENERAL - 00 HTTP Patch Original Read Tags`

**Donde esta ahora:** en la cadena principal, entre
`Compose Original Immutable ID` y `Compose Original Conversation ID`.

**Donde debe quedar:** dentro de la rama **False / Else** del condicional
`GENERAL - 00 Guard Already Routed`, como primera accion, justo antes de
`Compose FlowId`.

**Reconexiones que implica:**

- `Compose Original Conversation ID` pasa a ejecutarse despues de
  `GENERAL - 00 Get Current Immutable ID`.
- `Compose FlowId` pasa a ejecutarse despues de
  `GENERAL - 00 HTTP Patch Original Read Tags`.

La accion no cambia por dentro. Sus parametros siguen siendo:

- Uri: `v1.0/me/messages/@{encodeUriComponent(outputs('Compose_Original_Immutable_ID'))}`
- Method: `PATCH`
- CustomHeader1: `Prefer: IdType="ImmutableId"`
- ContentType: `application/json`
- Body: `{"isRead":true,"categories":["RPA","Cristhofer Marin"]}`

> Nota: en el Body original la palabra es `Marín`, con tilde. Debe conservarse
> con tilde, tal cual estaba.

---

## Que NO se debe tocar

- El trigger `When a new email arrives (V3)`, ni su carpeta ni sus condiciones.
- Los otros seis flujos de la solucion.
- El prompt que se envia a `ossgenai.im@pg.com`.
- Las connection references.
- El estado on/off de ningun flujo.

---

## Comprobacion final

1. El flujo guarda sin errores de validacion.
2. Enviar un correo de prueba con un asunto nuevo, desde un hilo nuevo.
3. En Run history, la ejecucion debe llegar hasta
   `GENERAL - 00 Move To Email Route 00`.
4. Debe aparecer un registro nuevo en la lista de SharePoint con
   `Phase = 00-RouterSent`.
5. Debe salir una respuesta dirigida a `ossgenai.im@pg.com`.
