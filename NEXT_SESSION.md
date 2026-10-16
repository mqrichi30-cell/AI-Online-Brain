# Próxima sesión: diagnosticar por qué no llegan las aprobaciones

Copia el bloque de abajo y pégalo tal cual en el chat de la extensión **Claude in Chrome**. Antes, inicia sesión en Power Automate en esa ventana de Chrome.

---

```text
Necesito que revises en este Chrome (ya inicié sesión en Power Automate) por qué NO me están llegando los correos de aprobación de mis flujos, me digas la causa exacta y lo corrijas si se puede desde la interfaz.

CONTEXTO
- Solución de Power Automate: "AWG Autoconsolidaciones" (versión 1.0.0.38), entorno "Personal Productivity (default)".
- Flujo a revisar primero (paso 0):
  https://make.powerautomate.com/environments/Default-3596192b-fdf5-4e2c-a6fa-acb706c963d8/flows/37e4b4ef-7856-432e-966b-9f96e3304802?v3=true
- Los 6 flujos tienen esta estructura: variables arriba → "Scope - Try" (todo el flujo) → "Scope - Catch" (no se modifica).
- Antes de cada correo a ossgenai.im@pg.com hay una aprobación. La acción se llama "<PREFIJO>_Aprobacion_Enviar_Solicitud", y prefijos como GENERAL_-_00, AWG_-_CON_-_01, _03, _04_Aprobacion_Tabla, _04_Aprobacion_10Dias, _05_Aprobacion_Cliente y _05_Aprobacion_NoCol.
  - Versión Outlook: es "Enviar correo electrónico con opciones" (Office 365 Outlook) y se envía a mi propio buzón (GET /me).
  - Versión SMTP: es "Iniciar y esperar una aprobación" (conector Aprobaciones).
  - Después viene "<PREFIJO>_Aprobacion_Guard_Aprobado". Si se rechaza o pasan 24 h, termina en Terminate Cancelled.
- En el flujo 00, antes de la aprobación hay: "Esperar_horario_laboral_L-V_5am_a_3pm_CR" (Wait), filtros (Guard_Block_*, Guard_Only_New_Email >24h, Guard_Already_Routed) y la clasificación. La aprobación solo se pide cuando la clasificación NO es Cancellation / Return / ZQs ni Resolved_Locally = YES.

QUÉ HACER (en orden)
1. Abre el flujo del enlace y revisa:
   - si está activado (On);
   - en "Detalles / Connections", si todas las conexiones están OK (Office 365 Outlook, SharePoint y, si existen, Aprobaciones, SMTP y Conversión de contenido).
2. Abre el historial de ejecuciones (28 días). Dime cuántas corridas hay y con qué estados (Succeeded / Running / Failed / Cancelled).
3. Abre las 3 corridas más recientes que deberían haber pedido aprobación. En cada una, expande "Scope - Try" y dime:
   a. hasta qué acción llegó, y qué rama de cada condición tomó (GENERAL_-_00_Guard_Block_*, Guard_Only_New_Email, Guard_Already_Routed, Guard_Rule_Is_Cancellation);
   b. el valor de GENERAL_-_00_DEBUG_Rule_Result (la clasificación);
   c. el estado de GENERAL_-_00_Aprobacion_Get_Aprobador (Outputs: mail / userPrincipalName) y de GENERAL_-_00_Aprobacion_Enviar_Solicitud: Running, Succeeded, Failed o Skipped, con el error completo si falló;
   d. si "Scope - Catch" se ejecutó y qué dice Compose_-_Error_Message.
4. Según el resultado:
   - Si la acción de aprobación está Running o Succeeded (el correo sí se envió):
     - abre Outlook Web (outlook.office.com) y busca "[AWG CON] Aprobación requerida" en todas las carpetas: Bandeja de entrada, Otros, Correo no deseado, la carpeta "Marín, Cristhofer - AWG + Wakefern" y las carpetas de reglas;
     - revisa Configuración → Correo → Reglas, por si alguna mueve correos con "[AWG CON]";
     - revisa si el remitente "Microsoft Power Automate" / flow-noreply@microsoft.com está bloqueado o en cuarentena;
     - si es la versión SMTP, revisa Power Automate → Aprobaciones → Recibidas.
   - Si está Skipped o no llegó: dime qué condición la cortó y por qué, con los valores de entrada.
   - Si está Failed: copia el error exacto y propón la corrección.
5. Repite los pasos 2 y 3 en los flujos "AWG - Consolidation - 01 - Send Con Prompt", "03 - Send MOC Form" y "04 - Query Con Status" si el flujo 00 no explica el problema.
6. Corrige solo lo que sea seguro desde la interfaz: reconectar una conexión, activar un flujo, aprobar una solicitud pendiente si yo te lo confirmo. Antes de editar acciones del flujo, guardar o reenviar corridas, pregúntame. No borres nada.

AL TERMINAR, dame:
- la causa exacta (con captura o texto del error);
- qué corregiste;
- qué cambio necesito en la solución (para pasarlo al repo AI-Online-Brain, archivo solutions/AWGAutoconsolidaciones/tools/build.py) si el problema está en la definición del flujo.
```

---

## Si luego hay que cambiar la solución

Abre una sesión de Claude Code sobre el repo `mqrichi30-cell/AI-Online-Brain`, rama `claude/try-catch-email-approval-7y2iaq`, y pega:

```text
Lee CLAUDE.md y NEXT_SESSION.md. El diagnóstico en Power Automate dio esto: <pega aquí el resumen de la extensión>.
Corrige la causa en solutions/AWGAutoconsolidaciones/tools/build.py, regenera las variantes Outlook y SMTP, corre tools/validate.py en las dos, reempaqueta los 4 zips (managed y unmanaged), actualiza el README, haz commit y push a la rama y avísame.
```
