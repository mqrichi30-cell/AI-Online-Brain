# Wakefern 11s block — Autopush (Wholesale 2 Alerts)

Reparación y ejecución de la automatización
*"Wakefern - 11s block - Autopush WHOLESALE 2 ALERTS"*, que no corrió con el
reporte del **03-08-2026 07:03**.

## Resumen de una línea

El correo estaba bien; el ZIP del flow venía **incompleto** (sin el Office
Script) y con un **bug de URI en las 4 acciones que envían los correos**. Aquí
está el script reconstruido, el flow corregido, y la corrida de hoy ya
ejecutada.

## Qué hay en este repo

| Ruta | Qué es |
|---|---|
| **[`chrome/INSTRUCCIONES-CLAUDE-CHROME.md`](chrome/INSTRUCCIONES-CLAUDE-CHROME.md)** | ⭐ **Empiece aquí.** Prompts listos para pegar en Claude para Chrome y que ejecute todo el trabajo en el navegador |
| [`docs/DIAGNOSTICO.md`](docs/DIAGNOSTICO.md) | Por qué falló, defecto por defecto, con la evidencia |
| [`office-script/ProcessWakefern11s_AUTOPUSH.ts`](office-script/ProcessWakefern11s_AUTOPUSH.ts) | El Office Script que **no venía en el ZIP**, reconstruido |
| [`flow/definition.json`](flow/definition.json) | Flow corregido (V15) |
| [`tools/build_flow_zip.py`](tools/build_flow_zip.py) | Arma el `.zip` importable a Power Automate |
| [`tools/run_local.py`](tools/run_local.py) | Ejecuta la misma lógica sobre un `.msg`/`.xlsx`, sin Power Automate |
| [`output/`](output/) | La corrida real del 03-08-2026 |

## Resultado de la corrida del 03-08-2026

**7 órdenes** WAKEFERN con bloqueo `11` y Good Issue hoy o mañana:

| SO # | PO # | Good Issue |
|---|---|---|
| 2066923910 | 01737647 | 2026-08-03 |
| 2066923927 | 01737651 | 2026-08-03 |
| 2066923928 | 01737653 | 2026-08-03 |
| 2066923929 | 01737648 | 2026-08-03 |
| 2066923911 | 01737654 | 2026-08-04 |
| 2066923925 | 01737646 | 2026-08-04 |
| 2066923926 | 01737652 | 2026-08-04 |

- **Push GI a:** `2026-08-05`
- **Nuevo RDD:** `2026-08-06` (*Thursday, August 6, 2026*)

Los dos correos ya redactados están en [`output/emails.json`](output/emails.json)
y se ven renderizados en [`output/preview.html`](output/preview.html).

## Uso rápido

```bash
pip install olefile openpyxl

# Ejecutar la automatización sobre el correo del día
python3 tools/run_local.py "Wholesale_2_Alerts_Report__08032026_Time_07_03.msg" \
    --today 2026-08-03

# Regenerar el paquete importable de Power Automate
python3 tools/build_flow_zip.py
```

## ⚠️ Dos cosas que debe confirmar antes de enviar

1. **La regla de empuje.** El script V14 original no venía en el ZIP, así que la
   regla exacta no se puede recuperar. Reconstruí la más común y la dejé como
   constantes al inicio del `.ts`:

   ```ts
   const PUSH_GI_BUSINESS_DAYS_AHEAD = 2;   // nuevo GI  = hoy + 2 días hábiles
   const RDD_OFFSET_BUSINESS_DAYS    = 1;   // nuevo RDD = nuevo GI + 1 día hábil
   ```

   Si su regla real es otra, cambie esos dos números.

2. **Los destinatarios son externos reales** (Wakefern). Tanto el flow V15 como
   las instrucciones de Chrome conservan los destinatarios de producción del
   V14. Para probar, apunte primero a su propio correo.

## Importar el flow corregido

1. `python3 tools/build_flow_zip.py`
2. En `make.powerautomate.com` → **Mis flujos** → **Importar** → *Paquete de
   importación (heredado)* → subir el `.zip` de `dist/`.
3. Asignar las 3 conexiones (Office 365 Outlook, OneDrive, Excel Online).
4. **Reabrir la acción `Run script ProcessWakefern11s` y volver a seleccionar el
   script desde el desplegable.** El `scriptId` del paquete apunta al OneDrive
   del autor original y no se puede transferir — hay que reapuntarlo a mano.
5. Probar con destinatarios propios antes de activar.
