# Guía de implementación en KNIME — Trigger NOVA vs SAP + aprobaciones en Outlook

## 1. Arquitectura

Son **dos workflows** de KNIME que comparten una base SQLite de estado:

```
WF-A  "NOVA_SAP_Trigger"  (programado cada 15 min en la ventana de operación, p. ej. 6:00–12:00)
 ┌────────────────────┐
 │ Variable Creator   │  lib_dir, config_path
 └─────────┬──────────┘
           │ (flow vars a todo)
 ┌─────────▼──────────┐      ┌───────────────────────────┐
 │ Python Script      │      │ Lectura SAP               │
 │ 01 Extraer NOVA    │      │ (SAP Reader / DB Query /  │
 └─────────┬──────────┘      │  Excel Reader de VA05)    │
           │                 └────────────┬──────────────┘
           └───────────┬──────────────────┘
             ┌─────────▼──────────┐
             │ Python Script      │──► resumen + detalle (Excel Writer para auditoría)
             │ 02 Match/Trigger   │
             └─────────┬──────────┘
                       │ flow var PortChoice = top / bottom
             ┌─────────▼──────────┐
             │ IF Switch          │
             └───┬────────────┬───┘
            top  │            │ bottom
   ┌─────────────▼───┐   ┌────▼───────────────────────┐
   │ Reportes        │   │ (opcional) aviso de        │
   │ (sus nodos)     │   │ diferencias / login NOVA   │
   └────────┬────────┘   └────────────────────────────┘
   ┌────────▼────────────────────┐
   │ Python Script               │  1 renglón por correo a enviar
   │ 03 Crear solicitudes        │──► manda a cada OMA el correo de aprobación
   └─────────────────────────────┘

WF-B  "NOVA_SAP_Aprobaciones"  (programado cada 5–10 min)
 Variable Creator ─► Python Script 04 Procesar respuestas ─► log / estado (Excel o DB)
```

¿Por qué dos workflows? Porque la aprobación del OMA puede tardar minutos u horas; un
workflow no debe quedarse bloqueado esperando. WF-B revisa el buzón periódicamente.

## 2. Regla del trigger (nodo 02)

- Se toman las POs **únicas** de NOVA y de SAP (duplicadas cuentan 1 vez).
- Se normalizan antes de comparar: espacios, `1234567890.0` de Excel, notación
  científica, ceros a la izquierda de SAP (configurable).
- **Se dispara** si `#únicas NOVA == #únicas SAP` **y** ninguna PO de NOVA falta en SAP.
- **No se dispara** si NOVA viene vacío (casi siempre es falla de extracción o sesión
  caducada), ni más de una vez por día (`fire_once_per_day`).
- Cada evaluación queda registrada en la tabla `trigger_log` del SQLite.

> ⚠️ **Alcance de SAP:** la consulta de SAP debe traer exactamente el mismo universo que
> muestra NOVA (mismo rango de fechas, mismo cliente Walmart / sold-to, mismos DCs y
> sin pedidos cancelados). Si SAP trae histórico, el conteo nunca va a cuadrar.

## 3. Requisitos de la máquina del robot

- KNIME Analytics Platform 5.x (o executor de KNIME Business Hub) con
  **KNIME Python Integration**. Recomendado: un entorno conda/pip propio
  (Preferences → KNIME → Python → "Conda" o "Manual") con `requirements.txt`:
  ```bash
  pip install -r requirements.txt
  ```
  ```bash
  python -m playwright install chromium
  ```
  (o usar `"browser_channel": "msedge"` para el Edge instalado en Windows).
- Por el SSO + MFA de Retail Link, la extracción de NOVA debe correr en una **VM / PC
  Windows fija** con un usuario de servicio que tenga acceso a NOVA. Un executor
  efímero en la nube no conserva la sesión.

### Login inicial de NOVA (una vez, y cada que caduque la sesión)
```bash
python -m nova_sap.nova_extract --config C:/KNIME_Robot/nova_sap/config.json --setup
```
(Ejecutar desde `knime/lib`.) Se abre el navegador; iniciar sesión con MFA, entrar a la
lista de órdenes y presionar ENTER. Si en una corrida la sesión caducó, el nodo 01 deja
`nova_error = LOGIN_REQUERIDO...` y el trigger **no** se dispara.

### Ajustar selectores de NOVA
Los selectores de `config.example.json` son de ejemplo. Para obtener los reales:
```bash
python -m playwright codegen https://retaillink2.wal-mart.com/novarl/
```
Grabar los clics (filtros, Export) y copiar los selectores a `nova.pre_steps` y
`nova.selectors`. Probar sin KNIME:
```bash
python -m nova_sap.nova_extract --config C:/KNIME_Robot/nova_sap/config.json --headed --out nova.csv
```
Si NOVA tiene botón de exportar, usar `"mode": "export"` (más estable que leer la tabla).

## 4. Armado del WF-A en KNIME, nodo por nodo

1. **Variable Creator**: `lib_dir` = ruta a `knime/lib` de este repo,
   `config_path` = ruta al `config.json`. Conectar su puerto de variables al nodo 01
   y a la lectura de SAP.
2. **Python Script (0 in / 1 out)** → pegar `knime/nodes/01_nova_extract.py`.
3. **Lectura SAP** — elegir según lo que tengan:
   - *SAP Reader (Theobald Software)* (requiere licencia de Xtract Universal): tabla /
     query sobre `VBAK`+`VBKD` o un reporte/BW query.
   - *DB Connector / DB Query Reader* si tienen réplica en HANA / BW / data lake.
   - *Excel Reader* sobre un export de `VA05` / ALV programado (job SAP que deja el
     archivo en una carpeta compartida).
   La PO de Walmart en un pedido de venta SAP es la **referencia de cliente**
   (`VBKD-BSTKD`, también en `VBAK-BSTNK`). Poner el nombre de esa columna en
   `config.sap.po_column`. Leerla como **String**.
4. **Python Script (2 in / 2 out)** → pegar `knime/nodes/02_match_trigger.py`.
   Entrada 0 = NOVA, entrada 1 = SAP. Salidas: resumen y detalle.
   Recomendado: *Excel Writer* del detalle a una carpeta de auditoría.
5. **IF Switch**: en *Flow Variables* del nodo, asignar `PortChoice` ← flow var
   `PortChoice` (lo genera el nodo 02). `top` = trigger disparado.
6. Rama **top**: sus nodos de reportes. Al final deben producir **una tabla con un
   renglón por correo**:

   | columna | ejemplo |
   |---|---|
   | `report_name` | Fill Rate |
   | `oma_email` | oma.cuenta@empresa.com |
   | `to` | comprador@walmart.com;otro@walmart.com |
   | `cc` | kam@empresa.com |
   | `subject` | Fill rate semana 45 |
   | `body_html` | `<p>Hola ...</p>` (puede salir de un *String Manipulation* o de una plantilla) |
   | `attachment_paths` | C:/Reportes/fill_rate_s45.xlsx |

   Termina con **End IF** si se quiere unir con la rama bottom.
7. **Python Script (1 in / 1 out)** → pegar `knime/nodes/03_create_approval_requests.py`.
8. Rama **bottom** (opcional): *Rule Engine Variable* / *Send Email* para avisar al
   equipo si `nova_error` trae `LOGIN_REQUERIDO` o si después de cierta hora todavía
   hay `SOLO_NOVA` (POs que no han bajado a SAP).

## 5. WF-B (aprobaciones)

1. **Variable Creator** (`lib_dir`, `config_path`).
2. **Python Script (0 in / 2 out)** → pegar `knime/nodes/04_process_approvals.py`.
3. *Excel Writer* / *DB Writer* con el log y el estado (tablero para el equipo).

### Qué ve el OMA

Recibe `"[APROBACIÓN REQ-1A2B3C4D] Fill rate semana 45"` con:
- resumen (para, CC, asunto, adjuntos) y vista previa del cuerpo,
- botones **APROBAR** / **RECHAZAR** (abren una respuesta ya escrita, sólo dar Enviar),
- adjunto **propuesta.eml** = el correo final, que en Outlook abre como **borrador
  editable** con destinatarios y adjuntos.

| El OMA hace… | Resultado |
|---|---|
| Clic en APROBAR → Enviar | KNIME envía el correo original (`APPROVED_SENT`) |
| Abre `propuesta.eml`, edita, y da **Enviar** él mismo | Sale desde su buzón; la copia BCC al buzón de aprobaciones cierra la solicitud (`EDITED_SENT`) |
| Edita el borrador, lo **arrastra/adjunta** a su respuesta "APROBAR" | KNIME envía la versión editada (`EDITED_SENT`) |
| Clic en RECHAZAR (+ motivo) | No se envía nada, se guarda el motivo (`REJECTED`) |
| No responde en `expires_hours` | `EXPIRED`, se le avisa, no se envía nada |

Controles: sólo cuenta la respuesta del OMA asignado (o sus `delegates`); una
solicitud se cierra una sola vez (doble clic en APROBAR no duplica el envío); si Graph
falla al enviar queda en `ERROR` para revisión manual.

## 6. Configuración de Outlook / Microsoft 365 (TI)

1. **Entra ID → App registrations → New**. Guardar *tenant id* y *client id*;
   crear un *client secret*.
2. **API permissions → Microsoft Graph → Application**: `Mail.Send` y
   `Mail.ReadWrite` → *Grant admin consent*.
3. Crear dos buzones compartidos (o usar uno para ambos):
   `sender_mailbox` (de donde salen los correos a Walmart) y `approvals_mailbox`
   (donde llegan las respuestas de los OMAs).
4. **Restringir la app a esos buzones** (Exchange Online PowerShell):
   ```powershell
   New-DistributionGroup -Name "KNIME-NovaSap-Mailboxes" -Type Security -Members reportes.walmart@tuempresa.com,aprobaciones.walmart@tuempresa.com
   ```
   ```powershell
   New-ApplicationAccessPolicy -AppId <client-id> -PolicyScopeGroupId KNIME-NovaSap-Mailboxes -AccessRight RestrictAccess -Description "KNIME NOVA SAP"
   ```
5. En la máquina del robot, variables de entorno del usuario que corre KNIME:
   `NOVA_SAP_TENANT_ID`, `NOVA_SAP_CLIENT_ID`, `NOVA_SAP_CLIENT_SECRET`.

> Límite: el envío por MIME de Graph acepta ~4 MB por correo. Si los reportes adjuntos
> pesan más, conviene mandar liga a SharePoint/OneDrive en lugar del archivo.

## 7. Programación

- **KNIME Business Hub / Server**: crear *schedules* para WF-A (cada 15 min,
  ventana de la mañana) y WF-B (cada 5–10 min, horario laboral).
- **Sólo Analytics Platform**: Programador de tareas de Windows con el modo batch:
  ```bash
  "C:\Program Files\KNIME\knime.exe" -nosplash -consoleLog -reset -application org.knime.product.KNIME_BATCH_APPLICATION -workflowDir="C:\KNIME_Robot\workflows\NOVA_SAP_Trigger"
  ```

## 8. Pendientes por confirmar con el negocio

- Pantalla exacta y filtros de NOVA que definen "la lista de órdenes del día"
  (fecha de creación, MABD, DC, estatus).
- Fuente SAP disponible (Theobald, HANA, export programado) y filtros equivalentes.
- Hora límite: si a cierta hora no hay match, ¿se corren los reportes de todos modos
  o se escala?
- Lista OMA ↔ cuenta/reporte y sus delegados.
- Validar en piloto que Outlook conserve el BCC al abrir `propuesta.eml`; si no, la
  opción "adjuntar la versión editada a la respuesta APROBAR" siempre funciona.
