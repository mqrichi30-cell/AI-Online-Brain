# AI-Online-Brain

## Automatización NOVA (Retail Link) vs SAP → trigger de reportes + aprobaciones en Outlook

Workflows de **KNIME** que:

1. Extraen la lista de órdenes (POs) de NOVA en `https://retaillink2.wal-mart.com/novarl/`.
2. La comparan contra las órdenes de SAP. **Banderazo** cuando el número de POs únicas
   coincide **y** todas las POs de NOVA existen en SAP (duplicadas cuentan 1 vez).
3. Con el banderazo corren los reportes y cada correo resultante se manda al **OMA**
   como solicitud de aprobación en Outlook (con el correo final adjunto, editable).
   El OMA puede **aprobar**, **editar** o **rechazar**; sólo lo aprobado se envía.

📘 Guía completa (arquitectura, armado nodo por nodo, Outlook/Azure, programación):
[`docs/guia_knime.md`](docs/guia_knime.md)

### Estructura

```
knime/
  config/config.example.json      # copiar a config.json y ajustar (no subir a git)
  lib/nova_sap/                   # lógica en Python que usan los nodos
    po_matching.py                #   normalización de POs + regla del trigger
    nova_extract.py               #   extracción NOVA con Playwright (sesión persistente SSO/MFA)
    approvals.py                  #   correos de aprobación, parseo de respuestas, envío
    graph_client.py               #   Microsoft Graph (Outlook / Exchange Online)
    state.py                      #   SQLite: log del trigger y solicitudes
  nodes/                          # código para pegar en cada nodo "Python Script" de KNIME
    01_nova_extract.py
    02_match_trigger.py
    03_create_approval_requests.py
    04_process_approvals.py
tests/                            # pytest (incluye prueba con navegador contra un NOVA simulado)
```

### Pruebas

```bash
pip install -r requirements.txt pytest
```
```bash
python -m pytest -q
```
