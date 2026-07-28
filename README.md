# AI-Online-Brain

Automatizaciones del equipo NA Order Management | Regional.

## AWG Ship With

Cadena de dos flows de Power Automate + un Office Script que detecta POs
subdimensionados en el reporte semanal de AWG y pide/manda el "Ship With" al
contacto del Ship-to.

| Ruta | Qué es |
|---|---|
| `office-scripts/shipWithReport.ts` | El Office Script que arma los borradores y resuelve el contacto |
| `office-scripts/tests/run.sh` | Tests del matching de contactos (`node` + `tsc`) |
| `flows/awg-shipwith-pos-report/` | `AWG - Ship with POs Repot` — procesa el adjunto y arma los borradores |
| `flows/awg-shipwith-add-contact-send/` | `AWG ShipWith - Add Contact & Send` — resuelve los Pending y manda el correo |
| `docs/awg-shipwith-soldto-contact-bug.md` | Diagnóstico del contacto que se volvía a pedir cada semana, y cómo desplegar |

Cada carpeta de flow tiene el `definition.original.json` exportado de Power
Automate y el `definition.patched.json` con las correcciones.
