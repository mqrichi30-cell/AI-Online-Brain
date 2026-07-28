# AI-Online-Brain

Automatizaciones del equipo NA Order Management | Regional.

## AWG Ship With

Cadena de dos flows de Power Automate + un Office Script que detecta POs
subdimensionados en el reporte semanal de AWG y pide/manda el "Ship With" al
contacto del Ship-to.

| Ruta | Qué es |
|---|---|
| `flows/awg-shipwith-pos-report/definition.original.json` | `AWG - Ship with POs Repot` — procesa el adjunto y arma los borradores |
| `flows/awg-shipwith-add-contact-send/definition.original.json` | `AWG ShipWith - Add Contact & Send` — resuelve los Pending y manda el correo |
| `flows/awg-shipwith-add-contact-send/definition.patched.json` | El mismo flow con el guard anti-duplicados en la Contacts Data Base |
| `docs/awg-shipwith-soldto-contact-bug.md` | Diagnóstico del bug de Sold To × Ship To y pasos de corrección |
