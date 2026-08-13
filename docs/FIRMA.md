# Firma corporativa — fuente unica

Esta es la firma que debe salir en **los nueve** correos que envia la solucion.
Si cambia algo aqui, cambia en los nueve sitios. No editar una copia suelta.

Sitios que la usan:

| Flujo | Accion | Mecanismo de envio |
|---|---|---|
| 00 | `GENERAL - 00 Rule Reply Comment` (rama consolidacion) | Graph `/reply` → pasa a borrador |
| 00 | `GENERAL - 00 Rule Reply Comment` (rama clasificador) | Graph `/reply` → pasa a borrador |
| 01 | `AWG - CON - 01 HTTP Reply Current OSS to OSSGenAI SameConversation` | Graph `/reply` → pasa a borrador |
| 02 | `AWG - CON - 02 Compose Reply HTML` | borrador Graph (ya correcto) |
| 03 | `AWG - CON - 03 HTTP Send Consolidation Request` | borrador Graph (ya correcto) |
| 03 | `AWG - CON - 03 Compose RDD Reply HTML` | Graph `/reply` → pasa a borrador |
| 04 | `AWG - CON - 04 Send HTTP OSS Response` | Graph `/reply` → pasa a borrador |
| 04 | `AWG - CON - 04 Send HTTP OSS Response 1` | Graph `/reply` → pasa a borrador |
| 05 | `Send an email (V2)` y `Send an email (V2) 1` | conector → pasa a Graph |

---

## El logo

Viaja **dentro** del correo como adjunto en linea, no como enlace externo. Asi se
ve siempre, sin que el destinatario tenga que autorizar la descarga de imagenes.

En cada envio, despues de crear el borrador y antes de enviarlo, se adjunta:

```json
{
  "@odata.type": "#microsoft.graph.fileAttachment",
  "name": "pglogo.png",
  "contentType": "image/png",
  "isInline": true,
  "contentId": "pglogo",
  "contentBytes": "<base64 del PNG>"
}
```

El base64 vive en una sola accion `Compose` por flujo, llamada
`Compose_Firma_Logo_Base64`, para no repetirlo dentro de cada cuerpo.

---

## Version A · atributos con comillas dobles

Para los cuerpos HTML que **no** se incrustan dentro de un JSON construido con
`concat` (flujos 02, 03, 04, 05).

```html
<p>Regards,</p>
<table cellpadding="0" cellspacing="0" border="0" style="border-collapse:collapse;font-family:'Segoe UI',Arial,sans-serif;font-size:10.5pt;color:#000000;">
  <tr>
    <td style="padding:0 14px 0 0;vertical-align:middle;">
      <img src="cid:pglogo" alt="P&amp;G" width="110" style="display:block;border:0;outline:none;">
    </td>
    <td style="vertical-align:middle;line-height:1.4;">
      Cristhofer Mar&iacute;n Quir&oacute;s<br>
      T# FA2978<br>
      pgcustservw2.im@pg.com<br>
      F&amp;A | NA Order Management Analyst | Regional<br>
      Sales to Cash
    </td>
  </tr>
</table>
```

---

## Version B · atributos con comillas simples

Identica a la A, pero los atributos usan comilla simple en vez de doble. Es
obligatoria en el flujo 00 y en el 01, donde el HTML se mete como valor de una
cadena JSON que se arma con `concat`: una comilla doble ahi rompe el JSON.

Dentro de la expresion de Power Automate cada comilla simple va **doblada**.

```html
<p>Regards,</p>
<table cellpadding='0' cellspacing='0' border='0' style='border-collapse:collapse;font-family:Segoe UI,Arial,sans-serif;font-size:10.5pt;color:#000000;'>
  <tr>
    <td style='padding:0 14px 0 0;vertical-align:middle;'>
      <img src='cid:pglogo' alt='P&amp;G' width='110' style='display:block;border:0;outline:none;'>
    </td>
    <td style='vertical-align:middle;line-height:1.4;'>
      Cristhofer Mar&iacute;n Quir&oacute;s<br>
      T# FA2978<br>
      pgcustservw2.im@pg.com<br>
      F&amp;A | NA Order Management Analyst | Regional<br>
      Sales to Cash
    </td>
  </tr>
</table>
```

---

## Notas de compatibilidad

- Tabla y no `flex`/`grid`: Outlook de escritorio usa el motor de Word y no
  entiende maquetado moderno. La tabla de dos celdas es lo unico que respeta.
- Los estilos van en linea. Outlook descarta la mayoria de las hojas de estilo.
- `&` va escrito `&amp;` (`F&amp;A`, `alt='P&amp;G'`).
- Los acentos van como entidad (`&iacute;`, `&oacute;`) para no depender de que
  la codificacion sobreviva al viaje por Graph.
- `width` fijo y sin `height`: la altura la calcula el cliente y no se deforma.
- Lo unico que es imagen es el logo. El nombre, el T#, el correo y los cargos
  son texto real: se pueden seleccionar y copiar, y se leen aunque el
  destinatario tenga las imagenes bloqueadas.

---

## Pendiente antes de aplicar

Solo falta el PNG del logo, para convertirlo a base64.
