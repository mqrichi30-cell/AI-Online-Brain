# AI Online Brain

Extensión de Chrome (Manifest V3) para chatear con Claude en un **panel lateral** o en una
**ventana flotante**, con la opción de pasarle el contenido de la pestaña que estés viendo.

No necesita instalar nada de escritorio: es la propia extensión la que llama a la API de
Anthropic desde el navegador con tu clave.

## Instalar

1. Descarga este repositorio en tu equipo (botón **Code → Download ZIP** en GitHub, y
   descomprime; o `git clone` si tienes git).
2. Abre `chrome://extensions` en Chrome.
3. Activa **Modo de desarrollador** (arriba a la derecha).
4. Pulsa **Cargar descomprimida** y selecciona la carpeta que contiene `manifest.json`.
5. Ancla el icono a la barra de herramientas (icono de puzzle → chincheta).

## Usar

- **Clic en el icono** → abre el chat en el panel lateral.
- **Ctrl+Shift+Y** (o **⌘+Shift+Y** en Mac) → abre el chat en una ventana flotante
  independiente. También lo hace el botón `⧉` de la cabecera.
- La primera vez, pulsa `⚙` y pega tu clave de API de Anthropic
  ([consola de Anthropic](https://console.anthropic.com/settings/keys)).
- Marca **«Usar el contenido de la pestaña actual»** para que el mensaje incluya el título,
  la URL y el texto visible de la página (hasta 40 000 caracteres). La primera vez Chrome
  pedirá permiso para leer páginas; si lo rechazas, la casilla se desmarca sola.
- `✕` borra la conversación.

## Cómo está montado

| Archivo | Qué hace |
| --- | --- |
| `manifest.json` | Permisos, panel lateral, atajo de teclado |
| `src/background.js` | Service worker: abre la ventana, recuerda la última pestaña, lee su contenido |
| `src/chat.html` / `chat.css` | Interfaz (la misma para panel y ventana) |
| `src/chat.js` | Estado, almacenamiento y streaming contra `api.anthropic.com` |

Modelo por defecto: `claude-opus-5`. Puedes cambiarlo a Sonnet 5 o Haiku 4.5 en `⚙`.

## Sobre la clave de API

La clave se guarda en `chrome.storage.local` (solo en tu navegador, sin sincronizar) y se
envía únicamente a `api.anthropic.com`. La llamada usa la cabecera
`anthropic-dangerous-direct-browser-access`, necesaria para hablar con la API desde un
contexto de navegador: cualquiera con acceso a tu perfil de Chrome podría leer la clave, así
que usa una clave dedicada y revócala si compartes el equipo.

## Limitaciones conocidas

- Leer la pestaña usa un permiso de host **opcional** (`http://*/*`, `https://*/*`): se pide
  al marcar la casilla, no durante la instalación, y puedes revocarlo en `chrome://extensions`.
- Chrome no permite leer páginas `chrome://`, la Chrome Web Store ni PDFs internos.
- El historial se guarda en local, recortado a los últimos 40 mensajes.
