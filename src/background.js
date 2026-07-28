// Service worker: abre el panel lateral / la ventana flotante y guarda cuál fue
// la última pestaña "real" que el usuario miró, para poder leer su contenido.

// El sufijo ?window permite a la UI saber que ya está en la ventana flotante.
const CHAT_URL = "src/chat.html?window";

let lastTabId = null;
let chatWindowId = null;

chrome.runtime.onInstalled.addListener(() => {
  // Clic en el icono de la barra -> abre el panel lateral.
  chrome.sidePanel
    .setPanelBehavior({ openPanelOnActionClick: true })
    .catch(() => {});
});

// Chrome solo deja inyectar scripts en páginas http(s).
function isReadable(url) {
  return /^https?:\/\//.test(url || "");
}

function rememberTab(tab) {
  if (tab?.id && isReadable(tab.url || tab.pendingUrl)) lastTabId = tab.id;
}

chrome.tabs.onActivated.addListener(async ({ tabId }) => {
  try {
    rememberTab(await chrome.tabs.get(tabId));
  } catch {
    /* la pestaña ya no existe */
  }
});

chrome.tabs.onUpdated.addListener((_tabId, changeInfo, tab) => {
  if (changeInfo.status === "complete" && tab.active) rememberTab(tab);
});

chrome.windows.onRemoved.addListener((windowId) => {
  if (windowId === chatWindowId) chatWindowId = null;
});

async function openChatWindow() {
  if (chatWindowId !== null) {
    try {
      await chrome.windows.update(chatWindowId, { focused: true });
      return;
    } catch {
      chatWindowId = null;
    }
  }
  const win = await chrome.windows.create({
    url: chrome.runtime.getURL(CHAT_URL),
    type: "popup",
    width: 460,
    height: 720,
  });
  chatWindowId = win.id ?? null;
}

chrome.commands.onCommand.addListener((command) => {
  if (command === "open-window") openChatWindow();
});

// Elige la pestaña a leer: la activa de la ventana normal enfocada; si no vale,
// la última que recordamos. Las ventanas de tipo popup (el propio chat) se ignoran.
async function pickTab() {
  const windows = await chrome.windows.getAll({ populate: true });
  const normal = windows
    .filter((w) => w.type === "normal")
    .sort((a, b) => Number(b.focused) - Number(a.focused));

  for (const win of normal) {
    const tab = win.tabs?.find((t) => t.active);
    if (tab && isReadable(tab.url)) return tab.id;
  }

  if (lastTabId !== null) {
    try {
      const tab = await chrome.tabs.get(lastTabId);
      if (isReadable(tab.url)) return tab.id;
    } catch {
      lastTabId = null;
    }
  }

  throw new Error(
    "No encuentro ninguna pestaña legible. Abre una página http/https " +
      "(Chrome no deja leer chrome://, la Web Store ni PDFs internos)."
  );
}

// Lee título, URL y texto visible de esa pestaña.
async function readPage() {
  const tabId = await pickTab();

  const [result] = await chrome.scripting.executeScript({
    target: { tabId },
    func: () => ({
      title: document.title,
      url: location.href,
      text: (document.body?.innerText || "").slice(0, 40000),
    }),
  });
  return result.result;
}

chrome.runtime.onMessage.addListener((msg, _sender, sendResponse) => {
  if (msg?.type === "open-window") {
    openChatWindow();
    return false;
  }
  if (msg?.type === "read-page") {
    readPage().then(
      (page) => sendResponse({ ok: true, page }),
      (err) => sendResponse({ ok: false, error: String(err.message || err) })
    );
    return true; // respuesta asíncrona
  }
  return false;
});
