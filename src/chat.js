// UI del chat. Habla directamente con la API de Anthropic desde el navegador,
// con la clave que el usuario guarda en chrome.storage.local.

const API_URL = "https://api.anthropic.com/v1/messages";
const DEFAULT_MODEL = "claude-opus-5";
const MAX_TOKENS = 8000;
const SYSTEM_PROMPT =
  "Eres un asistente que vive en una extensión de Chrome. Responde en el idioma del " +
  "usuario, de forma directa y sin preámbulos. Si te dan el contenido de una página, " +
  "úsalo como contexto y cita la parte relevante cuando ayude.";

const el = {
  log: document.getElementById("log"),
  input: document.getElementById("input"),
  send: document.getElementById("send"),
  clear: document.getElementById("clear"),
  popout: document.getElementById("popout"),
  toggleSettings: document.getElementById("toggle-settings"),
  settings: document.getElementById("settings"),
  apiKey: document.getElementById("api-key"),
  model: document.getElementById("model"),
  saveSettings: document.getElementById("save-settings"),
  usePage: document.getElementById("use-page"),
};

/** @type {{role: 'user'|'assistant', content: string}[]} */
let history = [];
let settings = { apiKey: "", model: DEFAULT_MODEL };
let busy = false;

init();

async function init() {
  const stored = await chrome.storage.local.get(["apiKey", "model", "history"]);
  settings.apiKey = stored.apiKey || "";
  settings.model = stored.model || DEFAULT_MODEL;
  el.apiKey.value = settings.apiKey;
  el.model.value = settings.model;

  history = Array.isArray(stored.history) ? stored.history : [];
  for (const m of history) addBubble(m.role, m.content);

  if (!settings.apiKey) {
    el.settings.hidden = false;
    addBubble("note", "Pega tu clave de API de Anthropic para empezar.");
  }

  // En la ventana flotante no tiene sentido ofrecer «abrir en ventana».
  if (location.search.includes("window")) el.popout.hidden = true;

  wireEvents();
  el.input.focus();
}

function wireEvents() {
  el.send.addEventListener("click", () => submit());

  el.input.addEventListener("keydown", (e) => {
    if (e.key === "Enter" && !e.shiftKey) {
      e.preventDefault();
      submit();
    }
  });

  el.input.addEventListener("input", () => {
    el.input.style.height = "auto";
    el.input.style.height = Math.min(el.input.scrollHeight, 180) + "px";
  });

  el.toggleSettings.addEventListener("click", () => {
    el.settings.hidden = !el.settings.hidden;
  });

  el.saveSettings.addEventListener("click", async () => {
    settings.apiKey = el.apiKey.value.trim();
    settings.model = el.model.value;
    await chrome.storage.local.set(settings);
    el.settings.hidden = true;
    addBubble("note", "Ajustes guardados.");
  });

  el.clear.addEventListener("click", async () => {
    history = [];
    el.log.replaceChildren();
    await chrome.storage.local.remove("history");
  });

  el.popout.addEventListener("click", () => {
    chrome.runtime.sendMessage({ type: "open-window" });
  });

  // Leer páginas necesita permiso de host. Se pide solo cuando el usuario lo activa,
  // aprovechando que el clic en la casilla es un gesto del usuario.
  el.usePage.addEventListener("change", async () => {
    if (!el.usePage.checked) return;
    const granted = await chrome.permissions.request({
      origins: ["http://*/*", "https://*/*"],
    });
    if (!granted) {
      el.usePage.checked = false;
      addBubble("note", "Sin permiso de lectura no puedo ver el contenido de las pestañas.");
    }
  });
}

function addBubble(role, text) {
  const node = document.createElement("div");
  node.className = "msg " + role;
  node.textContent = text;
  el.log.append(node);
  el.log.scrollTop = el.log.scrollHeight;
  return node;
}

async function submit() {
  if (busy) return;
  const text = el.input.value.trim();
  if (!text) return;

  if (!settings.apiKey) {
    el.settings.hidden = false;
    addBubble("error", "Falta la clave de API.");
    return;
  }

  let prompt = text;
  if (el.usePage.checked) {
    try {
      const page = await readCurrentPage();
      prompt =
        `Contexto de la pestaña abierta:\n` +
        `Título: ${page.title}\nURL: ${page.url}\n\n` +
        `<contenido>\n${page.text}\n</contenido>\n\n` +
        text;
    } catch (err) {
      addBubble("error", `No pude leer la página: ${err.message}`);
      return;
    }
  }

  el.input.value = "";
  el.input.style.height = "auto";
  addBubble("user", text);
  history.push({ role: "user", content: prompt });

  setBusy(true);
  const bubble = addBubble("assistant", "");
  try {
    await streamReply(bubble);
    history.push({ role: "assistant", content: bubble.textContent });
    await chrome.storage.local.set({ history: history.slice(-40) });
  } catch (err) {
    if (!bubble.textContent) bubble.remove();
    history.pop(); // descarta el turno del usuario que no llegó a respuesta
    addBubble("error", err.message);
  } finally {
    setBusy(false);
    el.input.focus();
  }
}

function setBusy(value) {
  busy = value;
  el.send.disabled = value;
  el.send.textContent = value ? "…" : "Enviar";
}

async function readCurrentPage() {
  const res = await chrome.runtime.sendMessage({ type: "read-page" });
  if (!res?.ok) throw new Error(res?.error || "error desconocido");
  return res.page;
}

async function streamReply(bubble) {
  const res = await fetch(API_URL, {
    method: "POST",
    headers: {
      "content-type": "application/json",
      "x-api-key": settings.apiKey,
      "anthropic-version": "2023-06-01",
      // Necesario para llamar a la API desde un contexto de navegador.
      "anthropic-dangerous-direct-browser-access": "true",
    },
    body: JSON.stringify({
      model: settings.model,
      max_tokens: MAX_TOKENS,
      stream: true,
      system: SYSTEM_PROMPT,
      messages: history,
    }),
  });

  if (!res.ok) {
    let detail = `HTTP ${res.status}`;
    try {
      const body = await res.json();
      detail = body?.error?.message || detail;
    } catch {
      /* respuesta no JSON */
    }
    throw new Error(detail);
  }

  const reader = res.body.getReader();
  const decoder = new TextDecoder();
  let buffer = "";

  while (true) {
    const { done, value } = await reader.read();
    if (done) break;
    buffer += decoder.decode(value, { stream: true });

    const lines = buffer.split("\n");
    buffer = lines.pop() ?? ""; // la última línea puede estar incompleta

    for (const line of lines) {
      if (!line.startsWith("data:")) continue;
      const payload = line.slice(5).trim();
      if (!payload) continue;

      let event;
      try {
        event = JSON.parse(payload);
      } catch {
        continue;
      }

      if (event.type === "content_block_delta" && event.delta?.type === "text_delta") {
        bubble.textContent += event.delta.text;
        el.log.scrollTop = el.log.scrollHeight;
      } else if (event.type === "message_delta" && event.delta?.stop_reason === "refusal") {
        throw new Error("El modelo declinó responder a esta petición.");
      } else if (event.type === "error") {
        throw new Error(event.error?.message || "Error del servidor.");
      }
    }
  }

  if (!bubble.textContent) throw new Error("La respuesta llegó vacía.");
}
