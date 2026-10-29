// OpenAI-compatible API of the Uni server (llm.scads.ai): model list and streamed answers with tool calls.

/** The recommended models as in the app (Einstellungen › Modelle); the rest comes from /models. */
export const FEATURED = [
  { id: "deepseek-ai/DeepSeek-V4.1-Flash", label: "DeepSeek V4.1 Flash", use: "Für fast alles (Standard)", vision: true },
  { id: "zai-org/GLM-5.3", label: "GLM-5.3", use: "Schwierigste Aufgaben, langsamer", vision: false },
  { id: "zai-org/GLM-5.3-Flash", label: "GLM-5.3 Flash", use: "Am schnellsten, sehr lange Dokumente", vision: true },
  { id: "Qwen/Qwen3.8-27B", label: "Qwen3.8 27B", use: "Programmieren", vision: true },
];

const NOT_CHAT = /embed|rerank|whisper|tts|speech|flux|diffusion|image|ocr|guard|alias|clip/i;

export async function models(s) {
  const res = await fetch(s.baseUrl.replace(/\/$/, "") + "/models", { headers: { Authorization: "Bearer " + s.apiKey } });
  if (!res.ok) throw await apiError(res);
  const data = await res.json();
  return (data.data || []).map((m) => m.id).filter((id) => !NOT_CHAT.test(id)).sort();
}

async function apiError(res) {
  let detail = "";
  try {
    const body = await res.text();
    try { detail = JSON.parse(body).error?.message || body; } catch { detail = body; }
  } catch { /* ignore */ }
  const msg = {
    401: "Der API-Schlüssel ist ungültig oder abgelaufen (Einstellungen › API-Schlüssel).",
    403: "Kein Zugriff auf dieses Modell mit diesem Schlüssel.",
    404: "Modell oder Adresse nicht gefunden – bitte ein anderes Modell wählen.",
    413: "Die Anfrage ist zu groß – weniger oder kürzere Dateien anhängen.",
    429: "Zu viele Anfragen – der Server bittet um eine kurze Pause.",
  }[res.status] || `Der Server antwortet mit Fehler ${res.status}.`;
  const e = new Error(detail ? `${msg} (${String(detail).slice(0, 300)})` : msg);
  e.status = res.status;
  e.retryAfter = Number(res.headers.get("retry-after")) || 0;
  return e;
}

const sleep = (ms, signal) => new Promise((resolve, reject) => {
  const t = setTimeout(resolve, ms);
  signal?.addEventListener("abort", () => { clearTimeout(t); reject(new DOMException("Abgebrochen", "AbortError")); }, { once: true });
});

/**
 * Sends [messages] and streams the answer. [onDelta] gets {content, reasoning} as they grow. Returns the full answer
 * with its tool calls. Retries 429 and gateway errors (502–504) a few times, like the app.
 */
export async function chat(s, { model, messages, tools, signal, onDelta }) {
  for (let attempt = 0; ; attempt++) {
    const res = await fetch(s.baseUrl.replace(/\/$/, "") + "/chat/completions", {
      method: "POST",
      headers: { Authorization: "Bearer " + s.apiKey, "Content-Type": "application/json" },
      body: JSON.stringify({ model, messages, stream: true, ...(tools?.length ? { tools } : {}) }),
      signal,
    });
    if (!res.ok) {
      const e = await apiError(res);
      if ((res.status === 429 || res.status >= 502) && attempt < 3) {
        onDelta?.({ content: "", reasoning: "", notice: `${e.message} Neuer Versuch …` });
        await sleep(Math.min(30, e.retryAfter || 3 * (attempt + 1)) * 1000, signal);
        continue;
      }
      throw e;
    }
    return read(res, onDelta);
  }
}

async function read(res, onDelta) {
  const reader = res.body.getReader();
  const decoder = new TextDecoder();
  let buffer = "";
  let content = "";
  let reasoning = "";
  let finish = null;
  const calls = [];
  for (;;) {
    const { done, value } = await reader.read();
    if (done) break;
    buffer += decoder.decode(value, { stream: true });
    let nl;
    while ((nl = buffer.indexOf("\n")) >= 0) {
      const line = buffer.slice(0, nl).trim();
      buffer = buffer.slice(nl + 1);
      if (!line.startsWith("data:")) continue;
      const data = line.slice(5).trim();
      if (data === "[DONE]") continue;
      let chunk;
      try { chunk = JSON.parse(data); } catch { continue; }
      if (chunk.error) throw new Error(chunk.error.message || "Fehler vom Server");
      const choice = chunk.choices?.[0];
      if (!choice) continue;
      const d = choice.delta || {};
      if (d.content) content += d.content;
      if (d.reasoning_content || d.reasoning) reasoning += d.reasoning_content || d.reasoning;
      for (const tc of d.tool_calls || []) {
        const i = tc.index ?? calls.length;
        calls[i] = calls[i] || { id: "", type: "function", function: { name: "", arguments: "" } };
        if (tc.id) calls[i].id = tc.id;
        if (tc.function?.name) calls[i].function.name += tc.function.name;
        if (tc.function?.arguments) calls[i].function.arguments += tc.function.arguments;
      }
      if (choice.finish_reason) finish = choice.finish_reason;
      onDelta?.({ content, reasoning });
    }
  }
  const toolCalls = calls.filter(Boolean).map((c, i) => ({ ...c, id: c.id || `call_${Date.now()}_${i}` }));
  return { content, reasoning, toolCalls, finish };
}
