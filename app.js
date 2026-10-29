// ADEU LLM im Browser: chats with the Uni server, projects with instructions and a working folder on this computer,
// attachments and an agent that reads and writes files in that folder (Word, Excel, Markdown, text …).
/* global marked, DOMPurify */
import * as store from "./js/store.js";
import * as folders from "./js/folders.js";
import * as llm from "./js/llm.js";
import { readAttachment, markdownToDocx, markdownToXlsx } from "./js/files.js";

const $ = (id) => document.getElementById(id);
const MAX_ROUNDS = 15;
const DEFAULT_KEY = "default";

const state = {
  settings: store.loadSettings(),
  chats: [],
  projects: [],
  /** Folder handles in memory (project id or "default"), so permission can be asked right in the click. */
  handles: {},
  current: null,
  activeProject: null,
  pending: [],
  abort: null,
  models: [],
};

marked.setOptions({ gfm: true, breaks: false });
DOMPurify.addHook("afterSanitizeAttributes", (node) => {
  if (node.tagName === "A") { node.setAttribute("target", "_blank"); node.setAttribute("rel", "noopener noreferrer"); }
});
const md = (text) => DOMPurify.sanitize(marked.parse(text || ""));
const esc = (s) => String(s ?? "").replace(/[&<>"']/g, (c) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[c]));

// ------------------------------------------------------------------ start

async function start() {
  state.chats = (await store.all("chats")).sort((a, b) => b.updatedAt - a.updatedAt);
  state.projects = (await store.all("projects")).sort((a, b) => a.name.localeCompare(b.name, "de"));
  for (const key of [DEFAULT_KEY, ...state.projects.map((p) => p.id)]) {
    const h = await folders.load(key).catch(() => null);
    if (h) state.handles[key] = h;
  }
  bind();
  renderModels();
  renderSidebar();
  openChat(null);
  if (!state.settings.apiKey) {
    openSettings();
    banner("Willkommen! Bitte zuerst den API-Schlüssel aus dem Self-Service der TU Dresden eintragen.");
  } else {
    loadModels();
  }
}

async function loadModels() {
  try {
    state.models = await llm.models(state.settings);
    renderModels();
  } catch (e) {
    banner("Modellliste nicht geladen: " + e.message);
  }
}

// ------------------------------------------------------------------ rendering

let pendingFrame = 0;

function renderModels() {
  const sel = $("model");
  const chosen = state.current?.model || state.settings.model;
  const featured = llm.FEATURED.map((m) => `<option value="${esc(m.id)}" title="${esc(m.use)}">${esc(m.label)}</option>`).join("");
  const others = state.models.filter((id) => !llm.FEATURED.some((f) => f.id === id));
  sel.innerHTML = `<optgroup label="Empfohlen">${featured}</optgroup>` +
    (others.length ? `<optgroup label="Weitere Modelle">${others.map((id) => `<option value="${esc(id)}">${esc(id.split("/").pop())}</option>`).join("")}</optgroup>` : "");
  if (![...sel.options].some((o) => o.value === chosen)) sel.insertAdjacentHTML("beforeend", `<option value="${esc(chosen)}">${esc(chosen.split("/").pop())}</option>`);
  sel.value = chosen;
}

function renderSidebar() {
  $("projects").innerHTML = state.projects.map((p) => `
    <li data-project="${p.id}" class="${state.activeProject === p.id ? "active" : ""}" title="${esc(p.instructions || p.name)}">
      <span>📁</span><span class="name">${esc(p.name)}</span>
      <button class="icon edit" data-edit="${p.id}" title="Projekt bearbeiten" aria-label="Projekt bearbeiten">✎</button>
    </li>`).join("") || `<li class="hint small" style="cursor:default">Noch keine Projekte.</li>`;
  const shown = state.chats.filter((c) => !state.activeProject || c.projectId === state.activeProject);
  $("chats").innerHTML = shown.map((c) => `
    <li data-chat="${c.id}" class="${state.current?.id === c.id ? "active" : ""}" title="${esc(c.title)}">
      <span class="name">${esc(c.title)}</span>${c.projectId && !state.activeProject ? `<span class="hint small">${esc(projectOf(c)?.name || "")}</span>` : ""}
    </li>`).join("") || `<li class="hint small" style="cursor:default">Noch keine Chats.</li>`;
}

const projectOf = (chat) => state.projects.find((p) => p.id === chat?.projectId) || null;

/** Subfolder of a project folder for everything the AI creates (like the app). */
const OUTPUT = "Ausgabe/";

/**
 * The folder for [chat]: its project's (files go into "Ausgabe"; without file access the AI sees only that), else the
 * standard one.
 */
function folderOf(chat) {
  const project = projectOf(chat) || state.projects.find((p) => p.id === state.activeProject) || null;
  if (project && state.handles[project.id]) {
    return { key: project.id, handle: state.handles[project.id], label: `Projektordner von „${project.name}“`, output: OUTPUT, readable: project.readFiles !== false };
  }
  if (state.handles[DEFAULT_KEY]) return { key: DEFAULT_KEY, handle: state.handles[DEFAULT_KEY], label: "Standard-Ordner", output: "", readable: true };
  return null;
}

/** Where a new file goes: into "Ausgabe" of a project folder. */
const outPath = (folder, path) => (folder.output && !path.startsWith(folder.output) ? folder.output + path : path);

/** What the AI may read: everything, or only "Ausgabe" without file access. */
function readPath(folder, path) {
  const p = String(path || "").replace(/\\/g, "/").replace(/^\/+/, "");
  return folder.readable || p.startsWith(folder.output) ? p : folder.output + p;
}

async function renderFolderStatus() {
  const el = $("folderStatus");
  const f = folderOf(state.current);
  if (!folders.SUPPORTED) {
    el.textContent = "Dateien werden heruntergeladen (Projektordner gibt es in Chrome und Edge).";
    return;
  }
  if (!f) {
    el.innerHTML = "Kein Ordner gewählt – Dateien werden heruntergeladen. ";
    return;
  }
  const ok = await folders.permitted(f.handle, false);
  el.innerHTML = `📂 ${esc(f.handle.name)} (${esc(f.label)})` + (ok ? "" : ` · <a href="#" id="grant">Zugriff erlauben</a>`);
  $("grant")?.addEventListener("click", async (e) => {
    e.preventDefault();
    await folders.permitted(f.handle, true);
    renderFolderStatus();
  });
}

function renderMessages() {
  if (pendingFrame) { cancelAnimationFrame(pendingFrame); pendingFrame = 0; }
  const box = $("messages");
  const chat = state.current;
  box.querySelectorAll(".msg").forEach((n) => n.remove());
  $("welcome").hidden = !!chat?.messages.length;
  if (!chat) return;
  const results = Object.fromEntries(chat.messages.filter((m) => m.role === "tool").map((m) => [m.toolCallId, m]));
  chat.messages.forEach((m, i) => {
    if (m.role === "tool") return;
    box.appendChild(messageEl(m, i, results));
  });
  box.scrollTop = box.scrollHeight;
}

function splitThinking(m) {
  let content = m.content || "";
  let thinking = m.reasoning || "";
  const t = /^\s*<think>([\s\S]*?)(<\/think>|$)/.exec(content);
  if (t) { thinking += t[1]; content = content.slice(t[0].length); }
  return { content: content.trim(), thinking: thinking.trim() };
}

function messageEl(m, index, results = {}) {
  const el = document.createElement("div");
  el.className = "msg " + m.role;
  if (m.role === "user") {
    const files = (m.files || []).map((f) => `<span class="chip"><span>📄 ${esc(f.name)}</span></span>`).join("") +
      (m.images || []).map((u) => `<img src="${u}" alt="Bild">`).join("");
    el.innerHTML = `<div>${files ? `<div class="files">${files}</div>` : ""}<div class="bubble">${esc(m.content)}</div></div>`;
    return el;
  }
  const { content, thinking } = splitThinking(m);
  const tools = (m.toolCalls || []).map((c) => {
    const r = results[c.id];
    let args = {};
    try { args = JSON.parse(c.function.arguments || "{}"); } catch { /* shown raw */ }
    return `<div class="tool ${r ? (r.ok ? "ok" : "err") : ""}">${esc(toolLabel(c.function.name, args))}${r ? ` – ${esc(r.summary || "")}` : " …"}</div>`;
  }).join("");
  el.innerHTML = `<div class="bubble ${m.streaming ? "cursor" : ""} ${m.error ? "error" : ""}">
      ${thinking ? `<details class="thinking"><summary>Gedankengang</summary>${md(thinking)}</details>` : ""}
      ${m.error ? esc(m.content) : md(content)}${tools}
    </div>`;
  if (!m.streaming && !m.error && content) {
    const actions = document.createElement("div");
    actions.className = "msg-actions";
    actions.innerHTML = `<button data-a="copy">Kopieren</button><button data-a="docx">Als Word</button>` +
      `<button data-a="xlsx" ${content.includes("|") ? "" : "hidden"}>Tabellen als Excel</button><button data-a="md">Als Markdown</button><button data-a="print">Drucken / PDF</button>`;
    actions.addEventListener("click", (e) => answerAction(e.target.dataset.a, content));
    el.appendChild(actions);
  }
  return el;
}

const toolLabel = (name, a) => ({
  write_file: `Datei schreiben: ${a.path || ""}`,
  read_file: `Datei lesen: ${a.path || ""}`,
  list_files: `Ordner ansehen: ${a.path || "/"}`,
}[name] || name);

function updateStreaming(m) {
  if (pendingFrame) return;
  pendingFrame = requestAnimationFrame(() => {
    pendingFrame = 0;
    const box = $("messages");
    const last = box.querySelector(".msg:last-of-type");
    const near = box.scrollHeight - box.scrollTop - box.clientHeight < 120;
    const el = messageEl(m);
    if (last && last.classList.contains("assistant") && last.dataset.live === "1") last.replaceWith(el);
    else box.appendChild(el);
    el.dataset.live = "1";
    if (near) box.scrollTop = box.scrollHeight;
  });
}

function renderAttachments() {
  $("attachments").innerHTML = state.pending.map((a, i) =>
    `<span class="chip" title="${esc(a.note || a.name)}"><span>${a.kind === "image" ? "🖼" : "📄"} ${esc(a.name)}${a.truncated ? " (gekürzt)" : ""}</span><button data-i="${i}" aria-label="Entfernen">✕</button></span>`).join("");
}

function banner(text) {
  const b = $("banner");
  b.textContent = text;
  b.hidden = !text;
}

function renderHeader() {
  const chat = state.current;
  $("chatTitle").textContent = chat?.title || "Neuer Chat";
  const p = projectOf(chat) || state.projects.find((x) => x.id === state.activeProject);
  $("chatMeta").textContent = p ? `Projekt: ${p.name}` : "";
  const name = state.settings.userName.trim();
  $("greeting").textContent = name ? `Hallo ${name}, wie kann ich helfen?` : "Wie kann ich helfen?";
  renderModels();
  renderFolderStatus();
}

// ------------------------------------------------------------------ chats

function openChat(id) {
  state.current = id ? state.chats.find((c) => c.id === id) || null : null;
  if (state.current?.projectId) state.activeProject = state.current.projectId;
  renderHeader();
  renderSidebar();
  renderMessages();
  $("sidebar").classList.remove("open");
}

async function saveChat(chat) {
  chat.updatedAt = Date.now();
  await store.put("chats", chat.id, chat);
  state.chats = [chat, ...state.chats.filter((c) => c.id !== chat.id)];
  renderSidebar();
}

function newChat() {
  return {
    id: store.uid(), title: "Neuer Chat", projectId: state.activeProject, model: $("model").value || state.settings.model,
    messages: [], createdAt: Date.now(), updatedAt: Date.now(),
  };
}

// ------------------------------------------------------------------ the agent

function systemPrompt(chat, folder) {
  const s = state.settings;
  const parts = [
    "Du bist ADEU LLM, ein Assistent für Lehre, Forschung und Verwaltung an der Professur für Allgemeine Didaktik und Empirische " +
      "Unterrichtsforschung der TU Dresden. Antworte auf Deutsch (außer der Nutzer schreibt in einer anderen Sprache), klar und " +
      "geschlechtergerecht, mit Markdown für Struktur. Erfinde keine Quellen oder Fakten; fehlen Angaben, frag nach oder setze [Platzhalter].",
  ];
  const about = [];
  if (s.userName.trim()) about.push(`Vorname: ${s.userName.trim()}`);
  if (s.senderInfo.trim()) about.push(`Absenderangaben (für Briefe und E-Mails): ${s.senderInfo.trim().replace(/\n/g, " / ")}`);
  about.push(...s.memory.split("\n").map((l) => l.trim()).filter(Boolean));
  if (about.length) parts.push("Was du über den Nutzer weißt:\n" + about.map((a) => "- " + a).join("\n"));
  const p = projectOf(chat);
  if (p?.instructions?.trim()) parts.push(`Dieser Chat gehört zum Projekt „${p.name}“. Anweisungen für das Projekt:\n${p.instructions.trim()}`);
  if (folder) {
    const out = folder.output ? ` Neue Dateien legt write_file automatisch im Unterordner „${folder.output.slice(0, -1)}“ ab.` : "";
    const read = folder.readable
      ? "Mit list_files und read_file liest du darin. Bevor du für das Projekt schreibst (Briefe, Texte, Zusammenfassungen), sieh nach, " +
        "welche Dateien zum Thema passen, lies sie und nutze sie."
      : `Lesen darfst du nur im Unterordner „${folder.output.slice(0, -1)}“ (list_files, read_file); die übrigen Dateien sind nicht freigegeben.`;
    parts.push(`Der Nutzer hat den Ordner „${folder.handle.name}“ (${folder.label}) auf seinem Rechner freigegeben. ${read} ` +
      "Mit write_file legst du Dateien an: .docx (Word, Inhalt als Markdown), " +
      ".xlsx (Excel, jede Markdown-Tabelle wird ein Tabellenblatt), .md, .txt, .csv, .html, .ics … Wenn der Nutzer ein Dokument, " +
      "eine Übersicht oder Tabelle haben möchte, speichere sie dort und nenne danach kurz den Dateinamen." + out + " Überschreibe vorhandene " +
      "Dateien nur, wenn der Nutzer das will (overwrite=true).");
  } else {
    parts.push("Mit write_file erstellst du Dateien (.docx aus Markdown, .xlsx aus Markdown-Tabellen, .md, .txt, .csv, .html …); " +
      "sie werden im Browser heruntergeladen, weil kein Ordner gewählt ist.");
  }
  parts.push(`Heute ist ${new Date().toLocaleDateString("de-DE", { weekday: "long", year: "numeric", month: "long", day: "numeric" })}.`);
  return parts.join("\n\n");
}

function tools(folder) {
  const defs = [{
    name: "write_file",
    description: "Legt eine Datei an (im freigegebenen Ordner, bei Projekten unter „Ausgabe“, sonst als Download). Word (.docx) und Excel (.xlsx) entstehen aus Markdown.",
    parameters: {
      type: "object",
      properties: {
        path: { type: "string", description: "Dateiname mit Endung, optional mit Unterordner, z.B. „Protokolle/Sitzung 3.docx“" },
        content: { type: "string", description: "Inhalt: Markdown für .docx/.xlsx/.md, sonst der Text der Datei" },
        overwrite: { type: "boolean", description: "true: eine vorhandene Datei überschreiben" },
      },
      required: ["path", "content"],
    },
  }];
  if (folder) {
    defs.push({
      name: "list_files",
      description: "Listet Dateien und Unterordner im freigegebenen Ordner (mit Größe und Datum).",
      parameters: { type: "object", properties: { path: { type: "string", description: "Unterordner (leer = oberste Ebene)" }, recursive: { type: "boolean" } } },
    }, {
      name: "read_file",
      description: "Liest eine Datei aus dem freigegebenen Ordner: Text, Markdown, CSV, Word, PowerPoint, Excel (als Markdown), PDF.",
      parameters: { type: "object", properties: { path: { type: "string" } }, required: ["path"] },
    });
  }
  return defs.map((f) => ({ type: "function", function: f }));
}

async function runTool(call, folder) {
  let a;
  try { a = JSON.parse(call.function.arguments || "{}"); } catch { return { ok: false, text: "FEHLER: Argumente sind kein gültiges JSON.", summary: "ungültige Argumente" }; }
  try {
    switch (call.function.name) {
      case "write_file": {
        const asked = String(a.path || "").trim().replace(/\\/g, "/").replace(/^\/+/, "");
        if (!asked) throw new Error("path fehlt.");
        const path = folder ? outPath(folder, asked) : asked;
        const ext = (path.split(".").pop() || "").toLowerCase();
        const data = ext === "docx" ? await markdownToDocx(a.content || "") : ext === "xlsx" ? await markdownToXlsx(a.content || "") : String(a.content ?? "");
        if (folder && await folders.permitted(folder.handle, false)) {
          if (!a.overwrite && await folders.exists(folder.handle, path)) {
            return { ok: false, text: `FEHLER: „${path}“ gibt es schon im Ordner. Frag den Nutzer, ob überschrieben werden soll (overwrite=true), oder wähle einen anderen Namen.`, summary: "existiert schon" };
          }
          await folders.writeFile(folder.handle, path, data);
          return { ok: true, text: `Gespeichert in „${folder.handle.name}“: ${path}`, summary: `gespeichert in „${folder.handle.name}“` };
        }
        folders.download(path.split("/").pop(), data);
        return {
          ok: true,
          text: `Die Datei „${path.split("/").pop()}“ wurde im Browser heruntergeladen${folder ? " (für den Ordner fehlt die Erlaubnis – der Nutzer kann sie unten mit „Zugriff erlauben“ erteilen)" : ""}.`,
          summary: "heruntergeladen",
        };
      }
      case "list_files": {
        if (!(await folders.permitted(folder.handle, false))) throw new Error("Für den Ordner fehlt die Erlaubnis (unten „Zugriff erlauben“).");
        const where = readPath(folder, a.path || "");
        const entries = await folders.list(folder.handle, where, !!a.recursive).catch((e) => {
          if (e?.name === "NotFoundError" && where.startsWith(OUTPUT)) return [];
          throw e;
        });
        return { ok: true, text: entries.length ? entries.join("\n") : "(leer)", summary: `${entries.length} Einträge` };
      }
      case "read_file": {
        if (!(await folders.permitted(folder.handle, false))) throw new Error("Für den Ordner fehlt die Erlaubnis (unten „Zugriff erlauben“).");
        const file = await folders.readFile(folder.handle, readPath(folder, a.path));
        const text = folders.isText(file.name) ? await file.text() : (await readAttachment(file)).text || "";
        return { ok: true, text: text.slice(0, 120000) || "(keine lesbaren Inhalte)", summary: `${Math.round(text.length / 1000)}k Zeichen` };
      }
      default:
        return { ok: false, text: `FEHLER: Unbekanntes Werkzeug ${call.function.name}.`, summary: "unbekannt" };
    }
  } catch (e) {
    const msg = folders.explain(e);
    return { ok: false, text: "FEHLER: " + msg, summary: msg };
  }
}

function apiMessages(chat, folder) {
  const out = [{ role: "system", content: systemPrompt(chat, folder) }];
  // Only tool calls with a result (a stopped run can leave some without), otherwise the API rejects the history.
  const answered = new Set(chat.messages.filter((m) => m.role === "tool").map((m) => m.toolCallId));
  for (const m of chat.messages) {
    if (m.error) continue;
    if (m.role === "user") {
      const files = (m.files || []).map((f) => `Angehängte Datei „${f.name}“${f.truncated ? " (gekürzt)" : ""}:\n<<<\n${f.text}\n>>>`).join("\n\n");
      const text = (files ? files + "\n\n" : "") + m.content + `\n\n(geschrieben: ${new Date(m.at || Date.now()).toLocaleString("de-DE")})`;
      out.push(m.images?.length
        ? { role: "user", content: [{ type: "text", text }, ...m.images.map((url) => ({ type: "image_url", image_url: { url } }))] }
        : { role: "user", content: text });
    } else if (m.role === "assistant") {
      const msg = { role: "assistant", content: splitThinking(m).content };
      const calls = (m.toolCalls || []).filter((c) => answered.has(c.id));
      if (calls.length) msg.tool_calls = calls;
      if (msg.content || msg.tool_calls) out.push(msg);
    } else if (m.role === "tool") {
      out.push({ role: "tool", tool_call_id: m.toolCallId, content: m.content });
    }
  }
  return out;
}

async function send() {
  const text = $("input").value.trim();
  if (!text && !state.pending.length) return;
  if (!state.settings.apiKey) { openSettings(); return; }
  let chat = state.current;
  if (!chat) { chat = newChat(); state.current = chat; }
  const folder = folderOf(chat);
  // Directly in the click: the browser shows its question about the folder only now.
  if (folder) await folders.permitted(folder.handle, true);
  chat.model = $("model").value;
  const files = state.pending.filter((a) => a.kind === "file").map(({ name, text: t, truncated }) => ({ name, text: t, truncated }));
  const images = state.pending.filter((a) => a.kind === "image").map((a) => a.url);
  chat.messages.push({ role: "user", content: text || "Siehe Anhang.", files, images, at: Date.now() });
  if (chat.title === "Neuer Chat") chat.title = (text || files[0]?.name || "Bild").replace(/\s+/g, " ").slice(0, 60);
  state.pending = [];
  renderAttachments();
  $("input").value = "";
  autosize();
  await saveChat(chat);
  renderHeader();
  renderMessages();
  await run(chat, folder);
}

async function run(chat, folder) {
  state.abort = new AbortController();
  $("send").hidden = true;
  $("stop").hidden = false;
  try {
    for (let round = 0; round < MAX_ROUNDS; round++) {
      const live = { role: "assistant", content: "", reasoning: "", streaming: true };
      const result = await llm.chat(state.settings, {
        model: chat.model,
        messages: apiMessages(chat, folder),
        tools: tools(folder),
        signal: state.abort.signal,
        onDelta: (d) => {
          live.content = d.content;
          live.reasoning = d.reasoning;
          if (d.notice) banner(d.notice);
          if (state.current === chat) updateStreaming(live);
        },
      });
      banner("");
      chat.messages.push({ role: "assistant", content: result.content, reasoning: result.reasoning, toolCalls: result.toolCalls });
      await saveChat(chat);
      if (state.current === chat) renderMessages();
      if (!result.toolCalls.length) return;
      for (const call of result.toolCalls) {
        const r = await runTool(call, folder);
        chat.messages.push({ role: "tool", toolCallId: call.id, name: call.function.name, content: r.text, ok: r.ok, summary: r.summary });
        await saveChat(chat);
        if (state.current === chat) renderMessages();
      }
    }
    chat.messages.push({ role: "assistant", error: true, content: `Nach ${MAX_ROUNDS} Werkzeugschritten angehalten – schreib „weiter“, um fortzufahren.` });
  } catch (e) {
    if (e.name !== "AbortError") chat.messages.push({ role: "assistant", error: true, content: e.message || String(e) });
  } finally {
    state.abort = null;
    $("send").hidden = false;
    $("stop").hidden = true;
    await saveChat(chat);
    if (state.current === chat) renderMessages();
    renderFolderStatus();
  }
}

// ------------------------------------------------------------------ actions

async function answerAction(action, content) {
  const base = (state.current?.title || "Antwort").replace(/[\\/:*?"<>|]/g, " ").slice(0, 60);
  const folder = folderOf(state.current);
  const saveOut = async (name, data) => {
    if (folder && await folders.permitted(folder.handle, true)) {
      const first = outPath(folder, name);
      let target = first;
      for (let n = 2; await folders.exists(folder.handle, target); n++) target = first.replace(/(\.[^.]+)$/, ` (${n})$1`);
      await folders.writeFile(folder.handle, target, data);
      banner(`Gespeichert in „${folder.handle.name}“: ${target}`);
    } else {
      folders.download(name, data);
    }
  };
  try {
    if (action === "copy") { await navigator.clipboard.writeText(content); banner("Kopiert."); }
    if (action === "md") await saveOut(base + ".md", content);
    if (action === "docx") await saveOut(base + ".docx", await markdownToDocx(content));
    if (action === "xlsx") await saveOut(base + ".xlsx", await markdownToXlsx(content));
    if (action === "print") window.print();
  } catch (e) {
    banner(folders.explain(e));
  }
}

async function addFiles(list) {
  for (const f of list) {
    try {
      const a = await readAttachment(f);
      state.pending.push(a);
      if (a.note) banner(a.note);
    } catch (e) {
      banner(e.message);
    }
  }
  renderAttachments();
}

// ------------------------------------------------------------------ dialogs

function openSettings() {
  const s = state.settings;
  $("apiKey").value = s.apiKey;
  $("baseUrl").value = s.baseUrl;
  $("userName").value = s.userName;
  $("senderInfo").value = s.senderInfo;
  $("memory").value = s.memory;
  showFolderInfo($("defaultFolderInfo"), state.handles[DEFAULT_KEY]);
  $("settingsDialog").showModal();
}

function showFolderInfo(el, handle) {
  el.textContent = !folders.SUPPORTED
    ? "Dieser Browser kann keine Ordner freigeben (geht in Chrome und Edge). Dateien werden heruntergeladen."
    : handle ? `Gewählt: „${handle.name}“` : "Kein Ordner gewählt – Dateien werden heruntergeladen.";
}

let editing = null;
let editingHandle;

function openProject(project) {
  editing = project;
  editingHandle = project ? state.handles[project.id] || null : null;
  $("projectDialogTitle").textContent = project ? "Projekt bearbeiten" : "Neues Projekt";
  $("projectName").value = project?.name || "";
  $("projectInstructions").value = project?.instructions || "";
  $("deleteProject").hidden = !project;
  $("projectReadFiles").checked = project?.readFiles !== false;
  showFolderInfo($("projectFolderInfo"), editingHandle);
  $("projectDialog").showModal();
}

async function exportData() {
  const data = { app: "ADEU LLM (Browser)", version: 1, exportedAt: new Date().toISOString(), settings: { ...state.settings, apiKey: "" }, projects: state.projects, chats: state.chats };
  folders.download(`ADEU-LLM-Browser-Sicherung-${new Date().toISOString().slice(0, 10)}.json`, new Blob([JSON.stringify(data)], { type: "application/json" }));
}

async function importData(file) {
  const data = JSON.parse(await file.text());
  for (const p of data.projects || []) await store.put("projects", p.id, p);
  for (const c of data.chats || []) await store.put("chats", c.id, c);
  state.projects = (await store.all("projects")).sort((a, b) => a.name.localeCompare(b.name, "de"));
  state.chats = (await store.all("chats")).sort((a, b) => b.updatedAt - a.updatedAt);
  renderSidebar();
  banner(`Eingespielt: ${(data.chats || []).length} Chats, ${(data.projects || []).length} Projekte.`);
}

// ------------------------------------------------------------------ events

function autosize() {
  const t = $("input");
  t.style.height = "auto";
  t.style.height = Math.min(220, t.scrollHeight) + "px";
}

const STARTERS = [
  ["Zusammenfassen", "Fasse die angehängte Datei in 5–7 Stichpunkten zusammen."],
  ["E-Mail schreiben", "Schreibe eine freundliche, förmliche E-Mail an "],
  ["Folien entwerfen", "Entwirf eine Gliederung für 10 Folien zum Thema "],
  ["Übersicht als Excel", "Erstelle eine Übersicht als Excel-Datei (.xlsx) mit Tabellenblättern zu "],
  ["Word-Dokument", "Schreibe ein Word-Dokument (.docx) mit "],
  ["Feedback", "Gib mir konstruktives Feedback zu folgendem Text:\n\n"],
];

function bind() {
  $("starters").innerHTML = STARTERS.map(([l], i) => `<button data-i="${i}">${esc(l)}</button>`).join("");
  $("starters").addEventListener("click", (e) => {
    const i = e.target.dataset.i;
    if (i == null) return;
    $("input").value = STARTERS[i][1];
    $("input").focus();
    autosize();
  });
  $("send").addEventListener("click", send);
  $("stop").addEventListener("click", () => state.abort?.abort());
  $("input").addEventListener("input", autosize);
  $("input").addEventListener("keydown", (e) => {
    if (e.key === "Enter" && !e.shiftKey && !e.isComposing) { e.preventDefault(); if (!state.abort) send(); }
  });
  $("attach").addEventListener("click", () => $("fileInput").click());
  $("fileInput").addEventListener("change", (e) => { addFiles([...e.target.files]); e.target.value = ""; });
  $("attachments").addEventListener("click", (e) => {
    const i = e.target.dataset.i;
    if (i == null) return;
    state.pending.splice(Number(i), 1);
    renderAttachments();
  });
  document.addEventListener("dragover", (e) => e.preventDefault());
  document.addEventListener("drop", (e) => { e.preventDefault(); if (e.dataTransfer?.files?.length) addFiles([...e.dataTransfer.files]); });
  document.addEventListener("paste", (e) => {
    const imgs = [...(e.clipboardData?.files || [])].filter((f) => f.type.startsWith("image/"));
    if (imgs.length) addFiles(imgs);
  });

  $("newChat").addEventListener("click", () => openChat(null));
  $("chats").addEventListener("click", (e) => { const li = e.target.closest("[data-chat]"); if (li) openChat(li.dataset.chat); });
  $("projects").addEventListener("click", (e) => {
    const edit = e.target.closest("[data-edit]");
    if (edit) { openProject(state.projects.find((p) => p.id === edit.dataset.edit)); return; }
    const li = e.target.closest("[data-project]");
    if (!li) return;
    state.activeProject = state.activeProject === li.dataset.project ? null : li.dataset.project;
    if (state.current && state.current.projectId !== state.activeProject) state.current = null;
    renderSidebar();
    renderHeader();
    renderMessages();
  });
  $("newProject").addEventListener("click", () => openProject(null));
  $("openSidebar").addEventListener("click", () => $("sidebar").classList.add("open"));
  $("closeSidebar").addEventListener("click", () => $("sidebar").classList.remove("open"));
  $("model").addEventListener("change", async () => {
    state.settings.model = $("model").value;
    store.saveSettings(state.settings);
    if (state.current) { state.current.model = $("model").value; await saveChat(state.current); }
  });

  $("openSettings").addEventListener("click", openSettings);
  $("saveSettings").addEventListener("click", () => {
    const s = state.settings;
    const keyChanged = s.apiKey !== $("apiKey").value.trim() || s.baseUrl !== $("baseUrl").value.trim();
    s.apiKey = $("apiKey").value.trim();
    s.baseUrl = $("baseUrl").value.trim() || store.DEFAULT_SETTINGS.baseUrl;
    s.userName = $("userName").value.trim();
    s.senderInfo = $("senderInfo").value;
    s.memory = $("memory").value;
    store.saveSettings(s);
    banner("");
    renderHeader();
    if (keyChanged && s.apiKey) loadModels();
  });
  $("pickDefaultFolder").addEventListener("click", async () => {
    if (!folders.SUPPORTED) return;
    try {
      const h = await folders.pick("standard");
      if (!h) return;
      state.handles[DEFAULT_KEY] = h;
      await folders.save(DEFAULT_KEY, h);
      showFolderInfo($("defaultFolderInfo"), h);
      renderFolderStatus();
    } catch (e) { $("defaultFolderInfo").textContent = e.message; }
  });
  $("clearDefaultFolder").addEventListener("click", async () => {
    delete state.handles[DEFAULT_KEY];
    await folders.save(DEFAULT_KEY, null);
    showFolderInfo($("defaultFolderInfo"), null);
    renderFolderStatus();
  });
  $("exportData").addEventListener("click", exportData);
  $("importData").addEventListener("click", () => $("importFile").click());
  $("importFile").addEventListener("change", async (e) => {
    const f = e.target.files[0];
    e.target.value = "";
    if (f) importData(f).catch((err) => banner("Sicherung nicht lesbar: " + err.message));
  });

  $("pickProjectFolder").addEventListener("click", async () => {
    if (!folders.SUPPORTED) return;
    try {
      const h = await folders.pick(editing?.id || "projekt");
      if (!h) return;
      editingHandle = h;
      showFolderInfo($("projectFolderInfo"), h);
    } catch (e) { $("projectFolderInfo").textContent = e.message; }
  });
  $("clearProjectFolder").addEventListener("click", () => { editingHandle = null; showFolderInfo($("projectFolderInfo"), null); });
  $("saveProject").addEventListener("click", async (e) => {
    const name = $("projectName").value.trim();
    if (!name) { e.preventDefault(); $("projectName").focus(); return; }
    const p = editing ? { ...editing } : { id: store.uid(), createdAt: Date.now() };
    p.name = name;
    p.instructions = $("projectInstructions").value;
    p.readFiles = $("projectReadFiles").checked;
    await store.put("projects", p.id, p);
    if (editingHandle) state.handles[p.id] = editingHandle; else delete state.handles[p.id];
    await folders.save(p.id, editingHandle);
    state.projects = [...state.projects.filter((x) => x.id !== p.id), p].sort((a, b) => a.name.localeCompare(b.name, "de"));
    if (!editing) state.activeProject = p.id;
    renderSidebar();
    renderHeader();
  });
  $("deleteProject").addEventListener("click", async () => {
    if (!editing || !confirm(`Projekt „${editing.name}“ löschen? Die Chats bleiben erhalten (ohne Projekt), die Dateien im Projektordner auch.`)) return;
    await store.remove("projects", editing.id);
    await folders.save(editing.id, null);
    delete state.handles[editing.id];
    for (const c of state.chats.filter((x) => x.projectId === editing.id)) { c.projectId = null; await store.put("chats", c.id, c); }
    state.projects = state.projects.filter((x) => x.id !== editing.id);
    if (state.activeProject === editing.id) state.activeProject = null;
    $("projectDialog").close();
    renderSidebar();
    renderHeader();
  });

  $("chatMenu").addEventListener("click", () => { $("chatMenuPop").hidden = !$("chatMenuPop").hidden; });
  document.addEventListener("click", (e) => { if (!e.target.closest("#chatMenu, #chatMenuPop")) $("chatMenuPop").hidden = true; });
  $("chatMenuPop").addEventListener("click", async (e) => {
    const act = e.target.dataset.act;
    $("chatMenuPop").hidden = true;
    const chat = state.current;
    if (!chat) return;
    if (act === "export-md") {
      const text = chat.messages.filter((m) => (m.role === "user" || m.role === "assistant") && !m.error && (m.content || "").trim())
        .map((m) => `## ${m.role === "user" ? "Ich" : "ADEU LLM"}\n\n${m.role === "user" ? m.content : splitThinking(m).content}`).join("\n\n");
      folders.download(chat.title.replace(/[\\/:*?"<>|]/g, " ") + ".md", `# ${chat.title}\n\n${text}\n`);
    }
    if (act === "rename") {
      const t = prompt("Neuer Name des Chats", chat.title);
      if (t?.trim()) { chat.title = t.trim(); await saveChat(chat); renderHeader(); }
    }
    if (act === "move") {
      $("moveTarget").innerHTML = `<option value="">Ohne Projekt</option>` + state.projects.map((p) => `<option value="${p.id}">${esc(p.name)}</option>`).join("");
      $("moveTarget").value = chat.projectId || "";
      $("moveDialog").showModal();
    }
    if (act === "delete" && confirm(`Chat „${chat.title}“ löschen?`)) {
      await store.remove("chats", chat.id);
      state.chats = state.chats.filter((c) => c.id !== chat.id);
      openChat(null);
    }
  });
  $("moveDialog").addEventListener("close", async () => {
    if ($("moveDialog").returnValue !== "ok" || !state.current) return;
    state.current.projectId = $("moveTarget").value || null;
    await saveChat(state.current);
    openChat(state.current.id);
  });
}

start().catch((e) => {
  console.error(e);
  banner("Start fehlgeschlagen: " + (e.message || e));
});
