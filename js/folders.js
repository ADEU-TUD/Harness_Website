// Working folders on this computer through the File System Access API (Chrome, Edge, Opera on Windows, macOS,
// Linux and ChromeOS). Firefox and Safari do not offer it: there files are downloaded instead.
import * as store from "./store.js";

export const SUPPORTED = typeof window.showDirectoryPicker === "function";

const TEXT = new Set(["md", "txt", "csv", "tsv", "json", "html", "htm", "xml", "tex", "bib", "ris", "yaml", "yml", "js", "py", "r", "sql", "css", "ics", "srt", "vtt"]);

/** Asks for a folder; null if cancelled. Must be called directly from a click. */
export async function pick(id) {
  try {
    return await window.showDirectoryPicker({ id: (id || "adeu").replace(/[^a-zA-Z0-9_-]/g, "").slice(0, 32) || "adeu", mode: "readwrite" });
  } catch (e) {
    if (e.name === "AbortError") return null;
    throw new Error(explain(e));
  }
}

export const load = (key) => store.get("handles", key);
export const save = (key, handle) => (handle ? store.put("handles", key, handle) : store.remove("handles", key));

/**
 * Write permission for [handle]. The browser asks again after a restart; [ask] = true shows its question (only
 * possible right after a click), otherwise the current state is returned.
 */
export async function permitted(handle, ask) {
  if (!handle) return false;
  const opts = { mode: "readwrite" };
  if ((await handle.queryPermission(opts)) === "granted") return true;
  if (!ask) return false;
  try {
    return (await handle.requestPermission(opts)) === "granted";
  } catch {
    return false;
  }
}

export function explain(e) {
  if (e?.name === "SecurityError") return "Der Browser erlaubt diesen Ordner nicht (Systemordner oder die oberste Ebene des Benutzerordners). Bitte einen Unterordner wählen, z. B. in Google Drive „G:\\Meine Ablage\\ADEU“.";
  if (e?.name === "NotAllowedError") return "Keine Berechtigung für den Ordner – bitte beim Nachfragen des Browsers „Dateien bearbeiten“ erlauben.";
  if (e?.name === "NotFoundError") return "Ordner oder Datei nicht gefunden – ist das Netzlaufwerk verbunden bzw. Google Drive gestartet?";
  if (e?.name === "NoModificationAllowedError" || e?.name === "InvalidModificationError") return "Die Datei ist gesperrt (in Word/Excel geöffnet?) oder der Ordner ist schreibgeschützt.";
  return e?.message || String(e);
}

/** "a/b/c.txt" inside [root]: the folder handle of "a/b" and the name "c.txt"; folders are created on [create]. */
async function resolve(root, path, create) {
  const parts = String(path || "").replace(/\\/g, "/").split("/").filter((p) => p && p !== ".");
  if (parts.some((p) => p === "..")) throw new Error("Pfade mit „..“ sind nicht erlaubt – nur innerhalb des freigegebenen Ordners.");
  let dir = root;
  for (const p of parts.slice(0, -1)) dir = await dir.getDirectoryHandle(p, { create });
  return { dir, name: parts[parts.length - 1] || "" };
}

export async function list(root, path = "", recursive = false) {
  const out = [];
  async function walk(dir, prefix, depth) {
    for await (const [name, h] of dir.entries()) {
      if (name.startsWith(".")) continue;
      if (h.kind === "directory") {
        out.push(prefix + name + "/");
        if (recursive && depth < 4 && out.length < 500) await walk(h, prefix + name + "/", depth + 1);
      } else {
        const f = await h.getFile();
        out.push(`${prefix}${name} (${Math.max(1, Math.round(f.size / 1024))} KB, ${new Date(f.lastModified).toLocaleString("de-DE")})`);
      }
      if (out.length >= 500) return;
    }
  }
  let dir = root;
  if (path) for (const p of path.replace(/\\/g, "/").split("/").filter(Boolean)) dir = await dir.getDirectoryHandle(p);
  await walk(dir, "", 0);
  return out.sort();
}

export async function readFile(root, path) {
  const { dir, name } = await resolve(root, path, false);
  return (await dir.getFileHandle(name)).getFile();
}

export const isText = (name) => TEXT.has((name.split(".").pop() || "").toLowerCase());

/** Writes [data] (string or Blob) to [path]; returns the path relative to the folder. */
export async function writeFile(root, path, data) {
  const { dir, name } = await resolve(root, path, true);
  if (!name) throw new Error("Dateiname fehlt.");
  const fh = await dir.getFileHandle(name, { create: true });
  const w = await fh.createWritable();
  try {
    await w.write(data);
  } finally {
    await w.close();
  }
  return path;
}

export async function exists(root, path) {
  try {
    const { dir, name } = await resolve(root, path, false);
    await dir.getFileHandle(name);
    return true;
  } catch {
    return false;
  }
}

/** Download in the browser (no working folder or no permission). */
export function download(name, data) {
  const blob = data instanceof Blob ? data : new Blob([data], { type: "text/plain;charset=utf-8" });
  const a = document.createElement("a");
  a.href = URL.createObjectURL(blob);
  a.download = name;
  document.body.appendChild(a);
  a.click();
  setTimeout(() => { URL.revokeObjectURL(a.href); a.remove(); }, 2000);
}
