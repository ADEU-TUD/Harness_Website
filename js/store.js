// Chats, projects and folder handles in IndexedDB (only in this browser); settings in localStorage.

const DB_NAME = "adeu-llm";
const STORES = ["chats", "projects", "handles"];
let dbPromise = null;

function db() {
  if (!dbPromise) {
    dbPromise = new Promise((resolve, reject) => {
      const req = indexedDB.open(DB_NAME, 1);
      req.onupgradeneeded = () => {
        for (const name of STORES) {
          if (!req.result.objectStoreNames.contains(name)) req.result.createObjectStore(name);
        }
      };
      req.onsuccess = () => resolve(req.result);
      req.onerror = () => reject(req.error);
    });
  }
  return dbPromise;
}

async function tx(store, mode, run) {
  const d = await db();
  return new Promise((resolve, reject) => {
    const t = d.transaction(store, mode);
    const result = run(t.objectStore(store));
    t.oncomplete = () => resolve(result && "result" in result ? result.result : undefined);
    t.onerror = () => reject(t.error);
    t.onabort = () => reject(t.error);
  });
}

export const get = (store, key) => tx(store, "readonly", (s) => s.get(key));
export const put = (store, key, value) => tx(store, "readwrite", (s) => s.put(value, key));
export const remove = (store, key) => tx(store, "readwrite", (s) => s.delete(key));
export const all = (store) => tx(store, "readonly", (s) => s.getAll());

const SETTINGS_KEY = "adeu.settings";

export const DEFAULT_SETTINGS = {
  apiKey: "",
  baseUrl: "https://llm.scads.ai/v1",
  model: "deepseek-ai/DeepSeek-V4.1-Flash",
  userName: "",
  senderInfo: "",
  memory: "",
};

export function loadSettings() {
  try {
    return { ...DEFAULT_SETTINGS, ...JSON.parse(localStorage.getItem(SETTINGS_KEY) || "{}") };
  } catch {
    return { ...DEFAULT_SETTINGS };
  }
}

export function saveSettings(s) {
  try {
    localStorage.setItem(SETTINGS_KEY, JSON.stringify(s));
  } catch (e) {
    console.warn("Einstellungen nicht gespeichert", e);
  }
}

export const uid = () => (crypto.randomUUID ? crypto.randomUUID() : String(Date.now()) + Math.random().toString(16).slice(2));
