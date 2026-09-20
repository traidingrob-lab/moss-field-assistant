// Local-first storage for Moss. Everything here lives in the browser's
// IndexedDB until Phase 2 wires it up to OneDrive/SharePoint via Microsoft
// Graph. Swap the functions in this file for Graph API calls later without
// touching any screen code, since every screen only talks to MossDB.

const MossDB = (() => {
  const DB_NAME = "moss-field-assistant";
  const DB_VERSION = 1;
  let dbPromise = null;

  function open() {
    if (dbPromise) return dbPromise;
    dbPromise = new Promise((resolve, reject) => {
      const req = indexedDB.open(DB_NAME, DB_VERSION);
      req.onupgradeneeded = () => {
        const db = req.result;
        if (!db.objectStoreNames.contains("projects")) {
          db.createObjectStore("projects", { keyPath: "id" });
        }
        if (!db.objectStoreNames.contains("issues")) {
          const s = db.createObjectStore("issues", { keyPath: "id" });
          s.createIndex("projectId", "projectId");
        }
        if (!db.objectStoreNames.contains("captures")) {
          const s = db.createObjectStore("captures", { keyPath: "id" });
          s.createIndex("projectId", "projectId");
          s.createIndex("type", "type");
        }
        if (!db.objectStoreNames.contains("logs")) {
          const s = db.createObjectStore("logs", { keyPath: "id" });
          s.createIndex("projectId", "projectId");
        }
      };
      req.onsuccess = () => resolve(req.result);
      req.onerror = () => reject(req.error);
    });
    return dbPromise;
  }

  async function tx(store, mode) {
    const db = await open();
    return db.transaction(store, mode).objectStore(store);
  }

  async function put(store, value) {
    const s = await tx(store, "readwrite");
    return new Promise((resolve, reject) => {
      const req = s.put(value);
      req.onsuccess = () => resolve(value);
      req.onerror = () => reject(req.error);
    });
  }

  async function getAll(store) {
    const s = await tx(store, "readonly");
    return new Promise((resolve, reject) => {
      const req = s.getAll();
      req.onsuccess = () => resolve(req.result);
      req.onerror = () => reject(req.error);
    });
  }

  async function getByIndex(store, index, value) {
    const s = await tx(store, "readonly");
    return new Promise((resolve, reject) => {
      const req = s.index(index).getAll(value);
      req.onsuccess = () => resolve(req.result);
      req.onerror = () => reject(req.error);
    });
  }

  async function get(store, id) {
    const s = await tx(store, "readonly");
    return new Promise((resolve, reject) => {
      const req = s.get(id);
      req.onsuccess = () => resolve(req.result);
      req.onerror = () => reject(req.error);
    });
  }

  function uid() {
    return Date.now().toString(36) + Math.random().toString(36).slice(2, 8);
  }

  const SEED_PROJECTS = [
    { id: "helm", name: "Helm", address: "2140 Rosewood Ave · Residential remodel" },
    { id: "redd", name: "Redd", address: "Active project" },
    { id: "liu", name: "Liu", address: "Active project" },
    { id: "haldeman", name: "Haldeman", address: "Active project" },
    { id: "barnhart", name: "Barnhart", address: "Active project" },
    { id: "schmitt", name: "Schmitt", address: "Active project" }
  ];

  async function seedIfEmpty() {
    const existing = await getAll("projects");
    if (existing.length) return;
    for (const p of SEED_PROJECTS) await put("projects", p);
  }

  return {
    uid,
    seedIfEmpty,
    projects: {
      all: () => getAll("projects"),
      get: (id) => get("projects", id),
      add: (project) => put("projects", { status: "Active", ...project }),
      update: async (id, patch) => {
        const existing = await get("projects", id);
        if (!existing) return null;
        return put("projects", { ...existing, ...patch });
      }
    },
    issues: {
      all: () => getAll("issues"),
      forProject: (projectId) => getByIndex("issues", "projectId", projectId),
      add: (issue) =>
        put("issues", {
          id: uid(),
          status: "Open",
          createdAt: new Date().toISOString(),
          ...issue
        })
    },
    captures: {
      forProject: (projectId) => getByIndex("captures", "projectId", projectId),
      add: (capture) =>
        put("captures", {
          id: uid(),
          createdAt: new Date().toISOString(),
          ...capture
        })
    },
    logs: {
      forProject: (projectId) => getByIndex("logs", "projectId", projectId),
      add: (log) =>
        put("logs", { id: uid(), createdAt: new Date().toISOString(), ...log })
    }
  };
})();
