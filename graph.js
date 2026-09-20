// Microsoft Graph calls — uploading captures into the OneDrive folder
// structure that already exists under "MOSS PROJECTS" (built in Phase 1).
// Everything here is best-effort: a failure here should never lose data
// that's already safe in MossDB/IndexedDB, only skip the OneDrive copy.

const GRAPH_BASE = "https://graph.microsoft.com/v1.0";
const MOSS_ROOT = "MOSS PROJECTS";
const MOSS_ACTIVE_ROOT = "MOSS PROJECTS/01 ACTIVE PROJECTS";
const PROJECTS_INDEX_PATH = `${MOSS_ROOT}/moss-index.json`;
const CAPTURES_INDEX_PATH = `${MOSS_ROOT}/moss-captures-index.json`;

// Mirrors the folder structure already created under each project in
// Phase 1. If you rename folders in OneDrive, update this list to match.
const PROJECT_SUBFOLDERS = [
  "01 PLANS & DRAWINGS/CURRENT",
  "01 PLANS & DRAWINGS/SUPERSEDED",
  "02 PHOTOS/01 BEFORE",
  "02 PHOTOS/02 PROGRESS",
  "02 PHOTOS/03 INSPECTIONS",
  "02 PHOTOS/04 CONCEALED WORK",
  "02 PHOTOS/05 COMPLETED",
  "03 INSPECTIONS",
  "04 DAILY LOGS",
  "05 SUBCONTRACTORS",
  "06 MATERIALS",
  "07 OPEN ISSUES",
  "08 CHANGE ORDERS",
  "09 EMAILS & COMMUNICATION",
  "10 COSTS",
  "11 PUNCH LIST"
];

function projectFolderPath(projectName) {
  return `${MOSS_ACTIVE_ROOT}/${projectName.toUpperCase()}`;
}

// Graph's path-based addressing (/root:/a/b/c:) needs each segment
// percent-encoded, but the slashes between segments must stay literal.
function graphPathEncode(path) {
  return path
    .split("/")
    .map(encodeURIComponent)
    .join("/");
}

async function graphFetch(path, token, options = {}) {
  const res = await fetch(`${GRAPH_BASE}${path}`, {
    ...options,
    headers: { Authorization: `Bearer ${token}`, ...(options.headers || {}) }
  });
  if (!res.ok) {
    const body = await res.text().catch(() => "");
    throw new Error(`Graph ${res.status} on ${path}: ${body.slice(0, 200)}`);
  }
  return res;
}

async function folderExists(path, token) {
  try {
    await graphFetch(`/me/drive/root:/${graphPathEncode(path)}`, token);
    return true;
  } catch (err) {
    return false;
  }
}

async function ensureFolder(path, token) {
  if (await folderExists(path, token)) return;
  const segments = path.split("/");
  const name = segments.pop();
  const parentPath = segments.join("/");
  // A top-level folder (no parent segments, e.g. "MOSS PROJECTS" itself)
  // has no colon-addressed parent path — Graph wants POST .../root/children
  // for that case, not .../root:/:children (which 400s on the empty path).
  const childrenUrl = parentPath
    ? `/me/drive/root:/${graphPathEncode(parentPath)}:/children`
    : `/me/drive/root/children`;
  await graphFetch(childrenUrl, token, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({
      name,
      folder: {},
      "@microsoft.graph.conflictBehavior": "replace"
    })
  });
}

// Creates the standard subfolder set under a project, for projects added
// in the app that don't exist in OneDrive yet. Existing Phase-1 projects
// (Helm, Redd, Liu, Haldeman, Barnhart, Schmitt) already have these, so
// this is a no-op for them (folderExists short-circuits each check).
async function ensureProjectFolders(projectName, token) {
  const base = projectFolderPath(projectName);
  await ensureFolder(base, token);
  for (const sub of PROJECT_SUBFOLDERS) {
    const parts = sub.split("/");
    for (let i = 1; i <= parts.length; i++) {
      await ensureFolder(`${base}/${parts.slice(0, i).join("/")}`, token);
    }
  }
}

// Uploads a Blob/File into a project's subfolder. Handles both the
// simple <4MB PUT and Graph's chunked upload session for larger files
// (real phone photos regularly cross 4MB).
async function uploadToOneDrive(projectName, subfolder, fileName, blob, token) {
  const path = `${projectFolderPath(projectName)}/${subfolder}/${fileName}`;
  const encodedPath = graphPathEncode(path);

  if (blob.size <= 4 * 1024 * 1024) {
    await graphFetch(`/me/drive/root:/${encodedPath}:/content`, token, {
      method: "PUT",
      headers: { "Content-Type": blob.type || "application/octet-stream" },
      body: blob
    });
    return;
  }

  const sessionRes = await graphFetch(`/me/drive/root:/${encodedPath}:/createUploadSession`, token, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ item: { "@microsoft.graph.conflictBehavior": "replace" } })
  });
  const session = await sessionRes.json();
  const uploadUrl = session.uploadUrl;

  const CHUNK_SIZE = 5 * 1024 * 1024; // must be a multiple of 320 KiB per Graph docs
  let start = 0;
  while (start < blob.size) {
    const end = Math.min(start + CHUNK_SIZE, blob.size);
    const chunk = blob.slice(start, end);
    // The session's uploadUrl is pre-authenticated — no Authorization header here.
    const res = await fetch(uploadUrl, {
      method: "PUT",
      headers: {
        "Content-Length": String(end - start),
        "Content-Range": `bytes ${start}-${end - 1}/${blob.size}`
      },
      body: chunk
    });
    if (!res.ok && res.status !== 202) {
      const body = await res.text().catch(() => "");
      throw new Error(`Upload session ${res.status}: ${body.slice(0, 200)}`);
    }
    start = end;
  }
}

// ---------- Cross-device project list ----------
// The project list itself (name, status, id) isn't a "capture", so it
// doesn't live under a project's own folder — it's a single shared JSON
// file at the root of MOSS PROJECTS, read and rewritten by whichever
// device last touched a project. This is how the same project list shows
// up whether you're on your phone or your laptop.

// Returns the shared project list, or null if the file doesn't exist yet
// (first run on this account) — callers treat that as "nothing to merge"
// rather than an error.
async function fetchProjectsIndex(token) {
  const encodedPath = graphPathEncode(PROJECTS_INDEX_PATH);
  const res = await fetch(`${GRAPH_BASE}/me/drive/root:/${encodedPath}:/content`, {
    headers: { Authorization: `Bearer ${token}` }
  });
  if (res.status === 404) return null;
  if (!res.ok) {
    const body = await res.text().catch(() => "");
    throw new Error(`Graph ${res.status} fetching project index: ${body.slice(0, 200)}`);
  }
  try {
    return await res.json();
  } catch {
    return null;
  }
}

// Overwrites the shared project list with the full merged set. Always well
// under the 4MB simple-upload limit, even at hundreds of projects.
async function saveProjectsIndex(token, projects) {
  await ensureFolder(MOSS_ROOT, token); // no-op if Phase 1 already created it
  const encodedPath = graphPathEncode(PROJECTS_INDEX_PATH);
  const blob = new Blob([JSON.stringify(projects, null, 2)], { type: "application/json" });
  await graphFetch(`/me/drive/root:/${encodedPath}:/content`, token, {
    method: "PUT",
    headers: { "Content-Type": "application/json" },
    body: blob
  });
}

// ---------- Cross-device capture text (captions/transcripts) ----------
// The captures themselves (photo/audio files) already reach OneDrive as
// real files via uploadToOneDrive, per-project — that's plenty for backup,
// but way too heavy to shuttle through a JSON index just so another device
// can read a caption. So this index carries only the lightweight text: no
// dataUrl. A device that pulls in a capture it didn't take locally gets a
// text-only "stub" record (see MossDB.captures.upsert / app.js) — enough
// for Ask AI and the reports to read, but with no photo/audio to display
// until that device's own capture is used (or the real file is fetched
// from OneDrive directly, which this does not do).

async function fetchCapturesIndex(token) {
  const encodedPath = graphPathEncode(CAPTURES_INDEX_PATH);
  const res = await fetch(`${GRAPH_BASE}/me/drive/root:/${encodedPath}:/content`, {
    headers: { Authorization: `Bearer ${token}` }
  });
  if (res.status === 404) return null;
  if (!res.ok) {
    const body = await res.text().catch(() => "");
    throw new Error(`Graph ${res.status} fetching captures index: ${body.slice(0, 200)}`);
  }
  try {
    return await res.json();
  } catch {
    return null;
  }
}

async function saveCapturesIndex(token, captures) {
  await ensureFolder(MOSS_ROOT, token);
  const encodedPath = graphPathEncode(CAPTURES_INDEX_PATH);
  // Strip dataUrl (and any other heavy fields) before it ever leaves this
  // device — only the text that makes a capture describable travels.
  const light = captures.map((c) => ({
    id: c.id,
    projectId: c.projectId,
    type: c.type,
    name: c.name || null,
    createdAt: c.createdAt,
    caption: c.caption || null,
    transcript: c.transcript || null
  }));
  const blob = new Blob([JSON.stringify(light, null, 2)], { type: "application/json" });
  await graphFetch(`/me/drive/root:/${encodedPath}:/content`, token, {
    method: "PUT",
    headers: { "Content-Type": "application/json" },
    body: blob
  });
}
