// Microsoft Graph calls — uploading captures into the OneDrive folder
// structure that already exists under "MOSS PROJECTS" (built in Phase 1).
// Everything here is best-effort: a failure here should never lose data
// that's already safe in MossDB/IndexedDB, only skip the OneDrive copy.

const GRAPH_BASE = "https://graph.microsoft.com/v1.0";
const MOSS_ACTIVE_ROOT = "MOSS PROJECTS/01 ACTIVE PROJECTS";

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
  await graphFetch(`/me/drive/root:/${graphPathEncode(parentPath)}:/children`, token, {
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
