// Moss AI Field Assistant — app shell + screens.
// Data currently lives in the browser (MossDB / IndexedDB, see js/db.js).
// Phase 2 swaps MossDB's internals for Microsoft Graph calls against
// OneDrive/SharePoint; screen code below does not need to change.

const $app = document.getElementById("app");
const $sheetBackdrop = document.getElementById("sheet-backdrop");
const $sheet = document.getElementById("sheet");
const $toast = document.getElementById("toast");

const TRADES = ["General", "HVAC", "Electrical", "Plumbing", "Framing", "Roofing", "Concrete", "Finishes"];

// User-typed text (issue titles, material names, inspector comments, etc.)
// gets inserted as HTML via template strings throughout this file, so it
// must be escaped first — otherwise typing something like `<img src=x
// onerror=...>` into a form field would inject and run as markup.
function escapeHtml(str) {
  return String(str ?? "").replace(/[&<>"']/g, (c) => ({
    "&": "&amp;",
    "<": "&lt;",
    ">": "&gt;",
    '"': "&quot;",
    "'": "&#39;"
  }[c]));
}

function toast(msg) {
  $toast.textContent = msg;
  $toast.classList.add("show");
  setTimeout(() => $toast.classList.remove("show"), 1800);
}

// If a sheet is open and gets cancelled (backdrop tap, or the user
// navigates away via hashchange/back button) without an explicit choice,
// this runs so any code `await`-ing a decision (e.g. pickProjectIfNeeded)
// resolves instead of hanging forever.
let onSheetCancelled = null;

function closeSheet() {
  $sheetBackdrop.classList.add("hidden");
  $sheet.innerHTML = "";
  if (onSheetCancelled) {
    const cb = onSheetCancelled;
    onSheetCancelled = null;
    cb();
  }
}

function openSheet(html, onCancel) {
  $sheet.innerHTML = html;
  $sheetBackdrop.classList.remove("hidden");
  onSheetCancelled = onCancel || null;
}

$sheetBackdrop.addEventListener("click", (e) => {
  if (e.target === $sheetBackdrop) closeSheet();
});

// ---------- Router ----------

function currentRoute() {
  const hash = location.hash.replace(/^#\/?/, "");
  const parts = hash.split("/").filter(Boolean);
  if (parts.length === 0) return { name: "home" };
  // decode: the link is built with encodeURIComponent(p.id) so ids with
  // spaces/unicode/special characters (real OneDrive folder names, once
  // Phase 2 lands) round-trip correctly instead of silently 404ing to Home.
  if (parts[0] === "project" && parts[1]) return { name: "project", id: decodeURIComponent(parts[1]) };
  if (parts[0] === "projects") return { name: "projects" };
  if (parts[0] === "issues") return { name: "issues" };
  if (parts[0] === "ask") return { name: "ask" };
  if (parts[0] === "settings") return { name: "settings" };
  return { name: "home" };
}

window.addEventListener("hashchange", () => {
  closeSheet(); // don't leave a sheet open over whatever screen navigation just landed on
  render();
});

window.addEventListener("DOMContentLoaded", async () => {
  try {
    // Fire-and-forget: resolves the redirect-back leg of Microsoft sign-in
    // as early as possible, whichever screen the user lands back on.
    // renderSettings() awaits the same (idempotent) promise before it
    // needs the result, so this doesn't need to block boot.
    initMsal();
    await MossDB.seedIfEmpty();
    // Ask the browser not to throw our data away when the phone is low on space.
    try { navigator.storage?.persist?.(); } catch {}
    const recovered = await recoverOrphanIssues();
    await render();
    if (recovered) toast(`Recovered ${recovered} issue${recovered === 1 ? "" : "s"} from saved voice notes`);
  } catch (err) {
    // IndexedDB can be unavailable (Safari private browsing, storage
    // disabled by device policy, quota errors) — without this, the app
    // silently shows a blank white screen with no explanation.
    console.error("Moss failed to start:", err);
    $app.innerHTML = `
      <div style="padding:40px 24px; text-align:center; display:flex; flex-direction:column; gap:12px; align-items:center;">
        <div style="font-size:32px;">⚠️</div>
        <div style="font-size:16px; font-weight:700;">Moss couldn't start</div>
        <div style="font-size:13px; color:#6B7280; max-width:280px;">
          This usually means private/incognito browsing or a storage restriction is blocking local storage on this device. Try a normal browser window, or check your browser's site data settings.
        </div>
      </div>
    `;
  }
  if ("serviceWorker" in navigator) {
    navigator.serviceWorker.register("sw.js").catch(() => {});
  }

  // Best-effort — runs after the first paint so it never blocks startup,
  // and re-renders only if it actually pulled in something new (a
  // project created on another device, or a status change made there).
  syncProjectsWithOneDrive().then((changed) => {
    if (changed) render();
  });
  syncCapturesWithOneDrive().then((changed) => {
    if (changed && currentRoute().name === "project") render();
  });
});

async function render() {
  const route = currentRoute();
  if (route.name === "home") return renderHome();
  if (route.name === "projects") return renderProjectsList();
  if (route.name === "project") return renderDashboard(route.id);
  if (route.name === "issues") return renderAllIssues();
  if (route.name === "ask") return renderAsk();
  if (route.name === "settings") return renderSettings();
}

function shell({ header, body, activeTab }) {
  $app.innerHTML = `
    ${header}
    <div class="view">${body}</div>
    ${tabbar(activeTab)}
  `;
}

function tabbar(active) {
  const tabs = [
    { id: "home", emoji: "🏠", label: "Home", href: "#/" },
    { id: "projects", emoji: "🏗️", label: "Projects", href: "#/projects" },
    { id: "ask", emoji: "🤖", label: "Ask AI", href: "#/ask" },
    { id: "settings", emoji: "⚙️", label: "Settings", href: "#/settings" }
  ];
  return `
    <div class="tabbar">
      ${tabs
        .map(
          (t) => `
        <a class="tab ${t.id === active ? "active" : ""}" href="${t.href}" style="text-decoration:none;">
          <span class="emoji">${t.emoji}</span>
          <span class="label">${t.label}</span>
        </a>`
        )
        .join("")}
    </div>
  `;
}

// ---------- Home ----------

function isActiveProject(p) {
  return p.status !== "Completed";
}

// Deleting a project doesn't erase its row — it flags it (a "tombstone"),
// the same way OneDrive sync already treats any edit: whichever device
// touched it last (via updatedAt) wins once synced. A real delete would
// otherwise just get silently undone next sync, by the other device's
// copy still being there and looking like "a project created elsewhere".
// isVisibleProject is what every screen filters through so a deleted
// project actually disappears everywhere despite still existing in
// IndexedDB (and, deliberately, in the OneDrive files themselves — see
// deleteProject below).
function isVisibleProject(p) {
  return !p.deleted;
}

async function renderHome() {
  const allProjects = (await MossDB.projects.all()).filter(isVisibleProject);
  const projects = allProjects.filter(isActiveProject);
  const visibleIds = new Set(allProjects.map((p) => p.id));
  const allIssues = (await MossDB.issues.all()).filter((i) => visibleIds.has(i.projectId));
  const openCount = allIssues.filter((i) => i.status === "Open").length;

  const header = `
    <div class="topbar">
      <div class="brand">
        <div class="mark">M</div>
        <div>
          <div class="name">MOSS</div>
          <div class="tag">AI Field Assistant</div>
        </div>
      </div>
    </div>
  `;

  const body = `
    <div>
      <div class="section-label">Today</div>
      <div class="card today-card" style="margin-top:10px;">
        <div>
          <div class="headline">${projects.length} active project${projects.length === 1 ? "" : "s"}</div>
          <div class="meta">${openCount} open issue${openCount === 1 ? "" : "s"}</div>
        </div>
        <div class="emoji">🏗️</div>
      </div>
    </div>

    <div>
      <div class="section-label">Quick Capture</div>
      <div class="grid-3" style="margin-top:10px;">
        <button class="tile-btn" data-capture="photo"><span class="emoji">📸</span><span class="label">Photo</span></button>
        <button class="tile-btn" data-capture="voice"><span class="emoji">🎤</span><span class="label">Voice Note</span></button>
        <button class="tile-btn" data-capture="issue"><span class="emoji">📋</span><span class="label">New Issue</span></button>
        <button class="tile-btn" data-capture="document"><span class="emoji">📄</span><span class="label">Document</span></button>
        <button class="tile-btn" data-capture="material"><span class="emoji">📦</span><span class="label">Material</span></button>
        <button class="tile-btn" data-capture="inspection"><span class="emoji">🏛️</span><span class="label">Inspection</span></button>
      </div>
    </div>

    <div>
      <div class="section-label">AI Actions</div>
      <div class="card card-list" style="margin-top:10px;">
        <button class="row" data-action="ask"><span class="icon">🤖</span><span class="main"><span class="title">Ask AI</span></span><span class="chev">›</span></button>
        <button class="row" data-action="open-issues"><span class="icon">📋</span><span class="main"><span class="title">Open Issues</span></span><span class="badge">${openCount}</span></button>
        <button class="row" data-action="next-steps"><span class="icon">➡️</span><span class="main"><span class="title">Next Steps</span></span><span class="chev">›</span></button>
        <button class="row" data-action="daily-report"><span class="icon">📝</span><span class="main"><span class="title">Daily Report</span></span><span class="chev">›</span></button>
        <button class="row" data-action="weekly-report"><span class="icon">📊</span><span class="main"><span class="title">Weekly Report</span></span><span class="chev">›</span></button>
      </div>
    </div>

    <div>
      <div class="section-label">
        Projects
        <span>
          <button class="link" data-action="new-project">+ Add</button>
          ${allProjects.some((p) => !isActiveProject(p)) ? `<a class="link" href="#/projects" style="text-decoration:none; margin-left:10px;">Completed ›</a>` : ""}
        </span>
      </div>
      <div style="display:flex; flex-direction:column; gap:8px; margin-top:10px;">
        ${projects.length ? projects.map(projectRow).join("") : `<div class="empty">No active projects. Tap "+ Add" to create one.</div>`}
      </div>
    </div>
  `;

  shell({ header, body, activeTab: "home" });
  wireQuickCapture();
  wireHomeActions(allIssues);
  wireProjectLinks();
  wireNewProjectButton();
}

function projectRow(p) {
  const initials = escapeHtml(p.name.slice(0, 2).toUpperCase());
  const completed = !isActiveProject(p);
  return `
    <a href="#/project/${encodeURIComponent(p.id)}" class="card" style="text-decoration:none; color:inherit; display:flex; align-items:center; gap:12px; padding:14px 16px; ${completed ? "opacity:.7;" : ""}">
      <div style="width:36px;height:36px;border-radius:10px;background:${completed ? "var(--ink-faint)" : "var(--ink)"};color:#fff;display:flex;align-items:center;justify-content:center;font-size:13px;font-weight:700;">${initials}</div>
      <div style="flex-grow:1; display:flex; flex-direction:column;">
        <span style="font-size:15px; font-weight:600;">${escapeHtml(p.name)}</span>
        <span style="font-size:12px; color:var(--ink-soft);">${escapeHtml(p.address || "")}</span>
      </div>
      ${completed ? `<span class="badge green">Completed</span>` : ""}
      <span style="color:var(--ink-faint); font-size:14px;">›</span>
    </a>
  `;
}

function wireProjectLinks() {
  // links are plain <a href="#/..."> — router picks up hashchange automatically
}

function wireNewProjectButton() {
  $app.querySelectorAll('[data-action="new-project"]').forEach((btn) => {
    btn.addEventListener("click", () => openNewProjectSheet());
  });
}

// ---------- Cross-device project sync ----------
// The project list lives in IndexedDB per device (see db.js), but a
// shared JSON file in OneDrive (graph.js's fetchProjectsIndex/
// saveProjectsIndex) lets every device signed into the same account see
// the same projects. This pulls the remote list, merges it with what's
// local (last-write-wins per project via updatedAt), writes the merged
// set back to both places, and reports whether anything actually changed
// so callers know whether to re-render.
let projectSyncInFlight = null;

async function syncProjectsWithOneDrive() {
  if (!msalConfigured()) return false;
  if (projectSyncInFlight) return projectSyncInFlight;

  projectSyncInFlight = (async () => {
    await initMsal();
    if (!msCurrentAccount()) return false;
    const token = await msGetToken();
    if (!token) return false;

    const [local, remote] = await Promise.all([MossDB.projects.all(), fetchProjectsIndex(token)]);

    const byId = new Map(local.map((p) => [p.id, p]));
    let changed = false;
    if (remote) {
      for (const rp of remote) {
        const lp = byId.get(rp.id);
        if (!lp) {
          byId.set(rp.id, rp); // a project created on another device
          changed = true;
        } else if (new Date(rp.updatedAt || 0) > new Date(lp.updatedAt || 0)) {
          byId.set(rp.id, rp); // remote has a newer edit (status change, etc.)
          changed = true;
        }
      }
    }

    const merged = [...byId.values()];
    for (const p of merged) await MossDB.projects.upsert(p);
    await saveProjectsIndex(token, merged);
    return changed;
  })();

  try {
    return await projectSyncInFlight;
  } catch (err) {
    console.error("Project sync failed", err);
    return false;
  } finally {
    projectSyncInFlight = null;
  }
}

// Syncs only the TEXT that makes a capture describable to AI (a photo's
// caption, a voice note's transcript) across devices — never the photo or
// audio file itself, which is too heavy for a JSON index and already has
// its own backup path (uploadToOneDrive, per capture, at the time it's
// taken). A capture pulled in from another device that this device never
// took locally becomes a text-only "stub": no thumbnail/audio to show
// here, but Ask AI and the reports can still read what it says. Mirrors
// syncProjectsWithOneDrive's shape.
let captureSyncInFlight = null;

async function syncCapturesWithOneDrive() {
  if (!msalConfigured()) return false;
  if (captureSyncInFlight) return captureSyncInFlight;

  captureSyncInFlight = (async () => {
    await initMsal();
    if (!msCurrentAccount()) return false;
    const token = await msGetToken();
    if (!token) return false;

    const [allLocal, remote] = await Promise.all([MossDB.captures.allRaw(), fetchCapturesIndex(token)]);
    // Scoped to photo/voice on purpose: those are the only capture types
    // this app writes a caption/transcript onto. Material and inspection
    // captures carry other fields (status, result, comments) that this
    // lightweight text index doesn't carry — syncing those in as stubs
    // would create records with those fields silently missing.
    const local = allLocal.filter((c) => c.type === "photo" || c.type === "voice");
    const byId = new Map(local.map((c) => [c.id, c]));
    let changed = false;

    if (remote) {
      for (const rc of remote.filter((c) => c.type === "photo" || c.type === "voice")) {
        const lc = byId.get(rc.id);
        if (rc.deleted) {
          // Deleted on another device: delete it here too (keeping only the
          // tombstone, so it is never brought back by a later sync).
          if (!lc) {
            const tomb = { id: rc.id, projectId: rc.projectId, type: rc.type, createdAt: rc.createdAt, deleted: true, deletedAt: new Date().toISOString() };
            byId.set(rc.id, tomb);
            await MossDB.captures.upsert(tomb);
            changed = true;
          } else if (!lc.deleted) {
            await MossDB.captures.remove(lc.id);
            byId.set(lc.id, { id: lc.id, projectId: lc.projectId, type: lc.type, createdAt: lc.createdAt, deleted: true });
            changed = true;
          }
          continue;
        }
        if (lc && lc.deleted) continue; // deleted here: stays deleted, and the merged index will tell other devices
        if (!lc) {
          // A capture taken on another device — save what we can (text
          // only; remoteOnly marks it so the UI doesn't try to show a
          // photo/audio player that has nothing to play).
          const stub = { ...rc, remoteOnly: true };
          byId.set(rc.id, stub);
          await MossDB.captures.upsert(stub);
          changed = true;
        } else {
          // Already have this capture locally (maybe with the real file).
          // Only ever fill in caption/transcript if this device doesn't
          // have one yet — these are set once at capture time and never
          // edited after, so there's no "newer wins" case to handle.
          const patch = {};
          if (rc.caption && !lc.caption) patch.caption = rc.caption;
          if (rc.transcript && !lc.transcript) patch.transcript = rc.transcript;
          if (rc.translation && !lc.translation) {
            patch.translation = rc.translation;
            patch.translationLang = rc.translationLang;
          }
          if (Object.keys(patch).length) {
            const updated = await MossDB.captures.update(lc.id, patch);
            if (updated) byId.set(lc.id, updated);
            changed = true;
          }
        }
      }
    }

    const merged = [...byId.values()];
    await saveCapturesIndex(token, merged);
    return changed;
  })();

  try {
    return await captureSyncInFlight;
  } catch (err) {
    console.error("Capture sync failed", err);
    return false;
  } finally {
    captureSyncInFlight = null;
  }
}

// Fetches the real photo/audio file for a text-only stub (one pulled in by
// syncCapturesWithOneDrive from another device) straight from OneDrive,
// using the exact filename it was uploaded under (remoteFileName), and
// saves it into this device's own copy of the record — so it only ever
// needs fetching once per device. Silently gives up if OneDrive isn't
// connected, the file was never actually uploaded, etc.: the caption/
// transcript text is still there either way, this only adds the media.
const hydrateAttempted = new Set();

async function hydrateRemoteCapture(capture) {
  if (!capture.remoteFileName || !msalConfigured()) return null;
  try {
    await initMsal();
    if (!msCurrentAccount()) return null;
    const token = await msGetToken();
    if (!token) return null;
    const project = await MossDB.projects.get(capture.projectId);
    if (!project) return null;
    const blob = await downloadCaptureFile(project.name, syncSubfolderFor(capture.type), capture.remoteFileName, token);
    const dataUrl = await blobToDataUrl(blob);
    return await MossDB.captures.update(capture.id, { dataUrl });
  } catch (err) {
    console.warn("Couldn't fetch capture file from OneDrive:", err.message);
    return null;
  }
}

// Best-effort: moves the project's OneDrive folder into a Trash folder
// there (see moveProjectToTrash in graph.js) instead of leaving it in
// Active Projects. Never blocks or reverses the app-side delete — if
// OneDrive isn't connected, or the move fails for any reason, the project
// is still gone from Moss either way; this only tidies up OneDrive.
async function trashProjectFolder(projectName) {
  if (!msalConfigured() || !msCurrentAccount()) return;
  try {
    const token = await msGetToken();
    if (!token) return;
    await moveProjectToTrash(projectName, token);
  } catch (err) {
    console.error("Couldn't move project folder to OneDrive Trash:", err);
  }
}

// Asks for confirmation, then soft-deletes a project (see isVisibleProject
// for why it's a flag and not a real removal) and syncs that deletion out
// to OneDrive/other devices. The project's own files (photos, voice notes,
// documents) aren't deleted — trashProjectFolder relocates that whole
// OneDrive folder into MOSS PROJECTS/99 TRASH rather than removing it.
function confirmDeleteProject(project) {
  openSheet(`
    <h2>Delete "${escapeHtml(project.name)}"?</h2>
    <p class="empty" style="text-align:left; margin-top:-4px;">
      This removes it from Moss on every device you sync with. Its OneDrive folder moves to MOSS PROJECTS ▸ 99 TRASH rather than being deleted, in case you need anything from it later.
    </p>
    <div class="sheet-actions">
      <button class="btn ghost" id="cancel-delete">Cancel</button>
      <button class="btn primary" id="confirm-delete" style="background:var(--red, #DC2626);">Delete Project</button>
    </div>
  `);
  document.getElementById("cancel-delete").addEventListener("click", closeSheet);
  document.getElementById("confirm-delete").addEventListener("click", async () => {
    closeSheet();
    await MossDB.projects.update(project.id, { deleted: true });
    syncProjectsWithOneDrive();
    trashProjectFolder(project.name);
    toast(`"${project.name}" deleted`);
    location.hash = "#/projects";
  });
}

// Full-size photo view — tapping a thumbnail in Recent Captures opens this
// instead of leaving the picture squeezed into a small grid tile.
function openPhotoViewer(photo) {
  openSheet(`
    <div class="photo-viewer">
      <img src="${photo.dataUrl}" alt="${photo.caption ? escapeHtml(photo.caption) : ""}">
      ${
        photo.caption
          ? `<p class="empty" style="text-align:left; margin-top:10px;">${escapeHtml(photo.caption)}</p>`
          : ""
      }
    </div>
    ${closeButton()}
  `);
  wireCloseButton();
}

async function openNewProjectSheet() {
  const existing = await MossDB.projects.all();
  const existingIds = new Set(existing.map((p) => p.id));

  openSheet(`
    <h2>New Project</h2>
    <div class="field">
      <label>Project name</label>
      <input id="f-name" placeholder='e.g. "Thompson Residence"'>
    </div>
    <div class="field">
      <label>Address / description</label>
      <input id="f-address" placeholder='e.g. "412 Maple St · Kitchen remodel"'>
    </div>
    <div class="sheet-actions">
      <button class="btn ghost" id="cancel">Cancel</button>
      <button class="btn primary" id="save">Create Project</button>
    </div>
  `);
  document.getElementById("cancel").addEventListener("click", closeSheet);
  document.getElementById("save").addEventListener("click", async () => {
    const name = document.getElementById("f-name").value.trim();
    if (!name) {
      toast("Give the project a name");
      return;
    }
    const address = document.getElementById("f-address").value.trim();

    // Slugify the name into an id, deduping against existing projects so
    // two "Smith" projects don't collide and silently overwrite each other.
    let base = name
      .toLowerCase()
      .replace(/[^a-z0-9]+/g, "-")
      .replace(/(^-|-$)/g, "") || "project";
    let id = base;
    let n = 2;
    while (existingIds.has(id)) {
      id = `${base}-${n}`;
      n++;
    }

    await MossDB.projects.add({ id, name, address });
    closeSheet();
    toast("Project created");
    render();

    // Best-effort: mirror the standard folder set into OneDrive so this
    // project has somewhere for captures to sync to, and push the updated
    // project list so it shows up on other devices too. Silent on
    // failure — the project still works locally either way.
    if (msalConfigured() && msCurrentAccount()) {
      try {
        const token = await msGetToken();
        if (token) await ensureProjectFolders(name, token);
      } catch (err) {
        console.error("Failed to create OneDrive folders for new project", err);
      }
      syncProjectsWithOneDrive();
    }
  });
}

function wireHomeActions(allIssues) {
  $app.querySelectorAll("[data-action]").forEach((btn) => {
    btn.addEventListener("click", () => {
      const action = btn.dataset.action;
      if (action === "ask") location.hash = "#/ask";
      else if (action === "open-issues") location.hash = "#/issues";
      else if (action === "next-steps") showNextSteps();
      else if (action === "daily-report") showDailyReport();
      else if (action === "weekly-report") showWeeklyReport();
    });
  });
}

// ---------- Projects list ----------

async function renderProjectsList() {
  const allProjects = (await MossDB.projects.all()).filter(isVisibleProject);
  const active = allProjects.filter(isActiveProject);
  const completed = allProjects.filter((p) => !isActiveProject(p));
  const header = `
    <div class="topbar">
      <div class="project-head">
        <div class="title" style="font-size:20px;">Projects</div>
        <button class="btn primary" style="flex:none; padding:8px 14px; font-size:13px;" data-action="new-project">+ Add Project</button>
      </div>
    </div>
  `;
  const body = `
    <div>
      <div class="section-label">Active (${active.length})</div>
      <div style="display:flex; flex-direction:column; gap:8px; margin-top:10px;">
        ${active.length ? active.map(projectRow).join("") : `<div class="empty">No active projects. Tap "+ Add Project" to create one.</div>`}
      </div>
    </div>
    ${
      completed.length
        ? `
    <div>
      <div class="section-label">Completed (${completed.length})</div>
      <div style="display:flex; flex-direction:column; gap:8px; margin-top:10px;">
        ${completed.map(projectRow).join("")}
      </div>
    </div>`
        : ""
    }
  `;
  shell({ header, body, activeTab: "projects" });
  wireNewProjectButton();
}

// ---------- Project dashboard ----------

async function renderDashboard(id) {
  const project = await MossDB.projects.get(id);
  if (!project || project.deleted) {
    location.hash = "#/";
    return;
  }
  const issues = await MossDB.issues.forProject(id);
  const captures = await MossDB.captures.forProject(id);
  // Newest first, so the compact view's single visible row is the latest one.
  const issueTime = (i) => String(i.recordedAt || i.createdAt || "");
  const openIssues = issues
    .filter((i) => i.status === "Open")
    .sort((a, b) => issueTime(b).localeCompare(issueTime(a)) || String(b.id).localeCompare(String(a.id)));
  const photos = captures.filter((c) => c.type === "photo");
  const inspections = captures
    .filter((c) => c.type === "inspection")
    .sort((a, b) => String(b.recordedAt || b.createdAt).localeCompare(String(a.recordedAt || a.createdAt)));
  const allMaterials = captures.filter((c) => c.type === "material");
  const materials = allMaterials.filter((c) => !c.archived); // bought ones live in the drawer
  const drawerMaterialCount = allMaterials.length - materials.length;
  const drawerIssueCount = issues.filter((i) => i.status === "Completed").length;

  // Any capture on this project that's still a text-only stub from another
  // device (has a caption/transcript but no dataUrl) gets its real photo/
  // audio fetched from OneDrive in the background, once per device — see
  // hydrateRemoteCapture. Re-renders this same screen when one lands.
  for (const c of captures) {
    if (!c.dataUrl && c.remoteFileName && !hydrateAttempted.has(c.id)) {
      hydrateAttempted.add(c.id);
      hydrateRemoteCapture(c).then((updated) => {
        if (updated && currentRoute().name === "project" && currentRoute().id === id) {
          render();
        } else if (!updated) {
          // Didn't work this time (maybe the source device hadn't finished
          // uploading yet, maybe a network hiccup) — un-mark it so the
          // next time this screen opens, it gets another try instead of
          // being stuck "unavailable" for the rest of the session.
          hydrateAttempted.delete(c.id);
        }
      });
    }
  }
  const initials = escapeHtml(project.name.slice(0, 2).toUpperCase());
  const completed = !isActiveProject(project);

  const header = `
    <div class="topbar">
      <a class="back-link" href="#/projects"><span>‹</span><span>All Projects</span></a>
      <div class="project-head">
        <div>
          <div class="title">${escapeHtml(project.name)}${completed ? ` <span class="badge green" style="vertical-align:middle;">Completed</span>` : ""}</div>
          <div class="sub">${escapeHtml(project.address || "")}</div>
        </div>
        <div class="avatar">${initials}</div>
      </div>
    </div>
  `;

  const body = `
    <div class="stat-grid">
      <div class="stat-tile"><span class="num" style="color:var(--green);">${issues.filter((i) => i.status === "Completed").length}</span><span class="label">COMPLETED</span></div>
      <div class="stat-tile warn"><span class="num" style="color:var(--amber-deep);">${openIssues.length}</span><span class="label">OPEN ISSUES</span></div>
      <div class="stat-tile"><span class="num">${issues.filter((i) => i.status === "Waiting").length}</span><span class="label">WAITING</span></div>
      <div class="stat-tile"><span class="num">${captures.filter((c) => c.type === "inspection").length}</span><span class="label">INSPECTIONS</span></div>
      <div class="stat-tile"><span class="num">${materials.length}</span><span class="label">MATERIALS</span></div>
      <div class="stat-tile"><span class="num">${new Set(openIssues.map((i) => i.trade)).size}</span><span class="label">TRADES ON OPEN</span></div>
    </div>

    <div>
      <div class="section-label">Quick Capture</div>
      <div class="grid-3" style="margin-top:10px;">
        <button class="tile-btn" data-capture="photo"><span class="emoji">📸</span><span class="label">Photo</span></button>
        <button class="tile-btn" data-capture="voice"><span class="emoji">🎤</span><span class="label">Voice Note</span></button>
        <button class="tile-btn" data-capture="issue"><span class="emoji">📋</span><span class="label">New Issue</span></button>
        <button class="tile-btn" data-capture="document"><span class="emoji">📄</span><span class="label">Document</span></button>
        <button class="tile-btn" data-capture="material"><span class="emoji">📦</span><span class="label">Material</span></button>
        <button class="tile-btn" data-capture="inspection"><span class="emoji">🏛️</span><span class="label">Inspection</span></button>
      </div>
    </div>

    <div>
      <div class="section-label">Open Issues<span><span class="link" data-quick="issue-pdf">📄 PDF</span><span class="link" data-quick="issue" style="margin-left:14px;">+ Add</span></span></div>
      <div class="card card-list" style="margin-top:10px;" data-acc="issues" data-acc-total="${openIssues.length}">
        ${
          openIssues.length
            ? openIssues
                .map(
                  (i) => `
          <div class="row" data-issue-id="${escapeHtml(i.id)}" style="cursor:pointer;">
            <span class="icon" style="color:var(--red);">●</span>
            <span class="main"><span class="title">${escapeHtml(i.title)}</span><span class="desc">${escapeHtml(i.trade)}${i.requirement ? " · " + escapeHtml(i.requirement.split("\n")[0]) : ""}</span></span>
            <span class="chev" style="white-space:nowrap;">${i.photoId ? "📷 " : ""}›</span>
          </div>`
                )
                .join("")
            : `<div class="row"><span class="empty">No open issues yet — capture one from the field.</span></div>`
        }
        <div class="row" data-drawer="issue" style="cursor:pointer;">
          <span class="icon">🗄️</span>
          <span class="main"><span class="title">Drawer</span><span class="desc">Fixed issues · ${drawerIssueCount}</span></span>
          <span class="chev">›</span>
        </div>
      </div>
    </div>

    <div>
      <div class="section-label">Materials<span><span class="link" data-quick="material-pdf">📄 PDF</span><span class="link" data-quick="material" style="margin-left:14px;">+ Add</span></span></div>
      <div class="card card-list" style="margin-top:10px;" data-acc="materials" data-acc-total="${materials.length}">
        ${
          materials.length
            ? materials
                .slice(-10)
                .reverse()
                .map(
                  (m) => `
          <div class="row" data-material-id="${escapeHtml(m.id)}" style="cursor:pointer;">
            <span class="icon">📦</span>
            <span class="main"><span class="title">${escapeHtml(m.name)}</span><span class="desc">${[m.dimensions, m.quantity ? "Qty " + m.quantity : "", m.status || ""].filter(Boolean).map(escapeHtml).join(" · ")}</span></span>
            <span class="chev" style="white-space:nowrap;">${m.photoId ? "📷 " : ""}›</span>
          </div>`
                )
                .join("") + (materials.length > 10 ? `<div class="row"><span class="empty">Showing the 10 most recent of ${materials.length}. The PDF list includes all of them.</span></div>` : "")
            : `<div class="row"><span class="empty">No materials yet — snap a photo of one to start a shopping list.</span></div>`
        }
        <div class="row" data-drawer="material" style="cursor:pointer;">
          <span class="icon">🗄️</span>
          <span class="main"><span class="title">Drawer</span><span class="desc">Bought materials · ${drawerMaterialCount}</span></span>
          <span class="chev">›</span>
        </div>
      </div>
    </div>

    <div>
      <div class="section-label">Inspections<span><span class="link" data-quick="inspection-pdf">📄 PDF</span><span class="link" data-quick="inspection" style="margin-left:14px;">+ Add</span></span></div>
      <div class="card card-list" style="margin-top:10px;" data-acc="inspections" data-acc-total="${inspections.length}">
        ${
          inspections.length
            ? inspections
                .slice(0, 10)
                .map((i) => {
                  const info = inspectionResultInfo(i.result);
                  return `
          <div class="row" data-inspection-id="${escapeHtml(i.id)}" style="cursor:pointer;">
            <span class="icon">${info.icon}</span>
            <span class="main"><span class="title">${escapeHtml(i.name || "Inspection")}</span><span class="desc">${[info.short, i.recordedAt || i.createdAt ? new Date(i.recordedAt || i.createdAt).toLocaleDateString() : "", i.inspector].filter(Boolean).map(escapeHtml).join(" · ")}</span></span>
            <span class="chev" style="white-space:nowrap;">${(i.photoIds && i.photoIds.length) ? "📷 " : ""}›</span>
          </div>`;
                })
                .join("") + (inspections.length > 10 ? `<div class="row"><span class="empty">Showing the 10 most recent of ${inspections.length}. The PDF includes all of them.</span></div>` : "")
            : `<div class="row"><span class="empty">No inspections yet — tap + Add to record one with the inspector.</span></div>`
        }
      </div>
    </div>

    <div>
      <div class="section-label">Recent Captures${photos.length ? `<span><span class="link" data-manage="photo">🗑 Clear</span></span>` : ""}</div>
      <div class="thumb-grid" style="margin-top:10px;" data-acc="photos" data-acc-total="${photos.length}">
        ${
          photos.length
            ? photos
                .slice(-4)
                .reverse()
                .map((p) => {
                  // A capture synced in from another device as text-only
                  // (see syncCapturesWithOneDrive) has no dataUrl here —
                  // there's no image to show, only whatever caption came
                  // with it, so render that instead of a broken <img>, and
                  // skip making it clickable (nothing to enlarge yet).
                  if (!p.dataUrl) {
                    return `
          <div class="thumb">
            <div class="thumb-img" style="display:flex; align-items:center; justify-content:center; text-align:center; padding:8px;">
              <span class="desc" style="font-size:11px; line-height:1.3;">Photo from another device</span>
            </div>
            ${p.caption ? `<span class="desc" style="font-size:12px; line-height:1.35;">${escapeHtml(p.caption)}</span>` : ""}
          </div>`;
                  }
                  return `
          <div class="thumb clickable" data-photo-id="${p.id}">
            <div class="thumb-img"><img src="${p.dataUrl}" alt="${p.caption ? escapeHtml(p.caption) : ""}"></div>
            ${
              p.caption
                ? `<span class="desc" style="font-size:12px; line-height:1.35;">${escapeHtml(p.caption)}</span>`
                : aiCaptionsEnabled()
                ? `<span class="desc" style="font-size:12px; opacity:.6;">Describing…</span>`
                : ""
            }
          </div>`;
                })
                .join("")
            : `<div class="empty">No photos captured for this project yet.</div>`
        }
      </div>
    </div>

    <div>
      <div class="section-label">Voice Notes${captures.some((c) => c.type === "voice") ? `<span><span class="link" data-manage="voice">🗑 Clear</span></span>` : ""}</div>
      <div class="card card-list" style="margin-top:10px;" data-acc="voice" data-acc-total="${captures.filter((c) => c.type === "voice").length}">
        ${
          captures.filter((c) => c.type === "voice").length
            ? captures
                .filter((c) => c.type === "voice")
                .slice(-10)
                .reverse()
                .map(
                  (v) => `
          <div class="row voice-row" style="cursor:default; flex-direction:column; align-items:stretch; gap:8px;">
            <span class="desc">${escapeHtml(new Date(v.createdAt).toLocaleString())}</span>
            ${
              v.dataUrl
                ? `<audio controls preload="none" style="width:100%; height:32px;" src="${v.dataUrl}"></audio>`
                : `<span class="desc" style="opacity:.6;">Recorded on another device — audio not available here.</span>`
            }
            ${
              v.transcript
                ? `<span class="desc" style="font-style:italic;">"${escapeHtml(v.transcript)}"</span>`
                : `<span class="desc" style="opacity:.6;">No transcript (speech-to-text wasn't available when this was recorded).</span>`
            }
            ${v.translation ? `<span class="desc">🌐 ${escapeHtml(translationLabel(v.translationLang))}: ${escapeHtml(v.translation)}</span>` : ""}
            ${v.transcript && aiConfigured() ? `<button class="btn ghost" data-retr-voice="${escapeHtml(v.id)}" style="padding:6px 10px; font-size:12px; align-self:flex-start;">🌐 ${v.translation ? "Translate again" : "Translate"}</button>` : ""}
          </div>`
                )
                .join("")
            : `<div class="row"><span class="empty">No voice notes for this project yet.</span></div>`
        }
      </div>
    </div>

    <div>
      <div class="section-label">Daily Logs</div>
      <div class="card card-list" style="margin-top:10px;" id="log-list"></div>
    </div>

    <div class="card card-list">
      <button class="row" id="toggle-project-status">
        <span class="icon">${completed ? "↩️" : "✅"}</span>
        <span class="main"><span class="title">${completed ? "Reactivate Project" : "Mark Project Complete"}</span></span>
      </button>
    </div>

    <div class="card card-list">
      <button class="row" id="delete-project" style="color:var(--red, #DC2626);">
        <span class="icon">🗑️</span>
        <span class="main"><span class="title">Delete Project</span></span>
      </button>
    </div>
  `;

  shell({ header, body, activeTab: "projects" });
  applyDashboardAccordions();

  wireQuickCapture();
  $app.querySelector('[data-quick="issue"]').addEventListener("click", () => openIssueSheet(id));
  $app.querySelector('[data-quick="issue-pdf"]').addEventListener("click", () => openIssuesPdfSheet(id));
  $app.querySelector('[data-quick="material"]').addEventListener("click", () => openMaterialSheet(id));
  $app.querySelector('[data-quick="material-pdf"]').addEventListener("click", () => openMaterialsPdfSheet(id));
  $app.querySelector('[data-quick="inspection"]').addEventListener("click", () => openInspectionSheet(id));
  $app.querySelector('[data-quick="inspection-pdf"]').addEventListener("click", () => openInspectionsPdfSheet(id));
  $app.querySelectorAll("[data-inspection-id]").forEach((el) => {
    el.addEventListener("click", () => {
      const insp = inspections.find((i) => i.id === el.dataset.inspectionId);
      if (insp) openInspectionDetail(insp, captures);
    });
  });
  $app.querySelectorAll("[data-manage]").forEach((el) => {
    el.addEventListener("click", () => openCaptureManager(el.dataset.manage, id));
  });
  $app.querySelectorAll("[data-drawer]").forEach((el) => {
    el.addEventListener("click", () => openDrawerSheet(el.dataset.drawer, id));
  });
  $app.querySelectorAll("[data-material-id]").forEach((el) => {
    el.addEventListener("click", () => {
      const material = materials.find((m) => m.id === el.dataset.materialId);
      if (material) openMaterialDetail(material, captures);
    });
  });
  $app.querySelector("#toggle-project-status").addEventListener("click", async () => {
    await MossDB.projects.update(id, { status: completed ? "Active" : "Completed" });
    syncProjectsWithOneDrive();
    toast(completed ? "Project reactivated" : "Project marked complete");
    renderDashboard(id);
  });
  $app.querySelector("#delete-project").addEventListener("click", () => confirmDeleteProject(project));
  $app.querySelectorAll("[data-issue-id]").forEach((el) => {
    el.addEventListener("click", () => {
      const issue = issues.find((i) => i.id === el.dataset.issueId);
      if (issue) openIssueDetail(issue, captures);
    });
  });
  $app.querySelectorAll("[data-retr-voice]").forEach((el) => {
    const v = captures.find((c) => c.id === el.dataset.retrVoice);
    if (!v) return;
    el.addEventListener("click", async () => {
      const label = el.textContent;
      el.disabled = true;
      el.textContent = "🌐 Translating…";
      const tr = await tryTranslate(v.transcript, true, v.translation || "");
      if (!tr) {
        toast("Couldn't translate — try again");
        el.disabled = false;
        el.textContent = label;
        return;
      }
      const patch = { translation: tr.translation, translationLang: tr.target };
      await MossDB.captures.update(v.id, patch);
      // A voice note that belongs to an issue or material: keep that record's copy in step.
      if (v.issueId) {
        const issue = (await MossDB.issues.all()).find((i) => i.id === v.issueId);
        if (issue) await MossDB.issues.add({ ...issue, ...patch });
      }
      if (v.materialId) await MossDB.captures.update(v.materialId, patch);
      toast("Translation updated");
      syncCapturesWithOneDrive();
      render();
    });
  });
  $app.querySelectorAll("[data-photo-id]").forEach((el) => {
    el.addEventListener("click", () => {
      const photo = photos.find((p) => p.id === el.dataset.photoId);
      if (photo) openPhotoViewer(photo);
    });
  });

  const logs = await MossDB.logs.forProject(id);
  const $logList = document.getElementById("log-list");
  $logList.innerHTML = logs.length
    ? logs
        .slice(-5)
        .reverse()
        .map(
          (l) => `<div class="row"><span class="icon">📝</span><span class="main"><span class="title">${escapeHtml(new Date(l.createdAt).toLocaleDateString())}</span><span class="desc">${escapeHtml(l.summary || "")}</span></span></div>`
        )
        .join("")
    : `<div class="row"><span class="empty">No daily logs yet — run "Daily Report" from Home to log today.</span></div>`;
}

// ---------- Open issues (all projects) ----------

async function renderAllIssues() {
  const projects = (await MossDB.projects.all()).filter(isVisibleProject);
  const visibleIds = new Set(projects.map((p) => p.id));
  const issues = (await MossDB.issues.all()).filter((i) => i.status === "Open" && visibleIds.has(i.projectId));
  const nameOf = (id) => projects.find((p) => p.id === id)?.name || id;

  const header = `
    <div class="topbar">
      <a class="back-link" href="#/"><span>‹</span><span>Home</span></a>
      <div class="project-head">
        <div class="title" style="font-size:20px;">Open Issues</div>
        <button class="btn primary" style="flex:none; padding:8px 14px; font-size:13px;" id="issues-pdf">📄 PDF</button>
      </div>
    </div>
  `;
  const body = `
    <div class="card card-list">
      ${
        issues.length
          ? issues
              .map(
                (i) => `
        <div class="row" style="cursor:default;">
          <span class="icon" style="color:var(--red);">●</span>
          <span class="main"><span class="title">${escapeHtml(i.title)}</span><span class="desc">${escapeHtml(nameOf(i.projectId))} · ${escapeHtml(i.trade)}</span></span>
        </div>`
              )
              .join("")
          : `<div class="row"><span class="empty">Nothing open right now.</span></div>`
      }
    </div>
  `;
  shell({ header, body, activeTab: "home" });
  document.getElementById("issues-pdf").addEventListener("click", () => openIssuesPdfSheet(null));
}

// ---------- Ask AI ----------

// Builds a compact text summary of everything MossDB knows — projects,
// open/closed issues, daily log summaries, and capture counts — for
// Claude to answer questions against. Capture content itself (photo/voice
// data URLs) is deliberately left out: it would blow past a reasonable
// prompt size fast, and nothing here reads photos or transcribes audio
// yet, so including the raw data would just be noise Claude can't use.
async function buildAIContext() {
  const projects = (await MossDB.projects.all()).filter(isVisibleProject);
  const allIssues = await MossDB.issues.all();
  const lines = [];

  for (const p of projects) {
    lines.push(`## ${p.name} (${p.status || "Active"})${p.address ? ` — ${p.address}` : ""}`);

    const issues = allIssues.filter((i) => i.projectId === p.id);
    if (issues.length) {
      for (const i of issues) {
        lines.push(`- Issue [${i.status}] (${i.trade}): ${i.title}${i.requirement ? ` — ${i.requirement}` : ""}`);
      }
    }

    const logs = await MossDB.logs.forProject(p.id);
    for (const l of logs.slice(-10)) {
      lines.push(`- Log (${new Date(l.createdAt).toLocaleDateString()}): ${l.summary}`);
    }

    const captures = await MossDB.captures.forProject(p.id);
    if (captures.length) {
      const counts = {};
      for (const c of captures) counts[c.type] = (counts[c.type] || 0) + 1;
      lines.push(`- Captures on file: ${Object.entries(counts).map(([type, n]) => `${n} ${type}`).join(", ")}`);

      // Anything we actually have text for (a photo caption from AI, a
      // voice-note transcript) gets surfaced so Ask AI can answer
      // questions about what's IN a capture, not just that it exists.
      for (const m of captures.filter((c) => c.type === "material" && !c.archived).slice(-20)) {
        const bits = [m.dimensions, m.quantity ? `qty ${m.quantity}` : "", m.note].filter(Boolean).join("; ");
        lines.push(`- Material [${m.status || "?"}]: ${m.name}${bits ? ` — ${bits}` : ""}`);
      }

      for (const i of captures.filter((c) => c.type === "inspection").slice(-10)) {
        lines.push(`- Inspection [${inspectionResultInfo(i.result).short || "?"}]: ${i.name}${i.inspector ? ` (inspector ${i.inspector})` : ""}${i.comments ? ` — ${i.comments}` : ""}`);
      }

      const described = captures.filter((c) => c.caption || c.transcript);
      for (const c of described.slice(-15)) {
        const when = new Date(c.createdAt).toLocaleDateString();
        if (c.caption) lines.push(`- Photo (${when}): ${c.caption}`);
        if (c.transcript) lines.push(`- Voice note (${when}): "${c.transcript}"`);
      }
    }

    lines.push("");
  }

  return lines.join("\n") || "No projects yet.";
}

async function renderAsk() {
  const configured = aiConfigured();
  const header = `
    <div class="topbar">
      <div class="project-head"><div class="title" style="font-size:20px;">Ask AI</div></div>
    </div>
  `;
  const body = `
    <div class="card" style="padding:16px;">
      ${
        configured
          ? `
        <div class="field">
          <label>Ask about any project</label>
          <textarea id="ask-input" placeholder='e.g. "What is still open at Helm?"'></textarea>
        </div>
        <button class="btn primary" id="ask-submit" style="margin-top:12px; width:100%;">Ask</button>
        <div id="ask-answer" style="margin-top:14px; font-size:14px; line-height:1.6; white-space:pre-wrap;"></div>
      `
          : `<p class="empty">Add a Claude API key in Settings to turn this on.</p>`
      }
    </div>
  `;
  shell({ header, body, activeTab: "ask" });
  if (!configured) return;

  const $input = document.getElementById("ask-input");
  const $submit = document.getElementById("ask-submit");
  const $answer = document.getElementById("ask-answer");

  $submit.addEventListener("click", async () => {
    const question = $input.value.trim();
    if (!question) {
      toast("Type a question first");
      return;
    }
    $submit.disabled = true;
    $submit.textContent = "Asking…";
    $answer.textContent = "";
    try {
      const context = await buildAIContext();
      const prompt = `You are answering questions about active construction projects for a general contractor, using ONLY the project data below. If the data doesn't answer the question, say so plainly instead of guessing.\n\n${context}\n\nQuestion: ${question}`;
      const answer = await askClaude(prompt);
      $answer.textContent = answer || "(No answer returned.)";
    } catch (err) {
      $answer.textContent = `⚠️ ${err.message}`;
    } finally {
      $submit.disabled = false;
      $submit.textContent = "Ask";
    }
  });
}

// ---------- Settings ----------

// ---------- Accordions (compact dashboard lists) ----------
// Each long list on a project shows only its latest item plus a count; tap
// "Show all" to open the rest. On by default; Settings can turn it off.
const ACCORDION_PREF = "moss_accordion";
const ACCORDION_OPEN = "moss_accordion_open";

function accordionEnabled() {
  try {
    return localStorage.getItem(ACCORDION_PREF) !== "off";
  } catch {
    return true;
  }
}

function setAccordionEnabled(on) {
  try {
    localStorage.setItem(ACCORDION_PREF, on ? "on" : "off");
  } catch {}
}

function accordionOpenSet() {
  try {
    return new Set(JSON.parse(localStorage.getItem(ACCORDION_OPEN) || "[]"));
  } catch {
    return new Set();
  }
}

function saveAccordionOpen(set) {
  try {
    localStorage.setItem(ACCORDION_OPEN, JSON.stringify([...set]));
  } catch {}
}

const ACCORDION_KINDS = {
  issues: { rows: "[data-issue-id]", noun: "open issues", singular: "issue" },
  materials: { rows: "[data-material-id]", noun: "materials", singular: "material" },
  inspections: { rows: "[data-inspection-id]", noun: "inspections", singular: "inspection" },
  voice: { rows: ".voice-row", noun: "voice notes", singular: "voice note" },
  photos: { rows: ".thumb", noun: "photos", singular: "photo" }
};

function applyDashboardAccordions() {
  const on = accordionEnabled();
  const openSet = accordionOpenSet();
  $app.querySelectorAll("[data-acc]").forEach((box) => {
    const kind = ACCORDION_KINDS[box.dataset.acc];
    if (!kind) return;
    const total = Number(box.dataset.accTotal) || 0;
    // Count next to the section title ("Voice Notes · 6").
    const label = box.previousElementSibling;
    if (label && label.classList.contains("section-label") && total && label.firstChild && label.firstChild.nodeType === 3) {
      label.firstChild.textContent = label.firstChild.textContent.trim() + " · " + total;
    }
    if (!on) return;
    const rows = [...box.querySelectorAll(kind.rows)];
    if (rows.length < 2) return;
    const isOpen = openSet.has(box.dataset.acc);
    const toggle = document.createElement("div");
    toggle.className = box.classList.contains("thumb-grid") ? "" : "row";
    toggle.setAttribute("role", "button");
    toggle.dataset.accToggle = box.dataset.acc;
    toggle.style.cursor = "pointer";
    const paint = (open) => {
      rows.slice(1).forEach((r) => (r.style.display = open ? "" : "none"));
      const text = open ? "Show only the latest" : `Show all ${total} ${kind.noun}`;
      toggle.innerHTML = box.classList.contains("thumb-grid")
        ? `<span class="link" style="color:var(--amber-deep); font-weight:600; font-size:13px;">${open ? "▴" : "▾"} ${text}</span>`
        : `<span class="icon">${open ? "▴" : "▾"}</span><span class="main"><span class="title">${text}</span>${open ? "" : `<span class="desc">Showing the latest ${kind.singular} · tap to see the rest</span>`}</span>`;
      toggle.setAttribute("aria-expanded", String(open));
    };
    paint(isOpen);
    toggle.addEventListener("click", () => {
      const set = accordionOpenSet();
      const nowOpen = !set.has(box.dataset.acc);
      if (nowOpen) set.add(box.dataset.acc);
      else set.delete(box.dataset.acc);
      saveAccordionOpen(set);
      paint(nowOpen);
    });
    if (box.classList.contains("thumb-grid")) {
      toggle.style.marginTop = "6px";
      box.after(toggle);
    } else {
      rows[0].after(toggle);
    }
  });
}

async function renderSettings() {
  await initMsal();
  const configured = msalConfigured();
  const account = msCurrentAccount();

  const oneDriveDesc = !configured
    ? "Setup needed — ask Claude to finish connecting your Azure app"
    : account
    ? `Connected as ${escapeHtml(account.username)}`
    : "Not connected";

  const header = `
    <div class="topbar">
      <div class="project-head"><div class="title" style="font-size:20px;">Settings</div></div>
    </div>
  `;
  const body = `
    <div class="card card-list">
      <div class="row">
        <span class="icon">☁️</span>
        <span class="main"><span class="title">OneDrive / SharePoint</span><span class="desc">${oneDriveDesc}</span></span>
        ${
          configured
            ? `<button class="btn ${account ? "ghost" : "primary"}" style="flex:none; padding:8px 14px; font-size:13px;" id="onedrive-toggle">${account ? "Disconnect" : "Connect"}</button>`
            : ""
        }
      </div>
      <div class="row" style="flex-direction:column; align-items:stretch; gap:8px;">
        <span class="main"><span class="icon">🤖</span> <span class="title">AI processing</span><span class="desc">${aiConfigured() ? "Claude API key saved on this device" : "Add a Claude API key to turn on Ask AI"}</span></span>
        <div style="display:flex; gap:8px;">
          <input id="f-ai-key" type="password" placeholder="sk-ant-..." style="flex:1;" value="${aiConfigured() ? "••••••••••••••••" : ""}">
          <button class="btn primary" id="ai-key-save" style="flex:none; padding:8px 14px; font-size:13px;">Save</button>
          ${aiConfigured() ? `<button class="btn ghost" id="ai-key-clear" style="flex:none; padding:8px 14px; font-size:13px;">Clear</button>` : ""}
        </div>
      </div>
      <div class="row">
        <span class="icon">🖼️</span>
        <span class="main"><span class="title">AI photo descriptions</span><span class="desc">${
          !aiConfigured()
            ? "Add a Claude API key above to use this"
            : aiCaptionsEnabled()
            ? "On — each new photo is described automatically"
            : "Off — photos are saved without an AI description"
        }</span></span>
        <input type="checkbox" class="switch" id="f-ai-captions" ${aiCaptionsEnabled() ? "checked" : ""} ${aiConfigured() ? "" : "disabled"} aria-label="AI photo descriptions">
      </div>
      <div class="row">
        <span class="icon">🌐</span>
        <span class="main"><span class="title">AI translate voice notes</span><span class="desc">${
          !aiConfigured()
            ? "Add a Claude API key above to use this"
            : aiTranslateEnabled()
            ? "On — Spanish notes get an English translation, English notes get Spanish"
            : "Off — voice notes are not translated"
        }</span></span>
        <input type="checkbox" class="switch" id="f-ai-translate" ${aiTranslateEnabled() ? "checked" : ""} ${aiConfigured() ? "" : "disabled"} aria-label="AI translate voice notes">
      </div>
      <div class="row">
        <span class="icon">🪜</span>
        <span class="main"><span class="title">Compact lists</span><span class="desc">${accordionEnabled() ? "On — each list shows its latest item; tap to open the rest" : "Off — lists are shown in full"}</span></span>
        <input type="checkbox" class="switch" id="f-accordion" ${accordionEnabled() ? "checked" : ""} aria-label="Compact lists">
      </div>
      <div class="row">
        <span class="icon">💾</span>
        <span class="main"><span class="title">Phone storage</span><span class="desc" id="storage-info">Checking…</span></span>
      </div>
      <div class="row" style="flex-direction:column; align-items:stretch; gap:8px;">
        <span class="main"><span class="icon">🎙️</span> <span class="title">Voice note language</span><span class="desc">What language you dictate voice notes in, for transcription</span></span>
        <select id="f-voice-lang" style="width:100%;">
          ${VOICE_LANG_OPTIONS.map(
            (o) => `<option value="${o.value}" ${voiceLangStored() === o.value ? "selected" : ""}>${o.label}</option>`
          ).join("")}
        </select>
      </div>
      <div class="row"><span class="icon">💾</span><span class="main"><span class="title">Data storage</span><span class="desc">Stored locally on this device${account ? ", synced to OneDrive" : ""}</span></span></div>
    </div>
  `;
  shell({ header, body, activeTab: "settings" });

  const $toggle = document.getElementById("onedrive-toggle");
  if ($toggle) {
    $toggle.addEventListener("click", async () => {
      if (account) await msSignOut();
      else await msSignIn();
    });
  }

  const $aiKeyInput = document.getElementById("f-ai-key");
  document.getElementById("ai-key-save").addEventListener("click", () => {
    const value = $aiKeyInput.value.trim();
    // The placeholder dots stand in for an already-saved key — clicking
    // Save without touching the field would otherwise overwrite the real
    // key with literal bullet characters.
    if (!value || value.startsWith("••")) {
      toast("Paste your API key first");
      return;
    }
    setAiKey(value);
    toast("API key saved");
    renderSettings();
  });
  document.getElementById("ai-key-clear")?.addEventListener("click", () => {
    setAiKey("");
    toast("API key removed");
    renderSettings();
  });

  document.getElementById("f-ai-captions").addEventListener("change", (e) => {
    setAiCaptions(e.target.checked);
    toast(e.target.checked ? "AI photo descriptions on" : "AI photo descriptions off");
    renderSettings();
  });

  document.getElementById("f-accordion").addEventListener("change", (e) => {
    setAccordionEnabled(e.target.checked);
    toast(e.target.checked ? "Compact lists on" : "Compact lists off");
    renderSettings();
  });

  document.getElementById("f-ai-translate").addEventListener("change", (e) => {
    setAiTranslate(e.target.checked);
    toast(e.target.checked ? "AI voice note translation on" : "AI voice note translation off");
    renderSettings();
  });

  (async () => {
    const el = document.getElementById("storage-info");
    if (!el) return;
    try {
      const est = navigator.storage?.estimate ? await navigator.storage.estimate() : null;
      const persisted = navigator.storage?.persisted ? await navigator.storage.persisted() : null;
      const mb = (n) => (n / 1048576 >= 1000 ? (n / 1073741824).toFixed(1) + " GB" : Math.round(n / 1048576) + " MB");
      el.textContent = est && est.quota
        ? `Using ${mb(est.usage || 0)} of ${mb(est.quota)}${persisted === true ? " · protected from auto-cleanup" : persisted === false ? " · the phone may clear it when low on space" : ""}`
        : "Not available on this browser";
    } catch {
      el.textContent = "Not available on this browser";
    }
  })();
  document.getElementById("f-voice-lang").addEventListener("change", (e) => {
    setVoiceLang(e.target.value);
    toast("Voice note language saved");
  });
}

// ---------- Quick capture ----------

const VOICE_LANG_STORAGE = "moss_voice_lang";

const VOICE_LANG_OPTIONS = [
  { value: "", label: "Auto (match this device's language)" },
  { value: "en-US", label: "English (US)" },
  { value: "es-US", label: "Español (Estados Unidos)" },
  { value: "es-419", label: "Español (Latinoamérica)" },
  { value: "es-ES", label: "Español (España)" }
];

// The raw stored choice — "" means "unset", i.e. the Settings dropdown
// should show "Auto", not a guessed language. Different from voiceLang()
// below, which is what recording actually uses (it resolves "" down to
// the device's language).
function voiceLangStored() {
  try {
    return localStorage.getItem(VOICE_LANG_STORAGE) || "";
  } catch {
    return "";
  }
}

// "" (unset) means "follow the device's own language" — the common case
// and why voice transcription worked for English devices without any
// setup. Only needs changing when someone dictates in a language their
// device/browser isn't set to.
function voiceLang() {
  // Normalized: a few devices report tags like "en-US@posix" or "es_MX",
  // which speech recognition and Intl both reject.
  return String(voiceLangStored() || navigator.language || "en-US").split("@")[0].replace(/_/g, "-");
}

function setVoiceLang(lang) {
  try {
    if (lang) localStorage.setItem(VOICE_LANG_STORAGE, lang);
    else localStorage.removeItem(VOICE_LANG_STORAGE);
  } catch {
    // Ignored, same as setAiKey — worst case it just falls back to the
    // device language again next time.
  }
}

function currentProjectId() {
  const route = currentRoute();
  return route.name === "project" ? route.id : null;
}

async function pickProjectIfNeeded(preselected) {
  if (preselected) return preselected;
  const projects = (await MossDB.projects.all()).filter(isVisibleProject);
  return new Promise((resolve) => {
    let resolved = false;
    const finish = (value) => {
      if (resolved) return; // closeSheet() also calls back here — only resolve once
      resolved = true;
      resolve(value);
    };
    openSheet(
      `
      <h2>Which project?</h2>
      <div class="card card-list">
        ${projects
          .map(
            (p) => `<button class="row" data-pid="${escapeHtml(p.id)}"><span class="main"><span class="title">${escapeHtml(p.name)}</span></span><span class="chev">›</span></button>`
          )
          .join("")}
      </div>
    `,
      () => finish(null) // backdrop tap / navigation away: don't leave the caller waiting forever
    );
    $sheet.querySelectorAll("[data-pid]").forEach((btn) => {
      btn.addEventListener("click", () => {
        // Resolve BEFORE closeSheet(): closeSheet triggers the cancel
        // callback too, and finish()'s resolved-guard makes that a no-op
        // only if the real answer was already recorded first.
        finish(btn.dataset.pid);
        closeSheet();
      });
    });
  });
}

function wireQuickCapture() {
  $app.querySelectorAll("[data-capture]").forEach((btn) => {
    btn.addEventListener("click", async () => {
      const type = btn.dataset.capture;
      // Inspection has its own "choose the project" step at the top of its sheet.
      if (type === "inspection") return openInspectionSheet(currentProjectId());
      const projectId = await pickProjectIfNeeded(currentProjectId());
      if (!projectId) return;
      if (type === "photo") capturePhoto(projectId);
      else if (type === "voice") captureVoice(projectId);
      else if (type === "document") captureDocument(projectId);
      else if (type === "issue") openIssueSheet(projectId);
      else if (type === "material") openMaterialSheet(projectId);
    });
  });
}

function capturePhoto(projectId) {
  const input = document.createElement("input");
  input.type = "file";
  input.accept = "image/*";
  input.capture = "environment";
  input.addEventListener("change", async () => {
    const file = input.files[0];
    if (!file) return;
    const dataUrl = await photoFileToDataUrl(file);
    // Recorded once and reused for both the OneDrive upload and the local
    // record, so another device can later ask OneDrive for this exact file
    // by name (see hydrateRemoteCapture) — relying on file.name alone
    // wasn't reliable enough to build a re-download path from.
    const remoteFileName = file.name || `photo-${Date.now()}.jpg`;
    let capture;
    try {
      capture = await MossDB.captures.add({ projectId, type: "photo", dataUrl, name: file.name, remoteFileName });
    } catch (err) {
      console.error("Saving photo failed", err);
      toast("Couldn't save the photo — the phone may be out of space");
      return;
    }
    toast("Photo saved");
    if (currentRoute().name === "project") render();
    // Kept so the captions chain below can wait for the real file to
    // actually finish reaching OneDrive before telling OTHER devices this
    // remoteFileName exists — pushing the captures index first would let
    // another device try to hydrate a file that isn't there yet (and
    // hydrateRemoteCapture doesn't retry a failed fetch until next visit).
    const uploadPromise = syncCaptureToOneDrive(projectId, "photo", file, remoteFileName);

    // Push the captures index exactly once, after the real file has
    // finished reaching OneDrive AND (if AI is on) after captioning has
    // had its chance to finish too — so other devices learn about this
    // photo (remoteFileName, so they can hydrate it) with its caption
    // already attached when there is one, and so a photo taken with no
    // API key configured doesn't stay invisible on other devices until
    // this device's app happens to reload. captionPhoto's own failure is
    // swallowed here (logged, not rethrown) so a captioning error never
    // skips announcing the photo itself.
    const captionPromise = aiCaptionsEnabled()
      ? captionPhoto(dataUrl)
          .then((caption) => caption && MossDB.captures.update(capture.id, { caption }))
          .then(() => { if (currentRoute().name === "project") render(); })
          .catch((err) => console.warn("Photo captioning skipped:", err.message))
      : Promise.resolve();

    Promise.all([uploadPromise, captionPromise]).then(() => syncCapturesWithOneDrive());
  });
  input.click();
}

function captureDocument(projectId) {
  const input = document.createElement("input");
  input.type = "file";
  input.accept = ".pdf,.doc,.docx,.png,.jpg,.jpeg,image/*,application/pdf";
  input.addEventListener("change", async () => {
    const file = input.files[0];
    if (!file) return;
    await MossDB.captures.add({ projectId, type: "document", name: file.name, size: file.size });
    toast(`Saved "${file.name}"`);
    if (currentRoute().name === "project") render();
    syncCaptureToOneDrive(projectId, "document", file, file.name);
  });
  input.click();
}

async function captureVoice(projectId) {
  if (!navigator.mediaDevices?.getUserMedia) {
    toast("Microphone not available in this browser");
    return;
  }
  try {
    const stream = await navigator.mediaDevices.getUserMedia({ audio: true });
    const recorder = new MediaRecorder(stream);
    const chunks = [];
    recorder.addEventListener("dataavailable", (e) => chunks.push(e.data));

    // Live speech-to-text runs alongside the recording, entirely in the
    // browser (the Web Speech API) — free, no API call, no cost per note.
    // So Ask AI can read what was said. Not every browser has this
    // (notably iOS Safari) — when it's missing, the voice note still saves
    // fine, it just won't have text Ask AI can read.
    //
    // Language: the Web Speech API can't auto-detect language mid-recording
    // — one recognizer, one language, chosen up front. voiceLang() defaults
    // to the device's own language setting (so it "just works" whichever
    // language the phone/browser is in), but is overridable in Settings
    // for anyone who dictates in a different language than their device UI.
    const SpeechRecognition = window.SpeechRecognition || window.webkitSpeechRecognition;
    let recognition = null;
    let transcript = "";
    if (SpeechRecognition) {
      recognition = new SpeechRecognition();
      recognition.continuous = true;
      recognition.interimResults = false;
      recognition.lang = voiceLang();
      recognition.addEventListener("result", (e) => {
        for (let i = e.resultIndex; i < e.results.length; i++) {
          if (e.results[i].isFinal) transcript += e.results[i][0].transcript + " ";
        }
      });
      // A recognition hiccup (e.g. a pause) shouldn't kill the recording —
      // just stop transcribing; the audio keeps recording regardless.
      recognition.addEventListener("error", () => {});
      try { recognition.start(); } catch {}
    }

    openSheet(`
      <h2>Voice note</h2>
      <div class="rec-indicator"><span class="dot"></span><span>Recording…</span></div>
      <div class="sheet-actions">
        <button class="btn ghost" id="cancel-rec">Cancel</button>
        <button class="btn primary" id="stop-rec">Stop &amp; Save</button>
      </div>
    `);

    recorder.start();

    const stop = (save) => {
      recorder.stop();
      stream.getTracks().forEach((t) => t.stop());
      if (recognition) { try { recognition.stop(); } catch {} }
      closeSheet();
      if (!save) return;
      recorder.addEventListener("stop", async () => {
        // Use the recorder's own mimeType (Safari records audio/mp4, not
        // webm) — hardcoding "audio/webm" produced a blob mislabeled on
        // iPhone, which is one of the target platforms.
        const blob = new Blob(chunks, { type: recorder.mimeType || "audio/webm" });
        const dataUrl = await blobToDataUrl(blob);
        const finalTranscript = transcript.trim() || null;
        // Same reasoning as the photo path: fix the filename once and
        // store it, so another device can re-download this exact file
        // from OneDrive later (hydrateRemoteCapture).
        const ext = (recorder.mimeType || "audio/webm").includes("mp4") ? "m4a" : "webm";
        const remoteFileName = `voice-${Date.now()}.${ext}`;
        const savedVoice = await MossDB.captures.add({ projectId, type: "voice", dataUrl, transcript: finalTranscript, remoteFileName });
        toast(finalTranscript ? "Voice note saved (transcribed)" : "Voice note saved");
        if (currentRoute().name === "project") render();
        // Spanish <-> English translation (if switched on) is added a moment
        // later, once the AI answers; the note is already saved without it.
        const translatePromise = finalTranscript
          ? tryTranslate(finalTranscript).then(async (tr) => {
              if (!tr) return;
              await MossDB.captures.update(savedVoice.id, { translation: tr.translation, translationLang: tr.target });
              if (currentRoute().name === "project") render();
            })
          : Promise.resolve();
        // Same ordering reason as capturePhoto: don't advertise this
        // remoteFileName to other devices until the real file has actually
        // finished uploading.
        const uploadPromise = syncCaptureToOneDrive(projectId, "voice", blob, remoteFileName);
        // Push the captures index once the real file is up, whether or not
        // this note got a transcript (no Web Speech API on this browser,
        // or nothing recognized) — same reasoning as capturePhoto above:
        // otherwise this voice note stays invisible to other devices until
        // this device's app happens to reload.
        Promise.all([uploadPromise, translatePromise]).then(() => syncCapturesWithOneDrive());
      });
    };

    document.getElementById("stop-rec").addEventListener("click", () => stop(true));
    document.getElementById("cancel-rec").addEventListener("click", () => stop(false));
  } catch (err) {
    toast("Microphone permission denied");
  }
}

// ---------- New Issue: photo + required voice note + AI ----------
// Flow: (1) take a photo, (2) record a voice note — the issue can't be
// saved without both — then (3) AI picks the trade, suggests a title and
// cleans up the transcript. The note is saved stamped with the day, date
// and time it was recorded. The photo and voice note are saved as normal
// captures too (linked to the issue by issueId), so they sync to OneDrive
// like any other photo/voice note.

// Day + date + time a voice note was recorded, in the voice-note language
// (e.g. "miércoles, 30 de septiembre de 2026 · 7:26 p. m.").
function issueStamp(date) {
  // Some devices report odd language tags ("en-US@posix", "es_MX") that
  // Intl rejects. Clean the tag up, and if it's still rejected fall back to
  // the device's default locale — the stamp must always carry the weekday,
  // date and time, never degrade to a short numeric date.
  const cleaned = String(voiceLang() || "").split("@")[0].replace(/_/g, "-");
  for (const loc of [cleaned, undefined]) {
    try {
      const d = date.toLocaleDateString(loc || undefined, { weekday: "long", year: "numeric", month: "long", day: "numeric" });
      const t = date.toLocaleTimeString(loc || undefined, { hour: "numeric", minute: "2-digit" });
      return `${d} · ${t}`;
    } catch {
      // try the next locale
    }
  }
  return date.toLocaleString();
}

// Starts the mic + (where the browser supports it) free live speech-to-text.
// Same approach as captureVoice, but returns a handle so the New Issue sheet
// can drive it. Throws if the mic isn't available or permission is denied.
async function startVoiceRecorder() {
  if (!navigator.mediaDevices?.getUserMedia || typeof MediaRecorder === "undefined") {
    throw new Error("Microphone not available in this browser");
  }
  const stream = await navigator.mediaDevices.getUserMedia({ audio: true });
  const recorder = new MediaRecorder(stream);
  const chunks = [];
  recorder.addEventListener("dataavailable", (e) => {
    if (e.data && e.data.size) chunks.push(e.data);
  });

  const SpeechRecognition = window.SpeechRecognition || window.webkitSpeechRecognition;
  let recognition = null;
  let recognitionActive = false;
  let stopping = false;
  let restarts = 0;
  let transcript = "";
  if (SpeechRecognition) {
    recognition = new SpeechRecognition();
    recognition.continuous = true;
    recognition.interimResults = false;
    recognition.lang = voiceLang();
    recognition.addEventListener("start", () => { recognitionActive = true; });
    recognition.addEventListener("result", (e) => {
      for (let i = e.resultIndex; i < e.results.length; i++) {
        if (e.results[i].isFinal) transcript += e.results[i][0].transcript + " ";
      }
    });
    recognition.addEventListener("error", (e) => {
      // Permission/service problems won't fix themselves — stop retrying.
      if (e.error === "not-allowed" || e.error === "service-not-allowed") stopping = true;
    });
    // Some browsers end recognition after a pause even in continuous mode;
    // restart it (a few times at most) so a longer note is still transcribed.
    recognition.addEventListener("end", () => {
      recognitionActive = false;
      if (!stopping && restarts < 20) {
        restarts++;
        try { recognition.start(); } catch {}
      }
    });
    try { recognition.start(); } catch {}
  }

  const startedAt = new Date();
  recorder.start();
  let released = false;
  const release = () => {
    if (released) return;
    released = true;
    stopping = true;
    stream.getTracks().forEach((t) => t.stop());
    if (recognition) { try { recognition.stop(); } catch {} }
  };

  return {
    startedAt,
    // Resolves with the finished recording, or null if nothing was recorded.
    async stop() {
      if (recorder.state === "inactive") {
        release();
        return null;
      }
      stopping = true;
      // Give the recognizer a moment to deliver its last phrase — otherwise
      // the final words said right before tapping Stop can get lost.
      const recognitionEnded =
        recognition && recognitionActive
          ? new Promise((resolve) => {
              recognition.addEventListener("end", resolve, { once: true });
              setTimeout(resolve, 1200);
            })
          : Promise.resolve();
      if (recognition) { try { recognition.stop(); } catch {} }
      const recorderStopped = new Promise((resolve) => recorder.addEventListener("stop", resolve, { once: true }));
      recorder.stop();
      await recorderStopped;
      await recognitionEnded;
      release();
      const mimeType = recorder.mimeType || "audio/webm";
      const blob = new Blob(chunks, { type: mimeType });
      if (!blob.size) return null;
      return { blob, dataUrl: await blobToDataUrl(blob), mimeType, transcript: transcript.trim(), startedAt };
    },
    cancel() {
      try { if (recorder.state !== "inactive") recorder.stop(); } catch {}
      release();
    }
  };
}

const MAX_ISSUE_PHOTOS = 5; // main photo + extras picked from the gallery

function openIssueSheet(projectId) {
  const state = {
    photo: null, // { dataUrl, file }
    extras: [], // more photos from the gallery: [{ dataUrl, file }]
    voice: null, // { blob, dataUrl, mimeType, transcript, startedAt }
    recorder: null,
    recording: false,
    analyzing: false,
    aiNote: "", // shown when AI couldn't run / failed
    description: "", // AI's one-line description of the photo
    translation: "", // Spanish <-> English translation of the note (if switched on)
    translationLang: "",
    translatedFrom: "", // the note text the translation belongs to
    translating: false,
    titleLang: issueTitleLang(), // "en" | "es": language the title is written in
    titleBusy: false, // AI rewriting the title in the other language
    title: "",
    trade: "General",
    note: "",
    titleTouched: false, // once the user edits a field by hand, AI never overwrites it
    tradeTouched: false,
    noteTouched: false,
    saving: false
  };
  let closed = false;
  let analysisRun = 0; // bumped on every analysis so a stale result can't overwrite a newer one
  let translateRun = 0;
  let titleRun = 0;

  const draw = () => {
    if (closed) return;
    const haveBoth = !!(state.photo && state.voice);
    const canSave = haveBoth && !state.recording && !state.analyzing && !state.translating && !state.titleBusy && !state.saving;

    const photoBlock = `
      ${state.photo ? `<img class="issue-photo" src="${state.photo.dataUrl}" alt="Issue photo">` : ""}
      <button class="btn ghost" id="i-photo" ${state.recording || state.saving ? "disabled" : ""}>${state.photo ? "📸 Retake photo" : "📸 Take photo"}</button>
      ${
        state.extras.length
          ? `<div style="display:flex; flex-wrap:wrap; gap:8px; margin-top:8px;">${state.extras
              .map(
                (p, i) => `<div style="position:relative;"><img src="${p.dataUrl}" alt="Extra photo ${i + 1}" style="width:72px; height:72px; object-fit:cover; border-radius:10px; display:block;"><button class="btn ghost" data-rm-extra="${i}" ${state.saving ? "disabled" : ""} style="position:absolute; top:-6px; right:-6px; width:26px; height:26px; padding:0; border-radius:50%; font-size:13px; line-height:1;" aria-label="Remove photo">✕</button></div>`
              )
              .join("")}</div>`
          : ""
      }
      <button class="btn ghost" id="i-gallery" ${state.recording || state.saving || (state.photo ? 1 : 0) + state.extras.length >= MAX_ISSUE_PHOTOS ? "disabled" : ""} style="margin-top:8px;">🖼 Choose from gallery${(state.photo ? 1 : 0) + state.extras.length ? ` (${(state.photo ? 1 : 0) + state.extras.length}/${MAX_ISSUE_PHOTOS})` : ` (up to ${MAX_ISSUE_PHOTOS})`}</button>
    `;

    let voiceBlock;
    if (!state.photo) {
      voiceBlock = `<button class="btn ghost" disabled>🎤 Take the photo first</button>`;
    } else if (state.recording) {
      voiceBlock = `
        <div class="rec-indicator"><span class="dot"></span><span>Recording… describe the issue</span></div>
        <button class="btn primary" id="i-stop">⏹ Stop recording</button>
      `;
    } else if (state.voice) {
      voiceBlock = `
        <audio controls preload="metadata" style="width:100%; height:36px;" src="${state.voice.dataUrl}"></audio>
        ${voiceLangPickerHtml("i-vlang", state.saving)}
        <button class="btn ghost" id="i-rec" ${state.saving ? "disabled" : ""}>🎤 Re-record</button>
      `;
    } else {
      voiceBlock = `${voiceLangPickerHtml("i-vlang", false)}<button class="btn primary" id="i-rec">🎤 Record voice note</button>`;
    }

    let detailsBlock = "";
    if (haveBoth && !state.recording) {
      detailsBlock = `
        <div class="field">
          <label>Note · ${escapeHtml(issueStamp(state.voice.startedAt))}</label>
          <textarea id="f-note" placeholder="${state.voice.transcript ? "" : "Couldn't transcribe this voice note — type what you said (optional)"}">${escapeHtml(state.note)}</textarea>
        </div>
        ${state.translating ? `<p class="empty" style="text-align:left;">🌐 AI is translating the note…</p>` : translationCardHtml(state.translation, state.translationLang) + retranslateButtonHtml("i-retr", !!state.translation, state.note)}
        <div class="field">
          <label>Trade</label>
          <select id="f-trade">${TRADES.map((t) => `<option ${t === state.trade ? "selected" : ""}>${t}</option>`).join("")}</select>
        </div>
        <div class="field">
          <label>Title · language</label>
          <select id="f-title-lang" ${state.analyzing || state.titleBusy || state.saving ? "disabled" : ""}>
            <option value="en" ${state.titleLang === "en" ? "selected" : ""}>English</option>
            <option value="es" ${state.titleLang === "es" ? "selected" : ""}>Español</option>
          </select>
        </div>
        <div class="field">
          <label>Title</label>
          <input id="f-title" value="${escapeHtml(state.title)}" placeholder='${state.titleLang === "es" ? "p. ej. &quot;Registro del piso del baño&quot;" : "e.g. &quot;Hall bathroom floor register&quot;"}'>
        </div>
        ${state.titleBusy ? `<p class="empty" style="text-align:left;">🌐 ${state.titleLang === "es" ? "Traduciendo el título…" : "Translating the title…"}</p>` : ""}
        ${state.analyzing ? `<p class="empty" style="text-align:left;">🤖 AI is identifying the trade and writing the note…</p>` : ""}
        ${state.aiNote ? `<p class="empty" style="text-align:left;">${escapeHtml(state.aiNote)}</p>` : ""}
      `;
    }

    $sheet.innerHTML = `
      <h2>New Issue</h2>
      <div class="field"><label>1 · Photo</label>${photoBlock}</div>
      <div class="field"><label>2 · Voice note (required)</label>${voiceBlock}</div>
      ${detailsBlock}
      <div class="sheet-actions">
        <button class="btn ghost" id="i-cancel">Cancel</button>
        <button class="btn primary" id="i-save" ${canSave ? "" : "disabled"} style="${canSave ? "" : "opacity:.5;"}">${state.saving ? "Saving…" : state.analyzing ? "Analyzing…" : state.translating || state.titleBusy ? "Translating…" : "Save Issue"}</button>
      </div>
    `;
    wire();
  };

  const wire = () => {
    const $ = (id) => document.getElementById(id);
    $("i-cancel")?.addEventListener("click", closeSheet);
    $("i-photo")?.addEventListener("click", () => takePhoto());
    $("i-gallery")?.addEventListener("click", pickFromGallery);
    $sheet.querySelectorAll("[data-rm-extra]").forEach((el) =>
      el.addEventListener("click", () => {
        state.extras.splice(Number(el.dataset.rmExtra), 1);
        draw();
      })
    );
    $("i-rec")?.addEventListener("click", startRecording);
    $("i-stop")?.addEventListener("click", stopRecording);
    $("i-save")?.addEventListener("click", save);
    // Keep state in sync as the user types, so redraws never lose edits.
    $("f-title")?.addEventListener("input", (e) => { state.title = e.target.value; state.titleTouched = true; });
    $("f-title-lang")?.addEventListener("change", (e) => changeTitleLang(e.target.value));
    $("i-retr")?.addEventListener("click", () => runTranslate(true));
    $("i-vlang")?.addEventListener("change", (e) => setVoiceLang(e.target.value));
    $("f-trade")?.addEventListener("change", (e) => { state.trade = e.target.value; state.tradeTouched = true; });
    $("f-note")?.addEventListener("input", (e) => { state.note = e.target.value; state.noteTouched = true; });
  };

  const takePhoto = (fromGallery = false) => {
    const input = document.createElement("input");
    input.type = "file";
    input.accept = "image/*";
    if (!fromGallery) input.capture = "environment";
    input.addEventListener("change", async () => {
      const file = input.files[0];
      if (!file || closed) return;
      let dataUrl;
      try {
        dataUrl = await photoFileToDataUrl(file);
      } catch {
        toast("Couldn't read that photo");
        return;
      }
      if (closed) return;
      state.photo = { dataUrl, file };
      state.description = "";
      draw();
      if (state.voice) runAnalysis(); // photo replaced after the note was recorded
    });
    input.click();
  };

  // Gallery: pick several photos at once (up to 5 per issue in total). If
  // there is no main photo yet, the first one picked becomes it.
  const pickFromGallery = () => {
    const input = document.createElement("input");
    input.type = "file";
    input.accept = "image/*";
    input.multiple = true;
    input.addEventListener("change", async () => {
      const files = [...(input.files || [])];
      if (!files.length || closed) return;
      const room = MAX_ISSUE_PHOTOS - (state.photo ? 1 : 0) - state.extras.length;
      if (files.length > room) toast(`Only ${MAX_ISSUE_PHOTOS} photos per issue`);
      const loaded = [];
      for (const file of files.slice(0, Math.max(0, room))) {
        try {
          loaded.push({ dataUrl: await photoFileToDataUrl(file), file });
        } catch {
          toast("Couldn't read one of the photos");
        }
      }
      if (closed || !loaded.length) return;
      let newMain = false;
      if (!state.photo) {
        state.photo = loaded.shift();
        state.description = "";
        newMain = true;
      }
      state.extras.push(...loaded);
      draw();
      if (newMain && state.voice) runAnalysis();
    });
    input.click();
  };

  const startRecording = async () => {
    try {
      state.recorder = await startVoiceRecorder();
    } catch (err) {
      toast(err.message === "Microphone not available in this browser" ? err.message : "Microphone permission denied");
      return;
    }
    if (closed) {
      state.recorder.cancel();
      state.recorder = null;
      return;
    }
    state.recording = true;
    draw();
  };

  const stopRecording = async () => {
    const rec = state.recorder;
    if (!rec) return;
    state.recorder = null;
    const stopBtn = document.getElementById("i-stop");
    if (stopBtn) { stopBtn.disabled = true; stopBtn.textContent = "Finishing…"; }
    let result = null;
    try {
      result = await rec.stop();
    } catch (err) {
      console.error("Stopping recording failed", err);
    }
    state.recording = false;
    if (closed) return;
    if (!result) {
      toast("Nothing was recorded — try again");
      draw();
      return;
    }
    state.voice = result;
    state.note = result.transcript; // editable; AI cleans it up below
    state.noteTouched = false;
    state.aiNote = "";
    state.translation = "";
    state.translatedFrom = "";
    translateRun++;
    state.translating = false;
    draw();
    runAnalysis();
  };

  // Picking the other title language: remember it, and have the AI rewrite
  // the title already on screen (if any) in that language.
  const changeTitleLang = async (lang) => {
    lang = lang === "es" ? "es" : "en";
    if (lang === state.titleLang) return;
    state.titleLang = lang;
    setIssueTitleLang(lang);
    const current = state.title.trim();
    if (!current || !aiConfigured()) {
      draw();
      return;
    }
    const run = ++titleRun;
    state.titleBusy = true;
    draw();
    try {
      const t = await withTimeout(translateIssueTitle(current, lang), 20000, "it took too long");
      if (closed || run !== titleRun) return;
      if (state.title.trim() === current) state.title = t; // don't clobber text typed meanwhile
    } catch {
      if (closed || run !== titleRun) return;
      toast("Couldn't translate the title — edit it by hand");
    }
    state.titleBusy = false;
    draw();
  };

  // Spanish <-> English translation of the final note (only when switched on).
  // `force` = "Translate again" was pressed: ignores the Settings switch and
  // keeps the old translation if the new try fails.
  const runTranslate = async (force = false) => {
    const run = ++translateRun;
    const text = state.note.trim();
    if (!(force ? aiConfigured() : aiTranslateEnabled()) || !text) {
      state.translating = false;
      state.translation = "";
      state.translatedFrom = "";
      draw();
      return;
    }
    state.translating = true;
    draw();
    const tr = await tryTranslate(text, force, force ? state.translation : "");
    if (closed || run !== translateRun) return;
    if (tr) {
      state.translation = tr.translation;
      state.translationLang = tr.target;
    } else if (force) {
      toast("Couldn't translate — try again");
    } else {
      state.translation = "";
      state.translationLang = "";
    }
    state.translatedFrom = text;
    state.translating = false;
    draw();
  };

  const runAnalysis = async () => {
    if (!state.photo || !state.voice) return;
    if (!aiConfigured()) {
      state.aiNote = "Add a Claude API key in Settings to have AI pick the trade and write the note. Choose the trade manually for now.";
      draw();
      return;
    }
    const run = ++analysisRun;
    state.analyzing = true;
    state.aiNote = "";
    draw();
    try {
      // Never leave Save locked on a bad connection: give the AI 30s, then
      // let the user finish the issue by hand.
      let timer;
      const timeout = new Promise((_, reject) => {
        timer = setTimeout(() => reject(new Error("it took too long")), 30000);
      });
      let r;
      try {
        r = await Promise.race([analyzeIssueCapture(state.photo.dataUrl, state.voice.transcript, TRADES, state.titleLang), timeout]);
      } finally {
        clearTimeout(timer);
      }
      if (closed || run !== analysisRun) return;
      if (!state.tradeTouched) state.trade = r.trade;
      if (!state.titleTouched && r.title) state.title = r.title;
      state.description = r.description;
      if (!state.noteTouched && r.note) state.note = r.note;
    } catch (err) {
      if (closed || run !== analysisRun) return;
      state.aiNote = `AI couldn't analyze this one (${err.message}). Pick the trade and add a title manually.`;
    }
    state.analyzing = false;
    draw();
    runTranslate();
  };

  const save = async () => {
    if (!state.photo || !state.voice || state.recording || state.analyzing || state.translating || state.titleBusy || state.saving) return;
    state.saving = true;
    const written = []; // ids saved so far, to clean up if the phone refuses part of it
    draw();
    try {
      // The note may have been edited after it was translated: refresh the
      // translation so the saved pair always matches (skipped if it fails).
      if (aiTranslateEnabled() && state.note.trim() && state.note.trim() !== state.translatedFrom) {
        const tr = await tryTranslate(state.note.trim());
        state.translation = tr ? tr.translation : "";
        state.translationLang = tr ? tr.target : "";
        state.translatedFrom = state.note.trim();
      } else if (!state.note.trim() || state.note.trim() !== state.translatedFrom) {
        state.translation = ""; // empty note, or switch off and note edited since: don't save a mismatched translation
      }
      const now = Date.now();
      const extOfPhoto = { "image/png": "png", "image/heic": "heic", "image/heif": "heif", "image/webp": "webp" }[state.photo.file.type] || "jpg";
      const extOfVoice = state.voice.mimeType.includes("mp4") ? "m4a" : "webm";
      // Unique file names: phones often name every camera photo "image.jpg",
      // which would overwrite earlier photos in OneDrive.
      const photoName = `issue-photo-${now}.${extOfPhoto}`;
      const voiceName = `issue-voice-${now}.${extOfVoice}`;
      const issueId = MossDB.uid();
      const photoId = MossDB.uid();
      const voiceId = MossDB.uid();
      const extraPhotoIds = state.extras.map(() => MossDB.uid());
      written.push(photoId, voiceId, ...extraPhotoIds);
      const stamp = issueStamp(state.voice.startedAt);
      const note = state.note.trim();
      const title = state.title.trim() || (note ? note.slice(0, 60) : `${state.titleLang === "es" ? "Problema" : "Issue"} — ${stamp}`);
      const trade = TRADES.includes(state.trade) ? state.trade : "General";

      await MossDB.captures.add({
        id: photoId,
        projectId,
        type: "photo",
        dataUrl: state.photo.dataUrl,
        name: state.photo.file.name,
        remoteFileName: photoName,
        caption: (aiCaptionsEnabled() && state.description) || title,
        issueId
      });
      const extraNames = [];
      for (let k = 0; k < state.extras.length; k++) {
        const ex = state.extras[k];
        const exExt = { "image/png": "png", "image/heic": "heic", "image/heif": "heif", "image/webp": "webp" }[ex.file.type] || "jpg";
        const exName = `issue-photo-${now}-${k + 2}.${exExt}`;
        extraNames.push(exName);
        await MossDB.captures.add({
          id: extraPhotoIds[k],
          projectId,
          type: "photo",
          dataUrl: ex.dataUrl,
          name: ex.file.name,
          remoteFileName: exName,
          caption: title,
          issueId
        });
      }
      await MossDB.captures.add({
        id: voiceId,
        projectId,
        type: "voice",
        dataUrl: state.voice.dataUrl,
        transcript: note || state.voice.transcript || null,
        remoteFileName: voiceName,
        issueId,
        ...(state.translation ? { translation: state.translation, translationLang: state.translationLang } : {})
      });
      await MossDB.issues.add({
        id: issueId,
        projectId,
        title,
        trade,
        requirement: note ? `${stamp}\n${note}` : stamp,
        photoId,
        voiceId,
        ...(extraPhotoIds.length ? { extraPhotoIds } : {}),
        recordedAt: state.voice.startedAt.toISOString(),
        ...(state.translation ? { translation: state.translation, translationLang: state.translationLang } : {})
      });

      const photoFile = state.photo.file;
      const voiceBlob = state.voice.blob;
      const extraFiles = state.extras.map((x) => x.file);
      closeSheet();
      toast("Issue created");
      if (currentRoute().name === "project" || currentRoute().name === "home") render();

      // Best-effort OneDrive copy of both files; announce them to other
      // devices only once both have actually finished uploading.
      Promise.all([
        syncCaptureToOneDrive(projectId, "photo", photoFile, photoName),
        ...extraFiles.map((f, k) => syncCaptureToOneDrive(projectId, "photo", f, extraNames[k])),
        syncCaptureToOneDrive(projectId, "voice", voiceBlob, voiceName)
      ]).then(() => syncCapturesWithOneDrive());
    } catch (err) {
      console.error("Saving issue failed", err);
      for (const id of written) {
        try { await MossDB.captures.remove(id); } catch {}
      }
      state.saving = false;
      toast("Couldn't save the issue — the phone is out of space? Nothing was kept; try again");
      draw();
    }
  };

  openSheet("", () => {
    // Runs whenever the sheet goes away (Cancel, backdrop tap, navigation,
    // or after a successful save): always release the microphone.
    closed = true;
    analysisRun++;
    translateRun++;
    titleRun++;
    if (state.recorder) {
      state.recorder.cancel();
      state.recorder = null;
    }
  });
  draw();
}

// Opens a saved issue with its linked photo, stamped note and voice note.
function openIssueDetail(issue, captures) {
  const photo = captures.find((c) => c.id === issue.photoId);
  const extraPhotos = (issue.extraPhotoIds || []).map((id) => captures.find((c) => c.id === id)).filter((c) => c?.dataUrl);
  const voice = captures.find((c) => c.id === issue.voiceId);
  const noteText = issuePdfData(issue, captures).note.trim();
  openSheet(`
    <h2>${escapeHtml(issue.title)}</h2>
    <span class="desc">${escapeHtml(issue.trade)}</span>
    ${photo?.dataUrl ? `<img class="issue-photo" src="${photo.dataUrl}" alt="Issue photo">` : ""}
    ${
      extraPhotos.length
        ? `<div style="display:grid; grid-template-columns:1fr 1fr; gap:8px; margin:8px 0;">${extraPhotos
            .map((p, i) => `<img src="${p.dataUrl}" alt="Issue photo ${i + 2}" style="width:100%; border-radius:10px; display:block;">`)
            .join("")}</div>`
        : ""
    }
    ${issue.requirement ? `<div class="card" style="padding:12px 14px; font-size:14px; line-height:1.5; white-space:pre-wrap;">${escapeHtml(issue.requirement)}</div>` : ""}
    ${translationCardHtml(issue.translation, issue.translationLang)}
    ${retranslateButtonHtml("d-retr", !!issue.translation, noteText)}
    ${voice?.dataUrl ? `<audio controls preload="metadata" style="width:100%; height:36px;" src="${voice.dataUrl}"></audio>` : ""}
    ${
      issue.status === "Completed"
        ? `<button class="btn ghost" id="d-restore" style="width:100%;">↩ Back to Open Issues</button>`
        : `<button class="btn ghost" id="d-archive" style="width:100%;">🗄️ Fixed — move to drawer</button>`
    }
    ${closeButton()}
  `);
  wireCloseButton();
  document.getElementById("d-archive")?.addEventListener("click", async () => {
    await MossDB.issues.add({ ...issue, status: "Completed", archivedAt: new Date().toISOString() });
    closeSheet();
    toast("Moved to the drawer");
    render();
  });
  document.getElementById("d-restore")?.addEventListener("click", async () => {
    await MossDB.issues.add({ ...issue, status: "Open" });
    closeSheet();
    toast("Back in Open Issues");
    render();
  });
  wireRetranslate(
    "d-retr",
    noteText,
    issue.translation || "",
    async (patch) => {
      Object.assign(issue, patch); // the dashboard's own copy too, so it shows the new one next time
      await MossDB.issues.add({ ...issue });
      if (voice) {
        Object.assign(voice, patch);
        await MossDB.captures.update(voice.id, patch);
      }
    },
    () => openIssueDetail(issue, captures)
  );
}

// ---------- Issues PDF export ----------
// Opens a picker of open issues (all pre-checked), builds a PDF from the
// ones left checked, then offers Share (phone share sheet → email, text,
// WhatsApp, etc.) and Open/Download. `projectId` limits it to one project;
// null/undefined covers every active project.

function issuePdfData(issue, captures) {
  const photo = captures.find((c) => c.id === issue.photoId);
  const req = issue.requirement || "";
  let stamp;
  let note;
  if (issue.recordedAt) {
    // Created with the photo + voice note flow: first line is the stamp.
    const nl = req.indexOf("\n");
    stamp = nl === -1 ? req : req.slice(0, nl);
    note = nl === -1 ? "" : req.slice(nl + 1);
  } else {
    // Older issue: no stamp saved, so show when it was created.
    stamp = issue.createdAt ? issueStamp(new Date(issue.createdAt)) : "";
    note = req;
  }
  return {
    title: issue.title,
    trade: issue.trade,
    stamp,
    note,
    translation: issue.translation || "",
    translationLabel: issue.translation ? translationLabel(issue.translationLang) : "",
    photoDataUrl: photo?.dataUrl || null,
    extraPhotoDataUrls: (issue.extraPhotoIds || []).map((id) => captures.find((c) => c.id === id)?.dataUrl).filter(Boolean)
  };
}

async function openIssuesPdfSheet(projectId) {
  if (!window.jspdf) {
    toast("PDF library missing — upload jspdf.umd.min.js to GitHub");
    return;
  }
  const projects = (await MossDB.projects.all()).filter(isVisibleProject).filter((p) => !projectId || p.id === projectId);
  const allIssues = await MossDB.issues.all();
  const groups = projects
    .map((p) => ({
      project: p,
      issues: allIssues
        .filter((i) => i.projectId === p.id && i.status === "Open")
        .sort((a, b) => String(a.createdAt).localeCompare(String(b.createdAt)))
    }))
    .filter((g) => g.issues.length);

  if (!groups.length) {
    toast("No open issues to put in a PDF");
    return;
  }

  openSheet(`
    <h2>Issues PDF</h2>
    <p class="empty" style="text-align:left; margin-top:-4px;">Choose what goes in the report. Each issue includes its photo and written note.</p>
    ${groups
      .map(
        (g) => `
      ${projectId ? "" : `<div class="section-label">${escapeHtml(g.project.name)}</div>`}
      <div class="card card-list">
        ${g.issues
          .map(
            (i) => `
          <label class="row" style="cursor:pointer; gap:12px;">
            <input type="checkbox" data-iid="${escapeHtml(i.id)}" checked style="width:20px; height:20px; flex:none;">
            <span class="main"><span class="title">${escapeHtml(i.title)}</span><span class="desc">${escapeHtml(i.trade)}${i.photoId ? " · 📷" : ""}</span></span>
          </label>`
          )
          .join("")}
      </div>`
      )
      .join("")}
    <div class="sheet-actions">
      <button class="btn ghost" id="pdf-cancel">Cancel</button>
      <button class="btn primary" id="pdf-create">Create PDF</button>
    </div>
  `);
  document.getElementById("pdf-cancel").addEventListener("click", closeSheet);
  document.getElementById("pdf-create").addEventListener("click", async () => {
    const chosen = new Set([...$sheet.querySelectorAll("[data-iid]")].filter((el) => el.checked).map((el) => el.dataset.iid));
    if (!chosen.size) {
      toast("Select at least one issue");
      return;
    }
    const btn = document.getElementById("pdf-create");
    btn.disabled = true;
    btn.textContent = "Building PDF…";
    try {
      const sections = [];
      for (const g of groups) {
        const picked = g.issues.filter((i) => chosen.has(i.id));
        if (!picked.length) continue;
        const captures = await MossDB.captures.forProject(g.project.id);
        sections.push({
          projectName: g.project.name,
          address: g.project.address || "",
          issues: picked.map((i) => issuePdfData(i, captures))
        });
      }
      const blob = await buildIssuesPdf(sections);
      const day = new Date().toISOString().slice(0, 10);
      const label = sections.length === 1 ? sections[0].projectName : "All projects";
      const fileName = `Issues - ${label} - ${day}.pdf`.replace(/[\\/:*?"<>|]+/g, "").replace(/\s+/g, " ");
      showPdfReady(blob, fileName, sections.reduce((n, s) => n + s.issues.length, 0));
    } catch (err) {
      console.error("PDF build failed", err);
      toast("Couldn't build the PDF — try again");
      btn.disabled = false;
      btn.textContent = "Create PDF";
    }
  });
}

// Second step: the PDF exists — share it or open/download it. Share has to
// start from its own tap (phones won't open the share sheet after a long
// wait that began with a different tap), which is why it's a separate screen.
function showPdfReady(blob, fileName, count, noun = "issue") {
  const url = URL.createObjectURL(blob);
  const file = new File([blob], fileName, { type: "application/pdf" });
  const canShare = !!(navigator.canShare && navigator.canShare({ files: [file] }));
  const sizeKb = Math.max(1, Math.round(blob.size / 1024));
  const sizeText = sizeKb >= 1024 ? `${(sizeKb / 1024).toFixed(1)} MB` : `${sizeKb} KB`;

  openSheet(
    `
    <h2>PDF ready</h2>
    <div class="card" style="padding:14px 16px; font-size:14px; line-height:1.5;">
      <strong>${escapeHtml(fileName)}</strong><br>
      <span class="desc">${count} ${noun}${count === 1 ? "" : "s"} · ${sizeText}</span>
    </div>
    <div class="sheet-actions" style="flex-direction:column;">
      ${canShare ? `<button class="btn primary" id="pdf-share">📤 Share / Send</button>` : ""}
      <button class="btn ${canShare ? "ghost" : "primary"}" id="pdf-open">${canShare ? "Open / Download" : "📥 Open / Download"}</button>
      <button class="btn ghost" id="pdf-done">Done</button>
    </div>
  `,
    () => setTimeout(() => URL.revokeObjectURL(url), 60000) // whenever this sheet goes away; delay so a just-opened PDF tab can finish loading
  );

  document.getElementById("pdf-done").addEventListener("click", closeSheet);
  document.getElementById("pdf-share")?.addEventListener("click", async () => {
    try {
      await navigator.share({ files: [file], title: fileName });
    } catch (err) {
      if (err && err.name !== "AbortError") toast("Couldn't open sharing — use Open / Download");
    }
  });
  document.getElementById("pdf-open").addEventListener("click", () => {
    const a = document.createElement("a");
    a.href = url;
    a.download = fileName;
    a.target = "_blank";
    a.rel = "noopener";
    document.body.appendChild(a);
    a.click();
    a.remove();
  });
}

// Small "what language are you speaking" picker shown above the record
// button. Speech-to-text works in ONE language per recording, so dictating
// Spanish with an English device language gives garbage: pick it here
// (same setting as Settings → Voice note language).
function voiceLangPickerHtml(id, disabled) {
  const cur = voiceLangStored();
  return `<select id="${id}" ${disabled ? "disabled" : ""} aria-label="Voice note language" style="width:100%; margin-bottom:6px;">${VOICE_LANG_OPTIONS.map(
    (o) => `<option value="${o.value}" ${cur === o.value ? "selected" : ""}>🎙️ ${escapeHtml(o.value ? o.label : "Language: Auto (this device)")}</option>`
  ).join("")}</select>`;
}

// ---------- Drawer: fixed issues and bought materials ----------
// Each project has two drawers. "Fixed" issues (status Completed) and
// "bought" materials (archived) leave the working lists but stay here; from
// the drawer they can be restored or deleted for good (with a confirmation).

async function deleteIssueCompletely(issue) {
  const caps = await MossDB.captures.forProject(issue.projectId);
  const ids = new Set([issue.photoId, issue.voiceId].filter(Boolean));
  for (const c of caps) if (c.issueId === issue.id) ids.add(c.id);
  for (const id of ids) await MossDB.captures.remove(id);
  await MossDB.issues.remove(issue.id);
}

async function deleteMaterialCompletely(material) {
  const caps = await MossDB.captures.forProject(material.projectId);
  const ids = new Set([material.photoId, material.voiceId].filter(Boolean));
  for (const c of caps) if (c.materialId === material.id) ids.add(c.id);
  for (const id of ids) await MossDB.captures.remove(id);
  await MossDB.captures.remove(material.id);
}

async function loadDrawerItems(kind, projectId) {
  if (kind === "issue") {
    return (await MossDB.issues.forProject(projectId))
      .filter((i) => i.status === "Completed")
      .sort((a, b) => String(b.archivedAt || b.createdAt).localeCompare(String(a.archivedAt || a.createdAt)));
  }
  return (await MossDB.captures.forProject(projectId))
    .filter((c) => c.type === "material" && c.archived)
    .sort((a, b) => String(b.archivedAt || b.createdAt).localeCompare(String(a.archivedAt || a.createdAt)));
}

async function openDrawerSheet(kind, projectId) {
  const isIssue = kind === "issue";
  const items = await loadDrawerItems(kind, projectId);
  const nameOf = (it) => (isIssue ? it.title : it.name);
  const descOf = (it) =>
    isIssue
      ? [it.trade, it.requirement ? it.requirement.split("\n")[0] : ""].filter(Boolean).join(" · ")
      : [it.dimensions, it.quantity ? "Qty " + it.quantity : "", it.status || ""].filter(Boolean).join(" · ");

  openSheet(`
    <h2>🗄️ ${isIssue ? "Issues drawer" : "Materials drawer"}</h2>
    <p class="empty" style="text-align:left; margin-top:-4px;">${
      isIssue
        ? "Fixed issues. Restore sends one back to Open Issues; Delete removes it for good."
        : "Materials already bought. Restore puts one back on the list; Delete removes it for good."
    }</p>
    ${
      items.length
        ? `<div class="card card-list">${items
            .map(
              (it) => `
        <div class="row" style="cursor:default; flex-direction:column; align-items:stretch; gap:8px;">
          <span class="main"><span class="title">${escapeHtml(nameOf(it))}</span><span class="desc">${escapeHtml(descOf(it))}</span></span>
          <div style="display:flex; gap:8px;">
            <button class="btn ghost" data-dr-view="${escapeHtml(it.id)}" style="flex:1; padding:8px 6px; font-size:13px;">👁 View</button>
            <button class="btn ghost" data-dr-restore="${escapeHtml(it.id)}" style="flex:1; padding:8px 6px; font-size:13px;">↩ Restore</button>
            <button class="btn ghost" data-dr-delete="${escapeHtml(it.id)}" style="flex:1; padding:8px 6px; font-size:13px; color:var(--red, #DC2626);">🗑 Delete</button>
          </div>
        </div>`
            )
            .join("")}</div>`
        : `<p class="empty" style="text-align:center;">The drawer is empty.</p>`
    }
    <div class="sheet-actions">
      ${items.length ? `<button class="btn ghost" id="dr-empty" style="color:var(--red, #DC2626);">🗑 Empty drawer</button>` : ""}
      <button class="btn primary" id="dr-close">Close</button>
    </div>
  `);
  document.getElementById("dr-close").addEventListener("click", closeSheet);
  const find = (id) => items.find((it) => it.id === id);

  $sheet.querySelectorAll("[data-dr-view]").forEach((el) =>
    el.addEventListener("click", async () => {
      const it = find(el.dataset.drView);
      if (!it) return;
      const caps = await MossDB.captures.forProject(projectId);
      if (isIssue) openIssueDetail(it, caps);
      else openMaterialDetail(it, caps);
    })
  );
  $sheet.querySelectorAll("[data-dr-restore]").forEach((el) =>
    el.addEventListener("click", async () => {
      const it = find(el.dataset.drRestore);
      if (!it) return;
      if (isIssue) await MossDB.issues.add({ ...it, status: "Open" });
      else await MossDB.captures.update(it.id, { archived: false, archivedAt: null });
      toast(isIssue ? "Back in Open Issues" : "Back on the list");
      render();
      openDrawerSheet(kind, projectId);
    })
  );
  $sheet.querySelectorAll("[data-dr-delete]").forEach((el) =>
    el.addEventListener("click", () => {
      const it = find(el.dataset.drDelete);
      if (it) confirmDrawerDelete(kind, projectId, [it]);
    })
  );
  document.getElementById("dr-empty")?.addEventListener("click", () => confirmDrawerDelete(kind, projectId, items));
}

// Asks before anything is deleted for good.
function confirmDrawerDelete(kind, projectId, items) {
  const isIssue = kind === "issue";
  const many = items.length > 1;
  const what = many ? `all ${items.length} ${isIssue ? "issues" : "materials"}` : `"${isIssue ? items[0].title : items[0].name}"`;
  openSheet(`
    <h2>Delete ${escapeHtml(what)}?</h2>
    <p class="empty" style="text-align:left; margin-top:-4px;">
      This permanently removes ${many ? "them" : "it"}, with the photo and voice note, from the app. It can't be undone.
      Copies already saved in OneDrive are not deleted.
    </p>
    <div class="sheet-actions">
      <button class="btn ghost" id="dr-cancel-delete">Cancel</button>
      <button class="btn primary" id="dr-confirm-delete" style="background:var(--red, #DC2626);">Delete${many ? " all" : ""}</button>
    </div>
  `);
  document.getElementById("dr-cancel-delete").addEventListener("click", () => openDrawerSheet(kind, projectId));
  document.getElementById("dr-confirm-delete").addEventListener("click", async () => {
    const btn = document.getElementById("dr-confirm-delete");
    btn.disabled = true;
    try {
      for (const it of items) {
        if (isIssue) await deleteIssueCompletely(it);
        else await deleteMaterialCompletely(it);
      }
    } catch (err) {
      console.error("Deleting from the drawer failed", err);
      toast("Couldn't delete — try again");
      btn.disabled = false;
      return;
    }
    toast("Deleted");
    syncCapturesWithOneDrive(); // tell the other devices (best effort)
    render();
    openDrawerSheet(kind, projectId);
  });
}

// ---------- Clear photos / voice notes ----------
// "🗑 Clear" next to Recent Captures and Voice Notes opens a list of ALL the
// project's photos (or voice notes) with a Delete button on each, plus
// "Clear all". Photos and voice notes that belong to an issue or a material
// can be deleted one by one (the issue/material just loses that photo or
// audio), but "Clear all" leaves them alone so it can't damage an issue.

function captureOwnerLabel(c, issues, materials, inspections = []) {
  if (c.issueId) {
    const i = issues.find((x) => x.id === c.issueId);
    return i ? `issue "${i.title}"` : "an issue";
  }
  if (c.materialId) {
    const m = materials.find((x) => x.id === c.materialId);
    return m ? `material "${m.name}"` : "a material";
  }
  if (c.inspectionId) {
    const x = inspections.find((i) => i.id === c.inspectionId);
    return x ? `inspection "${x.name}"` : "an inspection";
  }
  return "";
}

// Deletes one photo / voice note and clears the link on its issue/material.
async function deleteCaptureAndUnlink(c) {
  await MossDB.captures.remove(c.id);
  const field = c.type === "photo" ? "photoId" : "voiceId";
  if (c.issueId) {
    const issue = (await MossDB.issues.forProject(c.projectId)).find((i) => i.id === c.issueId);
    if (issue && issue[field] === c.id) await MossDB.issues.add({ ...issue, [field]: null });
    else if (issue && c.type === "photo" && (issue.extraPhotoIds || []).includes(c.id)) {
      await MossDB.issues.add({ ...issue, extraPhotoIds: issue.extraPhotoIds.filter((id) => id !== c.id) });
    }
  }
  if (c.materialId) {
    const m = (await MossDB.captures.forProject(c.projectId)).find((x) => x.id === c.materialId);
    if (m && m[field] === c.id) await MossDB.captures.update(m.id, { [field]: null });
  }
  if (c.inspectionId) {
    const x = (await MossDB.captures.forProject(c.projectId)).find((i) => i.id === c.inspectionId);
    if (x) {
      if (c.type === "photo") await MossDB.captures.update(x.id, { photoIds: (x.photoIds || []).filter((id) => id !== c.id) });
      else if (x.voiceId === c.id) await MossDB.captures.update(x.id, { voiceId: null });
    }
  }
}

async function openCaptureManager(kind, projectId) {
  const isPhoto = kind === "photo";
  const [caps, issues] = await Promise.all([MossDB.captures.forProject(projectId), MossDB.issues.forProject(projectId)]);
  const materials = caps.filter((c) => c.type === "material");
  const inspections = caps.filter((c) => c.type === "inspection");
  const items = caps
    .filter((c) => c.type === kind)
    .sort((a, b) => String(b.createdAt).localeCompare(String(a.createdAt)));
  const loose = items.filter((c) => !c.issueId && !c.materialId && !c.inspectionId);
  const noun = isPhoto ? "photos" : "voice notes";

  openSheet(`
    <h2>🗑 ${isPhoto ? "Photos" : "Voice notes"}</h2>
    <p class="empty" style="text-align:left; margin-top:-4px;">Delete any one, or clear them all. Deleting asks first and can't be undone.</p>
    ${
      items.length
        ? `<div class="card card-list">${items
            .map((c) => {
              const owner = captureOwnerLabel(c, issues, materials, inspections);
              const when = c.createdAt ? new Date(c.createdAt).toLocaleString() : "";
              return `
        <div class="row" style="cursor:default; gap:12px;">
          ${
            isPhoto
              ? c.dataUrl
                ? `<img src="${c.dataUrl}" alt="" style="width:56px; height:56px; object-fit:cover; border-radius:8px; flex:none;">`
                : `<span class="icon">📷</span>`
              : `<span class="icon">🎤</span>`
          }
          <span class="main"><span class="title">${escapeHtml(isPhoto ? c.caption || "Photo" : c.transcript ? c.transcript.slice(0, 60) : "Voice note")}</span><span class="desc">${escapeHtml(when)}${owner ? " · part of " + escapeHtml(owner) : ""}</span></span>
          <button class="btn ghost" data-cm-delete="${escapeHtml(c.id)}" style="flex:none; padding:8px 10px; font-size:13px; color:var(--red, #DC2626);">🗑</button>
        </div>`;
            })
            .join("")}</div>`
        : `<p class="empty" style="text-align:center;">No ${noun}.</p>`
    }
    <div class="sheet-actions">
      ${loose.length ? `<button class="btn ghost" id="cm-clear-all" style="color:var(--red, #DC2626);">🗑 Clear all (${loose.length})</button>` : ""}
      <button class="btn primary" id="cm-close">Close</button>
    </div>
    ${items.length > loose.length && loose.length ? `<p class="empty" style="text-align:left;">"Clear all" keeps the ${items.length - loose.length} ${noun} that belong to an issue or material.</p>` : ""}
  `);
  document.getElementById("cm-close").addEventListener("click", closeSheet);
  const byId = (id) => items.find((c) => c.id === id);

  $sheet.querySelectorAll("[data-cm-delete]").forEach((el) =>
    el.addEventListener("click", () => {
      const c = byId(el.dataset.cmDelete);
      if (c) confirmCaptureDelete(kind, projectId, [c], captureOwnerLabel(c, issues, materials, inspections));
    })
  );
  document.getElementById("cm-clear-all")?.addEventListener("click", () => confirmCaptureDelete(kind, projectId, loose, ""));
}

function confirmCaptureDelete(kind, projectId, items, owner) {
  const noun = kind === "photo" ? "photo" : "voice note";
  const many = items.length > 1;
  openSheet(`
    <h2>Delete ${many ? `all ${items.length} ${noun}s` : `this ${noun}`}?</h2>
    <p class="empty" style="text-align:left; margin-top:-4px;">
      ${
        owner
          ? `This ${noun} is part of ${escapeHtml(owner)}. The ${items[0].issueId ? "issue" : items[0].inspectionId ? "inspection" : "material"} stays, but loses its ${noun}. `
          : ""
      }This permanently removes ${many ? "them" : "it"} from the app. It can't be undone. Copies already saved in OneDrive are not deleted.
    </p>
    <div class="sheet-actions">
      <button class="btn ghost" id="cm-cancel-delete">Cancel</button>
      <button class="btn primary" id="cm-confirm-delete" style="background:var(--red, #DC2626);">Delete${many ? " all" : ""}</button>
    </div>
  `);
  document.getElementById("cm-cancel-delete").addEventListener("click", () => openCaptureManager(kind, projectId));
  document.getElementById("cm-confirm-delete").addEventListener("click", async () => {
    const btn = document.getElementById("cm-confirm-delete");
    btn.disabled = true;
    try {
      for (const c of items) await deleteCaptureAndUnlink(c);
    } catch (err) {
      console.error("Deleting captures failed", err);
      toast("Couldn't delete — try again");
      btn.disabled = false;
      return;
    }
    toast("Deleted");
    syncCapturesWithOneDrive(); // tell the other devices (best effort)
    render();
    openCaptureManager(kind, projectId);
  });
}

// ---------- Voice note translation (Spanish <-> English) ----------
// Controlled by the Settings switch (aiTranslateEnabled). The translation is
// saved next to the note as `translation` + `translationLang` (the language
// the translation is IN: "en" or "es").
function translationLabel(lang) {
  return lang === "es" ? "Traducción al español" : "English translation";
}

function translationCardHtml(text, lang) {
  if (!text) return "";
  return `<div class="card" style="padding:12px 14px; font-size:14px; line-height:1.5; white-space:pre-wrap;"><span class="desc">🌐 ${escapeHtml(translationLabel(lang))}</span><br>${escapeHtml(text)}</div>`;
}

// Never throws: returns { translation, target } or null (switch off, no key,
// nothing to translate, or the AI failed / took too long).
// `force` = the person pressed "Translate again": works even if the Settings
// switch is off (it only needs the API key).
async function tryTranslate(text, force = false, previous = "") {
  if (!(force ? aiConfigured() : aiTranslateEnabled()) || !String(text || "").trim()) return null;
  try {
    return await withTimeout(translateVoiceNote(text, previous), 20000, "it took too long");
  } catch (err) {
    console.error("Translation failed", err);
    return null;
  }
}

// "Translate again" button for a screen that shows a note. Shown whenever
// there is a note and an API key (also when no translation exists yet, e.g.
// the switch was off or the first try failed).
function retranslateButtonHtml(id, hasTranslation, noteText) {
  if (!aiConfigured() || !String(noteText || "").trim()) return "";
  return `<button class="btn ghost" id="${id}" style="width:100%;">🌐 ${hasTranslation ? "Translate again" : "Translate note"}</button>`;
}

// Wires that button on a SAVED issue/material/voice note. `apply(patch)` must
// store { translation, translationLang } wherever it belongs; `done()` redraws.
function wireRetranslate(id, noteText, previous, apply, done) {
  const btn = document.getElementById(id);
  if (!btn) return;
  btn.addEventListener("click", async () => {
    const label = btn.textContent;
    btn.disabled = true;
    btn.textContent = "🌐 Translating…";
    const tr = await tryTranslate(noteText, true, previous);
    if (!tr) {
      toast("Couldn't translate — try again");
      btn.disabled = false;
      btn.textContent = label;
      return;
    }
    try {
      await apply({ translation: tr.translation, translationLang: tr.target });
    } catch (err) {
      console.error("Saving translation failed", err);
      toast("Couldn't save the translation — try again");
      btn.disabled = false;
      btn.textContent = label;
      return;
    }
    toast("Translation updated");
    done();
    syncCapturesWithOneDrive();
  });
}

// ---------- Materials: photo + AI online lookup + voice note ----------
// Flow: take a photo → record a voice note → ONE AI call combines the photo
// and what was said, searches online, and fills the name, size, quantity and
// a written note → save. The photo and voice note are also saved as normal
// captures linked to the material (materialId), so they sync to OneDrive like
// any other photo / voice note. Typing a material in by hand still works.

const MATERIAL_STATUSES = ["To buy", "Ordered", "Delivered", "Backordered"];

function withTimeout(promise, ms, message) {
  let timer;
  const timeout = new Promise((_, reject) => {
    timer = setTimeout(() => reject(new Error(message)), ms);
  });
  return Promise.race([promise, timeout]).finally(() => clearTimeout(timer));
}

function describeMaterialLookup(r) {
  const parts = [];
  if (r.searched) {
    parts.push(r.sources.length ? `🔎 Checked online (${r.sources.join(", ")})` : "🔎 Checked online");
  } else {
    parts.push("⚠️ Couldn't search online (web search isn't enabled for this API key's organization) — identified from the photo only, so double-check the size");
  }
  if (r.confidence === "low") parts.push("Not sure about this one — please check the name and size");
  else if (r.confidence === "medium") parts.push("Fairly sure — worth a quick check");
  if (r.details) parts.push(r.details);
  return parts.join(" · ");
}

function openMaterialSheet(projectId) {
  const state = {
    photo: null, // { dataUrl, file }
    voice: null, // { blob, dataUrl, mimeType, transcript, startedAt }
    recorder: null,
    recording: false,
    identifying: false, // AI looking at photo + voice note, searching online
    aiRan: false,
    aiNote: "",
    translation: "",
    translationLang: "",
    translatedFrom: "",
    translating: false,
    item: "",
    dims: "",
    qty: "",
    status: "To buy",
    note: "",
    itemTouched: false, // once the user edits a field by hand, AI never overwrites it
    dimsTouched: false,
    qtyTouched: false,
    noteTouched: false,
    searched: false,
    sources: [],
    saving: false
  };
  let closed = false;
  let aiRun = 0; // bumped per analysis so a stale result can't overwrite a newer one
  let translateRun = 0;

  const draw = () => {
    if (closed) return;
    const busy = state.identifying || state.translating;
    const canSave = !!state.item.trim() && !state.recording && !busy && !state.saving;

    const photoBlock = `
      ${state.photo ? `<img class="issue-photo" src="${state.photo.dataUrl}" alt="Material photo">` : ""}
      <button class="btn ${state.photo ? "ghost" : "primary"}" id="m-photo" ${state.recording || state.saving ? "disabled" : ""}>${state.photo ? "📸 Retake photo" : "📸 Take photo"}</button>
      <button class="btn ghost" id="m-gallery" ${state.recording || state.saving ? "disabled" : ""} style="margin-top:8px;">🖼 ${state.photo ? "Choose another from gallery" : "Choose from gallery"}</button>
    `;

    let voiceBlock;
    if (!state.photo) {
      voiceBlock = `<button class="btn ghost" disabled>🎤 Take the photo first</button>`;
    } else if (state.recording) {
      voiceBlock = `
        <div class="rec-indicator"><span class="dot"></span><span>Recording… say what the material is and how many you need</span></div>
        <button class="btn primary" id="m-stop">⏹ Stop recording</button>
      `;
    } else if (state.voice) {
      voiceBlock = `
        <audio controls preload="metadata" style="width:100%; height:36px;" src="${state.voice.dataUrl}"></audio>
        ${voiceLangPickerHtml("m-vlang", state.saving || state.identifying)}
        <button class="btn ghost" id="m-rec" ${state.saving || state.identifying ? "disabled" : ""}>🎤 Re-record</button>
      `;
    } else {
      voiceBlock = `${voiceLangPickerHtml("m-vlang", state.saving)}<button class="btn primary" id="m-rec" ${state.saving ? "disabled" : ""}>🎤 Record voice note</button>`;
    }

    let aiBlock = "";
    if (state.photo && !state.recording) {
      if (state.identifying) {
        aiBlock = `<p class="empty" style="text-align:left;">🤖 AI is combining the photo${state.voice ? " and your voice note" : ""} and searching online… this can take up to a minute.</p>`;
      } else {
        const hint = !state.voice && !state.aiRan ? `<p class="empty" style="text-align:left;">Next: record a voice note — the AI will combine it with the photo.</p>` : "";
        const label = state.aiRan ? "🤖 Run AI again" : state.voice ? "🤖 Find name & size with AI" : "🤖 Identify from photo only";
        aiBlock = `
          ${hint}
          ${state.aiNote ? `<p class="empty" style="text-align:left;">${escapeHtml(state.aiNote)}</p>` : ""}
          <button class="btn ghost" id="m-ai" ${state.saving ? "disabled" : ""}>${label}</button>
        `;
      }
    }

    const noteBlock =
      state.voice && !state.recording
        ? `
        <div class="field">
          <label>Note · ${escapeHtml(issueStamp(state.voice.startedAt))}</label>
          <textarea id="f-note" placeholder="${state.voice.transcript ? "" : "Couldn't transcribe this voice note — type what you said (optional)"}">${escapeHtml(state.note)}</textarea>
        </div>
        ${state.translating ? `<p class="empty" style="text-align:left;">🌐 AI is translating the note…</p>` : translationCardHtml(state.translation, state.translationLang) + retranslateButtonHtml("m-retr", !!state.translation, state.note)}
      `
        : "";

    $sheet.innerHTML = `
      <h2>Material</h2>
      <div class="field"><label>1 · Photo</label>${photoBlock}</div>
      <div class="field"><label>2 · Voice note</label>${voiceBlock}</div>
      <div class="field"><label>3 · AI</label>${aiBlock || `<p class="empty" style="text-align:left;">Takes the photo + voice note and finds the name and size online.</p>`}</div>
      <div class="field"><label>Item</label><input id="f-item" value="${escapeHtml(state.item)}" placeholder='e.g. "2x4x8 stud" or "Hard pipe duct, 6&quot;"'></div>
      <div class="field"><label>Size / dimensions</label><input id="f-dims" value="${escapeHtml(state.dims)}" placeholder='e.g. 1.5" x 3.5" x 96"'></div>
      <div class="field"><label>Quantity</label><input id="f-qty" value="${escapeHtml(state.qty)}" placeholder='e.g. 12, or "2 boxes"'></div>
      <div class="field"><label>Status</label>
        <select id="f-status">${MATERIAL_STATUSES.map((s) => `<option ${s === state.status ? "selected" : ""}>${s}</option>`).join("")}</select>
      </div>
      ${noteBlock}
      <div class="sheet-actions">
        <button class="btn ghost" id="m-cancel">Cancel</button>
        <button class="btn primary" id="m-save" ${canSave ? "" : "disabled"} style="${canSave ? "" : "opacity:.5;"}">${state.saving ? "Saving…" : busy ? "Working…" : "Save Material"}</button>
      </div>
    `;
    wire();
  };

  const wire = () => {
    const $ = (id) => document.getElementById(id);
    $("m-cancel")?.addEventListener("click", closeSheet);
    $("m-photo")?.addEventListener("click", () => takePhoto());
    $("m-gallery")?.addEventListener("click", () => takePhoto(true));
    $("m-rec")?.addEventListener("click", startRecording);
    $("m-stop")?.addEventListener("click", stopRecording);
    $("m-ai")?.addEventListener("click", runIdentify);
    $("m-retr")?.addEventListener("click", () => runTranslate(true));
    $("m-vlang")?.addEventListener("change", (e) => setVoiceLang(e.target.value));
    $("m-save")?.addEventListener("click", save);
    // Keep state in sync as the user types, so redraws never lose edits. The
    // Save button is updated in place (no redraw while typing).
    const syncSave = () => {
      const btn = $("m-save");
      if (!btn) return;
      const ok = !!state.item.trim() && !state.recording && !state.identifying && !state.translating && !state.saving;
      btn.disabled = !ok;
      btn.style.opacity = ok ? "" : ".5";
    };
    $("f-item")?.addEventListener("input", (e) => { state.item = e.target.value; state.itemTouched = true; syncSave(); });
    $("f-dims")?.addEventListener("input", (e) => { state.dims = e.target.value; state.dimsTouched = true; });
    $("f-qty")?.addEventListener("input", (e) => { state.qty = e.target.value; state.qtyTouched = true; });
    $("f-status")?.addEventListener("change", (e) => { state.status = e.target.value; });
    $("f-note")?.addEventListener("input", (e) => { state.note = e.target.value; state.noteTouched = true; });
  };

  const takePhoto = (fromGallery = false) => {
    const input = document.createElement("input");
    input.type = "file";
    input.accept = "image/*";
    if (!fromGallery) input.capture = "environment";
    input.addEventListener("change", async () => {
      const file = input.files[0];
      if (!file || closed) return;
      let dataUrl;
      try {
        dataUrl = await photoFileToDataUrl(file);
      } catch {
        toast("Couldn't read that photo");
        return;
      }
      if (closed) return;
      state.photo = { dataUrl, file };
      state.aiNote = "";
      state.aiRan = false;
      draw();
      // A new photo with a voice note already recorded: re-run the combined
      // lookup. Otherwise wait for the voice note (step 2).
      if (state.voice) runIdentify();
    });
    input.click();
  };

  // ONE AI call that combines the photo and the voice-note transcript.
  const runIdentify = async () => {
    if (!state.photo || state.identifying) return;
    if (!aiConfigured()) {
      state.aiNote = "Add a Claude API key in Settings to have AI find the name and size online. Fill it in by hand for now.";
      draw();
      return;
    }
    const run = ++aiRun;
    state.identifying = true;
    state.aiNote = "";
    draw();
    const transcript = state.voice ? state.voice.transcript || "" : "";
    try {
      const r = await withTimeout(identifyMaterial(state.photo.dataUrl, transcript), 60000, "it took too long");
      if (closed || run !== aiRun) return;
      if (!state.itemTouched && r.item) state.item = r.item;
      if (!state.dimsTouched && r.dimensions) state.dims = r.dimensions;
      if (!state.qtyTouched && r.quantity) state.qty = r.quantity;
      if (state.voice && !state.noteTouched && r.note) state.note = r.note;
      state.searched = r.searched;
      state.sources = r.sources;
      state.aiNote = describeMaterialLookup(r);
    } catch (err) {
      if (closed || run !== aiRun) return;
      state.aiNote = `AI couldn't identify this one (${err.message}). Type the item by hand.`;
    }
    state.aiRan = true;
    state.identifying = false;
    draw();
    runTranslate();
  };

  // Spanish <-> English translation of the final note (only when switched on).
  // `force` = "Translate again" was pressed: ignores the Settings switch and
  // keeps the old translation if the new try fails.
  const runTranslate = async (force = false) => {
    const run = ++translateRun;
    const text = state.voice ? state.note.trim() : "";
    if (!(force ? aiConfigured() : aiTranslateEnabled()) || !text) {
      state.translating = false;
      state.translation = "";
      state.translatedFrom = "";
      draw();
      return;
    }
    state.translating = true;
    draw();
    const tr = await tryTranslate(text, force, force ? state.translation : "");
    if (closed || run !== translateRun) return;
    if (tr) {
      state.translation = tr.translation;
      state.translationLang = tr.target;
    } else if (force) {
      toast("Couldn't translate — try again");
    } else {
      state.translation = "";
      state.translationLang = "";
    }
    state.translatedFrom = text;
    state.translating = false;
    draw();
  };

  const startRecording = async () => {
    try {
      state.recorder = await startVoiceRecorder();
    } catch (err) {
      toast(err.message === "Microphone not available in this browser" ? err.message : "Microphone permission denied");
      return;
    }
    if (closed) {
      state.recorder.cancel();
      state.recorder = null;
      return;
    }
    state.recording = true;
    draw();
  };

  const stopRecording = async () => {
    const rec = state.recorder;
    if (!rec) return;
    state.recorder = null;
    const stopBtn = document.getElementById("m-stop");
    if (stopBtn) { stopBtn.disabled = true; stopBtn.textContent = "Finishing…"; }
    let result = null;
    try {
      result = await rec.stop();
    } catch (err) {
      console.error("Stopping recording failed", err);
    }
    state.recording = false;
    if (closed) return;
    if (!result) {
      toast("Nothing was recorded — try again");
      draw();
      return;
    }
    state.voice = result;
    state.note = result.transcript; // the AI rewrites this below; stays as-is if the AI fails
    state.noteTouched = false;
    state.aiNote = "";
    state.aiRan = false;
    state.translation = "";
    state.translatedFrom = "";
    translateRun++;
    state.translating = false;
    draw();
    runIdentify(); // photo + voice note → name, size, quantity, note
  };

  const save = async () => {
    const item = state.item.trim();
    if (!item) {
      toast("Add an item name");
      return;
    }
    if (state.recording || state.identifying || state.translating || state.saving) return;
    state.saving = true;
    draw();
    try {
      // The note may have been edited after it was translated: refresh the
      // translation so the saved pair always matches (skipped if it fails).
      if (state.voice && aiTranslateEnabled() && state.note.trim() && state.note.trim() !== state.translatedFrom) {
        const tr = await tryTranslate(state.note.trim());
        state.translation = tr ? tr.translation : "";
        state.translationLang = tr ? tr.target : "";
        state.translatedFrom = state.note.trim();
      } else if (!state.note.trim() || !state.voice || state.note.trim() !== state.translatedFrom) {
        state.translation = "";
      }
      const now = Date.now();
      const materialId = MossDB.uid();
      const photoId = state.photo ? MossDB.uid() : null;
      const voiceId = state.voice ? MossDB.uid() : null;
      const dims = state.dims.trim();
      const note = state.note.trim();
      const photoFile = state.photo ? state.photo.file : null;
      const voiceBlob = state.voice ? state.voice.blob : null;
      // Unique file names: phones often name every camera photo "image.jpg",
      // which would overwrite earlier photos in OneDrive.
      const extOfPhoto = photoFile
        ? { "image/png": "png", "image/heic": "heic", "image/heif": "heif", "image/webp": "webp" }[photoFile.type] || "jpg"
        : "jpg";
      const photoName = `material-photo-${now}.${extOfPhoto}`;
      const voiceName = state.voice ? `material-voice-${now}.${state.voice.mimeType.includes("mp4") ? "m4a" : "webm"}` : "";

      if (state.photo) {
        await MossDB.captures.add({
          id: photoId,
          projectId,
          type: "photo",
          dataUrl: state.photo.dataUrl,
          name: state.photo.file.name,
          remoteFileName: photoName,
          caption: dims ? `${item} — ${dims}` : item,
          materialId
        });
      }
      if (state.voice) {
        await MossDB.captures.add({
          id: voiceId,
          projectId,
          type: "voice",
          dataUrl: state.voice.dataUrl,
          transcript: note || state.voice.transcript || null,
          remoteFileName: voiceName,
          materialId,
          ...(state.translation ? { translation: state.translation, translationLang: state.translationLang } : {})
        });
      }
      await MossDB.captures.add({
        id: materialId,
        projectId,
        type: "material",
        name: item,
        dimensions: dims,
        quantity: state.qty.trim(),
        status: MATERIAL_STATUSES.includes(state.status) ? state.status : "To buy",
        note,
        recordedAt: state.voice ? state.voice.startedAt.toISOString() : null,
        photoId,
        voiceId,
        searchedOnline: state.searched,
        sources: state.sources,
        ...(state.translation ? { translation: state.translation, translationLang: state.translationLang } : {})
      });

      closeSheet();
      toast("Material saved");
      if (currentRoute().name === "project" || currentRoute().name === "home") render();

      // Best-effort OneDrive copy of the files; announce them to other
      // devices only once everything has actually finished uploading.
      const uploads = [];
      if (photoFile) uploads.push(syncCaptureToOneDrive(projectId, "photo", photoFile, photoName));
      if (voiceBlob) uploads.push(syncCaptureToOneDrive(projectId, "voice", voiceBlob, voiceName));
      if (uploads.length) Promise.all(uploads).then(() => syncCapturesWithOneDrive());
    } catch (err) {
      console.error("Saving material failed", err);
      state.saving = false;
      toast("Couldn't save the material — try again");
      draw();
    }
  };

  openSheet("", () => {
    // Runs whenever the sheet goes away (Cancel, backdrop tap, navigation,
    // or after a successful save): always release the microphone.
    closed = true;
    aiRun++;
    translateRun++;
    if (state.recorder) {
      state.recorder.cancel();
      state.recorder = null;
    }
  });
  draw();
}

// Opens a saved material with its photo, specs, written note and voice note.
function openMaterialDetail(material, captures) {
  const photo = captures.find((c) => c.id === material.photoId);
  const voice = captures.find((c) => c.id === material.voiceId);
  const facts = [material.dimensions, material.quantity ? `Qty ${material.quantity}` : "", material.status].filter(Boolean);
  openSheet(`
    <h2>${escapeHtml(material.name)}</h2>
    ${facts.length ? `<span class="desc">${facts.map(escapeHtml).join(" · ")}</span>` : ""}
    ${photo?.dataUrl ? `<img class="issue-photo" src="${photo.dataUrl}" alt="Material photo">` : ""}
    ${
      material.note
        ? `<div class="card" style="padding:12px 14px; font-size:14px; line-height:1.5; white-space:pre-wrap;">${
            material.recordedAt ? `<span class="desc">${escapeHtml(issueStamp(new Date(material.recordedAt)))}</span><br>` : ""
          }${escapeHtml(material.note)}</div>`
        : ""
    }
    ${translationCardHtml(material.translation, material.translationLang)}
    ${retranslateButtonHtml("d-retr", !!material.translation, material.note)}
    ${voice?.dataUrl ? `<audio controls preload="metadata" style="width:100%; height:36px;" src="${voice.dataUrl}"></audio>` : ""}
    ${material.sources && material.sources.length ? `<span class="desc">Checked online: ${material.sources.map(escapeHtml).join(", ")}</span>` : ""}
    <button class="btn ghost" id="d-archive" style="width:100%;">${material.archived ? "↩ Put back on the list" : "🗄️ Bought — move to drawer"}</button>
    ${closeButton()}
  `);
  wireCloseButton();
  document.getElementById("d-archive")?.addEventListener("click", async () => {
    const patch = material.archived ? { archived: false, archivedAt: null } : { archived: true, archivedAt: new Date().toISOString() };
    await MossDB.captures.update(material.id, patch);
    closeSheet();
    toast(material.archived ? "Back on the list" : "Moved to the drawer");
    render();
  });
  wireRetranslate(
    "d-retr",
    material.note || "",
    material.translation || "",
    async (patch) => {
      Object.assign(material, patch);
      await MossDB.captures.update(material.id, patch);
      if (voice) {
        Object.assign(voice, patch);
        await MossDB.captures.update(voice.id, patch);
      }
    },
    () => openMaterialDetail(material, captures)
  );
}

// Shopping-list PDF: pick which materials to include (the ones marked "To
// buy" start checked), then build, then share/open — same two-step flow as
// the issues PDF. `projectId` limits it to one project; null = all projects.
async function openMaterialsPdfSheet(projectId) {
  if (!window.jspdf) {
    toast("PDF library missing — upload jspdf.umd.min.js to GitHub");
    return;
  }
  const projects = (await MossDB.projects.all()).filter(isVisibleProject).filter((p) => !projectId || p.id === projectId);
  const groups = [];
  for (const p of projects) {
    const captures = await MossDB.captures.forProject(p.id);
    const materials = captures
      .filter((c) => c.type === "material" && !c.archived) // bought ones are in the drawer
      .sort((a, b) => String(a.createdAt).localeCompare(String(b.createdAt)));
    if (materials.length) groups.push({ project: p, captures, materials });
  }
  if (!groups.length) {
    toast("No materials to put in a list yet");
    return;
  }

  openSheet(`
    <h2>Shopping list PDF</h2>
    <p class="empty" style="text-align:left; margin-top:-4px;">Choose what goes on the list. Items marked "To buy" start checked.</p>
    ${groups
      .map(
        (g) => `
      ${projectId ? "" : `<div class="section-label">${escapeHtml(g.project.name)}</div>`}
      <div class="card card-list">
        ${g.materials
          .map(
            (m) => `
          <label class="row" style="cursor:pointer; gap:12px;">
            <input type="checkbox" data-mid="${escapeHtml(m.id)}" ${m.status === "To buy" ? "checked" : ""} style="width:20px; height:20px; flex:none;">
            <span class="main"><span class="title">${escapeHtml(m.name)}</span><span class="desc">${[m.dimensions, m.quantity ? "Qty " + m.quantity : "", m.status || ""].filter(Boolean).map(escapeHtml).join(" · ")}${m.photoId ? " · 📷" : ""}</span></span>
          </label>`
          )
          .join("")}
      </div>`
      )
      .join("")}
    <div class="sheet-actions">
      <button class="btn ghost" id="pdf-cancel">Cancel</button>
      <button class="btn primary" id="pdf-create">Create PDF</button>
    </div>
  `);
  document.getElementById("pdf-cancel").addEventListener("click", closeSheet);
  document.getElementById("pdf-create").addEventListener("click", async () => {
    const chosen = new Set([...$sheet.querySelectorAll("[data-mid]")].filter((el) => el.checked).map((el) => el.dataset.mid));
    if (!chosen.size) {
      toast("Select at least one material");
      return;
    }
    const btn = document.getElementById("pdf-create");
    btn.disabled = true;
    btn.textContent = "Building PDF…";
    try {
      const sections = [];
      for (const g of groups) {
        const picked = g.materials.filter((m) => chosen.has(m.id));
        if (!picked.length) continue;
        sections.push({
          projectName: g.project.name,
          address: g.project.address || "",
          items: picked.map((m) => ({
            name: m.name,
            dimensions: m.dimensions || "",
            quantity: m.quantity || "",
            status: m.status || "",
            note: m.note || "",
            translation: m.translation || "",
            translationLabel: m.translation ? translationLabel(m.translationLang) : "",
            stamp: m.recordedAt ? issueStamp(new Date(m.recordedAt)) : "",
            photoDataUrl: g.captures.find((c) => c.id === m.photoId)?.dataUrl || null
          }))
        });
      }
      const blob = await buildMaterialsPdf(sections);
      const day = new Date().toISOString().slice(0, 10);
      const label = sections.length === 1 ? sections[0].projectName : "All projects";
      const fileName = `Shopping list - ${label} - ${day}.pdf`.replace(/[\\/:*?"<>|]+/g, "").replace(/\s+/g, " ");
      showPdfReady(blob, fileName, sections.reduce((n, s) => n + s.items.length, 0), "item");
    } catch (err) {
      console.error("Materials PDF build failed", err);
      toast("Couldn't build the PDF — try again");
      btn.disabled = false;
      btn.textContent = "Create PDF";
    }
  });
}

// ---------- Inspections ----------
// Walk an inspection with the inspector while recording a voice note. The AI
// writes the note up word for word, names the inspection, and (if switched
// on) translates the note Spanish <-> English; photos are optional; the
// result (Passed / Failed / Scheduled) is picked at the bottom. A saved
// inspection can be turned into a PDF report and shared.
// Stored as a capture of type "inspection": { name (title), result, comments
// (the note), inspector, recordedAt, photoIds[], voiceId, translation,
// translationLang }. Its photos / voice note are normal captures that carry
// inspectionId, so they sync to OneDrive like any other.

const INSPECTION_RESULTS = [
  { value: "Passed", label: "✅ Passed", short: "Passed", icon: "✅" },
  { value: "Failed / corrections required", label: "❌ Failed", short: "Failed", icon: "❌" },
  { value: "Scheduled", label: "🗓 Scheduled", short: "Scheduled", icon: "🗓" }
];

function inspectionResultInfo(value) {
  return INSPECTION_RESULTS.find((r) => r.value === value) || { value, short: value || "", icon: "🏛️" };
}

const MAX_INSPECTION_PHOTOS = 15;

async function openInspectionSheet(preProjectId) {
  const visible = (await MossDB.projects.all()).filter(isVisibleProject);
  const projects = [...visible.filter(isActiveProject), ...visible.filter((p) => !isActiveProject(p))];
  const state = {
    projectId: preProjectId && projects.some((p) => p.id === preProjectId) ? preProjectId : "",
    photos: [], // [{ dataUrl, file }]
    voice: null, // { blob, dataUrl, mimeType, transcript, startedAt }
    recorder: null,
    recording: false,
    analyzing: false,
    aiNote: "",
    aiRan: false,
    title: "",
    titleLang: issueTitleLang(),
    titleBusy: false,
    titleTouched: false,
    note: "",
    noteTouched: false,
    translation: "",
    translationLang: "",
    translatedFrom: "",
    translating: false,
    inspector: "",
    result: "",
    saving: false
  };
  let closed = false;
  let analysisRun = 0;
  let translateRun = 0;
  let titleRun = 0;

  const busyNow = () => state.analyzing || state.translating || state.titleBusy;

  const draw = () => {
    if (closed) return;
    const haveProject = !!state.projectId;
    const canSave = haveProject && !!state.result && !state.recording && !busyNow() && !state.saving;

    const projectBlock = `
      <select id="in-project" ${state.recording || state.saving ? "disabled" : ""}>
        ${haveProject ? "" : `<option value="">Choose a project…</option>`}
        ${projects.map((p) => `<option value="${escapeHtml(p.id)}" ${p.id === state.projectId ? "selected" : ""}>${escapeHtml(p.name)}</option>`).join("")}
      </select>`;

    let voiceBlock;
    if (!haveProject) {
      voiceBlock = `<button class="btn ghost" disabled>🎤 Choose the project first</button>`;
    } else if (state.recording) {
      voiceBlock = `
        <div class="rec-indicator"><span class="dot"></span><span>Recording the inspection… talk through it with the inspector</span></div>
        <button class="btn primary" id="in-stop">⏹ Stop recording</button>`;
    } else if (state.voice) {
      voiceBlock = `
        <audio controls preload="metadata" style="width:100%; height:36px;" src="${state.voice.dataUrl}"></audio>
        ${voiceLangPickerHtml("in-vlang", state.saving || busyNow())}
        <button class="btn ghost" id="in-rec" ${state.saving || busyNow() ? "disabled" : ""}>🎤 Re-record</button>`;
    } else {
      voiceBlock = `${voiceLangPickerHtml("in-vlang", state.saving)}<button class="btn primary" id="in-rec" ${state.saving ? "disabled" : ""}>🎤 Start recording</button>`;
    }

    let noteBlock = "";
    if (state.voice && !state.recording) {
      noteBlock = `
        <div class="field">
          <label>Note · ${escapeHtml(issueStamp(state.voice.startedAt))}</label>
          <textarea id="in-note" rows="5" placeholder="${state.voice.transcript ? "" : "Couldn't transcribe this voice note — type what was said (optional)"}">${escapeHtml(state.note)}</textarea>
        </div>
        ${state.translating ? `<p class="empty" style="text-align:left;">🌐 AI is translating the note…</p>` : translationCardHtml(state.translation, state.translationLang) + retranslateButtonHtml("in-retr", !!state.translation, state.note)}
        ${state.analyzing ? `<p class="empty" style="text-align:left;">🤖 AI is writing up the note and naming the inspection…</p>` : ""}
        ${state.aiNote ? `<p class="empty" style="text-align:left;">${escapeHtml(state.aiNote)}</p>` : ""}
        ${aiConfigured() && !state.analyzing ? `<button class="btn ghost" id="in-ai" ${state.saving || state.translating ? "disabled" : ""}>🤖 ${state.aiRan ? "Run AI again" : "Name it with AI"}</button>` : ""}`;
    }

    const photosBlock = haveProject
      ? `
        ${
          state.photos.length
            ? `<div style="display:flex; flex-wrap:wrap; gap:8px; margin-bottom:8px;">${state.photos
                .map(
                  (p, i) => `<div style="position:relative;"><img src="${p.dataUrl}" alt="Inspection photo ${i + 1}" style="width:84px; height:84px; object-fit:cover; border-radius:10px; display:block;"><button class="btn ghost" data-rm-photo="${i}" ${state.saving ? "disabled" : ""} style="position:absolute; top:-6px; right:-6px; width:26px; height:26px; padding:0; border-radius:50%; font-size:13px; line-height:1;" aria-label="Remove photo">✕</button></div>`
                )
                .join("")}</div>`
            : ""
        }
        <button class="btn ghost" id="in-photo" ${state.recording || state.saving || state.photos.length >= MAX_INSPECTION_PHOTOS ? "disabled" : ""}>📸 ${state.photos.length ? "Add another photo" : "Add photo"} ${state.photos.length ? `(${state.photos.length}/${MAX_INSPECTION_PHOTOS})` : ""}</button>
        <button class="btn ghost" id="in-gallery" ${state.recording || state.saving || state.photos.length >= MAX_INSPECTION_PHOTOS ? "disabled" : ""} style="margin-top:8px;">🖼 Choose from gallery</button>`
      : `<button class="btn ghost" disabled>📸 Choose the project first</button>`;

    $sheet.innerHTML = `
      <h2>Inspection</h2>
      <div class="field"><label>1 · Project</label>${projectBlock}</div>
      <div class="field"><label>2 · Voice note — record during the inspection</label>${voiceBlock}</div>
      ${noteBlock}
      <div class="field"><label>3 · Photos (optional)</label>${photosBlock}</div>
      <div class="field">
        <label>Inspection · title language</label>
        <select id="in-title-lang" ${busyNow() || state.saving ? "disabled" : ""}>
          <option value="en" ${state.titleLang === "en" ? "selected" : ""}>English</option>
          <option value="es" ${state.titleLang === "es" ? "selected" : ""}>Español</option>
        </select>
      </div>
      <div class="field"><label>Inspection (type / title)</label><input id="in-title" value="${escapeHtml(state.title)}" placeholder='${state.titleLang === "es" ? "p. ej. &quot;Inspección de estructura, segundo piso&quot;" : "e.g. &quot;Framing inspection, second floor&quot;"}'></div>
      ${state.titleBusy ? `<p class="empty" style="text-align:left;">🌐 ${state.titleLang === "es" ? "Traduciendo el título…" : "Translating the title…"}</p>` : ""}
      <div class="field"><label>Inspector (optional)</label><input id="in-inspector" value="${escapeHtml(state.inspector)}" placeholder="Name"></div>
      <div class="field">
        <label>Result</label>
        <div style="display:flex; gap:8px;">
          ${INSPECTION_RESULTS.map(
            (r) => `<button class="btn ${state.result === r.value ? "primary" : "ghost"}" data-result="${escapeHtml(r.value)}" aria-pressed="${state.result === r.value}" ${state.saving ? "disabled" : ""} style="flex:1; padding:12px 6px; font-size:14px;">${r.label}</button>`
          ).join("")}
        </div>
      </div>
      <div class="sheet-actions">
        <button class="btn ghost" id="in-cancel">Cancel</button>
        <button class="btn primary" id="in-save" ${canSave ? "" : "disabled"} style="${canSave ? "" : "opacity:.5;"}">${state.saving ? "Saving…" : busyNow() ? "Working…" : "Save"}</button>
      </div>
      <div class="sheet-actions">
        <button class="btn ghost" id="in-save-pdf" ${canSave ? "" : "disabled"} style="width:100%; ${canSave ? "" : "opacity:.5;"}">📄 Save &amp; create PDF</button>
      </div>`;
    wire();
  };

  const wire = () => {
    const $ = (id) => document.getElementById(id);
    $("in-cancel")?.addEventListener("click", closeSheet);
    $("in-project")?.addEventListener("change", (e) => { state.projectId = e.target.value; draw(); });
    $("in-rec")?.addEventListener("click", startRecording);
    $("in-stop")?.addEventListener("click", stopRecording);
    $("in-photo")?.addEventListener("click", () => takePhoto());
    $("in-gallery")?.addEventListener("click", () => takePhoto(true));
    $("in-ai")?.addEventListener("click", runAnalysis);
    $("in-retr")?.addEventListener("click", () => runTranslate(true));
    $("in-vlang")?.addEventListener("change", (e) => setVoiceLang(e.target.value));
    $("in-title-lang")?.addEventListener("change", (e) => changeTitleLang(e.target.value));
    $("in-save")?.addEventListener("click", () => save(false));
    $("in-save-pdf")?.addEventListener("click", () => save(true));
    $sheet.querySelectorAll("[data-rm-photo]").forEach((el) =>
      el.addEventListener("click", () => { state.photos.splice(Number(el.dataset.rmPhoto), 1); draw(); })
    );
    $sheet.querySelectorAll("[data-result]").forEach((el) =>
      el.addEventListener("click", () => { state.result = el.dataset.result; draw(); })
    );
    // Keep state in sync while typing (no redraw, so the caret never jumps);
    // Save buttons are updated in place.
    $("in-note")?.addEventListener("input", (e) => { state.note = e.target.value; state.noteTouched = true; });
    $("in-title")?.addEventListener("input", (e) => { state.title = e.target.value; state.titleTouched = true; });
    $("in-inspector")?.addEventListener("input", (e) => { state.inspector = e.target.value; });
  };

  const takePhoto = (fromGallery = false) => {
    const input = document.createElement("input");
    input.type = "file";
    input.accept = "image/*";
    if (!fromGallery) input.capture = "environment";
    input.addEventListener("change", async () => {
      const file = input.files[0];
      if (!file || closed) return;
      let dataUrl;
      try {
        dataUrl = await photoFileToDataUrl(file);
      } catch {
        toast("Couldn't read that photo");
        return;
      }
      if (closed || state.photos.length >= MAX_INSPECTION_PHOTOS) return;
      state.photos.push({ dataUrl, file });
      draw();
    });
    input.click();
  };

  const startRecording = async () => {
    try {
      state.recorder = await startVoiceRecorder();
    } catch (err) {
      toast(err.message === "Microphone not available in this browser" ? err.message : "Microphone permission denied");
      return;
    }
    if (closed) {
      state.recorder.cancel();
      state.recorder = null;
      return;
    }
    state.recording = true;
    draw();
  };

  const stopRecording = async () => {
    const rec = state.recorder;
    if (!rec) return;
    state.recorder = null;
    const stopBtn = document.getElementById("in-stop");
    if (stopBtn) { stopBtn.disabled = true; stopBtn.textContent = "Finishing…"; }
    let result = null;
    try {
      result = await rec.stop();
    } catch (err) {
      console.error("Stopping recording failed", err);
    }
    state.recording = false;
    if (closed) return;
    if (!result) {
      toast("Nothing was recorded — try again");
      draw();
      return;
    }
    state.voice = result;
    state.note = result.transcript; // the AI tidies it below; stays as-is if the AI fails
    state.noteTouched = false;
    state.aiNote = "";
    state.aiRan = false;
    state.translation = "";
    state.translatedFrom = "";
    translateRun++;
    state.translating = false;
    draw();
    runAnalysis();
  };

  // One AI call: names the inspection and writes the note word for word.
  const runAnalysis = async () => {
    if (!state.voice || state.analyzing) return;
    if (!aiConfigured()) {
      state.aiNote = "Add a Claude API key in Settings to have AI write up the note and name the inspection. Type the title by hand for now.";
      draw();
      return;
    }
    const run = ++analysisRun;
    state.analyzing = true;
    state.aiNote = "";
    draw();
    try {
      const r = await withTimeout(
        analyzeInspectionCapture(state.photos.map((p) => p.dataUrl), state.voice.transcript, state.titleLang),
        60000,
        "it took too long"
      );
      if (closed || run !== analysisRun) return;
      if (!state.titleTouched && r.title) state.title = r.title;
      if (!state.noteTouched && r.note) state.note = r.note;
    } catch (err) {
      if (closed || run !== analysisRun) return;
      state.aiNote = `AI couldn't analyze this one (${err.message}). Type the title by hand.`;
    }
    state.aiRan = true;
    state.analyzing = false;
    draw();
    runTranslate();
  };

  // Spanish <-> English translation of the note (only when switched on, or
  // when "Translate again" is pressed).
  const runTranslate = async (force = false) => {
    const run = ++translateRun;
    const text = state.voice ? state.note.trim() : "";
    if (!(force ? aiConfigured() : aiTranslateEnabled()) || !text) {
      state.translating = false;
      state.translation = "";
      state.translatedFrom = "";
      draw();
      return;
    }
    state.translating = true;
    draw();
    const tr = await tryTranslate(text, force, force ? state.translation : "");
    if (closed || run !== translateRun) return;
    if (tr) {
      state.translation = tr.translation;
      state.translationLang = tr.target;
    } else if (force) {
      toast("Couldn't translate — try again");
    } else {
      state.translation = "";
      state.translationLang = "";
    }
    state.translatedFrom = text;
    state.translating = false;
    draw();
  };

  const changeTitleLang = async (lang) => {
    lang = lang === "es" ? "es" : "en";
    if (lang === state.titleLang) return;
    state.titleLang = lang;
    setIssueTitleLang(lang);
    const current = state.title.trim();
    if (!current || !aiConfigured()) {
      draw();
      return;
    }
    const run = ++titleRun;
    state.titleBusy = true;
    draw();
    try {
      const t = await withTimeout(translateIssueTitle(current, lang), 20000, "it took too long");
      if (closed || run !== titleRun) return;
      if (state.title.trim() === current) state.title = t;
    } catch {
      if (closed || run !== titleRun) return;
      toast("Couldn't translate the title — edit it by hand");
    }
    state.titleBusy = false;
    draw();
  };

  const save = async (thenPdf) => {
    if (!state.projectId || !state.result || state.recording || busyNow() || state.saving) return;
    state.saving = true;
    draw();
    try {
      // The note may have been edited after it was translated: refresh the
      // translation so the saved pair always matches (skipped if it fails).
      const noteNow = state.note.trim();
      if (state.voice && aiTranslateEnabled() && noteNow && noteNow !== state.translatedFrom) {
        const tr = await tryTranslate(noteNow);
        state.translation = tr ? tr.translation : "";
        state.translationLang = tr ? tr.target : "";
        state.translatedFrom = noteNow;
      } else if (!noteNow || !state.voice || noteNow !== state.translatedFrom) {
        state.translation = "";
      }

      const projectId = state.projectId;
      const now = Date.now();
      const inspectionId = MossDB.uid();
      const stamp = state.voice ? issueStamp(state.voice.startedAt) : issueStamp(new Date());
      const title = state.title.trim() || `${state.titleLang === "es" ? "Inspección" : "Inspection"} — ${stamp}`;
      const photoCaps = state.photos.map((p, i) => {
        const ext = { "image/png": "png", "image/heic": "heic", "image/heif": "heif", "image/webp": "webp" }[p.file.type] || "jpg";
        return { id: MossDB.uid(), file: p.file, dataUrl: p.dataUrl, remoteFileName: `inspection-photo-${now}-${i + 1}.${ext}` };
      });
      const voiceId = state.voice ? MossDB.uid() : null;
      const voiceName = state.voice ? `inspection-voice-${now}.${state.voice.mimeType.includes("mp4") ? "m4a" : "webm"}` : "";

      for (const p of photoCaps) {
        await MossDB.captures.add({ id: p.id, projectId, type: "photo", dataUrl: p.dataUrl, name: p.file.name, remoteFileName: p.remoteFileName, caption: title, inspectionId });
      }
      if (state.voice) {
        await MossDB.captures.add({
          id: voiceId,
          projectId,
          type: "voice",
          dataUrl: state.voice.dataUrl,
          transcript: noteNow || state.voice.transcript || null,
          remoteFileName: voiceName,
          inspectionId,
          ...(state.translation ? { translation: state.translation, translationLang: state.translationLang } : {})
        });
      }
      const saved = await MossDB.captures.add({
        id: inspectionId,
        projectId,
        type: "inspection",
        name: title,
        result: state.result,
        comments: noteNow,
        inspector: state.inspector.trim(),
        recordedAt: (state.voice ? state.voice.startedAt : new Date()).toISOString(),
        photoIds: photoCaps.map((p) => p.id),
        voiceId,
        ...(state.translation ? { translation: state.translation, translationLang: state.translationLang } : {})
      });

      const voiceBlob = state.voice ? state.voice.blob : null;
      closeSheet();
      toast("Inspection saved");
      if (currentRoute().name === "project" || currentRoute().name === "home") render();

      // Best-effort OneDrive copy of the files; announce them to other
      // devices only once everything has actually finished uploading.
      const uploads = photoCaps.map((p) => syncCaptureToOneDrive(projectId, "photo", p.file, p.remoteFileName));
      if (voiceBlob) uploads.push(syncCaptureToOneDrive(projectId, "voice", voiceBlob, voiceName));
      if (uploads.length) Promise.all(uploads).then(() => syncCapturesWithOneDrive());

      if (thenPdf) {
        const project = await MossDB.projects.get(projectId);
        const caps = await MossDB.captures.forProject(projectId);
        makeInspectionsPdf([{ project, inspections: [saved], captures: caps }]);
      }
    } catch (err) {
      console.error("Saving inspection failed", err);
      state.saving = false;
      toast("Couldn't save the inspection — try again");
      draw();
    }
  };

  openSheet("", () => {
    // Runs whenever the sheet goes away (Cancel, backdrop tap, navigation,
    // or after a save): always release the microphone.
    closed = true;
    analysisRun++;
    translateRun++;
    titleRun++;
    if (state.recorder) {
      state.recorder.cancel();
      state.recorder = null;
    }
  });
  draw();
}

// What the PDF builder (and the detail screen) need from a saved inspection.
function inspectionPdfData(insp, captures) {
  const photos = (insp.photoIds || []).map((id) => captures.find((c) => c.id === id)?.dataUrl).filter(Boolean);
  return {
    title: insp.name || "Inspection",
    result: insp.result || "",
    stamp: insp.recordedAt ? issueStamp(new Date(insp.recordedAt)) : insp.createdAt ? issueStamp(new Date(insp.createdAt)) : "",
    inspector: insp.inspector || "",
    note: insp.comments || "",
    translation: insp.translation || "",
    translationLabel: insp.translation ? translationLabel(insp.translationLang) : "",
    photos
  };
}

// groups = [{ project, inspections: [record...], captures: [...] }] -> PDF -> share sheet.
async function makeInspectionsPdf(groups) {
  if (!window.jspdf) {
    toast("PDF library missing — upload jspdf.umd.min.js to GitHub");
    return;
  }
  try {
    const sections = groups
      .filter((g) => g.inspections.length)
      .map((g) => ({
        projectName: g.project.name,
        address: g.project.address || "",
        inspections: g.inspections.map((i) => inspectionPdfData(i, g.captures))
      }));
    const blob = await buildInspectionsPdf(sections);
    const day = new Date().toISOString().slice(0, 10);
    const count = sections.reduce((n, s) => n + s.inspections.length, 0);
    const label = sections.length === 1 ? sections[0].projectName : "All projects";
    const what = count === 1 ? `Inspection - ${sections[0].inspections[0].title}` : "Inspections";
    const fileName = `${what} - ${label} - ${day}.pdf`.replace(/[\\/:*?"<>|]+/g, "").replace(/\s+/g, " ").slice(0, 120).replace(/\s+\.pdf$/, ".pdf");
    showPdfReady(blob, fileName.endsWith(".pdf") ? fileName : fileName + ".pdf", count, "inspection");
  } catch (err) {
    console.error("Inspection PDF build failed", err);
    toast("Couldn't build the PDF — try again");
  }
}

// Picker for the project's inspections (all start checked), then the PDF.
async function openInspectionsPdfSheet(projectId) {
  if (!window.jspdf) {
    toast("PDF library missing — upload jspdf.umd.min.js to GitHub");
    return;
  }
  const project = await MossDB.projects.get(projectId);
  const captures = await MossDB.captures.forProject(projectId);
  const inspections = captures
    .filter((c) => c.type === "inspection")
    .sort((a, b) => String(a.recordedAt || a.createdAt).localeCompare(String(b.recordedAt || b.createdAt)));
  if (!project || !inspections.length) {
    toast("No inspections to put in a PDF yet");
    return;
  }
  openSheet(`
    <h2>Inspection report PDF</h2>
    <p class="empty" style="text-align:left; margin-top:-4px;">Choose what goes in the report. Each inspection includes its result, photos and written note.</p>
    <div class="card card-list">
      ${inspections
        .map(
          (i) => `
        <label class="row" style="cursor:pointer; gap:12px;">
          <input type="checkbox" data-iid="${escapeHtml(i.id)}" checked style="width:20px; height:20px; flex:none;">
          <span class="main"><span class="title">${escapeHtml(i.name || "Inspection")}</span><span class="desc">${escapeHtml(inspectionResultInfo(i.result).short)}${i.recordedAt || i.createdAt ? " · " + escapeHtml(new Date(i.recordedAt || i.createdAt).toLocaleDateString()) : ""}</span></span>
        </label>`
        )
        .join("")}
    </div>
    <div class="sheet-actions">
      <button class="btn ghost" id="pdf-cancel">Cancel</button>
      <button class="btn primary" id="pdf-create">Create PDF</button>
    </div>
  `);
  document.getElementById("pdf-cancel").addEventListener("click", closeSheet);
  document.getElementById("pdf-create").addEventListener("click", async () => {
    const chosen = new Set([...$sheet.querySelectorAll("[data-iid]")].filter((el) => el.checked).map((el) => el.dataset.iid));
    if (!chosen.size) {
      toast("Select at least one inspection");
      return;
    }
    const btn = document.getElementById("pdf-create");
    btn.disabled = true;
    btn.textContent = "Building PDF…";
    await makeInspectionsPdf([{ project, inspections: inspections.filter((i) => chosen.has(i.id)), captures }]);
  });
}

// A saved inspection: result, date, inspector, photos, note, translation,
// audio — plus Translate again, PDF and Delete.
function openInspectionDetail(insp, captures) {
  const data = inspectionPdfData(insp, captures);
  const voice = captures.find((c) => c.id === insp.voiceId);
  const info = inspectionResultInfo(insp.result);
  openSheet(`
    <h2>${escapeHtml(data.title)}</h2>
    <span class="desc">${[`${info.icon} ${info.short}`, data.stamp, data.inspector ? "Inspector: " + data.inspector : ""].filter(Boolean).map(escapeHtml).join(" · ")}</span>
    ${data.photos.length ? `<div style="display:grid; grid-template-columns:1fr 1fr; gap:8px;">${data.photos.map((p) => `<img src="${p}" alt="Inspection photo" style="width:100%; border-radius:10px; display:block;">`).join("")}</div>` : ""}
    ${data.note ? `<div class="card" style="padding:12px 14px; font-size:14px; line-height:1.5; white-space:pre-wrap;">${escapeHtml(data.note)}</div>` : ""}
    ${translationCardHtml(insp.translation, insp.translationLang)}
    ${retranslateButtonHtml("d-retr", !!insp.translation, data.note)}
    ${voice?.dataUrl ? `<audio controls preload="metadata" style="width:100%; height:36px;" src="${voice.dataUrl}"></audio>` : ""}
    <button class="btn primary" id="d-pdf" style="width:100%;">📄 Create PDF</button>
    <button class="btn ghost" id="d-delete" style="width:100%; color:var(--red, #DC2626);">🗑 Delete inspection</button>
    ${closeButton()}
  `);
  wireCloseButton();
  wireRetranslate(
    "d-retr",
    data.note,
    insp.translation || "",
    async (patch) => {
      Object.assign(insp, patch); // the dashboard's own copy too
      await MossDB.captures.update(insp.id, patch);
      if (voice) {
        Object.assign(voice, patch);
        await MossDB.captures.update(voice.id, patch);
      }
    },
    () => openInspectionDetail(insp, captures)
  );
  document.getElementById("d-pdf").addEventListener("click", async () => {
    const project = await MossDB.projects.get(insp.projectId);
    makeInspectionsPdf([{ project: project || { name: "Project" }, inspections: [insp], captures }]);
  });
  document.getElementById("d-delete").addEventListener("click", () => confirmDeleteInspection(insp));
}

async function deleteInspectionCompletely(insp) {
  const caps = await MossDB.captures.forProject(insp.projectId);
  const ids = new Set([...(insp.photoIds || []), insp.voiceId].filter(Boolean));
  for (const c of caps) if (c.inspectionId === insp.id) ids.add(c.id);
  for (const id of ids) await MossDB.captures.remove(id);
  await MossDB.captures.remove(insp.id);
}

function confirmDeleteInspection(insp) {
  openSheet(`
    <h2>Delete "${escapeHtml(insp.name || "Inspection")}"?</h2>
    <p class="empty" style="text-align:left; margin-top:-4px;">
      This permanently removes the inspection, with its photos and voice note, from the app. It can't be undone.
      Copies already saved in OneDrive are not deleted.
    </p>
    <div class="sheet-actions">
      <button class="btn ghost" id="di-cancel">Cancel</button>
      <button class="btn primary" id="di-confirm" style="background:var(--red, #DC2626);">Delete</button>
    </div>
  `);
  document.getElementById("di-cancel").addEventListener("click", closeSheet);
  document.getElementById("di-confirm").addEventListener("click", async () => {
    const btn = document.getElementById("di-confirm");
    btn.disabled = true;
    try {
      await deleteInspectionCompletely(insp);
    } catch (err) {
      console.error("Deleting inspection failed", err);
      toast("Couldn't delete — try again");
      btn.disabled = false;
      return;
    }
    closeSheet();
    toast("Inspection deleted");
    syncCapturesWithOneDrive();
    render();
  });
}

// ---------- Reports (local, pre-AI versions) ----------

function closeButton() {
  return `<div class="sheet-actions"><button class="btn primary" id="sheet-close" style="flex:1;">Close</button></div>`;
}
function wireCloseButton() {
  document.getElementById("sheet-close")?.addEventListener("click", closeSheet);
}

async function showNextSteps() {
  const issues = (await MossDB.issues.all()).filter((i) => i.status === "Open");

  let aiSummary = "";
  if (aiConfigured()) {
    try {
      const context = await buildAIContext();
      aiSummary = await askClaude(
        `You are a general contractor's assistant. Using ONLY the project data below, list the top ` +
        `priority next steps across all active projects — the things most urgent or most likely to ` +
        `block progress (open issues, anything time-sensitive mentioned in logs or notes). Keep it ` +
        `short: a plain-text prioritized list, no markdown headers.\n\n${context}`
      );
    } catch (err) {
      aiSummary = `⚠️ ${err.message}`;
    }
  }

  openSheet(`
    <h2>Next Steps</h2>
    ${
      aiSummary
        ? `<div class="card" style="padding:14px 16px; font-size:14px; line-height:1.6; white-space:pre-wrap;">${escapeHtml(aiSummary)}</div>`
        : ""
    }
    <div class="card card-list">
      ${
        issues.length
          ? issues.slice(0, 8).map((i) => `<div class="row"><span class="main"><span class="title">${escapeHtml(i.title)}</span><span class="desc">${escapeHtml(i.trade)}</span></span></div>`).join("")
          : `<div class="row"><span class="empty">Nothing outstanding.</span></div>`
      }
    </div>
    ${!aiConfigured() ? `<p class="empty">Add a Claude API key in Settings for an AI-prioritized summary here.</p>` : ""}
    ${closeButton()}
  `);
  wireCloseButton();
}

async function showDailyReport() {
  const projects = (await MossDB.projects.all()).filter(isVisibleProject);
  const today = new Date().toDateString();
  let lines = [];
  for (const p of projects) {
    const issues = await MossDB.issues.forProject(p.id);
    const captures = await MossDB.captures.forProject(p.id);
    const todays = captures.filter((c) => new Date(c.createdAt).toDateString() === today);
    const todaysIssues = issues.filter((i) => new Date(i.createdAt).toDateString() === today);
    if (todays.length || todaysIssues.length) {
      const summary = `${todaysIssues.length} new issue(s), ${todays.length} capture(s) today.`;
      lines.push(`<strong>${escapeHtml(p.name)}</strong>: ${summary}`);
      // One log entry per project per day — re-opening the report later
      // the same day (after more captures) updates that same entry's
      // count instead of piling up duplicate rows, which used to inflate
      // both the Daily Logs list and the AI context sent to Ask AI.
      const existingLogs = await MossDB.logs.forProject(p.id);
      const todaysLog = existingLogs.find((l) => new Date(l.createdAt).toDateString() === today);
      if (todaysLog) {
        await MossDB.logs.update(todaysLog.id, { summary });
      } else {
        await MossDB.logs.add({ projectId: p.id, summary });
      }
    }
  }

  let narrative = "";
  if (aiConfigured() && lines.length) {
    try {
      const context = await buildAIContext();
      narrative = await askClaude(
        `You are writing today's daily log narrative for a general contractor, using ONLY the project ` +
        `data below. Write 2-4 short plain-professional sentences for each project that had activity ` +
        `today (today is ${today}) — only mention projects with today's date in their captures/issues/logs.\n\n${context}`
      );
    } catch (err) {
      narrative = `⚠️ ${err.message}`;
    }
  }

  openSheet(`
    <h2>Today's Daily Log</h2>
    <div class="card" style="padding:14px 16px; font-size:14px; line-height:1.6; white-space:pre-wrap;">
      ${narrative ? escapeHtml(narrative) : (lines.length ? lines.join("<br>") : "Nothing captured yet today.")}
    </div>
    <p class="empty">${
      !aiConfigured()
        ? "Add a Claude API key in Settings for a written narrative log."
        : lines.length
        ? "Saved to each project's Daily Logs."
        : ""
    }</p>
    ${closeButton()}
  `);
  wireCloseButton();
  if (currentRoute().name === "project" && lines.length) render();
}

async function showWeeklyReport() {
  let narrative = "";
  if (aiConfigured()) {
    try {
      const context = await buildAIContext();
      narrative = await askClaude(
        `You are writing a weekly report for a general contractor, using ONLY the project data below ` +
        `(issues, logs, and captures on file). For each active project, summarize: completed/recent ` +
        `work, open issues or inspections, and a brief note on what's likely next. Be concise and ` +
        `professional, plain text, no markdown headers.\n\n${context}`
      );
    } catch (err) {
      narrative = `⚠️ ${err.message}`;
    }
  }

  openSheet(`
    <h2>Weekly Report</h2>
    ${
      narrative
        ? `<div class="card" style="padding:14px 16px; font-size:14px; line-height:1.6; white-space:pre-wrap;">${escapeHtml(narrative)}</div>`
        : `<p class="empty">Add a Claude API key in Settings to generate weekly reports (completed work, inspections, change orders, next week's plan).</p>`
    }
    ${closeButton()}
  `);
  wireCloseButton();
}

// ---------- helpers ----------

// Photos are stored at up to 2000px (about 0.3-1 MB) instead of the phone's
// full 5-12 MB original: five full-size photos per issue filled up the
// phone's storage, and then saves were refused. The original file still goes
// to OneDrive untouched. Falls back to the original if shrinking fails.
async function photoFileToDataUrl(file) {
  const original = await fileToDataUrl(file);
  if (original.length < 1_500_000 || typeof prepareImageForPdf !== "function") return original;
  try {
    const small = await prepareImageForPdf(original, 2000, 0.85);
    if (small && small.dataUrl && small.dataUrl.length < original.length) return small.dataUrl;
  } catch {}
  return original;
}

// If issues went missing (a save the phone refused half-way) but their voice
// notes / photos are still stored, rebuild the issue from them. Returns how
// many were rebuilt.
async function recoverOrphanIssues() {
  try {
    const [issues, caps] = await Promise.all([MossDB.issues.all(), MossDB.captures.all()]);
    const have = new Set(issues.map((i) => i.id));
    const groups = new Map();
    for (const c of caps) {
      if (!c.issueId || have.has(c.issueId) || (c.type !== "photo" && c.type !== "voice")) continue;
      if (!c.dataUrl) continue; // text-only copies synced from another device aren't ours to rebuild
      if (Date.now() - new Date(c.createdAt || 0).getTime() < 120000) continue; // may be mid-save right now
      if (!groups.has(c.issueId)) groups.set(c.issueId, []);
      groups.get(c.issueId).push(c);
    }
    let n = 0;
    for (const [issueId, list] of groups) {
      list.sort((a, b) => String(a.createdAt).localeCompare(String(b.createdAt)));
      const photos = list.filter((c) => c.type === "photo");
      const voice = list.find((c) => c.type === "voice");
      const when = new Date((voice || list[0]).createdAt || Date.now());
      const stamp = issueStamp(when);
      const note = ((voice && voice.transcript) || "").trim();
      const caption = ((photos.find((p) => p.caption) || {}).caption || "").trim();
      const title = (caption || note).slice(0, 70) || "Recovered issue";
      await MossDB.issues.add({
        id: issueId,
        projectId: list[0].projectId,
        title,
        trade: "General",
        requirement: note ? `${stamp}\n${note}` : stamp,
        photoId: photos[0] ? photos[0].id : null,
        voiceId: voice ? voice.id : null,
        ...(photos.length > 1 ? { extraPhotoIds: photos.slice(1).map((p) => p.id) } : {}),
        recordedAt: when.toISOString(),
        recovered: true,
        ...(voice && voice.translation ? { translation: voice.translation, translationLang: voice.translationLang } : {})
      });
      n++;
    }
    return n;
  } catch (err) {
    console.error("Recovering issues failed", err);
    return 0;
  }
}

function fileToDataUrl(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = () => resolve(reader.result);
    reader.onerror = reject;
    reader.readAsDataURL(file);
  });
}

function blobToDataUrl(blob) {
  return fileToDataUrl(blob);
}

// ---------- OneDrive sync (Phase 2) ----------
// Best-effort: MossDB/IndexedDB is always the source of truth for the app
// itself, so a sync failure here should never block or undo a capture —
// only skip the OneDrive copy and let the user know via toast.

function syncSubfolderFor(type) {
  if (type === "photo") return "02 PHOTOS/02 PROGRESS";
  if (type === "voice") return "04 DAILY LOGS";
  return "01 PLANS & DRAWINGS/CURRENT"; // document
}

async function syncCaptureToOneDrive(projectId, type, blob, fileName) {
  if (!msalConfigured() || !msCurrentAccount()) return; // local-only until connected
  try {
    const project = await MossDB.projects.get(projectId);
    if (!project) return;
    const token = await msGetToken();
    if (!token) return; // msGetToken() already kicked off a re-auth redirect if needed
    await uploadToOneDrive(project.name, syncSubfolderFor(type), fileName, blob, token);
    toast("Synced to OneDrive");
  } catch (err) {
    console.error("OneDrive sync failed", err);
    toast("Saved locally (OneDrive sync failed)");
  }
}
