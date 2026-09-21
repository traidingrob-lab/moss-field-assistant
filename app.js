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
    await render();
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

    const [allLocal, remote] = await Promise.all([MossDB.captures.all(), fetchCapturesIndex(token)]);
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
  const openIssues = issues.filter((i) => i.status === "Open");
  const photos = captures.filter((c) => c.type === "photo");

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
      <div class="stat-tile"><span class="num">${captures.filter((c) => c.type === "material").length}</span><span class="label">MATERIALS</span></div>
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
      <div class="section-label">Open Issues<span class="link" data-quick="issue">+ Add</span></div>
      <div class="card card-list" style="margin-top:10px;">
        ${
          openIssues.length
            ? openIssues
                .map(
                  (i) => `
          <div class="row" style="cursor:default;">
            <span class="icon" style="color:var(--red);">●</span>
            <span class="main"><span class="title">${escapeHtml(i.title)}</span><span class="desc">${escapeHtml(i.trade)}${i.requirement ? " · " + escapeHtml(i.requirement) : ""}</span></span>
          </div>`
                )
                .join("")
            : `<div class="row"><span class="empty">No open issues yet — capture one from the field.</span></div>`
        }
      </div>
    </div>

    <div>
      <div class="section-label">Recent Captures</div>
      <div class="thumb-grid" style="margin-top:10px;">
        ${
          photos.length
            ? photos
                .slice(-4)
                .reverse()
                .map((p) => {
                  // A capture synced in from another device as text-only
                  // (see syncCapturesWithOneDrive) has no dataUrl here —
                  // there's no image to show, only whatever caption came
                  // with it, so render that instead of a broken <img>.
                  if (!p.dataUrl) {
                    return `
          <div class="thumb" style="display:flex; align-items:center; justify-content:center; text-align:center; padding:8px; background:var(--bg-2, #F3F4F6);">
            <span class="desc" style="font-size:11px; line-height:1.3;">${
              p.caption ? escapeHtml(p.caption) : "Photo from another device"
            }</span>
          </div>`;
                  }
                  return `
          <div class="thumb" title="${p.caption ? escapeHtml(p.caption) : ""}">
            <img src="${p.dataUrl}" alt="${p.caption ? escapeHtml(p.caption) : ""}">
            ${
              p.caption
                ? `<span class="desc" style="display:block; font-size:11px; margin-top:4px; line-height:1.3;">${escapeHtml(p.caption)}</span>`
                : aiConfigured()
                ? `<span class="desc" style="display:block; font-size:11px; margin-top:4px; opacity:.6;">Describing…</span>`
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
      <div class="section-label">Voice Notes</div>
      <div class="card card-list" style="margin-top:10px;">
        ${
          captures.filter((c) => c.type === "voice").length
            ? captures
                .filter((c) => c.type === "voice")
                .slice(-5)
                .reverse()
                .map(
                  (v) => `
          <div class="row" style="cursor:default; flex-direction:column; align-items:stretch; gap:8px;">
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

  wireQuickCapture();
  $app.querySelector('[data-quick="issue"]').addEventListener("click", () => openIssueSheet(id));
  $app.querySelector("#toggle-project-status").addEventListener("click", async () => {
    await MossDB.projects.update(id, { status: completed ? "Active" : "Completed" });
    syncProjectsWithOneDrive();
    toast(completed ? "Project reactivated" : "Project marked complete");
    renderDashboard(id);
  });
  $app.querySelector("#delete-project").addEventListener("click", () => confirmDeleteProject(project));

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
      <div class="project-head"><div class="title" style="font-size:20px;">Open Issues</div></div>
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
  return voiceLangStored() || navigator.language || "en-US";
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
      const projectId = await pickProjectIfNeeded(currentProjectId());
      if (!projectId) return;
      if (type === "photo") capturePhoto(projectId);
      else if (type === "voice") captureVoice(projectId);
      else if (type === "document") captureDocument(projectId);
      else if (type === "issue") openIssueSheet(projectId);
      else if (type === "material") openMaterialSheet(projectId);
      else if (type === "inspection") openInspectionSheet(projectId);
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
    const dataUrl = await fileToDataUrl(file);
    // Recorded once and reused for both the OneDrive upload and the local
    // record, so another device can later ask OneDrive for this exact file
    // by name (see hydrateRemoteCapture) — relying on file.name alone
    // wasn't reliable enough to build a re-download path from.
    const remoteFileName = file.name || `photo-${Date.now()}.jpg`;
    const capture = await MossDB.captures.add({ projectId, type: "photo", dataUrl, name: file.name, remoteFileName });
    toast("Photo saved");
    if (currentRoute().name === "project") render();
    // Kept so the captions chain below can wait for the real file to
    // actually finish reaching OneDrive before telling OTHER devices this
    // remoteFileName exists — pushing the captures index first would let
    // another device try to hydrate a file that isn't there yet (and
    // hydrateRemoteCapture doesn't retry a failed fetch until next visit).
    const uploadPromise = syncCaptureToOneDrive(projectId, "photo", file, remoteFileName);

    // Caption it in the background so Ask AI and the reports can describe
    // what's in the photo later. Best-effort: no API key yet, or the call
    // fails, and the photo is still saved fine — it just won't be
    // describable by AI until a key is added or the next capture works.
    if (aiConfigured()) {
      captionPhoto(dataUrl)
        .then((caption) => caption && MossDB.captures.update(capture.id, { caption }))
        .then(() => { if (currentRoute().name === "project") render(); })
        .then(() => uploadPromise)
        .then(() => syncCapturesWithOneDrive())
        .catch((err) => console.warn("Photo captioning skipped:", err.message));
    }
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
        await MossDB.captures.add({ projectId, type: "voice", dataUrl, transcript: finalTranscript, remoteFileName });
        toast(finalTranscript ? "Voice note saved (transcribed)" : "Voice note saved");
        if (currentRoute().name === "project") render();
        // Same ordering reason as capturePhoto: don't advertise this
        // remoteFileName to other devices until the real file has actually
        // finished uploading.
        const uploadPromise = syncCaptureToOneDrive(projectId, "voice", blob, remoteFileName);
        if (finalTranscript) uploadPromise.then(() => syncCapturesWithOneDrive());
      });
    };

    document.getElementById("stop-rec").addEventListener("click", () => stop(true));
    document.getElementById("cancel-rec").addEventListener("click", () => stop(false));
  } catch (err) {
    toast("Microphone permission denied");
  }
}

function openIssueSheet(projectId) {
  openSheet(`
    <h2>New Issue</h2>
    <div class="field">
      <label>Title</label>
      <input id="f-title" placeholder='e.g. "Hall bathroom floor register"'>
    </div>
    <div class="field">
      <label>Trade</label>
      <select id="f-trade">${TRADES.map((t) => `<option>${t}</option>`).join("")}</select>
    </div>
    <div class="field">
      <label>Requirement / notes</label>
      <textarea id="f-req" placeholder="What does it need to pass?"></textarea>
    </div>
    <div class="sheet-actions">
      <button class="btn ghost" id="cancel">Cancel</button>
      <button class="btn primary" id="save">Save Issue</button>
    </div>
  `);
  document.getElementById("cancel").addEventListener("click", closeSheet);
  document.getElementById("save").addEventListener("click", async () => {
    const title = document.getElementById("f-title").value.trim();
    if (!title) {
      toast("Give the issue a title");
      return;
    }
    await MossDB.issues.add({
      projectId,
      title,
      trade: document.getElementById("f-trade").value,
      requirement: document.getElementById("f-req").value.trim()
    });
    closeSheet();
    toast("Issue created");
    if (currentRoute().name === "project" || currentRoute().name === "home") render();
  });
}

function openMaterialSheet(projectId) {
  openSheet(`
    <h2>Material</h2>
    <div class="field"><label>Item</label><input id="f-item" placeholder="e.g. Hard pipe duct, 6&quot;"></div>
    <div class="field"><label>Status</label>
      <select id="f-status"><option>Ordered</option><option>Delivered</option><option>Backordered</option></select>
    </div>
    <div class="sheet-actions">
      <button class="btn ghost" id="cancel">Cancel</button>
      <button class="btn primary" id="save">Save</button>
    </div>
  `);
  document.getElementById("cancel").addEventListener("click", closeSheet);
  document.getElementById("save").addEventListener("click", async () => {
    const item = document.getElementById("f-item").value.trim();
    if (!item) { toast("Add an item name"); return; }
    await MossDB.captures.add({ projectId, type: "material", name: item, status: document.getElementById("f-status").value });
    closeSheet();
    toast("Material logged");
    if (currentRoute().name === "project") render();
  });
}

function openInspectionSheet(projectId) {
  openSheet(`
    <h2>Inspection</h2>
    <div class="field"><label>Type</label><input id="f-type" placeholder="e.g. Mechanical rough-in"></div>
    <div class="field"><label>Result</label>
      <select id="f-result"><option>Scheduled</option><option>Passed</option><option>Failed / corrections required</option></select>
    </div>
    <div class="field"><label>Inspector comments</label><textarea id="f-comments"></textarea></div>
    <div class="sheet-actions">
      <button class="btn ghost" id="cancel">Cancel</button>
      <button class="btn primary" id="save">Save</button>
    </div>
  `);
  document.getElementById("cancel").addEventListener("click", closeSheet);
  document.getElementById("save").addEventListener("click", async () => {
    const type = document.getElementById("f-type").value.trim();
    if (!type) { toast("Add an inspection type"); return; }
    await MossDB.captures.add({
      projectId,
      type: "inspection",
      name: type,
      result: document.getElementById("f-result").value,
      comments: document.getElementById("f-comments").value.trim()
    });
    closeSheet();
    toast("Inspection logged");
    if (currentRoute().name === "project") render();
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
