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

async function renderHome() {
  const allProjects = await MossDB.projects.all();
  const projects = allProjects.filter(isActiveProject);
  const allIssues = await MossDB.issues.all();
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
  const allProjects = await MossDB.projects.all();
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
  if (!project) {
    location.hash = "#/";
    return;
  }
  const issues = await MossDB.issues.forProject(id);
  const captures = await MossDB.captures.forProject(id);
  const openIssues = issues.filter((i) => i.status === "Open");
  const photos = captures.filter((c) => c.type === "photo");
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
                .map((p) => `<div class="thumb"><img src="${p.dataUrl}" alt=""></div>`)
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
            <audio controls preload="none" style="width:100%; height:32px;" src="${v.dataUrl}"></audio>
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
  const issues = (await MossDB.issues.all()).filter((i) => i.status === "Open");
  const projects = await MossDB.projects.all();
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

// ---------- Ask AI (placeholder for Phase 3) ----------

async function renderAsk() {
  const header = `
    <div class="topbar">
      <div class="project-head"><div class="title" style="font-size:20px;">Ask AI</div></div>
    </div>
  `;
  const body = `
    <div class="card" style="padding:16px;">
      <div class="field">
        <label>Ask about any project</label>
        <textarea id="ask-input" placeholder='e.g. "What is still open at Helm?"'></textarea>
      </div>
      <button class="btn primary" id="ask-submit" style="margin-top:12px; width:100%;">Ask</button>
      <p class="empty" style="margin-top:12px;">The AI layer connects in Phase 3 — this will search everything captured here and in OneDrive once it's wired up. For now, try "Open Issues" from the home screen to see live local data.</p>
    </div>
  `;
  shell({ header, body, activeTab: "ask" });
  document.getElementById("ask-submit").addEventListener("click", () => {
    toast("AI search comes online in Phase 3");
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
      <div class="row"><span class="icon">🤖</span><span class="main"><span class="title">AI processing</span><span class="desc">Not connected yet — Phase 3</span></span></div>
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
}

// ---------- Quick capture ----------

function currentProjectId() {
  const route = currentRoute();
  return route.name === "project" ? route.id : null;
}

async function pickProjectIfNeeded(preselected) {
  if (preselected) return preselected;
  const projects = await MossDB.projects.all();
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
    await MossDB.captures.add({ projectId, type: "photo", dataUrl, name: file.name });
    toast("Photo saved");
    if (currentRoute().name === "project") render();
    syncCaptureToOneDrive(projectId, "photo", file, file.name || `photo-${Date.now()}.jpg`);
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
      closeSheet();
      if (!save) return;
      recorder.addEventListener("stop", async () => {
        // Use the recorder's own mimeType (Safari records audio/mp4, not
        // webm) — hardcoding "audio/webm" produced a blob mislabeled on
        // iPhone, which is one of the target platforms.
        const blob = new Blob(chunks, { type: recorder.mimeType || "audio/webm" });
        const dataUrl = await blobToDataUrl(blob);
        await MossDB.captures.add({ projectId, type: "voice", dataUrl });
        toast("Voice note saved");
        if (currentRoute().name === "project") render();
        const ext = (recorder.mimeType || "audio/webm").includes("mp4") ? "m4a" : "webm";
        syncCaptureToOneDrive(projectId, "voice", blob, `voice-${Date.now()}.${ext}`);
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
  openSheet(`
    <h2>Next Steps</h2>
    <div class="card card-list">
      ${
        issues.length
          ? issues.slice(0, 8).map((i) => `<div class="row"><span class="main"><span class="title">${escapeHtml(i.title)}</span><span class="desc">${escapeHtml(i.trade)}</span></span></div>`).join("")
          : `<div class="row"><span class="empty">Nothing outstanding.</span></div>`
      }
    </div>
    <p class="empty">This gets smarter once AI processing (Phase 3) can prioritize by inspection dates and blockers.</p>
    ${closeButton()}
  `);
  wireCloseButton();
}

async function showDailyReport() {
  const projects = await MossDB.projects.all();
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
      // Persist so the project's own Daily Logs section (and future
      // reports) actually have something to show — previously nothing
      // ever wrote to the logs store, so that section stayed empty forever.
      await MossDB.logs.add({ projectId: p.id, summary });
    }
  }
  openSheet(`
    <h2>Today's Daily Log</h2>
    <div class="card" style="padding:14px 16px; font-size:14px; line-height:1.6;">
      ${lines.length ? lines.join("<br>") : "Nothing captured yet today."}
    </div>
    <p class="empty">A written narrative daily log comes with Phase 3's AI processing. ${lines.length ? "Saved to each project's Daily Logs." : ""}</p>
    ${closeButton()}
  `);
  wireCloseButton();
  if (currentRoute().name === "project" && lines.length) render();
}

async function showWeeklyReport() {
  openSheet(`
    <h2>Weekly Report</h2>
    <p class="empty">Weekly summaries (completed work, inspections, change orders, next week's plan) generate once AI processing is connected — Phase 3.</p>
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
