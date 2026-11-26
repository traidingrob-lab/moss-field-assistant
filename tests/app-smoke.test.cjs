// Prueba de humo: la app arranca en un navegador simulado (jsdom + IndexedDB
// falso), cada pantalla principal se dibuja sin errores, y Settings muestra
// la versión y el interruptor "Economy AI" (v54).
// Correr: cd tests && npm install && node app-smoke.test.cjs
"use strict";

const fs = require("fs");
const path = require("path");
const assert = require("assert");
const { JSDOM } = require("jsdom");

const ROOT = path.join(__dirname, "..");
const SCRIPTS = [
  "config.js",
  "auth.js",
  "graph.js",
  "ai.js",
  "jspdf.umd.min.js",
  "report-pdf.js",
  "db.js",
  "app.js",
  "icons.js"
];

const sleep = (ms) => new Promise((r) => setTimeout(r, ms));

async function boot() {
  const html = fs.readFileSync(path.join(ROOT, "index.html"), "utf8").replace(/<script src="[^"]*"><\/script>/g, "");
  const dom = new JSDOM(html, {
    url: "https://traidingrob-lab.github.io/moss-field-assistant/",
    // Los archivos se inyectan como <script> (como en la app real) para que
    // compartan el mismo ámbito global; jsdom no aplica la CSP.
    runScripts: "dangerously",
    pretendToBeVisual: true
  });
  const w = dom.window;
  const errors = [];
  w.addEventListener("error", (e) => errors.push(e.error || e.message));
  w.addEventListener("unhandledrejection", (e) => errors.push(e.reason));
  w.console.error = (...a) => errors.push(a.map(String).join(" "));

  // IndexedDB falso y APIs que jsdom no trae.
  require("fake-indexeddb/auto");
  w.indexedDB = globalThis.indexedDB;
  w.IDBKeyRange = globalThis.IDBKeyRange;
  w.scrollTo = () => {};
  w.HTMLElement.prototype.scrollIntoView = () => {};
  w.matchMedia = () => ({ matches: false, addListener() {}, removeListener() {}, addEventListener() {}, removeEventListener() {} });
  w.navigator.storage = { persisted: async () => true, persist: async () => true, estimate: async () => ({ usage: 1, quota: 100 }) };
  w.localStorage.setItem("moss_anthropic_api_key", "sk-ant-test");
  // Sin Microsoft en la prueba: MSAL "no configurado".
  w.msal = undefined;

  for (const f of SCRIPTS) {
    let src = fs.readFileSync(path.join(ROOT, f), "utf8");
    if (f === "config.js") src = 'const MSAL_CLIENT_ID = "REPLACE_FOR_TESTS";';
    const el = w.document.createElement("script");
    el.textContent = src + "\n//# sourceURL=" + f;
    w.document.body.appendChild(el);
  }
  await sleep(300);
  return { w, errors };
}

async function go(w, hash) {
  w.location.hash = hash;
  await sleep(250);
}

(async () => {
  const { w, errors } = await boot();
  const doc = w.document;
  const results = [];
  const check = (name, fn) => {
    try {
      fn();
      results.push([true, name]);
    } catch (e) {
      results.push([false, name + " — " + e.message]);
    }
  };

  check("la pantalla de inicio se dibuja", () => assert.ok(doc.getElementById("app").innerHTML.length > 100));

  for (const [hash, name] of [["#/projects", "Projects"], ["#/issues", "All issues"], ["#/ask", "Ask AI"]]) {
    await go(w, hash);
    check(`la pantalla ${name} se dibuja`, () => assert.ok(doc.getElementById("app").innerHTML.length > 100));
  }

  const APP_VERSION = /APP_VERSION = "(\d+)"/.exec(fs.readFileSync(path.join(ROOT, "app.js"), "utf8"))[1];
  await go(w, "#/settings");
  check(`Settings muestra la versión ${APP_VERSION}`, () =>
    assert.match(doc.getElementById("app").textContent, new RegExp("App version\\s*" + APP_VERSION))
  );
  check("Settings muestra cuánto ocupa Moss y la lista de problemas", () => {
    assert.match(doc.getElementById("storage-moss").textContent, /Moss data: \d+ photos/);
    assert.match(doc.getElementById("error-log").textContent, /No problems recorded/);
  });
  const sw = doc.getElementById("f-ai-economy");
  check("Settings tiene el interruptor Economy AI, encendido por defecto", () => {
    assert.ok(sw, "no existe #f-ai-economy");
    assert.strictEqual(sw.checked, true);
  });

  if (sw) {
    sw.checked = false;
    sw.dispatchEvent(new w.Event("change"));
    await sleep(250);
    check("apagar Economy AI se guarda y se ve en pantalla", () => {
      assert.strictEqual(w.localStorage.getItem("moss_ai_economy"), "off");
      assert.match(doc.getElementById("app").textContent, /everything uses the stronger model/);
      assert.strictEqual(doc.getElementById("f-ai-economy").checked, false);
    });
  }

  check("la versión coincide en app.js y sw.js", () => {
    const sw = fs.readFileSync(path.join(ROOT, "sw.js"), "utf8");
    const app = fs.readFileSync(path.join(ROOT, "app.js"), "utf8");
    const a = /APP_VERSION = "(\d+)"/.exec(app)[1];
    const c = /moss-cache-v(\d+)/.exec(sw)[1];
    assert.strictEqual(a, c);
  });

  // ---- Tomar una foto (v55): el selector queda dentro de la página y la foto se guarda ----
  const toastText = () => doc.getElementById("toast").textContent;
  const photoFile = () => new w.File([new Uint8Array(2048).fill(7)], "IMG_0001.JPG", { type: "image/jpeg" });
  w.fetch = async () => ({
    ok: true,
    status: 200,
    json: async () => ({ content: [{ type: "text", text: "Drywall hung in kitchen." }], stop_reason: "end_turn" }),
    text: async () => "",
    clone() { return this; }
  });

  const pickAndSend = async (file) => {
    w.capturePhoto("proj-test");
    const input = doc.querySelector("input[data-moss-picker]");
    if (!input) return null;
    Object.defineProperty(input, "files", { value: file ? [file] : [] });
    input.dispatchEvent(new w.Event("change"));
    return input;
  };

  const picker = await pickAndSend(photoFile());
  check("el selector de cámara queda dentro de la página (arreglo iPhone)", () => assert.ok(picker, "no se encontró el input en la página"));
  await sleep(600);
  check("después de elegir la foto, el selector se quita de la página", () =>
    assert.strictEqual(doc.querySelector("input[data-moss-picker]"), null)
  );
  check('al tomar una foto aparece "Photo saved"', () => assert.match(toastText(), /Photo saved/));
  const saved = (await w.eval("MossDB").captures.all()).filter((c) => c.projectId === "proj-test" && c.type === "photo");
  check("la foto queda guardada en la base de datos", () => assert.strictEqual(saved.length, 1));
  await sleep(4800); // la descripción espera hasta 4 s a que la imagen se reduzca
  const savedAfter = (await w.eval("MossDB").captures.all()).find((c) => c.projectId === "proj-test");
  check("la descripción de la IA se guarda en la foto", () => assert.strictEqual(savedAfter && savedAfter.caption, "Drywall hung in kitchen."));

  check("sin errores de JavaScript al usar la app", () => {
    const real = errors.filter((e) => !/MSAL|msal/i.test(String(e && (e.stack || e))));
    assert.deepStrictEqual(real.map((e) => String(e && (e.stack || e)).split("\n")[0]), []);
  });

  // ---- Fallas: ahora se ven en pantalla y quedan en la lista ----
  w.fetch = async () => ({ ok: false, status: 529, text: async () => '{"error":{"message":"Overloaded"}}', clone() { return this; } });
  await pickAndSend(photoFile());
  await sleep(5200);
  check("si la IA falla, se avisa en pantalla y queda en la lista", () => {
    assert.match(toastText(), /AI description failed: .*529/);
    assert.ok(w.errorLogEntries().some((e) => e.where === "AI photo description"));
  });

  const realAdd = w.eval("MossDB").captures.add;
  w.eval("MossDB").captures.add = async () => {
    const e = new Error("The quota has been exceeded.");
    e.name = "QuotaExceededError";
    throw e;
  };
  await pickAndSend(photoFile());
  await sleep(600);
  w.eval("MossDB").captures.add = realAdd;
  check("si el teléfono no tiene espacio, lo dice claramente y queda en la lista", () => {
    assert.match(toastText(), /out of space/);
    assert.ok(w.errorLogEntries().some((e) => /^photo save/.test(e.where) && /QuotaExceededError/.test(e.msg)));
  });

  await pickAndSend(null);
  await sleep(200);
  check("si la cámara no devuelve foto, se avisa", () => assert.match(toastText(), /No photo came back/));

  await go(w, "#/projects");
  await go(w, "#/settings");
  check("Settings ▸ Recent problems muestra los problemas", () => {
    const t = doc.getElementById("error-log").textContent;
    assert.match(t, /QuotaExceededError/);
    assert.match(t, /AI photo description/);
  });

  // ---- v56: botón Test AI y traducción que ya no falla en silencio ----
  const sent = [];
  w.fetch = async (url, init) => {
    const body = JSON.parse(init.body);
    sent.push(body.model);
    return {
      ok: true,
      status: 200,
      json: async () => ({ content: [{ type: "thinking", thinking: "" }, { type: "text", text: "OK" }], stop_reason: "end_turn" }),
      text: async () => "",
      clone() { return this; }
    };
  };
  doc.getElementById("ai-test").dispatchEvent(new w.Event("click"));
  await sleep(400);
  check("Test AI muestra el resultado de Haiku y de Sonnet", () => {
    const t = doc.getElementById("ai-test-result").textContent;
    assert.match(t, /✅ claude-haiku-5-5 — [\d.]+ s — "OK"/);
    assert.match(t, /✅ claude-sonnet-5-5 — [\d.]+ s — "OK"/);
    assert.deepStrictEqual(sent, ["claude-haiku-5-5", "claude-sonnet-5-5"]);
  });

  w.fetch = async () => ({ ok: false, status: 404, text: async () => '{"error":{"message":"model not found"}}', clone() { return this; } });
  const tr = await w.tryTranslate("La tubería gotea.", true);
  check("si la traducción falla, se avisa y queda en Recent problems", () => {
    assert.strictEqual(tr, null);
    assert.match(toastText(), /Translation failed: .*404/);
    assert.ok(w.errorLogEntries().some((e) => e.where === "AI translation" && /model not found/.test(e.msg)));
  });

  const db = w.eval("MossDB");
  await db.projects.add({ id: "p-del", name: "Mafakhame" });
  // ---- v57: ✨ Make it easy to understand en una nota de voz ----
  const ORIGINAL = "eh cambiar tres tubos de media pulgada en la cocina y la fuga del baño revisar";
  const CLEAR = "- Cambiar 3 tubos de 1/2 pulgada en la cocina.\n- Revisar la fuga del baño.";
  const voice = await db.captures.add({ projectId: "p-del", type: "voice", dataUrl: "data:audio/mp4;base64,AAAA", transcript: ORIGINAL });
  const voiceRec = await db.captures.get ? await db.captures.get(voice) : (await db.captures.all()).find((c) => c.transcript === ORIGINAL);
  const v = voiceRec || (await db.captures.all()).find((c) => c.transcript === ORIGINAL);
  const asked = [];
  w.fetch = async (url, init) => {
    const body = JSON.parse(init.body);
    asked.push(body.messages[0].content);
    const isTranslate = /Translate this construction/.test(body.messages[0].content);
    const text = isTranslate ? '{"language":"es","translation":"- Replace 3 half-inch pipes in the kitchen.\\n- Check the bathroom leak."}' : CLEAR;
    return { ok: true, status: 200, json: async () => ({ content: [{ type: "text", text }], stop_reason: "end_turn" }), text: async () => "", clone() { return this; } };
  };
  w.openVoiceEditSheet(v);
  await sleep(100);
  check("Edit text tiene el botón ✨ Make it easy to understand", () => assert.ok(doc.getElementById("ve-clarify")));
  doc.getElementById("ve-clarify").dispatchEvent(new w.Event("click"));
  await sleep(300);
  check("✨ pone el texto claro en el cuadro (sin guardar todavía)", () => {
    assert.strictEqual(doc.getElementById("ve-text").value, CLEAR);
    assert.notStrictEqual(doc.getElementById("ve-undo").style.display, "none");
  });
  doc.getElementById("ve-undo").dispatchEvent(new w.Event("click"));
  check("Undo devuelve el texto original", () => assert.strictEqual(doc.getElementById("ve-text").value, ORIGINAL));
  doc.getElementById("ve-clarify").dispatchEvent(new w.Event("click"));
  await sleep(300);
  doc.getElementById("ve-save").dispatchEvent(new w.Event("click"));
  await sleep(400);
  const savedVoice = (await db.captures.all()).find((c) => c.id === v.id);
  check("al guardar: texto claro, se conserva lo que se dijo y se traduce de nuevo", () => {
    assert.strictEqual(savedVoice.transcript, CLEAR);
    assert.strictEqual(savedVoice.originalTranscript, ORIGINAL);
    assert.match(savedVoice.translation, /Replace 3 half-inch pipes/);
  });
  w.openVoiceEditSheet(savedVoice);
  await sleep(100);
  check('al volver a editar se puede ver "What was originally said"', () => assert.match(doc.getElementById("sheet").textContent, /What was originally said[\s\S]*cambiar tres tubos/));
  w.closeSheet();

  // ---- v58: botón ✨ Make clear directo en la lista de notas de voz ----
  await db.captures.add({ projectId: "p-del", type: "voice", dataUrl: "data:audio/mp4;base64,AAAA", transcript: ORIGINAL + " otra" });
  await go(w, "#/project/p-del");
  await sleep(300);
  const direct = doc.querySelector("[data-clarify-voice]");
  check("la lista de notas de voz muestra ✨ Make clear al lado de Edit text", () => assert.ok(direct, "no aparece el botón"));
  if (direct) {
    direct.dispatchEvent(new w.Event("click"));
    await sleep(400);
    check("✨ Make clear abre Edit text con el texto ya reescrito", () => {
      assert.ok(doc.getElementById("ve-text"));
      assert.strictEqual(doc.getElementById("ve-text").value, CLEAR);
    });
    w.closeSheet();
  }

  // ---- v58: ✨ en las pantallas de crear nota (New Issue / Material / Inspección) ----
  const host = doc.createElement("div");
  doc.body.appendChild(host);
  const st = { note: ORIGINAL };
  let retranslated = 0;
  const redraw = () => {
    host.innerHTML = w.clarifyControlsHtml("t", st);
    w.wireClarify("t", st, redraw, () => { retranslated++; redraw(); });
  };
  redraw();
  check("la pantalla de crear nota tiene el botón ✨ (sin Undo todavía)", () => {
    assert.ok(doc.getElementById("t-clarify"));
    assert.strictEqual(doc.getElementById("t-undo"), null);
  });
  doc.getElementById("t-clarify").dispatchEvent(new w.Event("click"));
  await sleep(300);
  check("✨ reescribe la nota, guarda la anterior y rehace la traducción", () => {
    assert.strictEqual(st.note, CLEAR);
    assert.strictEqual(st.beforeClarify, ORIGINAL);
    assert.strictEqual(retranslated, 1);
    assert.ok(doc.getElementById("t-undo"), "debe aparecer Undo");
  });
  doc.getElementById("t-undo").dispatchEvent(new w.Event("click"));
  check("Undo devuelve la nota original y la vuelve a traducir", () => {
    assert.strictEqual(st.note, ORIGINAL);
    assert.strictEqual(st.beforeClarify, null);
    assert.strictEqual(retranslated, 2);
  });
  st.note = "  ";
  redraw();
  doc.getElementById("t-clarify").dispatchEvent(new w.Event("click"));
  await sleep(50);
  check("si la nota está vacía, avisa en lugar de llamar a la IA", () => assert.match(toastText(), /no text to rewrite/));
  // Las tres pantallas reales tienen su botón y su conexión.
  check("New Issue, Material e Inspección incluyen el botón ✨", () => {
    const src = fs.readFileSync(path.join(ROOT, "app.js"), "utf8");
    for (const p of ["i", "m", "in"]) {
      assert.match(src, new RegExp(`clarifyControlsHtml\\("${p}", state\\)`));
      assert.match(src, new RegExp(`wireClarify\\("${p}", state, draw`));
    }
  });

  // ---- v60: borrar un issue abierto desde su detalle ----
  const ph = await db.captures.add({ projectId: "p-del", type: "photo", dataUrl: "data:image/jpeg;base64,AAAA", name: "leak.jpg" });
  const vo = await db.captures.add({ projectId: "p-del", type: "voice", dataUrl: "data:audio/mp4;base64,AAAA", transcript: "gotea" });
  const allCaps = await db.captures.all();
  const photoCap = allCaps.find((c) => c.name === "leak.jpg");
  const voiceCap = allCaps.find((c) => c.transcript === "gotea");
  await db.issues.add({ id: "iss-1", projectId: "p-del", title: "Leak under sink", trade: "Plumbing", photoId: photoCap.id, voiceId: voiceCap.id });
  await db.captures.update(photoCap.id, { issueId: "iss-1" });
  await db.captures.update(voiceCap.id, { issueId: "iss-1" });
  const issue1 = (await db.issues.all()).find((i) => i.id === "iss-1");
  const capsNow = await db.captures.forProject("p-del");

  w.openIssueDetail(issue1, capsNow);
  await sleep(100);
  check("el detalle de un issue abierto tiene 🗑 Delete issue", () => assert.ok(doc.getElementById("d-delete")));
  doc.getElementById("d-delete").dispatchEvent(new w.Event("click"));
  await sleep(100);
  check("pide confirmación y dice qué se borra", () =>
    assert.match(doc.getElementById("sheet").textContent, /Delete "Leak under sink"\?[\s\S]*1 photo and its voice note/)
  );
  doc.getElementById("di-cancel").dispatchEvent(new w.Event("click"));
  await sleep(100);
  check("Cancel vuelve al detalle del issue", () => {
    assert.ok(doc.getElementById("d-delete"), "debería verse otra vez el issue");
    assert.match(doc.getElementById("sheet").textContent, /Leak under sink/);
  });
  const stillThere = (await db.issues.all()).some((i) => i.id === "iss-1");
  check("Cancel no borra el issue", () => assert.strictEqual(stillThere, true));

  doc.getElementById("d-delete").dispatchEvent(new w.Event("click"));
  await sleep(100);
  doc.getElementById("di-confirm").dispatchEvent(new w.Event("click"));
  await sleep(400);
  const issuesAfter = await db.issues.all();
  const capsAfter = await db.captures.forProject("p-del");
  check("Delete issue borra el issue con su foto y su nota de voz", () => {
    assert.ok(!issuesAfter.some((i) => i.id === "iss-1"));
    assert.ok(!capsAfter.some((c) => c.id === photoCap.id || c.id === voiceCap.id));
    assert.match(toastText(), /Issue deleted/);
  });
  check("después de borrar no se sale de la app (sigue en la misma pantalla)", () =>
    assert.match(w.location.hash, /#\/project\/p-del/)
  );

  // ---- v61: lista de tareas (en lugar del botón Document) ----
  w.fetch = async (url, init) => {
    const body = JSON.parse(init.body);
    const isTranslate = /Translate this construction/.test(body.messages[0].content);
    const text = isTranslate ? '{"language":"es","translation":"Change the bathroom faucet"}' : "ok";
    return { ok: true, status: 200, json: async () => ({ content: [{ type: "text", text }], stop_reason: "end_turn" }), text: async () => "", clone() { return this; } };
  };
  await go(w, "#/project/p-del");
  check("el botón Tasks reemplaza a Document", () => {
    assert.ok(doc.querySelector('[data-capture="task"]'), "falta el botón Tasks");
    assert.ok(!doc.querySelector('[data-capture="document"]'), "Document sigue ahí");
  });
  doc.querySelector('[data-capture="task"]').dispatchEvent(new w.Event("click"));
  await sleep(200);
  check("Tasks abre la lista vacía del proyecto", () => assert.match(doc.getElementById("sheet").textContent, /Tasks · Mafakhame[\s\S]*No tasks yet/));
  check("v66: el botón ➕ dice el número de la próxima tarea", () => assert.match(doc.getElementById("tl-new").textContent, /➕ New task #1/));
  doc.getElementById("tl-new").dispatchEvent(new w.Event("click"));
  await sleep(50);
  check("v66: la tarea nueva pide la foto primero (con opción de saltar)", () => {
    assert.ok(doc.getElementById("tp-camera"), "falta Take photo");
    assert.ok(doc.getElementById("tp-gallery"), "falta Gallery");
    assert.ok(doc.getElementById("tp-skip"), "falta Skip");
    assert.ok(!doc.getElementById("f-note"), "el texto no debe salir antes de la foto");
    assert.match(doc.getElementById("sheet").textContent, /New task #1/);
  });
  doc.getElementById("tp-skip").dispatchEvent(new w.Event("click"));
  await sleep(20);
  const fNote = doc.getElementById("f-note");
  fNote.value = "Cambiar la llave del baño";
  fNote.dispatchEvent(new w.Event("input"));
  doc.querySelector('[data-prio="red"]').dispatchEvent(new w.Event("click"));
  await sleep(20);
  check("la tarea muestra los 3 colores y ✨ Make it easy to understand", () => {
    assert.strictEqual(doc.querySelectorAll("[data-prio]").length, 3);
    assert.ok(doc.getElementById("t-clarify"));
    assert.strictEqual(doc.getElementById("f-note").value, "Cambiar la llave del baño", "el texto se pierde al cambiar color");
  });
  doc.getElementById("t-save").dispatchEvent(new w.Event("click"));
  await sleep(300);
  const taskList = async () => (await db.captures.forProject("p-del")).filter((c) => c.type === "task");
  let savedTask = (await taskList())[0];
  check("Save guarda la tarea en rojo (urgente)", () => {
    assert.ok(savedTask, "no se guardó");
    assert.strictEqual(savedTask.priority, "red");
    assert.strictEqual(savedTask.text, "Cambiar la llave del baño");
    assert.strictEqual(savedTask.number, 1, "v66: la primera tarea es la #1");
    assert.match(doc.getElementById("sheet").textContent, /Cambiar la llave del baño[\s\S]*Urgent/);
  });
  check("v62: al guardar, la tarea se traduce y la traducción se guarda", () => {
    assert.strictEqual(savedTask.translation, "Change the bathroom faucet");
    assert.strictEqual(savedTask.translationLang, "en");
    assert.strictEqual(savedTask.translatedFrom, "Cambiar la llave del baño");
  });
  w.eval("openTaskEditSheet")("p-del", savedTask);
  await sleep(50);
  check("v62: editar la tarea muestra la traducción y el botón 🌐", () => {
    assert.match(doc.getElementById("sheet").textContent, /English translation[\s\S]*Change the bathroom faucet/);
    assert.ok(doc.getElementById("t-retr"), "falta el botón Translate");
  });
  w.localStorage.removeItem("moss_anthropic_api_key");
  w.localStorage.setItem("moss_share_lang", "en");
  w.eval("openTaskShareSheet")({ name: "Mafakhame" }, [savedTask]);
  await sleep(200);
  check("v62: el mensaje en inglés usa la traducción guardada", () =>
    assert.match(doc.getElementById("sh-text").value, /URGENT\n- #1 Change the bathroom faucet/)
  );
  w.localStorage.setItem("moss_anthropic_api_key", "sk-ant-test");
  await db.captures.add({ projectId: "p-del", type: "task", text: "Pintar la cerca", priority: "green" });
  await db.captures.add({ projectId: "p-del", type: "task", text: "Pedir azulejo", priority: "yellow" });
  w.eval("openTaskListSheet('p-del')");
  await sleep(200);
  check("la lista va en orden: rojo, amarillo, verde", () => {
    const titles = [...doc.querySelectorAll("[data-task-id] .title")].map((e) => e.textContent);
    assert.deepStrictEqual(titles.map((x) => x.replace(/^#\d+ · /, "")), ["Cambiar la llave del baño", "Pedir azulejo", "Pintar la cerca"]);
    assert.strictEqual(titles[0], "#1 · Cambiar la llave del baño");
    // v66: las tareas que no tenían número reciben #2 y #3
    assert.deepStrictEqual(titles.slice(1).map((x) => x.slice(0, 2)).sort(), ["#2", "#3"]);
  });
  doc.querySelector(`[data-task-done="${savedTask.id}"]`).dispatchEvent(new w.Event("click", { bubbles: true }));
  await sleep(300);
  savedTask = (await taskList()).find((t) => t.id === savedTask.id);
  check("✓ marca la tarea como hecha y la pasa al final", () => {
    assert.strictEqual(savedTask.done, true);
    const titles = [...doc.querySelectorAll("[data-task-id] .title")].map((e) => e.textContent);
    assert.strictEqual(titles[2], "#1 · Cambiar la llave del baño");
  });

  w.localStorage.removeItem("moss_anthropic_api_key"); // sin IA: el mensaje sale tal cual
  doc.getElementById("tl-share").dispatchEvent(new w.Event("click"));
  await sleep(300);
  check("Send the list as text arma el mensaje por colores (solo pendientes)", () => {
    const msg = doc.getElementById("sh-text").value;
    assert.match(msg, /Tasks: Mafakhame/);
    assert.match(msg, /SOON\n- #\d Pedir azulejo/);
    assert.match(msg, /CAN WAIT\n- #\d Pintar la cerca/);
    assert.ok(!/llave/.test(msg), "no debe incluir la tarea hecha");
  });
  w.localStorage.setItem("moss_anthropic_api_key", "sk-ant-test");

  w.eval("openTaskListSheet('p-del')");
  await sleep(200);
  const greenRow = [...doc.querySelectorAll("[data-task-id]")].find((r) => /Pintar/.test(r.textContent));
  greenRow.querySelector(".title").dispatchEvent(new w.Event("click", { bubbles: true }));
  await sleep(50);
  doc.querySelector('[data-prio="yellow"]').dispatchEvent(new w.Event("click"));
  const fn2 = doc.getElementById("f-note");
  fn2.value = "Pintar la cerca de blanco";
  fn2.dispatchEvent(new w.Event("input"));
  doc.getElementById("t-save").dispatchEvent(new w.Event("click"));
  await sleep(300);
  const edited = (await taskList()).find((t) => /Pintar/.test(t.text));
  check("editar cambia el texto y el color sin crear otra tarea", () => {
    assert.strictEqual(edited.text, "Pintar la cerca de blanco");
    assert.strictEqual(edited.priority, "yellow");
  });
  const nums = (await taskList()).map((t) => t.number).sort();
  check("v66: las 3 tareas quedaron guardadas como #1, #2 y #3", () => assert.deepStrictEqual(nums, [1, 2, 3]));
  const countBefore = (await taskList()).length;
  check("siguen siendo 3 tareas", () => assert.strictEqual(countBefore, 3));
  await sleep(200);
  check("v63: el cuadro TASKS reemplaza a COMPLETED y cuenta las pendientes", () => {
    const tile = doc.querySelector('[data-sec-toggle="tasks"]');
    assert.ok(tile, "falta el cuadro TASKS");
    assert.ok(!doc.querySelector('[data-sec-toggle="completed"]'), "COMPLETED sigue ahí");
    assert.strictEqual(tile.querySelector(".num").textContent, "2");
  });
  check("v63: la sección Tasks del proyecto lista las pendientes", () => {
    const rows = [...doc.querySelectorAll("[data-task-row]")].map((r) => r.querySelector(".title").textContent);
    // Las dos son amarillas y se crearon en el mismo milisegundo: el orden entre ellas no importa.
    assert.deepStrictEqual(rows.map((x) => x.replace(/^#\d+ · /, "")).sort(), ["Pedir azulejo", "Pintar la cerca de blanco"]);
  });
  w.eval("openTaskListSheet('p-del')");
  await sleep(200);
  [...doc.querySelectorAll("[data-task-id]")].find((r) => /Pintar/.test(r.textContent)).dispatchEvent(new w.Event("click"));
  await sleep(50);
  doc.getElementById("t-delete").dispatchEvent(new w.Event("click"));
  await sleep(20);
  doc.getElementById("td-confirm").dispatchEvent(new w.Event("click"));
  await sleep(300);
  const after = await taskList();
  check("🗑 Delete task borra solo esa tarea", () => {
    assert.strictEqual(after.length, countBefore - 1);
    assert.ok(!after.some((t) => /Pintar/.test(t.text)));
    assert.match(toastText(), /Task deleted/);
  });
  w.eval("openTaskListSheet('p-del')");
  await sleep(200);
  check("v66: después de borrar, la próxima tarea sigue el número más alto", () => {
    const left = after.map((t) => t.number).sort();
    assert.match(doc.getElementById("tl-new").textContent, new RegExp("New task #" + (Math.max(...left) + 1)));
  });

  // ---- v67: PDF de la lista de tareas ----
  w.eval("openTaskListSheet('p-del')");
  await sleep(200);
  check("v67: la lista tiene el botón 📄 Create PDF", () => assert.ok(doc.getElementById("tl-pdf"), "falta el botón PDF"));
  doc.getElementById("tl-pdf").dispatchEvent(new w.Event("click"));
  await sleep(200);
  check("v67: el PDF deja escoger las tareas (las hechas sin marcar)", () => {
    const boxes = [...doc.querySelectorAll("[data-tpid]")];
    assert.strictEqual(boxes.length, 2);
    assert.strictEqual(boxes.filter((b) => b.checked).length, 1, "solo la pendiente va marcada");
  });
  const rep = w.eval("taskPdfReport")({ name: "Mafakhame" }, await taskList());
  check("v67: el PDF lleva número, color, texto y traducción", () => {
    assert.strictEqual(rep.title, "Task List");
    const txt = JSON.stringify(rep.blocks);
    assert.match(txt, /SOON \(Pronto\)/);
    assert.match(txt, /#\d  -  Soon/);
    assert.match(txt, /Pedir azulejo/);
    assert.match(txt, /DONE \(Hechas\)/);
    assert.match(txt, /#1  -  Done/);
    assert.match(txt, /Change the bathroom faucet/);
  });
  if (!w.URL.createObjectURL) { w.URL.createObjectURL = () => "blob:test"; w.URL.revokeObjectURL = () => {}; } // jsdom no los tiene (el teléfono sí)
  doc.getElementById("tp-pdf-create").dispatchEvent(new w.Event("click"));
  await sleep(1500);
  check("v67: Create PDF arma el PDF y ofrece mandarlo", () => assert.match(doc.getElementById("sheet").textContent, /PDF ready[\s\S]*Tasks - Mafakhame[\s\S]*1 task\b/));

  // ---- v64: Recent Captures — editar el texto de la foto, ✨ y traducir ----
  doc.getElementById("sheet-close")?.dispatchEvent(new w.Event("click"));
  await db.captures.add({ id: "zz-cap", projectId: "p-del", type: "photo", dataUrl: "data:image/jpeg;base64,AAAA", caption: "Water damage on ceiling" });
  w.eval("render()"); // ya estamos en #/project/p-del: volver a dibujar
  await sleep(250);
  doc.querySelector('[data-photo-id="zz-cap"]').dispatchEvent(new w.Event("click"));
  await sleep(50);
  check("v64: la foto tiene Edit text, ✨ y 🌐 Translate", () => {
    assert.ok(doc.getElementById("pv-edit"));
    assert.ok(doc.getElementById("pv-clarify"));
    assert.ok(doc.getElementById("pv-retr"));
    assert.match(doc.getElementById("sheet").textContent, /Water damage on ceiling/);
  });
  doc.getElementById("pv-edit").dispatchEvent(new w.Event("click"));
  await sleep(20);
  doc.getElementById("pe-clarify").dispatchEvent(new w.Event("click"));
  await sleep(200);
  check("v64: ✨ reescribe el texto de la foto y Undo lo regresa", () => {
    assert.strictEqual(doc.getElementById("pe-text").value, "ok");
    doc.getElementById("pe-undo").dispatchEvent(new w.Event("click"));
    assert.strictEqual(doc.getElementById("pe-text").value, "Water damage on ceiling");
  });
  doc.getElementById("pe-text").value = "  ";
  doc.getElementById("pe-save").dispatchEvent(new w.Event("click"));
  await sleep(50);
  check("v64: no deja guardar el texto vacío", () => assert.match(toastText(), /can't be empty/));
  doc.getElementById("pe-text").value = "Mancha de agua en el techo de la cocina";
  doc.getElementById("pe-save").dispatchEvent(new w.Event("click"));
  await sleep(300);
  const phAfter = (await db.captures.forProject("p-del")).find((c) => c.id === "zz-cap");
  check("v64: Save guarda el texto nuevo, el original y la traducción", () => {
    assert.strictEqual(phAfter.caption, "Mancha de agua en el techo de la cocina");
    assert.strictEqual(phAfter.originalCaption, "Water damage on ceiling");
    assert.ok(phAfter.translation, "falta la traducción");
    assert.strictEqual(phAfter.translatedFrom, phAfter.caption);
    assert.match(doc.getElementById("sheet").textContent, /Mancha de agua[\s\S]*English translation/);
  });


  let failed = 0;
  for (const [pass, name] of results) {
    if (!pass) failed++;
    console.log((pass ? "  ✓ " : "  ✗ ") + name);
  }
  console.log(`\n${results.length - failed}/${results.length} pruebas pasaron`);
  process.exit(failed ? 1 : 0);
})().catch((e) => {
  console.error(e);
  process.exit(1);
});
