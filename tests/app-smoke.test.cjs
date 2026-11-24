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
