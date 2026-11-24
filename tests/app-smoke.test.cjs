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

  await go(w, "#/settings");
  check("Settings muestra la versión 54", () => assert.match(doc.getElementById("app").textContent, /App version\s*54/));
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

  check("sin errores de JavaScript al usar la app", () => {
    const real = errors.filter((e) => !/MSAL|msal/i.test(String(e && (e.stack || e))));
    assert.deepStrictEqual(real.map((e) => String(e && (e.stack || e)).split("\n")[0]), []);
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
