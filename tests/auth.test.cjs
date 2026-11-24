// Pruebas de auth.js (v59): una sincronización en segundo plano NUNCA debe
// sacar la app hacia la página de Microsoft; solo una acción que la persona
// tocó (interactive) puede hacerlo. MSAL es simulado.
// Correr: node tests/auth.test.cjs
"use strict";

const fs = require("fs");
const path = require("path");
const vm = require("vm");
const assert = require("assert");

function loadAuth({ silentFails = false, signedIn = true } = {}) {
  const log = { redirects: 0, notices: [] };
  const account = { username: "rob@example.com" };
  class PublicClientApplication {
    constructor() {
      this.active = signedIn ? account : null;
    }
    initialize() { return Promise.resolve(); }
    handleRedirectPromise() { return Promise.resolve(null); }
    getActiveAccount() { return this.active; }
    getAllAccounts() { return signedIn ? [account] : []; }
    setActiveAccount(a) { this.active = a; }
    acquireTokenSilent() {
      return silentFails ? Promise.reject(new Error("interaction_required")) : Promise.resolve({ accessToken: "tok-123" });
    }
    acquireTokenRedirect() { log.redirects++; return Promise.resolve(); }
    loginRedirect() { log.redirects++; return Promise.resolve(); }
  }
  const ctx = {
    console: { log() {}, warn() {}, error() {} },
    msal: { PublicClientApplication },
    MSAL_CLIENT_ID: "bdddaef4-test",
    onMsSignInNeeded: (err) => log.notices.push(err && err.message)
  };
  vm.createContext(ctx);
  vm.runInContext(fs.readFileSync(path.join(__dirname, "..", "auth.js"), "utf8") + "\n;this.__a = { msGetToken };", ctx);
  return { auth: ctx.__a, log };
}

const tests = [];
const test = (name, fn) => tests.push({ name, fn });

test("con sesión válida devuelve el token sin salir de la app", async () => {
  const { auth, log } = loadAuth();
  assert.strictEqual(await auth.msGetToken(), "tok-123");
  assert.strictEqual(log.redirects, 0);
  assert.deepStrictEqual(log.notices, []);
});

test("sesión vencida en segundo plano: NO sale de la app y avisa", async () => {
  const { auth, log } = loadAuth({ silentFails: true });
  assert.strictEqual(await auth.msGetToken(), null);
  assert.strictEqual(log.redirects, 0, "no debe ir a la página de Microsoft");
  assert.deepStrictEqual(log.notices, ["interaction_required"]);
});

test("sesión vencida al tocar Save now (interactive): sí va a iniciar sesión", async () => {
  const { auth, log } = loadAuth({ silentFails: true });
  assert.strictEqual(await auth.msGetToken({ interactive: true }), null);
  assert.strictEqual(log.redirects, 1);
});

test("sin cuenta conectada no hace nada", async () => {
  const { auth, log } = loadAuth({ signedIn: false });
  assert.strictEqual(await auth.msGetToken(), null);
  assert.strictEqual(log.redirects, 0);
  assert.deepStrictEqual(log.notices, []);
});

test("todas las llamadas de app.js son en segundo plano salvo Save now", async () => {
  const src = fs.readFileSync(path.join(__dirname, "..", "app.js"), "utf8");
  const calls = src.match(/msGetToken\([^)]*\)/g) || [];
  const interactive = calls.filter((c) => /interactive:\s*true/.test(c));
  assert.strictEqual(interactive.length, 1, "solo backupToOneDrive (Save now) puede ser interactiva");
  const idx = src.indexOf("msGetToken({ interactive: true })");
  const fnStart = src.lastIndexOf("async function ", idx);
  assert.match(src.slice(fnStart, fnStart + 60), /async function backupToOneDrive/);
});

(async () => {
  let failed = 0;
  for (const t of tests) {
    try {
      await t.fn();
      console.log("  ✓ " + t.name);
    } catch (e) {
      failed++;
      console.log("  ✗ " + t.name + "\n    " + (e && e.message));
    }
  }
  console.log(`\n${tests.length - failed}/${tests.length} pruebas pasaron`);
  process.exit(failed ? 1 : 0);
})();
