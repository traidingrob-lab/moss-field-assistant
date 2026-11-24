// Pruebas de ai.js (v54): modelo, esfuerzo, límites, respaldo ante rechazos.
// No gastan créditos: `fetch` es simulado.
// Correr: node tests/ai.test.cjs
"use strict";

const fs = require("fs");
const path = require("path");
const vm = require("vm");
const assert = require("assert");

function loadAi({ economy = true, responses = [] } = {}) {
  const store = new Map([["moss_anthropic_api_key", "sk-ant-test"]]);
  if (!economy) store.set("moss_ai_economy", "off");
  const calls = [];
  const queue = responses.slice();
  const ctx = {
    console,
    URL,
    Intl,
    setTimeout,
    clearTimeout,
    AbortController,
    localStorage: {
      getItem: (k) => (store.has(k) ? store.get(k) : null),
      setItem: (k, v) => store.set(k, String(v)),
      removeItem: (k) => store.delete(k)
    },
    // Sin canvas en Node: la imagen "falla" y se usa la original.
    Image: class {
      set src(_) {
        setTimeout(() => this.onerror && this.onerror(), 0);
      }
    },
    fetch: async (url, init) => {
      calls.push({ url, headers: { ...init.headers }, body: JSON.parse(init.body) });
      const next = queue.shift() || { status: 200, json: { content: [{ type: "text", text: "ok" }], stop_reason: "end_turn" } };
      if (next.networkError) throw new TypeError("Failed to fetch");
      if (next.hang) {
        // Nunca responde: solo termina si la app cancela (timeout).
        return new Promise((_, reject) => {
          init.signal && init.signal.addEventListener("abort", () => {
            const e = new Error("The operation was aborted.");
            e.name = "AbortError";
            reject(e);
          });
        });
      }
      const text = typeof next.json === "string" ? next.json : JSON.stringify(next.json);
      const make = () => ({
        ok: next.status >= 200 && next.status < 300,
        status: next.status,
        text: async () => text,
        json: async () => JSON.parse(text),
        clone: make
      });
      return make();
    }
  };
  vm.createContext(ctx);
  const src = fs.readFileSync(path.join(__dirname, "..", "ai.js"), "utf8");
  vm.runInContext(src + "\n;this.__ai = { callClaude, askClaude, captionPhoto, translateVoiceNote, translateIssueTitle, analyzeIssueCapture, analyzeInspectionCapture, identifyMaterial, aiEconomyEnabled, setAiEconomy, aiSelfTest, callClaudeFull };", ctx);
  return { ai: ctx.__ai, calls, store };
}

const IMG = "data:image/jpeg;base64,AAAA";
const ok = (text, extra = {}) => ({ status: 200, json: { content: [{ type: "text", text }], stop_reason: "end_turn", ...extra } });

const tests = [];
const test = (name, fn) => tests.push({ name, fn });

test("caption usa Haiku 5.5, esfuerzo bajo, sin respaldo", async () => {
  const { ai, calls } = loadAi({ responses: [ok("Drywall hung, not taped.")] });
  const out = await ai.captionPhoto(IMG);
  assert.strictEqual(out, "Drywall hung, not taped.");
  const b = calls[0].body;
  assert.strictEqual(b.model, "claude-haiku-5-5");
  assert.deepStrictEqual(b.output_config, { effort: "low" });
  assert.strictEqual(b.max_tokens, 150 + 2000);
  assert.strictEqual(b.fallbacks, undefined);
  assert.strictEqual(calls[0].headers["anthropic-beta"], undefined);
  assert.strictEqual(b.thinking, undefined);
  assert.strictEqual(b.temperature, undefined);
});

test("con Economy apagado, caption usa Sonnet 5.5 con respaldo", async () => {
  const { ai, calls } = loadAi({ economy: false, responses: [ok("x")] });
  await ai.captionPhoto(IMG);
  const b = calls[0].body;
  assert.strictEqual(b.model, "claude-sonnet-5-5");
  assert.deepStrictEqual(b.output_config, { effort: "low" });
  assert.strictEqual(b.fallbacks, "default");
  assert.strictEqual(calls[0].headers["anthropic-beta"], "server-side-fallback-2026-07-01");
});

test("Ask AI usa Sonnet 5.5, esfuerzo medio y margen para pensar", async () => {
  const { ai, calls } = loadAi({ responses: [ok("Answer")] });
  assert.strictEqual(await ai.askClaude("q"), "Answer");
  const b = calls[0].body;
  assert.strictEqual(b.model, "claude-sonnet-5-5");
  assert.deepStrictEqual(b.output_config, { effort: "medium" });
  assert.strictEqual(b.max_tokens, 1024 + 4000);
  assert.strictEqual(b.fallbacks, "default");
});

test("los bloques de pensamiento no se mezclan con la respuesta", async () => {
  const { ai } = loadAi({
    responses: [{ status: 200, json: { content: [{ type: "thinking", thinking: "" }, { type: "text", text: "Hola" }], stop_reason: "end_turn" } }]
  });
  assert.strictEqual(await ai.askClaude("q"), "Hola");
});

test("si el API rechaza el respaldo (400), reintenta sin él y no lo vuelve a enviar", async () => {
  const { ai, calls } = loadAi({
    responses: [
      { status: 400, json: { error: { message: "fallbacks: unknown beta" } } },
      ok("first"),
      ok("second")
    ]
  });
  assert.strictEqual(await ai.askClaude("q"), "first");
  assert.strictEqual(calls.length, 2);
  assert.strictEqual(calls[1].body.fallbacks, undefined);
  assert.strictEqual(calls[1].headers["anthropic-beta"], undefined);
  await ai.askClaude("q2");
  assert.strictEqual(calls[2].body.fallbacks, undefined);
});

test("si el navegador bloquea la llamada con respaldo (error de red), reintenta sin él", async () => {
  const { ai, calls } = loadAi({ responses: [{ networkError: true }, ok("fine")] });
  assert.strictEqual(await ai.askClaude("q"), "fine");
  assert.strictEqual(calls.length, 2);
  assert.strictEqual(calls[1].headers["anthropic-beta"], undefined);
});

test("en Haiku un error de red no se reintenta", async () => {
  const { ai, calls } = loadAi({ responses: [{ networkError: true }] });
  await assert.rejects(ai.captionPhoto(IMG), /Failed to fetch/);
  assert.strictEqual(calls.length, 1);
});

test("un 400 que no es del respaldo no se reintenta", async () => {
  const { ai, calls } = loadAi({ responses: [{ status: 400, json: { error: { message: "image too large" } } }] });
  await assert.rejects(ai.askClaude("q"), /image too large/);
  assert.strictEqual(calls.length, 1);
});

test("un rechazo por seguridad da un error claro", async () => {
  const { ai } = loadAi({ responses: [{ status: 200, json: { content: [], stop_reason: "refusal" } }] });
  await assert.rejects(ai.askClaude("q"), /declined/);
});

test("clave inválida (401) da el mensaje de siempre", async () => {
  const { ai } = loadAi({ responses: [{ status: 401, json: { error: { message: "invalid x-api-key" } } }] });
  await assert.rejects(ai.askClaude("q"), /API key was rejected/);
});

test("traducción de nota de voz usa Haiku y se interpreta bien", async () => {
  const { ai, calls } = loadAi({ responses: [ok('{"language":"es","translation":"The pipe leaks."}')] });
  const r = await ai.translateVoiceNote("El tubo gotea.");
  assert.deepStrictEqual(JSON.parse(JSON.stringify(r)), { lang: "es", target: "en", translation: "The pipe leaks." });
  assert.strictEqual(calls[0].body.model, "claude-haiku-5-5");
  assert.strictEqual(calls[0].body.max_tokens, 1000 + 2000);
});

test("título de issue usa Haiku", async () => {
  const { ai, calls } = loadAi({ responses: [ok("Fuga en tubería")] });
  assert.strictEqual(await ai.translateIssueTitle("Pipe leak", "es"), "Fuga en tubería");
  assert.strictEqual(calls[0].body.model, "claude-haiku-5-5");
});

test("análisis de New Issue usa Haiku y conserva la nota fiel", async () => {
  const reply = '{"trade":"Plumbing","title":"Leak under sink","description":"Wet cabinet floor.","note":"La tubería gotea debajo del fregadero."}';
  const { ai, calls } = loadAi({ responses: [ok(reply)] });
  const r = await ai.analyzeIssueCapture(IMG, "la tubería gotea debajo del fregadero", ["General", "Plumbing"], "en");
  assert.strictEqual(r.trade, "Plumbing");
  assert.strictEqual(r.note, "La tubería gotea debajo del fregadero.");
  assert.strictEqual(calls[0].body.model, "claude-haiku-5-5");
});

test("inspección usa Sonnet con esfuerzo bajo", async () => {
  const { ai, calls } = loadAi({ responses: [ok('{"title":"Framing inspection","note":"Todo bien."}')] });
  const r = await ai.analyzeInspectionCapture([IMG], "todo bien", "en");
  assert.strictEqual(r.title, "Framing inspection");
  assert.strictEqual(calls[0].body.model, "claude-sonnet-5-5");
  assert.deepStrictEqual(calls[0].body.output_config, { effort: "low" });
});

test("materiales: Sonnet con búsqueda web, continúa tras pause_turn y conserva el modelo", async () => {
  const final = '{"item":"Colonial baseboard, primed MDF","dimensions":"3-1/4 in x 9/16 in x 16 ft","quantity":"","note":"","confidence":"high","details":"ok"}';
  const { ai, calls } = loadAi({
    responses: [
      { status: 200, json: { content: [{ type: "web_search_tool_result", content: [{ url: "https://www.homedepot.com/x" }] }], stop_reason: "pause_turn" } },
      { status: 200, json: { content: [{ type: "text", text: final, citations: [{ url: "https://www.lowes.com/y" }] }], stop_reason: "end_turn" } }
    ]
  });
  const r = await ai.identifyMaterial(IMG, "");
  assert.strictEqual(r.item, "Colonial baseboard, primed MDF");
  assert.strictEqual(r.searched, true);
  assert.deepStrictEqual(Array.from(r.sources), ["lowes.com"]);
  assert.strictEqual(calls.length, 2);
  for (const c of calls) {
    assert.strictEqual(c.body.model, "claude-sonnet-5-5");
    assert.strictEqual(c.body.tools[0].type, "web_search_20250305");
    assert.strictEqual(c.body.max_tokens, 1800 + 4000);
  }
});

test("materiales: sin búsqueda web habilitada (400) reintenta sin herramientas", async () => {
  const final = '{"item":"2x4 stud","dimensions":"1.5 x 3.5 in","quantity":"","note":"","confidence":"medium","details":"x"}';
  const { ai, calls } = loadAi({
    responses: [{ status: 400, json: { error: { message: "web search is not enabled" } } }, ok(final)]
  });
  const r = await ai.identifyMaterial(IMG, "");
  assert.strictEqual(r.searched, false);
  assert.strictEqual(calls[1].body.tools, undefined);
});

test("si Claude nunca responde, se cancela con un error claro (v56)", async () => {
  const { ai, calls } = loadAi({ responses: [{ hang: true }] });
  const started = Date.now();
  await assert.rejects(
    ai.callClaudeFull([{ role: "user", content: "x" }], 10, undefined, { tier: "fast", timeoutMs: 300 }),
    /No answer from Claude after 0 seconds \(claude-haiku-5-5\)/
  );
  assert.ok(Date.now() - started < 3000);
  assert.strictEqual(calls.length, 1, "un timeout no debe reintentarse");
});

test("un error de red dice qué modelo falló", async () => {
  const { ai } = loadAi({ responses: [{ networkError: true }] });
  await assert.rejects(ai.captionPhoto(IMG), /Couldn't reach Claude \(claude-haiku-5-5\): Failed to fetch/);
});

test("Test AI prueba Haiku y Sonnet y reporta cada resultado", async () => {
  const { ai, calls } = loadAi({
    responses: [ok("OK"), { status: 404, json: { error: { message: "model: claude-sonnet-5-5 not found" } } }]
  });
  const r = JSON.parse(JSON.stringify(await ai.aiSelfTest()));
  assert.strictEqual(r.length, 2);
  assert.strictEqual(r[0].model, "claude-haiku-5-5");
  assert.strictEqual(r[0].ok, true);
  assert.strictEqual(r[0].text, "OK");
  assert.strictEqual(r[1].model, "claude-sonnet-5-5");
  assert.strictEqual(r[1].ok, false);
  assert.match(r[1].error, /404.*not found/);
  assert.strictEqual(calls[0].body.model, "claude-haiku-5-5");
  assert.strictEqual(calls[1].body.model, "claude-sonnet-5-5");
});

test("Test AI usa el modelo pedido aunque Economy esté apagado", async () => {
  const { ai, calls } = loadAi({ economy: false, responses: [ok("OK"), ok("OK")] });
  await ai.aiSelfTest();
  assert.strictEqual(calls[0].body.model, "claude-haiku-5-5");
});

test("el interruptor Economy se guarda y se lee", async () => {
  const { ai, store } = loadAi();
  assert.strictEqual(ai.aiEconomyEnabled(), true);
  ai.setAiEconomy(false);
  assert.strictEqual(store.get("moss_ai_economy"), "off");
  assert.strictEqual(ai.aiEconomyEnabled(), false);
  ai.setAiEconomy(true);
  assert.strictEqual(ai.aiEconomyEnabled(), true);
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
