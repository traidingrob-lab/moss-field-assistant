// Client-side calls to the Claude API for AI features (Phase 3): "Ask AI",
// photo captioning, and the AI-written report narratives.
//
// This app is a static site with no backend server, so there's nowhere to
// keep an API key secret the way a normal web app would. Instead, the key
// is entered by the user in Settings and stored in this browser's
// localStorage, and calls go straight from the browser to Anthropic's API
// using the "anthropic-dangerous-direct-browser-access" header — the
// header Anthropic added specifically so client-only apps like this one
// can call the API without a server in front of it. The name is a real
// warning, not boilerplate: anyone with access to this device/browser can
// read the key back out of localStorage. That's an acceptable tradeoff for
// a private, single-user deployment like this one, but this key should
// never be pasted into a shared or public device, and this app should
// never be made public with a key already saved in it.

const AI_KEY_STORAGE = "moss_anthropic_api_key";
const AI_MODEL = "claude-sonnet-5";

function aiKey() {
  try {
    return localStorage.getItem(AI_KEY_STORAGE) || "";
  } catch {
    return "";
  }
}

function aiConfigured() {
  return !!aiKey();
}

// Whether photos get an automatic AI description (caption). On by default
// once a key is saved, matching how the app behaved before this switch
// existed. Turning it off only stops NEW descriptions — captions already
// saved stay as they are. The key itself is still needed for Ask AI,
// reports, and New Issue's trade/note analysis; this only governs captions.
const AI_CAPTIONS_STORAGE = "moss_ai_photo_captions";

function aiCaptionsEnabled() {
  if (!aiConfigured()) return false;
  try {
    return localStorage.getItem(AI_CAPTIONS_STORAGE) !== "off";
  } catch {
    return true;
  }
}

function setAiCaptions(enabled) {
  try {
    if (enabled) localStorage.removeItem(AI_CAPTIONS_STORAGE);
    else localStorage.setItem(AI_CAPTIONS_STORAGE, "off");
  } catch {
    // localStorage unavailable — the switch just won't stick; harmless.
  }
}

function setAiKey(key) {
  try {
    if (key) localStorage.setItem(AI_KEY_STORAGE, key);
    else localStorage.removeItem(AI_KEY_STORAGE);
  } catch {
    // localStorage unavailable (private browsing, etc.) — caller's UI
    // will just show "not connected" again next render, which is enough
    // signal without throwing here.
  }
}

// Shared low-level call. `messages` is the Messages API array; `tools` is
// optional (used for web search). Returns the whole response object. Errors
// carry `.status` (HTTP status) so callers can tell "web search isn't
// enabled for this key's organization" (400) from a bad key (401), etc.
async function callClaudeFull(messages, maxTokens = 1024, tools) {
  const key = aiKey();
  if (!key) throw new Error("No Claude API key set — add one in Settings first.");

  const body = { model: AI_MODEL, max_tokens: maxTokens, messages };
  if (tools) body.tools = tools;

  const res = await fetch("https://api.anthropic.com/v1/messages", {
    method: "POST",
    headers: {
      "content-type": "application/json",
      "x-api-key": key,
      "anthropic-version": "2023-06-01",
      "anthropic-dangerous-direct-browser-access": "true"
    },
    body: JSON.stringify(body)
  });

  if (!res.ok) {
    const text = await res.text().catch(() => "");
    let detail = text.slice(0, 300);
    try {
      detail = JSON.parse(text)?.error?.message || detail;
    } catch {
      // body wasn't JSON — fall back to the raw (truncated) text above
    }
    const err = new Error(
      res.status === 401
        ? "That API key was rejected — double check it in Settings."
        : `Claude API error (${res.status}): ${detail}`
    );
    err.status = res.status;
    err.detail = detail;
    throw err;
  }

  return res.json();
}

// `content` is either a plain string (text-only) or an array of Claude
// content blocks (e.g. an image block + a text block), per the Messages API.
// Returns the reply as plain text.
async function callClaude(content, maxTokens = 1024) {
  const data = await callClaudeFull([{ role: "user", content }], maxTokens);
  return (data.content || []).map((block) => block.text || "").join("").trim();
}

// Sends one question to Claude along with whatever project context the
// caller built, and returns the plain-text answer.
async function askClaude(prompt) {
  return callClaude(prompt);
}

// Describes a photo in one short sentence, so Ask AI and the reports can
// later answer questions about what's IN a photo using only that saved
// text — without re-sending the image (cheaper, and it still works if the
// image is later cleared to save space). Called once, right after a photo
// is captured.
//
// NOTE: unlike text questions, this costs one small API call per photo —
// only fires when an API key is present (aiConfigured()), and failures are
// swallowed by the caller so a captioning problem never blocks saving the
// photo itself.
async function captionPhoto(dataUrl) {
  const match = /^data:([^;]+);base64,(.*)$/.exec(dataUrl || "");
  if (!match) throw new Error("Invalid image data.");
  const [, mediaType, base64Data] = match;
  const prompt =
    "Describe this construction job-site photo in one short, specific sentence for a daily log " +
    "(what's shown and its apparent stage/condition — no generic filler like \"a photo of a room\").";
  return callClaude(
    [
      { type: "image", source: { type: "base64", media_type: mediaType, data: base64Data } },
      { type: "text", text: prompt }
    ],
    150
  );
}

// Shrinks a photo before it's sent to Claude. Phone photos can be several
// MB, and the API rejects any single image over 5 MB (base64) — resizing
// the long edge to ~1568px keeps it well under that and is all the detail
// the model uses anyway. Only used for the AI call; the saved photo keeps
// its full resolution. Falls back to the original if anything goes wrong.
function downscaleImageForAI(dataUrl, maxEdge = 1568) {
  return new Promise((resolve) => {
    // If the image never finishes decoding, don't hang the caller — fall
    // back to the original after a few seconds.
    const giveUp = setTimeout(() => resolve(dataUrl), 4000);
    const img = new Image();
    img.onload = () => {
      clearTimeout(giveUp);
      try {
        const scale = Math.min(1, maxEdge / Math.max(img.width, img.height));
        const canvas = document.createElement("canvas");
        canvas.width = Math.max(1, Math.round(img.width * scale));
        canvas.height = Math.max(1, Math.round(img.height * scale));
        canvas.getContext("2d").drawImage(img, 0, 0, canvas.width, canvas.height);
        resolve(canvas.toDataURL("image/jpeg", 0.85));
      } catch {
        resolve(dataUrl);
      }
    };
    img.onerror = () => {
      clearTimeout(giveUp);
      resolve(dataUrl);
    };
    img.src = dataUrl;
  });
}

// New Issue helper: looks at the issue photo plus what was said in the
// voice note, and returns { trade, title, description, note }:
//   trade       — one of `trades` (falls back to "General")
//   title       — short issue title
//   description — one sentence on what the photo shows (used as photo caption)
//   note        — the transcript cleaned up (speech-recognition errors,
//                 punctuation), never embellished
// One API call per new issue; only runs when an API key is saved.
async function analyzeIssueCapture(dataUrl, transcript, trades) {
  const small = await downscaleImageForAI(dataUrl);
  const match = /^data:([^;]+);base64,(.*)$/.exec(small || "");
  if (!match) throw new Error("Invalid image data.");
  const [, mediaType, base64Data] = match;
  const spoken = (transcript || "").trim();

  const prompt =
    "You are helping a general contractor log a job-site issue from a photo and a spoken note.\n" +
    `Pick the single best trade for fixing or handling the issue from this list: ${trades.join(", ")}. ` +
    'Use "General" if it is unclear.\n' +
    `Voice note transcript (may be empty or contain speech-recognition mistakes): """${spoken}"""\n\n` +
    "Reply with ONLY a JSON object (no markdown, no other text) with these keys:\n" +
    '"trade": exactly one item from the list above,\n' +
    '"title": a specific issue title, max 8 words,\n' +
    '"description": one short, specific sentence describing what the photo shows,\n' +
    '"note": the transcript cleaned up — fix obvious speech-recognition errors, punctuation and ' +
    "capitalization, but do not add, remove or invent any information (empty string if the transcript is empty).\n" +
    "Write title, description and note in the same language as the transcript; if there is no transcript, use English.";

  const text = await callClaude(
    [
      { type: "image", source: { type: "base64", media_type: mediaType, data: base64Data } },
      { type: "text", text: prompt }
    ],
    700
  );

  const start = text.indexOf("{");
  const end = text.lastIndexOf("}");
  if (start === -1 || end <= start) throw new Error("AI reply wasn't readable.");
  let parsed;
  try {
    parsed = JSON.parse(text.slice(start, end + 1));
  } catch {
    throw new Error("AI reply wasn't readable.");
  }
  const str = (v) => (typeof v === "string" ? v.trim() : "");
  const trade = trades.find((t) => t.toLowerCase() === str(parsed.trade).toLowerCase()) || "General";
  return { trade, title: str(parsed.title), description: str(parsed.description), note: str(parsed.note) };
}


// ---------- Materials ----------

// Pulls the JSON object out of a model reply. The model is told to answer
// with only JSON, but after web searching it may add a stray sentence, so
// this looks for the last object that actually parses.
function parseJsonReply(text) {
  const s = String(text || "");
  const end = s.lastIndexOf("}");
  if (end === -1) return null;
  let i = s.lastIndexOf("{", end);
  while (i !== -1) {
    try {
      const obj = JSON.parse(s.slice(i, end + 1));
      if (obj && typeof obj === "object") return obj;
    } catch {
      // not valid starting from this "{" — try the one before it
    }
    i = i === 0 ? -1 : s.lastIndexOf("{", i - 1);
  }
  return null;
}

function hostnameOf(url) {
  try {
    return new URL(url).hostname.replace(/^www\./, "");
  } catch {
    return "";
  }
}

// Runs a request that may use web search. When the search runs long the API
// answers "pause_turn"; the paused reply is sent back to let it continue.
// Returns { data, urls } — the final response plus the pages it cited (or,
// failing that, the pages the search returned), as hostnames.
async function runWithSearch(messages, maxTokens, tools) {
  let msgs = messages.slice();
  const cited = [];
  const found = [];
  for (let turn = 0; turn < 4; turn++) {
    const data = await callClaudeFull(msgs, maxTokens, tools);
    for (const block of data.content || []) {
      if (block.type === "web_search_tool_result" && Array.isArray(block.content)) {
        for (const r of block.content) if (r && r.url) found.push(hostnameOf(r.url));
      }
      if (block.type === "text" && Array.isArray(block.citations)) {
        for (const c of block.citations) if (c && c.url) cited.push(hostnameOf(c.url));
      }
    }
    if (data.stop_reason !== "pause_turn") {
      const pick = (cited.length ? cited : found).filter(Boolean);
      return { data, urls: [...new Set(pick)].slice(0, 3) };
    }
    msgs = [...msgs, { role: "assistant", content: data.content }];
  }
  throw new Error("the online search took too long");
}

// Material photo + voice note → { item, dimensions, quantity, note, confidence,
// details, searched, sources }. One request combines BOTH: the photo shows
// what the product is, the dictated note often adds what the photo can't
// (the exact name, brand or size, and how many are needed). Claude then uses
// its web search tool to confirm the real product name and dimensions
// (billed by Anthropic at about $10 per 1,000 searches, at most 3 per request
// here). `transcript` may be empty — then it works from the photo alone.
// If web search isn't available (an organization admin can switch it off in
// the Claude Console, which makes the API answer 400), it retries without it
// and reports `searched: false` — the answer then comes from the photo/note
// alone.
async function identifyMaterial(dataUrl, transcript = "") {
  const small = await downscaleImageForAI(dataUrl);
  const match = /^data:([^;]+);base64,(.*)$/.exec(small || "");
  if (!match) throw new Error("Invalid image data.");
  const [, mediaType, base64Data] = match;
  const spoken = (transcript || "").trim();

  const noteSection = spoken
    ? "The contractor also dictated this voice note about the material. It may contain speech-recognition mistakes, and it often " +
      "names the product, brand or size and says how many are needed:\n" +
      `"""${spoken}"""\n` +
      "Combine the photo and the voice note: the photo shows what it is, the note gives details the photo can't. " +
      "Where the note states a name, size or brand, use it (then verify it online). If the photo and the note disagree, say so in \"details\".\n"
    : "No voice note was recorded, so identify it from the photo alone.\n";

  const prompt =
    "You are helping a general contractor build a materials shopping list from a job-site photo.\n" +
    "Identify the construction material or product in the photo (read any visible label, brand, model or size markings). " +
    "Then use web search to find its exact product name and real dimensions/specs (nominal and actual size, length, thickness, " +
    "gauge, rating, pack or box quantity — whatever applies). Search once or twice and prefer manufacturer or major supplier pages. " +
    "If it is a generic commodity (e.g. 2x4 lumber, 1/2\" drywall), give the standard name and standard dimensions.\n" +
    noteSection +
    "If you cannot tell what it is, say so — do not invent specs.\n" +
    "Do all searching first. Your final message must be ONLY a JSON object (no markdown, no other text) with these keys:\n" +
    '"item": specific product name in English, max 12 words (include brand/model only if visible, stated or confirmed),\n' +
    '"dimensions": size/specs on one line in US units (inches/feet), metric in parentheses only when it is standard; "" if unknown,\n' +
    '"quantity": how many to buy, ONLY if the voice note clearly says (e.g. "12", "2 boxes", "3 sheets"), otherwise "",\n' +
    '"note": the voice note cleaned up — fix obvious recognition errors, punctuation and capitalization, but do not add, remove or ' +
    'invent information, and keep the same language as the note; "" if there is no voice note,\n' +
    '"confidence": "high", "medium" or "low",\n' +
    '"details": one short sentence in English saying what you identified and what you could not confirm.';

  const userMsg = {
    role: "user",
    content: [
      { type: "image", source: { type: "base64", media_type: mediaType, data: base64Data } },
      { type: "text", text: prompt }
    ]
  };

  const search = { type: "web_search_20250305", name: "web_search", max_uses: 3 };
  try {
    const tz = Intl.DateTimeFormat().resolvedOptions().timeZone;
    if (tz) search.user_location = { type: "approximate", timezone: tz };
  } catch {
    // no timezone available — searching works without a location hint
  }

  let result;
  let searched = true;
  try {
    result = await runWithSearch([userMsg], 1800, [search]);
  } catch (err) {
    // 400/403: web search not enabled for this organization. Anything else
    // (bad key, network, rate limit) is a real failure the caller reports.
    if (err.status !== 400 && err.status !== 403) throw err;
    searched = false;
    result = { data: await callClaudeFull([userMsg], 1500), urls: [] };
  }

  const text = (result.data.content || []).map((b) => (b.type === "text" ? b.text || "" : "")).join("");
  const parsed = parseJsonReply(text);
  if (!parsed) throw new Error("AI reply wasn't readable.");
  const str = (v) => (typeof v === "string" ? v.trim() : typeof v === "number" ? String(v) : "");
  const confidence = ["high", "medium", "low"].includes(str(parsed.confidence).toLowerCase())
    ? str(parsed.confidence).toLowerCase()
    : "low";
  return {
    item: str(parsed.item),
    dimensions: str(parsed.dimensions),
    quantity: str(parsed.quantity),
    note: spoken ? str(parsed.note) : "",
    confidence,
    details: str(parsed.details),
    searched,
    sources: searched ? result.urls : []
  };
}

