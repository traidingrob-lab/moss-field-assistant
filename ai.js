// Client-side calls to the Claude API for "Ask AI" (Phase 3).
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

// Sends one question to Claude along with whatever project context the
// caller built, and returns the plain-text answer.
async function askClaude(prompt) {
  const key = aiKey();
  if (!key) throw new Error("No Claude API key set — add one in Settings first.");

  const res = await fetch("https://api.anthropic.com/v1/messages", {
    method: "POST",
    headers: {
      "content-type": "application/json",
      "x-api-key": key,
      "anthropic-version": "2023-06-01",
      "anthropic-dangerous-direct-browser-access": "true"
    },
    body: JSON.stringify({
      model: AI_MODEL,
      max_tokens: 1024,
      messages: [{ role: "user", content: prompt }]
    })
  });

  if (!res.ok) {
    const body = await res.text().catch(() => "");
    let detail = body.slice(0, 300);
    try {
      detail = JSON.parse(body)?.error?.message || detail;
    } catch {
      // body wasn't JSON — fall back to the raw (truncated) text above
    }
    if (res.status === 401) throw new Error("That API key was rejected — double check it in Settings.");
    throw new Error(`Claude API error (${res.status}): ${detail}`);
  }

  const data = await res.json();
  return (data.content || []).map((block) => block.text || "").join("").trim();
}
