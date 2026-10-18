// Replaces emoji with line icons in a colored border (v53). The emoji is kept
// (hidden) inside <span class="ei ei-NAME"> so text content never changes;
// style.css draws the icon. Runs on load and on every DOM change.
(function () {
  const MAP = {"🎤": "mic", "🎙": "mic", "📋": "clip", "📷": "cam", "📸": "cam", "🗑": "trash", "📄": "pdf", "🌐": "globe", "🤖": "ai", "✅": "ok", "⚠": "alert", "🏛": "bldg", "🗄": "arch", "📦": "box", "💬": "msg", "🖼": "img", "⏹": "stop", "🏗": "hat", "☁": "cloud", "📅": "cal", "🗓": "cal", "💾": "save", "📤": "up", "📥": "down", "🔊": "vol", "🔎": "search", "❌": "xc", "🏠": "home", "⚙": "gear", "➡": "right", "✏": "pen", "📝": "note", "🪜": "steps", "⬆": "top", "📲": "phone", "👁": "eye", "⏳": "hour", "ℹ": "info", "↩": "undo"};
  const RE = new RegExp("(" + Object.keys(MAP).join("|") + ")\\uFE0F?", "g");
  const SKIP = { SCRIPT: 1, STYLE: 1, TEXTAREA: 1, OPTION: 1, SELECT: 1, TITLE: 1, INPUT: 1, NOSCRIPT: 1 };
  function walk(root) {
    if (!root) return;
    const w = document.createTreeWalker(root, NodeFilter.SHOW_TEXT, {
      acceptNode(n) {
        const p = n.parentNode;
        if (!p || SKIP[p.nodeName] || (p.classList && p.classList.contains("ei"))) return NodeFilter.FILTER_REJECT;
        RE.lastIndex = 0;
        return RE.test(n.nodeValue) ? NodeFilter.FILTER_ACCEPT : NodeFilter.FILTER_REJECT;
      }
    });
    const nodes = [];
    while (w.nextNode()) nodes.push(w.currentNode);
    nodes.forEach((n) => {
      const text = n.nodeValue, frag = document.createDocumentFragment();
      let last = 0, m;
      RE.lastIndex = 0;
      while ((m = RE.exec(text))) {
        if (m.index > last) frag.appendChild(document.createTextNode(text.slice(last, m.index)));
        const s = document.createElement("span");
        s.className = "ei ei-" + MAP[m[1]];
        s.setAttribute("aria-hidden", "true");
        s.textContent = m[0];
        frag.appendChild(s);
        last = m.index + m[0].length;
      }
      if (last < text.length) frag.appendChild(document.createTextNode(text.slice(last)));
      n.parentNode.replaceChild(frag, n);
    });
  }
  let queued = false;
  function run() { queued = false; if (document.body) walk(document.body); }
  function queue() { if (!queued) { queued = true; (window.requestAnimationFrame || setTimeout)(run); } }
  window.mossIcons = { run: run, walk: walk, MAP: MAP };
  function start() { run(); new MutationObserver(queue).observe(document.body, { childList: true, subtree: true, characterData: true }); }
  if (document.readyState === "loading") document.addEventListener("DOMContentLoaded", start); else start();
})();
