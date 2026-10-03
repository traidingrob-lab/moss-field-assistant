// Builds the "Open Issues" PDF (project name, each issue with its photo and
// the written text of its voice note). Runs entirely in the browser using
// jsPDF, vendored locally as jspdf.umd.min.js (index.html's CSP only allows
// scripts from this same site, so it can't come from a CDN).
//
// buildIssuesPdf(sections) takes plain data — no database access in here:
//   sections = [{ projectName, address, issues: [{ title, trade, stamp, note, photoDataUrl }] }]
// and resolves to a PDF Blob.

// jsPDF's built-in fonts only cover Latin-1 (English + Spanish accents/ñ/¿/¡
// are fine). Anything outside that — emoji, curly quotes, the narrow no-break
// space newer browsers put before "PM" — would print as garbage, so it's
// swapped for a plain equivalent or dropped.
function pdfSafe(str) {
  return String(str ?? "")
    .replace(/\r/g, "")
    .replace(/\t/g, " ")
    .replace(/[‘’‚′]/g, "'")
    .replace(/[“”„″]/g, '"')
    .replace(/[–—−]/g, "-")
    .replace(/…/g, "...")
    .replace(/[  -​  　]/g, " ")
    .replace(/[^\n -~¡-ÿ]/g, "");
}

// Shrinks a photo for the PDF. Phone photos are several MB each; embedding
// them as-is would make a PDF far too large to email or text. ~1100px on the
// long edge at JPEG 80% is sharp on a page and ~100-250 KB per photo.
// Resolves to { dataUrl, width, height }, or null if the image can't be
// decoded (caller just leaves the photo out).
function prepareImageForPdf(dataUrl, maxEdge = 1100, quality = 0.8) {
  // Phones often hand back a blank (all-white) picture when a big photo is
  // drawn straight onto a canvas, with no error at all. So: shrink it in a
  // memory-friendly way, LOOK at the result, and if it is blank try again
  // with another method / a smaller size. Resolves to
  // { dataUrl, width, height } or null (caller shows a "couldn't be
  // included" note instead of a blank box).
  const withTimeout = (promise, ms) =>
    Promise.race([promise, new Promise((resolve) => setTimeout(() => resolve(null), ms))]);

  const loadImg = () =>
    new Promise((resolve) => {
      const img = new Image();
      img.onload = async () => {
        try { if (img.decode) await img.decode(); } catch {}
        resolve(img);
      };
      img.onerror = () => resolve(null);
      img.src = dataUrl;
    });

  const isBlank = (canvas) => {
    try {
      const probe = document.createElement("canvas");
      probe.width = 16;
      probe.height = 16;
      const pctx = probe.getContext("2d");
      pctx.drawImage(canvas, 0, 0, 16, 16);
      const d = pctx.getImageData(0, 0, 16, 16).data;
      let min = 255;
      let max = 0;
      for (let i = 0; i < d.length; i += 4) {
        const v = (d[i] + d[i + 1] + d[i + 2]) / 3;
        if (v < min) min = v;
        if (v > max) max = v;
      }
      probe.width = probe.height = 1;
      return min > 247 || (max < 8 && min < 8); // all white, or all black
    } catch {
      return false; // can't check: trust it
    }
  };

  const toResult = (src, srcW, srcH, edge) => {
    const scale = Math.min(1, edge / Math.max(srcW, srcH));
    const canvas = document.createElement("canvas");
    canvas.width = Math.max(1, Math.round(srcW * scale));
    canvas.height = Math.max(1, Math.round(srcH * scale));
    const ctx = canvas.getContext("2d");
    ctx.fillStyle = "#fff"; // JPEG has no transparency
    ctx.fillRect(0, 0, canvas.width, canvas.height);
    ctx.drawImage(src, 0, 0, canvas.width, canvas.height);
    if (isBlank(canvas)) {
      canvas.width = canvas.height = 1;
      return null;
    }
    const out = { dataUrl: canvas.toDataURL("image/jpeg", quality), width: canvas.width, height: canvas.height };
    canvas.width = canvas.height = 1; // let the phone free the memory
    return out;
  };

  const attempt = async () => {
    if (!dataUrl) return null;
    const img = await loadImg();
    let w = img ? img.naturalWidth || img.width : 0;
    let h = img ? img.naturalHeight || img.height : 0;
    const edges = [maxEdge, Math.min(maxEdge, 800), Math.min(maxEdge, 500)];
    let blob = null;
    const getBlob = async () => {
      if (!blob) {
        // (No fetch(): the app's security policy blocks data: URLs there.)
        const m = /^data:([^;,]+)?(;base64)?,(.*)$/s.exec(dataUrl);
        if (!m) throw new Error("bad data URL");
        const bin = m[2] ? atob(m[3]) : decodeURIComponent(m[3]);
        const bytes = new Uint8Array(bin.length);
        for (let i = 0; i < bin.length; i++) bytes[i] = bin.charCodeAt(i);
        blob = new Blob([bytes], { type: m[1] || "image/jpeg" });
      }
      return blob;
    };
    const release = () => { if (img) img.src = ""; };
    for (const edge of edges) {
      // Method 1: decode straight from the file at the small size (needs far
      // less memory than drawing the full-size photo).
      if (typeof createImageBitmap === "function") {
        try {
          const bl = await getBlob();
          let opts;
          if (w && h) {
            const sc = Math.min(1, edge / Math.max(w, h));
            opts = { resizeWidth: Math.max(1, Math.round(w * sc)), resizeHeight: Math.max(1, Math.round(h * sc)), resizeQuality: "medium" };
          }
          const bmp = await createImageBitmap(bl, opts);
          const r = toResult(bmp, bmp.width, bmp.height, edge);
          if (bmp.close) bmp.close();
          if (r) { release(); return r; }
        } catch {}
      }
      // Method 2: draw the <img> element.
      if (img && w && h) {
        try {
          const r = toResult(img, w, h, edge);
          if (r) { release(); return r; }
        } catch {}
      }
    }
    release();
    return null;
  };

  // One at a time with a short breather, so Safari can free the previous
  // photo's memory before the next one is decoded.
  return withTimeout(
    attempt().then(async (r) => {
      await new Promise((res) => setTimeout(res, 60));
      return r;
    }),
    40000
  );
}

async function buildIssuesPdf(sections, opts = {}) {
  if (!window.jspdf || !window.jspdf.jsPDF) throw new Error("PDF library (jspdf.umd.min.js) isn't loaded.");
  const { jsPDF } = window.jspdf;
  const prepareImage = opts.prepareImage || prepareImageForPdf;
  const generatedAt = opts.generatedAt || new Date();

  const doc = new jsPDF({ unit: "pt", format: "letter" });
  const pageW = doc.internal.pageSize.getWidth();
  const pageH = doc.internal.pageSize.getHeight();
  const M = 42; // page margin
  const contentW = pageW - M * 2;
  const bottom = pageH - 56; // leave room for the footer
  const INK = [28, 37, 48];
  const SOFT = [107, 114, 128];
  const AMBER = [217, 119, 6];
  let y = M;

  const ensure = (h) => {
    if (y + h > bottom) {
      doc.addPage();
      y = M;
    }
  };
  const color = (c) => doc.setTextColor(c[0], c[1], c[2]);
  const font = (style, size) => {
    doc.setFont("helvetica", style);
    doc.setFontSize(size);
  };
  const rule = (c, width) => {
    doc.setDrawColor(c[0], c[1], c[2]);
    doc.setLineWidth(width);
    doc.line(M, y, pageW - M, y);
  };

  let dateText;
  try {
    dateText = generatedAt.toLocaleDateString(undefined, { weekday: "long", year: "numeric", month: "long", day: "numeric" });
  } catch {
    dateText = generatedAt.toDateString();
  }
  dateText = pdfSafe(dateText);

  const totalIssues = sections.reduce((n, s) => n + s.issues.length, 0);

  // ----- Report header -----
  font("bold", 10);
  color(AMBER);
  doc.text("MOSS AI FIELD ASSISTANT", M, y + 8);
  y += 28;
  font("bold", 22);
  color(INK);
  doc.text("Open Issues Report", M, y);
  y += 18;
  font("normal", 10);
  color(SOFT);
  doc.text(`${dateText}  -  ${totalIssues} issue${totalIssues === 1 ? "" : "s"}`, M, y);
  y += 12;
  rule(AMBER, 1.5);
  y += 22;

  for (let s = 0; s < sections.length; s++) {
    const section = sections[s];

    // ----- Project heading -----
    ensure(70);
    font("bold", 16);
    color(INK);
    const nameLines = doc.splitTextToSize(pdfSafe(section.projectName), contentW);
    doc.text(nameLines, M, y + 12);
    y += nameLines.length * 19 + 2;
    if (section.address) {
      font("normal", 10);
      color(SOFT);
      const addrLines = doc.splitTextToSize(pdfSafe(section.address), contentW);
      doc.text(addrLines, M, y + 8);
      y += addrLines.length * 12 + 2;
    }
    font("normal", 9);
    color(SOFT);
    doc.text(`${section.issues.length} open issue${section.issues.length === 1 ? "" : "s"}`, M, y + 8);
    y += 18;

    for (let i = 0; i < section.issues.length; i++) {
      const issue = section.issues[i];
      const photo = issue.photoDataUrl ? await prepareImage(issue.photoDataUrl) : null;

      // Photo box: up to 300pt wide / 225pt tall, keeping its proportions.
      let pw = 0;
      let ph = 0;
      if (photo) {
        const scale = Math.min(300 / photo.width, 225 / photo.height);
        pw = photo.width * scale;
        ph = photo.height * scale;
      }

      font("bold", 12.5);
      const titleLines = doc.splitTextToSize(`${i + 1}. ${pdfSafe(issue.title)}`, contentW);
      const titleH = titleLines.length * 15;
      const metaText = pdfSafe([issue.trade, issue.stamp].filter(Boolean).join("  -  "));
      font("normal", 9);
      const metaLines = doc.splitTextToSize(metaText, contentW);
      const metaH = metaLines.length * 11;

      // Keep the title, its details and its photo together on one page.
      ensure(titleH + metaH + (ph ? ph + 10 : 0) + 8);

      font("bold", 12.5);
      color(INK);
      doc.text(titleLines, M, y + 11);
      y += titleH + 1;

      font("normal", 9);
      color(SOFT);
      doc.text(metaLines, M, y + 8);
      y += metaH + 6;

      if (photo) {
        doc.addImage(photo.dataUrl, "JPEG", M, y, pw, ph);
        doc.setDrawColor(220, 224, 229);
        doc.setLineWidth(0.5);
        doc.rect(M, y, pw, ph);
        y += ph + 10;
      } else if (issue.photoDataUrl) {
        font("italic", 9);
        color(SOFT);
        ensure(14);
        doc.text("(Photo couldn't be included in this PDF)", M, y + 9);
        y += 16;
      }

      // Extra photos (picked from the gallery): a row of small pictures.
      const extraList = [];
      let missingExtras = 0;
      for (const d of issue.extraPhotoDataUrls || []) {
        const pr = await prepareImage(d, 700, 0.75);
        if (pr) extraList.push(pr);
        else missingExtras++;
      }
      if (missingExtras) {
        font("italic", 9);
        color(SOFT);
        ensure(14);
        doc.text(`(${missingExtras} photo${missingExtras === 1 ? "" : "s"} couldn't be included in this PDF)`, M, y + 9);
        y += 16;
      }
      if (extraList.length) {
        const gap = 6;
        const cellW = (contentW - gap * 3) / 4;
        const cellH = 105;
        ensure(cellH + 8);
        extraList.forEach((pr, k) => {
          const sc = Math.min(cellW / pr.width, cellH / pr.height);
          const w2 = pr.width * sc;
          const h2 = pr.height * sc;
          const x2 = M + k * (cellW + gap);
          doc.addImage(pr.dataUrl, "JPEG", x2, y, w2, h2);
          doc.setDrawColor(220, 224, 229);
          doc.setLineWidth(0.5);
          doc.rect(x2, y, w2, h2);
        });
        y += cellH + 10;
      }

      const note = pdfSafe(issue.note).trim();
      if (note) {
        font("normal", 10.5);
        color(INK);
        const lines = doc.splitTextToSize(note, contentW);
        // Line by line, so a long note can continue onto the next page.
        for (const line of lines) {
          ensure(14);
          doc.text(line, M, y + 10);
          y += 14;
        }
      } else {
        font("italic", 10);
        color(SOFT);
        ensure(14);
        doc.text("(No written note)", M, y + 10);
        y += 14;
      }

      // Spanish <-> English translation of the note, when there is one.
      const tr = pdfSafe(issue.translation).trim();
      if (tr) {
        y += 4;
        font("bold", 9);
        color(SOFT);
        ensure(13);
        doc.text(pdfSafe(issue.translationLabel || "Translation"), M, y + 9);
        y += 13;
        font("italic", 10.5);
        color(INK);
        for (const line of doc.splitTextToSize(tr, contentW)) {
          ensure(14);
          doc.text(line, M, y + 10);
          y += 14;
        }
      }

      y += 10;
      // Separator line — skipped when the next issue starts a new page anyway,
      // so a stray line is never left alone at the top of a page.
      if (i < section.issues.length - 1 && y + 20 <= bottom) {
        rule([229, 231, 235], 0.5);
        y += 16;
      }
    }
    y += 14;
  }

  // ----- Footer on every page -----
  const pages = doc.getNumberOfPages();
  for (let p = 1; p <= pages; p++) {
    doc.setPage(p);
    font("normal", 8);
    color(SOFT);
    doc.text(`Moss AI Field Assistant  -  ${dateText}`, M, pageH - 28);
    doc.text(`Page ${p} of ${pages}`, pageW - M, pageH - 28, { align: "right" });
  }

  return doc.output("blob");
}

// ---------- Materials shopping list ----------
// sections = [{ projectName, address, items: [{ name, dimensions, quantity, status, note, stamp, photoDataUrl }] }]
// One row per material: a tick box (to cross off at the store), a small photo,
// the name in bold, size/specs, quantity, and the written voice note.
async function buildMaterialsPdf(sections, opts = {}) {
  if (!window.jspdf || !window.jspdf.jsPDF) throw new Error("PDF library (jspdf.umd.min.js) isn't loaded.");
  const { jsPDF } = window.jspdf;
  const prepareImage = opts.prepareImage || ((d) => prepareImageForPdf(d, 500, 0.75)); // thumbnails: small is fine
  const generatedAt = opts.generatedAt || new Date();

  const doc = new jsPDF({ unit: "pt", format: "letter" });
  const pageW = doc.internal.pageSize.getWidth();
  const pageH = doc.internal.pageSize.getHeight();
  const M = 42;
  const bottom = pageH - 56;
  const INK = [28, 37, 48];
  const SOFT = [107, 114, 128];
  const AMBER = [217, 119, 6];
  const PHOTO_BOX = 92;
  const BOX = 12; // tick box size
  const textX0 = M + BOX + 10; // text starts here when there's no photo
  let y = M;

  const ensure = (h) => {
    if (y + h > bottom) {
      doc.addPage();
      y = M;
    }
  };
  const color = (c) => doc.setTextColor(c[0], c[1], c[2]);
  const font = (style, size) => {
    doc.setFont("helvetica", style);
    doc.setFontSize(size);
  };
  const rule = (c, width) => {
    doc.setDrawColor(c[0], c[1], c[2]);
    doc.setLineWidth(width);
    doc.line(M, y, pageW - M, y);
  };

  let dateText;
  try {
    dateText = generatedAt.toLocaleDateString(undefined, { weekday: "long", year: "numeric", month: "long", day: "numeric" });
  } catch {
    dateText = generatedAt.toDateString();
  }
  dateText = pdfSafe(dateText);
  const total = sections.reduce((n, s) => n + s.items.length, 0);

  font("bold", 10);
  color(AMBER);
  doc.text("MOSS AI FIELD ASSISTANT", M, y + 8);
  y += 28;
  font("bold", 22);
  color(INK);
  doc.text("Materials Shopping List", M, y);
  y += 18;
  font("normal", 10);
  color(SOFT);
  doc.text(`${dateText}  -  ${total} item${total === 1 ? "" : "s"}`, M, y);
  y += 12;
  rule(AMBER, 1.5);
  y += 22;

  for (const section of sections) {
    ensure(90);
    font("bold", 16);
    color(INK);
    const nameLines = doc.splitTextToSize(pdfSafe(section.projectName), pageW - M * 2);
    doc.text(nameLines, M, y + 12);
    y += nameLines.length * 19 + 2;
    if (section.address) {
      font("normal", 10);
      color(SOFT);
      const addr = doc.splitTextToSize(pdfSafe(section.address), pageW - M * 2);
      doc.text(addr, M, y + 8);
      y += addr.length * 12 + 2;
    }
    font("normal", 9);
    color(SOFT);
    doc.text(`${section.items.length} item${section.items.length === 1 ? "" : "s"} to buy`, M, y + 8);
    y += 20;

    for (let i = 0; i < section.items.length; i++) {
      const item = section.items[i];
      const photo = item.photoDataUrl ? await prepareImage(item.photoDataUrl) : null;
      let pw = 0;
      let ph = 0;
      if (photo) {
        const scale = Math.min(PHOTO_BOX / photo.width, PHOTO_BOX / photo.height);
        pw = photo.width * scale;
        ph = photo.height * scale;
      }
      const textX = photo ? textX0 + PHOTO_BOX + 12 : textX0;
      const textW = pageW - M - textX;

      // Measure the text column first so the whole row can be kept on one page.
      font("bold", 12);
      const nameL = doc.splitTextToSize(pdfSafe(item.name), textW);
      font("normal", 10);
      const dimL = item.dimensions ? doc.splitTextToSize(pdfSafe(item.dimensions), textW) : [];
      const qtyText = pdfSafe(
        [item.quantity ? `Qty: ${item.quantity}` : "", item.status && item.status !== "To buy" ? item.status : ""].filter(Boolean).join("   -   ")
      );
      let note = pdfSafe(item.note).trim();
      if (note.length > 1200) note = note.slice(0, 1200).trimEnd() + "...";
      font("normal", 9.5);
      const noteL = note ? doc.splitTextToSize(note, textW) : [];
      let trText = pdfSafe(item.translation).trim();
      if (trText.length > 1200) trText = trText.slice(0, 1200).trimEnd() + "...";
      font("italic", 9.5);
      const trL = trText ? doc.splitTextToSize(trText, textW) : [];
      font("normal", 8.5);
      const stampL = item.stamp ? doc.splitTextToSize(pdfSafe(item.stamp), textW) : [];

      const textH =
        nameL.length * 14.5 + dimL.length * 12.5 + (qtyText ? 13 : 0) + noteL.length * 12 + (trL.length ? 12 + trL.length * 12 : 0) + stampL.length * 10.5 + 6;
      const rowH = Math.max(textH, ph, BOX + 4);
      ensure(rowH + 14);

      // tick box, aligned with the first line of the name
      doc.setDrawColor(INK[0], INK[1], INK[2]);
      doc.setLineWidth(1);
      doc.rect(M, y + 2, BOX, BOX);

      if (photo) {
        doc.addImage(photo.dataUrl, "JPEG", M + BOX + 10, y, pw, ph);
        doc.setDrawColor(220, 224, 229);
        doc.setLineWidth(0.5);
        doc.rect(M + BOX + 10, y, pw, ph);
      }

      let ty = y;
      font("bold", 12);
      color(INK);
      doc.text(nameL, textX, ty + 12);
      ty += nameL.length * 14.5;
      if (dimL.length) {
        font("normal", 10);
        color(INK);
        doc.text(dimL, textX, ty + 9);
        ty += dimL.length * 12.5;
      }
      if (qtyText) {
        font("bold", 10);
        color(AMBER);
        doc.text(qtyText, textX, ty + 9);
        ty += 13;
      }
      if (noteL.length) {
        font("normal", 9.5);
        color(INK);
        doc.text(noteL, textX, ty + 9);
        ty += noteL.length * 12;
      }
      if (trL.length) {
        font("bold", 8.5);
        color(SOFT);
        doc.text(pdfSafe(item.translationLabel || "Translation"), textX, ty + 9);
        ty += 12;
        font("italic", 9.5);
        color(INK);
        doc.text(trL, textX, ty + 9);
        ty += trL.length * 12;
      }
      if (stampL.length) {
        font("normal", 8.5);
        color(SOFT);
        doc.text(stampL, textX, ty + 9);
      }

      y += rowH + 10;
      if (i < section.items.length - 1 && y + 16 <= bottom) {
        rule([229, 231, 235], 0.5);
        y += 12;
      }
    }
    y += 16;
  }

  const pages = doc.getNumberOfPages();
  for (let p = 1; p <= pages; p++) {
    doc.setPage(p);
    font("normal", 8);
    color(SOFT);
    doc.text(`Moss AI Field Assistant  -  ${dateText}`, M, pageH - 28);
    doc.text(`Page ${p} of ${pages}`, pageW - M, pageH - 28, { align: "right" });
  }
  return doc.output("blob");
}


// ---------- Inspection report ----------
// sections = [{ projectName, address, inspections: [{ title, result, stamp, inspector, note,
//   translation, translationLabel, photos: [dataUrl, ...] }] }]
// One block per inspection: title, date/inspector, a colored RESULT chip, a
// photo grid, the written note and (if there is one) its translation.
async function buildInspectionsPdf(sections, opts = {}) {
  if (!window.jspdf || !window.jspdf.jsPDF) throw new Error("PDF library (jspdf.umd.min.js) isn't loaded.");
  const { jsPDF } = window.jspdf;
  const prepareImage = opts.prepareImage || ((d) => prepareImageForPdf(d, 800, 0.75));
  const generatedAt = opts.generatedAt || new Date();

  const doc = new jsPDF({ unit: "pt", format: "letter" });
  const pageW = doc.internal.pageSize.getWidth();
  const pageH = doc.internal.pageSize.getHeight();
  const M = 42;
  const contentW = pageW - M * 2;
  const bottom = pageH - 56;
  const INK = [28, 37, 48];
  const SOFT = [107, 114, 128];
  const AMBER = [217, 119, 6];
  const GREEN = [22, 128, 70];
  const RED = [200, 40, 40];
  let y = M;

  const ensure = (h) => {
    if (y + h > bottom) {
      doc.addPage();
      y = M;
    }
  };
  const color = (c) => doc.setTextColor(c[0], c[1], c[2]);
  const font = (style, size) => {
    doc.setFont("helvetica", style);
    doc.setFontSize(size);
  };
  const rule = (c, width) => {
    doc.setDrawColor(c[0], c[1], c[2]);
    doc.setLineWidth(width);
    doc.line(M, y, pageW - M, y);
  };

  let dateText;
  try {
    dateText = generatedAt.toLocaleDateString(undefined, { weekday: "long", year: "numeric", month: "long", day: "numeric" });
  } catch {
    dateText = generatedAt.toDateString();
  }
  dateText = pdfSafe(dateText);
  const total = sections.reduce((n, s) => n + s.inspections.length, 0);

  font("bold", 10);
  color(AMBER);
  doc.text("MOSS AI FIELD ASSISTANT", M, y + 8);
  y += 28;
  font("bold", 22);
  color(INK);
  doc.text("Inspection Report", M, y);
  y += 18;
  font("normal", 10);
  color(SOFT);
  doc.text(`${dateText}  -  ${total} inspection${total === 1 ? "" : "s"}`, M, y);
  y += 12;
  rule(AMBER, 1.5);
  y += 22;

  const chipFor = (result) => {
    const r = String(result || "");
    if (/^pass/i.test(r)) return { text: "PASSED", fill: GREEN };
    if (/^fail/i.test(r)) return { text: "FAILED / CORRECTIONS REQUIRED", fill: RED };
    if (/^sched/i.test(r)) return { text: "SCHEDULED", fill: AMBER };
    return r ? { text: pdfSafe(r).toUpperCase(), fill: SOFT } : null;
  };

  for (const section of sections) {
    ensure(70);
    font("bold", 16);
    color(INK);
    const nameLines = doc.splitTextToSize(pdfSafe(section.projectName), contentW);
    doc.text(nameLines, M, y + 12);
    y += nameLines.length * 19 + 2;
    if (section.address) {
      font("normal", 10);
      color(SOFT);
      const addrLines = doc.splitTextToSize(pdfSafe(section.address), contentW);
      doc.text(addrLines, M, y + 8);
      y += addrLines.length * 12 + 2;
    }
    font("normal", 9);
    color(SOFT);
    doc.text(`${section.inspections.length} inspection${section.inspections.length === 1 ? "" : "s"}`, M, y + 8);
    y += 18;

    for (let i = 0; i < section.inspections.length; i++) {
      const insp = section.inspections[i];

      font("bold", 12.5);
      const titleLines = doc.splitTextToSize(`${i + 1}. ${pdfSafe(insp.title)}`, contentW);
      const metaText = pdfSafe([insp.stamp, insp.inspector ? `Inspector: ${insp.inspector}` : ""].filter(Boolean).join("  -  "));
      font("normal", 9);
      const metaLines = metaText ? doc.splitTextToSize(metaText, contentW) : [];
      // Keep the title, details and result chip together on one page.
      ensure(titleLines.length * 15 + metaLines.length * 11 + 40);

      font("bold", 12.5);
      color(INK);
      doc.text(titleLines, M, y + 11);
      y += titleLines.length * 15 + 1;
      if (metaLines.length) {
        font("normal", 9);
        color(SOFT);
        doc.text(metaLines, M, y + 8);
        y += metaLines.length * 11 + 4;
      }
      const chip = chipFor(insp.result);
      if (chip) {
        font("bold", 9);
        const w = doc.getTextWidth(chip.text) + 16;
        doc.setFillColor(chip.fill[0], chip.fill[1], chip.fill[2]);
        doc.roundedRect(M, y + 2, w, 16, 3, 3, "F");
        doc.setTextColor(255, 255, 255);
        doc.text(chip.text, M + 8, y + 13);
        y += 26;
      }

      // Photo grid, two per row.
      const photos = [];
      for (const p of insp.photos || []) {
        const prepared = p ? await prepareImage(p) : null;
        if (prepared) photos.push(prepared);
      }
      const gap = 10;
      const cellW = (contentW - gap) / 2;
      const cellMaxH = 180;
      for (let k = 0; k < photos.length; k += 2) {
        const row = photos.slice(k, k + 2).map((ph) => {
          const scale = Math.min(cellW / ph.width, cellMaxH / ph.height);
          return { ph, w: ph.width * scale, h: ph.height * scale };
        });
        const rowH = Math.max(...row.map((c) => c.h));
        ensure(rowH + 8);
        row.forEach((c, idx) => {
          const x = M + idx * (cellW + gap);
          doc.addImage(c.ph.dataUrl, "JPEG", x, y, c.w, c.h);
          doc.setDrawColor(220, 224, 229);
          doc.setLineWidth(0.5);
          doc.rect(x, y, c.w, c.h);
        });
        y += rowH + 8;
      }
      if (photos.length) y += 2;

      const note = pdfSafe(insp.note).trim();
      if (note) {
        font("normal", 10.5);
        color(INK);
        for (const line of doc.splitTextToSize(note, contentW)) {
          ensure(14);
          doc.text(line, M, y + 10);
          y += 14;
        }
      } else {
        font("italic", 10);
        color(SOFT);
        ensure(14);
        doc.text("(No written note)", M, y + 10);
        y += 14;
      }

      const tr = pdfSafe(insp.translation).trim();
      if (tr) {
        y += 4;
        font("bold", 9);
        color(SOFT);
        ensure(13);
        doc.text(pdfSafe(insp.translationLabel || "Translation"), M, y + 9);
        y += 13;
        font("italic", 10.5);
        color(INK);
        for (const line of doc.splitTextToSize(tr, contentW)) {
          ensure(14);
          doc.text(line, M, y + 10);
          y += 14;
        }
      }

      y += 10;
      if (i < section.inspections.length - 1 && y + 20 <= bottom) {
        rule([229, 231, 235], 0.5);
        y += 16;
      }
    }
    y += 14;
  }

  const pages = doc.getNumberOfPages();
  for (let p = 1; p <= pages; p++) {
    doc.setPage(p);
    font("normal", 8);
    color(SOFT);
    doc.text(`Moss AI Field Assistant  -  ${dateText}`, M, pageH - 28);
    doc.text(`Page ${p} of ${pages}`, pageW - M, pageH - 28, { align: "right" });
  }

  return doc.output("blob");
}
