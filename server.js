const express = require("express");
const cors = require("cors");
const fs = require("fs");
const path = require("path");
const { execFile } = require("child_process");
const ExcelJS = require("exceljs");
const PDFDocument = require("pdfkit");
const Tesseract = require("tesseract.js");

const app = express();
app.use(cors());
app.use(express.json({ limit: "12mb" }));
app.use(express.static(__dirname));

const LOG_FILE = path.join(__dirname, "events.jsonl");
const OUT_DIR = path.join(__dirname, "output");
const FONT = path.join(__dirname, "fonts", "NotoSansGujarati-Regular.ttf");
if (!fs.existsSync(OUT_DIR)) fs.mkdirSync(OUT_DIR);

function logEvent(kind, payload, req) {
  const rec = {
    ts: new Date().toISOString(),
    kind,
    ip: req && req.ip,
    ua: req && req.headers["user-agent"],
    payload: payload || {},
  };
  fs.appendFileSync(LOG_FILE, JSON.stringify(rec) + "\n");
}

function setVal(ws, addr, v) {
  if (!ws) return;
  ws.getCell(addr).value = v;
}

function cellText(cell) {
  const v = cell.value;
  if (v == null || v === "") return "";
  if (typeof v === "object") {
    if (v.result != null && v.result !== "") return String(v.result);
    if (Array.isArray(v.richText)) return v.richText.map((t) => t.text).join("");
    if (v.text) return String(v.text);
    if (v.hyperlink) return String(v.text || v.hyperlink);
    return "";
  }
  return String(v);
}

function colLetterToNum(letter) {
  let n = 0;
  for (let i = 0; i < letter.length; i++) n = n * 26 + (letter.charCodeAt(i) - 64);
  return n;
}

const PRINT_AREA = {
  "FACE SHEET": "A1:I40",
  Abstract: "A1:F32",
  Measurement: "A1:L32",
  RA: "A1:I39",
  Lead: "A1:H40",
  Schedule: "A1:I20",
};

const PRINT_AREA_PAVER = {
  Estimate: "A1:I40",
  Abstract: "A1:G29",
  Measurement: "A1:L33",
  Lead: "A1:I54",
};

function applyOnePage(wb, areas) {
  const spec = areas || PRINT_AREA;
  Object.keys(spec).forEach((name) => {
    const ws = wb.getWorksheet(name);
    if (!ws) return;
    const area = spec[name];
    const m = area.match(/^([A-Z]+)(\d+):([A-Z]+)(\d+)$/);
    if (!m) return;
    const lastC = colLetterToNum(m[3]);
    const lastR = Number(m[4]);

    ws.pageSetup.paperSize = 9;
    ws.pageSetup.orientation = "portrait";
    ws.pageSetup.fitToPage = true;
    ws.pageSetup.fitToWidth = 1;
    ws.pageSetup.fitToHeight = 1;
    ws.pageSetup.horizontalCentered = true;
    ws.pageSetup.horizontalDpi = 300;
    ws.pageSetup.verticalDpi = 300;
    ws.pageSetup.margins = {
      left: 0.3,
      right: 0.3,
      top: 0.35,
      bottom: 0.3,
      header: 0.12,
      footer: 0.12,
    };
    ws.pageSetup.printArea = area;

    for (let c = lastC + 1; c <= 80; c++) ws.getColumn(c).hidden = true;
    const rowMax = Math.max(lastR + 50, ws.rowCount || lastR);
    for (let r = lastR + 1; r <= rowMax; r++) ws.getRow(r).hidden = true;
  });
}

async function patchFitXml(xlsxPath) {
  let JSZip;
  try {
    JSZip = require("jszip");
  } catch (_e) {
    return;
  }
  const zip = await JSZip.loadAsync(fs.readFileSync(xlsxPath));
  const files = Object.keys(zip.files).filter(
    (n) => n.startsWith("xl/worksheets/sheet") && n.endsWith(".xml")
  );
  for (const name of files) {
    let xml = await zip.file(name).async("string");
    if (!/pageSetUpPr/.test(xml)) {
      if (/<sheetPr\b[^>]*\/>/.test(xml)) {
        xml = xml.replace(/<sheetPr\b([^>]*)\/>/, '<sheetPr$1><pageSetUpPr fitToPage="1"/></sheetPr>');
      } else if (/<sheetPr\b/.test(xml)) {
        xml = xml.replace(/<sheetPr\b([^>]*)>/, '<sheetPr$1><pageSetUpPr fitToPage="1"/>');
      } else {
        xml = xml.replace(/<dimension /, '<sheetPr><pageSetUpPr fitToPage="1"/></sheetPr><dimension ');
      }
    } else {
      xml = xml.replace(/<pageSetUpPr[^/]*\/>/, '<pageSetUpPr fitToPage="1"/>');
    }
    xml = xml.replace(/<pageSetup\b([^>]*)\/>/, (_all, attrs) => {
      let a = String(attrs)
        .replace(/\s+scale="[^"]*"/g, "")
        .replace(/\s+horizontalDpi="[^"]*"/g, "")
        .replace(/\s+verticalDpi="[^"]*"/g, "")
        .replace(/\s+fitToWidth="[^"]*"/g, "")
        .replace(/\s+fitToHeight="[^"]*"/g, "")
        .replace(/\s+paperSize="[^"]*"/g, "");
      return `<pageSetup${a} paperSize="9" fitToWidth="1" fitToHeight="1" horizontalDpi="300" verticalDpi="300"/>`;
    });
    xml = xml.replace(/<pageMargins[^/]*\/>/, '<pageMargins left="0.5" right="0.5" top="0.5" bottom="0.5" header="0.25" footer="0.25"/>');
    if (!/<pageMargins /.test(xml)) {
      xml = xml.replace(/<pageSetup /, '<pageMargins left="0.5" right="0.5" top="0.5" bottom="0.5" header="0.25" footer="0.25"/><pageSetup ');
    }
    zip.file(name, xml);
  }
  const out = await zip.generateAsync({ type: "nodebuffer", compression: "DEFLATE" });
  fs.writeFileSync(xlsxPath, out);
}

function sofficeBin() {
  const list = ["soffice", "libreoffice", "/usr/bin/soffice", "/usr/bin/libreoffice"];
  for (const b of list) {
    try {
      if (b.startsWith("/") && fs.existsSync(b)) return b;
    } catch (_e) {}
  }
  return "soffice";
}

function convertWithSoffice(xlsxPath) {
  return new Promise((resolve, reject) => {
    const dir = path.dirname(xlsxPath);
    const pdfPath = xlsxPath.replace(/\.xlsx$/i, ".pdf");
    const filter =
      'pdf:calc_pdf_Export:{"SinglePageSheets":{"type":"boolean","value":"true"}}';
    execFile(
      sofficeBin(),
      ["--headless", "--norestore", "--nolockcheck", "--convert-to", filter, "--outdir", dir, xlsxPath],
      { timeout: 120000 },
      (err) => {
        if (err) {
          execFile(
            sofficeBin(),
            ["--headless", "--norestore", "--nolockcheck", "--convert-to", "pdf", "--outdir", dir, xlsxPath],
            { timeout: 120000 },
            (err2) => {
              if (err2) return reject(err2);
              if (!fs.existsSync(pdfPath)) return reject(new Error("pdf missing"));
              resolve(pdfPath);
            }
          );
          return;
        }
        if (!fs.existsSync(pdfPath)) return reject(new Error("pdf missing"));
        resolve(pdfPath);
      }
    );
  });
}

async function addPdfMargins(pdfPath) {
  let PDFLib;
  try {
    PDFLib = require("pdf-lib");
  } catch (_e) {
    return;
  }
  const { PDFDocument } = PDFLib;
  const src = await PDFDocument.load(fs.readFileSync(pdfPath));
  const out = await PDFDocument.create();
  const margin = 36;
  const srcPages = src.getPages();
  for (let i = 0; i < srcPages.length; i++) {
    const sp = srcPages[i];
    const { width, height } = sp.getSize();
    const [emb] = await out.embedPages([sp]);
    const page = out.addPage([width, height]);
    page.drawPage(emb, {
      x: margin,
      y: margin,
      width: width - margin * 2,
      height: height - margin * 2,
    });
  }
  fs.writeFileSync(pdfPath, await out.save());
}

function sheetsToPdf(wb, pdfPath) {
  return new Promise((resolve, reject) => {
    const doc = new PDFDocument({ autoFirstPage: false, margin: 24 });
    const stream = fs.createWriteStream(pdfPath);
    doc.pipe(stream);
    stream.on("finish", () => resolve(pdfPath));
    stream.on("error", reject);

    const hasFont = fs.existsSync(FONT);
    if (hasFont) {
      try {
        doc.registerFont("Gu", FONT);
      } catch (_e) {}
    }

    wb.worksheets.forEach((ws, idx) => {
      const landscape = (ws.pageSetup && ws.pageSetup.orientation) === "landscape";
      doc.addPage({ size: "A4", layout: landscape ? "landscape" : "portrait", margin: 22 });
      const pageW = doc.page.width - 44;
      const pageH = doc.page.height - 50;

      let maxR = 0;
      let maxC = 0;
      const grid = [];
      ws.eachRow({ includeEmpty: false }, (row, r) => {
        if (r > maxR) maxR = r;
        row.eachCell({ includeEmpty: false }, (cell, c) => {
          if (c > maxC) maxC = c;
          if (!grid[r]) grid[r] = [];
          grid[r][c] = cellText(cell);
        });
      });
      if (maxR < 1) maxR = 1;
      if (maxC < 1) maxC = 1;

      doc.font(hasFont ? "Gu" : "Helvetica").fontSize(9).fillColor("#1B5D45");
      doc.text(ws.name, 22, 12, { width: pageW, align: "left" });

      const top = 28;
      const fontSize = Math.max(5, Math.min(8, (pageH - 8) / Math.max(maxR, 1) - 1.2));
      const rowH = Math.min(14, (pageH - 8) / maxR);
      const colW = pageW / maxC;

      doc.fontSize(fontSize).fillColor("#14211A");
      for (let r = 1; r <= maxR; r++) {
        const y = top + (r - 1) * rowH;
        if (y > top + pageH - 6) break;
        for (let c = 1; c <= maxC; c++) {
          const t = (grid[r] && grid[r][c]) || "";
          if (!t) continue;
          const x = 22 + (c - 1) * colW;
          doc.text(t, x, y, { width: colW - 2, height: rowH - 0.5, ellipsis: true, lineBreak: false });
        }
      }
      if (idx === wb.worksheets.length - 1) {
        /* last */
      }
    });
    doc.end();
  });
}

async function writeAndRespond(req, res, wb, d, prefix, areas, kind) {
  if (wb.calcProperties) wb.calcProperties.fullCalcOnLoad = true;
  applyOnePage(wb, areas);
  const output = d.output || "xlsx";
  const safe = String(d.village || "gam").replace(/[^a-zA-Z0-9._-]+/g, "_");
  const stamp = Date.now();
  const xlsxName = `${prefix}_${safe}_${stamp}.xlsx`;
  const pdfName = `${prefix}_${safe}_${stamp}.pdf`;
  const xlsxFull = path.join(OUT_DIR, xlsxName);
  const pdfFull = path.join(OUT_DIR, pdfName);
  await wb.xlsx.writeFile(xlsxFull);
  await patchFitXml(xlsxFull);

  let pdfUrl = null;
  let pdfError = null;
  if (output === "pdf" || output === "both") {
    try {
      await convertWithSoffice(xlsxFull);
      const produced = xlsxFull.replace(/\.xlsx$/i, ".pdf");
      if (produced !== pdfFull && fs.existsSync(produced)) fs.copyFileSync(produced, pdfFull);
      if (!fs.existsSync(pdfFull) && fs.existsSync(produced)) fs.copyFileSync(produced, pdfFull);
      if (!fs.existsSync(pdfFull)) throw new Error("pdf missing");
      await addPdfMargins(pdfFull);
      pdfUrl = `/api/download/${pdfName}`;
    } catch (e1) {
      pdfError =
        "PDF LibreOffice vagar nathi. Render Settings ma Runtime = Docker karo (Dockerfile repo ma che).";
      logEvent("pdf_fail", { error: String(e1.message || e1), kind }, req);
    }
  }

  logEvent(kind, { village: d.village, output, xlsx: xlsxName, pdf: pdfUrl }, req);
  const wantXlsx = output === "xlsx" || output === "both";
  res.json({
    ok: true,
    xlsx: wantXlsx ? `/api/download/${xlsxName}` : null,
    pdf: pdfUrl,
    pdf_error: pdfError,
  });
}

function detectTypeFromText(t) {
  const s = String(t || "").toLowerCase().replace(/\s+/g, " ");
  if (/ગટર|ગટ્ટર|gutter|સેનિટેશન/.test(s)) return "gutter";
  if (/પાઇપ|પાઈપ|pipe line|pipeline|hume/.test(s)) return "pipe";
  if (/પેવર|पेवर|paver|interlock|ઇન્ટરલોક|બ્લોકનું|બ્લોક નું/.test(s)) return "paver";
  if (/સી\s*સી|સીસી|સી\.?\s*સી|सी\s*सी|\bcc\b|सीसी|સી સી રોડ|સીસી રોડ|cement concrete/.test(s)) return "cc";
  if (/બોર|bore|પમ્પ|પંપ|મશીનરી/.test(s)) return "unknown";
  return "unknown";
}

function detectVillage(t) {
  const m = String(t || "").match(/([^\s,]{2,20})\s*ગામે/);
  if (m) return m[1];
  const m2 = String(t || "").match(/ગામ(?:નું નામ)?\s*[:\-–]?\s*([^\s,]{2,20})/);
  return m2 ? m2[1] : "";
}

function detectWorks(t) {
  const g = "૦૧૨૩૪૫૬૭૮૯";
  const raw = String(t || "").replace(/[૦-૯]/g, (ch) => String(g.indexOf(ch)));
  const parts = raw.split(/(?=(?:^|\n)\s*(?:[1-9]|1[0-9]|2[0-2])\s+)/);
  let works = [];
  parts.forEach((p) => {
    const sm = p.match(/^\s*(\d{1,2})\s+/);
    const sr = sm ? Number(sm[1]) : 0;
    if (sr < 1 || sr > 22) return;
    const nums = String(p).replace(/,/g, "").match(/\d{5,8}/g) || [];
    const amts = nums.map(Number).filter((n) => n >= 25000 && n <= 2000000);
    const amt = amts.length ? amts[amts.length - 1] : 0;
    works.push({
      sr,
      work_name: p.replace(/\s+/g, " ").trim().slice(0, 120),
      amounting: amt,
      type: detectTypeFromText(p),
      village: detectVillage(p)
    });
  });
  if (works.length < 2) {
    String(raw).split(/\n+/).forEach((line) => {
      const ln = line.replace(/,/g, " ").replace(/\s+/g, " ").trim();
      const m = ln.match(/(\d{5,8})\s*$/);
      if (!m) return;
      const amt = Number(m[1]);
      if (amt < 25000 || amt > 2000000) return;
      if (/કુલ/.test(ln) && !/ગામે|ફળીયા/.test(ln)) return;
      works.push({
        sr: works.length + 1,
        work_name: ln.replace(m[1], "").trim(),
        amounting: amt,
        type: detectTypeFromText(ln),
        village: detectVillage(ln)
      });
    });
  }
  let max = 0;
  works.forEach((w) => { if (w.amounting > max) max = w.amounting; });
  if (max >= 800000) works = works.filter((w) => w.amounting !== max);
  return works.slice(0, 25);
}

function detectAmount(t) {
  const nums = String(t || "").replace(/,/g, "").match(/\d{4,9}/g) || [];
  const vals = nums.map(Number).filter((n) => n >= 10000 && n <= 99999999);
  if (!vals.length) return 0;
  return Math.max.apply(null, vals);
}

function detectYear(t) {
  const m = String(t || "").match(/20\d{2}\s*[-–]\s*\d{2,4}/);
  return m ? m[0].replace(/\s/g, "") : "";
}

async function geminiWorks(b64, mime) {
  const key = process.env.GEMINI_API_KEY;
  if (!key) return null;
  const body = {
    contents: [{
      parts: [
        { text: "This is a Gujarati Panchayat parishisht / SS sanction table. Extract EVERY work row. Ignore header and TOTAL (કુલ) rows. Reply ONLY JSON array: [{\"sr\":1,\"type\":\"cc|paver|gutter|pipe|bore|unknown\",\"work_name\":\"gujarati name\",\"amounting\":100000,\"village\":\"\",\"taluka\":\"\"}]. type: સીસી/સી સી/CC=cc, પેવર=paver, ગટર=gutter, પાઇપ=pipe, બોર/પમ્પ=bore." },
        { inline_data: { mime_type: mime || "image/jpeg", data: b64 } }
      ]
    }]
  };
  const url = "https://generativelanguage.googleapis.com/v1beta/models/gemini-2.0-flash:generateContent?key=" + encodeURIComponent(key);
  const r = await fetch(url, { method: "POST", headers: { "Content-Type": "application/json" }, body: JSON.stringify(body) });
  const j = await r.json();
  const txt = (((j.candidates || [])[0] || {}).content || {}).parts
    ? j.candidates[0].content.parts.map((p) => p.text || "").join("\n")
    : "";
  const m = String(txt).match(/\[[\s\S]*\]/);
  if (!m) return { works: [], raw: txt.slice(0, 1500) };
  let works = [];
  try { works = JSON.parse(m[0]); } catch (e) { works = []; }
  works = (works || []).map((w, i) => ({
    sr: Number(w.sr || i + 1),
    type: detectTypeFromText(String(w.type || "") + " " + String(w.work_name || "")) || w.type || "unknown",
    work_name: String(w.work_name || ""),
    amounting: Number(String(w.amounting || "0").replace(/[^\d]/g, "")) || 0,
    village: String(w.village || detectVillage(w.work_name || "")),
    taluka: String(w.taluka || "")
  })).filter((w) => w.work_name || w.amounting);
  return { works, raw: txt.slice(0, 1500) };
}

app.post("/api/scan", (req, res) => {
  const img = req.body && req.body.image;
  if (!img || typeof img !== "string") return res.json({ ok: false, error: "no image" });
  const m = img.match(/^data:image\/\w+;base64,(.+)$/);
  const b64 = m ? m[1] : img;
  const id = "scan-" + Date.now();
  const ext = (img.match(/^data:image\/([\w+]+);/) || [])[1] === "png" ? ".png" : ".jpg";
  const tmp = path.join(OUT_DIR, id + ext);
  const outBase = path.join(OUT_DIR, id + "-out");
  try {
    fs.writeFileSync(tmp, Buffer.from(b64, "base64"));
  } catch (e) {
    return res.json({ ok: false, error: "save fail" });
  }
  const mime = (img.match(/^data:(image\/[\w+]+);/) || [])[1] || "image/jpeg";
  geminiWorks(b64, mime).then((g) => {
    if (g && g.works && g.works.length) {
      try { fs.unlinkSync(tmp); } catch (e) {}
      logEvent("scan_gemini", { n: g.works.length }, req);
      return res.json({
        ok: true,
        type: g.works[0].type,
        amounting: g.works[0].amounting,
        year: "",
        work_name: g.works[0].work_name,
        works: g.works,
        raw: g.raw,
        ocr: true,
        error: ""
      });
    }
    runTess();
  }).catch(() => runTess());
  function runTess() {
  function readOut() {
    try { return fs.readFileSync(outBase + ".txt", "utf8"); } catch (e) { return ""; }
  }
  function cleanup() {
    try { fs.unlinkSync(tmp); } catch (e) {}
    try { fs.unlinkSync(outBase + ".txt"); } catch (e) {}
  }
  function finish(err, text) {
    cleanup();
    const type = detectTypeFromText(text);
    const works = detectWorks(text);
    logEvent("scan_ocr", { type, chars: String(text||"").length, n: works.length, err: err ? String(err.message || err) : "" }, req);
    res.json({
      ok: true,
      type: works[0] ? works[0].type : type,
      amounting: works[0] ? works[0].amounting : detectAmount(text),
      year: detectYear(text),
      work_name: works[0] ? works[0].work_name : "",
      works: works,
      raw: String(text || "").slice(0, 1500),
      ocr: !err && !!String(text || "").trim(),
      error: err ? String(err.message || err).slice(0, 180) : ""
    });
  }
  function ocrLang(lang) {
    return Tesseract.recognize(tmp, lang, { logger: function () {} }).then(function (out) {
      return (out && out.data && out.data.text) || "";
    });
  }
  ocrLang("guj+eng")
    .then(function (t) {
      if (String(t).trim().length > 15) return t;
      return ocrLang("guj");
    })
    .then(function (t) {
      if (String(t).trim().length > 15) return t;
      return ocrLang("eng");
    })
    .then(function (t) {
      finish(null, t);
    })
    .catch(function (e) {
      finish(e, "");
    });
  }
});

app.get("/api/ocr", (_req, res) => {
  execFile("tesseract", ["--version"], { timeout: 8000 }, (err, stdout, stderr) => {
    res.json({
      ok: !err,
      version: String(stdout || stderr || "").split("\n")[0] || "",
      err: err ? String(err.message || err) : ""
    });
  });
});

app.get("/api/health", (_req, res) => {
  res.json({ ok: true, service: "parastate-mvp" });
});

app.post("/api/log", (req, res) => {
  logEvent(req.body && req.body.kind ? req.body.kind : "client", req.body || {}, req);
  res.json({ ok: true });
});

app.get("/api/logs", (_req, res) => {
  if (!fs.existsSync(LOG_FILE)) return res.json({ count: 0, events: [] });
  const lines = fs.readFileSync(LOG_FILE, "utf8").trim().split("\n").filter(Boolean);
  const events = lines
    .slice(-50)
    .map((l) => JSON.parse(l))
    .reverse();
  res.json({ count: lines.length, events });
});

app.post("/api/estimate/cc", async (req, res) => {
  try {
    const d = req.body || {};
    const wb = new ExcelJS.Workbook();
    await wb.xlsx.readFile(path.join(__dirname, "skeleton-cc.xlsx"));

    const face = wb.getWorksheet("FACE SHEET");
    const meas = wb.getWorksheet("Measurement");
    const lead = wb.getWorksheet("Lead");
    const abs = wb.getWorksheet("Abstract");
    const ra = wb.getWorksheet("RA");
    const sch = wb.getWorksheet("Schedule");
    if (!face || !meas || !lead) {
      throw new Error("skeleton-cc.xlsx sheets missing");
    }

    const L = Number(d.length_m);
    const W = Number(d.width_m);
    const boxT = Number(d.box_thick_m);
    const btT = Number(d.bt_thick_m);
    const voids = Number(d.voids);
    const murPct = Number(d.murrum_pct);
    const ccT = Number(d.cc_thick_m);
    const area = L * W;
    const boxQty = area * boxT;
    const btBase = area * btT;
    const btQty = btBase * (1 + voids);
    const murQty = btQty * (murPct / 100);
    const ccQty = area * ccT;
    const work = d.work_name || "";
    const taluka = d.taluka || "";
    const say = Number(d.amounting || 0);

    const a1 = boxQty * 156.56;
    const a2 = btQty * 677.83;
    const a3 = murQty * 171.8;
    const a4 = btQty * 247.28;
    const a5 = murQty * 146.01;
    const a6r = 0;
    const a6c = ccQty * 4866.35;
    const a7 = 2656;
    const a8 = 306.14;
    const tot = a1 + a2 + a3 + a4 + a5 + a6r + a6c + a7 + a8;
    const gst = tot * 0.18;
    const grand = tot + gst;

    setVal(face, "F3", d.division);
    setVal(face, "G3", d.jilla);
    setVal(face, "F5", d.subdiv_address);
    setVal(face, "I5", d.nani_address);
    setVal(face, "D9", d.fund_head);
    setVal(face, "F35", d.fund_head);
    setVal(face, "E35", "મોજે");
    setVal(face, "H19", taluka);
    setVal(face, "H39", "");
    setVal(face, "F40", taluka);
    setVal(face, "C21", work);
    setVal(face, "G22", say);
    setVal(face, "D28", d.prepared_by);
    setVal(face, "B34", d.sr_no);
    setVal(face, "C34", d.ss_details);
    setVal(face, "B35", d.village);

    setVal(meas, "C1", work);
    setVal(meas, "C3", L);
    setVal(meas, "C4", W);
    setVal(meas, "I3", L);
    setVal(meas, "J3", W);
    setVal(meas, "K3", area);
    setVal(meas, "K4", area);
    setVal(meas, "G6", area);
    setVal(meas, "I6", boxT);
    setVal(meas, "K6", boxQty);
    setVal(meas, "G9", area);
    setVal(meas, "I9", btT);
    setVal(meas, "K9", btBase);
    setVal(meas, "E10", btBase);
    setVal(meas, "G10", voids);
    setVal(meas, "K10", btBase * voids);
    setVal(meas, "K11", btQty);
    setVal(meas, "I14", murPct);
    setVal(meas, "K14", murQty);
    setVal(meas, "A16", "Item No. :- 4 Spreading the stone aggregates for soiling and W. B. M. including filling the inter stices forming the surface to required camber and gradient (excluding spreading of blindage) (ii) 40 mm to 63 mm size aggreates (HB)");
    setVal(meas, "K22", area);
    setVal(meas, "G26", area);
    setVal(meas, "I26", ccT);
    setVal(meas, "K26", ccQty);

    if (abs) {
      setVal(abs, "B2", work);
      setVal(abs, "A4", boxQty);
      setVal(abs, "F4", a1);
      setVal(abs, "A6", btQty);
      setVal(abs, "F6", a2);
      setVal(abs, "A8", murQty);
      setVal(abs, "F8", a3);
      setVal(abs, "C10", "Item No. :- 4 Spreading the stone aggregates for soiling and W. B. M. including filling the inter stices forming the surface to required camber and gradient (excluding spreading of blindage) (ii) 40 mm to 63 mm size aggreates (HB)");
      setVal(abs, "F10", a4);
      setVal(abs, "A12", murQty);
      setVal(abs, "F12", a5);
      setVal(abs, "A14", 0);
      setVal(abs, "F14", a6r);
      setVal(abs, "F16", a6c);
      setVal(abs, "F22", tot);
      setVal(abs, "F23", gst);
      setVal(abs, "F24", grand);
      setVal(abs, "F25", say);
      abs.getCell("A32").value = taluka;
    }
    if (ra) {
      setVal(ra, "C1", work);
      ra.getCell("D38").value = taluka;
    }
    if (sch) {
      setVal(sch, "C1", work);
      sch.getCell("C18").value = taluka;
    }
    setVal(lead, "B1", work);
    lead.getCell("C5").value = taluka;
    lead.getCell("A6").value = taluka;
    lead.getCell("B38").value = taluka;

    await writeAndRespond(req, res, wb, d, "CC", PRINT_AREA, "estimate_cc");
  } catch (err) {
    logEvent("estimate_cc_error", { error: String(err) }, req);
    res.status(500).json({ ok: false, error: String(err.message || err) });
  }
});

app.post("/api/estimate/paver", async (req, res) => {
  try {
    const d = req.body || {};
    const wb = new ExcelJS.Workbook();
    await wb.xlsx.readFile(path.join(__dirname, "skeleton-paver.xlsx"));

    const face = wb.getWorksheet("Estimate");
    const meas = wb.getWorksheet("Measurement");
    const lead = wb.getWorksheet("Lead");
    if (!face || !meas || !lead) {
      throw new Error("skeleton-paver.xlsx sheets missing");
    }

    setVal(face, "F2", d.jilla);
    setVal(face, "F4", d.subdiv_address);
    setVal(face, "I4", d.nani_address);
    setVal(face, "F5", d.taluka);
    setVal(face, "D8", d.fund_head);
    setVal(face, "C20", d.work_name);
    setVal(face, "G21", Number(d.amounting || 0));
    setVal(face, "D27", d.prepared_by);
    setVal(face, "D29", d.prepared_by);
    setVal(face, "B34", d.sr_no);
    setVal(face, "C34", d.ss_details);
    setVal(face, "H18", d.taluka);
    setVal(face, "G40", d.taluka);

    const L = Number(d.length_m);
    const W = Number(d.width_m);
    const area = L * W;
    const boxQty = area * Number(d.box_thick_m);
    const murQty = area * 0.1;
    const vata = 2 * L + 2 * W;
    const trunc2 = (x) => Math.trunc(x * 100) / 100;
    const brass = area * 10.7584 / 100;
    const f4 = trunc2(156.56 * boxQty);
    const f6 = trunc2(210.42 * murQty);
    const f12 = trunc2(740.51 * area);
    const f14 = trunc2(23.68 * vata);
    const f18 = 608;
    const f20 = 306.14;
    const f22 = f4 + f6 + f12 + f14 + f18 + f20;
    const f23 = f22 * 0.18;
    const f24 = f22 + f23;

    setVal(meas, "C3", L);
    setVal(meas, "C4", W);
    setVal(meas, "D2", d.work_name);
    setVal(meas, "I3", area);
    setVal(meas, "I6", Number(d.box_thick_m));
    setVal(meas, "G6", area);
    setVal(meas, "K6", boxQty);
    setVal(meas, "K7", boxQty);
    setVal(meas, "G10", area);
    setVal(meas, "K10", murQty);
    setVal(meas, "K11", murQty);
    setVal(meas, "K13", murQty);
    setVal(meas, "G16", W);
    setVal(meas, "E16", L);
    setVal(meas, "K16", trunc2(L * W));
    setVal(meas, "K19", area);
    setVal(meas, "K20", area);
    setVal(meas, "M20", area * 10.7584);
    setVal(meas, "N20", brass);
    setVal(meas, "E20", brass);
    setVal(meas, "E22", L);
    setVal(meas, "E23", W);
    setVal(meas, "K22", trunc2(2 * L));
    setVal(meas, "K23", trunc2(2 * W));
    setVal(meas, "K24", vata);
    setVal(meas, "K30", boxQty);

    setVal(lead, "B1", d.work_name);
    setVal(lead, "C5", d.taluka);
    setVal(lead, "A6", d.taluka);
    setVal(lead, "G54", d.taluka);

    const abs = wb.getWorksheet("Abstract");
    if (abs) {
      setVal(abs, "C2", d.work_name);
      setVal(abs, "A4", boxQty);
      setVal(abs, "A6", murQty);
      setVal(abs, "A12", area);
      setVal(abs, "A14", vata);
      setVal(abs, "F4", f4);
      setVal(abs, "F6", f6);
      setVal(abs, "F12", f12);
      setVal(abs, "F14", f14);
      setVal(abs, "F18", f18);
      setVal(abs, "F20", f20);
      setVal(abs, "F22", f22);
      setVal(abs, "F23", f23);
      setVal(abs, "F24", f24);
      setVal(abs, "F25", Number(d.amounting || 0));
      setVal(abs, "A28", d.taluka);
    }

    await writeAndRespond(req, res, wb, d, "PAVER", PRINT_AREA_PAVER, "estimate_paver");
  } catch (err) {
    logEvent("estimate_paver_error", { error: String(err) }, req);
    res.status(500).json({ ok: false, error: String(err.message || err) });
  }
});

app.get("/api/download/:name", (req, res) => {
  const name = path.basename(req.params.name);
  const full = path.join(OUT_DIR, name);
  if (!fs.existsSync(full)) return res.status(404).json({ ok: false });
  res.download(full, name);
});

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => {
  console.log("ParaState MVP on " + PORT);
});
