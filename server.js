const express = require("express");
const cors = require("cors");
const fs = require("fs");
const path = require("path");
const { execFile, execFileSync, spawn } = require("child_process");
const ExcelJS = require("exceljs");
const PDFDocument = require("pdfkit");
const Tesseract = require("tesseract.js");

const app = express();
app.use(cors());
app.use(express.json({ limit: "12mb" }));
app.use(express.static(__dirname));

function istDay(input) {
  const dt = input ? new Date(input) : new Date();
  const t = dt.getTime();
  const base = isNaN(t) ? Date.now() : t;
  return new Date(base + 330 * 60 * 1000).toISOString().slice(0, 10);
}
function recordDay(r) {
  if (r && r.ts) return istDay(r.ts);
  return String((r && r.day) || "").slice(0, 10);
}
const OUT_DIR = path.join(__dirname, "output");
const LOG_FILE = path.join(__dirname, "events.jsonl");
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


function unshareFormulas(wb) {
  (wb.worksheets || []).forEach(function (ws) {
    ws.eachRow(function (row) {
      row.eachCell(function (cell) {
        try {
          if (cell.formula && String(cell.formulaType || "") === "shared") {
            const f = cell.formula;
            cell.value = { formula: f };
          }
        } catch (_e) {}
      });
    });
  });
}

function setVal(ws, addr, v) {
  if (!ws) return;
  ws.getCell(addr).value = v;
}

function findLabelRow(ws, re) {
  let hit = 0;
  if (!ws) return 0;
  ws.eachRow(function (row, n) {
    if (hit) return;
    if (re.test(String(cellText(row.getCell(1)) || ""))) hit = n;
  });
  return hit;
}

function bakeFormulaResults(wb) {
  const PENDING = { pending: true };
  const formulas = [];
  wb.worksheets.forEach(function (ws) {
    if (!ws) return;
    ws.eachRow(function (row) {
      row.eachCell(function (cell) {
        const v = cell.value;
        if (v && typeof v === "object" && v.formula) {
          formulas.push({ ws: ws, addr: cell.address, formula: String(v.formula).replace(/^=/, "") });
        }
      });
    });
  });
  const done = {};
  function findWs(name) {
    const n = String(name || "").replace(/^'|'$/g, "");
    return wb.getWorksheet(n) || wb.worksheets.find(function (w) { return w && String(w.name).toLowerCase() === n.toLowerCase(); });
  }
  function rawOf(ws, addr) {
    const cell = ws.getCell(addr);
    const v = cell.value;
    if (v && typeof v === "object" && v.formula) return undefined;
    if (v == null || v === "") return 0;
    if (typeof v === "number") return v;
    if (typeof v === "string") return v;
    if (typeof v === "object") {
      if (Array.isArray(v.richText)) return v.richText.map(function (t) { return t.text; }).join("");
      if (v.text != null) return v.text;
      if (v.result != null) return v.result;
    }
    return 0;
  }
  function isPend(v) { return !!(v && v.pending); }
  function num(v) {
    if (isPend(v)) return v;
    if (typeof v === "number") return v;
    if (typeof v === "string") {
      const n = Number(String(v).replace(/,/g, ""));
      return isNaN(n) ? 0 : n;
    }
    if (Array.isArray(v)) {
      let s = 0;
      for (let i = 0; i < v.length; i++) {
        const n = num(v[i]);
        if (isPend(n)) return n;
        s += n;
      }
      return s;
    }
    return 0;
  }
  function cellVal(sheet, addr) {
    const ws = findWs(sheet);
    if (!ws) return 0;
    const k = ws.name + "!" + String(addr).toUpperCase();
    if (Object.prototype.hasOwnProperty.call(done, k)) return done[k];
    const raw = rawOf(ws, addr);
    if (raw === undefined) return PENDING;
    return raw;
  }
  function rangeVals(sheet, a, b) {
    const pa = String(a).toUpperCase().replace(/\$/g, "").match(/^([A-Z]+)(\d+)$/);
    const pb = String(b).toUpperCase().replace(/\$/g, "").match(/^([A-Z]+)(\d+)$/);
    if (!pa || !pb) return [0];
    const c1 = colLetterToNum(pa[1]);
    const c2 = colLetterToNum(pb[1]);
    const r1 = Number(pa[2]);
    const r2 = Number(pb[2]);
    const out = [];
    for (let r = Math.min(r1, r2); r <= Math.max(r1, r2); r++) {
      for (let c = Math.min(c1, c2); c <= Math.max(c1, c2); c++) {
        let col = "";
        let n = c;
        while (n > 0) { const m = (n - 1) % 26; col = String.fromCharCode(65 + m) + col; n = Math.floor((n - 1) / 26); }
        const v = cellVal(sheet, col + r);
        if (isPend(v)) return PENDING;
        out.push(v);
      }
    }
    return out;
  }
  function evalFormula(formula, sheet) {
    const s = String(formula || "");
    let i = 0;
    function skip() { while (s[i] === " ") i++; }
    function pendJoin(a, b, op) {
      if (isPend(a) || isPend(b)) return PENDING;
      if (op === "+") return num(a) + num(b);
      if (op === "-") return num(a) - num(b);
      if (op === "*") return num(a) * num(b);
      if (op === "/") { const d = num(b); return d ? num(a) / d : 0; }
      return 0;
    }
    function parseExpr() { return parseAdd(); }
    function parseAdd() {
      let v = parseMul();
      while (true) {
        skip();
        if (s[i] === "+") { i++; v = pendJoin(v, parseMul(), "+"); }
        else if (s[i] === "-") { i++; v = pendJoin(v, parseMul(), "-"); }
        else break;
      }
      return v;
    }
    function parseMul() {
      let v = parseUnary();
      while (true) {
        skip();
        if (s[i] === "*") { i++; v = pendJoin(v, parseUnary(), "*"); }
        else if (s[i] === "/") { i++; v = pendJoin(v, parseUnary(), "/"); }
        else if (s[i] === "%") { i++; v = isPend(v) ? v : num(v) / 100; }
        else break;
      }
      return v;
    }
    function parseUnary() {
      skip();
      if (s[i] === "-") { i++; const v = parseUnary(); return isPend(v) ? v : -num(v); }
      if (s[i] === "+") { i++; return parseUnary(); }
      return parsePrim();
    }
    function parseRef(sheetName) {
      const m = s.slice(i).match(/^(\$?[A-Z]{1,3}\$?\d+)(?::(\$?[A-Z]{1,3}\$?\d+))?/i);
      if (!m) return 0;
      i += m[0].length;
      if (m[2]) {
        const vals = rangeVals(sheetName, m[1], m[2]);
        if (isPend(vals)) return vals;
        return vals.length ? vals[0] : 0;
      }
      return cellVal(sheetName, m[1].replace(/\$/g, ""));
    }
    function parsePrim() {
      skip();
      if (s[i] === "(") { i++; const v = parseExpr(); skip(); if (s[i] === ")") i++; return v; }
      if (s[i] === "'") {
        i++;
        let name = "";
        while (i < s.length && s[i] !== "'") name += s[i++];
        if (s[i] === "'") i++;
        if (s[i] === "!") i++;
        return parseRef(name);
      }
      if (/[0-9.]/.test(s[i] || "")) {
        let j = i;
        while (/[0-9.]/.test(s[j] || "")) j++;
        const n = Number(s.slice(i, j));
        i = j;
        skip();
        if (s[i] === "%") { i++; return n / 100; }
        return n;
      }
      let k = i;
      while (k < s.length && s[k] !== "!" && s[k] !== "(" && "+-*/, )".indexOf(s[k]) < 0) k++;
      if (s[k] === "!") {
        const name = s.slice(i, k);
        i = k + 1;
        return parseRef(name);
      }
      const token = s.slice(i, k);
      const upper = token.toUpperCase();
      if ((upper === "SUM" || upper === "ROUNDUP" || upper === "ROUNDDOWN" || upper === "TRUNC") && s[k] === "(") {
        i = k + 1;
        const args = [];
        while (i < s.length) {
          skip();
          if (s[i] === ")") { i++; break; }
          args.push(parseArg());
          skip();
          if (s[i] === ",") { i++; continue; }
          if (s[i] === ")") i++;
          break;
        }
        if (args.some(isPend)) return PENDING;
        if (upper === "SUM") {
          return args.reduce(function (a, b) {
            if (Array.isArray(b)) return a + b.reduce(function (x, y) { return x + num(y); }, 0);
            return a + num(b);
          }, 0);
        }
        const digits = args.length > 1 ? num(args[1]) : 0;
        const p = Math.pow(10, digits || 0);
        const x = num(args[0]) * p;
        if (upper === "ROUNDUP") return (x < 0 ? Math.floor(x) : Math.ceil(x)) / p;
        return (x < 0 ? Math.ceil(x) : Math.floor(x)) / p;
      }
      if (/^\$?[A-Z]{1,3}\$?\d+$/i.test(token)) {
        i = k;
        return cellVal(sheet, token.replace(/\$/g, ""));
      }
      i = Math.max(k, i + 1);
      return 0;
    }
    function parseArg() {
      skip();
      const m = s.slice(i).match(/^(\$?[A-Z]{1,3}\$?\d+):(\$?[A-Z]{1,3}\$?\d+)/i);
      if (m) {
        i += m[0].length;
        return rangeVals(sheet, m[1], m[2]);
      }
      return parseExpr();
    }
    const val = parseExpr();
    skip();
    if (i < s.length) return PENDING;
    return val;
  }
  for (let pass = 0; pass < formulas.length + 2; pass++) {
    let moved = 0;
    formulas.forEach(function (item) {
      const k = item.ws.name + "!" + item.addr.toUpperCase();
      if (Object.prototype.hasOwnProperty.call(done, k)) return;
      const val = evalFormula(item.formula, item.ws.name);
      if (isPend(val)) return;
      if (typeof val === "number" && !isFinite(val)) return;
      done[k] = val;
      item.ws.getCell(item.addr).value = { formula: item.formula, result: val };
      moved++;
    });
    if (!moved) break;
  }
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
  const names = Array.isArray(spec)
    ? wb.worksheets.map(function (w) { return w.name; })
    : Object.keys(spec);
  names.forEach((name) => {
    const ws = wb.getWorksheet(name);
    if (!ws) return;
    const area = Array.isArray(spec) ? (spec[0] || "A1:I34") : spec[name];
    const m = area.match(/^([A-Z]+)(\d+):([A-Z]+)(\d+)$/);
    if (!m) return;
    const lastC = colLetterToNum(m[3]);
    const lastR = Number(m[4]);

    ws.pageSetup.paperSize = 9;
    ws.pageSetup.orientation = lastC > 10 ? "landscape" : "portrait";
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

function dropBadNames(wb) {
  const dn = wb && wb.definedNames;
  if (!dn || !Array.isArray(dn.model)) return;
  const keep = [];
  dn.model.forEach(function (n) {
    const blob = JSON.stringify(n || {});
    if (blob.indexOf("[") >= 0) return;
    keep.push(n);
  });
  dn.model.length = 0;
  keep.forEach(function (n) { dn.model.push(n); });
}

async function patchFitXml(xlsxPath) {
  let JSZip;
  try {
    JSZip = require("jszip");
  } catch (_e1) {
    try { JSZip = require("exceljs/node_modules/jszip"); } catch (_e2) { return; }
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
  const wbName = "xl/workbook.xml";
  if (zip.file(wbName)) {
    let wbXml = await zip.file(wbName).async("string");
    wbXml = wbXml.replace(/<definedName\b[^>]*>[^<]*\[[^<]*<\/definedName>/g, "");
    wbXml = wbXml.replace(/<definedNames>\s*<\/definedNames>/g, "");
    zip.file(wbName, wbXml);
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

const LO_XCU =
  '<?xml version="1.0" encoding="UTF-8"?>\n' +
  '<oor:items xmlns:oor="http://openoffice.org/2001/registry" xmlns:xs="http://www.w3.org/2001/XMLSchema" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">\n' +
  '<item oor:path="/org.openoffice.Office.Calc/Formula/Load"><prop oor:name="OOXMLRecalcMode" oor:op="fuse"><value>0</value></prop></item>\n' +
  "</oor:items>\n";

let loChain = Promise.resolve();
function loQueue(job) {
  const run = loChain.then(job, job);
  loChain = run.then(
    function () {},
    function () {}
  );
  return run;
}

const LO_DIR = "/tmp/lo-profile";

function resetLoProfile() {
  try { fs.rmSync(LO_DIR, { recursive: true, force: true }); } catch (_e) {}
  fs.mkdirSync(LO_DIR, { recursive: true });
}

function loSharedArg() {
  if (!fs.existsSync(LO_DIR)) fs.mkdirSync(LO_DIR, { recursive: true });
  return "-env:UserInstallation=file:///tmp/lo-profile";
}

function stopSoffice() {
  try { execFileSync("pkill", ["-9", "-f", "soffice"], { stdio: "ignore" }); } catch (_e) {}
  [path.join(LO_DIR, ".lock"), path.join(LO_DIR, "user", ".lock")].forEach(function (p) {
    try { fs.rmSync(p, { force: true }); } catch (_e) {}
  });
}

function runSoffice(args, timeoutMs) {
  return new Promise((resolve, reject) => {
    let done = false;
    const child = spawn(sofficeBin(), args, { stdio: "ignore" });
    const timer = setTimeout(() => {
      if (done) return;
      done = true;
      try { child.kill("SIGKILL"); } catch (_e) {}
      stopSoffice();
      resetLoProfile();
      const err = new Error("soffice timeout");
      err.timeout = true;
      reject(err);
    }, timeoutMs || 55000);
    child.on("error", (err) => {
      if (done) return;
      done = true;
      clearTimeout(timer);
      reject(err);
    });
    child.on("close", (code) => {
      if (done) return;
      done = true;
      clearTimeout(timer);
      if (code) reject(new Error("soffice exit " + code));
      else resolve();
    });
  });
}

function recalcXlsxFile(xlsxPath) {
  return loQueue(async function () {
    const tmp = path.join(path.dirname(xlsxPath), "recalc-" + Date.now() + "-" + Math.random().toString(36).slice(2, 6));
    fs.mkdirSync(tmp, { recursive: true });
    const name = path.basename(xlsxPath);
    try {
      await runSoffice(
        [loSharedArg(), "--headless", "--norestore", "--nolockcheck", "--convert-to", "xlsx", "--outdir", tmp, xlsxPath],
        60000
      );
      const produced = path.join(tmp, name);
      if (fs.existsSync(produced) && fs.statSync(produced).size > 1000) fs.copyFileSync(produced, xlsxPath);
      else console.error("xlsx resave fail", name, "no output");
    } catch (e) {
      console.error("xlsx resave fail", name, e.message || e);
    } finally {
      try { fs.rmSync(tmp, { recursive: true, force: true }); } catch (_e) {}
    }
  });
}

function convertWithSoffice(xlsxPath) {
  return loQueue(function () {
    return new Promise((resolve, reject) => {
      const dir = path.dirname(xlsxPath);
      const pdfPath = xlsxPath.replace(/\.xlsx$/i, ".pdf");
      try { if (fs.existsSync(pdfPath)) fs.unlinkSync(pdfPath); } catch (_e) {}
      if (!fs.existsSync(LO_DIR)) fs.mkdirSync(LO_DIR, { recursive: true });
      const child = spawn(sofficeBin(), [
        "-env:UserInstallation=file:///tmp/lo-profile",
        "--headless", "--norestore", "--nolockcheck",
        "--convert-to", "pdf", "--outdir", dir, xlsxPath
      ], {
        stdio: "ignore",
        env: Object.assign({}, process.env, { SAL_USE_VCLPLUGIN: "svp", HOME: "/tmp" })
      });
      let done = false;
      const finish = (err) => {
        if (done) return;
        done = true;
        clearTimeout(timer);
        clearInterval(poll);
        let ok = false;
        try { ok = fs.existsSync(pdfPath) && fs.statSync(pdfPath).size > 500; } catch (_e) {}
        if (ok) return resolve(pdfPath);
        try { child.kill("SIGKILL"); } catch (_e) {}
        reject(err || new Error("pdf missing"));
      };
      const poll = setInterval(() => {
        try {
          if (fs.existsSync(pdfPath) && fs.statSync(pdfPath).size > 500) finish(null);
        } catch (_e) {}
      }, 400);
      const timer = setTimeout(() => finish(new Error("soffice timeout")), 180000);
      child.on("error", (err) => finish(err));
      child.on("close", () => setTimeout(() => finish(new Error("pdf missing")), 500));
    });
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

function unlockSheets(wb) {
  wb.worksheets.forEach(function (ws) {
    ws.sheetProtection = undefined;
  });
}

const pdfJobs = {};
function startPdfJob(xlsxFull, pdfName) {
  const id = pdfName.replace(/\.pdf$/i, "");
  pdfJobs[id] = { status: "run" };
  convertWithSoffice(xlsxFull).then(function (produced) {
    const pdfFull = path.join(OUT_DIR, pdfName);
    if (produced !== pdfFull && fs.existsSync(produced)) fs.copyFileSync(produced, pdfFull);
    if (!fs.existsSync(pdfFull)) throw new Error("pdf missing");
    pdfJobs[id] = { status: "ok", pdf: "/api/download/" + pdfName };
  }).catch(function (e) {
    pdfJobs[id] = { status: "fail", error: String(e.message || e) };
  });
  return id;
}

async function writeAndRespond(req, res, wb, d, prefix, areas, kind) {
  const output = d.output || "xlsx";
  if (output !== "pdf" && wb.calcProperties) wb.calcProperties.fullCalcOnLoad = true;
  if (output !== "pdf") {
    try { bakeFormulaResults(wb); } catch (_e) {}
  }
  dropBadNames(wb);
  unlockSheets(wb);
  applyOnePage(wb, areas);
  const safe = String(d.village || "gam").replace(/[^a-zA-Z0-9._-]+/g, "_");
  const stamp = Date.now();
  const xlsxName = `${prefix}_${safe}_${stamp}.xlsx`;
  const pdfName = `${prefix}_${safe}_${stamp}.pdf`;
  const xlsxFull = path.join(OUT_DIR, xlsxName);
  const pdfFull = path.join(OUT_DIR, pdfName);
  await wb.xlsx.writeFile(xlsxFull);
  await patchFitXml(xlsxFull);
  if ((output === "xlsx" || output === "both") && String(process.env.XLSX_RESAVE || "0") === "1") {
    await recalcXlsxFile(xlsxFull);
    await patchFitXml(xlsxFull);
  }

  let pdfUrl = null;
  let pdfError = null;
  let pdfJob = null;
  if (output === "pdf" || output === "both") {
    pdfJob = startPdfJob(xlsxFull, pdfName);
  }

  const Lm = Number(d.length_m || 0);
  const Wm = Number(d.width_m || 0);
  const areaM = Lm * Wm;
  logEvent(kind, {
    village: d.village,
    taluka: d.taluka,
    jilla: d.jilla,
    amounting: Number(d.amounting || 0),
    work_name: d.work_name || "",
    length_m: Lm,
    width_m: Wm,
    area: areaM,
    brass: kind === "estimate_paver" ? areaM * 10.7584 / 100 : 0,
    output,
    xlsx: xlsxName,
    pdf: pdfUrl
  }, req);
  try {
    const kindStr = String(kind || "");
    const isBill = /^bill_(paver|cc|gutter|pipe)$/.test(kindStr);
    const isEst = /^estimate_/.test(kindStr);
    if (isBill || isEst) {
      const net = Number(d.net || 0);
      const rec = dbAppend(isBill ? DB_BILL : DB_EST, {
        kind: isBill ? "bill" : "estimate",
        type: d.type || (kind === "estimate_paver" || kind === "bill_paver" ? "Paver" : kind === "bill_gutter" ? "Gutter" : kind === "bill_pipe" ? "Pipe" : "CC"),
        village: d.village, taluka: d.taluka, jilla: d.jilla,
        work_name: d.work_name || "",
        fund_head: d.fund_head || d.grant || "",
        amounting: isBill ? (net || Number(d.amounting || 0)) : Number(d.amounting || 0),
        net: net,
        day: d.day || istDay(),
        length_m: Lm, width_m: Wm, area: areaM,
        brass: (kind === "estimate_paver" || kind === "bill_paver") ? areaM * 10.7584 / 100 : 0,
        prepared_by: d.prepared_by || "",
        mb_no: d.mb_no || ""
      });
      if (typeof sbOn === "function" && sbOn()) {
        sbInsert(isBill ? "bills" : "estimates", rec).catch(function (e) { console.error(e.message); });
      }
    }
  } catch (_e) {}
  const wantXlsx = output === "xlsx" || output === "both";
  res.json({
    ok: true,
    xlsx: wantXlsx ? `/api/download/${xlsxName}` : null,
    pdf: pdfUrl,
    pdf_job: pdfJob,
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


app.get("/api/stats", (_req, res) => {
  const empty = { total: 0, cc: 0, paver: 0, amount: 0, today: 0, talukas: [], recent: [] };
  if (!fs.existsSync(LOG_FILE)) return res.json(empty);
  const lines = fs.readFileSync(LOG_FILE, "utf8").trim().split("\n").filter(Boolean);
  const today = istDay();
  const talMap = {};
  let total = 0, cc = 0, paver = 0, amount = 0, todayN = 0;
  const recent = [];
  for (const line of lines) {
    let ev;
    try { ev = JSON.parse(line); } catch (_e) { continue; }
    if (ev.kind !== "estimate_cc" && ev.kind !== "estimate_paver") continue;
    const pl = ev.payload || {};
    total++;
    if (ev.kind === "estimate_cc") cc++;
    else paver++;
    amount += Number(pl.amounting || 0);
    if (String(ev.ts || "") && recordDay({ ts: ev.ts }) === today) todayN++;
    const tk = pl.taluka || "—";
    talMap[tk] = (talMap[tk] || 0) + 1;
    recent.push({
      ts: ev.ts,
      type: ev.kind === "estimate_paver" ? "Paver" : "CC",
      village: pl.village || "",
      taluka: tk,
      amounting: Number(pl.amounting || 0),
      work_name: pl.work_name || "",
      length_m: Number(pl.length_m || 0),
      width_m: Number(pl.width_m || 0),
      area: Number(pl.area || 0),
      brass: Number(pl.brass || 0)
    });
  }
  const talukas = Object.keys(talMap)
    .map((k) => ({ name: k, n: talMap[k] }))
    .sort((a, b) => b.n - a.n)
    .slice(0, 8);
  res.json({
    total, cc, paver, amount, today: todayN, talukas,
    recent: recent.reverse().slice(0, 8)
  });
});


app.get("/api/bills", (req, res) => {
  try {
    const day = String(req.query.day || req.query.date || "").slice(0, 10);
    const file = DB_BILL;
    let rows = notDeleted(dbRead(file, 500)).filter(function (r) { return r.kind === "bill" || String(r.kind||"").indexOf("bill")>=0; });
    if (day) {
      rows = rows.filter(function (r) {
        const d = recordDay(r);
        return d === day;
      });
    }
    res.json({ ok: true, bills: rows.slice(-80).reverse() });
  } catch (e) {
    res.json({ ok: false, bills: [], error: String(e.message || e) });
  }
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
    setVal(face, "D30", d.prepared_by);
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
    setVal(meas, "G14", btQty);
    setVal(meas, "K17", btQty);
    setVal(meas, "K20", murQty);
    setVal(meas, "A16", "Item No. :- 4 Spreading the stone aggregates for soiling and W. B. M. including filling the inter stices forming the surface to required camber and gradient (excluding spreading of blindage) (ii) 40 mm to 63 mm size aggreates (HB)");
    setVal(meas, "K22", area);
    setVal(meas, "G26", area);
    setVal(meas, "I26", ccT);
    setVal(meas, "K26", ccQty);
    setVal(meas, "A28", "Item No :- 7 Testing charges for Kapchi,Metal,Sand,Cement,C.C. Cube as per schedule of testing");
    setVal(meas, "A30", "Item No :- 8 Providing and fixing number plate of marble stone of required size set in C. M. 1 : 4 including finishing and engraving letters etc. complete.");
    setVal(meas, "K31", 1);
    try { meas.unMergeCells("A30:L30"); } catch (e) {}
    try { meas.mergeCells("A30:L30"); } catch (e) {}
    meas.getCell("A30").alignment = { wrapText: true, horizontal: "left", vertical: "middle" };
    meas.getRow(30).height = 36;

    if (abs) {
      setVal(abs, "B2", work);
      setVal(abs, "A4", boxQty);
      setVal(abs, "F4", a1);
      setVal(abs, "A6", btQty);
      setVal(abs, "F6", a2);
      setVal(abs, "A8", murQty);
      setVal(abs, "F8", a3);
      setVal(abs, "C10", "Item No. :- 4 Spreading the stone aggregates for soiling and W. B. M. including filling the inter stices forming the surface to required camber and gradient (excluding spreading of blindage) (ii) 40 mm to 63 mm size aggreates (HB)");
      setVal(abs, "A10", btQty);
      setVal(abs, "F10", a4);
      setVal(abs, "A12", murQty);
      setVal(abs, "F12", a5);
      setVal(abs, "A14", 0);
      setVal(abs, "F14", a6r);
      setVal(abs, "A16", ccQty);
      setVal(abs, "F16", a6c);
      setVal(abs, "C18", "Item No :- 7 Testing charges for Kapchi,Metal,Sand,Cement,C.C. Cube as per schedule of testing");
      setVal(abs, "C20", "Item No :- 8 Providing and fixing number plate of marble stone of required size set in C. M. 1 : 4 including finishing and engraving letters etc. complete.");
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
    setVal(face, "I2", d.jilla);
    setVal(face, "F4", d.subdiv_address);
    setVal(face, "I4", d.nani_address);
    setVal(face, "F5", d.subdiv_address);
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

app.get("/api/pdf-job/:id", (req, res) => {
  const job = pdfJobs[req.params.id];
  if (!job) return res.json({ ok: false, error: "job નથી" });
  if (job.status === "ok") return res.json({ ok: true, pdf: job.pdf });
  if (job.status === "fail") return res.json({ ok: false, error: job.error || "PDF ન બની" });
  res.json({ ok: true, pending: true });
});


const DB_DIR = path.join(__dirname, "db");
if (!fs.existsSync(DB_DIR)) fs.mkdirSync(DB_DIR);
const DB_EST = path.join(DB_DIR, "estimates.jsonl");
const DB_BILL = path.join(DB_DIR, "bills.jsonl");
const DB_SITE = path.join(DB_DIR, "site.jsonl");
const DB_MEDIA = path.join(DB_DIR, "media.jsonl");
const DB_KACHU = path.join(DB_DIR, "kachu.jsonl");
const MEDIA_DIR = path.join(DB_DIR, "media");
if (!fs.existsSync(MEDIA_DIR)) fs.mkdirSync(MEDIA_DIR);

function dbAppend(file, rec) {
  rec.id = rec.id || (Date.now() + "-" + Math.random().toString(36).slice(2, 8));
  rec.ts = rec.ts || new Date().toISOString();
  fs.appendFileSync(file, JSON.stringify(rec) + "\n");
  return rec;
}
function dbRemove(file, id) {
  if (!file || !id || !fs.existsSync(file)) return;
  const lines = fs.readFileSync(file, "utf8").split("\n").filter(Boolean);
  const keep = lines.filter(function (line) {
    try { return String(JSON.parse(line).id || "") !== String(id); } catch (_e) { return true; }
  });
  fs.writeFileSync(file, keep.length ? keep.join("\n") + "\n" : "");
}
function tombPath() { return path.join(DB_DIR, "deleted.json"); }
function tombRead() {
  try {
    const arr = JSON.parse(fs.readFileSync(tombPath(), "utf8"));
    return Array.isArray(arr) ? arr.map(String) : [];
  } catch (_e) { return []; }
}
function tombAdd(id) {
  if (!id) return;
  const arr = tombRead();
  if (arr.indexOf(String(id)) < 0) arr.push(String(id));
  fs.writeFileSync(tombPath(), JSON.stringify(arr));
}
function forgetRows(file, body) {
  const id = String((body && body.id) || "");
  const work = String((body && body.work_name) || "");
  const ts = String((body && body.ts) || "");
  const mb = String((body && body.mb_no) || "");
  if (id) tombAdd(id);
  if (!file || !fs.existsSync(file)) return;
  const lines = fs.readFileSync(file, "utf8").split("\n").filter(Boolean);
  const keep = [];
  lines.forEach(function (line) {
    let r = null;
    try { r = JSON.parse(line); } catch (_e) { keep.push(line); return; }
    const sameId = id && String(r.id || "") === id;
    const sameWork = work && String(r.work_name || "") === work && ((ts && String(r.ts || "") === ts) || (mb && String(r.mb_no || "") === mb));
    if (sameId || sameWork) tombAdd(r.id);
    else keep.push(line);
  });
  fs.writeFileSync(file, keep.length ? keep.join("\n") + "\n" : "");
}
function notDeleted(items) {
  const gone = {};
  tombRead().forEach(function (id) { gone[id] = 1; });
  return (items || []).filter(function (x) { return x && !gone[String(x.id || "")]; });
}
function dbRead(file, limit) {
  if (!fs.existsSync(file)) return [];
  const lines = fs.readFileSync(file, "utf8").trim().split("\n").filter(Boolean);
  const out = [];
  for (const line of lines) {
    try { out.push(JSON.parse(line)); } catch (_e) {}
  }
  return out.reverse().slice(0, limit || 200);
}

const SB_URL = (process.env.SUPABASE_URL || "").replace(/\/$/, "");
const SB_KEY = process.env.SUPABASE_ANON_KEY || process.env.SUPABASE_KEY || "";
function sbOn() { return !!(SB_URL && SB_KEY); }
function mergeItems(remote, local) {
  const out = [];
  const seen = {};
  function key(x) {
    return String(x.id || "") || (String(x.work_name || "") + "|" + String(x.ts || "") + "|" + String(x.amounting || ""));
  }
  (remote || []).concat(local || []).forEach(function (x) {
    const k = key(x);
    if (!k || seen[k]) return;
    seen[k] = 1;
    out.push(x);
  });
  out.sort(function (a, b) { return String(b.ts || "").localeCompare(String(a.ts || "")); });
  return out;
}

function sbPick(table, row) {
  const cols = {
    estimates: ["id","ts","type","village","taluka","jilla","work_name","fund_head","amounting","length_m","width_m","area","brass","prepared_by"],
    site_measures: ["id","ts","type","work_name","amounting","rows","area","brass","bill","gps","estimate_id","taluka","village","fund_head","grant_head","contractor"],
    media: ["id","ts","kind","work_name","gps","url"],
    kachu_bills: ["id","ts","type","work_name","village","amounting","total","net","test_qty","name_plate","preview","xlsx"],
    bills: ["id","ts","type","work_name","village","taluka","fund_head","amounting","prepared_by","mb_no"]
  }[table] || Object.keys(row);
  const o = {};
  cols.forEach(function (k) { if (row[k] !== undefined) o[k] = row[k]; });
  return o;
}
async function sbInsert(table, row) {
  const r = await fetch(SB_URL + "/rest/v1/" + table, {
    method: "POST",
    headers: {
      apikey: SB_KEY,
      Authorization: "Bearer " + SB_KEY,
      "Content-Type": "application/json",
      Prefer: "return=representation"
    },
    body: JSON.stringify(sbPick(table, row))
  });
  if (!r.ok) {
    const err = await r.text();
    if (table === "estimates") {
      const slim = Object.assign({}, sbPick(table, row));
      delete slim.fund_head;
      const r2 = await fetch(SB_URL + "/rest/v1/" + table, {
        method: "POST",
        headers: {
          apikey: SB_KEY,
          Authorization: "Bearer " + SB_KEY,
          "Content-Type": "application/json",
          Prefer: "return=representation"
        },
        body: JSON.stringify(slim)
      });
      if (r2.ok) {
        const js2 = await r2.json();
        return Array.isArray(js2) ? js2[0] : js2;
      }
    }
    throw new Error(err);
  }
  const js = await r.json();
  return Array.isArray(js) ? js[0] : js;
}
async function sbSelect(table, limit) {
  const r = await fetch(SB_URL + "/rest/v1/" + table + "?select=*&order=ts.desc&limit=" + (limit || 200), {
    headers: { apikey: SB_KEY, Authorization: "Bearer " + SB_KEY }
  });
  if (!r.ok) throw new Error(await r.text());
  return await r.json();
}
async function sbDelete(table, id) {
  if (!sbOn() || !id) return 0;
  const key = process.env.SUPABASE_SERVICE_ROLE || process.env.SUPABASE_SERVICE_KEY || SB_KEY;
  const r = await fetch(SB_URL + "/rest/v1/" + table + "?id=eq." + encodeURIComponent(id), {
    method: "DELETE",
    headers: {
      apikey: key,
      Authorization: "Bearer " + key,
      Prefer: "return=representation"
    }
  });
  if (!r.ok) {
    console.error("delete", table, id, r.status, await r.text());
    return 0;
  }
  const js = await r.json().catch(function () { return []; });
  return Array.isArray(js) ? js.length : 0;
}
async function sbUpload(name, buf) {
  const r = await fetch(SB_URL + "/storage/v1/object/site-media/" + name, {
    method: "POST",
    headers: {
      apikey: SB_KEY,
      Authorization: "Bearer " + SB_KEY,
      "Content-Type": "image/jpeg",
      "x-upsert": "true"
    },
    body: buf
  });
  if (!r.ok) throw new Error(await r.text());
  return SB_URL + "/storage/v1/object/public/site-media/" + name;
}


app.post("/api/estimate/gutter", async (req, res) => {
  try {
    const d = req.body || {};
    const wb = new ExcelJS.Workbook();
    await wb.xlsx.readFile(path.join(__dirname, "skeleton-gutter.xlsx"));
    const face = wb.getWorksheet("Estimate");
    const abs = wb.getWorksheet("Abstract");
    const meas = wb.getWorksheet("Measurement");
    const test = wb.getWorksheet("TEST-SITE");
    if (!face || !abs || !meas) throw new Error("skeleton-gutter.xlsx sheets missing");

    const n = (k) => Number(d[k] || 0);
    const work = d.work_name || "";
    const taluka = d.taluka || "";
    const village = d.village || "";
    const say = Number(d.amounting || 0);
    const fund = d.fund_head || "";
    const gEx = Array.isArray(d.gEx) ? d.gEx : [];
    const gDe = Array.isArray(d.gDe) ? d.gDe : [];
    const sumDia = (dia) => gEx.filter(r => Number(r.dia)===dia).reduce((a,r)=>a+Number(r.l||0),0);
    const lastWD = (dia, dw, dd) => {
      const arr=gEx.filter(r => Number(r.dia)===dia);
      const r=arr[arr.length-1];
      return {w:Number((r&&r.w)||dw), d:Number((r&&r.d)||dd)};
    };

    setVal(face, "F2", d.division || d.jilla || "");
    setVal(face, "I2", d.jilla || "");
    setVal(face, "F4", d.subdiv_address || "");
    setVal(face, "F5", d.subdiv_address || "");
    setVal(face, "I4", d.nani_address || "");
    setVal(face, "D8", fund);
    setVal(face, "H18", taluka);
    setVal(face, "C20", work);
    setVal(face, "G21", say);
    setVal(face, "D27", d.prepared_by || "");
    setVal(face, "D29", d.prepared_by || "");
    setVal(face, "B34", d.sr_no || "");
    setVal(face, "C34", d.ss_details || "");
    setVal(face, "G40", taluka);

    const wd225=lastWD(225,0.45,0.825), wd300=lastWD(300,0.45,0.90), wd450=lastWD(450,0.75,1.15);
    const wd600=lastWD(600,0.90,1.35), wd900=lastWD(900,1.20,1.80), wd1200=lastWD(1200,1.50,2.20);
    const rows = [
      { L: sumDia(225), W: wd225.w, D: wd225.d, E: "E9", G: "G9", I: "I9" },
      { L: sumDia(300), W: wd300.w, D: wd300.d, E: "E10", G: "G10", I: "I10" },
      { L: sumDia(450), W: wd450.w, D: wd450.d, E: "E11", G: "G11", I: "I11" },
      { L: sumDia(600), W: wd600.w, D: wd600.d, E: "E12", G: "G12", I: "I12" },
      { L: sumDia(900), W: wd900.w, D: wd900.d, E: "E13", G: "G13", I: "I13" },
      { L: sumDia(1200), W: wd1200.w, D: wd1200.d, E: "E14", G: "G14", I: "I14" }
    ];
    rows.forEach(function (r) {
      setVal(meas, r.E, r.L);
      setVal(meas, r.G, r.W);
      setVal(meas, r.I, r.D);
    });
    const demoL = gDe.reduce((a,r)=>a+Number(r.l||0),0);
    const demoW = Number((gDe[0]&&gDe[0].w)||0.45);
    const demoD = Number((gDe[0]&&gDe[0].d)||0.10);
    setVal(meas, "E5", demoL);
    setVal(meas, "G5", demoW);
    setVal(meas, "I5", demoD);
    setVal(meas, "E35", n("gCh60"));
    setVal(meas, "E36", n("gCh90"));
    setVal(meas, "E37", n("gCh139"));
    setVal(meas, "E38", n("gCh1313"));
    setVal(meas, "E62", n("gPlate") || 1);
    const pipeL = rows.reduce((a, r) => a + r.L, 0);
    if (demoL > 0) {
      setVal(meas, "G58", n("gBedW") || 0.45);
      setVal(meas, "I58", n("gBedT") || 0.05);
    } else {
      setVal(meas, "G58", 0);
      setVal(meas, "I58", 0);
    }
    meas.getCell("K58").value = { formula: "C58*E58*G58*I58" };
    setVal(meas, "C2", work);
    setVal(abs, "C2", work);
    if (test) setVal(test, "B1", work);

    await writeAndRespond(req, res, wb, d, "GUTTER", {
      Estimate: "A1:I41",
      Abstract: "A1:F46",
      Measurement: "A1:K63",
      "TEST-SITE": "A1:G36"
    }, "estimate_gutter");
  } catch (e) {
    console.error(e);
    res.status(500).json({ ok: false, error: String(e.message || e) });
  }
});

app.post("/api/db/estimate", (req, res) => {
  const b = req.body || {};
  const rec = dbAppend(DB_EST, {
    kind: "estimate",
    type: b.type || "",
    village: b.village || "",
    taluka: b.taluka || "",
    jilla: b.jilla || "",
    work_name: b.work_name || "",
    fund_head: b.fund_head || b.grant || "",
    amounting: Number(b.amounting || 0),
    length_m: Number(b.length_m || 0),
    width_m: Number(b.width_m || 0),
    brass: Number(b.brass || 0),
    prepared_by: b.prepared_by || ""
  });
  if (sbOn()) sbInsert("estimates", rec).catch(function(e){ console.error(e.message); });
  res.json({ ok: true, id: rec.id });
});
app.get("/api/db/bills", async (_req, res) => {
  const local = notDeleted(dbRead(DB_BILL, 200));
  try {
    if (sbOn()) return res.json({ ok: true, items: notDeleted(await sbSelect("bills", 200)) });
  } catch (e) { console.error(e.message); }
  res.json({ ok: true, items: local });
});
app.post("/api/db/bills/delete", async (req, res) => {
  const id = String((req.body && req.body.id) || "");
  if (!id) return res.json({ ok: false });
  forgetRows(DB_BILL, req.body || {});
  let removed = 0;
  try { removed = await sbDelete("bills", id); } catch (e) { console.error(e.message); }
  res.json({ ok: true, removed: removed });
});
app.get("/api/db/tour", async (req, res) => {
  try {
    if (!sbOn()) return res.json({ ok: true, items: [] });
    const email = String((req.query && req.query.email) || "").trim();
    let url = SB_URL + "/rest/v1/tour_days?select=*&order=day.desc&limit=400";
    if (email) url += "&email=eq." + encodeURIComponent(email);
    const r = await fetch(url, {
      headers: { apikey: SB_KEY, Authorization: "Bearer " + SB_KEY }
    });
    if (!r.ok) throw new Error(await r.text());
    res.json({ ok: true, items: await r.json() });
  } catch (e) {
    res.json({ ok: false, items: [], error: String(e.message || e) });
  }
});
app.post("/api/db/tour", async (req, res) => {
  try {
    if (!sbOn()) return res.json({ ok: false, error: "no supabase" });
    const b = req.body || {};
    const row = {
      email: String(b.email || "").trim() || "shared",
      day: String(b.day || "").slice(0, 10),
      act: b.act || "none",
      note: b.note || "",
      meet: b.meet || "",
      time_from: b.time_from || "",
      time_to: b.time_to || "",
      leave_type: b.leave_type || ""
    };
    if (!row.day) return res.json({ ok: false, error: "તારીખ નથી" });
    const r = await fetch(SB_URL + "/rest/v1/tour_days?on_conflict=email,day", {
      method: "POST",
      headers: {
        apikey: SB_KEY,
        Authorization: "Bearer " + SB_KEY,
        "Content-Type": "application/json",
        Prefer: "resolution=merge-duplicates,return=minimal"
      },
      body: JSON.stringify(row)
    });
    if (!r.ok) throw new Error(await r.text());
    res.json({ ok: true });
  } catch (e) {
    res.json({ ok: false, error: String(e.message || e) });
  }
});
app.get("/api/db/tour-profile", async (req, res) => {
  try {
    if (!sbOn()) return res.json({ ok: true, item: null });
    const email = String((req.query && req.query.email) || "").trim();
    if (!email) return res.json({ ok: true, item: null });
    const url = SB_URL + "/rest/v1/tour_profile?select=*&email=eq." + encodeURIComponent(email) + "&limit=1";
    const r = await fetch(url, { headers: { apikey: SB_KEY, Authorization: "Bearer " + SB_KEY } });
    if (!r.ok) throw new Error(await r.text());
    const rows = await r.json();
    res.json({ ok: true, item: rows[0] || null });
  } catch (e) {
    res.json({ ok: false, item: null, error: String(e.message || e) });
  }
});
app.post("/api/db/tour-profile", async (req, res) => {
  try {
    if (!sbOn()) return res.json({ ok: false, error: "no supabase" });
    const b = req.body || {};
    const row = {
      email: String(b.email || "").trim() || "shared",
      name: b.name || "",
      office: b.office || "",
      mobile: b.mobile || "",
      leave: b.leave || {}
    };
    const r = await fetch(SB_URL + "/rest/v1/tour_profile?on_conflict=email", {
      method: "POST",
      headers: {
        apikey: SB_KEY,
        Authorization: "Bearer " + SB_KEY,
        "Content-Type": "application/json",
        Prefer: "resolution=merge-duplicates,return=minimal"
      },
      body: JSON.stringify(row)
    });
    if (!r.ok) throw new Error(await r.text());
    res.json({ ok: true });
  } catch (e) {
    res.json({ ok: false, error: String(e.message || e) });
  }
});
app.get("/api/db/estimates", async (_req, res) => {
  const local = notDeleted(dbRead(DB_EST, 200));
  try {
    if (sbOn()) {
      const remote = notDeleted(await sbSelect("estimates", 200));
      return res.json({ ok: true, items: remote });
    }
  } catch (e) { console.error(e.message); }
  res.json({ ok: true, items: local });
});
app.post("/api/db/site", (req, res) => {
  const b = req.body || {};
  const rec = dbAppend(DB_SITE, {
    kind: "site",
    type: b.type || "",
    work_name: b.work_name || "",
    amounting: Number(b.amounting || 0),
    rows: b.rows || [],
    area: Number(b.area || 0),
    brass: Number(b.brass || 0),
    bill: Number(b.bill || 0),
    gps: b.gps || "",
    estimate_id: b.estimate_id || "",
    taluka: b.taluka || "",
    village: b.village || "",
    fund_head: b.fund_head || b.grant || b.grant_head || "",
    grant_head: b.grant_head || b.grant || b.fund_head || "",
    contractor: b.contractor || ""
  });
  if (sbOn()) sbInsert("site_measures", rec).catch(function(e){ console.error(e.message); });
  res.json({ ok: true, id: rec.id });
});
app.get("/api/db/site", async (_req, res) => {
  const local = notDeleted(dbRead(DB_SITE, 200));
  try {
    if (sbOn()) {
      const remote = notDeleted(await sbSelect("site_measures", 200));
      return res.json({ ok: true, items: mergeItems(remote, local) });
    }
  } catch (e) { console.error(e.message); }
  res.json({ ok: true, items: local });
});
app.post("/api/db/site/delete", async (req, res) => {
  const id = String((req.body && req.body.id) || "");
  if (!id) return res.json({ ok: false });
  forgetRows(DB_SITE, req.body || {});
  let removed = 0;
  try { removed = await sbDelete("site_measures", id); } catch (e) { console.error(e.message); }
  res.json({ ok: true, removed: removed });
});

app.post("/api/db/kachu", (req, res) => {
  const b = req.body || {};
  const rec = dbAppend(DB_KACHU, {
    type: b.type || "paver",
    work_name: b.work_name || "",
    village: b.village || "",
    amounting: Number(b.amounting || 0),
    total: Number(b.total || 0),
    net: Number(b.net || 0),
    test_qty: Number(b.test_qty || 0),
    name_plate: Number(b.name_plate || 0),
    preview: b.preview || {},
    xlsx: b.xlsx || ""
  });
  if (sbOn()) sbInsert("kachu_bills", rec).catch(function(e){ console.error(e.message); });
  res.json({ ok: true, id: rec.id });
});
app.get("/api/db/kachu", async (_req, res) => {
  try { if (sbOn()) return res.json({ ok: true, items: await sbSelect("kachu_bills", 200) }); }
  catch (e) { console.error(e.message); }
  res.json({ ok: true, items: dbRead(DB_KACHU, 200) });
});
app.post("/api/db/media", (req, res) => {
  const b = req.body || {};
  let file = "";
  let url = "";
  if (b.data && String(b.data).startsWith("data:")) {
    const raw = String(b.data);
    const comma = raw.indexOf(",");
    const buf = Buffer.from(raw.slice(comma + 1), "base64");
    if (buf.length <= 4 * 1024 * 1024) {
      file = (b.kind || "media") + "-" + Date.now() + ".jpg";
      fs.writeFileSync(path.join(MEDIA_DIR, file), buf);
      url = "/db/media/" + file;
    }
  }
  const rec = dbAppend(DB_MEDIA, {
    kind: b.kind || "photo",
    work_name: b.work_name || "",
    gps: b.gps || "",
    file, url
  });
  if (sbOn() && b.data && String(b.data).startsWith("data:")) {
    const raw = String(b.data);
    const comma = raw.indexOf(",");
    const buf = Buffer.from(raw.slice(comma + 1), "base64");
    const name = (b.kind || "media") + "-" + Date.now() + ".jpg";
    sbUpload(name, buf).then(function (u) {
      rec.url = u;
      sbInsert("media", { id: rec.id, ts: rec.ts, kind: rec.kind, work_name: rec.work_name, gps: rec.gps, url: u }).catch(function(){});
      res.json({ ok: true, id: rec.id, file: u });
    }).catch(function (e) {
      console.error(e.message);
      res.json({ ok: true, id: rec.id, file: url });
    });
    return;
  }
  res.json({ ok: true, id: rec.id, file: url });
});
app.get("/api/db/media", async (_req, res) => {
  try { if (sbOn()) return res.json({ ok: true, items: await sbSelect("media", 80) }); }
  catch (e) { console.error(e.message); }
  const items = dbRead(DB_MEDIA, 80).map((m) => ({
    id: m.id, ts: m.ts, kind: m.kind, work_name: m.work_name, gps: m.gps,
    url: m.url || (m.file ? ("/db/media/" + m.file) : "")
  }));
  res.json({ ok: true, items });
});
app.use("/db/media", express.static(MEDIA_DIR));


app.get("/api/db/bundle", async (req, res) => {
  const work = String(req.query.work || "");
  const id = String(req.query.id || "");
  function match(x) {
    if (id && String(x.id||x.estimate_id||"") === id) return true;
    if (work && String(x.work_name||"") === work) return true;
    return false;
  }
  try {
    const ests = sbOn() ? await sbSelect("estimates", 200) : dbRead(DB_EST, 200);
    const sites = sbOn() ? await sbSelect("site_measures", 200) : dbRead(DB_SITE, 200);
    const media = sbOn() ? await sbSelect("media", 80) : dbRead(DB_MEDIA, 80);
    const estimate = ests.find(match) || ests.find((x)=> work && String(x.work_name||"").indexOf(work)>=0) || null;
    res.json({
      ok: true,
      estimate,
      site: sites.filter(match),
      media: media.filter(match)
    });
  } catch (e) {
    res.json({ ok: false, error: String(e.message||e) });
  }
});
app.get("/api/db/health", (_req, res) => {
  res.json({ ok: true, supabase: sbOn(), url: SB_URL ? SB_URL.replace(/https:\/\//,"") : "" });
});
app.get("/api/auth/config", (_req, res) => {
  res.json({ ok: true, url: SB_URL || "", anon: SB_KEY || "" });
});
async function sbProfiles(method, path, body) {
  const r = await fetch(SB_URL + "/rest/v1/" + path, {
    method: method,
    headers: {
      apikey: SB_KEY,
      Authorization: "Bearer " + SB_KEY,
      "Content-Type": "application/json",
      Prefer: "return=representation,resolution=merge-duplicates"
    },
    body: body ? JSON.stringify(body) : undefined
  });
  const txt = await r.text();
  let js; try { js = txt ? JSON.parse(txt) : null; } catch(e) { js = { raw: txt }; }
  if (!r.ok) throw new Error(typeof js==="string"?js:JSON.stringify(js));
  return js;
}
app.post("/api/admin/profile", async (req, res) => {
  try {
    if (!sbOn()) return res.json({ ok: false, error: "no supabase" });
    const d = req.body || {};
    const isAd = String(d.email||"").toLowerCase()==="vanraj2592@gmail.com" || d.app_role==="super_admin" || d.role==="admin";
    const row = {
      email: String(d.email||"").toLowerCase(),
      full_name: d.full_name||d.name||"",
      mobile_number: d.mobile_number||d.mobile||"",
      designation: d.designation||d.role||"AAE",
      department: d.department||"Panchayat",
      office_location: d.office_location||d.taluka||"",
      sub_division: d.sub_division||d.subdiv||"",
      subscription_status: d.subscription_status||d.status||(isAd?"Active":"Trial"),
      role: isAd?"admin":"user"
    };
    if (d.id) row.id = d.id;
    if (d.subscription_end_date) row.subscription_end_date = d.subscription_end_date;
    const js = await sbProfiles("POST", "profiles?on_conflict=email", row);
    res.json({ ok: true, item: Array.isArray(js)?js[0]:js });
  } catch (e) {
    res.json({ ok: false, error: String(e.message||e) });
  }
});
app.get("/api/admin/pending", async (_req, res) => {
  try {
    if (!sbOn()) return res.json({ ok: true, items: [] });
    const js = await sbProfiles("GET", "profiles?select=*&order=email.asc", null);
    res.json({ ok: true, items: Array.isArray(js)?js:[] });
  } catch (e) {
    res.json({ ok: false, items: [], error: String(e.message||e) });
  }
});
app.post("/api/admin/approve", async (req, res) => {
  try {
    const email = String((req.body||{}).email||"").toLowerCase();
    const status = (req.body||{}).status || "Active";
    if (!email) return res.json({ ok: false, error: "email" });
    const js = await sbProfiles("PATCH", "profiles?email=eq." + encodeURIComponent(email), { subscription_status: status==="Approved"?"Active":status });
    res.json({ ok: true, item: Array.isArray(js)?js[0]:js });
  } catch (e) {
    res.json({ ok: false, error: String(e.message||e) });
  }
});



app.post("/api/bill/paver", async (req, res) => {
  try {
    const d = req.body || {};
    const wb = new ExcelJS.Workbook();
    await wb.xlsx.readFile(path.join(__dirname, "skeleton-paver-bill.xlsx"));
    const pa = wb.getWorksheet("21 No. P.A. Form");
    const bill = wb.getWorksheet("Bill");
    const comp = wb.getWorksheet("COMP-14MU NAN");
    if (!pa || !bill || !comp) throw new Error("bill skeleton sheets missing");

    const work = d.work_name || "";
    const grant = d.fund_head || d.grant || "";
    const gam = d.village || d.gam || "";
    let tal = String(d.taluka || "").trim();
    tal = tal.replace(/^તા\.\s*/, "");
    if (/city|district|જિલ્લો|ahmedabad city/i.test(tal) && d.taluka_real) tal = String(d.taluka_real);
    const talLabel = tal ? ("તા. " + tal) : "";
    const agency = d.agency || "સરપંચ શ્રી ગ્રામ પંચાયત";
    const wo = d.work_order || "";
    const asDet = d.as_details || d.as_detail || "";
    const tsDet = d.ts_details || "";
    const tsAmt = Number(d.ts_amount || d.amounting || 0);
    const startDate = d.start_date || "";
    const measDate = d.meas_date || d.date || "";
    const mb = d.mb_no || "";
    const pg1 = d.page_from || "";
    const pg2 = d.page_to || "";
    const subdiv = d.subdiv || d.subdiv_address || "";
    const aae = d.prepared_by || d.aae || "";

    const q = d.qty || {};
    const box = Number(q.box ?? d.box_qty ?? 0);
    const mur = Number(q.murum ?? d.mur_qty ?? 0);
    const pav = Number(q.paver ?? d.area ?? 0);
    const vata = Number(q.vata ?? 0);
    const test = Number(q.test ?? 1);
    const plate = Number(q.plate ?? 0);
    const e4 = box * 156.56;
    const e5 = mur * 210.42;
    const e6 = pav * 740.51;
    const e7 = vata * 23.68;
    const e8 = test * 608;
    const e9 = plate * 306.14;
    const e10 = e4 + e5 + e6 + e7 + e8 + e9;
    const e11 = e10 * 0.18;
    const e12 = e10 + e11;
    const e13 = Math.floor(e12);

    setVal(pa, "H2", grant);
    setVal(pa, "G6", subdiv);
    setVal(pa, "C7", work);
    setVal(pa, "J7", talLabel);
    setVal(pa, "G9", talLabel);
    setVal(pa, "J8", gam);
    setVal(pa, "G8", agency);
    setVal(pa, "I9", wo);

    setVal(bill, "A1", work);
    setVal(bill, "B4", box);
    setVal(bill, "B5", mur);
    setVal(bill, "B6", pav);
    setVal(bill, "B7", vata);
    setVal(bill, "B8", test);
    setVal(bill, "B9", plate);
    setVal(bill, "E4", e4);
    setVal(bill, "E5", e5);
    setVal(bill, "E6", e6);
    setVal(bill, "E7", e7);
    setVal(bill, "E8", e8);
    setVal(bill, "E9", e9);
    setVal(bill, "E10", e10);
    setVal(bill, "E11", e11);
    setVal(bill, "E12", e12);
    setVal(bill, "E13", e13);
    if (aae) setVal(bill, "A14", "શ્રી- " + aae);
    setVal(bill, "B15", measDate);
    setVal(bill, "B16", mb);
    setVal(bill, "D16", pg1);
    setVal(bill, "F16", pg2);

    setVal(comp, "C2", grant);
    setVal(comp, "C3", work);
    setVal(comp, "C4", tsDet);
    setVal(comp, "C5", tsAmt);
    setVal(comp, "C6", asDet);
    setVal(comp, "C7", tsAmt);
    setVal(comp, "C8", agency);
    setVal(comp, "F8", gam);
    setVal(comp, "C9", startDate);
    setVal(comp, "C10", measDate);
    setVal(comp, "C11", e13);
    setVal(comp, "C12", "MB NO -");
    setVal(comp, "D12", mb);
    setVal(comp, "E12", "PAGE NO");
    setVal(comp, "F12", pg1);
    setVal(comp, "G12", "TO");
    setVal(comp, "H12", pg2);
    setVal(comp, "A20", tal || "");

    const areas = {
      "21 No. P.A. Form": "A1:J36",
      "Bill": "A1:G19",
      "COMP-14MU NAN": "A1:H21"
    };
    await writeAndRespond(req, res, wb, Object.assign({}, d, { net: e13 }), "BILL_PAVER", areas, "bill_paver");
  } catch (err) {
    res.status(500).json({ ok: false, error: String(err.message || err) });
  }
});


app.post("/api/bill/cc", async (req, res) => {
  try {
    const d = req.body || {};
    const wb = new ExcelJS.Workbook();
    await wb.xlsx.readFile(path.join(__dirname, "skeleton-cc-bill.xlsx"));
    const pa = wb.getWorksheet("21 No. P.A. Form");
    const bill = wb.getWorksheet("Bill");
    const comp = wb.getWorksheet("COMP-14MU NAN");
    if (!pa || !bill || !comp) throw new Error("cc bill skeleton missing");

    const work = d.work_name || "";
    const grant = d.fund_head || d.grant || "";
    const gam = d.village || d.gam || "";
    let tal = String(d.taluka || "").trim().replace(/^તા\.\s*/, "");
    const talLabel = tal ? ("તા. " + tal) : "";
    const agency = d.agency || "સરપંચ શ્રી ગ્રામ પંચાયત";
    const wo = d.work_order || "";
    const asDet = d.as_details || d.as_detail || "";
    const tsDet = d.ts_details || "";
    const tsAmt = Number(d.ts_amount || d.amounting || 0);
    const startDate = d.start_date || "";
    const measDate = d.meas_date || d.date || "";
    const mb = d.mb_no || "";
    const pg1 = d.page_from || "";
    const pg2 = d.page_to || "";
    const subdiv = d.subdiv || d.subdiv_address || "";
    const aae = d.prepared_by || d.aae || "";

    const q = d.qty || {};
    const q1 = Number(q.box ?? q.q1 ?? 0);
    const q2 = Number(q.metal ?? q.q2 ?? 0);
    const q3 = Number(q.murum ?? q.q3 ?? 0);
    const q4 = Number(q.spread_metal ?? q.q4 ?? q2);
    const q5 = Number(q.spread_mur ?? q.q5 ?? q3);
    const q6 = Number(q.cc ?? q.q6 ?? 0);
    const q7 = Number(q.test ?? q.q7 ?? 1);
    const q8 = Number(q.plate ?? q.q8 ?? 0);
    const rates = [158.12, 684.6, 173.51, 249.75, 147.47, 4915.01, 2656, 306.14];
    const qtys = [q1, q2, q3, q4, q5, q6, q7, q8];
    let sub = 0;
    qtys.forEach(function (qty, i) {
      const amt = qty * rates[i];
      setVal(bill, "B" + (4 + i), qty);
      setVal(bill, "E" + (4 + i), amt);
      sub += amt;
    });
    const gst = sub * 0.18;
    const tot = sub + gst;
    const net = Math.floor(tot);

    setVal(pa, "H2", grant);
    setVal(pa, "G6", subdiv);
    setVal(pa, "C7", work);
    setVal(pa, "J7", talLabel);
    setVal(pa, "G9", talLabel);
    setVal(pa, "J8", gam);
    setVal(pa, "G8", agency);
    setVal(pa, "I9", wo);

    setVal(bill, "A1", work);
    setVal(bill, "E12", sub);
    setVal(bill, "E13", gst);
    setVal(bill, "E14", tot);
    setVal(bill, "E15", net);
    if (aae) setVal(bill, "A16", "શ્રી- " + aae);
    setVal(bill, "B17", measDate);
    setVal(bill, "B18", mb);
    setVal(bill, "D18", pg1);
    setVal(bill, "F18", pg2);

    setVal(comp, "C2", grant);
    setVal(comp, "C3", work);
    setVal(comp, "C4", tsDet);
    setVal(comp, "C5", tsAmt);
    setVal(comp, "C6", asDet);
    setVal(comp, "C7", tsAmt);
    setVal(comp, "C8", agency);
    setVal(comp, "F8", gam);
    setVal(comp, "C9", startDate);
    setVal(comp, "C10", measDate);
    setVal(comp, "C11", net);
    setVal(comp, "D12", mb);
    setVal(comp, "F12", pg1);
    setVal(comp, "H12", pg2);
    setVal(comp, "A20", talLabel || tal || "");

    const areas = {
      "21 No. P.A. Form": "A1:J36",
      "Bill": "A1:G21",
      "COMP-14MU NAN": "A1:H21"
    };
    await writeAndRespond(req, res, wb, Object.assign({}, d, { type: "CC", net: net }), "BILL_CC", areas, "bill_cc");
  } catch (err) {
    res.status(500).json({ ok: false, error: String(err.message || err) });
  }
});


app.post("/api/mb/paver", (req, res) => { req.url = "/api/mb"; req.body = Object.assign({}, req.body||{}, { type: "paver" }); return app._router.handle(req, res, function(){}); });

app.post("/api/mb", async (req, res) => {
  try {
    const d = req.body || {};
    function mix(v) {
      const parts = String(v == null ? "" : v).split("+").map(function (x) { return parseFloat(String(x).trim()); }).filter(function (n) { return !isNaN(n); });
      if (!parts.length) return Number(v) || 0;
      return parts.reduce(function (a, b) { return a + b; }, 0) / parts.length;
    }
    const type = String(d.type || "paver").toLowerCase().indexOf("cc") >= 0 ? "cc" : "paver";
    const rows = Array.isArray(d.rows) ? d.rows : [];
    const files = ["skeleton-mb.xlsx","skeleton-mb-paver-cc.xlsx","skeleton-mb-paver.xlsx"];
    const found = files.map(function(n){ return path.join(__dirname, n); }).filter(function(f){ return require("fs").existsSync(f); });
    if (!found.length) throw new Error("skeleton-mb.xlsx missing on server");
    let file = found[0];
    let wb = new ExcelJS.Workbook();
    await wb.xlsx.readFile(file);
    unshareFormulas(wb);
    function hasCc(wbb){
      return (wbb.worksheets||[]).some(function(w){ return /cc/i.test(String(w.name||"")); });
    }
    if (!hasCc(wb)) {
      for (let i=1;i<found.length;i++){
        const w2 = new ExcelJS.Workbook();
        await w2.xlsx.readFile(found[i]);
        if (hasCc(w2)) { wb = w2; file = found[i]; break; }
      }
    }
    function findWs(keys){
      const all = wb.worksheets || [];
      for (let i=0;i<keys.length;i++){
        const k=keys[i];
        const hit=all.find(function(w){ return String(w.name||"").toLowerCase().indexOf(k)>=0; });
        if (hit) return hit;
      }
      return null;
    }
    const ws = type === "cc"
      ? (wb.getWorksheet("mb-cc road") || findWs(["cc","road"]) || wb.worksheets[1])
      : (findWs(["paver"]) || wb.worksheets[0]);
    if (!ws) throw new Error("mb sheet missing: " + wb.worksheets.map(function(w){return w.name;}).join(", "));

    const segs = [];
    rows.forEach(function (r) {
      const L = mix(r.l || r.L);
      const W = mix(r.w || r.W);
      if (!L && !W) return;
      segs.push({ L, W, d: mix(r.d || r.D) });
    });

    if (type === "paver") {
      const excD = Number(d.exc_d || 0.2);
      const dustD = Number(d.dust_d || 0.1);
      for (let r = 3; r <= 14; r++) {
        setVal(ws, "G" + r, 0);
        setVal(ws, "H" + r, 0);
      }
      setVal(ws, "I3", excD);
      segs.forEach(function (s, i) {
        if (i > 11) return;
        const r = 3 + i;
        setVal(ws, "G" + r, s.L);
        setVal(ws, "H" + r, s.W);
      });
      setVal(ws, "C6", Number(d.test_qty == null ? 1 : d.test_qty));
      setVal(ws, "C7", Number(d.name_plate == null ? 0 : d.name_plate));
      setVal(ws, "E11", Number(d.amounting || 0));
    } else {
      const boxT = Number(d.exc_d || 0);
      const ccT = Number(d.cc_t || 0);
      for (let r = 3; r <= 10; r++) {
        ["G","H","I","J","L","M","N","O","P"].forEach(function (col) {
          const c = ws.getCell(col + r);
          if (c) { c.value = (r === 3 && col === "I") ? boxT : (r === 3 && col === "N") ? ccT : 0; }
        });
      }
      const m2 = ws.getCell("M2");
      if (m2) { m2.numFmt = "@"; m2.value = ""; }
      segs.forEach(function (s, i) {
        if (i > 7) return;
        const r = 3 + i;
        setVal(ws, "G" + r, s.L);
        setVal(ws, "H" + r, s.W);
        setVal(ws, "L" + r, s.L);
        setVal(ws, "M" + r, s.W);
        setVal(ws, "N" + r, ccT);
        setVal(ws, "O" + r, s.L * s.W * ccT);
        setVal(ws, "P" + r, s.L * s.W);
        setVal(ws, "J" + r, s.L * s.W * (s.d || boxT));
        if (i === 0) {
          setVal(ws, "I3", s.d || boxT);
          setVal(ws, "N3", ccT);
        }
      });
      setVal(ws, "C8", Number(d.test_qty == null ? 0 : d.test_qty));
      setVal(ws, "C9", Number(d.name_plate == null ? 0 : d.name_plate));
      setVal(ws, "E13", Number(d.amounting || 0));
    }

    const prefix = type === "cc" ? "MB_CC" : "MB_PAVER";
    const areas = {};
    areas[ws.name] = type === "cc" ? "A1:P20" : "A1:M32";
    wb.worksheets.slice().forEach(function (w) {
      if (w && ws && w.id !== ws.id) {
        try { wb.removeWorksheet(w.id); } catch (_e) { w.state = "hidden"; }
      }
    });
    await writeAndRespond(req, res, wb, Object.assign({}, d, { output: d.output || "xlsx" }), prefix, areas, "bill_mb");
  } catch (err) {
    res.status(500).json({ ok: false, error: String(err.message || err) });
  }
});

const PORT = process.env.PORT || 3000;

app.post("/api/letter/fwd", async (req, res) => {
  try {
    const d = req.body || {};
    const all = Array.isArray(d.items) ? d.items : [];
    const rb = Array.isArray(d.rb) ? d.rb : all.filter(function (x) { return x.office !== "nani"; });
    const nani = Array.isArray(d.nani) ? d.nani : all.filter(function (x) { return x.office === "nani"; });
    if (!all.length) return res.status(400).json({ ok: false, error: "no items" });
    if (!String(d.letter_no || "").trim() || !d.date) {
      return res.status(400).json({ ok: false, error: "vashi and date required" });
    }
    const fileA = path.join(__dirname, "skeleton-paver-frwd-letter.xlsx");
    const fileB = path.join(__dirname, "skeleton-letter-3.xlsx");
    const wb = new ExcelJS.Workbook();
    await wb.xlsx.readFile(require("fs").existsSync(fileA) ? fileA : fileB);
    const wsRb = wb.getWorksheet("DEE-R&B") || wb.getWorksheet("DEE R&B");
    const wsNani = wb.getWorksheet("DEE-Nani Sinchai") || wb.getWorksheet("DEE Nani Sinchai");
    const wsAud = wb.getWorksheet("Audit") || wb.getWorksheet("AANTRIK ODIT nana");
    const taluka = String(d.taluka || (all[0] && all[0].taluka) || "").trim();
    const date = d.date || "";
        const rawNo = String(d.letter_no || "").trim();
    let no = parseInt(fromGujDigits(rawNo).replace(/[^0-9]/g, ""), 10);
    if (!no) no = Number(fromGujDigits(rawNo)) || 0;
    const useGuj = hasGuj(d.letter_no) || /[૦-૯]/.test(String(d.letter_no||"")) || hasGuj(taluka) || hasGuj(d.audit_office) ||
      all.some(function (x) { return hasGuj(x.work_name) || hasGuj(x.taluka); });
        function hasGuj(t) {
      return /[઀-૿]/.test(String(t || ""));
    }
    function toGujDigits(v) {
      const map = "૦૧૨૩૪૫૬૭૮૯";
      return String(v == null ? "" : v).replace(/[0-9]/g, function (d) { return map[Number(d)]; });
    }
    function fromGujDigits(v) {
      return String(v == null ? "" : v).replace(/[૦-૯]/g, function (d) { return "0123456789"["૦૧૨૩૪૫૬૭૮૯".indexOf(d)]; });
    }
    function showVal(v, guj) {
      if (v == null || v === "") return v;
      return guj ? toGujDigits(v) : v;
    }
    function set(ws, addr, v) {
      if (!ws) return;
      ws.getCell(addr).value = v;
    }
    function uniqMb(items) {
      const u = [];
      (items || []).forEach(function (it) {
        const m = String(it.mb_no || "").trim();
        if (m && u.indexOf(m) < 0) u.push(m);
      });
      return u;
    }
    function addWorkRows(ws, templateRow, n) {
      const extra = Math.max(0, Number(n || 0) - 1);
      if (!ws || extra < 1) return 0;
      function colName(num) {
        let s = "";
        let c = num;
        while (c > 0) { const m = (c - 1) % 26; s = String.fromCharCode(65 + m) + s; c = Math.floor((c - 1) / 26); }
        return s;
      }
      function listMerges() {
        const out = [];
        Object.keys(ws._merges || {}).forEach(function (key) {
          const m = ws._merges[key];
          const model = m && (m.model || m);
          if (model && model.top) out.push({ top: model.top, left: model.left, bottom: model.bottom, right: model.right });
        });
        return out;
      }
      function clearMerges() {
        Object.keys(ws._merges || {}).slice().forEach(function (key) {
          try { ws.unMergeCells(key); } catch (_e) {}
        });
      }
      function applyMerge(top, left, bottom, right) {
        try { ws.mergeCells(colName(left) + top + ":" + colName(right) + bottom); } catch (_e) {}
      }
      const saved = listMerges();
      const borderSave = saved.map(function (m) {
        const cells = [];
        for (let r = m.top; r <= m.bottom; r++) {
          for (let c = m.left; c <= m.right; c++) {
            let border = null;
            try { border = JSON.parse(JSON.stringify(ws.getRow(r).getCell(c).border || {})); } catch (_e) {}
            if (border && (border.top || border.left || border.bottom || border.right)) {
              cells.push({ dr: r - m.top, dc: c - m.left, border: border });
            }
          }
        }
        return cells;
      });
      const spans = [];
      saved.forEach(function (m) {
        if (m.top === templateRow && m.bottom === templateRow) spans.push([m.left, m.right]);
      });
      if (!spans.length) {
        spans.push([2, 6]);
        spans.push(templateRow === 20 ? [7, 9] : [7, 8]);
      }
      const src = ws.getRow(templateRow);
      const height = src.height;
      const styles = [];
      for (let c = 1; c <= 9; c++) {
        const cell = src.getCell(c);
        let style = {};
        try { style = JSON.parse(JSON.stringify(cell.style || {})); } catch (_e) {}
        styles.push({ style: style, numFmt: cell.numFmt });
      }
      ws.spliceRows(templateRow + 1, 0, ...new Array(extra).fill([]));
      for (let i = 1; i <= extra; i++) {
        const row = ws.getRow(templateRow + i);
        if (height) row.height = height;
        for (let c = 1; c <= 9; c++) {
          const dst = row.getCell(c);
          dst.value = null;
          dst.style = styles[c - 1].style || {};
          if (styles[c - 1].numFmt) dst.numFmt = styles[c - 1].numFmt;
        }
      }
      clearMerges();
      saved.forEach(function (m, mi) {
        let top = m.top;
        let bottom = m.bottom;
        if (top >= templateRow + 1) { top += extra; bottom += extra; }
        else if (bottom >= templateRow + 1) bottom += extra;
        applyMerge(top, m.left, bottom, m.right);
        (borderSave[mi] || []).forEach(function (b) {
          try { ws.getRow(top + b.dr).getCell(m.left + b.dc).border = b.border; } catch (_e) {}
        });
      });
      for (let i = 1; i <= extra; i++) {
        const r = templateRow + i;
        spans.forEach(function (sp) { applyMerge(r, sp[0], r, sp[1]); });
      }
      return extra;
    }
    function fillDee(ws, items, letterNo) {
      if (!ws) return;
      const first = items[0] || {};
      set(ws, "I1", letterNo);
      set(ws, "I2", taluka);
      set(ws, "H3", date);
      set(ws, "A8", first.office_addr || first.subdiv || d.subdiv || "");
      const list = items || [];
      const extra = addWorkRows(ws, 19, list.length);
      let total = 0;
      list.forEach(function (it, i) {
        const r = 19 + i;
        set(ws, "A" + r, i + 1);
        set(ws, "B" + r, it.work_name || "");
        set(ws, "G" + r, Number(it.amounting || 0));
        total += Number(it.amounting || 0);
      });
      const tr = 19 + list.length;
      set(ws, "A" + tr, "");
      set(ws, "B" + tr, "TOTAL");
      set(ws, "G" + tr, total);
      const mbs = uniqMb(items);
      set(ws, "D" + (28 + extra), mbs[0] || "");
      set(ws, "E" + (28 + extra), mbs[1] || "");
      set(ws, "F" + (28 + extra), mbs[2] || "");
      set(ws, "G" + (28 + extra), mbs[3] || "");
      set(ws, "F" + (33 + extra), taluka);
      ws._fitBottom = 33 + extra;
    }
    function fillAudit(ws, items, letterNo) {
      if (!ws) return;
      set(ws, "I1", letterNo);
      set(ws, "I2", taluka);
      set(ws, "H3", date);
      set(ws, "A8", d.audit_office || "");
      const list = items || [];
      const extra = addWorkRows(ws, 20, list.length);
      let total = 0;
      list.forEach(function (it, i) {
        const r = 20 + i;
        set(ws, "A" + r, i + 1);
        set(ws, "B" + r, it.work_name || "");
        set(ws, "G" + r, Number(it.amounting || 0));
        total += Number(it.amounting || 0);
      });
      const tr = 20 + list.length;
      set(ws, "B" + tr, "TOTAL");
      set(ws, "G" + tr, total);
      set(ws, "B" + (30 + extra), "માપ બુક નંબર");
      set(ws, "D" + (30 + extra), uniqMb(items).join(", "));
      set(ws, "H" + (35 + extra), taluka);
      ws._fitBottom = 35 + extra;
    }
    if (rb.length && wsRb) { fillDee(wsRb, rb, no); no += 1; }
    else if (wsRb) wb.removeWorksheet(wsRb.id);
    if (nani.length && wsNani) { fillDee(wsNani, nani, no); no += 1; }
    else if (wsNani) wb.removeWorksheet(wsNani.id);
    if (wsAud) fillAudit(wsAud, all, no);
    const orderNames = ["DEE-R&B", "DEE R&B", "DEE-Nani Sinchai", "DEE Nani Sinchai", "Audit"];
    if (wb._worksheets) {
      const map = {};
      wb.worksheets.forEach(function (w) { map[w.name] = w; });
      const ordered = [];
      orderNames.forEach(function (n) {
        if (map[n] && ordered.indexOf(map[n]) < 0) ordered.push(map[n]);
      });
      if (ordered.length) wb._worksheets = [undefined].concat(ordered);
    }
    const areaMap = {};
    wb.worksheets.forEach(function (ws) {
      const bottom = ws._fitBottom || (ws.name === "Audit" ? 35 : 33);
      const area = "A1:I" + bottom;
      ws.pageSetup.paperSize = 9;
      ws.pageSetup.orientation = "portrait";
      ws.pageSetup.fitToPage = true;
      ws.pageSetup.fitToWidth = 1;
      ws.pageSetup.fitToHeight = 1;
      ws.pageSetup.printArea = area;
      for (let c = 10; c <= 256; c++) ws.getColumn(c).hidden = true;
      areaMap[ws.name] = area;
    });
    return writeAndRespond(req, res, wb, Object.assign({}, d, { output: d.output || "both", village: taluka || "letter" }), "LETTER", areaMap, "letter_fwd");
  } catch (e) {
    res.status(500).json({ ok: false, error: String(e.message || e) });
  }
});




app.post("/api/estimate/pipe", async (req, res) => {
  try {
    const d = req.body || {};
    const wb = new ExcelJS.Workbook();
    await wb.xlsx.readFile(path.join(__dirname, "skeleton-pipe.xlsx"));
    const face = wb.getWorksheet("Estimate");
    const abs = wb.getWorksheet("Abstract");
    const meas = wb.getWorksheet("Measurement");
    const test = wb.getWorksheet("TEST-SITE");
    if (!face || !abs || !meas) throw new Error("skeleton-pipe.xlsx sheets missing");
    const work = d.work_name || "";
    const taluka = d.taluka || "";
    const village = d.village || "";
    const say = Number(d.amounting || 0);
    const fund = d.fund_head || "";
    const pEx = Array.isArray(d.pEx) ? d.pEx : [];
    const pDe = Array.isArray(d.pDe) ? d.pDe : [];
    const sumDia = (dia) => pEx.filter(r => Number(r.dia)===dia).reduce((a,r)=>a+Number(r.l||0),0);
    const lastWD = (dia, dw, dd) => {
      const arr=pEx.filter(r => Number(r.dia)===dia);
      const r=arr[arr.length-1];
      return {w:Number((r&&r.w)||dw), d:Number((r&&r.d)||dd)};
    };
    setVal(face, "F2", d.division || d.jilla || "");
    setVal(face, "I2", d.jilla || "");
    setVal(face, "F4", d.subdiv_address || "");
    setVal(face, "F5", d.subdiv_address || "");
    setVal(face, "I4", d.nani_address || "");
    setVal(face, "D8", fund);
    setVal(face, "H18", taluka);
    setVal(face, "C20", work);
    setVal(face, "G21", say);
    setVal(face, "D27", d.prepared_by || "");
    setVal(face, "D29", d.prepared_by || "");
    setVal(face, "B34", d.sr_no || "");
    setVal(face, "C34", d.ss_details || "");
    setVal(face, "G40", taluka);
    setVal(abs, "C2", work);
    setVal(meas, "C2", work);
    if (test) setVal(test, "B1", work);
    const demoL = pDe.reduce((a,r)=>a+Number(r.l||0),0);
    const demoW = Number((pDe[0]&&pDe[0].w)||0.45);
    setVal(meas, "E5", demoL);
    setVal(meas, "G5", demoW);
    [[63,"E9","G9","I9",0.45,0.9],[75,"E10","G10","I10",0.45,0.9],[90,"E11","G11","I11",0.45,0.9],[110,"E12","G12","I12",0.45,0.9]].forEach(function(row){
      const wd=lastWD(row[0], row[4], row[5]);
      setVal(meas, row[1], sumDia(row[0]));
      setVal(meas, row[2], wd.w);
      setVal(meas, row[3], wd.d);
    });
    setVal(meas, "E40", Number(d.pPlate||1));
    if (test) setVal(test, "C12", 1);
    await writeAndRespond(req, res, wb, d, "PIPE", {
      Estimate: "A1:I41", Abstract: "A1:F30", Measurement: "A1:K42", "TEST-SITE": "A1:G36"
    }, "estimate_pipe");
  } catch (e) {
    res.status(500).json({ ok: false, error: String(e.message || e) });
  }
});

function fillBillPaComp(pa, bill, comp, d){
  const work = d.work_name || "";
  const grant = d.fund_head || d.grant || "";
  const gam = d.village || d.gam || "";
  let tal = String(d.taluka || "").trim().replace(/^તા\.\s*/, "");
  const talLabel = tal ? ("તા. " + tal) : "";
  const agency = d.agency || "સરપંચ શ્રી ગ્રામ પંચાયત";
  const wo = d.work_order || d.as_details || "";
  const subdiv = d.subdiv || d.subdiv_address || "";
  const aae = d.prepared_by || d.aae || "";
  const measDate = d.meas_date || d.date || "";
  const mb = d.mb_no || "";
  const pg1 = d.page_from || "";
  const pg2 = d.page_to || "";
  if (pa) {
    setVal(pa, "H2", grant);
    setVal(pa, "G6", subdiv);
    setVal(pa, "C7", work);
    setVal(pa, "J7", talLabel);
    setVal(pa, "G9", talLabel);
    setVal(pa, "J8", gam);
    setVal(pa, "G8", agency);
    setVal(pa, "I9", wo);
  }
  if (bill) {
    setVal(bill, "A1", work);
    const mbRow = findLabelRow(bill, /^MB NO/i);
    if (mbRow) {
      setVal(bill, "A" + (mbRow - 2), aae ? ("શ્રી- " + aae) : "");
      setVal(bill, "B" + (mbRow - 1), measDate);
      setVal(bill, "B" + mbRow, mb);
      setVal(bill, "D" + mbRow, pg1);
      setVal(bill, "F" + mbRow, pg2);
    }
  }
  if (comp) {
    setVal(comp, "C2", grant);
    setVal(comp, "C3", work);
    setVal(comp, "C4", d.ts_details || "");
    setVal(comp, "C5", d.ts_amount || d.amounting || "");
    setVal(comp, "C6", d.as_details || d.work_order || "");
    setVal(comp, "C7", d.as_amount || d.amounting || "");
    setVal(comp, "C8", agency);
    setVal(comp, "F8", gam);
    setVal(comp, "C9", d.start_date || "");
    setVal(comp, "C10", measDate);
    setVal(comp, "D12", mb);
    setVal(comp, "F12", pg1);
    setVal(comp, "H12", pg2);
    setVal(comp, "A20", talLabel);
  }
}

function putBillLine(bill, addr, qty, rate) {
  const q = Number(qty || 0);
  const amt = Math.round(q * rate * 100) / 100;
  setVal(bill, addr, q);
  setVal(bill, "E" + String(addr).replace(/^[A-Z]+/, ""), amt);
  return amt;
}
function putBillTotals(bill, sub, row) {
  const gst = Math.round(sub * 0.18 * 100) / 100;
  const tot = Math.round((sub + gst) * 100) / 100;
  const net = Math.floor(sub * 1.18);
  setVal(bill, "E" + row, Math.round(sub * 100) / 100);
  setVal(bill, "E" + (row + 1), gst);
  setVal(bill, "E" + (row + 2), tot);
  setVal(bill, "E" + (row + 3), net);
  return net;
}

app.post("/api/bill/gutter", async (req, res) => {
  try {
    const d = req.body || {};
    const wb = new ExcelJS.Workbook();
    await wb.xlsx.readFile(path.join(__dirname, "skeleton-gutter-bill.xlsx"));
    const pa = wb.getWorksheet("21 No. P.A. Form");
    const bill = wb.getWorksheet("Bill");
    const comp = wb.getWorksheet("COMP-14MU NAN");
    if (!pa || !bill || !comp) throw new Error("gutter bill skeleton missing");
    fillBillPaComp(pa, bill, comp, d);
    const q = d.qty || {};
    const n = (k) => Number(q[k] || 0);
    let sub = 0;
    sub += putBillLine(bill, "B4", n("demo"), 1030.81);
    sub += putBillLine(bill, "B5", n("exc"), 89);
    [[ "B7", 225, 421 ], [ "B8", 300, 672 ], [ "B9", 450, 817 ], [ "B10", 600, 1331 ], [ "B11", 900, 2476 ], [ "B12", 1200, 4121 ]].forEach(function (x) {
      sub += putBillLine(bill, x[0], n("p" + x[1]), x[2]);
    });
    [[ "B14", 225, 88 ], [ "B15", 300, 119 ], [ "B16", 450, 171 ], [ "B17", 600, 228 ], [ "B18", 900, 340 ], [ "B19", 1200, 440 ]].forEach(function (x) {
      sub += putBillLine(bill, x[0], n("p" + x[1]), x[2]);
    });
    [[ "B21", "c60", 5138 ], [ "B22", "c90", 7343 ], [ "B23", "c139", 8882 ], [ "B24", "c1313", 10698 ]].forEach(function (x) {
      sub += putBillLine(bill, x[0], n(x[1]), x[2]);
    });
    sub += putBillLine(bill, "B25", n("refill"), 22);
    const ch = n("frame") || (n("c60") + n("c90") + n("c139") + n("c1313"));
    sub += putBillLine(bill, "B27", ch, 1121);
    sub += putBillLine(bill, "B28", n("cover") || ch, 1173);
    sub += putBillLine(bill, "B29", n("cc"), 3652.31);
    sub += putBillLine(bill, "B30", n("plate"), 306.14);
    const net = putBillTotals(bill, sub, 31);
    if (comp) setVal(comp, "C11", net);
    d.net = net;
    await writeAndRespond(req, res, wb, d, "BILL_GUTTER", {
      "21 No. P.A. Form": "A1:I36", Bill: "A1:G40", "COMP-14MU NAN": "A1:H21"
    }, "bill_gutter");
  } catch (e) {
    res.status(500).json({ ok: false, error: String(e.message || e) });
  }
});

app.post("/api/bill/pipe", async (req, res) => {
  try {
    const d = req.body || {};
    const wb = new ExcelJS.Workbook();
    await wb.xlsx.readFile(path.join(__dirname, "skeleton-pipe-bill.xlsx"));
    const pa = wb.getWorksheet("21 No. P.A. Form");
    const bill = wb.getWorksheet("Bill");
    const comp = wb.getWorksheet("COMP-14MU NAN");
    if (!pa || !bill || !comp) throw new Error("pipe bill skeleton missing");
    fillBillPaComp(pa, bill, comp, d);
    const q = d.qty || {};
    const n = (k) => Number(q[k] || 0);
    let sub = 0;
    sub += putBillLine(bill, "B4", n("demo"), 202.2);
    sub += putBillLine(bill, "B5", n("exc"), 89);
    [[ "B7", 63, 69 ], [ "B8", 75, 96 ], [ "B9", 90, 139 ], [ "B10", 110, 199 ]].forEach(function (x) {
      sub += putBillLine(bill, x[0], n("p" + x[1]), x[2]);
    });
    [[ "B12", 63, 12 ], [ "B13", 75, 15 ], [ "B14", 90, 17 ], [ "B15", 110, 19 ]].forEach(function (x) {
      sub += putBillLine(bill, x[0], n("p" + x[1]), x[2]);
    });
    sub += putBillLine(bill, "B16", n("refill"), 22);
    sub += putBillLine(bill, "B17", n("plate"), 306.14);
    const net = putBillTotals(bill, sub, 18);
    if (comp) setVal(comp, "C11", net);
    d.net = net;
    await writeAndRespond(req, res, wb, d, "BILL_PIPE", {
      "21 No. P.A. Form": "A1:I36", Bill: "A1:G40", "COMP-14MU NAN": "A1:H21"
    }, "bill_pipe");
  } catch (e) {
    res.status(500).json({ ok: false, error: String(e.message || e) });
  }
});


if (require.main === module) {
  resetLoProfile();
  app.listen(PORT, () => {
    console.log("ParaState MVP on " + PORT);
  });
}
module.exports = { bakeFormulaResults };
