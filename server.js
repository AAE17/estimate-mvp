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
app.get("/", (_req, res) => res.sendFile(path.join(__dirname, "index.html")));
app.get("/index.html", (_req, res) => res.sendFile(path.join(__dirname, "index.html")));
const PUBLIC_PAGES = ["login.html", "signup.html", "auth-pages.js", "gujarat-talukas.js", "gj-geo.json", "NotoSansGujarati-Regular.ttf"];
PUBLIC_PAGES.forEach(name => app.get("/" + name, (_req, res) => res.sendFile(path.join(__dirname, name))));
// Auth gates are mounted BEFORE any route, so every /api/db/* handler (and the list endpoints) is covered.
function userActive(req, res, next) { requireUser(req, res, function () { requireActive(req, res, next); }); }
app.use("/api/db", userActive);
app.use("/api/bills", userActive);
app.use("/api/stats", userActive);
app.use("/api/logs", requireAdmin);

function istDay(input) {
  const dt = input ? new Date(input) : new Date();
  const t = dt.getTime();
  const base = isNaN(t) ? Date.now() : t;
  return new Date(base + 330 * 60 * 1000).toISOString().slice(0, 10);
}
function recordDay(r) {
  const d = String((r && r.day) || "").slice(0, 10);
  if (/^\d{4}-\d{2}-\d{2}$/.test(d)) return d;
  if (r && r.ts) return istDay(r.ts);
  return "";
}
const OUT_DIR = path.join(__dirname, "output");
const LOG_FILE = path.join(__dirname, "events.jsonl");
const FONT_CANDIDATES = [
  path.join(__dirname, "fonts", "NotoSansGujarati-Regular.ttf"),
  path.join(__dirname, "NotoSansGujarati-Regular.ttf"),
];
const FONT = FONT_CANDIDATES.find(function (p) { return fs.existsSync(p); }) || FONT_CANDIDATES[0];
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
          const v = cell.value;
          if (v && typeof v === "object" && (v.sharedFormula || v.shareType === "shared")) {
            const f = cell.formula;
            if (f) cell.value = { formula: f };
          }
        } catch (_e) {}
      });
    });
  });
}

function setVal(ws, addr, v) {
  if (!ws) return;
  if (typeof v === "number" && !isFinite(v)) v = 0;
  ws.getCell(addr).value = v;
}
function mixNum(v) {
  if (typeof v === "number") return isFinite(v) ? v : 0;
  const parts = String(v == null ? "" : v).split("+").map(function (x) { return parseFloat(String(x).trim()); }).filter(function (n) { return !isNaN(n); });
  if (!parts.length) return 0;
  return parts.reduce(function (a, b) { return a + b; }, 0) / parts.length;
}
function wdOf(list, dw, dd) {
  const arr = (list || []).filter(function (r) { return mixNum(r.l) > 0; });
  const L = arr.reduce(function (a, r) { return a + mixNum(r.l); }, 0);
  if (!L) {
    const r = (list || [])[(list || []).length - 1];
    return { l: 0, w: Number((r && r.w) || dw), d: Number((r && r.d) || dd) };
  }
  const LW = arr.reduce(function (a, r) { return a + mixNum(r.l) * Number(r.w || dw); }, 0);
  const LWD = arr.reduce(function (a, r) { return a + mixNum(r.l) * Number(r.w || dw) * Number(r.d || dd); }, 0);
  return { l: L, w: LW / L, d: LW ? LWD / LW : Number(dd) };
}
function billDateText(v) {
  const s = String(v || "").trim();
  const m = s.match(/^(\d{4})-(\d{2})-(\d{2})/);
  if (m) return m[3] + "/" + m[2] + "/" + m[1];
  return s;
}
function fitPaLabels(pa) {
  if (!pa) return;
  [6, 8, 10, 14, 21, 24, 26, 27, 31].forEach(function (r) {
    try { pa.mergeCells("A" + r + ":F" + r); } catch (_e) {
      // template has a smaller merge here (paver A6:B6 + text in C6): join the texts, then merge A:F
      const parts = [];
      for (let c = 1; c <= 6; c++) {
        const t = String(cellText(pa.getCell(r, c)) || "").trim();
        if (t && parts.indexOf(t) < 0) parts.push(t);
      }
      try {
        pa.unMergeCells("A" + r + ":F" + r);
        for (let c = 2; c <= 6; c++) pa.getCell(r, c).value = null;
        pa.getCell("A" + r).value = parts.join(" ");
        pa.mergeCells("A" + r + ":F" + r);
      } catch (_e2) {}
    }
    const cell = pa.getCell("A" + r);
    cell.alignment = { horizontal: "left", vertical: "center", wrapText: true };
  });
  [6, 8, 10].forEach(function (r) {
    const row = pa.getRow(r);
    if (!row.height || row.height < 36) row.height = 36;
  });
  try { pa.unMergeCells("G6:J6"); pa.mergeCells("G6:J6"); } catch (_e) {}
  pa.getCell("G6").alignment = { horizontal: "left", vertical: "top", wrapText: true };
  if (String(cellText(pa.getCell("G6")) || "").length > 30) pa.getRow(6).height = 54;
  for (let r = 7; r <= 11; r++) pa.getCell("Q" + r).value = null;
  for (let c = 11; c <= 20; c++) pa.getColumn(c).hidden = true;
}
function fitDeeName(ws) {
  if (!ws) return;
  [6, 7, 8].forEach(function (r) {
    try { ws.mergeCells("A" + r + ":E" + r); } catch (_e) {}
    const cell = ws.getCell("A" + r);
    cell.alignment = { horizontal: "left", vertical: "center", wrapText: true };
  });
  if (!ws.getColumn(9).width || ws.getColumn(9).width < 18) ws.getColumn(9).width = 18;
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
        if (cell.isMerged && cell.master && cell.master.address !== cell.address) return;
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
    if (cell.isMerged && cell.master && cell.master.address !== cell.address) return 0;
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
      if ((upper === "SUM" || upper === "ROUND" || upper === "ROUNDUP" || upper === "ROUNDDOWN" || upper === "TRUNC") && s[k] === "(") {
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
        if (upper === "ROUND") {
          const sign = x < 0 ? -1 : 1;
          return (sign * Math.round(Math.abs(x))) / p;
        }
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
  let out = "";
  if (v == null || v === "") return "";
  if (typeof v === "object") {
    if (v.result != null && v.result !== "") out = v.result;
    else if (Array.isArray(v.richText)) return v.richText.map((t) => t.text).join("");
    else if (v.text) return String(v.text);
    else if (v.hyperlink) return String(v.text || v.hyperlink);
    else return "";
  } else out = v;
  if (typeof out === "number" && isFinite(out)) {
    const fmt = String((cell && cell.numFmt) || "");
    if (fmt === "0.00" || fmt === "#,##0.00" || fmt.indexOf("0.00") === 0) return (Math.round(out * 100) / 100).toFixed(2);
    return tidyNum(out);
  }
  return String(out);
}

function tidyNum(n) {
  const c2 = Math.round(n * 100) / 100;
  if (Math.abs(n - c2) < 1e-6) {
    if (Math.abs(c2 - Math.round(c2)) < 1e-9) return String(Math.round(c2));
    return c2.toFixed(2);
  }
  const c4 = Math.round(n * 10000) / 10000;
  if (Math.abs(n - c4) < 1e-6) return String(c4);
  return String(c2);
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
    ws.pageSetup.orientation = lastC > 12 ? "landscape" : "portrait";
    ws.pageSetup.fitToPage = true;
    ws.pageSetup.fitToWidth = 1;
    ws.pageSetup.fitToHeight = 1;
    delete ws.pageSetup.scale;
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

async function printEndByFile(zip) {
  const out = {};
  const wbFile = zip.file("xl/workbook.xml");
  const relFile = zip.file("xl/_rels/workbook.xml.rels");
  if (!wbFile || !relFile) return out;
  const wbXml = await wbFile.async("string");
  const relXml = await relFile.async("string");
  const rels = {};
  const rre = /<Relationship\b([^>]*)\/?>/g;
  let m;
  while ((m = rre.exec(relXml))) {
    const id = (m[1].match(/\bId="([^"]+)"/) || [])[1];
    const target = (m[1].match(/\bTarget="([^"]+)"/) || [])[1];
    const type = (m[1].match(/\bType="([^"]+)"/) || [])[1] || "";
    if (!id || !target || type.indexOf("worksheet") < 0) continue;
    let t = String(target).replace(/\\/g, "/");
    if (t.indexOf("xl/") !== 0) t = "xl/" + t.replace(/^\//, "");
    rels[id] = t;
  }
  const order = [];
  const sre = /<sheet\b([^>]*)\/?>/g;
  while ((m = sre.exec(wbXml))) {
    order.push((m[1].match(/\br:id="([^"]+)"/) || [])[1] || "");
  }
  const ends = {};
  const dre = /<definedName\b([^>]*)>([^<]*)<\/definedName>/g;
  while ((m = dre.exec(wbXml))) {
    if (!/name="_xlnm\.Print_Area"/.test(m[1])) continue;
    const sid = (m[1].match(/\blocalSheetId="(\d+)"/) || [])[1];
    if (sid == null) continue;
    let end = 0;
    String(m[2]).split(",").forEach(function (ref) {
      const cols = ref.toUpperCase().replace(/\$/g, "").match(/[A-Z]{1,3}(?=\d)/g);
      if (!cols || !cols.length) return;
      end = Math.max(end, colLetterToNum(cols[cols.length - 1]));
    });
    if (end) ends[sid] = end;
  }
  order.forEach(function (rid, i) {
    const file = rels[rid];
    if (file && ends[String(i)]) out[file] = ends[String(i)];
  });
  return out;
}

async function patchFitXml(xlsxPath) {
  let JSZip;
  try {
    JSZip = require("jszip");
  } catch (_e1) {
    try { JSZip = require("exceljs/node_modules/jszip"); } catch (_e2) { return; }
  }
  const zip = await JSZip.loadAsync(fs.readFileSync(xlsxPath));
  const printEnds = await printEndByFile(zip);
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
    const endC = printEnds[name] || 0;
    xml = xml.replace(/<pageSetup\b([^>]*)\/>/, (_all, attrs) => {
      let a = String(attrs)
        .replace(/\s+scale="[^"]*"/g, "")
        .replace(/\s+orientation="[^"]*"/g, "")
        .replace(/\s+horizontalDpi="[^"]*"/g, "")
        .replace(/\s+verticalDpi="[^"]*"/g, "")
        .replace(/\s+fitToWidth="[^"]*"/g, "")
        .replace(/\s+fitToHeight="[^"]*"/g, "")
        .replace(/\s+paperSize="[^"]*"/g, "")
        .replace(/\s+r:id="[^"]*"/g, "");
      const hadLandscape = /orientation="landscape"/.test(String(attrs));
      const landscape = endC ? endC > 12 : hadLandscape;
      return `<pageSetup${a} paperSize="9" orientation="${landscape ? "landscape" : "portrait"}" fitToWidth="1" fitToHeight="1" horizontalDpi="300" verticalDpi="300"/>`;
    });
    xml = xml.replace(/<pageMargins[^/]*\/>/, '<pageMargins left="0.5" right="0.5" top="0.5" bottom="0.5" header="0.25" footer="0.25"/>');
    if (!/<pageMargins /.test(xml)) {
      xml = xml.replace(/<pageSetup /, '<pageMargins left="0.5" right="0.5" top="0.5" bottom="0.5" header="0.25" footer="0.25"/><pageSetup ');
    }
    zip.file(name, xml);
  }
  const relFiles = Object.keys(zip.files).filter((n) => /xl\/worksheets\/_rels\/sheet\d+\.xml\.rels$/.test(n));
  for (const rn of relFiles) {
    const relFile = zip.file(rn);
    if (!relFile) continue;
    let rel = await relFile.async("string");
    rel = rel.replace(/<Relationship\b[^>]*printerSettings[^>]*\/>/g, "");
    zip.file(rn, rel);
  }
  Object.keys(zip.files).filter((n) => n.indexOf("printerSettings/") >= 0).forEach((n) => {
    zip.remove(n);
  });
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
  const list = ["/usr/bin/soffice", "/usr/bin/libreoffice", "/usr/lib/libreoffice/program/soffice.bin"];
  for (const b of list) {
    try { if (fs.existsSync(b)) return b; } catch (_e) {}
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
  [path.join(LO_DIR, ".lock"), path.join(LO_DIR, "user", ".lock")].forEach(function (p) {
    try { fs.rmSync(p, { force: true }); } catch (_e) {}
  });
}

function runSoffice(args, timeoutMs) {
  return new Promise((resolve, reject) => {
    let done = false;
    const child = spawn(sofficeBin(), args, { stdio: "ignore", detached: true });
    const timer = setTimeout(() => {
      if (done) return;
      done = true;
      try { process.kill(-child.pid, "SIGKILL"); } catch (_e) { try { child.kill("SIGKILL"); } catch (_e2) {} }
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
  if (process.env.SOFFICE_DISABLED === "1") return Promise.resolve(); // tests/CI only: skip LibreOffice recalc
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
      const profile = "/tmp/lo-" + Date.now();
      fs.mkdirSync(profile, { recursive: true });
      try { if (fs.existsSync(pdfPath)) fs.unlinkSync(pdfPath); } catch (_e) {}
      let log = "";
      const child = spawn(sofficeBin(), [
        "-env:UserInstallation=file://" + profile,
        "--headless", "--norestore", "--nolockcheck", "--nologo",
        "--convert-to", "pdf", "--outdir", dir, xlsxPath
      ], {
        detached: true,
        stdio: ["ignore", "pipe", "pipe"],
        env: Object.assign({}, process.env, { SAL_USE_VCLPLUGIN: "svp", HOME: "/tmp" })
      });
      if (child.stdout) child.stdout.on("data", (b) => { log += b.toString(); });
      if (child.stderr) child.stderr.on("data", (b) => { log += b.toString(); });
      let done = false;
      const ready = () => {
        try { return fs.existsSync(pdfPath) && fs.statSync(pdfPath).size > 800; } catch (_e) { return false; }
      };
      const finish = (err) => {
        if (done) return;
        if (!ready() && !err) return;
        done = true;
        clearTimeout(timer);
        clearInterval(poll);
        try { fs.rmSync(profile, { recursive: true, force: true }); } catch (_e) {}
        if (ready()) return resolve(pdfPath);
        try { process.kill(-child.pid, "SIGKILL"); } catch (_e) { try { child.kill("SIGKILL"); } catch (_e2) {} }
        const extra = log.replace(/\s+/g, " ").trim().slice(0, 160);
        reject(new Error((err && err.message ? err.message : "pdf missing") + (extra ? " | " + extra : "")));
      };
      const poll = setInterval(() => { if (ready()) finish(null); }, 500);
      const timer = setTimeout(() => finish(new Error("soffice timeout")), 100000);
      child.on("error", (err) => finish(err));
      child.on("close", () => setTimeout(() => { if (ready()) finish(null); }, 1500));
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

function sheetMerges(ws) {
  const skip = {};
  const span = {};
  function add(top, left, bottom, right) {
    top = Number(top); left = Number(left); bottom = Number(bottom); right = Number(right);
    if (!top || !left || !bottom || !right) return;
    span[top + "," + left] = { c: Math.max(1, right - left + 1), r: Math.max(1, bottom - top + 1) };
    for (let r = top; r <= bottom; r++) {
      for (let c = left; c <= right; c++) {
        if (r === top && c === left) continue;
        skip[r + "," + c] = 1;
      }
    }
  }
  Object.keys(ws._merges || {}).forEach(function (key) {
    const m = ws._merges[key];
    const model = m && (m.model || m);
    if (model && model.top) add(model.top, model.left, model.bottom, model.right);
    else {
      const mm = String(key).match(/^([A-Z]+)(\d+):([A-Z]+)(\d+)$/i);
      if (mm) add(Number(mm[2]), colLetterToNum(mm[1]), Number(mm[4]), colLetterToNum(mm[3]));
    }
  });
  const list = (ws.model && ws.model.merges) || [];
  list.forEach(function (key) {
    const mm = String(key).match(/^([A-Z]+)(\d+):([A-Z]+)(\d+)$/i);
    if (mm) add(Number(mm[2]), colLetterToNum(mm[1]), Number(mm[4]), colLetterToNum(mm[3]));
  });
  return { skip: skip, span: span };
}

function drawPlain(doc, text, x, y, opt, size) {
  doc.font("Helvetica");
  doc.fontSize(size || 9);
  doc.fillColor("#14211A");
  doc.text(String(text == null ? "" : text), x, y, Object.assign({}, opt, { continued: false }));
}

function drawSafe(doc, text, x, y, opt, size, hasFont) {
  try {
    drawMixed(doc, text, x, y, opt, size, hasFont);
  } catch (_e) {
    try { drawPlain(doc, text, x, y, opt, size); } catch (_e2) {}
  }
}

function drawMixed(doc, text, x, y, opt, size, hasFont) {
  const runs = String(text == null ? "" : text).match(/[\u0A80-\u0AFF]+|[A-Za-z]+|[^A-Za-z\u0A80-\u0AFF]+/g) || [String(text || "")];
  function use(g) {
    doc.font(g && hasFont ? "Gu" : "Helvetica");
    doc.fontSize(size || 9);
    doc.fillColor("#14211A");
  }
  if (runs.length === 1) {
    use(/[\u0A80-\u0AFF]/.test(runs[0]));
    doc.text(runs[0], x, y, opt);
    return;
  }
  runs.forEach(function (run, i) {
    use(/[\u0A80-\u0AFF]/.test(run));
    const o = Object.assign({}, opt, { continued: i < runs.length - 1 });
    if (i === 0) doc.text(run, x, y, o);
    else doc.text(run, o);
  });
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
      const area = String((ws.pageSetup && ws.pageSetup.printArea) || "");
      const am = area.match(/([A-Z]+)(\d+):([A-Z]+)(\d+)/);
      const capC = am ? colLetterToNum(am[3]) : 40;
      const capR = am ? Number(am[4]) : 80;
      const landscape = am ? capC > 12 : (ws.pageSetup && ws.pageSetup.orientation) === "landscape";
      doc.addPage({ size: "A4", layout: landscape ? "landscape" : "portrait", margin: 18 });
      const pageW = doc.page.width - 36;
      const pageH = doc.page.height - 40;

      const merges = sheetMerges(ws);
      let maxC = 0;
      const grid = [];
      const rows = [];
      ws.eachRow({ includeEmpty: false }, (row, r) => {
        if (r > capR) return;
        if (row.hidden || row.height === 0) return;
        rows.push(r);
        row.eachCell({ includeEmpty: false }, (cell, c) => {
          if (c > capC) return;
          if (merges.skip[r + "," + c]) {
            if (c > maxC) maxC = c;
            return;
          }
          if (c > maxC) maxC = c;
          if (!grid[r]) grid[r] = [];
          grid[r][c] = cellText(cell);
        });
      });
      const nRows = rows.length || 1;
      if (maxC < 1) maxC = 1;

      const top = 26;
      const rowH = (pageH - 8) / nRows;
      const widths = [];
      let sumW = 0;
      for (let c = 1; c <= maxC; c++) {
        let w = 10;
        try {
          const cw = Number(ws.getColumn(c).width);
          if (cw > 0) w = cw;
        } catch (_e) {}
        widths[c] = w;
        sumW += w;
      }
      if (!sumW) sumW = maxC * 10;
      const fontSize = Math.max(7, Math.min(11, rowH - 2));
      const xs = [];
      let xCursor = 18;
      for (let c = 1; c <= maxC; c++) {
        xs[c] = xCursor;
        xCursor += pageW * (widths[c] / sumW);
      }
      function spanW(c, n) {
        let w = 0;
        for (let i = 0; i < n; i++) w += pageW * ((widths[c + i] || 10) / sumW);
        return Math.max(8, w - 2);
      }
      drawSafe(doc, ws.name, 18, 10, { width: pageW, height: 14 }, fontSize, hasFont);

      doc.fillColor("#14211A");
      rows.forEach(function (r, i) {
        const y = top + i * rowH;
        if (y > top + pageH - 6) return;
        for (let c = 1; c <= maxC; c++) {
          if (merges.skip[r + "," + c]) continue;
          const t = (grid[r] && grid[r][c]) || "";
          if (!t) continue;
          const sp = merges.span[r + "," + c] || { c: 1, r: 1 };
          drawSafe(doc, t, xs[c], y, {
            width: spanW(c, sp.c),
            height: Math.max(4, rowH * sp.r - 0.5),
            ellipsis: true,
            lineBreak: false
          }, fontSize, hasFont);
        }
      });
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
  pdfFromXlsx(xlsxFull).then(function (produced) {
    const pdfFull = path.join(OUT_DIR, pdfName);
    if (produced !== pdfFull && fs.existsSync(produced)) fs.copyFileSync(produced, pdfFull);
    if (!fs.existsSync(pdfFull)) throw new Error("pdf missing");
    pdfJobs[id] = { status: "ok", pdf: "/api/download/" + pdfName };
  }).catch(function (e) {
    pdfJobs[id] = { status: "fail", error: String(e.message || e) };
  });
  return id;
}

async function pdfFromXlsx(xlsxFull) {
  try {
    return await convertWithSoffice(xlsxFull);
  } catch (e) {
    const msg = String(e && e.message || e);
    if (!/ENOENT|not found|spawn/i.test(msg)) throw e;
    const wb = new ExcelJS.Workbook();
    await wb.xlsx.readFile(xlsxFull);
    return sheetsToPdf(wb, xlsxFull.replace(/\.xlsx$/i, ".pdf"));
  }
}

function billMoneyFormat(ws) {
  if (!ws) return;
  for (let r = 1; r <= 60; r++) {
    const c = ws.getCell("E" + r);
    const f = String(c.formula || (c.value && c.value.formula) || "");
    if (/^ROUND\(/i.test(f) && (!c.numFmt || c.numFmt === "0" || c.numFmt === "General")) c.numFmt = "0.00";
  }
  if ((ws.getColumn(5).width || 0) < 13) ws.getColumn(5).width = 13; // room for 10,00,000.00+
}
async function writeAndRespond(req, res, wb, d, prefix, areas, kind) {
  const output = d.output || "xlsx";
  if (/^bill_/.test(String(kind || ""))) billMoneyFormat(wb.getWorksheet("Bill"));
  if (wb.calcProperties) wb.calcProperties.fullCalcOnLoad = true;
  try { unshareFormulas(wb); } catch (_e) {}
  try { bakeFormulaResults(wb); } catch (_e) {}
  dropBadNames(wb);
  unlockSheets(wb);
  applyOnePage(wb, areas);
  const safe = String(d.village || "gam").replace(/[^a-zA-Z0-9._-]+/g, "_");
  const stamp = Date.now();
  const xlsxName = `${prefix}_${safe}_${stamp}.xlsx`;
  const pdfName = `${prefix}_${safe}_${stamp}.pdf`;
  const xlsxFull = path.join(OUT_DIR, xlsxName);
  const pdfFull = path.join(OUT_DIR, pdfName);
  wb.creator = "";
  wb.lastModifiedBy = "";
  await wb.xlsx.writeFile(xlsxFull);
  await patchFitXml(xlsxFull);
  const wantXlsx = output === "xlsx" || output === "both";
  const wantPdf = output === "pdf" || output === "both";
  let pdfJob = null;
  if (wantPdf) pdfJob = pdfName.replace(/\.pdf$/i, "");
  if (output === "xlsx") {
    await recalcXlsxFile(xlsxFull);
    await patchFitXml(xlsxFull);
    persistOut(xlsxFull);
  }
  if (wantPdf) {
    pdfJobs[pdfJob] = { status: "run" };
    setImmediate(function () {
      const prep = output === "both" ? recalcXlsxFile(xlsxFull).then(function () { return patchFitXml(xlsxFull); }).then(function () { persistOut(xlsxFull); }) : Promise.resolve();
      prep.then(function () { return pdfFromXlsx(xlsxFull); }).then(function (produced) {
        const pdfFull = path.join(OUT_DIR, pdfName);
        if (produced !== pdfFull && fs.existsSync(produced)) fs.copyFileSync(produced, pdfFull);
        if (!fs.existsSync(pdfFull)) throw new Error("pdf missing");
        persistOut(pdfFull);
        pdfJobs[pdfJob] = { status: "ok", pdf: "/api/download/" + pdfName };
      }).catch(function (e) {
        pdfJobs[pdfJob] = { status: "fail", error: String(e.message || e) };
      });
    });
  }

  const Lm = Number(d.length_m || 0);
  const Wm = Number(d.width_m || 0);
  const areaM = Lm * Wm;
  const pdfUrl = null;
  const who = await callerEmail(req);
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
    pdf: pdfUrl,
    user_email: who
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
        mb_no: d.mb_no || "",
        user_email: who
      });
      if (typeof sbOn === "function" && sbOn()) {
        sbInsert(isBill ? "bills" : "estimates", rec).catch(function (e) { console.error(e.message); });
      }
    }
  } catch (_e) {}
  const wantXlsxOut = output === "xlsx" || output === "both";
  res.json({
    ok: true,
    xlsx: wantXlsxOut ? `/api/download/${xlsxName}` : null,
    pdf: null,
    pdf_job: pdfJob,
    pdf_error: null,
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


app.get("/api/stats", async (req, res) => {
  const empty = { total: 0, cc: 0, paver: 0, amount: 0, today: 0, talukas: [], recent: [] };
  const who = await callerEmail(req);
  if (!who) return res.json(empty);
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
    if (!ownsRow({ user_email: pl.user_email }, who)) continue;
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


function dedupeBills(rows) {
  const seen = {};
  return (rows || []).slice().sort(function (a, b) { return String(b.ts || "").localeCompare(String(a.ts || "")); })
    .filter(function (r) {
      const mb = String(r.mb_no || "").trim();
      if (!mb) return true;
      const k = [String(r.type || "").toLowerCase(), String(r.work_name || "").trim().toLowerCase(), mb].join("|");
      if (seen[k]) return false;
      seen[k] = 1;
      return true;
    });
}
app.get("/api/bills", async (req, res) => {
  try {
    const day = String(req.query.day || req.query.date || "").slice(0, 10);
    const localBills = notDeleted(dbRead(DB_BILL, 500));
    let remote = [];
    if (sbOn()) {
      try { remote = notDeleted(await sbSelect("bills", 500)); } catch (_e) { remote = []; }
      remote.forEach(function (r) {
        const loc = localBills.find(function (x) { return String(x.id) === String(r.id); });
        if (loc) {
          if (!r.day && loc.day) r.day = loc.day;
          if (r.net == null && loc.net != null) r.net = loc.net;
          if (!r.kind && loc.kind) r.kind = loc.kind;
        }
      });
    }
    let rows = mineMerged(remote, localBills, await callerEmail(req));
    if (day) rows = rows.filter(function (r) { return recordDay(r) === day; });
    rows = dedupeBills(rows);
    res.json({ ok: true, bills: rows.slice(0, 80) });
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

    unshareFormulas(wb);
    const r2 = (n) => Math.round(Number(n || 0) * 100) / 100;
    const rawSegs = Array.isArray(d.segs) && d.segs.length ? d.segs : [{ l: L, w: W, exc: boxT, d: ccT }];
    const lines = rawSegs.map(function (sg) {
      return { l: mixNum(sg.l), w: mixNum(sg.w), t: Number(sg.exc || boxT || 0.3), cc: Number(ccT || sg.d || 0.1) };
    }).filter(function (ln) { return ln.l > 0 && ln.w > 0; });
    if (!lines.length) lines.push({ l: L || 0, w: W || 0, t: boxT || 0.3, cc: ccT || 0.1 });
    const extra = Math.max(0, lines.length - 1);
    if (extra) {
      meas.spliceRows(5, 0, ...new Array(extra).fill([]));
      for (let r = 20; r <= 22 + extra; r++) meas.getRow(r).hidden = r >= 20 + extra;
      const src4 = meas.getRow(4);
      for (let i = 1; i <= extra; i++) {
        const row = meas.getRow(4 + i);
        if (src4.height) row.height = src4.height;
        for (let c = 1; c <= 12; c++) {
          try { row.getCell(c).style = JSON.parse(JSON.stringify(src4.getCell(c).style || {})); } catch (_e) {}
        }
      }
    }
    lines.forEach(function (ln, i) {
      const r = 4 + i;
      setVal(meas, "C" + r, 1);
      setVal(meas, "D" + r, "x");
      setVal(meas, "E" + r, ln.l);
      setVal(meas, "F" + r, "x");
      setVal(meas, "G" + r, ln.w);
      setVal(meas, "H" + r, "x");
      setVal(meas, "I" + r, ln.t);
      setVal(meas, "J" + r, "=");
      meas.getCell("K" + r).value = { formula: "C" + r + "*E" + r + "*G" + r + "*I" + r };
    });
    const boxTot = 4 + lines.length;
    meas.getCell("K" + boxTot).value = { formula: "SUM(K4:K" + (boxTot - 1) + ")" };
    const areaParts = lines.map(function (_ln, i) { return "E" + (4 + i) + "*G" + (4 + i); }).join("+");
    const ccParts = lines.map(function (ln, i) { return "E" + (4 + i) + "*G" + (4 + i) + "*" + Number(ln.cc || ccT || 0.1); }).join("+");
    const metalRow = 7 + extra;
    const voidRow = 8 + extra;
    const metalTotRow = 9 + extra;
    const murRow = 12 + extra;
    const spreadMetalRow = 15 + extra;
    const spreadMurRow = 18 + extra;
    const rollRow = 21 + extra;
    const ccRow = 24 + extra;
    const plateRow = 30 + extra;
    meas.getCell("E" + metalRow).value = { formula: areaParts };
    meas.getCell("G" + metalRow).value = 1;
    meas.getCell("I" + metalRow).value = btT;
    meas.getCell("K" + metalRow).value = { formula: "(" + areaParts + ")*" + btT };
    meas.getCell("E" + voidRow).value = { formula: "K" + metalRow };
    meas.getCell("G" + voidRow).value = voids;
    meas.getCell("K" + voidRow).value = { formula: "E" + voidRow + "*G" + voidRow };
    meas.getCell("K" + metalTotRow).value = { formula: "K" + metalRow + "+K" + voidRow };
    meas.getCell("G" + murRow).value = { formula: "K" + metalTotRow };
    meas.getCell("I" + murRow).value = murPct;
    meas.getCell("K" + murRow).value = { formula: "G" + murRow + "*I" + murRow + "/100" };
    meas.getCell("K" + spreadMetalRow).value = { formula: "K" + metalTotRow };
    meas.getCell("K" + spreadMurRow).value = { formula: "K" + murRow };
    meas.getCell("E" + rollRow).value = { formula: areaParts };
    meas.getCell("G" + rollRow).value = 1;
    meas.getCell("K" + rollRow).value = { formula: areaParts };
    meas.getCell("E" + ccRow).value = { formula: areaParts };
    meas.getCell("G" + ccRow).value = 1;
    meas.getCell("I" + ccRow).value = Number(lines[0].cc || ccT || 0.1);
    meas.getCell("K" + ccRow).value = { formula: ccParts };
    meas.getCell("K" + plateRow).value = 1;
    setVal(meas, "C1", work);
    if (abs) {
      setVal(abs, "B2", work);
      abs.getCell("A4").value = { formula: "Measurement!K" + boxTot };
      abs.getCell("C4").value = { formula: "Measurement!A3" };
      abs.getCell("A6").value = { formula: "Measurement!K" + metalTotRow };
      abs.getCell("B6").value = "Cmt";
      abs.getCell("C6").value = { formula: "Measurement!A" + (6 + extra) };
      abs.getCell("A8").value = { formula: "Measurement!K" + murRow };
      abs.getCell("C8").value = { formula: "Measurement!A" + (11 + extra) };
      abs.getCell("A10").value = { formula: "Measurement!K" + spreadMetalRow };
      abs.getCell("C10").value = { formula: "Measurement!A" + (14 + extra) };
      abs.getCell("A12").value = { formula: "Measurement!K" + spreadMurRow };
      abs.getCell("C12").value = { formula: "Measurement!A" + (17 + extra) };
      abs.getCell("A14").value = { formula: "Measurement!K" + ccRow };
      abs.getCell("C14").value = { formula: "Measurement!A" + (23 + extra) };
      abs.getCell("F4").value = { formula: "ROUND(A4*D5,2)" };
      abs.getCell("F6").value = { formula: "ROUND(A6*D7,2)" };
      abs.getCell("F8").value = { formula: "ROUND(A8*D9,2)" };
      abs.getCell("F10").value = { formula: "ROUND(A10*D11,2)" };
      abs.getCell("F12").value = { formula: "ROUND(A12*D13,2)" };
      abs.getCell("F14").value = { formula: "ROUND(A14*D15,2)" };
      setVal(abs, "A16", 1);
      abs.getCell("F16").value = { formula: "ROUND(A16*D16,2)" };
      setVal(abs, "A18", 1);
      abs.getCell("F18").value = { formula: "ROUND(A18*D19,2)" };
      abs.getCell("F20").value = { formula: "ROUND(F4+F6+F8+F10+F12+F14+F16+F18,2)" };
      abs.getCell("F21").value = { formula: "ROUND(F20*0.18,2)" };
      abs.getCell("F22").value = { formula: "ROUND(F20+F21,2)" };
      abs.getCell("F23").value = { formula: "ROUND(F22,0)" };
      face.getCell("G22").value = { formula: "Abstract!F23" };
      ["C5", "C19"].forEach(function (a) {
        const v = abs.getCell(a).value;
        if (typeof v === "string") abs.getCell(a).value = v.replace(/^ITEM NO\. 0\s+/i, "");
      });
      setVal(abs, "G6", null);
      abs.getCell("A30").value = taluka;
    }
    if (ra) {
      setVal(ra, "C1", work);
      ra.getCell("A3").value = { formula: "Measurement!A" + (23 + extra) };
      ra.getCell("D38").value = taluka;
    }
    if (sch) {
      setVal(sch, "C1", work);
      sch.getCell("C18").value = taluka;
      sch.getCell("F4").value = { formula: "C4" };
    }
    setVal(lead, "B1", work);
    if (Number(d.lead_sevaliya_to_taluka_km) > 0) setVal(lead, "D5", Number(d.lead_sevaliya_to_taluka_km));
    if (Number(d.lead_taluka_to_site_km) > 0) setVal(lead, "D6", Number(d.lead_taluka_to_site_km));
    lead.getCell("C5").value = taluka;
    lead.getCell("A6").value = taluka;
    lead.getCell("B38").value = taluka;
    if (Number(W) > 0) lead.getCell("F18").value = "W= " + Number(W).toFixed(2) + " MT";
    if (Number(ccT) > 0) lead.getCell("D20").value = "CC " + Math.round(Number(ccT) * 100) + " CM THICK";

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

    const segsP = (Array.isArray(d.segs) ? d.segs : []).map(function (s) {
      return { l: mixNum(s.l), w: mixNum(s.w), t: Number(s.exc || d.box_thick_m || 0.2) };
    }).filter(function (s) { return s.l > 0 && s.w > 0; });
    const L = segsP.length ? segsP.reduce(function (a, s) { return a + s.l; }, 0) : Number(d.length_m);
    const area = segsP.length ? segsP.reduce(function (a, s) { return a + s.l * s.w; }, 0) : Number(d.length_m) * Number(d.width_m);
    const W = L ? area / L : Number(d.width_m);
    const boxQty = segsP.length ? segsP.reduce(function (a, s) { return a + s.l * s.w * s.t; }, 0) : area * Number(d.box_thick_m);
    const boxT = area ? boxQty / area : Number(d.box_thick_m);
    const murQty = area * 0.1;
    const vata = 2 * L + 2 * W;
    const trunc2 = (x) => Math.trunc(Math.round(x * 100 * 1e6) / 1e6) / 100;
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
    setVal(meas, "I6", boxT);
    setVal(meas, "G6", area);
    setVal(meas, "K6", boxQty);
    setVal(meas, "K7", boxQty);
    setVal(meas, "G10", area);
    setVal(meas, "K10", murQty);
    setVal(meas, "K11", murQty);
    setVal(meas, "K13", murQty);
    setVal(meas, "G16", W);
    setVal(meas, "E16", L);
    setVal(meas, "K16", trunc2(area));
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
    if (Number(d.lead_sevaliya_to_taluka_km) > 0) setVal(lead, "D5", Number(d.lead_sevaliya_to_taluka_km));
    if (Number(d.lead_taluka_to_site_km) > 0) setVal(lead, "D6", Number(d.lead_taluka_to_site_km));
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
      setVal(abs, "F25", Math.round(f24));
      setVal(face, "G21", Math.round(f24));
      setVal(abs, "A28", d.taluka);
    }

    await writeAndRespond(req, res, wb, d, "PAVER", PRINT_AREA_PAVER, "estimate_paver");
  } catch (err) {
    logEvent("estimate_paver_error", { error: String(err) }, req);
    res.status(500).json({ ok: false, error: String(err.message || err) });
  }
});

function persistOut(full) {
  if (!sbOn()) return;
  try {
    const buf = fs.readFileSync(full);
    const ct = /\.pdf$/i.test(full) ? "application/pdf" : "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
    sbUploadFile("out/" + path.basename(full), buf, ct).catch(function (e) { console.error("persistOut", e.message); });
  } catch (_e) {}
}
app.get("/api/download/:name", (req, res) => {
  const name = path.basename(req.params.name);
  const full = path.join(OUT_DIR, name);
  if (!fs.existsSync(full)) {
    if (sbOn() && /^[A-Za-z0-9._-]+\.(xlsx|pdf)$/.test(name)) return res.redirect(302, SB_URL + "/storage/v1/object/public/site-media/out/" + name);
    return res.status(404).json({ ok: false });
  }
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
function sbAuthKey() {
  return process.env.SUPABASE_SERVICE_ROLE || process.env.SUPABASE_SERVICE_KEY || SB_KEY;
}
if (SB_URL && sbAuthKey() === SB_KEY) console.warn("SUPABASE_SERVICE_ROLE is not set: the server falls back to the anon key, so RLS cannot be closed yet");
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

function normMail(s) { return String(s || "").trim().toLowerCase(); }
function mineMerged(remote, local, who) {
  const loc = local || [];
  function hit(r) {
    return loc.find(function (x) {
      if (!normMail(x && x.user_email)) return false;
      if (r && x && r.id && x.id && String(r.id) === String(x.id)) return true;
      const wa = String((r && r.work_name) || "").trim().toLowerCase();
      const wb = String((x && x.work_name) || "").trim().toLowerCase();
      if (!wa || wa !== wb) return false;
      const ta = String((r && r.ts) || "").slice(0, 16);
      const tb = String((x && x.ts) || "").slice(0, 16);
      return ta && tb && ta === tb;
    });
  }
  const stamped = (remote || []).map(function (r) {
    if (normMail(r && r.user_email)) return r;
    const h = hit(r);
    return h ? Object.assign({}, r, { user_email: normMail(h.user_email) }) : r;
  });
  const rows = onlyMine(stamped, who).concat(onlyMine(loc, who));
  const out = [];
  const seen = {};
  rows.forEach(function (r) {
    const id = String((r && r.id) || "");
    const k = String((r && r.work_name) || "").trim().toLowerCase() + "|" + String((r && r.ts) || "").slice(0, 16) + "|" + String((r && r.type) || "").toLowerCase();
    if ((id && seen["i" + id]) || (k !== "||" && seen["k" + k])) return;
    if (id) seen["i" + id] = 1;
    if (k !== "||") seen["k" + k] = 1;
    out.push(r);
  });
  out.sort(function (a, b) { return String(b.ts || "").localeCompare(String(a.ts || "")); });
  return out;
}
async function callerEmail(req) {
  try {
    const u = await authUser(req);
    return normMail(u && u.email);
  } catch (_e) { return ""; }
}
function ownsRow(row, email) {
  email = normMail(email);
  if (!email) return false;
  const own = normMail(row && row.user_email);
  if (own) return own === email;
  return email === OWNER_EMAIL;
}
function onlyMine(rows, email) {
  return (rows || []).filter(function (r) { return ownsRow(r, email); });
}
function ownsTour(row, email) {
  email = normMail(email);
  if (!email) return false;
  const own = normMail(row && row.email);
  if (own && own !== "shared") return own === email;
  return email === OWNER_EMAIL;
}
async function rowAllowed(table, file, id, email) {
  if (!id || !email) return false;
  let row = null;
  if (sbOn()) {
    try {
      const key = sbAuthKey();
      const r = await fetch(SB_URL + "/rest/v1/" + table + "?id=eq." + encodeURIComponent(id) + "&select=*&limit=1", {
        headers: { apikey: key, Authorization: "Bearer " + key }
      });
      if (r.ok) {
        const js = await r.json();
        row = (Array.isArray(js) && js[0]) || null;
      }
    } catch (_e) {}
  }
  if (!row && file) {
    row = dbRead(file, 5000).find(function (x) { return String(x.id || "") === String(id); }) || null;
  }
  return !!(row && ownsRow(row, email));
}
function forgetId(file, id) {
  if (!id) return;
  tombAdd(id);
  if (!file || !fs.existsSync(file)) return;
  const lines = fs.readFileSync(file, "utf8").split("\n").filter(Boolean);
  const keep = [];
  lines.forEach(function (line) {
    let r = null;
    try { r = JSON.parse(line); } catch (_e) { keep.push(line); return; }
    if (String(r.id || "") === String(id)) tombAdd(r.id);
    else keep.push(line);
  });
  fs.writeFileSync(file, keep.length ? keep.join("\n") + "\n" : "");
}
function sbPick(table, row) {
  const cols = {
    estimates: ["id","ts","type","village","taluka","jilla","work_name","fund_head","amounting","length_m","width_m","area","brass","prepared_by","user_email"],
    site_measures: ["id","ts","type","work_name","amounting","rows","area","brass","bill","gps","estimate_id","taluka","village","fund_head","grant_head","contractor","cc_t","test_qty","name_plate","prepared_by","done","user_email"],
    media: ["id","ts","kind","work_name","gps","url","user_email"],
    kachu_bills: ["id","ts","type","work_name","village","amounting","total","net","test_qty","name_plate","preview","xlsx","pdf","user_email"],
    bills: ["id","ts","type","work_name","village","taluka","fund_head","amounting","prepared_by","mb_no","day","net","kind","user_email"],
  }[table] || Object.keys(row);
  const o = {};
  cols.forEach(function (k) { if (row[k] !== undefined) o[k] = row[k]; });
  return o;
}
async function sbInsert(table, row) {
  let body = sbPick(table, row);
  let lastErr = "supabase insert failed";
  const key = sbAuthKey();
  for (let n = 0; n < 12; n++) {
    const r = await fetch(SB_URL + "/rest/v1/" + table, {
      method: "POST",
      headers: {
        apikey: key,
        Authorization: "Bearer " + key,
        "Content-Type": "application/json",
        Prefer: "return=representation"
      },
      body: JSON.stringify(body)
    });
    if (r.ok) {
      const js = await r.json();
      return Array.isArray(js) ? js[0] : js;
    }
    const err = await r.text();
    lastErr = err;
    const missing = err.match(/Could not find the '([^']+)' column/i);
    if (missing && missing[1] !== "user_email" && Object.prototype.hasOwnProperty.call(body, missing[1])) {
      delete body[missing[1]];
      continue;
    }
    if (/uuid/i.test(err) && body.id) {
      delete body.id;
      continue;
    }
    throw new Error(err);
  }
  throw new Error(lastErr);
}
async function sbSelect(table, limit) {
  const key = sbAuthKey();
  const r = await fetch(SB_URL + "/rest/v1/" + table + "?select=*&order=ts.desc&limit=" + (limit || 200), {
    headers: { apikey: key, Authorization: "Bearer " + key }
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
async function sbUploadFile(name, buf, contentType) {
  const key = sbAuthKey();
  const safe = String(name || "file").replace(/^\/+/, "");
  const r = await fetch(SB_URL + "/storage/v1/object/site-media/" + safe, {
    method: "POST",
    headers: {
      apikey: key,
      Authorization: "Bearer " + key,
      "Content-Type": contentType || "application/octet-stream",
      "x-upsert": "true"
    },
    body: buf
  });
  if (!r.ok) throw new Error(await r.text());
  return SB_URL + "/storage/v1/object/public/site-media/" + safe;
}

async function keepFile(url) {
  const s = String(url || "");
  if (!s || /^https?:\/\//i.test(s) || !sbOn()) return { url: s, error: "" };
  const name = path.basename(s.split("?")[0]);
  if (!name) return { url: s, error: "" };
  const full = path.join(OUT_DIR, name);
  if (!fs.existsSync(full)) return { url: s, error: "ફાઇલ ડિસ્ક પર નથી" };
  const ext = path.extname(name).toLowerCase();
  const type = ext === ".pdf"
    ? "application/pdf"
    : "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
  try {
    const pub = await sbUploadFile("bills/" + name, fs.readFileSync(full), type);
    return { url: pub, error: "" };
  } catch (e) {
    return { url: s, error: String(e.message || e) };
  }
}

async function sbUpload(name, buf) {
  return sbUploadFile(name, buf, "image/jpeg");
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
    const gNr = Array.isArray(d.gNr) ? d.gNr : [];
    const sumDia = (dia) => gEx.filter(r => Number(r.dia)===dia).reduce((a,r)=>a+mixNum(r.l),0);
    const lastWD = (dia, dw, dd) => wdOf(gEx.filter(r => Number(r.dia)===dia), dw, dd);

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
    const demo = wdOf(gDe, 0.45, 0.10);
    const demoL = demo.l;
    setVal(meas, "E5", demoL);
    setVal(meas, "G5", demo.w);
    setVal(meas, "I5", demo.d);
    setVal(meas, "E35", n("gCh60"));
    setVal(meas, "E36", n("gCh90"));
    setVal(meas, "E37", n("gCh139"));
    setVal(meas, "E38", n("gCh1313"));
    setVal(meas, "E62", n("gPlate") || 1);
    const bed = wdOf(gNr, 0.45, 0.10);
    setVal(meas, "E58", bed.l);
    setVal(meas, "G58", bed.l ? bed.w : 0);
    setVal(meas, "I58", bed.l ? bed.d : 0);
    meas.getCell("K58").value = { formula: "C58*E58*G58*I58" };
    abs.getCell("F43").value = { formula: "ROUND(F42,0)" };
    face.getCell("G21").value = { formula: "Abstract!F43" };
    setVal(meas, "C2", work);
    setVal(abs, "C2", work);
    if (test) setVal(test, "B1", work);

    await writeAndRespond(req, res, wb, d, "GUTTER", {
      Estimate: "A1:I41",
      Abstract: "A1:F46",
      Measurement: "A1:L63",
      "TEST-SITE": "A1:G36"
    }, "estimate_gutter");
  } catch (e) {
    console.error(e);
    res.status(500).json({ ok: false, error: String(e.message || e) });
  }
});

app.post("/api/db/estimate", async (req, res) => {
  const b = req.body || {};
  const who = await callerEmail(req);
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
    prepared_by: b.prepared_by || "",
    user_email: who
  });
  if (sbOn()) sbInsert("estimates", rec).catch(function(e){ console.error(e.message); });
  res.json({ ok: true, id: rec.id });
});
app.post("/api/db/bills", async (req, res) => {
  const b = req.body || {};
  const who = await callerEmail(req);
  const rec = dbAppend(DB_BILL, {
    kind: "bill",
    type: b.type || "",
    village: b.village || "",
    taluka: b.taluka || "",
    work_name: b.work_name || "",
    fund_head: b.fund_head || b.grant || "",
    amounting: Number(b.amounting || b.net || 0),
    net: Number(b.net || b.amounting || 0),
    day: String(b.day || istDay()).slice(0, 10),
    mb_no: b.mb_no || "",
    prepared_by: b.prepared_by || "",
    user_email: who
  });
  if (sbOn()) sbInsert("bills", rec).catch(function (e) { console.error(e.message); });
  res.json({ ok: true, id: rec.id });
});
app.get("/api/db/bills", async (req, res) => {
  const who = await callerEmail(req);
  const local = notDeleted(dbRead(DB_BILL, 500));
  let remote = [];
  try { if (sbOn()) remote = notDeleted(await sbSelect("bills", 500)); }
  catch (e) { console.error(e.message); }
  res.json({ ok: true, items: dedupeBills(mineMerged(remote, local, who)) });
});
app.post("/api/db/bills/delete", requireUser, async (req, res) => {
  const id = String((req.body && req.body.id) || "");
  const who = normMail(req.user && req.user.email);
  if (!id) return res.json({ ok: false });
  if (!(await rowAllowed("bills", DB_BILL, id, who))) return res.json({ ok: false, error: "આ તમારું નથી" });
  forgetId(DB_BILL, id);
  let removed = 0;
  try { removed = await sbDelete("bills", id); } catch (e) { console.error(e.message); }
  res.json({ ok: true, removed: removed });
});
app.get("/api/db/tour", async (req, res) => {
  try {
    const email = await callerEmail(req);
    if (!email) return res.json({ ok: false, items: [], error: "login" });
    if (!sbOn()) return res.json({ ok: true, items: [] });
    const key = sbAuthKey();
    const r = await fetch(SB_URL + "/rest/v1/tour_days?select=*&order=day.desc&limit=400", {
      headers: { apikey: key, Authorization: "Bearer " + key }
    });
    if (!r.ok) throw new Error(await r.text());
    const items = (await r.json()).filter(function (x) { return ownsTour(x, email); });
    res.json({ ok: true, items: items });
  } catch (e) {
    res.json({ ok: false, items: [], error: String(e.message || e) });
  }
});
app.post("/api/db/tour", async (req, res) => {
  try {
    if (!sbOn()) return res.json({ ok: false, error: "no supabase" });
    const b = req.body || {};
    const email = await callerEmail(req);
    if (!email) return res.json({ ok: false, error: "login" });
    const row = {
      email: email,
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
        apikey: sbAuthKey(),
        Authorization: "Bearer " + sbAuthKey(),
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
    const email = await callerEmail(req);
    if (!email) return res.json({ ok: true, item: null });
    const url = SB_URL + "/rest/v1/tour_profile?select=*&email=eq." + encodeURIComponent(email) + "&limit=1";
    const r = await fetch(url, { headers: { apikey: sbAuthKey(), Authorization: "Bearer " + sbAuthKey() } });
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
    const email = await callerEmail(req);
    if (!email) return res.json({ ok: false, error: "login" });
    const row = {
      email: email,
      name: b.name || "",
      office: b.office || "",
      mobile: b.mobile || "",
      leave: b.leave || {}
    };
    const r = await fetch(SB_URL + "/rest/v1/tour_profile?on_conflict=email", {
      method: "POST",
      headers: {
        apikey: sbAuthKey(),
        Authorization: "Bearer " + sbAuthKey(),
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
app.get("/api/db/estimates", async (req, res) => {
  const who = await callerEmail(req);
  const local = notDeleted(dbRead(DB_EST, 500));
  let remote = [];
  try { if (sbOn()) remote = notDeleted(await sbSelect("estimates", 500)); }
  catch (e) { console.error(e.message); }
  res.json({ ok: true, items: mineMerged(remote, local, who) });
});
app.post("/api/db/site", async (req, res) => {
  const b = req.body || {};
  const who = await callerEmail(req);
  const rows = Array.isArray(b.rows) ? JSON.parse(JSON.stringify(b.rows)) : [];
  const meta = {
    cc_t: Number(b.cc_t || 0) || 0.1,
    test_qty: Number(b.test_qty || 0),
    name_plate: Number(b.name_plate || b.plate || 0),
    aae: b.prepared_by || b.aae || "",
    done: !!b.done
  };
  if (Number(b.exc_d) > 0) meta.exc_d = Number(b.exc_d);
  if (Number(b.dust_d) > 0) meta.dust_d = Number(b.dust_d);
  if (rows[0] && typeof rows[0] === "object") Object.assign(rows[0], meta);
  const rec = dbAppend(DB_SITE, {
    kind: "site",
    type: b.type || "",
    work_name: b.work_name || "",
    amounting: Number(b.amounting || 0),
    rows: rows,
    area: Number(b.area || 0),
    brass: Number(b.brass || 0),
    bill: Number(b.bill || 0),
    gps: b.gps || "",
    estimate_id: b.estimate_id || "",
    taluka: b.taluka || "",
    village: b.village || "",
    fund_head: b.fund_head || b.grant || b.grant_head || "",
    grant_head: b.grant_head || b.grant || b.fund_head || "",
    contractor: b.contractor || "",
    cc_t: meta.cc_t,
    test_qty: meta.test_qty,
    name_plate: meta.name_plate,
    prepared_by: meta.aae,
    done: meta.done,
    user_email: who
  });
  let remote = false;
  let error = "";
  if (sbOn()) {
    try { await sbInsert("site_measures", rec); remote = true; }
    catch (e) { error = String(e.message || e); console.error("site save", error); }
  } else error = "Supabase જોડાયેલું નથી — રીડિપ્લોય પર માપ રહેશે નહીં";
  res.json({ ok: !error, id: rec.id, remote: remote, error: error });
});
app.get("/api/db/site", async (req, res) => {
  const who = await callerEmail(req);
  const local = notDeleted(dbRead(DB_SITE, 500));
  let remote = [];
  try { if (sbOn()) remote = notDeleted(await sbSelect("site_measures", 500)); }
  catch (e) { console.error(e.message); }
  res.json({ ok: true, items: mineMerged(remote, local, who) });
});
app.post("/api/db/site/delete", requireUser, async (req, res) => {
  const id = String((req.body && req.body.id) || "");
  const who = normMail(req.user && req.user.email);
  if (!id) return res.json({ ok: false });
  if (!(await rowAllowed("site_measures", DB_SITE, id, who))) return res.json({ ok: false, error: "આ તમારું નથી" });
  forgetId(DB_SITE, id);
  let removed = 0;
  try { removed = await sbDelete("site_measures", id); } catch (e) { console.error(e.message); }
  res.json({ ok: true, removed: removed });
});

app.post("/api/db/kachu", async (req, res) => {
  const b = req.body || {};
  const who = await callerEmail(req);
  const xlsxKeep = await keepFile(b.xlsx || (b.preview && b.preview.xlsx) || "");
  const pdfKeep = await keepFile(b.pdf || (b.preview && b.preview.pdf) || "");
  const fileError = [xlsxKeep.error, pdfKeep.error].filter(Boolean).join(" | ");
  const rec = dbAppend(DB_KACHU, {
    type: b.type || "paver",
    work_name: b.work_name || "",
    village: b.village || "",
    amounting: Number(b.amounting || 0),
    total: Number(b.total || 0),
    net: Number(b.net || 0),
    test_qty: Number(b.test_qty || 0),
    name_plate: Number(b.name_plate || 0),
    preview: Object.assign({}, b.preview || {}, { xlsx: xlsxKeep.url, pdf: pdfKeep.url }),
    xlsx: xlsxKeep.url,
    pdf: pdfKeep.url,
    user_email: who
  });
  let error = "";
  if (sbOn()) {
    try { await sbInsert("kachu_bills", rec); }
    catch (e) { error = String(e.message || e); console.error("kachu save", error); }
  } else error = "Supabase જોડાયેલું નથી — રીડિપ્લોય પર કાચું બિલ રહેશે નહીં";
  res.json({ ok: !error, id: rec.id, error: error, file_error: fileError, xlsx: xlsxKeep.url, pdf: pdfKeep.url });
});
app.get("/api/db/kachu", async (req, res) => {
  const who = await callerEmail(req);
  const local = dbRead(DB_KACHU, 500);
  let remote = [];
  try { if (sbOn()) remote = await sbSelect("kachu_bills", 500); }
  catch (e) { console.error(e.message); }
  res.json({ ok: true, items: mineMerged(remote, local, who) });
});
app.post("/api/db/media", async (req, res) => {
  const b = req.body || {};
  const who = await callerEmail(req);
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
    file, url,
    user_email: who
  });
  if (sbOn() && b.data && String(b.data).startsWith("data:")) {
    const raw = String(b.data);
    const comma = raw.indexOf(",");
    const buf = Buffer.from(raw.slice(comma + 1), "base64");
    const name = (b.kind || "media") + "-" + Date.now() + ".jpg";
    sbUpload(name, buf).then(function (u) {
      rec.url = u;
      sbInsert("media", { id: rec.id, ts: rec.ts, kind: rec.kind, work_name: rec.work_name, gps: rec.gps, url: u, user_email: rec.user_email }).catch(function(){});
      res.json({ ok: true, id: rec.id, file: u });
    }).catch(function (e) {
      console.error(e.message);
      res.json({ ok: true, id: rec.id, file: url });
    });
    return;
  }
  res.json({ ok: true, id: rec.id, file: url });
});
app.get("/api/db/media", async (req, res) => {
  const who = await callerEmail(req);
  const local = dbRead(DB_MEDIA, 200);
  let remote = [];
  try { if (sbOn()) remote = await sbSelect("media", 200); }
  catch (e) { console.error(e.message); }
  const items = mineMerged(remote, local, who).map((m) => ({
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
    const who = await callerEmail(req);
    let estRemote = [], siteRemote = [], mediaRemote = [];
    if (sbOn()) {
      estRemote = await sbSelect("estimates", 500);
      siteRemote = await sbSelect("site_measures", 500);
      mediaRemote = await sbSelect("media", 200);
    }
    const ests = mineMerged(estRemote, notDeleted(dbRead(DB_EST, 500)), who);
    const sites = mineMerged(siteRemote, notDeleted(dbRead(DB_SITE, 500)), who);
    const media = mineMerged(mediaRemote, dbRead(DB_MEDIA, 200), who);
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
  res.json({ ok: true, supabase: sbOn(), service: !!(SB_URL && sbAuthKey() !== SB_KEY), url: SB_URL ? SB_URL.replace(/https:\/\//,"") : "" });
});
app.get("/api/auth/config", (_req, res) => {
  res.json({ ok: true, url: SB_URL || "", anon: SB_KEY || "" });
});
async function sbProfiles(method, path, body) {
  const r = await fetch(SB_URL + "/rest/v1/" + path, {
    method: method,
    headers: {
      apikey: sbAuthKey(),
      Authorization: "Bearer " + sbAuthKey(),
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
const OWNER_EMAIL = String(process.env.ADMIN_EMAIL || "").trim().toLowerCase();
if (!OWNER_EMAIL) console.warn("ADMIN_EMAIL is not set: nobody is Super Admin until it is set in Render env");
function isOwnerEmail(e) { return !!OWNER_EMAIL && normMail(e) === OWNER_EMAIL; }
async function authUser(req) {
  const tok = String(req.headers.authorization || "").replace(/^Bearer\s+/i, "");
  if (!tok || !SB_URL) return null;
  const keys = [SB_KEY, process.env.SUPABASE_SERVICE_ROLE, process.env.SUPABASE_SERVICE_KEY].filter(Boolean);
  const seen = {};
  for (let i = 0; i < keys.length; i++) {
    const key = keys[i];
    if (seen[key]) continue;
    seen[key] = 1;
    try {
      const r = await fetch(SB_URL + "/auth/v1/user", { headers: { apikey: key, Authorization: "Bearer " + tok } });
      if (r.ok) return await r.json();
    } catch (_e) {}
  }
  return null;
}
async function requireUser(req, res, next) {
  const u = await authUser(req);
  if (!u || !u.email) return res.status(401).json({ ok: false, error: "login required" });
  req.user = u;
  next();
}
async function requireAdmin(req, res, next) {
  const u = await authUser(req);
  const email = String((u && u.email) || "").toLowerCase();
  if (!email) return res.status(401).json({ ok: false, error: "login required" });
  req.user = u;
  if (OWNER_EMAIL && email === OWNER_EMAIL) return next();
  return res.status(403).json({ ok: false, error: "admin only" });
}
const activeCache = {};
async function requireActive(req, res, next) {
  const email = normMail(req.user && req.user.email);
  if (!email) return res.status(401).json({ ok: false, error: "login required" });
  if (isOwnerEmail(email)) return next();
  const hit = activeCache[email];
  if (hit && Date.now() - hit.t < 60000) {
    return hit.ok ? next() : res.status(403).json({ ok: false, error: "approval pending" });
  }
  let ok = false;
  try {
    const js = await sbProfiles("GET", "profiles?select=role,subscription_status,subscription_end_date&email=eq." + encodeURIComponent(email), null);
    const p = Array.isArray(js) ? js[0] : null;
    if (p) {
      const st = String(p.subscription_status || "");
      const end = p.subscription_end_date ? new Date(p.subscription_end_date).getTime() : 0;
      ok = st === "Active" || st === "Approved" || (st === "Trial" && end > Date.now());
    }
  } catch (_e) {}
  activeCache[email] = { ok: ok, t: Date.now() };
  return ok ? next() : res.status(403).json({ ok: false, error: "approval pending" });
}
async function meHandler(req, res) {
  const email = normMail(req.user && req.user.email);
  const admin = isOwnerEmail(email);
  try {
    const js = await sbProfiles("GET", "profiles?select=*&email=eq." + encodeURIComponent(email), null);
    let item = Array.isArray(js) ? (js[0] || null) : null;
    if (admin) item = Object.assign({ email: email }, item || {}, { role: "admin", subscription_status: "Active" });
    res.json({ ok: true, admin: admin, item: item });
  } catch (e) {
    res.json({ ok: false, admin: admin, item: admin ? { email: email, role: "admin", subscription_status: "Active" } : null, error: String(e.message || e) });
  }
}
app.get("/api/me", requireUser, meHandler);
app.get("/api/admin/me", requireUser, meHandler);
async function sbUpsertProfile(row) {
  let body = Object.assign({}, row);
  Object.keys(body).forEach(function (k) { if (body[k] === undefined) delete body[k]; });
  let lastErr = "profile save failed";
  const key = sbAuthKey();
  for (let n = 0; n < 16; n++) {
    const r = await fetch(SB_URL + "/rest/v1/profiles?on_conflict=email", {
      method: "POST",
      headers: {
        apikey: key,
        Authorization: "Bearer " + key,
        "Content-Type": "application/json",
        Prefer: "return=representation,resolution=merge-duplicates"
      },
      body: JSON.stringify(body)
    });
    if (r.ok) {
      const js = await r.json();
      return Array.isArray(js) ? js[0] : js;
    }
    const err = await r.text();
    lastErr = err;
    const missing = err.match(/Could not find the '([^']+)' column/i);
    if (missing && Object.prototype.hasOwnProperty.call(body, missing[1])) {
      delete body[missing[1]];
      continue;
    }
    throw new Error(err);
  }
  throw new Error(lastErr);
}
async function profileHandler(req, res) {
  try {
    if (!sbOn()) return res.json({ ok: false, error: "no supabase" });
    const d = req.body || {};
    const email = normMail(req.user.email);
    const isAd = isOwnerEmail(email);
    let prev = null;
    try {
      const old = await sbProfiles("GET", "profiles?select=*&email=eq." + encodeURIComponent(email), null);
      prev = Array.isArray(old) ? (old[0] || null) : null;
    } catch (_e) {}
    const row = { email: email };
    const name = String(d.full_name || d.name || "").trim();
    const mobile = String(d.mobile_number || d.mobile || "").trim();
    const jilla = String(d.jilla || "").trim();
    const taluka = String(d.taluka || d.office_location || "").trim();
    const job = String(d.designation || d.role || "").trim();
    if (name) { row.name = name; row.full_name = name; }
    if (mobile) { row.mobile = mobile; row.mobile_number = mobile; }
    if (jilla) row.jilla = jilla;
    if (taluka) { row.taluka = taluka; row.office_location = taluka; }
    if (job && job !== "admin" && job !== "user") row.designation = job;
    const subdiv = String(d.sub_division || d.subdiv || "").trim();
    if (subdiv) row.sub_division = subdiv;
    if (isAd) {
      row.subscription_status = "Active";
      row.role = "admin";
    } else {
      // Never change role/status of an existing non-admin row here: only /api/admin/approve may do that.
      // A brand-new row (or one with no status yet) starts a 15-day Trial.
      const st = String((prev && prev.subscription_status) || "");
      const hasEnd = !!(prev && prev.subscription_end_date);
      if (!prev || !st) {
        row.subscription_status = "Trial";
        row.role = "user";
      }
      if ((!prev || !st || st === "Trial") && !hasEnd) {
        const base = prev && prev.created_at ? new Date(prev.created_at) : new Date();
        const start = isNaN(base.getTime()) ? new Date() : base;
        row.subscription_end_date = new Date(start.getTime() + 15 * 864e5).toISOString();
      }
    }
    delete activeCache[email];
    const js = await sbUpsertProfile(row);
    res.json({ ok: true, admin: isAd, item: js });
  } catch (e) {
    res.json({ ok: false, error: String(e.message||e) });
  }
}
app.post("/api/me/profile", requireUser, profileHandler);
app.post("/api/admin/profile", requireUser, profileHandler);
app.get("/api/admin/pending", requireAdmin, async (_req, res) => {
  try {
    if (!sbOn()) return res.json({ ok: true, items: [] });
    const js = await sbProfiles("GET", "profiles?select=*&order=email.asc", null);
    res.json({ ok: true, items: Array.isArray(js)?js:[] });
  } catch (e) {
    res.json({ ok: false, items: [], error: String(e.message||e) });
  }
});
app.post("/api/admin/approve", requireAdmin, async (req, res) => {
  try {
    const email = String((req.body||{}).email||"").toLowerCase();
    const status = (req.body||{}).status || "Active";
    if (!email) return res.json({ ok: false, error: "email" });
    delete activeCache[email];
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
    const asAmt = Number(d.as_amount != null && d.as_amount !== "" ? d.as_amount : (d.amounting || 0));
    const startDate = billDateText(d.start_date || "");
    const measDate = billDateText(d.meas_date || d.date || "");
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
    const pavRates = [156.56, 210.42, 740.51, 23.68, 608, 306.14];

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
    [4, 5, 6, 7, 8, 9].forEach(function (r, i) {
      bill.getCell("E" + r).value = { formula: "ROUND(B" + r + "*" + pavRates[i] + ",2)" };
      bill.getCell("E" + r).numFmt = "0.00";
    });
    bill.getCell("E10").value = { formula: "ROUND(SUM(E4:E9),2)" };
    bill.getCell("E11").value = { formula: "ROUND(E10*0.18,2)" };
    bill.getCell("E12").value = { formula: "ROUND(E10+E11,2)" };
    bill.getCell("E13").value = { formula: "ROUNDDOWN(E12,0)" };
    bill.getCell("E10").numFmt = "0.00";
    bill.getCell("E11").numFmt = "0.00";
    bill.getCell("E12").numFmt = "0.00";
    bill.getCell("E13").numFmt = "0";
    setVal(bill, "A14", aae ? ("શ્રી- " + aae) : "");
    setVal(bill, "B15", measDate);
    setVal(bill, "B16", mb);
    setVal(bill, "D16", pg1);
    setVal(bill, "F16", pg2);

    setVal(comp, "C2", grant);
    setVal(comp, "C3", work);
    setVal(comp, "C4", tsDet);
    setVal(comp, "C5", tsAmt);
    setVal(comp, "C6", asDet);
    setVal(comp, "C7", asAmt);
    setVal(comp, "C8", agency);
    setVal(comp, "F8", gam);
    setVal(comp, "C9", startDate);
    setVal(comp, "C10", measDate);
    comp.getCell("C11").value = { formula: "Bill!E13" };
    setVal(comp, "C12", "MB NO -");
    setVal(comp, "D12", mb);
    setVal(comp, "E12", "PAGE NO");
    setVal(comp, "F12", pg1);
    setVal(comp, "G12", "TO");
    setVal(comp, "H12", pg2);
    setVal(comp, "A20", talLabel || tal || "");
    fitPaLabels(pa);

    const paSheet = wb.getWorksheet("21 No. P.A. Form");
    if (paSheet) ["Q7","Q8","Q9","Q10","Q11","O20"].forEach(function (a) { paSheet.getCell(a).value = null; });
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
    const asAmt = Number(d.as_amount != null && d.as_amount !== "" ? d.as_amount : (d.amounting || 0));
    const startDate = billDateText(d.start_date || "");
    const measDate = billDateText(d.meas_date || d.date || "");
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
      const row = 4 + i;
      const amt = qty * rates[i];
      setVal(bill, "B" + row, qty);
      bill.getCell("E" + row).value = { formula: "ROUND(B" + row + "*" + rates[i] + ",2)" };
      bill.getCell("E" + row).numFmt = "0.00";
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
    bill.getCell("E12").value = { formula: "ROUND(SUM(E4:E11),2)" };
    bill.getCell("E13").value = { formula: "ROUND(E12*0.18,2)" };
    bill.getCell("E14").value = { formula: "ROUND(E12+E13,2)" };
    bill.getCell("E15").value = { formula: "ROUNDDOWN(E14,0)" };
    bill.getCell("E12").numFmt = "0.00";
    bill.getCell("E13").numFmt = "0.00";
    bill.getCell("E14").numFmt = "0.00";
    bill.getCell("E15").numFmt = "0";
    setVal(bill, "A16", aae ? ("શ્રી- " + aae) : "");
    setVal(bill, "B17", measDate);
    setVal(bill, "B18", mb);
    setVal(bill, "D18", pg1);
    setVal(bill, "F18", pg2);

    setVal(comp, "C2", grant);
    setVal(comp, "C3", work);
    setVal(comp, "C4", tsDet);
    setVal(comp, "C5", tsAmt);
    setVal(comp, "C6", asDet);
    setVal(comp, "C7", asAmt);
    setVal(comp, "C8", agency);
    setVal(comp, "F8", gam);
    setVal(comp, "C9", startDate);
    setVal(comp, "C10", measDate);
    comp.getCell("C11").value = { formula: "Bill!E15" };
    setVal(comp, "D12", mb);
    setVal(comp, "F12", pg1);
    setVal(comp, "H12", pg2);
    setVal(comp, "A20", talLabel || tal || "");
    fitPaLabels(pa);

    const paSheet = wb.getWorksheet("21 No. P.A. Form");
    if (paSheet) ["Q7","Q8","Q9","Q10","Q11","O20"].forEach(function (a) { paSheet.getCell(a).value = null; });
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

/* v15d: more rows than the MB form has lines -> the extra rows are merged into the last line
   (same length sum, area-weighted width/depth) so no measurement is dropped */
function ccRows(segs, max) {
  if (segs.length <= max) return segs;
  const head = segs.slice(0, max - 1), tail = segs.slice(max - 1);
  const L = tail.reduce(function (a, s) { return a + s.L; }, 0);
  const A = tail.reduce(function (a, s) { return a + s.L * s.W; }, 0);
  const B = tail.reduce(function (a, s) { return a + s.L * s.W * (s.d || 0); }, 0);
  return head.concat([{ L: L, W: L ? A / L : 0, d: A ? B / A : 0 }]);
}
function paverRows(segs, max) {
  if (segs.length <= max) return segs;
  const head = segs.slice(0, max - 1), tail = segs.slice(max - 1);
  const L = tail.reduce(function (a, s) { return a + s.L; }, 0);
  const A = tail.reduce(function (a, s) { return a + s.L * s.W; }, 0);
  const V = tail.reduce(function (a, s) { return a + 2 * s.L + 2 * s.W; }, 0);
  return head.concat([{ L: L, W: L ? A / L : 0, vata: V }]);
}
function fillGpMbMeas(ws, type, gp) {
  const blocks = type === "gutter"
    ? { 225: 9, 300: 12, 450: 15, 600: 18, 900: 21, 1200: 24 }
    : { 63: 9, 75: 12, 90: 15, 110: 18 };
  const pipes = Array.isArray(gp.pipes) ? gp.pipes : [];
  if (pipes) {
    Object.keys(blocks).forEach(function (dia) {
      const r0 = blocks[dia];
      const list = pipes.filter(function (p) { return Number(p.dia) === Number(dia) && mixNum(p.l) > 0; });
      for (let k = 0; k < 3; k++) {
        const p = list[k];
        const r = r0 + k;
        if (p) {
          setVal(ws, "I" + r, 1);
          setVal(ws, "J" + r, mixNum(p.l));
          setVal(ws, "K" + r, Number(p.w || 0));
          setVal(ws, "L" + r, Number(p.d || 0));
          ws.getCell("M" + r).value = { formula: "I" + r + "*J" + r + "*K" + r + "*L" + r };
        } else {
          /* v15d: unused measurement rows stay blank (no 0 length with template width/depth) */
          ["I", "J", "K", "L", "M"].forEach(function (col) { setVal(ws, col + r, null); });
        }
      }
      /* v15e: unused dia block -> no green "0" subtotal (formula =SUM(J..) left only where pipes exist) */
      if (!list.length) setVal(ws, "N" + (r0 + 2), null);
      if (list.length > 3) {
        const eq = wdOf(list.slice(2), 0, 0);
        setVal(ws, "J" + (r0 + 2), eq.l); setVal(ws, "K" + (r0 + 2), eq.w); setVal(ws, "L" + (r0 + 2), eq.d);
      }
    });
  }
  /* v15d: demolition counts only when opened AND it has a quantity; an empty demolition is hidden
     (abstract row 2 hidden, measurement block H2:N6 blank, numbering shifted) instead of printing 0 with the template width 0.45 */
  const dm = gp.demoRow || null;
  const demoQty = dm
    ? (dm.open ? mixNum(dm.l) * Number(dm.w || 0) * (type === "gutter" ? Number(dm.t || 0) : 1) : 0)
    : Number(gp.demo || 0);
  const demoOn = demoQty > 0;
  for (let r = 3; r <= 5; r++) ["I", "J", "K", "L", "M"].forEach(function (col) { setVal(ws, col + r, null); });
  if (demoOn && dm) {
    setVal(ws, "I3", 1);
    setVal(ws, "J3", mixNum(dm.l));
    setVal(ws, "K3", Number(dm.w || 0));
    if (type === "gutter") setVal(ws, "L3", Number(dm.t || 0));
    ws.getCell("M3").value = { formula: type === "gutter" ? "I3*J3*K3*L3" : "I3*J3*K3" };
  }
  if (!demoOn) {
    for (let r = 2; r <= 6; r++) ["H", "I", "J", "K", "L", "M", "N"].forEach(function (col) { setVal(ws, col + r, null); });
    ws.getRow(2).hidden = true;
  }
  let ccQty = 0;
  if (type === "gutter") {
    const cc = gp.ccRow || null;
    ccQty = !demoOn ? 0 : cc ? mixNum(cc.l) * Number(cc.w || 0) * Number(cc.t || 0) : Number(gp.cc || 0);
    for (let r = 36; r <= 38; r++) ["B", "C", "D", "E", "F"].forEach(function (col) { setVal(ws, col + r, null); });
    if (ccQty > 0 && cc) {
      setVal(ws, "B36", 1); setVal(ws, "C36", mixNum(cc.l)); setVal(ws, "D36", Number(cc.w || 0)); setVal(ws, "E36", Number(cc.t || 0));
      ws.getCell("F36").value = { formula: "B36*C36*D36*E36" };
    } else if (!(ccQty > 0)) {
      for (let r = 35; r <= 39; r++) ["A", "B", "C", "D", "E", "F", "G"].forEach(function (col) { setVal(ws, col + r, null); });
    }
    const ch = gp.ch || {};
    setVal(ws, "I30", Number(ch["60"] || 0));
    setVal(ws, "I31", Number(ch["90"] || 0));
    setVal(ws, "I32", Number(ch["139"] || 0));
    setVal(ws, "I33", Number(ch["1313"] || 0));
    setVal(ws, "H33", "1.30*1.30"); /* v15e: template label said 1.30*0.90 twice */
  }
  return { demoOn: demoOn, demoQty: demoQty, ccQty: ccQty };
}

app.post("/api/mb", async (req, res) => {
  try {
    const d = req.body || {};
    function mix(v) {
      const parts = String(v == null ? "" : v).split("+").map(function (x) { return parseFloat(String(x).trim()); }).filter(function (n) { return !isNaN(n); });
      if (!parts.length) return Number(v) || 0;
      return parts.reduce(function (a, b) { return a + b; }, 0) / parts.length;
    }
    const raw = String(d.type || "paver").toLowerCase();
    const type = raw.indexOf("gutter") >= 0 ? "gutter" : raw.indexOf("pipe") >= 0 ? "pipe" : raw.indexOf("cc") >= 0 ? "cc" : "paver";
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
      ? (wb.getWorksheet("mb-cc road") || findWs(["cc", "road"]) || wb.worksheets[1])
      : type === "gutter"
        ? (wb.getWorksheet("mb-gutter line") || findWs(["gutter"]) || wb.worksheets[2])
        : type === "pipe"
          ? (wb.getWorksheet("mb-pipe line") || findWs(["pipe"]) || wb.worksheets[3])
          : (findWs(["paver"]) || wb.worksheets[0]);
    if (!ws) throw new Error("mb sheet missing: " + wb.worksheets.map(function(w){return w.name;}).join(", "));
    if ((ws.getColumn(5).width || 0) < 16) ws.getColumn(5).width = 16;

    const segs = [];
    rows.forEach(function (r) {
      const L = mix(r.l || r.L);
      const W = mix(r.w || r.W);
      if (!L && !W) return;
      segs.push({ L, W, d: mix(r.d || r.D) });
    });

    const hdr = ws.getCell("A1");
    hdr.value = "કામ નું નામ : " + String(d.work_name || "") + (d.village ? "  (ગામ: " + d.village + ")" : "");
    hdr.alignment = { horizontal: "left", vertical: "middle", wrapText: false };
    hdr.font = Object.assign({}, hdr.font || {}, { bold: true });
    if (type === "gutter" || type === "pipe") {
      const gp = d.gp || {};
      const by = gp.byDia || {};
      const meas = fillGpMbMeas(ws, type, gp);
      const ch = gp.ch || {};
      const num = function (v) { const n = Number(v || 0); return isFinite(n) ? n : 0; };
      const put = function (addr, qty, rate) {
        const q = num(qty);
        const eAddr = "E" + String(addr).replace(/^[A-Z]+/, "");
        if (!(q > 0)) { setVal(ws, addr, null); setVal(ws, eAddr, null); return 0; } /* v15d: empty item = blank, not 0 */
        setVal(ws, addr, q);
        setVal(ws, eAddr, Math.round(q * rate * 100) / 100);
        return q * rate;
      };
      const renum = function (map) { Object.keys(map).forEach(function (a) { setVal(ws, a, map[a]); }); };
      let sub = 0;
      if (type === "gutter") {
        sub += put("C2", meas.demoQty, 1030.81);
        sub += put("C3", gp.exc, 89);
        [[ "C5", 225, 421 ], [ "C6", 300, 672 ], [ "C7", 450, 817 ], [ "C8", 600, 1331 ], [ "C9", 900, 2476 ], [ "C10", 1200, 4121 ]].forEach(function (x) {
          sub += put(x[0], by[x[1]], x[2]);
        });
        [[ "C12", 225, 88 ], [ "C13", 300, 119 ], [ "C14", 450, 171 ], [ "C15", 600, 228 ], [ "C16", 900, 340 ], [ "C17", 1200, 440 ]].forEach(function (x) {
          sub += put(x[0], by[x[1]], x[2]);
        });
        [[ "C19", "60", 5138 ], [ "C20", "90", 7343 ], [ "C21", "139", 8882 ], [ "C22", "1313", 10698 ]].forEach(function (x) {
          sub += put(x[0], ch[x[1]], x[2]);
        });
        sub += put("C23", gp.refill, 22);
        sub += put("C24", gp.frame, 1121);
        sub += put("C25", gp.cover, 1173);
        sub += put("C26", meas.ccQty, 3652.31);
        sub += put("C27", gp.plate, 306.14);
        setVal(ws, "E28", Math.round(sub * 100) / 100);
        setVal(ws, "E29", Math.round(sub * 0.18 * 100) / 100);
        setVal(ws, "E31", Number(d.amounting || 0));
        ws.getCell("E32").value = { formula: "ROUND(E31-E30,2)" };
        if (meas.demoOn) renum({ B24: 7.1, B25: 7.2 });
        else renum({ B3: 1, B4: 2, B11: 3, B18: 4, B23: 5, B24: 6.1, B25: 6.2, B26: 7, B27: 8, H8: 1, H29: "4+6.1+6.2", H35: 5 });
      } else {
        sub += put("C2", meas.demoQty, 202.2);
        sub += put("C3", gp.exc, 89);
        [[ "C5", 63, 69 ], [ "C6", 75, 96 ], [ "C7", 90, 139 ], [ "C8", 110, 199 ]].forEach(function (x) {
          sub += put(x[0], by[x[1]], x[2]);
        });
        [[ "C10", 63, 12 ], [ "C11", 75, 15 ], [ "C12", 90, 17 ], [ "C13", 110, 19 ]].forEach(function (x) {
          sub += put(x[0], by[x[1]], x[2]);
        });
        sub += put("C14", gp.refill, 22);
        sub += put("C15", gp.plate, 306.14);
        setVal(ws, "E16", Math.round(sub * 100) / 100);
        setVal(ws, "E17", Math.round(sub * 0.18 * 100) / 100);
        setVal(ws, "E18", Math.round(sub * 1.18 * 100) / 100);
        setVal(ws, "E19", Number(d.amounting || 0));
        ws.getCell("E20").value = { formula: "ROUND(E19-E18,2)" };
        if (!meas.demoOn) renum({ B3: 1, B4: 2, B9: 3, B14: 4, B15: 5, H8: 1, H24: 4 });
      }
    } else if (type === "paver") {
      const excD = Number(d.exc_d || 0.2);
      const dustD = Number(d.dust_d || 0.1);
      setVal(ws, "I3", excD);
      const pv = paverRows(segs, 12);
      for (let r = 3; r <= 14; r++) {
        const s = pv[r - 3];
        const v = 21 + (r - 3);
        if (s) {
          setVal(ws, "G" + r, s.L);
          setVal(ws, "H" + r, s.W);
          /* v15d: binding (dust) depth from the form, template had 0.1 hard-coded */
          ws.getCell("L" + r).value = { formula: "G" + r + "*H" + r + "*" + dustD };
          if (s.vata != null) { setVal(ws, "H" + v, s.vata / 4); setVal(ws, "L" + v, s.vata / 4); } /* merged line keeps the full vata of the merged rows */
        } else {
          /* v15d: unused rows blank (no 0 × 0 rows, no stale template data) */
          ["G", "H", "I", "J", "K", "L"].forEach(function (col) { setVal(ws, col + r, null); });
          ["G", "H", "I", "K", "L", "M"].forEach(function (col) { setVal(ws, col + v, null); });
        }
      }
      setVal(ws, "C6", Number(d.test_qty == null ? 1 : d.test_qty));
      setVal(ws, "C7", Number(d.name_plate == null ? 0 : d.name_plate));
      setVal(ws, "E11", Number(d.amounting || 0));
    } else {
      const boxT = Number(d.exc_d || 0);
      const ccT = Number(d.cc_t || 0) || 0.1;
      for (let r = 3; r <= 10; r++) {
        ["G","H","I","J","L","M","N","O","P"].forEach(function (col) {
          const c = ws.getCell(col + r);
          if (c) c.value = null;
        });
      }
      const a6 = ws.getCell("A6");
      a6.value = "SPREADING BINDING";
      a6.numFmt = "@";
      const colA = ws.getColumn(1);
      if (!colA.width || colA.width < 24) colA.width = 24;
      setVal(ws, "P1", null);
      setVal(ws, "P2", null);
      const a7 = ws.getCell("A7");
      a7.value = "CC 1:2:4";
      a7.numFmt = "@";
      const m2 = ws.getCell("M2");
      m2.value = "CC 1:2:4";
      m2.numFmt = "@";
      ccRows(segs, 8).forEach(function (s, i) {
        const r = 3 + i;
        /* v15d: blank depth = 0 box cutting, same as the live preview and the bill (was 0.2 via the duplicate exc_d key) */
        const depth = s.d || 0;
        setVal(ws, "G" + r, s.L);
        setVal(ws, "H" + r, s.W);
        setVal(ws, "I" + r, depth);
        setVal(ws, "L" + r, s.L);
        setVal(ws, "M" + r, s.W);
        setVal(ws, "N" + r, ccT);
        const r6 = function (x) { return Math.round(x * 1e6) / 1e6; };
        setVal(ws, "O" + r, r6(s.L * s.W * ccT));
        setVal(ws, "P" + r, r6(s.L * s.W));
        setVal(ws, "J" + r, r6(s.L * s.W * depth));
      });
      for (let r = 2; r <= 9; r++) {
        ws.getCell("E" + r).value = { formula: "ROUND(C" + r + "*D" + r + ",2)" };
      }
      ws.getCell("E10").value = { formula: "ROUND(SUM(E2:E9),2)" };
      ws.getCell("E11").value = { formula: "ROUND(E10*0.18,2)" };
      ws.getCell("E12").value = { formula: "ROUND(E10+E11,2)" };
      ws.getCell("E14").value = { formula: "ROUND(E13-E12,2)" };
      for (let r = 2; r <= 12; r++) ws.getCell("E" + r).numFmt = "0.00";
      ws.getCell("E14").numFmt = "0.00";
      setVal(ws, "C8", Number(d.test_qty == null ? 0 : d.test_qty));
      setVal(ws, "C9", Number(d.name_plate == null ? 0 : d.name_plate));
      setVal(ws, "E13", Number(d.amounting || 0));
    }

    const prefix = type === "cc" ? "MB_CC" : type === "gutter" ? "MB_GUTTER" : type === "pipe" ? "MB_PIPE" : "MB_PAVER";
    const areas = {};
    /* v15e: print areas cover the whole MB form (pipe Net Qty row 31, gutter rows 33-44 = 4th chamber, CC Road measurement, Filling; paver vata total row 33; "Rmt" labels in column O) */
    areas[ws.name] = type === "cc" ? "A1:P20" : type === "gutter" ? "A1:O44" : type === "pipe" ? "A1:O31" : "A1:M33";
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
    let letterSeq = 0;
    function nextNo() {
      const src = fromGujDigits(rawNo);
      const m = src.match(/^(.*?)(\d+)(\D*)$/);
      if (!m) return showVal(rawNo, useGuj);
      const n = String(parseInt(m[2], 10) + letterSeq).padStart(m[2].length, "0");
      letterSeq += 1;
      return showVal(m[1] + n + (m[3] || ""), useGuj);
    }
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
      fitDeeName(ws);
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
      ws.getCell("G" + tr).value = { formula: "SUM(G19:G" + (tr - 1) + ")" };
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
      fitDeeName(ws);
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
      ws.getCell("G" + tr).value = { formula: "SUM(G20:G" + (tr - 1) + ")" };
      set(ws, "B" + (30 + extra), "માપ બુક નંબર");
      set(ws, "D" + (30 + extra), uniqMb(items).join(", "));
      set(ws, "H" + (35 + extra), taluka);
      ws._fitBottom = 35 + extra;
    }
    if (rb.length && wsRb) fillDee(wsRb, rb, nextNo());
    else if (wsRb) wb.removeWorksheet(wsRb.id);
    if (nani.length && wsNani) fillDee(wsNani, nani, nextNo());
    else if (wsNani) wb.removeWorksheet(wsNani.id);
    if (wsAud) fillAudit(wsAud, all, nextNo());
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
    const sumDia = (dia) => pEx.filter(r => Number(r.dia)===dia).reduce((a,r)=>a+mixNum(r.l),0);
    const lastWD = (dia, dw, dd) => wdOf(pEx.filter(r => Number(r.dia)===dia), dw, dd);
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
    const demo = wdOf(pDe.map(function (r) { return { l: r.l, w: r.w, d: 1 }; }), 0.45, 1);
    setVal(meas, "E5", demo.l);
    setVal(meas, "G5", demo.w);
    [[63,"E9","G9","I9",0.45,0.9],[75,"E10","G10","I10",0.45,0.9],[90,"E11","G11","I11",0.45,0.9],[110,"E12","G12","I12",0.45,0.9]].forEach(function(row){
      const wd=lastWD(row[0], row[4], row[5]);
      setVal(meas, row[1], sumDia(row[0]));
      setVal(meas, row[2], wd.w);
      setVal(meas, row[3], wd.d);
    });
    setVal(meas, "E40", Number(d.pPlate||1));
    if (test) { setVal(test, "C12", 0); setVal(test, "C13", 1); setVal(test, "C14", 0); }
    abs.getCell("F27").value = { formula: "ROUND(F26,0)" };
    face.getCell("G21").value = { formula: "Abstract!F27" };
    await writeAndRespond(req, res, wb, d, "PIPE", {
      Estimate: "A1:I41", Abstract: "A1:F30", Measurement: "A1:L42", "TEST-SITE": "A1:G36"
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
  const measDate = billDateText(d.meas_date || d.date || "");
  const mb = d.mb_no || "";
  const pg1 = d.page_from || "";
  const pg2 = d.page_to || "";
  if (pa) {
    ["Q7","Q8","Q9","Q10","Q11","O20"].forEach(function (a) { pa.getCell(a).value = null; });
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
    setVal(comp, "C9", billDateText(d.start_date || ""));
    setVal(comp, "C10", measDate);
    setVal(comp, "D12", mb);
    setVal(comp, "F12", pg1);
    setVal(comp, "H12", pg2);
    setVal(comp, "A20", talLabel);
  }
}

function putBillLine(bill, addr, qty, rate, rows) {
  const q = Number(qty || 0);
  const row = String(addr).replace(/^[A-Z]+/, "");
  setVal(bill, addr, q);
  bill.getCell("E" + row).value = { formula: "ROUND(" + addr + "*" + rate + ",2)" };
  bill.getCell("E" + row).numFmt = "0.00";
  if (rows) rows.push(row);
  return Math.round(q * rate * 100) / 100;
}
function putBillTotals(bill, rows, row) {
  const sum = (rows && rows.length) ? rows.map(function (r) { return "E" + r; }).join(",") : "0";
  bill.getCell("E" + row).value = { formula: "ROUND(SUM(" + sum + "),2)" };
  bill.getCell("E" + (row + 1)).value = { formula: "ROUND(E" + row + "*0.18,2)" };
  bill.getCell("E" + (row + 2)).value = { formula: "ROUND(E" + row + "+E" + (row + 1) + ",2)" };
  bill.getCell("E" + (row + 3)).value = { formula: "ROUNDDOWN(E" + (row + 2) + ",0)" };
  bill.getCell("E" + row).numFmt = "0.00";
  bill.getCell("E" + (row + 1)).numFmt = "0.00";
  bill.getCell("E" + (row + 2)).numFmt = "0.00";
  bill.getCell("E" + (row + 3)).numFmt = "0";
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
    fitPaLabels(pa);
    const q = d.qty || {};
    const n = (k) => Number(q[k] || 0);
    const rows = [];
    let sub = 0;
    sub += putBillLine(bill, "B4", n("demo"), 1030.81, rows);
    sub += putBillLine(bill, "B5", n("exc"), 89, rows);
    [[ "B7", 225, 421 ], [ "B8", 300, 672 ], [ "B9", 450, 817 ], [ "B10", 600, 1331 ], [ "B11", 900, 2476 ], [ "B12", 1200, 4121 ]].forEach(function (x) {
      sub += putBillLine(bill, x[0], n("p" + x[1]), x[2], rows);
    });
    [[ "B14", 225, 88 ], [ "B15", 300, 119 ], [ "B16", 450, 171 ], [ "B17", 600, 228 ], [ "B18", 900, 340 ], [ "B19", 1200, 440 ]].forEach(function (x) {
      sub += putBillLine(bill, x[0], n("p" + x[1]), x[2], rows);
    });
    [[ "B21", "c60", 5138 ], [ "B22", "c90", 7343 ], [ "B23", "c139", 8882 ], [ "B24", "c1313", 10698 ]].forEach(function (x) {
      sub += putBillLine(bill, x[0], n(x[1]), x[2], rows);
    });
    sub += putBillLine(bill, "B25", n("refill"), 22, rows);
    const ch = n("frame") || (n("c60") + n("c90") + n("c139") + n("c1313"));
    sub += putBillLine(bill, "B27", ch, 1121, rows);
    sub += putBillLine(bill, "B28", n("cover") || ch, 1173, rows);
    sub += putBillLine(bill, "B29", n("cc"), 3652.31, rows);
    sub += putBillLine(bill, "B30", n("plate"), 306.14, rows);
    putBillTotals(bill, rows, 31);
    const net = Math.floor(sub * 1.18);
    if (comp) comp.getCell("C11").value = { formula: "Bill!E34" };
    d.net = net;
    await writeAndRespond(req, res, wb, d, "BILL_GUTTER", {
      "21 No. P.A. Form": "A1:J36", Bill: "A1:G40", "COMP-14MU NAN": "A1:H21"
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
    fitPaLabels(pa);
    const q = d.qty || {};
    const n = (k) => Number(q[k] || 0);
    const rows = [];
    let sub = 0;
    sub += putBillLine(bill, "B4", n("demo"), 202.2, rows);
    sub += putBillLine(bill, "B5", n("exc"), 89, rows);
    [[ "B7", 63, 69 ], [ "B8", 75, 96 ], [ "B9", 90, 139 ], [ "B10", 110, 199 ]].forEach(function (x) {
      sub += putBillLine(bill, x[0], n("p" + x[1]), x[2], rows);
    });
    [[ "B12", 63, 12 ], [ "B13", 75, 15 ], [ "B14", 90, 17 ], [ "B15", 110, 19 ]].forEach(function (x) {
      sub += putBillLine(bill, x[0], n("p" + x[1]), x[2], rows);
    });
    sub += putBillLine(bill, "B16", n("refill"), 22, rows);
    sub += putBillLine(bill, "B17", n("plate"), 306.14, rows);
    putBillTotals(bill, rows, 18);
    const net = Math.floor(sub * 1.18);
    if (comp) comp.getCell("C11").value = { formula: "Bill!E21" };
    d.net = net;
    await writeAndRespond(req, res, wb, d, "BILL_PIPE", {
      "21 No. P.A. Form": "A1:J36", Bill: "A1:G40", "COMP-14MU NAN": "A1:H21"
    }, "bill_pipe");
  } catch (e) {
    res.status(500).json({ ok: false, error: String(e.message || e) });
  }
});


// v15: a bug inside an async route must never take the whole server down (Express 4 does not catch rejected promises).
process.on("unhandledRejection", function (e) { console.error("unhandledRejection:", (e && e.stack) || e); });

if (require.main === module) {
  resetLoProfile();
  app.listen(PORT, () => {
    console.log("ParaState MVP on " + PORT);
  });
}
module.exports = { bakeFormulaResults };
