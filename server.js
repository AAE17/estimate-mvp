const express = require("express");
const cors = require("cors");
const fs = require("fs");
const path = require("path");
const { execFile } = require("child_process");
const ExcelJS = require("exceljs");
const PDFDocument = require("pdfkit");

const app = express();
app.use(cors());
app.use(express.json({ limit: "2mb" }));
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

function applyFit(ws) {
  let maxR = 1;
  let maxC = 1;
  ws.eachRow({ includeEmpty: false }, (row, r) => {
    if (r > maxR) maxR = r;
    row.eachCell({ includeEmpty: false }, (_c, c) => {
      if (c > maxC) maxC = c;
    });
  });
  const lastCol = ws.getColumn(maxC).letter;
  ws.pageSetup = Object.assign({}, ws.pageSetup, {
    paperSize: 9,
    orientation: maxC > 10 ? "landscape" : "portrait",
    fitToPage: true,
    fitToWidth: 1,
    fitToHeight: 1,
    horizontalCentered: true,
    verticalCentered: false,
    margins: { left: 0.35, right: 0.35, top: 0.4, bottom: 0.4, header: 0.15, footer: 0.15 },
  });
  try {
    ws.pageSetup.printArea = `A1:${lastCol}${maxR}`;
  } catch (_e) {}
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
    execFile(
      sofficeBin(),
      ["--headless", "--norestore", "--nolockcheck", "--convert-to", "pdf", "--outdir", dir, xlsxPath],
      { timeout: 90000 },
      (err) => {
        if (err) return reject(err);
        if (!fs.existsSync(pdfPath)) return reject(new Error("pdf missing"));
        resolve(pdfPath);
      }
    );
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
    const output = d.output || "xlsx";
    const wb = new ExcelJS.Workbook();
    await wb.xlsx.readFile(path.join(__dirname, "skeleton-cc.xlsx"));

    const face = wb.getWorksheet("FACE SHEET");
    const meas = wb.getWorksheet("Measurement");
    const lead = wb.getWorksheet("Lead");
    if (!face || !meas || !lead) {
      throw new Error("skeleton-cc.xlsx sheets missing");
    }

    setVal(face, "F3", d.division);
    setVal(face, "G3", d.jilla);
    setVal(face, "F5", d.subdiv_address);
    setVal(face, "I5", d.nani_address);
    setVal(face, "D9", d.fund_head);
    setVal(face, "H19", d.taluka);
    setVal(face, "C21", d.work_name);
    setVal(face, "G22", Number(d.amounting || 0));
    setVal(face, "D28", d.prepared_by);
    setVal(face, "B34", d.sr_no);
    setVal(face, "C34", d.ss_details);
    setVal(face, "B35", d.village);

    setVal(meas, "C3", Number(d.length_m));
    setVal(meas, "C4", Number(d.width_m));
    setVal(meas, "I6", Number(d.box_thick_m));
    setVal(meas, "I9", Number(d.bt_thick_m));
    setVal(meas, "G10", Number(d.voids));
    setVal(meas, "I14", Number(d.murrum_pct));
    setVal(meas, "I26", Number(d.cc_thick_m));

    setVal(lead, "D5", Number(d.lead_sevaliya_to_taluka_km));
    setVal(lead, "D6", Number(d.lead_taluka_to_site_km));
    setVal(lead, "D11", 5);

    const abs = wb.getWorksheet("Abstract");
    const ra = wb.getWorksheet("RA");
    const sch = wb.getWorksheet("Schedule");
    const taluka = d.taluka || "";
    const faceH39 = "'FACE SHEET'!H39";

    face.getCell("H39").value = { formula: "H19", result: taluka };
    if (abs) abs.getCell("A32").value = { formula: faceH39, result: taluka };
    lead.getCell("C5").value = { formula: faceH39, result: taluka };
    lead.getCell("A6").value = { formula: "C5", result: taluka };
    lead.getCell("B38").value = { formula: faceH39, result: taluka };
    if (ra) ra.getCell("D38").value = { formula: faceH39, result: taluka };
    if (sch) sch.getCell("C18").value = { formula: faceH39, result: taluka };

    wb.worksheets.forEach(applyFit);
    if (wb.calcProperties) wb.calcProperties.fullCalcOnLoad = true;

    const safe = String(d.village || "gam").replace(/[^a-zA-Z0-9._-]+/g, "_");
    const stamp = Date.now();
    const xlsxName = `CC_${safe}_${stamp}.xlsx`;
    const pdfName = `CC_${safe}_${stamp}.pdf`;
    const xlsxFull = path.join(OUT_DIR, xlsxName);
    const pdfFull = path.join(OUT_DIR, pdfName);
    await wb.xlsx.writeFile(xlsxFull);

    let pdfUrl = null;
    let pdfError = null;
    if (output === "pdf" || output === "both") {
      try {
        await convertWithSoffice(xlsxFull);
        const produced = xlsxFull.replace(/\.xlsx$/i, ".pdf");
        if (produced !== pdfFull && fs.existsSync(produced)) {
          fs.copyFileSync(produced, pdfFull);
        }
        const finalPdf = fs.existsSync(pdfFull) ? pdfFull : produced;
        if (!fs.existsSync(finalPdf)) throw new Error("pdf missing");
        pdfUrl = `/api/download/${path.basename(finalPdf)}`;
      } catch (e1) {
        pdfError =
          "PDF LibreOffice vagar nathi. Render Settings ma Runtime = Docker karo (Dockerfile repo ma che).";
        logEvent("pdf_fail", { error: String(e1.message || e1) }, req);
      }
    }

    logEvent("estimate_cc", { village: d.village, output, xlsx: xlsxName, pdf: pdfUrl }, req);

    const wantXlsx = output === "xlsx" || output === "both";
    res.json({
      ok: true,
      xlsx: wantXlsx ? `/api/download/${xlsxName}` : null,
      pdf: pdfUrl,
      pdf_error: pdfError,
    });
  } catch (err) {
    logEvent("estimate_cc_error", { error: String(err) }, req);
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
