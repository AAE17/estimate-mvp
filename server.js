const express = require("express");
const cors = require("cors");
const fs = require("fs");
const path = require("path");
const ExcelJS = require("exceljs");

const app = express();
app.use(cors());
app.use(express.json({ limit: "2mb" }));
app.use(express.static(__dirname));

const LOG_FILE = path.join(__dirname, "events.jsonl");
const OUT_DIR = path.join(__dirname, "output");
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
  const cell = ws.getCell(addr);
  cell.value = v;
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
    const wb = new ExcelJS.Workbook();
    await wb.xlsx.readFile(path.join(__dirname, "skeleton-cc.xlsx"));

    const face = wb.getWorksheet("FACE SHEET");
    const meas = wb.getWorksheet("Measurement");
    const lead = wb.getWorksheet("Lead");
    if (!face || !meas || !lead) {
      throw new Error("skeleton-cc.xlsx sheets missing");
    }

    // Only user-input cells. Do not rewrite formulas — that corrupts Excel.
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
    // Do not write D35 — it is merged with B35 (મોજે). Taluka stays H19 → H39.

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

    if (face) {
      face.getCell("H39").value = { formula: "H19", result: taluka };
    }
    if (abs) {
      abs.getCell("A32").value = { formula: faceH39, result: taluka };
    }
    if (lead) {
      lead.getCell("C5").value = { formula: faceH39, result: taluka };
      lead.getCell("A6").value = { formula: "C5", result: taluka };
      lead.getCell("B38").value = { formula: faceH39, result: taluka };
    }
    if (ra) {
      ra.getCell("D38").value = { formula: faceH39, result: taluka };
    }
    if (sch) {
      sch.getCell("C18").value = { formula: faceH39, result: taluka };
    }

    if (wb.calcProperties) {
      wb.calcProperties.fullCalcOnLoad = true;
    }

    const safe = String(d.village || "gam").replace(/[^a-zA-Z0-9._-]+/g, "_");
    const name = `CC_${safe}_${Date.now()}.xlsx`;
    const full = path.join(OUT_DIR, name);
    await wb.xlsx.writeFile(full);

    logEvent("estimate_cc", { user: d.user_name, village: d.village, file: name }, req);
    res.json({
      ok: true,
      xlsx: `/api/download/${name}`,
      pdf: null,
      pdf_error: "PDF Render par pachi. Have Excel download thao.",
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
