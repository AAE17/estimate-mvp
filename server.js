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
  const events = lines.slice(-50).map((l) => JSON.parse(l)).reverse();
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

    face.getCell("F3").value = d.division;
    face.getCell("G3").value = d.jilla;
    face.getCell("F5").value = d.subdiv_address;
    face.getCell("I5").value = d.nani_address;
    face.getCell("D9").value = d.fund_head;
    face.getCell("H19").value = d.taluka;
    face.getCell("C21").value = d.work_name;
    face.getCell("G22").value = Number(d.amounting || 0);
    face.getCell("D28").value = d.prepared_by;
    face.getCell("B34").value = d.sr_no;
    face.getCell("C34").value = d.ss_details;
    face.getCell("B35").value = d.village;
    face.getCell("D35").value = d.taluka;

    meas.getCell("C3").value = Number(d.length_m);
    meas.getCell("C4").value = Number(d.width_m);
    meas.getCell("I6").value = Number(d.box_thick_m);
    meas.getCell("I9").value = Number(d.bt_thick_m);
    meas.getCell("G10").value = Number(d.voids);
    meas.getCell("I14").value = Number(d.murrum_pct);
    meas.getCell("I26").value = Number(d.cc_thick_m);

    lead.getCell("D5").value = Number(d.lead_sevaliya_to_taluka_km);
    lead.getCell("D6").value = Number(d.lead_taluka_to_site_km);
    lead.getCell("D11").value = 5;

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
