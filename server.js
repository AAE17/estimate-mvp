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

    // Formula cells keep formula, but cached result must match new inputs
    // otherwise Excel/mobile shows old Golaj / 40x3 numbers.
    function fres(ws, addr, formula, result) {
      ws.getCell(addr).value = { formula: formula, result: result };
    }

    const L = Number(d.length_m);
    const W = Number(d.width_m);
    const box = Number(d.box_thick_m);
    const bt = Number(d.bt_thick_m);
    const voids = Number(d.voids);
    const murPct = Number(d.murrum_pct);
    const ccT = Number(d.cc_thick_m);
    const area = L * W;
    const boxQty = 1 * area * box;
    const metal = 1 * area * bt;
    const voidQty = metal * voids;
    const metalTot = metal + voidQty;
    const murQty = metalTot * (murPct / 100);
    const ccQty = 1 * area * ccT;
    const leadKm = Number(d.lead_sevaliya_to_taluka_km) + Number(d.lead_taluka_to_site_km);

    face.getCell("D30").value = { formula: "D28", result: d.prepared_by };
    face.getCell("F35").value = { formula: "D9", result: d.fund_head };
    face.getCell("H39").value = { formula: "H19", result: d.taluka };

    fres(meas, "C1", "Abstract!B2", d.work_name);
    fres(meas, "I3", "C3", L);
    fres(meas, "J3", "C4", W);
    fres(meas, "K3", "I3*J3", area);
    fres(meas, "K4", "C3*C4", area);
    fres(meas, "G6", "K3", area);
    fres(meas, "K6", "E6*G6*I6", boxQty);
    fres(meas, "E9", "E6", 1);
    fres(meas, "G9", "G6", area);
    fres(meas, "K9", "E9*G9*I9", metal);
    fres(meas, "E10", "K9", metal);
    fres(meas, "K10", "E10*G10", voidQty);
    fres(meas, "K11", "SUM(K9:K10)", metalTot);
    fres(meas, "G14", "K11", metalTot);
    fres(meas, "K14", "I14*G14%", murQty);
    fres(meas, "K17", "K11", metalTot);
    fres(meas, "K20", "K14", murQty);
    fres(meas, "E23", "C3", L);
    fres(meas, "G23", "K3", area);
    fres(meas, "K23", "TRUNC((G23*E23),2)", Math.trunc(area * L * 100) / 100);
    fres(meas, "E26", "E9", 1);
    fres(meas, "G26", "K4", area);
    fres(meas, "K26", "E26*G26*I26", ccQty);

    fres(lead, "B1", "Abstract!B2", d.work_name);
    fres(lead, "C5", "'FACE SHEET'!H39", d.taluka);
    fres(lead, "A6", "C5", d.taluka);
    fres(lead, "D7", "D5+D6", leadKm);
    fres(lead, "D8", "TRUNC(D7,0)", Math.trunc(leadKm));

    const remLead = Math.max(Math.trunc(leadKm) - 5, 0);
    const metalRate = 417.53 + 81.8 + remLead * 5.95;
    const schSheet = wb.getWorksheet("Schedule");
    if (schSheet) {
      fres(schSheet, "C4", "Lead!D8", Math.trunc(leadKm));
      schSheet.getCell("F4").value = remLead;
      fres(schSheet, "H4", "G4*F4", remLead * 5.95);
      fres(schSheet, "I4", "D4+E4+H4", metalRate);
    }

    const abs = wb.getWorksheet("Abstract");
    if (abs) {
      const murRate = 90 + 81.8;
      const cess = (r) => Math.trunc(r * 1.01 * 100) / 100;
      const f4 = boxQty * cess(156.56);
      const f6 = metalTot * cess(metalRate);
      const f8 = murQty * cess(murRate);
      const f10 = metalTot * cess(247.28);
      const f12 = murQty * cess(146.01);
      const f16 = ccQty * cess(4866.35);
      const f18 = 2656;
      const f20 = cess(303.11);
      const f22 = f4 + f6 + f8 + f10 + f12 + f16 + f18 + f20;
      const f23 = f22 * 0.18;
      const f24 = f22 + f23;

      fres(abs, "A2", "'FACE SHEET'!A21", "કામ નું નામ :");
      fres(abs, "B2", "'FACE SHEET'!C21", d.work_name);
      fres(abs, "A4", "Measurement!K6", boxQty);
      fres(abs, "I4", "Measurement!K6", boxQty);
      fres(abs, "A6", "Measurement!K11", metalTot);
      fres(abs, "D6", "Schedule!I4", metalRate);
      fres(abs, "D7", "TRUNC((D6*1.01),2)", Math.trunc(metalRate * 1.01 * 100) / 100);
      fres(abs, "I5", "Measurement!K11", metalTot);
      fres(abs, "A8", "Measurement!K14", murQty);
      fres(abs, "I6", "Measurement!K14", murQty);
      fres(abs, "A10", "Measurement!K17", metalTot);
      fres(abs, "I7", "Abstract!I5", metalTot);
      fres(abs, "A12", "Measurement!K20", murQty);
      fres(abs, "I8", "I6", murQty);
      fres(abs, "I9", "Measurement!K4", area);
      fres(abs, "A16", "Measurement!K26", ccQty);
      fres(abs, "I10", "Measurement!K26", ccQty);
      fres(abs, "A20", "Measurement!K29", 1);
      fres(abs, "F4", "A4*D5", f4);
      fres(abs, "F6", "A6*D7", f6);
      fres(abs, "F8", "A8*D9", f8);
      fres(abs, "F10", "A10*D11", f10);
      fres(abs, "F12", "A12*D13", f12);
      fres(abs, "F16", "A16*D17", f16);
      fres(abs, "F18", "A18*D18", f18);
      fres(abs, "F20", "A20*D21", f20);
      fres(abs, "F22", "SUM(F4:F21)", f22);
      fres(abs, "F23", "F22*18%", f23);
      fres(abs, "F24", "F22+F23", f24);
      fres(abs, "F25", "'FACE SHEET'!G22", Number(d.amounting || 0));
    }

    const ra = wb.getWorksheet("RA");
    if (ra) {
      fres(ra, "C1", "Abstract!B2", d.work_name);
      ra.getCell("D38").value = d.taluka;
    }
    const sch = wb.getWorksheet("Schedule");
    if (sch) {
      fres(sch, "C1", "Abstract!B2", d.work_name);
      sch.getCell("C18").value = d.taluka;
    }

    if (wb.calcProperties) wb.calcProperties.fullCalcOnLoad = true;

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
