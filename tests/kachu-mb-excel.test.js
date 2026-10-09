// 34 kachu-bill cases through POST /api/mb: the Excel written by server.js must equal the hand calculation,
// empty sections hidden/blank, no stale template rows, nothing dropped.
const test = require("node:test");
const assert = require("node:assert");
const ExcelJS = require("exceljs");
const cases = require("./fixtures/kachu-cases.json");
const { loadCode, runCase } = require("./helpers/preview");
const { expected, mix } = require("./helpers/expected");
const { startApp } = require("./helpers/app");

const ROWS = {
  pipe: { demo: 2, exc: 3, pipe63: 5, pipe75: 6, pipe90: 7, pipe110: 8, lay63: 10, lay75: 11, lay90: 12, lay110: 13, refill: 14, plate: 15, _sub: 16, _gst: 17, _tot: 18, _est: 19 },
  gutter: { demo: 2, exc: 3, pipe225: 5, pipe300: 6, pipe450: 7, pipe600: 8, pipe900: 9, pipe1200: 10, lay225: 12, lay300: 13, lay450: 14, lay600: 15, lay900: 16, lay1200: 17,
    ch60: 19, ch90: 20, ch139: 21, ch1313: 22, refill: 23, frame: 24, cover: 25, cc: 26, plate: 27, _sub: 28, _gst: 29, _est: 31 }
};
const n = function (v) { if (v && typeof v === "object" && "result" in v) v = v.result; return typeof v === "number" ? v : Number(v || 0); };
const blank = function (v) { return v === null || v === undefined || v === ""; };

let app, code;
test.before(async function () { app = await startApp(); code = loadCode(); });
test.after(function () { if (app) app.stop(); });

for (const c of cases) {
  test("MB excel " + c.id + " " + c.type + ": " + c.t, async function () {
    const pv = runCase(code, c);
    const exp = expected(c, pv.state);
    const body = Object.assign(pv.body, { work_name: "QA TEST " + c.id, village: "QA TEST", contractor: "QA TEST", aae: "QA TEST", output: "xlsx" });
    const r = await app.post("/api/mb", body);
    assert.equal(r.status, 200, JSON.stringify(r.json));
    const wb = new ExcelJS.Workbook();
    await wb.xlsx.load(await app.download(r.json.xlsx));
    const ws = wb.worksheets[0];
    const v = function (a) { return ws.getCell(a).value; };
    if (c.type === "pipe" || c.type === "gutter") {
      const R = ROWS[c.type];
      for (const [k, [q, rate]] of Object.entries(exp.items)) {
        const row = R[k];
        if (q > 0) {
          assert.ok(Math.abs(n(v("C" + row)) - q) < 1e-3, k + " C" + row + "=" + v("C" + row) + " != " + q);
          assert.ok(Math.abs(n(v("E" + row)) - Math.round(q * rate * 100) / 100) < 0.011, k + " E" + row);
        } else if (k !== "demo") {
          assert.ok(blank(v("C" + row)) && blank(v("E" + row)), "empty " + k + " printed in row " + row);
        }
      }
      assert.ok(Math.abs(n(v("E" + R._sub)) - exp.sub) < 0.6, "sub " + v("E" + R._sub) + " != " + exp.sub);
      assert.ok(Math.abs(n(v("E" + R._gst)) - exp.sub * 0.18) < 0.6, "GST");
      if (R._tot) assert.ok(Math.abs(n(v("E" + R._tot)) - exp.sub * 1.18) < 0.8, "total");
      assert.equal(n(v("E" + R._est)), 100000, "EST");
      const demo = exp.items.demo[0];
      if (demo > 0) {
        assert.ok(!ws.getRow(2).hidden, "demolition row hidden");
        ["J4", "K4", "L4", "J5", "K5", "L5"].forEach(function (a) { assert.ok(!n(v(a)), "demolition leftover " + a); });
      } else {
        assert.ok(ws.getRow(2).hidden, "empty demolition row 2 not hidden");
        ["J3", "K3", "L3", "M3", "K4", "K5", "M6"].forEach(function (a) { assert.ok(blank(v(a)), "empty demolition shows " + a + "=" + JSON.stringify(v(a))); });
      }
      const blocks = c.type === "pipe" ? { 63: 9, 75: 12, 90: 15, 110: 18 } : { 225: 9, 300: 12, 450: 15, 600: 18, 900: 21, 1200: 24 };
      for (const [dia, r0] of Object.entries(blocks)) {
        let sum = 0;
        for (let k = 0; k < 3; k++) {
          const J = v("J" + (r0 + k));
          sum += n(J);
          if (blank(J)) assert.ok(blank(v("K" + (r0 + k))) && blank(v("L" + (r0 + k))), "0-length row " + (r0 + k) + " has width/depth");
        }
        assert.ok(Math.abs(sum - exp.items["pipe" + dia][0]) < 1e-3, "detail " + dia + " lengths " + sum);
      }
      if (c.type === "gutter" && exp.items.cc[0] === 0) ["C36", "D36", "E36"].forEach(function (a) { assert.ok(blank(v(a)), "empty CC block shows " + a); });
    } else {
      const max = c.type === "paver" ? 12 : 8;
      let rows = c.rows.filter(function (r) { return mix(r.l) || mix(r.w); });
      if (rows.length > max) {
        const tail = rows.slice(max - 1);
        const L = tail.reduce(function (a, r) { return a + mix(r.l); }, 0), A = tail.reduce(function (a, r) { return a + mix(r.l) * mix(r.w); }, 0);
        rows = rows.slice(0, max - 1).concat([{ l: L, w: A / L }]);
      }
      for (let i = 0; i < max; i++) {
        const r = 3 + i;
        if (i < rows.length) {
          assert.ok(Math.abs(n(v("G" + r)) - mix(rows[i].l)) < 1e-6 && Math.abs(n(v("H" + r)) - mix(rows[i].w)) < 1e-6, "G/H" + r);
        } else {
          assert.ok(!n(v("G" + r)) && !n(v("H" + r)), "stale template row " + r);
        }
      }
      if (c.type === "cc") {
        let J = 0, O = 0, P = 0;
        for (let r = 3; r <= 10; r++) { J += n(v("J" + r)); O += n(v("O" + r)); P += n(v("P" + r)); }
        assert.ok(Math.abs(J - exp.items.box[0]) < 1e-3, "box " + J + " != " + exp.items.box[0]);
        assert.ok(Math.abs(O - exp.items.cc[0]) < 1e-3, "cc " + O + " != " + exp.items.cc[0]);
        assert.ok(Math.abs(P - exp.items._area) < 1e-3, "area " + P);
        assert.equal(n(v("C8")), exp.items.test[0], "test");
        assert.equal(n(v("C9")), exp.items.plate[0], "plate");
        assert.equal(n(v("E13")), 100000, "EST");
      } else {
        assert.ok(Math.abs(n(v("I3")) - exp.items._exc) < 1e-9, "exc depth");
        const f = v("L3") && v("L3").formula;
        if (rows.length) assert.ok(f && f.endsWith("*" + exp.items._dust), "binding depth formula " + f);
        assert.equal(n(v("C6")), exp.items.test[0], "test");
        assert.equal(n(v("C7")), exp.items.plate[0], "plate");
        assert.equal(n(v("E11")), 100000, "EST");
      }
    }
  });
}
