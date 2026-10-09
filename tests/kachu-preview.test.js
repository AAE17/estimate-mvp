// 34 kachu-bill (MB) cases: the on-screen preview (index.html functions) must equal the independent hand calculation.
const test = require("node:test");
const assert = require("node:assert");
const cases = require("./fixtures/kachu-cases.json");
const { loadCode, runCase } = require("./helpers/preview");
const { expected, previewKey } = require("./helpers/expected");

const code = loadCode();

for (const c of cases) {
  test("preview " + c.id + " " + c.type + ": " + c.t, function () {
    const pv = runCase(code, c);
    const exp = expected(c, pv.state);
    const got = {};
    pv.items.forEach(function (it) {
      const k = previewKey(c.type, it[0]);
      got[k] = (got[k] || 0) + it[1];
      if (c.type === "pipe" || c.type === "gutter") assert.ok(Math.abs(it[1]) > 1e-12, "zero row shown: " + it[0]);
    });
    for (const [k, v] of Object.entries(exp.items)) {
      if (k[0] === "_") continue;
      assert.ok(Math.abs((got[k] || 0) - v[0]) < 1e-6, k + " qty " + (got[k] || 0) + " != " + v[0]);
    }
    assert.ok(Math.abs(pv.sub - exp.sub) < 0.5, "sub " + pv.sub + " != " + exp.sub);
    assert.ok(Math.abs(pv.tot - exp.sub * 1.18) < 0.5, "total " + pv.tot + " != " + exp.sub * 1.18);
    assert.ok(!/NaN|undefined|Infinity/.test(pv.html), "NaN/undefined on screen");
  });
}
