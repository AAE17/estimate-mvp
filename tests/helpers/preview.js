// Runs the REAL kachu-bill preview functions from index.html inside node (no browser), so tests check
// exactly what the user sees on screen. Later (step 4) these functions move to public/js/shared/calc,
// and this helper will simply require that file.
const fs = require("fs");
const vm = require("vm");
const path = require("path");

const NAMES = ["parseMix", "diaryArea", "diaryType", "ccThick", "mbPaverTotals", "mbCcTotals", "gpDefWD", "gpRefillQty",
  "gpBlank", "gpState", "gpSiteItems", "kachuMoney", "kachuBody", "liveMbBill", "renderKachuSheets", "renderGpSite",
  "gpDemoToggle", "gpSiteDemo", "gpSiteCc", "gpSitePlate", "gpSitePipe", "gpSitePipeAdd", "gpSiteDia", "gpSiteCh", "gpSiteChAdd"];

function grab(src, name) {
  const i = src.indexOf("function " + name + "(");
  if (i < 0) throw new Error("index.html: function " + name + " not found");
  let d = 0, q = null;
  for (let k = src.indexOf("{", i); k < src.length; k++) {
    const c = src[k];
    if (q) { if (c === "\\") { k++; continue; } if (c === q) q = null; continue; }
    if (c === '"' || c === "'" || c === "`") { q = c; continue; }
    if (c === "{") d++;
    else if (c === "}") { d--; if (!d) return src.slice(i, k + 1); }
  }
  throw new Error("unbalanced " + name);
}
function grabVar(src, re) { const m = src.match(re); if (!m) throw new Error("index.html: " + re); return m[0]; }

function loadCode(indexPath) {
  const src = fs.readFileSync(indexPath || path.join(__dirname, "..", "..", "index.html"), "utf8");
  let code = grabVar(src, /var GPDIA=\{[\s\S]*?\n  \};/) + "\n" + grabVar(src, /var GPCH=\[.*?\];/) + "\nvar SITEGP={};\n";
  NAMES.forEach(function (n) { code += grab(src, n) + "\n"; });
  return code;
}

// c = {id, type, gp?, rows?, dom?, actions?, work_name?}
function runCase(code, c) {
  const dom = {};
  Object.entries(c.dom || {}).forEach(function ([k, v]) { dom[k] = { value: String(v) }; });
  dom.dType = { value: c.type };
  ["dBill", "gpAutoLine", "kachuAbs", "kachuMeas", "gpSite"].forEach(function (k) {
    dom[k] = { innerHTML: "", textContent: "", classList: { remove() {}, add() {}, toggle() {} } };
  });
  const sb = {
    document: { getElementById: function (id) { return id in dom ? dom[id] : null; } },
    DIARY: c.rows || [{ l: "", w: "", d: "" }], diaryEst: null,
    siteWorkName: function () { return c.work_name || "QA TEST"; }, siteFundHead: function () { return "QA"; }
  };
  vm.createContext(sb);
  vm.runInContext(code, sb);
  if (c.gp) vm.runInContext("SITEGP[" + JSON.stringify(c.type) + "]=" + JSON.stringify(c.gp) + ";", sb);
  (c.actions || []).forEach(function (a) { vm.runInContext(a, sb); });
  const state = JSON.parse(JSON.stringify(vm.runInContext("SITEGP", sb)));
  const m = vm.runInContext("kachuMoney()", sb);
  const body = vm.runInContext('kachuBody("xlsx")', sb);
  vm.runInContext("liveMbBill(); renderKachuSheets();", sb);
  return { items: m.items, sub: m.sub, gst: m.gst, tot: m.tot, body: JSON.parse(JSON.stringify(body)), state: state,
    html: dom.dBill.innerHTML + dom.kachuAbs.innerHTML + dom.kachuMeas.innerHTML + dom.gpAutoLine.textContent };
}

module.exports = { loadCode, runCase };
