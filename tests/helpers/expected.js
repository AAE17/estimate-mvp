// Independent hand calculation of the kachu bill (MB) — written separately from index.html/server.js on purpose.
// Rules (owner decisions, Oct 2026): empty field = 0, no auto defaults; demolition counts only when opened
// (old drafts without the open flag: when a length was typed); gutter New C.C. only with demolition.
function mix(v) {
  const ps = String(v == null ? "" : v).split("+").filter(function (x) { return /^\s*-?[\d.]+\s*$/.test(x); }).map(parseFloat);
  return ps.length ? ps.reduce(function (a, b) { return a + b; }, 0) / ps.length : 0;
}
const num = function (v) { return Number(v || 0) || 0; };
const GPDIA = {
  pipe: [[63, 69, 12], [75, 96, 15], [90, 139, 17], [110, 199, 19]],
  gutter: [[225, 421, 88], [300, 672, 119], [450, 817, 171], [600, 1331, 228], [900, 2476, 340], [1200, 4121, 440]]
};
const GPCH = [["60", 5138], ["90", 7343], ["139", 8882], ["1313", 10698]];
const FAC = { pipe: { 63: .065, 75: .075, 90: .09, 110: .11 }, gutter: { 225: .3, 300: .35, 450: .5, 600: .65, 900: 1.0, 1200: 1.3 } };

// returns { items: {key: [qty, rate]}, sub }
function expected(c, state) {
  const t = c.type, e = {};
  if (t === "pipe" || t === "gutter") {
    const s = (state && state[t]) || c.gp;
    const dm = s.demo || {};
    const open = ("open" in dm) ? !!dm.open : mix(dm.l) > 0;
    const demo = open ? mix(dm.l) * num(dm.w) * (t === "pipe" ? 1 : num(dm.t)) : 0;
    let exc = 0; const by = {};
    s.pipes.forEach(function (p) { const L = mix(p.l); exc += L * num(p.w) * num(p.d); by[+p.dia] = (by[+p.dia] || 0) + L; });
    let ded = 0; Object.keys(by).forEach(function (d) { ded += .785 * Math.pow(FAC[t][d], 2) * by[d]; });
    const net = Math.max(0, exc - ded);
    const refill = t === "gutter" ? Math.ceil(net - 1e-9) : net;
    e.demo = [demo, t === "pipe" ? 202.2 : 1030.81]; e.exc = [exc, 89];
    GPDIA[t].forEach(function (x) { e["pipe" + x[0]] = [by[x[0]] || 0, x[1]]; });
    GPDIA[t].forEach(function (x) { e["lay" + x[0]] = [by[x[0]] || 0, x[2]]; });
    if (t === "gutter") {
      const chm = {}; s.ch.forEach(function (x) { chm[x.k] = (chm[x.k] || 0) + num(x.n); });
      GPCH.forEach(function (x) { e["ch" + x[0]] = [chm[x[0]] || 0, x[1]]; });
      const n = Object.values(chm).reduce(function (a, b) { return a + b; }, 0);
      const cc = s.cc || {};
      e.refill = [refill, 22]; e.frame = [n, 1121]; e.cover = [n, 1173];
      e.cc = [demo > 0 ? mix(cc.l) * num(cc.w) * num(cc.t) : 0, 3652.31];
    } else {
      e.refill = [refill, 22];
    }
    e.plate = [num(s.plate), 306.14];
  } else if (t === "cc") {
    const dm = c.dom || {}; const ccT = num(dm.dCcT) || 0.1;
    let box = 0, area = 0;
    c.rows.forEach(function (r) { const L = mix(r.l), W = mix(r.w); area += L * W; box += L * W * mix(r.d); });
    const metal = area * 0.12 * 1.15, mq = Math.floor(metal / 1.5 + 1e-9) * 1.5, bq = Math.floor(mq * 0.25 / 1.5 + 1e-9) * 1.5;
    Object.assign(e, { box: [box, 158.12], metal: [mq, 684.6], bind: [bq, 173.51], smetal: [mq, 249.75], sbind: [bq, 147.47],
      cc: [area * ccT, 4915.01], test: [num(dm.kTest), 2656], plate: [num(dm.kPlate), 306.14] });
    e._area = area;
  } else {
    const dm = c.dom || {}; const ex = num(dm.dExcD) || 0.2, du = num(dm.dDustD) || 0.1;
    let area = 0, va = 0;
    c.rows.forEach(function (r) { const L = mix(r.l), W = mix(r.w); area += L * W; if (L || W) va += 2 * L + 2 * W; });
    const test = (dm.kTest === undefined || dm.kTest === "") ? 1 : num(dm.kTest);
    Object.assign(e, { exc: [area * ex, 156.56], collect: [Math.floor(area * du / 1.5 + 1e-9) * 1.5, 210.42], block: [area, 740.51],
      vata: [va, 23.68], test: [test, 608], plate: [num(dm.kPlate), 306.14] });
    e._area = area; e._dust = du; e._exc = ex;
  }
  let sub = 0;
  Object.keys(e).forEach(function (k) { if (k[0] !== "_") sub += e[k][0] * e[k][1]; });
  return { items: e, sub: sub };
}

// preview label -> key
function previewKey(t, label) {
  const pairs = t === "cc"
    ? [["Box", "box"], ["Spreading Metal", "smetal"], ["Spreading binding", "sbind"], ["Metal", "metal"], ["Binding", "bind"], ["CC 1", "cc"], ["Test", "test"], ["Name", "plate"]]
    : t === "paver"
      ? [["EXCAVATION", "exc"], ["COLLECTING", "collect"], ["BLOCK", "block"], ["VATA", "vata"], ["Test", "test"], ["તકતી", "plate"]]
      : [["Demolition", "demo"], ["Excavation", "exc"], ["Refilling", "refill"], ["Name Plate", "plate"], ["Frame", "frame"], ["Cover", "cover"], ["New C.C", "cc"]];
  for (const [p, k] of pairs) if (label.indexOf(p) >= 0) return k;
  let m = label.match(/(PVC|NP2) (\d+)/); if (m) return "pipe" + m[2];
  m = label.match(/Laying (\d+)/); if (m) return "lay" + m[1];
  m = label.match(/Chamber (\S+)/); if (m) return "ch" + ({ "0.60×0.60": "60", "0.90×0.90": "90", "1.30×0.90": "139", "1.30×1.30": "1313" })[m[1]];
  return "?" + label;
}

module.exports = { expected, previewKey, mix, GPDIA };
