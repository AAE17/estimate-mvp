// index.html must not define the same function twice (the last copy silently wins and hides the good one).
const test = require("node:test");
const assert = require("node:assert");
const fs = require("fs");
const path = require("path");

const html = fs.readFileSync(path.join(__dirname, "..", "index.html"), "utf8");
const KNOWN = ["fromGu"]; // existing duplicate, identical copies; remove from this list when cleaned up

test("index.html: no function defined twice", function () {
  const seen = {};
  html.split("\n").forEach(function (l, i) {
    const m = l.match(/^\s*(?:async\s+)?function\s+([A-Za-z_$][\w$]*)\s*\(/);
    if (m) (seen[m[1]] = seen[m[1]] || []).push(i + 1);
  });
  const dup = Object.keys(seen).filter(function (k) { return seen[k].length > 1 && KNOWN.indexOf(k) < 0; })
    .map(function (k) { return k + " @ " + seen[k].join(","); });
  assert.deepStrictEqual(dup, []);
});

test("site photo/sketch save syncs to /api/db/media", function () {
  const i = html.indexOf("function pushMedia(");
  const body = html.slice(i, html.indexOf("function renderMedia(", i));
  assert.ok(body.indexOf('"/api/db/media"') > 0, "pushMedia must POST to /api/db/media");
});

test("tour diary month is not hard-coded", function () {
  assert.ok(!/tourYM\s*=\s*\{\s*y\s*:\s*\d{4}/.test(html), "tourYM must start from the current month");
});
