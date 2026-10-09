// Step 4: the core loader must serve exactly the same app as before (legacy feature mounted unchanged).
const test = require("node:test");
const assert = require("node:assert");
const fs = require("fs");
const path = require("path");
const { startApp } = require("./helpers/app");

const ROOT = path.join(__dirname, "..");

test("config/features.js lists folders that export a valid manifest", function () {
  const names = require("../config/features");
  assert.ok(Array.isArray(names) && names.indexOf("legacy") >= 0);
  names.forEach(function (n) {
    const src = fs.readFileSync(path.join(ROOT, "features", n, "index.js"), "utf8");
    assert.ok(src.indexOf('name: "' + n + '"') >= 0, n + " manifest name");
    assert.ok(/mount:\s*function/.test(src), n + " manifest mount");
  });
});

test("no repo-root paths broken: legacy server uses ROOT_DIR, not __dirname", function () {
  const src = fs.readFileSync(path.join(ROOT, "features", "legacy", "server.js"), "utf8");
  const uses = src.split("\n").filter(function (l) { return l.indexOf("__dirname") >= 0; });
  assert.deepStrictEqual(uses, ['const ROOT_DIR = path.join(__dirname, "..", "..");']);
});

let app;
test.before(async function () { app = await startApp(); });
test.after(function () { if (app) app.stop(); });

test("GET / serves the root index.html byte-for-byte", async function () {
  const r = await app.get("/");
  assert.equal(r.status, 200);
  assert.equal(r.text, fs.readFileSync(path.join(ROOT, "index.html"), "utf8"));
});
test("public pages and health still served", async function () {
  for (const p of ["/login.html", "/signup.html", "/auth-pages.js", "/gj-geo.json"]) {
    assert.equal((await app.get(p)).status, 200, p);
  }
  const h = await app.get("/api/health");
  assert.equal(h.status, 200);
  assert.equal(JSON.parse(h.text).ok, true);
});
test("unknown route is still 404", async function () {
  assert.equal((await app.get("/no-such-page-qa")).status, 404);
});
