// Private APIs must need login; logs are admin-only (admin = ADMIN_EMAIL only).
const test = require("node:test");
const assert = require("node:assert");
const { startApp } = require("./helpers/app");

let app;
test.before(async function () { app = await startApp(); });
test.after(function () { if (app) app.stop(); });

for (const p of ["/api/db/health", "/api/bills", "/api/stats", "/api/logs", "/api/db/estimates", "/api/db/site", "/api/db/kachu"]) {
  test("GET " + p + " without login -> 401", async function () {
    assert.equal((await app.get(p)).status, 401);
  });
}
test("expired Trial account -> 403 on /api/db/health", async function () {
  assert.equal((await app.get("/api/db/health", "tokPend")).status, 403);
});
test("/api/logs: normal user 403, profiles.role=admin is NOT admin 403, ADMIN_EMAIL 200", async function () {
  assert.equal((await app.get("/api/logs", "tokUser")).status, 403);
  assert.equal((await app.get("/api/logs", "tokEvil")).status, 403);
  assert.equal((await app.get("/api/logs", "tokOwner")).status, 200);
});
test("/api/db/health with active user -> 200", async function () {
  assert.equal((await app.get("/api/db/health", "tokUser")).status, 200);
});
