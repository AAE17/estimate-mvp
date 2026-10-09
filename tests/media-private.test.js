// Site photos/sketches are private: each user sees and opens ONLY their own files (bucket user-media).
const test = require("node:test");
const assert = require("node:assert");
const { startApp } = require("./helpers/app");

const A = "11111111-1111-4111-8111-111111111111"; // tokUser
const B = "22222222-2222-4222-8222-222222222222"; // tokUser2
const JPG = "data:image/jpeg;base64," + Buffer.from("fake-jpeg-bytes").toString("base64");
const UUID = "[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}";

let app, idA;
test.before(async function () {
  app = await startApp();
  const r1 = await app.post("/api/db/media", { kind: "photo", work_name: "QA A road", gps: "22.1, 72.9", data: JPG }, "tokUser");
  const r2 = await app.post("/api/db/media", { kind: "plan", work_name: "QA A road", data: JPG }, "tokUser");
  assert.equal(r1.status, 200, JSON.stringify(r1.json)); assert.equal(r2.status, 200);
  idA = r1.json.id;
});
test.after(function () { if (app) app.stop(); });

test("no login -> 401 (upload, list, signed url)", async function () {
  assert.equal((await app.post("/api/db/media", { kind: "photo", data: JPG })).status, 401);
  assert.equal((await app.get("/api/db/media")).status, 401);
  assert.equal((await app.get("/api/db/media/" + idA + "/url")).status, 401);
});
test("approval pending user -> 403", async function () {
  assert.equal((await app.post("/api/db/media", { kind: "photo", data: JPG }, "tokPend")).status, 403);
});
test("old public folder /db/media is closed (even if a file is on disk)", async function () {
  const fs = require("fs"), path = require("path");
  const dir = path.join(__dirname, "..", "db", "media");
  fs.mkdirSync(dir, { recursive: true });
  const f = path.join(dir, "qa-old-public.jpg");
  fs.writeFileSync(f, "x");
  try {
    const r = await app.get("/db/media/qa-old-public.jpg");
    assert.equal(r.status, 410);
  } finally { fs.unlinkSync(f); }
});
test("upload goes to private bucket, own folder, random name, service key", async function () {
  const ups = app.sb.calls.filter(function (c) { return c.op === "upload"; });
  assert.equal(ups.length, 2);
  ups.forEach(function (c) {
    assert.match(c.name, new RegExp("^user-media/" + A + "/(photo|plan)/" + UUID + "\\.jpg$"));
    assert.equal(c.key, "service");
  });
  assert.notEqual(ups[0].name.split("/").pop(), ups[1].name.split("/").pop());
  const ins = app.sb.calls.filter(function (c) { return c.op === "insert"; });
  ins.forEach(function (c) { assert.equal(c.row.user_id, A); assert.equal(c.row.url, ""); assert.ok(c.row.path.startsWith(A + "/")); });
});
test("user id in body cannot change the folder", async function () {
  const r = await app.post("/api/db/media", { kind: "photo", data: JPG, user_id: B, path: B + "/photo/x.jpg", user_email: "user2@qa.invalid" }, "tokUser");
  assert.equal(r.status, 200);
  const last = app.sb.calls.filter(function (c) { return c.op === "upload"; }).pop();
  assert.ok(last.name.startsWith("user-media/" + A + "/"), last.name);
  const row = app.sb.media.find(function (m) { return m.id === r.json.id; });
  assert.equal(row.user_id, A);
});
test("bad kind / not an image -> 400", async function () {
  assert.equal((await app.post("/api/db/media", { kind: "../../x", data: JPG }, "tokUser")).status, 400);
  assert.equal((await app.post("/api/db/media", { kind: "photo", data: "data:text/html;base64,PGI+" }, "tokUser")).status, 400);
});
test("owner A lists own files with short-lived signed URLs (no public URLs)", async function () {
  const r = await app.get("/api/db/media", "tokUser");
  assert.equal(r.status, 200);
  const items = JSON.parse(r.text).items;
  assert.equal(items.length, 3);
  items.forEach(function (m) {
    assert.match(m.url, new RegExp("/storage/v1/object/sign/user-media/" + A + "/"));
    assert.ok(m.url.indexOf("token=") > 0);
    assert.ok(m.url.indexOf("/object/public/") < 0);
  });
});
test("user B cannot list or open A's files", async function () {
  const r = await app.get("/api/db/media", "tokUser2");
  assert.equal(r.status, 200);
  assert.deepStrictEqual(JSON.parse(r.text).items, []);
  assert.equal((await app.get("/api/db/media/" + idA + "/url", "tokUser2")).status, 404);
  const bundle = JSON.parse((await app.get("/api/db/bundle?work=" + encodeURIComponent("QA A road"), "tokUser2")).text);
  assert.deepStrictEqual(bundle.media || [], []);
});
test("A gets a signed URL for own file only", async function () {
  const r = await app.get("/api/db/media/" + idA + "/url", "tokUser");
  assert.equal(r.status, 200);
  assert.match(JSON.parse(r.text).url, new RegExp("/object/sign/user-media/" + A + "/photo/" + UUID + "\\.jpg\\?token="));
  assert.equal((await app.get("/api/db/media/not-a-real-id/url", "tokUser")).status, 404);
});
test("signing requests only ever contain the caller's own paths", async function () {
  app.sb.calls.filter(function (c) { return c.op === "sign"; }).forEach(function (c) {
    assert.equal(c.bucket, "user-media");
    (c.paths || []).forEach(function (p) { assert.ok(p.startsWith(A + "/"), p); });
  });
});
