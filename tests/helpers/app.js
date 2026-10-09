// Starts a fake Supabase (auth + REST, dummy QA users only) and the real server.js on free ports.
// SOFFICE_DISABLED=1 skips LibreOffice recalculation, so Excel formula cells are not recalculated in these tests;
// the tests therefore check the values server.js itself writes.
const http = require("http");
const path = require("path");
const { spawn } = require("child_process");

const users = {
  tokOwner: { id: "00000000-0000-4000-8000-000000000001", email: "owner@qa.invalid", user_metadata: {} },
  tokUser: { id: "11111111-1111-4111-8111-111111111111", email: "user@qa.invalid", user_metadata: { app_role: "aae" } },
  tokUser2: { id: "22222222-2222-4222-8222-222222222222", email: "user2@qa.invalid", user_metadata: { app_role: "aae" } },
  tokPend: { id: "33333333-3333-4333-8333-333333333333", email: "pend@qa.invalid", user_metadata: {} },
  tokEvil: { id: "44444444-4444-4444-8444-444444444444", email: "evil@qa.invalid", user_metadata: { app_role: "super_admin" } }
};
const profiles = [
  { email: "owner@qa.invalid", role: "admin", subscription_status: "Active" },
  { email: "user@qa.invalid", role: "user", subscription_status: "Active" },
  { email: "user2@qa.invalid", role: "user", subscription_status: "Active" },
  { email: "pend@qa.invalid", role: "user", subscription_status: "Trial", subscription_end_date: "2026-01-01T00:00:00Z" },
  { email: "evil@qa.invalid", role: "admin", subscription_status: "Active" }
];

// In-memory fake of Supabase Storage + the "media" table (only what server.js uses).
function readBody(req) {
  return new Promise(function (resolve) { const ch = []; req.on("data", function (c) { ch.push(c); }); req.on("end", function () { resolve(Buffer.concat(ch)); }); });
}
function eqParam(u, k) { const v = u.searchParams.get(k); return v && v.indexOf("eq.") === 0 ? v.slice(3) : null; }
async function storageAndMedia(req, res, u, tok, state) {
  const p = u.pathname;
  if (p.startsWith("/storage/v1/object/sign/")) {
    const bucket = decodeURIComponent(p.slice("/storage/v1/object/sign/".length));
    const body = JSON.parse((await readBody(req)).toString() || "{}");
    state.calls.push({ op: "sign", bucket: bucket, key: tok, paths: body.paths });
    res.end(JSON.stringify((body.paths || []).map(function (pp) {
      const ok = !!state.objects[bucket + "/" + pp];
      return { path: pp, error: ok ? null : "not found", signedURL: ok ? "/object/sign/" + bucket + "/" + pp + "?token=signed-" + Math.random().toString(36).slice(2) : null };
    })));
    return true;
  }
  if (p.startsWith("/storage/v1/object/") && req.method === "POST") {
    const rest = decodeURIComponent(p.slice("/storage/v1/object/".length));
    const buf = await readBody(req);
    state.calls.push({ op: "upload", key: tok, name: rest, type: req.headers["content-type"], size: buf.length });
    state.objects[rest] = { type: req.headers["content-type"], size: buf.length };
    res.end(JSON.stringify({ Key: rest }));
    return true;
  }
  if (p === "/rest/v1/media") {
    if (req.method === "POST") {
      const row = JSON.parse((await readBody(req)).toString() || "{}");
      state.calls.push({ op: "insert", key: tok, row: row });
      state.media.push(row); res.statusCode = 201; res.end("");
      return true;
    }
    const uid = eqParam(u, "user_id"), id = eqParam(u, "id");
    let rows = state.media.filter(function (r) { return (!uid || r.user_id === uid) && (!id || r.id === id); });
    rows = rows.slice().sort(function (a, b) { return String(b.ts).localeCompare(String(a.ts)); });
    res.end(JSON.stringify(rows.slice(0, Number(u.searchParams.get("limit") || 1000))));
    return true;
  }
  return false;
}

function startMock() {
  const state = { objects: {}, media: [], calls: [] };
  return new Promise(function (resolve) {
    const srv = http.createServer(async function (req, res) {
      const u = new URL(req.url, "http://x");
      const tok = String(req.headers.authorization || "").replace(/^Bearer\s+/i, "");
      res.setHeader("Content-Type", "application/json");
      if (await storageAndMedia(req, res, u, tok, state)) return;
      if (u.pathname === "/auth/v1/user") {
        if (users[tok]) return res.end(JSON.stringify(users[tok]));
        res.statusCode = 401; return res.end("{}");
      }
      if (u.pathname === "/rest/v1/profiles") {
        const em = (u.searchParams.get("email") || "").replace(/^eq\./, "");
        let out = em ? profiles.filter(function (p) { return p.email === em; }) : profiles;
        if (req.method !== "GET") out = [{ email: em || "x" }];
        return res.end(JSON.stringify(out));
      }
      if (u.pathname.startsWith("/rest/v1/")) {
        if (req.method === "GET") return res.end("[]");
        let b = ""; req.on("data", function (c) { b += c; }); req.on("end", function () { res.end(b && b[0] === "{" ? "[" + b + "]" : "[]"); });
        return;
      }
      res.statusCode = 404; res.end("{}");
    });
    srv.state = state;
    srv.listen(0, "127.0.0.1", function () { resolve(srv); });
  });
}

async function startApp() {
  const mock = await startMock();
  const port = 20000 + Math.floor(Math.random() * 20000);
  const env = Object.assign({}, process.env, {
    PORT: String(port), SOFFICE_DISABLED: "1", ADMIN_EMAIL: "owner@qa.invalid",
    SUPABASE_URL: "http://127.0.0.1:" + mock.address().port, SUPABASE_ANON_KEY: "anon", SUPABASE_SERVICE_ROLE: "service"
  });
  delete env.GEMINI_API_KEY;
  const child = spawn(process.execPath, [path.join(__dirname, "..", "..", "server.js")], { env: env, stdio: ["ignore", "pipe", "pipe"] });
  let log = "";
  await new Promise(function (resolve, reject) {
    const t = setTimeout(function () { reject(new Error("server did not start:\n" + log)); }, 30000);
    const on = function (d) { log += d; if (log.indexOf(" on " + port) >= 0) { clearTimeout(t); resolve(); } };
    child.stdout.on("data", on); child.stderr.on("data", on);
    child.on("exit", function (code) { clearTimeout(t); reject(new Error("server exited " + code + "\n" + log)); });
  });
  const base = "http://127.0.0.1:" + port;
  return {
    base: base,
    sb: mock.state,
    post: async function (p, body, token) {
      const h = { "Content-Type": "application/json" };
      if (token) h.Authorization = "Bearer " + token;
      const r = await fetch(base + p, { method: "POST", headers: h, body: JSON.stringify(body) });
      return { status: r.status, json: await r.json() };
    },
    get: async function (p, token) {
      const r = await fetch(base + p, { headers: token ? { Authorization: "Bearer " + token } : {} });
      return { status: r.status, text: await r.text() };
    },
    download: async function (p) { const r = await fetch(base + p); return Buffer.from(await r.arrayBuffer()); },
    stop: function () { child.removeAllListeners("exit"); child.kill(); mock.close(); }
  };
}

module.exports = { startApp };
