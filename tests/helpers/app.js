// Starts a fake Supabase (auth + REST, dummy QA users only) and the real server.js on free ports.
// SOFFICE_DISABLED=1 skips LibreOffice recalculation, so Excel formula cells are not recalculated in these tests;
// the tests therefore check the values server.js itself writes.
const http = require("http");
const path = require("path");
const { spawn } = require("child_process");

const users = {
  tokOwner: { email: "owner@qa.invalid", user_metadata: {} },
  tokUser: { email: "user@qa.invalid", user_metadata: { app_role: "aae" } },
  tokPend: { email: "pend@qa.invalid", user_metadata: {} },
  tokEvil: { email: "evil@qa.invalid", user_metadata: { app_role: "super_admin" } }
};
const profiles = [
  { email: "owner@qa.invalid", role: "admin", subscription_status: "Active" },
  { email: "user@qa.invalid", role: "user", subscription_status: "Active" },
  { email: "pend@qa.invalid", role: "user", subscription_status: "Trial", subscription_end_date: "2026-01-01T00:00:00Z" },
  { email: "evil@qa.invalid", role: "admin", subscription_status: "Active" }
];

function startMock() {
  return new Promise(function (resolve) {
    const srv = http.createServer(function (req, res) {
      const u = new URL(req.url, "http://x");
      const tok = String(req.headers.authorization || "").replace(/^Bearer\s+/i, "");
      res.setHeader("Content-Type", "application/json");
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
