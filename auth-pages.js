(function () {
  var ADMIN = "vanraj2592@gmail.com";
  var sb = null;
  function msg(t) {
    var m = document.getElementById("authMsg");
    if (m) m.textContent = t || "";
  }
  function errText(s) {
    s = String(s || "");
    if (/only request this after/i.test(s)) return "થોડી સેકન્ડ રાહ જુઓ, પછી ફરી દબાવો.";
    if (/already registered/i.test(s)) return "આ ઈમેલ પહેલાંથી છે. Log in કરો.";
    if (/Invalid login/i.test(s)) return "ઈમેલ કે પાસવર્ડ ખોટો.";
    if (/confirm/i.test(s)) return "સુપાબેઝમાં Confirm email બંધ કરો.";
    return s;
  }
  async function client() {
    if (sb) return sb;
    if (!window.supabase) throw new Error("લૉગિન સેટ નથી. પેજ ફરી ખોલો.");
    var cfg = await fetch("/api/auth/config").then(function (r) { return r.json(); });
    if (!cfg.url || !cfg.anon) throw new Error("લૉગિન સેટ નથી. પેજ ફરી ખોલો.");
    sb = window.supabase.createClient(cfg.url, cfg.anon);
    return sb;
  }
  async function goIfIn() {
    try {
      var c = await client();
      var s = await c.auth.getSession();
      if (s.data && s.data.session) location.replace("index.html");
    } catch (e) {}
  }
  window.soLogin = async function () {
    var id = ((document.getElementById("id") || {}).value || "").trim();
    var password = (document.getElementById("pw") || {}).value || "";
    if (!id || !password) { msg("ઈમેલ અને પાસવર્ડ લખો"); return; }
    if (id.indexOf("@") < 0) { msg("હાલ ઈમેલથી લૉગિન થાય છે."); return; }
    msg("ચાલુ…");
    try {
      var c = await client();
      var out = await c.auth.signInWithPassword({ email: id, password: password });
      if (out.error) { msg(errText(out.error.message)); return; }
      location.replace("index.html");
    } catch (e) { msg(errText(e.message || e)); }
  };
  window.soForgot = async function (ev) {
    if (ev) ev.preventDefault();
    var id = ((document.getElementById("id") || {}).value || "").trim();
    if (id.indexOf("@") < 0) { msg("પાસવર્ડ રીસેટ માટે ઈમેલ લખો."); return; }
    msg("મોકલાઈ રહ્યું છે…");
    try {
      var c = await client();
      var out = await c.auth.resetPasswordForEmail(id, { redirectTo: location.origin + "/login.html" });
      if (out.error) { msg(errText(out.error.message)); return; }
      msg("ઈમેલ પર પાસવર્ડ રીસેટ લિંક મોકલી.");
    } catch (e) { msg(errText(e.message || e)); }
  };
  window.soSignup = async function () {
    var name = ((document.getElementById("name") || {}).value || "").trim();
    var mobile = ((document.getElementById("mobile") || {}).value || "").replace(/\s+/g, "");
    var email = ((document.getElementById("email") || {}).value || "").trim();
    var password = (document.getElementById("pw") || {}).value || "";
    var password2 = (document.getElementById("pw2") || {}).value || "";
    var role = ((document.getElementById("designation") || {}).value || "").trim();
    var jilla = ((document.getElementById("district") || {}).value || "").trim();
    var taluka = ((document.getElementById("taluka") || {}).value || "").trim();
    if (!name || !email || !password) { msg("નામ, ઈમેલ અને પાસવર્ડ લખો"); return; }
    if (password.length < 6) { msg("પાસવર્ડ ઓછામાં 6 અક્ષર"); return; }
    if (password !== password2) { msg("બંને પાસવર્ડ એકસરખા નથી"); return; }
    if (!role || !jilla || !taluka) { msg("હોદ્દો, જિલ્લો અને તાલુકો પસંદ કરો"); return; }
    msg("બની રહ્યું છે…");
    var isAd = email.toLowerCase() === ADMIN;
    try {
      var c = await client();
      var up = await c.auth.signUp({
        email: email,
        password: password,
        options: { data: { name: name, role: role, mobile: mobile, jilla: jilla, taluka: taluka, app_role: isAd ? "super_admin" : "aae", approved: isAd } }
      });
      if (up.error) { msg(errText(up.error.message)); return; }
      var inn = await c.auth.signInWithPassword({ email: email, password: password });
      if (inn.error) { msg("અકાઉન્ટ બન્યું. 30 સેકન્ડ પછી Log in દબાવો."); return; }
      try {
        var token = inn.data && inn.data.session && inn.data.session.access_token;
        await fetch("/api/admin/profile", {
          method: "POST",
          headers: { "Content-Type": "application/json", Authorization: "Bearer " + token },
          body: JSON.stringify({ email: email, name: name, role: role, mobile: mobile, jilla: jilla, taluka: taluka })
        });
      } catch (e) {}
      location.replace("index.html");
    } catch (e) { msg(errText(e.message || e)); }
  };
  goIfIn();
})();
