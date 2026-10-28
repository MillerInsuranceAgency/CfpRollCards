// Builds the rewrite-options card from one Quote Options row.
// Pure functions only (no Office.js), so it can be tested outside Excel.
(function (root) {
  "use strict";

  var CARRIERS = ["Farmers", "Foremost", "Aegis", "Bamboo", "Delos", "Lightspeed", "Obie", "SageSure", "Slide", "Solara", "Swyfft"];
  var FIXED_LABEL = { Delos: "indication", Lightspeed: "indication", SageSure: "prequote", Slide: "prequote", Solara: "estimate" };
  var FF = { Farmers: true, Foremost: true };

  function clean(v) { return typeof v === "string" ? v.trim() : v; }
  function num(v) { return typeof v === "number" && isFinite(v) && v !== 0 ? v : null; }
  function txt(v) { return v === null || v === undefined ? "" : String(v).trim(); }

  // Map header names to column indexes. h1 = group row (row 1), h2 = field row (row 2).
  function buildLayout(h1, h2) {
    h1 = h1.map(clean); h2 = h2.map(clean);
    var field = {};
    var starts = [];
    h1.forEach(function (v, i) { if (v) starts.push(i); });
    var blocks = {};
    CARRIERS.forEach(function (c) {
      var s = h1.indexOf(c);
      if (s < 0) return;
      var next = starts.filter(function (x) { return x > s; })[0];
      var e = next === undefined ? h2.length : next;
      var b = {};
      for (var i = s; i < e; i++) { if (h2[i] && b[h2[i]] === undefined) b[h2[i]] = i; }
      blocks[c] = b;
    });
    // First occurrence of each field name outside the carrier blocks.
    var inBlock = {};
    Object.keys(blocks).forEach(function (c) { Object.keys(blocks[c]).forEach(function (k) { inBlock[blocks[c][k]] = true; }); });
    h2.forEach(function (v, i) { if (v && !inBlock[i] && field[v] === undefined) field[v] = i; });
    return { field: field, blocks: blocks };
  }

  function excelDate(v) {
    if (typeof v === "number" && v > 20000) return new Date(Date.UTC(1899, 11, 30) + Math.round(v) * 86400000);
    if (v instanceof Date) return v;
    return null;
  }

  function buildCard(layout, row, extra, today) {
    extra = extra || {};
    today = today || new Date();
    var F = layout.field;
    function g(name) { var i = F[name]; return i === undefined ? null : row[i]; }

    var exp = num(g("CFP Rate"));
    var ren = num(g("CFP Renewal Premium"));
    var dicCarrier = txt(g("DIC Carrier"));
    var dicPrem = num(g("DIC Rate"));
    var base = ren !== null ? ren : exp;
    var total = base === null ? null : base + (dicPrem || 0);
    var rd = excelDate(g("Renewal Date"));
    var days = null;
    if (rd) {
      var t = Date.UTC(today.getFullYear(), today.getMonth(), today.getDate());
      days = Math.round((rd.getTime() - t) / 86400000);
    }
    var occ = txt(g("Occupancy")).replace(" Occupied", "");
    if (occ === "Primary") occ = "Owner";

    var full = [], dic = [];
    CARRIERS.forEach(function (c) {
      var b = layout.blocks[c];
      if (!b || b.Premium === undefined) return;
      var prem = num(row[b.Premium]);
      if (prem === null || prem < 0) return;
      var status = txt(row[b.Status]);
      var type = b.Type !== undefined ? txt(row[b.Type]) : "Full";
      var opt = {
        carrier: c,
        label: FIXED_LABEL[c] || (status === "Prequoted" ? "prequote" : ""),
        premium: prem,
        quote: b["Quote #"] !== undefined ? txt(row[b["Quote #"]]) : "",
        deductible: b.Deductible !== undefined ? num(row[b.Deductible]) : null,
        dwelling: b.Dwelling !== undefined ? num(row[b.Dwelling]) : null
      };
      if (type === "DIC") {
        opt.diff = dicCarrier && dicPrem !== null ? prem - dicPrem : null;
        dic.push(opt);
      } else {
        opt.diff = total !== null ? prem - total : null;
        full.push(opt);
      }
    });
    full.sort(function (a, b) { return a.premium - b.premium; });
    dic.sort(function (a, b) { return a.premium - b.premium; });

    return {
      policy: txt(g("Policy #")),
      agent: txt(g("Agent")),
      name: txt(g("Account Name")),
      address: txt(g("Property Address")),
      sfUrl: txt(g("SF Account URL")),
      mapUrl: txt(g("Map URL")),
      propertyType: txt(g("Property Type")),
      occupancy: occ,
      status: txt(extra.status),
      direction: txt(extra.direction),
      phone: txt(g("Phone")) || txt(g("Primary Phone")),
      email: txt(g("Email")) || txt(g("Primary Email")),
      yearBuilt: num(g("Year Built")),
      construction: txt(g("Construction")),
      renewalDate: rd,
      daysToRenewal: days,
      expiring: exp,
      renewal: ren,
      increase: ren !== null && exp !== null ? ren - exp : null,
      increasePct: ren !== null && exp ? ren / exp - 1 : null,
      dicCarrier: dicCarrier,
      dicPremium: dicPrem,
      dicPolicy: txt(g("DIC Policy #")),
      dicAlreadyFF: !!FF[dicCarrier],
      total: total,
      full: full,
      dic: dic,
      notes: txt(g("Notes"))
    };
  }

  // ---------- HTML ----------
  function esc(s) {
    return String(s === null || s === undefined ? "" : s).replace(/[&<>"']/g, function (c) {
      return { "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[c];
    });
  }
  function money(v) { return v === null || v === undefined ? "?" : "$" + Math.round(Math.abs(v)).toLocaleString("en-US"); }
  function signed(v) { return (v < 0 ? "-" : "+") + money(v); }
  function mdy(d) {
    if (!d) return "?";
    return ("0" + (d.getUTCMonth() + 1)).slice(-2) + "/" + ("0" + d.getUTCDate()).slice(-2) + "/" + d.getUTCFullYear();
  }
  function safeUrl(u) { return /^https:\/\//i.test(u) ? u : ""; }

  function optionRow(o, dicMode, isNewDic) {
    var cls = o.diff === null ? "neutral" : (o.diff <= 0 ? "good" : "bad");
    if (dicMode && isNewDic) cls = "new";
    var meta = [];
    if (o.deductible) meta.push(money(o.deductible) + " ded");
    if (o.quote) meta.push("#" + o.quote);
    var right = money(o.premium);
    var sub = dicMode && isNewDic ? "new DIC" : (o.diff === null ? "" : signed(o.diff));
    return '<div class="opt ' + cls + '">' +
      '<div class="opt-l"><span class="opt-name">' + esc(o.carrier + (dicMode ? " DIC" : "")) + "</span>" +
      (o.label ? ' <span class="tag">' + esc(o.label) + "</span>" : "") +
      (meta.length ? '<div class="opt-meta">' + esc(meta.join(" · ")) + "</div>" : "") + "</div>" +
      '<div class="opt-r"><div class="opt-prem">' + right + "</div>" +
      (sub ? '<div class="opt-diff">' + esc(sub) + "</div>" : "") + "</div></div>";
  }

  function renderCard(c) {
    var h = "";
    var sf = safeUrl(c.sfUrl), map = safeUrl(c.mapUrl);
    h += '<div class="head">';
    h += sf ? '<a class="name" href="' + esc(sf) + '" target="_blank" rel="noopener">' + esc(c.name) + "</a>" : '<div class="name">' + esc(c.name) + "</div>";
    h += '<div class="addr">' + (map ? '<a href="' + esc(map) + '" target="_blank" rel="noopener">' + esc(c.address) + "</a>" : esc(c.address)) + "</div>";
    h += '<div class="chips">';
    h += '<span class="chip">' + esc(c.agent) + "</span>";
    if (c.propertyType) h += '<span class="chip">' + esc(c.propertyType) + "</span>";
    if (c.occupancy) h += '<span class="chip">' + esc(c.occupancy) + "</span>";
    if (c.status) h += '<span class="chip status">' + esc(c.status) + "</span>";
    h += "</div>";
    var contact = [];
    if (c.phone) contact.push('<a href="tel:' + esc(c.phone.replace(/[^0-9+]/g, "")) + '">' + esc(c.phone) + "</a>");
    if (c.email) contact.push('<a href="mailto:' + esc(c.email) + '">' + esc(c.email) + "</a>");
    if (contact.length) h += '<div class="contact">' + contact.join(" · ") + "</div>";
    if (c.direction) h += '<div class="direction"><span class="d-lbl">Agent direction</span>' + esc(c.direction) + "</div>";
    h += "</div>";

    // Tiles
    var dueTxt = c.daysToRenewal === null ? "" : (c.daysToRenewal >= 0 ? "in " + c.daysToRenewal + " days" : "renewed " + (-c.daysToRenewal) + " days ago");
    h += '<div class="tiles">';
    h += '<div class="tile"><div class="t-lbl">FAIR Plan renewal ' + esc(mdy(c.renewalDate)) + "</div>";
    if (c.renewal !== null) {
      h += '<div class="t-val">' + money(c.renewal) + "</div>";
      if (c.increase !== null) h += '<div class="t-sub ' + (c.increase > 0 ? "up" : "down") + '">' + signed(c.increase) + " / " + (c.increasePct >= 0 ? "+" : "") + Math.round(c.increasePct * 100) + "%</div>";
    } else {
      h += '<div class="t-val">' + money(c.expiring) + '</div><div class="t-sub muted">expiring, renewal not out yet</div>';
    }
    h += '<div class="t-sub muted">' + esc(c.policy) + (dueTxt ? " · " + dueTxt : "") + "</div></div>";
    h += '<div class="tile"><div class="t-lbl">Current DIC</div>';
    if (c.dicCarrier) {
      h += '<div class="t-val">' + (c.dicPremium !== null ? money(c.dicPremium) : "?") + "</div>";
      h += '<div class="t-sub">' + esc(c.dicCarrier) + (c.dicPremium === null ? " (premium missing)" : "") + "</div>";
    } else {
      h += '<div class="t-val">None</div><div class="t-sub muted">no DIC on file</div>';
    }
    h += "</div></div>";
    h += '<div class="total">Current total <b>' + money(c.total) + "</b></div>";

    // Options
    h += '<div class="sec">Full policy <span class="muted">vs current total</span></div>';
    h += c.full.length ? c.full.map(function (o) { return optionRow(o, false); }).join("") : '<div class="none">No full options</div>';
    h += '<div class="sec">DIC <span class="muted">vs current DIC</span></div>';
    if (c.dicAlreadyFF) {
      h += '<div class="opt good"><div class="opt-l"><span class="opt-name">Already with ' + esc(c.dicCarrier) + '</span></div><div class="opt-r"><div class="opt-diff">no change needed</div></div></div>';
    } else if (c.dic.length) {
      h += c.dic.map(function (o) { return optionRow(o, true, !c.dicCarrier); }).join("");
    } else {
      h += '<div class="none">None quoted' + (c.full.length ? " (full option quoted instead)" : "") + "</div>";
    }
    if (c.notes) h += '<details class="notes"><summary>Quote notes</summary><div>' + esc(c.notes).replace(/ \| /g, "<br>") + "</div></details>";
    return h;
  }

  var api = { buildLayout: buildLayout, buildCard: buildCard, renderCard: renderCard, CARRIERS: CARRIERS };
  if (typeof module !== "undefined" && module.exports) module.exports = api; else root.CfpCard = api;
})(this);
