/* global Office, Excel, CfpCard */
// Excel side panel: shows the rewrite-options card for the policy on the selected row.
// Works on any sheet that has a "Policy #" header in its first 3 rows (Compare, Quote Options,
// Work Queue, No Option). All data is read from the open workbook; nothing is sent anywhere.
(function () {
  "use strict";

  var QO_SHEET = "Quote Options";
  var COMPARE_TABLE = "Compare";
  var state = { layout: null, qoCols: 0, qoIndex: null, policyCol: {}, timer: null, last: "" };

  function $(id) { return document.getElementById(id); }
  function show(html) { $("card").innerHTML = html; }
  function msg(text) { show('<div class="msg">' + text + "</div>"); }
  function norm(v) { return v === null || v === undefined ? "" : String(v).trim(); }

  Office.onReady(function (info) {
    if (info.host !== Office.HostType.Excel) { msg("Open this in Excel."); return; }
    $("refresh").addEventListener("click", function () { state.layout = null; state.qoIndex = null; state.policyCol = {}; state.last = ""; update(); });
    Office.context.document.addHandlerAsync(Office.EventType.DocumentSelectionChanged, function () {
      clearTimeout(state.timer);
      state.timer = setTimeout(update, 200);
    });
    update();
  });

  function loadQuoteOptions(ctx) {
    var ws = ctx.workbook.worksheets.getItem(QO_SHEET);
    var used = ws.getUsedRange(true);
    used.load("rowCount,columnCount");
    return ctx.sync().then(function () {
      var head = ws.getRangeByIndexes(0, 0, 2, used.columnCount);
      var colA = ws.getRangeByIndexes(0, 0, used.rowCount, 1);
      head.load("values"); colA.load("values");
      return ctx.sync().then(function () {
        state.qoCols = used.columnCount;
        state.layout = CfpCard.buildLayout(head.values[0], head.values[1]);
        state.qoIndex = {};
        colA.values.forEach(function (r, i) { var p = norm(r[0]); if (p && i >= 2 && state.qoIndex[p] === undefined) state.qoIndex[p] = i; });
      });
    });
  }

  function findPolicyCol(ctx, sheet, name) {
    if (state.policyCol[name] !== undefined) return Promise.resolve(state.policyCol[name]);
    var used = sheet.getUsedRange(true);
    used.load("columnCount");
    return ctx.sync().then(function () {
      var top = sheet.getRangeByIndexes(0, 0, 3, Math.min(used.columnCount, 250));
      top.load("values");
      return ctx.sync().then(function () {
        var col = -1;
        top.values.forEach(function (r) { r.forEach(function (v, j) { if (col < 0 && norm(v) === "Policy #") col = j; }); });
        state.policyCol[name] = col;
        return col;
      });
    });
  }

  // Team-owned columns on the Compare table, matched by Policy #.
  function readTeamColumns(ctx, policy) {
    var empty = { status: "", direction: "" };
    var t = ctx.workbook.tables.getItemOrNullObject(COMPARE_TABLE);
    t.load("isNullObject");
    return ctx.sync().then(function () {
      if (t.isNullObject) return empty;
      var pol = t.columns.getItemOrNullObject("Policy #");
      var st = t.columns.getItemOrNullObject("Status");
      var dir = t.columns.getItemOrNullObject("Agent Direction");
      pol.load("isNullObject"); st.load("isNullObject"); dir.load("isNullObject");
      return ctx.sync().then(function () {
        if (pol.isNullObject) return empty;
        var a = pol.getDataBodyRange(); a.load("values");
        var b = st.isNullObject ? null : st.getDataBodyRange(); if (b) b.load("values");
        var c = dir.isNullObject ? null : dir.getDataBodyRange(); if (c) c.load("values");
        return ctx.sync().then(function () {
          for (var i = 0; i < a.values.length; i++) {
            if (norm(a.values[i][0]) === policy) {
              return { status: b ? norm(b.values[i][0]) : "", direction: c ? norm(c.values[i][0]) : "" };
            }
          }
          return empty;
        });
      });
    });
  }

  function update() {
    Excel.run(function (ctx) {
      var cell = ctx.workbook.getActiveCell();
      var sheet = ctx.workbook.worksheets.getActiveWorksheet();
      cell.load("rowIndex"); sheet.load("name");
      var policy = "", rowIndex = 0;
      return ctx.sync()
        .then(function () { if (!state.layout) return loadQuoteOptions(ctx); })
        .then(function () { rowIndex = cell.rowIndex; return findPolicyCol(ctx, sheet, sheet.name); })
        .then(function (col) {
          if (col < 0) return null;
          var c = sheet.getRangeByIndexes(rowIndex, col, 1, 1);
          c.load("values");
          return ctx.sync().then(function () { return norm(c.values[0][0]); });
        })
        .then(function (p) {
          policy = p || "";
          if (!/^CFP /i.test(policy)) {
            state.last = "";
            msg("Click any row with a FAIR Plan policy number (Compare, Quote Options, Work Queue or No Option) to see its rewrite options.");
            $("info").textContent = "";
            return null;
          }
          if (policy === state.last) return null;
          var r = state.qoIndex[policy];
          if (r === undefined) return loadQuoteOptions(ctx).then(function () { return state.qoIndex[policy]; });
          return r;
        })
        .then(function (r) {
          if (r === null) return;
          if (r === undefined) { msg(policy + " isn't on the Quote Options sheet."); return; }
          var rng = ctx.workbook.worksheets.getItem(QO_SHEET).getRangeByIndexes(r, 0, 1, state.qoCols);
          rng.load("values");
          return ctx.sync().then(function () {
            var row = rng.values[0];
            if (norm(row[0]) !== policy) { state.qoIndex = null; state.layout = null; msg("Rows moved. Click Refresh."); return; }
            return readTeamColumns(ctx, policy).then(function (team) {
              var card = CfpCard.buildCard(state.layout, row, team, new Date());
              show(CfpCard.renderCard(card));
              state.last = policy;
              $("info").textContent = sheet.name + " row " + (rowIndex + 1);
            });
          });
        });
    }).catch(function (e) {
      msg("Couldn't read the workbook: " + (e && e.message ? e.message : e) + ". Click Refresh.");
    });
  }
})();
