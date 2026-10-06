/**
 * SAC Custom Widget - Excel Export Widget v1.1.0
 * Exports a bound SAC table to formatted Excel.
 * Hierarchy, frozen header, logo, totals, number formats.
 * SheetJS (xlsx-js-style) loaded from jsDelivr CDN on first export click.
 */
(function () {
  "use strict";

  var XLSX_CDN = "https://cdn.jsdelivr.net/npm/xlsx-js-style@1.2.0/dist/xlsx.bundle.js";
  var MAX_OUTLINE = 7;
  var C = {
    HDR_BG:"FFFF00", HDR_FG:"000000",
    TOT_BG:"D9D9D9", TOT_FG:"000000",
    ODD:"F5F5F5",    EVEN:"FFFFFF",
    HIER:"FAFAFA",   BDR:"CCCCCC"
  };
  var TB = { style:"thin", color:{ rgb:C.BDR } };
  var FB = { top:TB, bottom:TB, left:TB, right:TB };

  /* Shadow-DOM template (CSS inlined) */
  var tmpl = document.createElement("template");
  tmpl.innerHTML = [
    "<style>",
    ":host{display:flex;align-items:center;justify-content:center;",
          "width:100%;height:100%;box-sizing:border-box;",
          "font-family:'72','SAP 72',Arial,sans-serif}",
    "#wrap{display:flex;flex-direction:column;align-items:center;",
           "justify-content:center;gap:6px;width:100%;padding:4px}",
    "#btn{display:inline-flex;align-items:center;justify-content:center;",
         "gap:6px;padding:8px 20px;min-width:120px;border:none;border-radius:4px;",
         "font-size:14px;font-weight:600;cursor:pointer;white-space:nowrap;",
         "background:#0070F2;color:#fff;outline:none;",
         "transition:opacity .18s,box-shadow .18s}",
    "#btn:hover:not([disabled]){opacity:.88;box-shadow:0 2px 8px rgba(0,0,0,.22)}",
    "#btn:active:not([disabled]){opacity:.72}",
    "#btn[disabled]{opacity:.48;cursor:not-allowed}",
    "#msg{font-size:11px;color:#666;min-height:14px;text-align:center}",
    "</style>",
    '<div id="wrap">',
      '<button id="btn">Export to Excel</button>',
      '<span id="msg"></span>',
    "</div>"
  ].join("");

  /* Web Component */
  function ExcelWidget() {
    var self = HTMLElement.call(this) || this;
    self._sh  = self.attachShadow({ mode:"open" });
    self._sh.appendChild(tmpl.content.cloneNode(true));
    self._props = {
      fileNamePrefix:"SAC_Export", logoUrl:"",
      buttonLabel:"Export to Excel", buttonColor:"#0070F2", buttonTextColor:"#FFFFFF",
      headerBgColor:"FFFF00", tableSubtitle:"Export", username:"",
      includeGrandTotal:true, logoRowHeight:60
    };
    self._xlReady = false;
    self._xlLoading = false;
    self.tableDataBinding = null;
    return self;
  }
  ExcelWidget.prototype = Object.create(HTMLElement.prototype);
  ExcelWidget.prototype.constructor = ExcelWidget;

  ExcelWidget.prototype.connectedCallback = function () {
    var self = this;
    self._btn = self._sh.getElementById("btn");
    self._msg = self._sh.getElementById("msg");
    if (self._btn) {
      self._btn.addEventListener("click", function () { self.exportToExcel(); });
      self._applyBtn();
    }
    self._loadXLSX().catch(function () {});
  };

  ExcelWidget.prototype.onCustomWidgetBeforeUpdate = function () {};

  ExcelWidget.prototype.onCustomWidgetAfterUpdate = function (cp) {
    var self = this;
    Object.keys(self._props).forEach(function (k) {
      if (Object.prototype.hasOwnProperty.call(cp, k)) { self._props[k] = cp[k]; }
    });
    if (self._btn) { self._applyBtn(); }
  };

  ExcelWidget.prototype.onCustomWidgetDataChanged = function () {
    if (self._msg) { self._msg.textContent = ""; }
  };

  /* Public method - callable from SAC scripting */
  ExcelWidget.prototype.exportToExcel = function () {
    var self = this;
    if (self._btn) { self._btn.disabled = true; }
    self._setMsg("Preparing\u2026");
    self._loadXLSX()
      .then(function () { return self._runExport(); })
      .then(function () {
        self._setMsg("Downloaded \u2713");
        setTimeout(function () { self._setMsg(""); }, 4000);
      })
      .catch(function (e) {
        console.error("[ExcelExportWidget]", e);
        self._setMsg("Error \u2013 see console");
        self.dispatchEvent(new CustomEvent("onExportError",
          { bubbles:true, detail:{ error:String(e.message || e) } }));
      })
      .then(function () {
        if (self._btn) { self._btn.disabled = false; }
      });
  };

  /* Button styling */
  ExcelWidget.prototype._applyBtn = function () {
    if (!this._btn) { return; }
    this._btn.textContent = this._props.buttonLabel || "Export to Excel";
    this._btn.style.backgroundColor = this._props.buttonColor    || "#0070F2";
    this._btn.style.color            = this._props.buttonTextColor|| "#FFFFFF";
  };

  ExcelWidget.prototype._setMsg = function (m) {
    if (this._msg) { this._msg.textContent = m; }
  };

  /* SheetJS CDN loader */
  ExcelWidget.prototype._loadXLSX = function () {
    var self = this;
    if (window.XLSXStyle) { return Promise.resolve(); }
    if (self._xlLoading) {
      return new Promise(function (res, rej) {
        var t0 = Date.now();
        var id = setInterval(function () {
          if (window.XLSXStyle) { clearInterval(id); res(); }
          else if (Date.now() - t0 > 20000) { clearInterval(id); rej(new Error("SheetJS load timeout")); }
        }, 100);
      });
    }
    self._xlLoading = true;
    return new Promise(function (res, rej) {
      var s = document.createElement("script");
      s.src = XLSX_CDN;
      s.onload  = function () { self._xlLoading = false; self._xlReady = true; res(); };
      s.onerror = function () { self._xlLoading = false; rej(new Error("Cannot load SheetJS CDN. Check CSP settings.")); };
      document.head.appendChild(s);
    });
  };

  /* Core export */
  ExcelWidget.prototype._runExport = function () {
    var self = this;
    var binding = self.tableDataBinding;
    if (!binding) {
      return Promise.reject(new Error("No data binding. Bind a data source in the widget properties panel."));
    }
    return Promise.resolve(binding.getData ? binding.getData() : binding)
      .then(function (result) {
        if (!result || !result.data) {
          throw new Error("Data binding returned no data.");
        }
        var X    = window.XLSXStyle;
        var meta = result.metadata || {};
        var fDim = (meta.feeds && meta.feeds.dimensions && meta.feeds.dimensions.members) || [];
        var fMeas= (meta.feeds && meta.feeds.measures  && meta.feeds.measures.members)   || [];
        var vDim = fDim.filter(function (m) { return m.visible !== false; });
        var vMeas= fMeas.filter(function (m) { return m.visible !== false; });
        var cols = vDim.concat(vMeas);
        if (!cols.length) { throw new Error("No visible columns in the data binding."); }

        var cMap = new Map();
        fDim.forEach(function (m, i) { cMap.set(m.id, i); });
        fMeas.forEach(function (m, i) { cMap.set(m.id, fDim.length + i); });

        var raw     = result.data;
        var hasHier = self._detectHier(raw);
        var hasLogo = !!(self._props.logoUrl && self._props.logoUrl.trim());

        /* Build AOA */
        var aoa = [], totFlags = [], lvlFlags = [], drillFlags = [];

        if (hasLogo) {
          aoa.push(new Array(cols.length).fill(""));
          totFlags.push(false); lvlFlags.push(0); drillFlags.push("leaf");
        }
        var HR = aoa.length; /* header row index */

        /* Header */
        aoa.push(cols.map(function (c) { return c.description || c.label || c.id || ""; }));
        totFlags.push(false); lvlFlags.push(0); drillFlags.push("leaf");

        /* Data rows */
        raw.forEach(function (row) {
          var isTotal = self._isTotal(row, fMeas);
          if (isTotal && !self._props.includeGrandTotal) { return; }
          var lv = hasHier ? self._rowLevel(row) : 0;
          var dr = hasHier ? self._rowDrill(row) : "leaf";
          aoa.push(cols.map(function (col) {
            var idx = cMap.get(col.id);
            if (idx === undefined) { return ""; }
            var cell = row[idx];
            if (cell === null || cell === undefined) { return ""; }
            var isMeas = fMeas.some(function (m) { return m.id === col.id; });
            if (isMeas) {
              var raw2 = (cell.value !== undefined) ? cell.value : cell;
              var num  = Number(raw2);
              return isNaN(num) ? (cell.formattedValue || cell.value || "") : num;
            }
            if (cell.formattedValue !== undefined) { return cell.formattedValue; }
            if (cell.value         !== undefined) { return cell.value; }
            return String(cell);
          }));
          totFlags.push(isTotal); lvlFlags.push(lv); drillFlags.push(dr);
        });

        /* Create worksheet */
        var ws  = X.utils.aoa_to_sheet(aoa);
        var nR  = aoa.length, nC = cols.length;
        ws["!ref"]    = X.utils.encode_range({ r:0, c:0 }, { r:nR-1, c:nC-1 });
        ws["!cols"]   = cols.map(function (c) {
          return { wch: Math.max(14, (c.description || c.label || c.id || "").length + 6) };
        });
        ws["!freeze"] = { xSplit:0, ySplit:HR + 1 };
        ws["!merges"] = hasLogo
          ? [{ s:{ r:0, c:0 }, e:{ r:0, c:nC-1 } }]
          : [];

        /* Row heights */
        var rh = [];
        if (hasLogo) { rh[0] = { hpt: Number(self._props.logoRowHeight) || 60 }; }
        rh[HR] = { hpt: 22 };
        ws["!rows"] = rh;

        /* Hierarchy grouping */
        if (hasHier) { self._applyHier(ws, HR, lvlFlags, drillFlags, totFlags); }

        /* Cell styles */
        for (var ri = 0; ri < nR; ri++) {
          var isLogoR  = hasLogo && ri === 0;
          var isHdrR   = ri === HR;
          var isTotR   = totFlags[ri] === true;
          var rowLevel = lvlFlags[ri] || 0;
          for (var ci = 0; ci < nC; ci++) {
            var ref = X.utils.encode_cell({ r:ri, c:ci });
            if (!ws[ref]) { ws[ref] = { v:"", t:"s" }; }
            if      (isLogoR) { ws[ref].s = self._sLogo(); }
            else if (isHdrR)  { ws[ref].s = self._sHdr(); }
            else if (isTotR)  { ws[ref].s = self._sTot(ci, vDim.length, rowLevel); }
            else              { ws[ref].s = self._sData(ri, ci, vDim.length, rowLevel, hasHier); }
          }
        }

        /* Number formats */
        var mStart = vDim.length;
        for (var dr2 = HR + 1; dr2 < nR; dr2++) {
          vMeas.forEach(function (m, mi) {
            if (!m.format) { return; }
            var r2 = X.utils.encode_cell({ r:dr2, c:mStart + mi });
            if (ws[r2]) { ws[r2].z = self._fmtNum(m.format); }
          });
        }

        /* Logo text placeholder */
        if (hasLogo) {
          var lRef = X.utils.encode_cell({ r:0, c:0 });
          if (!ws[lRef]) { ws[lRef] = {}; }
          ws[lRef].v = "[Logo: " + self._props.logoUrl + "]";
          ws[lRef].t = "s";
          ws[lRef].s = self._sLogo();
        }

        /* Build workbook and download */
        var wb = X.utils.book_new();
        X.utils.book_append_sheet(wb, ws, self._sheetName());
        X.writeFile(wb, self._fileName(), { compression:true });

        self.dispatchEvent(new CustomEvent("onExportSuccess", { bubbles:true, detail:{} }));
      });
  };

  /* Hierarchy helpers */
  ExcelWidget.prototype._detectHier = function (rows) {
    for (var i = 0; i < rows.length; i++) {
      var row = rows[i];
      if (!Array.isArray(row)) { continue; }
      for (var j = 0; j < row.length; j++) {
        var c = row[j];
        if (!c || typeof c !== "object") { continue; }
        if (typeof c.level === "number" && c.level > 0) { return true; }
        if (typeof c.hierarchyLevel === "number" && c.hierarchyLevel > 0) { return true; }
        if (c.drillState === "expanded" || c.drillState === "collapsed") { return true; }
      }
    }
    return false;
  };

  ExcelWidget.prototype._rowLevel = function (row) {
    for (var j = 0; j < row.length; j++) {
      var c = row[j];
      if (!c || typeof c !== "object") { continue; }
      if (typeof c.level === "number") { return c.level; }
      if (typeof c.hierarchyLevel === "number") { return c.hierarchyLevel; }
    }
    return 0;
  };

  ExcelWidget.prototype._rowDrill = function (row) {
    for (var j = 0; j < row.length; j++) {
      var c = row[j];
      if (!c || typeof c !== "object") { continue; }
      if (c.drillState) { return c.drillState; }
    }
    return "leaf";
  };

  ExcelWidget.prototype._applyHier = function (ws, HR, lvl, drill, tot) {
    if (!ws["!rows"]) { ws["!rows"] = []; }
    var stk = [], ds = HR + 1;
    for (var i = ds; i < lvl.length; i++) {
      var lv = lvl[i] || 0;
      var dr = drill[i] || "leaf";
      var it = tot[i] === true;
      while (ws["!rows"].length <= i) { ws["!rows"].push(null); }
      if (!ws["!rows"][i]) { ws["!rows"][i] = {}; }
      while (stk.length > 0 && stk[stk.length - 1].l >= lv) { stk.pop(); }
      var hidden = stk.some(function (e) { return e.c && e.l < lv; });
      if (lv > 0 && !it) {
        ws["!rows"][i].level  = Math.min(lv, MAX_OUTLINE);
        ws["!rows"][i].hidden = hidden;
      }
      if (dr === "expanded" || dr === "collapsed") {
        stk.push({ l:lv, c:dr === "collapsed" });
      }
    }
    ws["!outline"] = { above:true };
  };

  /* Style factories */
  ExcelWidget.prototype._sLogo = function () {
    return { font:{ bold:true, sz:12, color:{ rgb:"444444" } },
             fill:{ patternType:"solid", fgColor:{ rgb:"FFFFFF" } },
             alignment:{ horizontal:"left", vertical:"center" } };
  };
  ExcelWidget.prototype._sHdr = function () {
    var bg = this._normHex(this._props.headerBgColor, C.HDR_BG);
    return { font:{ bold:true, sz:11, color:{ rgb:C.HDR_FG } },
             fill:{ patternType:"solid", fgColor:{ rgb:bg } },
             alignment:{ horizontal:"center", vertical:"center", wrapText:false },
             border:FB };
  };
  ExcelWidget.prototype._sTot = function (ci, mStart, lv) {
    var bg = lv > 0 ? "E8E8E8" : C.TOT_BG;
    return { font:{ bold:true, sz:11, color:{ rgb:C.TOT_FG } },
             fill:{ patternType:"solid", fgColor:{ rgb:bg } },
             alignment:{ horizontal:ci >= mStart ? "right" : "left", vertical:"center",
                         indent: ci === 0 ? Math.min(lv * 2, 14) : 0 },
             border:FB };
  };
  ExcelWidget.prototype._sData = function (ri, ci, mStart, lv, hasH) {
    var bg = hasH && lv > 0 ? (ri % 2 === 0 ? C.HIER : C.ODD) : (ri % 2 === 0 ? C.EVEN : C.ODD);
    return { font:{ sz:11 },
             fill:{ patternType:"solid", fgColor:{ rgb:bg } },
             alignment:{ horizontal:ci >= mStart ? "right" : "left", vertical:"center",
                         indent: hasH && ci === 0 && lv > 0 ? Math.min(lv * 2, 14) : 0 },
             border:FB };
  };

  /* Utility helpers */
  ExcelWidget.prototype._isTotal = function (row, fMeas) {
    if (!Array.isArray(row)) { return false; }
    return row.some(function (c) {
      if (!c || typeof c !== "object") { return false; }
      return c.type === "ResultCell" || c.type === "TOTAL" ||
             c.type === "SUBTOTAL"   || c.type === "GRAND_TOTAL" ||
             c.isGrandTotal === true || c.isSubTotal === true || c.isTotal === true;
    });
  };
  ExcelWidget.prototype._fmtNum = function (f) {
    if (!f) { return "@"; }
    var map = { "0":"0","0.0":"0.0","0.00":"0.00","#,##0":"#,##0",
                "#,##0.0":"#,##0.0","#,##0.00":"#,##0.00",
                "0%":"0%","0.0%":"0.0%","0.00%":"0.00%" };
    if (map[f]) { return map[f]; }
    if (f.indexOf("$")  >= 0) { return '"$"#,##0.00'; }
    if (f.indexOf("\u20ac") >= 0) { return '[$\u20ac-407]#,##0.00'; }
    if (f.indexOf("\u00a3") >= 0) { return '[$\u00a3-809]#,##0.00'; }
    return f;
  };
  ExcelWidget.prototype._sheetName = function () {
    var sub  = (this._props.tableSubtitle || "Export").trim().replace(/[\\\/\[\]\*\?:]/g, "_");
    var ts   = this._tsShort(new Date());
    var user = (this._props.username || "").trim().replace(/[\\\/\[\]\*\?:]/g, "_");
    return [sub, ts, user].filter(Boolean).join(" ").substring(0, 31);
  };
  ExcelWidget.prototype._fileName = function () {
    var p = (this._props.fileNamePrefix || "SAC_Export").trim().replace(/[^a-zA-Z0-9_\-]/g, "_");
    return p + "_" + this._tsLong(new Date()) + ".xlsx";
  };
  ExcelWidget.prototype._tsShort = function (d) {
    return d.getFullYear() + "-" + this._p2(d.getMonth()+1) + "-" + this._p2(d.getDate()) +
           " " + this._p2(d.getHours()) + ":" + this._p2(d.getMinutes());
  };
  ExcelWidget.prototype._tsLong = function (d) {
    return d.getFullYear() + "-" + this._p2(d.getMonth()+1) + "-" + this._p2(d.getDate()) +
           "_" + this._p2(d.getHours()) + this._p2(d.getMinutes());
  };
  ExcelWidget.prototype._p2 = function (n) { return String(n).padStart(2, "0"); };
  ExcelWidget.prototype._normHex = function (h, fb) {
    if (!h) { return fb; }
    var c = h.replace(/^#/, "").toUpperCase();
    return /^[0-9A-F]{6}$/.test(c) ? c : fb;
  };

  customElements.define("com-custom-sac-excel-export-widget", ExcelWidget);
}());
