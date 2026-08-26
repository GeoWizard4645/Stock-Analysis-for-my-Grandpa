/* ==========================================================================
   Daily Stock Analysis — front end
   Plain JavaScript. No framework, no bundler, no dependencies.
   Everything is drawn from data/latest.json, which the Python script writes.
   ========================================================================== */

(function () {
  "use strict";

  var DATA_URL = "data/latest.json";
  var XLSX_URL = "data/latest.xlsx";

  var state = {
    data: null,
    rows: [],
    sortKey: null,
    sortDir: -1,
    group: "all",
    query: "",
    previewGroup: "all"
  };

  /* ---------------------------------------------------------------- utils */

  var $ = function (id) { return document.getElementById(id); };

  function el(tag, cls, text) {
    var n = document.createElement(tag);
    if (cls) n.className = cls;
    if (text !== undefined && text !== null) n.textContent = text;
    return n;
  }

  function pct(v, digits) {
    if (v === null || v === undefined || isNaN(v)) return "—";
    var d = digits === undefined ? 2 : digits;
    return (v >= 0 ? "+" : "") + (v * 100).toFixed(d) + "%";
  }

  function num(v, digits) {
    if (v === null || v === undefined || isNaN(v)) return "—";
    return Number(v).toLocaleString("en-US", {
      minimumFractionDigits: digits === undefined ? 2 : digits,
      maximumFractionDigits: digits === undefined ? 2 : digits
    });
  }

  function compact(v) {
    if (v === null || v === undefined || isNaN(v)) return "—";
    var a = Math.abs(v);
    if (a >= 1e12) return (v / 1e12).toFixed(2) + "T";
    if (a >= 1e9) return (v / 1e9).toFixed(2) + "B";
    if (a >= 1e6) return (v / 1e6).toFixed(1) + "M";
    if (a >= 1e3) return (v / 1e3).toFixed(1) + "K";
    return String(Math.round(v));
  }

  function dirClass(v) {
    if (v === null || v === undefined || isNaN(v) || v === 0) return "flat";
    return v > 0 ? "up" : "down";
  }

  /* Red → grey → green, for heat-map cells. `scale` is the value that reaches
     full colour, so the palette stays comparable between columns. */
  function heatColour(v, scale) {
    if (v === null || v === undefined || isNaN(v)) return "transparent";
    var t = Math.max(-1, Math.min(1, v / scale));
    var a = Math.abs(t);
    if (a < 0.04) return "rgba(120,140,175,.10)";
    return t > 0
      ? "rgba(0, 214, 143, " + (0.10 + a * 0.42).toFixed(3) + ")"
      : "rgba(255, 77, 106, " + (0.10 + a * 0.42).toFixed(3) + ")";
  }

  function scoreColour(score) {
    if (score >= 8) return "#00d68f";
    if (score >= 6) return "#5ee7a8";
    if (score >= 4) return "#ffb020";
    if (score >= 2) return "#ff8a5c";
    return "#ff4d6a";
  }

  var VERDICT_STYLE = {
    "Very Strong": ["rgba(0,214,143,.16)", "#00d68f"],
    "Strong": ["rgba(0,214,143,.10)", "#5ee7a8"],
    "Neutral": ["rgba(147,164,196,.12)", "#93a4c4"],
    "Weak": ["rgba(255,176,32,.13)", "#ffb020"],
    "Very Weak": ["rgba(255,77,106,.14)", "#ff4d6a"]
  };

  var GROUP_COLOUR = { sectors: "#22d3ee", indices: "#00d68f", stocks: "#a78bfa" };

  function chip(value) {
    var v = (value || "").toString();
    var cls = v === "Yes" ? "yes" : v === "No" ? "no" : v === "Same" ? "same" : "na";
    var span = el("span", "chip " + cls, v || "—");
    return span;
  }

  function toast(message) {
    var t = $("toast");
    t.textContent = message;
    t.classList.add("show");
    clearTimeout(toast._timer);
    toast._timer = setTimeout(function () { t.classList.remove("show"); }, 3200);
  }

  /* ------------------------------------------------------------- sparkline */

  function sparkline(values, opts) {
    opts = opts || {};
    var w = opts.width || 150;
    var h = opts.height || 38;
    var pad = 2;
    var ns = "http://www.w3.org/2000/svg";
    var svg = document.createElementNS(ns, "svg");
    svg.setAttribute("viewBox", "0 0 " + w + " " + h);
    svg.setAttribute("width", w);
    svg.setAttribute("height", h);
    svg.style.display = "block";
    svg.style.overflow = "visible";

    if (!values || values.length < 2) return svg;

    var lo = Math.min.apply(null, values);
    var hi = Math.max.apply(null, values);
    var span = (hi - lo) || 1;
    var step = (w - pad * 2) / (values.length - 1);

    var pts = values.map(function (v, i) {
      return [pad + i * step, pad + (h - pad * 2) * (1 - (v - lo) / span)];
    });

    var rising = values[values.length - 1] >= values[0];
    var stroke = opts.colour || (rising ? "#00d68f" : "#ff4d6a");
    var id = "g" + Math.random().toString(36).slice(2, 9);

    var d = pts.map(function (p, i) {
      return (i ? "L" : "M") + p[0].toFixed(2) + " " + p[1].toFixed(2);
    }).join(" ");

    if (opts.fill !== false) {
      var grad = document.createElementNS(ns, "linearGradient");
      grad.setAttribute("id", id);
      grad.setAttribute("x1", "0"); grad.setAttribute("y1", "0");
      grad.setAttribute("x2", "0"); grad.setAttribute("y2", "1");
      [[0, 0.32], [1, 0]].forEach(function (s) {
        var stop = document.createElementNS(ns, "stop");
        stop.setAttribute("offset", s[0]);
        stop.setAttribute("stop-color", stroke);
        stop.setAttribute("stop-opacity", s[1]);
        grad.appendChild(stop);
      });
      var defs = document.createElementNS(ns, "defs");
      defs.appendChild(grad);
      svg.appendChild(defs);

      var area = document.createElementNS(ns, "path");
      area.setAttribute("d", d + " L" + pts[pts.length - 1][0].toFixed(2) + " " + h +
                             " L" + pts[0][0].toFixed(2) + " " + h + " Z");
      area.setAttribute("fill", "url(#" + id + ")");
      svg.appendChild(area);
    }

    var line = document.createElementNS(ns, "path");
    line.setAttribute("d", d);
    line.setAttribute("fill", "none");
    line.setAttribute("stroke", stroke);
    line.setAttribute("stroke-width", opts.weight || 1.6);
    line.setAttribute("stroke-linecap", "round");
    line.setAttribute("stroke-linejoin", "round");
    svg.appendChild(line);

    if (opts.dot !== false) {
      var last = pts[pts.length - 1];
      var dot = document.createElementNS(ns, "circle");
      dot.setAttribute("cx", last[0]); dot.setAttribute("cy", last[1]);
      dot.setAttribute("r", 2.2); dot.setAttribute("fill", stroke);
      svg.appendChild(dot);
    }
    return svg;
  }

  /* ------------------------------------------------------------ donut ring */

  function donut(fraction, size) {
    var ns = "http://www.w3.org/2000/svg";
    var s = size || 152;
    var r = s / 2 - 13;
    var c = 2 * Math.PI * r;
    var svg = document.createElementNS(ns, "svg");
    svg.setAttribute("viewBox", "0 0 " + s + " " + s);
    svg.setAttribute("width", s); svg.setAttribute("height", s);

    function ring(colour, dash, width, opacity) {
      var circle = document.createElementNS(ns, "circle");
      circle.setAttribute("cx", s / 2); circle.setAttribute("cy", s / 2);
      circle.setAttribute("r", r);
      circle.setAttribute("fill", "none");
      circle.setAttribute("stroke", colour);
      circle.setAttribute("stroke-width", width);
      circle.setAttribute("stroke-linecap", "round");
      if (dash) circle.setAttribute("stroke-dasharray", dash);
      if (opacity) circle.setAttribute("opacity", opacity);
      circle.setAttribute("transform", "rotate(-90 " + s / 2 + " " + s / 2 + ")");
      return circle;
    }

    svg.appendChild(ring("#ff4d6a", null, 13, 0.55));
    svg.appendChild(ring("#00d68f", (c * fraction).toFixed(2) + " " + c.toFixed(2), 13));

    var value = document.createElementNS(ns, "text");
    value.setAttribute("x", s / 2); value.setAttribute("y", s / 2 - 2);
    value.setAttribute("text-anchor", "middle");
    value.setAttribute("fill", "#e8eefb");
    value.setAttribute("font-family", "JetBrains Mono, monospace");
    value.setAttribute("font-size", "27"); value.setAttribute("font-weight", "700");
    value.textContent = Math.round(fraction * 100) + "%";
    svg.appendChild(value);

    var label = document.createElementNS(ns, "text");
    label.setAttribute("x", s / 2); label.setAttribute("y", s / 2 + 17);
    label.setAttribute("text-anchor", "middle");
    label.setAttribute("fill", "#5d6f92");
    label.setAttribute("font-family", "JetBrains Mono, monospace");
    label.setAttribute("font-size", "9.5"); label.setAttribute("letter-spacing", "1.6");
    label.textContent = "ADVANCING";
    svg.appendChild(label);
    return svg;
  }

  /* ============================================================== rendering */

  function renderStamp() {
    var d = state.data;
    $("stamp").textContent =
      "Close of " + d.datePretty + " · built " + d.generated;
    $("tapeDate").textContent = d.date;
    document.title = "Daily Stock Analysis · " + d.date;
  }

  function renderTape() {
    var host = $("tapeRows");
    host.innerHTML = "";
    state.rows.filter(function (r) { return r.group === "indices"; })
      .forEach(function (r) {
        var row = el("div", "tape-row");
        row.appendChild(el("div", "tape-sym", r.ticker));
        row.appendChild(el("div", "tape-name", r.name));
        row.appendChild(el("div", "tape-px num", num(r.last)));
        var chg = el("div", "tape-chg num " + dirClass(r.pct_change), pct(r.pct_change));
        row.appendChild(chg);
        row.addEventListener("click", function () { openDrawer(r.ticker); });
        row.style.cursor = "pointer";
        host.appendChild(row);
      });
  }

  function renderTiles() {
    var s = state.data.summary;
    var host = $("tiles");
    host.innerHTML = "";
    var tiles = [
      { label: "Advancing", value: s.advancing, sub: "of " + s.count + " names", colour: "#00d68f" },
      { label: "Declining", value: s.declining, sub: "of " + s.count + " names", colour: "#ff4d6a" },
      { label: "Breadth", value: s.breadth.toFixed(0) + "%", sub: "closed higher",
        colour: s.breadth >= 50 ? "#00d68f" : "#ff4d6a" },
      { label: "Average score", value: s.avgScore.toFixed(1), sub: "out of 9", colour: "#22d3ee" },
      { label: "Above 200 DMA", value: s.above200, sub: "in a long-term uptrend", colour: "#60a5fa" },
      { label: "Signals", value: s.signals, sub: "events fired today", colour: "#a78bfa" }
    ];
    tiles.forEach(function (t) {
      var card = el("div", "tile");
      card.style.setProperty("--accent", t.colour);
      card.appendChild(el("div", "tile-label", t.label));
      var v = el("div", "tile-value", t.value);
      v.style.color = t.colour;
      card.appendChild(v);
      card.appendChild(el("div", "tile-sub", t.sub));
      host.appendChild(card);
    });
  }

  /* ------------------------------------------------------------- preview */

  /* key, short header for the table, full wording for the drawer */
  var PREVIEW_COLS = [
    ["pct_change", "% Change", "% change"],
    ["strength", "Rank", "Strength rank in its group"],
    ["c_gt_prev", "Price ><br>Pre-Day", "Price above yesterday's close"],
    ["v_gt_prev", "Volume ><br>Pre-Day", "Volume above yesterday's"],
    ["macd_gt_sig", "MACD<br>Green>Red", "MACD line above its signal"],
    ["c_gt_ema8", "> 8<br>EMA", "Price above the 8 EMA"],
    ["c_gt_ema21", "> 21<br>EMA", "Price above the 21 EMA"],
    ["c_gt_sma50", "> 50<br>DMA", "Price above the 50 DMA"],
    ["c_gt_bbu", "> Upper<br>BB", "Price above the upper band"],
    ["c_gt_bbl", "> Lower<br>BB", "Price above the lower band"],
    ["c_gt_bbm", "> Mid<br>BB", "Price above the middle band"]
  ];

  function renderPreviewFilters() {
    var host = $("previewFilters");
    host.innerHTML = "";
    [["all", "All 40"], ["sectors", "Sectors"], ["indices", "Index"], ["stocks", "Stocks"]]
      .forEach(function (g) {
        var b = el("button", "pill" + (state.previewGroup === g[0] ? " on" : ""), g[1]);
        b.addEventListener("click", function () {
          state.previewGroup = g[0];
          renderPreviewFilters();
          renderPreview();
        });
        host.appendChild(b);
      });
  }

  function renderPreview() {
    var head = $("previewHead");
    var body = $("previewBody");
    head.innerHTML = "";
    body.innerHTML = "";

    var tr = el("tr");
    tr.appendChild(Object.assign(el("th", "left"), { textContent: "Ticker" }));
    PREVIEW_COLS.forEach(function (c) {
      var th = el("th");
      th.innerHTML = c[1];
      th.title = c[2];
      tr.appendChild(th);
    });
    tr.appendChild(el("th", null, "Score"));
    tr.appendChild(el("th", null, "90 days"));
    head.appendChild(tr);

    var shown = state.rows.filter(function (r) {
      return state.previewGroup === "all" || r.group === state.previewGroup;
    });

    var lastGroup = null;
    shown.forEach(function (r) {
      if (r.group !== lastGroup) {
        lastGroup = r.group;
        var gr = el("tr", "group-row");
        var gtd = el("td");
        gtd.colSpan = PREVIEW_COLS.length + 3;
        var tag = el("span", "group-tag", r.groupTitle);
        tag.style.color = GROUP_COLOUR[r.group] || "#93a4c4";
        gtd.appendChild(tag);
        gr.appendChild(gtd);
        body.appendChild(gr);
      }

      var row = el("tr");
      row.style.cursor = "pointer";
      row.addEventListener("click", function () { openDrawer(r.ticker); });

      var symCell = el("td", "left");
      var box = el("div", "sym-cell");
      var dot = el("span", "dot");
      dot.style.background = GROUP_COLOUR[r.group] || "#93a4c4";
      box.appendChild(dot);
      box.appendChild(el("span", "sym", r.ticker));
      symCell.appendChild(box);
      row.appendChild(symCell);

      PREVIEW_COLS.forEach(function (c) {
        var key = c[0];
        var td = el("td");
        if (key === "pct_change") {
          td.className = "num " + dirClass(r.pct_change);
          td.textContent = pct(r.pct_change);
        } else if (key === "strength") {
          td.className = "num";
          td.textContent = r.strength === null || r.strength === undefined ? "—" : r.strength;
          td.style.color = "#93a4c4";
        } else {
          td.appendChild(chip(r[key]));
        }
        row.appendChild(td);
      });

      var scoreTd = el("td");
      scoreTd.appendChild(scoreCell(r.score));
      row.appendChild(scoreTd);

      var sparkTd = el("td");
      sparkTd.appendChild(sparkline(r.spark, { width: 96, height: 26 }));
      row.appendChild(sparkTd);

      body.appendChild(row);
    });
  }

  function scoreCell(score) {
    var wrap = el("div", "score-cell");
    var meter = el("div", "meter");
    var fill = el("i");
    fill.style.width = (score / 9 * 100) + "%";
    fill.style.background = scoreColour(score);
    meter.appendChild(fill);
    wrap.appendChild(meter);
    var n = el("span", "score-num", score);
    n.style.color = scoreColour(score);
    wrap.appendChild(n);
    return wrap;
  }

  /* --------------------------------------------------------------- pulse */

  function renderPulse() {
    var s = state.data.summary;
    var frac = s.breadth / 100;

    var d = $("donut");
    d.innerHTML = "";
    d.appendChild(donut(frac));

    var legend = $("donutLegend");
    legend.innerHTML = "";
    [["#00d68f", "Higher", s.advancing],
     ["#ff4d6a", "Lower", s.declining],
     ["#5d6f92", "Unchanged", s.unchanged]].forEach(function (item) {
      var row = el("div", "legend-item");
      var sw = el("span", "legend-swatch");
      sw.style.background = item[0];
      row.appendChild(sw);
      row.appendChild(el("span", null, item[1]));
      row.appendChild(el("b", null, item[2]));
      legend.appendChild(row);
    });

    $("adRatio").textContent = s.declining
      ? (s.advancing / s.declining).toFixed(2) + " : 1"
      : "all higher";

    var bar = $("adBar");
    bar.innerHTML = "";
    var g = el("i"); g.style.background = "linear-gradient(90deg,#00b876,#00d68f)";
    g.style.width = (s.advancing / s.count * 100) + "%";
    var r = el("i"); r.style.background = "linear-gradient(90deg,#ff4d6a,#e03a56)";
    r.style.width = (s.declining / s.count * 100) + "%";
    var f = el("i"); f.style.background = "#22304d";
    f.style.width = (s.unchanged / s.count * 100) + "%";
    bar.appendChild(g); bar.appendChild(r); bar.appendChild(f);

    var labels = $("adLabels");
    labels.innerHTML = "";
    var left = el("span", "up", s.advancing + " up");
    var right = el("span", "down", s.declining + " down");
    labels.appendChild(left); labels.appendChild(right);

    // trend participation bars
    var host = $("participation");
    host.innerHTML = "";
    var checks = [
      ["Above the 8 EMA", "c_gt_ema8"],
      ["Above the 21 EMA", "c_gt_ema21"],
      ["Above the 50 DMA", "c_gt_sma50"],
      ["Above the mid band", "c_gt_bbm"]
    ];
    checks.forEach(function (c) {
      var n = state.rows.filter(function (row) { return row[c[1]] === "Yes"; }).length;
      var share = n / state.rows.length;
      var line = el("div");
      line.style.cssText = "display:grid;grid-template-columns:130px 1fr 42px;align-items:center;gap:11px;margin-bottom:9px";
      var label = el("span", null, c[0]);
      label.style.cssText = "font-size:12px;color:var(--text-dim)";
      var track = el("div", "meter");
      track.style.width = "100%";
      var fill = el("i");
      fill.style.width = (share * 100) + "%";
      fill.style.background = share >= 0.5 ? "#00d68f" : "#ff4d6a";
      track.appendChild(fill);
      var val = el("span", "num", n + "/" + state.rows.length);
      val.style.cssText = "font-size:11px;color:var(--text-faint);text-align:right";
      line.appendChild(label); line.appendChild(track); line.appendChild(val);
      host.appendChild(line);
    });

    // score histogram
    var hist = $("hist");
    hist.innerHTML = "";
    var buckets = [];
    for (var i = 0; i <= 9; i++) buckets.push(0);
    state.rows.forEach(function (row) { buckets[row.score] += 1; });
    var peak = Math.max.apply(null, buckets) || 1;
    buckets.forEach(function (count, score) {
      var col = el("div", "hist-col");
      col.appendChild(el("div", "hist-count", count || ""));
      var bar = el("div", "hist-bar");
      bar.style.height = Math.max(3, count / peak * 78) + "px";
      bar.style.background = scoreColour(score);
      bar.style.opacity = count ? 1 : 0.22;
      col.appendChild(bar);
      col.appendChild(el("div", "hist-label", score));
      col.title = count + " ticker" + (count === 1 ? "" : "s") + " scored " + score;
      hist.appendChild(col);
    });
  }

  /* ------------------------------------------------------------ rotation */

  var HEAT_COLS = [
    ["pct_change", "Today", 0.02],
    ["r5", "1 Week", 0.05],
    ["r21", "1 Month", 0.10],
    ["r63", "3 Months", 0.18],
    ["r126", "6 Months", 0.30],
    ["ytd", "Year to date", 0.35]
  ];

  function renderRotation() {
    var head = $("heatHead"), body = $("heatBody");
    head.innerHTML = ""; body.innerHTML = "";

    var tr = el("tr");
    tr.appendChild(Object.assign(el("th"), { textContent: "Sector", style: "text-align:left;padding-left:12px" }));
    tr.appendChild(el("th", null, "Fund"));
    HEAT_COLS.forEach(function (c) { tr.appendChild(el("th", null, c[1])); });
    tr.appendChild(el("th", null, "Score"));
    tr.appendChild(el("th", null, "90 days"));
    head.appendChild(tr);

    var sectors = state.rows.filter(function (r) { return r.group === "sectors"; })
      .slice()
      .sort(function (a, b) { return (b.r21 || -9) - (a.r21 || -9); });

    sectors.forEach(function (r) {
      var row = el("tr");
      row.style.cursor = "pointer";
      row.addEventListener("click", function () { openDrawer(r.ticker); });

      var nameTd = el("td");
      nameTd.appendChild(el("div", "name-cell", shortSectorName(r.name)));
      row.appendChild(nameTd);

      var symTd = el("td");
      var sym = el("span", "sym", r.ticker);
      sym.style.color = "#e8eefb";
      symTd.appendChild(sym);
      row.appendChild(symTd);

      HEAT_COLS.forEach(function (c) {
        var td = el("td");
        var cell = el("span", "cell", pct(r[c[0]], 1));
        cell.style.background = heatColour(r[c[0]], c[2]);
        cell.style.color = r[c[0]] > 0 ? "#8bf5cd" : r[c[0]] < 0 ? "#ffa3b4" : "#93a4c4";
        td.appendChild(cell);
        row.appendChild(td);
      });

      var scoreTd = el("td");
      scoreTd.appendChild(scoreCell(r.score));
      row.appendChild(scoreTd);

      var sparkTd = el("td");
      var box = el("div");
      box.style.cssText = "display:flex;justify-content:center;padding:0 8px";
      box.appendChild(sparkline(r.spark, { width: 110, height: 30 }));
      sparkTd.appendChild(box);
      row.appendChild(sparkTd);

      body.appendChild(row);
    });
  }

  function shortSectorName(name) {
    return (name || "")
      .replace(/^State Street\s+/i, "")
      .replace(/\s*Select Sector SPDR ETF$/i, "")
      .replace(/\s*Select Sector SPDR Fund$/i, "")
      .replace(/\s*SPDR ETF$/i, "");
  }

  /* -------------------------------------------------------------- matrix */

  var MATRIX_COLS = [
    { key: "ticker", label: "Ticker", align: "left", type: "sym" },
    { key: "last", label: "Last", type: "num", digits: 2 },
    { key: "pct_change", label: "% Chg", type: "pct" },
    { key: "score", label: "Score", type: "score" },
    { key: "verdict", label: "Signal", type: "verdict" },
    { key: "rsi", label: "RSI", type: "num", digits: 1, heat: "rsi" },
    { key: "stoch", label: "Stoch %K", type: "num", digits: 1 },
    { key: "adx", label: "ADX", type: "num", digits: 1 },
    { key: "macd_hist", label: "MACD Hist", type: "signed", digits: 3 },
    { key: "rvol", label: "Rel Vol", type: "x" },
    { key: "volume", label: "Volume", type: "compact" },
    { key: "pct_b", label: "%B", type: "num", digits: 2 },
    { key: "bb_width", label: "Band W", type: "pctpoint", digits: 1 },
    { key: "atr_pct", label: "ATR", type: "pctpoint", digits: 2 },
    { key: "ema8", label: "8 EMA", type: "num", digits: 2 },
    { key: "ema21", label: "21 EMA", type: "num", digits: 2 },
    { key: "sma50", label: "50 DMA", type: "num", digits: 2 },
    { key: "sma200", label: "200 DMA", type: "num", digits: 2 },
    { key: "off_hi", label: "Off 52W Hi", type: "pctpoint", digits: 1 },
    { key: "range_pos", label: "52W Pos", type: "rangebar" },
    { key: "r5", label: "5 Day", type: "pct" },
    { key: "r21", label: "1 Month", type: "pct" },
    { key: "ytd", label: "YTD", type: "pct" },
    { key: "spark", label: "90 Days", type: "spark", sortable: false }
  ];

  function renderMatrixHead() {
    var head = $("matrixHead");
    head.innerHTML = "";
    var tr = el("tr");
    MATRIX_COLS.forEach(function (c) {
      var th = el("th", (c.align === "left" ? "left " : "") +
                        (c.sortable === false ? "" : "sortable") +
                        (state.sortKey === c.key ? " sorted" : ""));
      th.textContent = c.label;
      if (state.sortKey === c.key) {
        var arrow = el("span", "arrow", state.sortDir < 0 ? "▼" : "▲");
        th.appendChild(arrow);
      }
      if (c.sortable !== false) {
        th.addEventListener("click", function () {
          if (state.sortKey === c.key) state.sortDir *= -1;
          else { state.sortKey = c.key; state.sortDir = -1; }
          renderMatrixHead();
          renderMatrixBody();
        });
      }
      tr.appendChild(th);
    });
    head.appendChild(tr);
  }

  function renderMatrixBody() {
    var body = $("matrixBody");
    body.innerHTML = "";

    var rows = state.rows.filter(function (r) {
      if (state.group !== "all" && r.group !== state.group) return false;
      if (!state.query) return true;
      var q = state.query.toLowerCase();
      return r.ticker.toLowerCase().indexOf(q) >= 0 ||
             (r.name || "").toLowerCase().indexOf(q) >= 0;
    });

    if (state.sortKey) {
      var k = state.sortKey, dir = state.sortDir;
      rows = rows.slice().sort(function (a, b) {
        var x = a[k], y = b[k];
        if (typeof x === "string" || typeof y === "string") {
          return String(x).localeCompare(String(y)) * -dir;
        }
        if (x === null || x === undefined) return 1;
        if (y === null || y === undefined) return -1;
        return (x - y) * dir;
      });
    }

    rows.forEach(function (r) {
      var tr = el("tr");
      tr.style.cursor = "pointer";
      tr.addEventListener("click", function () { openDrawer(r.ticker); });
      MATRIX_COLS.forEach(function (c) { tr.appendChild(matrixCell(r, c)); });
      body.appendChild(tr);
    });

    if (!rows.length) {
      var tr2 = el("tr");
      var td = el("td", null, "Nothing matches that filter.");
      td.colSpan = MATRIX_COLS.length;
      td.style.cssText = "padding:28px;color:var(--text-faint)";
      tr2.appendChild(td);
      body.appendChild(tr2);
    }
  }

  function matrixCell(r, c) {
    var td = el("td", c.align === "left" ? "left" : null);
    var v = r[c.key];

    switch (c.type) {
      case "sym": {
        var box = el("div", "sym-cell");
        var dot = el("span", "dot");
        dot.style.background = GROUP_COLOUR[r.group] || "#93a4c4";
        box.appendChild(dot);
        box.appendChild(el("span", "sym", r.ticker));
        td.appendChild(box);
        break;
      }
      case "pct":
        td.className = "num " + dirClass(v);
        td.textContent = pct(v);
        break;
      case "signed":
        td.className = "num " + dirClass(v);
        td.textContent = v === null || v === undefined ? "—" : num(v, c.digits);
        break;
      case "pctpoint":
        td.className = "num";
        td.textContent = v === null || v === undefined ? "—" : num(v, c.digits) + "%";
        if (c.key === "off_hi") td.style.color = v > -5 ? "#00d68f" : v < -25 ? "#ff4d6a" : "#93a4c4";
        break;
      case "x":
        td.className = "num";
        td.textContent = v === null || v === undefined ? "—" : num(v, 2) + "x";
        td.style.color = v >= 1.75 ? "#22d3ee" : v >= 1 ? "#e8eefb" : "#5d6f92";
        break;
      case "compact":
        td.className = "num";
        td.style.color = "#93a4c4";
        td.textContent = compact(v);
        break;
      case "score":
        td.appendChild(scoreCell(r.score));
        break;
      case "verdict": {
        var style = VERDICT_STYLE[v] || VERDICT_STYLE.Neutral;
        var badge = el("span", "verdict", v);
        badge.style.background = style[0];
        badge.style.color = style[1];
        td.appendChild(badge);
        break;
      }
      case "rangebar": {
        var wrap = el("div", "score-cell");
        var meter = el("div", "meter");
        var fill = el("i");
        fill.style.width = Math.max(0, Math.min(100, v || 0)) + "%";
        fill.style.background = v > 70 ? "#00d68f" : v < 30 ? "#ff4d6a" : "#ffb020";
        meter.appendChild(fill);
        wrap.appendChild(meter);
        var label = el("span", "score-num", v === null || v === undefined ? "—" : Math.round(v));
        label.style.width = "22px";
        label.style.color = "#93a4c4";
        wrap.appendChild(label);
        td.appendChild(wrap);
        break;
      }
      case "spark": {
        var holder = el("div");
        holder.style.cssText = "display:flex;justify-content:center";
        holder.appendChild(sparkline(r.spark, { width: 104, height: 26 }));
        td.appendChild(holder);
        break;
      }
      default:
        td.className = "num";
        td.textContent = num(v, c.digits);
        if (c.heat === "rsi" && v !== null && v !== undefined) {
          td.style.color = v >= 70 ? "#ff4d6a" : v <= 30 ? "#00d68f" : "#e8eefb";
        }
    }
    return td;
  }

  /* ------------------------------------------------------------- signals */

  var BULLISH = ["crossed up", "crossed above", "above upper", "52-week high",
                 "oversold", "reclaimed", "gapped up", "back above 50"];

  function isBullish(text) {
    var t = text.toLowerCase();
    return BULLISH.some(function (k) { return t.indexOf(k) >= 0; });
  }

  function renderSignals() {
    var feed = $("feed");
    feed.innerHTML = "";
    var items = [];
    state.rows.forEach(function (r) {
      (r.events || []).forEach(function (e) { items.push({ row: r, event: e }); });
    });
    items.sort(function (a, b) { return b.row.score - a.row.score; });
    $("signalCount").textContent = items.length + " event" + (items.length === 1 ? "" : "s");

    if (!items.length) {
      var quiet = el("p", null,
        "Nothing crossed, broke out or spiked today. A quiet session.");
      quiet.style.cssText = "color:var(--text-faint);margin:6px 0";
      feed.appendChild(quiet);
      return;
    }

    items.forEach(function (item) {
      var good = isBullish(item.event);
      var node = el("div", "feed-item");
      node.style.setProperty("--accent", good ? "#00d68f" : "#ff4d6a");
      node.addEventListener("click", function () { openDrawer(item.row.ticker); });

      var sym = el("div", "sym", item.row.ticker);
      node.appendChild(sym);

      var text = el("div", "feed-text", item.event);
      node.appendChild(text);

      var meta = el("div", "feed-meta " + dirClass(item.row.pct_change), pct(item.row.pct_change));
      node.appendChild(meta);
      feed.appendChild(node);
    });
  }

  function renderLeaders() {
    var host = $("leaders");
    host.innerHTML = "";
    var ranked = state.rows.slice().sort(function (a, b) {
      return (b.score - a.score) || ((b.pct_change || 0) - (a.pct_change || 0));
    });
    board(host, "Strongest", ranked.slice(0, 6));
    board(host, "Weakest", ranked.slice(-6).reverse());
  }

  function board(host, title, rows) {
    host.appendChild(el("div", "sub-head", title));
    rows.forEach(function (r) {
      var line = el("div");
      line.style.cssText = "display:grid;grid-template-columns:64px 1fr 60px 34px;align-items:center;gap:10px;padding:7px 0;cursor:pointer";
      line.addEventListener("click", function () { openDrawer(r.ticker); });
      line.appendChild(el("span", "sym", r.ticker));
      var spark = el("span");
      spark.appendChild(sparkline(r.spark, { width: 100, height: 22, fill: false }));
      line.appendChild(spark);
      var chg = el("span", "num " + dirClass(r.pct_change), pct(r.pct_change));
      chg.style.cssText += ";font-size:12px;text-align:right";
      line.appendChild(chg);
      var sc = el("span", "num", r.score);
      sc.style.cssText = "font-size:12px;text-align:right;font-weight:700;color:" + scoreColour(r.score);
      line.appendChild(sc);
      host.appendChild(line);
    });
  }

  /* ------------------------------------------------------------ news wall */

  function renderNews() {
    var host = $("newsWall");
    host.innerHTML = "";
    var movers = state.rows.slice()
      .filter(function (r) { return (r.news || []).length; })
      .sort(function (a, b) {
        return Math.abs(b.pct_change || 0) - Math.abs(a.pct_change || 0);
      })
      .slice(0, 6);

    if (!movers.length) {
      host.innerHTML = '<p style="color:var(--text-faint);margin:0">' +
        "No headlines came back this run.</p>";
      return;
    }

    var grid = el("div");
    grid.style.cssText = "display:grid;grid-template-columns:repeat(auto-fit,minmax(320px,1fr));gap:22px";

    movers.forEach(function (r) {
      var col = el("div");
      var head = el("div");
      head.style.cssText = "display:flex;align-items:baseline;gap:10px;margin-bottom:12px;cursor:pointer";
      head.addEventListener("click", function () { openDrawer(r.ticker); });
      var sym = el("span", "sym", r.ticker);
      sym.style.fontSize = "15px";
      head.appendChild(sym);
      var chg = el("span", "num " + dirClass(r.pct_change), pct(r.pct_change));
      chg.style.fontSize = "13px";
      head.appendChild(chg);
      var nm = el("span", null, r.name);
      nm.style.cssText = "font-size:11.5px;color:var(--text-faint);overflow:hidden;text-overflow:ellipsis;white-space:nowrap";
      head.appendChild(nm);
      col.appendChild(head);

      r.news.slice(0, 3).forEach(function (n) {
        col.appendChild(newsCard(n));
      });
      grid.appendChild(col);
    });
    host.appendChild(grid);
  }

  function newsCard(n) {
    var a = el("a", "news-item");
    a.href = n.url;
    a.target = "_blank";
    a.rel = "noopener noreferrer";
    a.appendChild(el("div", "news-title", n.title));
    var meta = el("div", "news-meta");
    meta.appendChild(el("span", null, n.source || "Yahoo Finance"));
    if (n.published) meta.appendChild(el("span", null, tidyDate(n.published)));
    a.appendChild(meta);
    return a;
  }

  function tidyDate(text) {
    var d = new Date(text);
    if (isNaN(d.getTime())) return "";
    return d.toLocaleDateString("en-US", { month: "short", day: "numeric" });
  }

  /* -------------------------------------------------------------- drawer */

  function openDrawer(ticker) {
    var r = state.rows.filter(function (x) { return x.ticker === ticker; })[0];
    if (!r) return;

    var head = $("drawerHead");
    head.innerHTML = "";
    var left = el("div");
    var symRow = el("div");
    symRow.style.cssText = "display:flex;align-items:baseline;gap:12px";
    var sym = el("span", "sym", r.ticker);
    sym.style.cssText += ";font-size:26px";
    symRow.appendChild(sym);
    var badge = el("span", "verdict", r.verdict);
    var vs = VERDICT_STYLE[r.verdict] || VERDICT_STYLE.Neutral;
    badge.style.background = vs[0]; badge.style.color = vs[1];
    symRow.appendChild(badge);
    left.appendChild(symRow);
    var nm = el("div", null, r.name);
    nm.style.cssText = "color:var(--text-dim);font-size:13px;margin-top:5px";
    left.appendChild(nm);
    var px = el("div");
    px.style.cssText = "display:flex;align-items:baseline;gap:12px;margin-top:12px";
    var last = el("span", "num", num(r.last));
    last.style.cssText += ";font-size:30px;font-weight:700";
    px.appendChild(last);
    var chg = el("span", "num " + dirClass(r.pct_change),
                 (r.chg >= 0 ? "+" : "") + num(r.chg) + "  (" + pct(r.pct_change) + ")");
    chg.style.cssText += ";font-size:15px;font-weight:600";
    px.appendChild(chg);
    left.appendChild(px);
    head.appendChild(left);

    var close = el("button", "drawer-close", "✕");
    close.setAttribute("aria-label", "Close");
    close.addEventListener("click", closeDrawer);
    head.appendChild(close);

    var body = $("drawerBody");
    body.innerHTML = "";

    // chart
    var chartBox = el("div");
    chartBox.style.cssText = "background:var(--surface-2);border:1px solid var(--line-soft);border-radius:11px;padding:16px;margin-bottom:6px";
    var chart = sparkline(r.spark, { width: 480, height: 132, weight: 2 });
    chart.style.width = "100%";
    chart.setAttribute("preserveAspectRatio", "none");
    chartBox.appendChild(chart);
    var range = el("div");
    range.style.cssText = "display:flex;justify-content:space-between;font-family:var(--mono);font-size:10.5px;color:var(--text-faint);margin-top:9px";
    range.appendChild(el("span", null, "90 sessions"));
    range.appendChild(el("span", null,
      "low " + num(Math.min.apply(null, r.spark)) + " · high " + num(Math.max.apply(null, r.spark))));
    chartBox.appendChild(range);
    body.appendChild(chartBox);

    // the checklist, exactly as it appears in the workbook
    body.appendChild(el("div", "sub-head", "The daily checklist"));
    var list = el("div");
    list.style.cssText = "display:grid;gap:1px;background:var(--line-soft);border:1px solid var(--line-soft);border-radius:9px;overflow:hidden";
    PREVIEW_COLS.slice(2).forEach(function (c) {
      var line = el("div");
      line.style.cssText = "display:flex;align-items:center;justify-content:space-between;gap:12px;background:var(--surface-2);padding:9px 13px";
      var label = el("span", null, c[2]);
      label.style.cssText = "font-size:12.5px;color:var(--text-dim)";
      line.appendChild(label);
      line.appendChild(chip(r[c[0]]));
      list.appendChild(line);
    });
    body.appendChild(list);

    // score
    body.appendChild(el("div", "sub-head", "Bull score"));
    var scoreBox = el("div");
    scoreBox.style.cssText = "display:flex;align-items:center;gap:16px";
    var bigScore = el("span", "num", r.score + " / 9");
    bigScore.style.cssText += ";font-size:26px;font-weight:700;color:" + scoreColour(r.score);
    scoreBox.appendChild(bigScore);
    var track = el("div", "meter");
    track.style.cssText = "flex:1;height:9px";
    var fill = el("i");
    fill.style.width = (r.score / 9 * 100) + "%";
    fill.style.background = scoreColour(r.score);
    track.appendChild(fill);
    scoreBox.appendChild(track);
    body.appendChild(scoreBox);
    var trendNote = el("p", null, r.trend);
    trendNote.style.cssText = "margin:11px 0 0;color:var(--text-dim);font-size:13px";
    body.appendChild(trendNote);

    // numbers
    body.appendChild(el("div", "sub-head", "The numbers"));
    var kv = el("dl", "kv");
    [["Open", num(r.open)], ["High", num(r.high)], ["Low", num(r.low)],
     ["Prev close", num(r.prev_close)],
     ["Volume", compact(r.volume)], ["Vs 20-day avg", num(r.rvol, 2) + "x"],
     ["RSI (14)", num(r.rsi, 1)], ["Stochastic %K", num(r.stoch, 1)],
     ["ADX (14)", num(r.adx, 1)], ["ATR (14)", num(r.atr_pct, 2) + "%"],
     ["MACD", num(r.macd, 3)], ["Signal", num(r.macd_signal, 3)],
     ["8 EMA", num(r.ema8)], ["21 EMA", num(r.ema21)],
     ["50 DMA", num(r.sma50)], ["200 DMA", num(r.sma200)],
     ["Band width", num(r.bb_width, 1) + "%"], ["%B", num(r.pct_b, 2)],
     ["5-day", pct(r.r5)], ["1-month", pct(r.r21)],
     ["3-month", pct(r.r63)], ["Year to date", pct(r.ytd)]
    ].forEach(function (pair) {
      var cell = el("div");
      cell.appendChild(el("dt", null, pair[0]));
      cell.appendChild(el("dd", null, pair[1]));
      kv.appendChild(cell);
    });
    body.appendChild(kv);

    // bollinger position
    body.appendChild(el("div", "sub-head", "Inside the Bollinger Bands"));
    body.appendChild(rail(r.pct_b === null ? 0.5 : r.pct_b, 0, 1,
                          num(r.bb_low), num(r.bb_up), "Lower band", "Upper band"));

    body.appendChild(el("div", "sub-head", "52-week range"));
    body.appendChild(rail((r.range_pos || 0) / 100, 0, 1,
                          num(r.lo52), num(r.hi52), "52-week low", "52-week high"));

    // events
    if ((r.events || []).length) {
      body.appendChild(el("div", "sub-head", "Signals today"));
      r.events.forEach(function (e) {
        var line = el("div");
        var good = isBullish(e);
        line.style.cssText = "display:flex;align-items:center;gap:10px;padding:9px 12px;margin-bottom:7px;border-radius:8px;background:" +
          (good ? "rgba(0,214,143,.08)" : "rgba(255,77,106,.08)") +
          ";border:1px solid " + (good ? "rgba(0,214,143,.22)" : "rgba(255,77,106,.22)") +
          ";font-size:13px";
        line.appendChild(el("span", null, good ? "▲" : "▼"));
        line.appendChild(el("span", null, e));
        line.firstChild.style.color = good ? "#00d68f" : "#ff4d6a";
        body.appendChild(line);
      });
    }

    // news
    if ((r.news || []).length) {
      body.appendChild(el("div", "sub-head", "Top news"));
      r.news.forEach(function (n) { body.appendChild(newsCard(n)); });
    }

    $("drawer").classList.add("open");
    $("drawer").setAttribute("aria-hidden", "false");
    $("scrim").classList.add("open");
    document.body.style.overflow = "hidden";

    // keep the address bar in step, so a read-out can be linked to directly
    if (history.replaceState) history.replaceState(null, "", "#t=" + r.ticker);
  }

  function rail(fraction, lo, hi, loLabel, hiLabel, loText, hiText) {
    var wrap = el("div", "rail");
    var track = el("div", "rail-track");
    var f = Math.max(0, Math.min(1, fraction));
    var fill = el("div", "rail-fill");
    fill.style.width = (f * 100) + "%";
    fill.style.background = "linear-gradient(90deg,#ff4d6a,#ffb020 50%,#00d68f)";
    track.appendChild(fill);
    var knob = el("div", "rail-knob");
    knob.style.left = (f * 100) + "%";
    knob.style.background = "#e8eefb";
    track.appendChild(knob);
    wrap.appendChild(track);
    var ends = el("div", "rail-ends");
    ends.appendChild(el("span", null, loText + " " + loLabel));
    ends.appendChild(el("span", null, hiText + " " + hiLabel));
    wrap.appendChild(ends);
    return wrap;
  }

  function closeDrawer() {
    $("drawer").classList.remove("open");
    $("drawer").setAttribute("aria-hidden", "true");
    $("scrim").classList.remove("open");
    document.body.style.overflow = "";
    if (history.replaceState && location.hash.indexOf("#t=") === 0) {
      history.replaceState(null, "", location.pathname + location.search);
    }
  }

  /* Open whatever #t=TICKER asks for, on load and on back/forward. */
  function openFromHash() {
    var m = /^#t=([A-Za-z.\-]{1,12})$/.exec(location.hash || "");
    if (m) openDrawer(m[1].toUpperCase());
  }

  /* ------------------------------------------------------------- download */
  /*
     This is the button the whole site exists for, so it is deliberately
     stubborn. It is wired up before the market data is fetched, so it works
     even if everything else on the page fails to load. A failed fetch is
     retried once, and if fetching is blocked outright it falls back to a plain
     link and lets the browser do the downloading itself.
  */

  function fileName() {
    var stamp = (state.data && state.data.date)
      ? state.data.date.slice(5)
      : new Date().toISOString().slice(5, 10);
    return "Stock_Analysis_" + stamp + ".xlsx";
  }

  function saveBlob(blob, name) {
    if (window.navigator && window.navigator.msSaveOrOpenBlob) {   // old Edge
      window.navigator.msSaveOrOpenBlob(blob, name);
      return;
    }
    var url = URL.createObjectURL(blob);
    var a = document.createElement("a");
    a.href = url;
    a.download = name;
    a.rel = "noopener";
    a.style.display = "none";
    document.body.appendChild(a);
    a.click();
    setTimeout(function () {
      document.body.removeChild(a);
      URL.revokeObjectURL(url);
    }, 4000);
  }

  /* Last resort: hand the URL straight to the browser's own downloader. */
  function saveByLink(name) {
    var a = document.createElement("a");
    a.href = XLSX_URL + "?t=" + Date.now();
    a.download = name;
    a.rel = "noopener";
    a.style.display = "none";
    document.body.appendChild(a);
    a.click();
    setTimeout(function () { document.body.removeChild(a); }, 4000);
  }

  function fetchWorkbook(attempt) {
    return fetch(XLSX_URL + "?t=" + Date.now(), { cache: "no-store" })
      .then(function (res) {
        if (!res.ok) throw new Error("HTTP " + res.status);
        return res.blob();
      })
      .then(function (blob) {
        // A valid .xlsx is a zip file: it must start with the bytes "PK".
        if (blob.size < 5000) throw new Error("file looks truncated");
        return blob;
      })
      .catch(function (err) {
        if (attempt < 2) return fetchWorkbook(attempt + 1);
        throw err;
      });
  }

  function wireDownload() {
    var btn = $("generateBtn");
    var label = $("generateLabel");
    var idle = label.textContent;
    var busy = false;

    btn.addEventListener("click", function (ev) {
      // Take over from the plain link so the file can be given today's date in
      // its name. If anything below throws, the fallbacks still deliver it.
      ev.preventDefault();
      if (busy) return;
      busy = true;
      btn.classList.add("is-busy");
      label.textContent = "Building your spreadsheet\u2026";

      var name = fileName();

      fetchWorkbook(0)
        .then(function (blob) {
          saveBlob(blob, name);
          label.textContent = "Saved \u2014 click again for another copy";
          toast("Saved " + name + " to your Downloads folder.");
        })
        .catch(function () {
          // Fetching failed. Let the browser fetch it the ordinary way.
          saveByLink(name);
          label.textContent = idle;
          toast("Downloading " + name + " \u2026");
        })
        .then(function () {
          busy = false;
          btn.classList.remove("is-busy");
          setTimeout(function () { label.textContent = idle; }, 6000);
        });
    });
  }

  function wireControls() {
    $("matrixSearch").addEventListener("input", function (e) {
      state.query = e.target.value.trim();
      renderMatrixBody();
    });

    Array.prototype.forEach.call(
      document.querySelectorAll(".pill[data-group]"), function (btn) {
        btn.addEventListener("click", function () {
          Array.prototype.forEach.call(
            document.querySelectorAll(".pill[data-group]"), function (b) {
              b.classList.remove("on");
            });
          btn.classList.add("on");
          state.group = btn.getAttribute("data-group");
          renderMatrixBody();
        });
      });

    $("scrim").addEventListener("click", closeDrawer);
    window.addEventListener("hashchange", openFromHash);
    document.addEventListener("keydown", function (e) {
      if (e.key === "Escape") closeDrawer();
    });
  }

  /* ----------------------------------------------------------------- boot */

  function renderAll() {
    renderStamp();
    renderTape();
    renderTiles();
    renderPreviewFilters();
    renderPreview();
    renderPulse();
    renderRotation();
    renderMatrixHead();
    renderMatrixBody();
    renderSignals();
    renderLeaders();
    renderNews();
    openFromHash();
  }

  function fail(message) {
    $("stamp").textContent = "Data unavailable";
    var host = $("previewBody");
    var tr = el("tr");
    var td = el("td");
    td.colSpan = 14;
    td.innerHTML = '<div class="error-box"><b>Could not load the market data.</b><br>' +
      message + "<br><br>Run <code>python stock_analysis.py</code> locally to build " +
      "the spreadsheet from your own machine.</div>";
    tr.appendChild(td);
    host.appendChild(tr);
  }

  wireDownload();

  fetch(DATA_URL, { cache: "no-store" })
    .then(function (res) {
      if (!res.ok) throw new Error("HTTP " + res.status);
      return res.json();
    })
    .then(function (data) {
      state.data = data;
      state.rows = data.rows;
      wireControls();
      renderAll();
    })
    .catch(function (err) { fail(err.message); });
})();
