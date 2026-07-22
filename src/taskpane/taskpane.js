/* global Office, Excel */

(function () {
  "use strict";

  /** Max pixels on the long edge of the sampled map bitmap. */
  var MAX_SAMPLE = 160;
  /** How often to sync the viewport rect from Excel (ms). */
  var VIEW_POLL_MS = 250;
  /** Rebuild the map bitmap this often (ms), unless Refresh is clicked. */
  var MAP_REFRESH_MS = 8000;

  var state = {
    /** Used range in 0-based indexes: { row, col, rowCount, colCount, address, sheetName } */
    used: null,
    /** Visible range relative to sheet: { row, col, rowCount, colCount } */
    visible: null,
    zoom: 100,
    /** Offscreen sample: { cols, rows, colors: string[][] } */
    sample: null,
    /** Layout of map inside the stage: { x, y, w, h } in CSS pixels */
    mapRect: null,
    dragging: false,
    unsupported: false,
    busy: false
  };

  var ui = {};
  var viewTimer = null;
  var mapTimer = null;
  var drag = null;

  Office.onReady(function (info) {
    if (info.host !== Office.HostType.Excel) {
      return;
    }
    bindUi();

    if (!Office.context.requirements.isSetSupported("ExcelApiDesktop", "1.1")) {
      state.unsupported = true;
      setStatus("Excel Map needs Excel desktop (Windows/Mac) for scroll/zoom APIs.");
      ui.metaText.textContent = "Desktop Excel required";
      return;
    }

    setStatus("Building map…");
    refreshMap()
      .then(function () {
        return syncViewport();
      })
      .then(function () {
        setStatus("Drag the box to scroll. Scroll or zoom in Excel to move it.");
      })
      .catch(function (err) {
        setStatus("Could not build map: " + err.message);
      });

    viewTimer = setInterval(function () {
      if (state.dragging || state.busy) { return; }
      syncViewport().catch(function () { /* ignore transient */ });
    }, VIEW_POLL_MS);

    mapTimer = setInterval(function () {
      if (state.dragging || state.busy) { return; }
      refreshMap().catch(function () { /* ignore transient */ });
    }, MAP_REFRESH_MS);

    registerSheetEvents();
  });

  function bindUi() {
    ui.stage = document.getElementById("map-stage");
    ui.canvas = document.getElementById("map-canvas");
    ui.viewport = document.getElementById("viewport");
    ui.empty = document.getElementById("empty-state");
    ui.metaText = document.getElementById("meta-text");
    ui.statusText = document.getElementById("status-text");
    ui.refreshBtn = document.getElementById("refresh-btn");

    ui.refreshBtn.addEventListener("click", function () {
      setStatus("Refreshing map…");
      refreshMap()
        .then(function () { return syncViewport(); })
        .then(function () { setStatus("Map updated."); })
        .catch(function (err) { setStatus("Refresh failed: " + err.message); });
    });

    window.addEventListener("resize", function () {
      drawMap();
      layoutViewport();
    });

    ui.viewport.addEventListener("pointerdown", onViewportPointerDown);
    ui.canvas.addEventListener("pointerdown", onCanvasPointerDown);
    window.addEventListener("pointermove", onPointerMove);
    window.addEventListener("pointerup", onPointerUp);
    window.addEventListener("pointercancel", onPointerUp);
  }

  function registerSheetEvents() {
    Excel.run(function (context) {
      var sheets = context.workbook.worksheets;
      sheets.onActivated.add(function () {
        return refreshMap().then(function () { return syncViewport(); });
      });
      sheets.onChanged.add(function () {
        // Debounced via the next map timer / explicit refresh; light touch here.
        return Promise.resolve();
      });
      return context.sync();
    }).catch(function () { /* optional */ });
  }

  // ---------------------------------------------------------------------------
  // Excel: used range → color sample
  // ---------------------------------------------------------------------------

  function refreshMap() {
    if (state.unsupported) { return Promise.resolve(); }
    state.busy = true;
    ui.refreshBtn.disabled = true;

    return Excel.run(function (context) {
      var sheet = context.workbook.worksheets.getActiveWorksheet();
      sheet.load("name");
      var used = sheet.getUsedRangeOrNullObject(false);
      used.load(["address", "rowIndex", "columnIndex", "rowCount", "columnCount"]);
      return context.sync().then(function () {
        if (used.isNullObject || used.rowCount < 1 || used.columnCount < 1) {
          state.used = null;
          state.sample = null;
          return null;
        }
        state.used = {
          row: used.rowIndex,
          col: used.columnIndex,
          rowCount: used.rowCount,
          colCount: used.columnCount,
          address: used.address,
          sheetName: sheet.name
        };
        return sampleUsedRange(context, sheet, state.used);
      });
    }).then(function () {
      drawMap();
      updateMeta();
    }).finally(function () {
      state.busy = false;
      ui.refreshBtn.disabled = false;
    });
  }

  function sampleUsedRange(context, sheet, used) {
    var dims = sampleDims(used.colCount, used.rowCount);
    var colors = [];
    var r = 0;

    function nextRow() {
      if (r >= dims.rows) {
        state.sample = { cols: dims.cols, rows: dims.rows, colors: colors };
        return Promise.resolve();
      }

      // One Excel row-band per sample row: pick a representative source row.
      var srcRow = used.row + Math.min(
        used.rowCount - 1,
        Math.floor((r + 0.5) * used.rowCount / dims.rows)
      );
      var props = [];
      var colIndexes = [];
      for (var c = 0; c < dims.cols; c++) {
        var srcCol = used.col + Math.min(
          used.colCount - 1,
          Math.floor((c + 0.5) * used.colCount / dims.cols)
        );
        colIndexes.push(srcCol);
        props.push(sheet.getCell(srcRow, srcCol).format.fill);
      }
      props.forEach(function (fill) {
        fill.load(["color", "tintAndShade"]);
      });

      // Also detect non-empty cells for a subtle “ink” mark when fill is default.
      var valuesRange = sheet.getRangeByIndexes(srcRow, used.col, 1, used.colCount);
      valuesRange.load("values");

      return context.sync().then(function () {
        var rowColors = [];
        var values = valuesRange.values[0] || [];
        for (var c = 0; c < dims.cols; c++) {
          var fill = props[c];
          var color = normalizeFill(fill.color);
          var srcCol = colIndexes[c];
          var localCol = srcCol - used.col;
          var hasValue = localCol >= 0 && localCol < values.length &&
            values[localCol] !== null && values[localCol] !== "";
          if ((!color || color === "#FFFFFF" || color === "#FFF") && hasValue) {
            color = "#5B6570";
          } else if (!color) {
            color = "#F3F3F3";
          }
          rowColors.push(color);
        }
        colors.push(rowColors);
        r += 1;
        if (r % 8 === 0) {
          setStatus("Scanning sheet… " + r + " / " + dims.rows);
        }
        return nextRow();
      });
    }

    return nextRow();
  }

  function sampleDims(colCount, rowCount) {
    var aspect = colCount / Math.max(rowCount, 1);
    var cols;
    var rows;
    if (aspect >= 1) {
      cols = Math.min(MAX_SAMPLE, Math.max(1, colCount));
      rows = Math.max(1, Math.round(cols / aspect));
    } else {
      rows = Math.min(MAX_SAMPLE, Math.max(1, rowCount));
      cols = Math.max(1, Math.round(rows * aspect));
    }
    cols = Math.min(cols, colCount);
    rows = Math.min(rows, rowCount);
    return { cols: cols, rows: rows };
  }

  function normalizeFill(color) {
    if (!color || color === "null" || color === "") { return null; }
    var c = String(color).trim().toUpperCase();
    if (c.charAt(0) !== "#") { c = "#" + c; }
    return c;
  }

  // ---------------------------------------------------------------------------
  // Excel: viewport sync + scroll
  // ---------------------------------------------------------------------------

  function syncViewport() {
    if (state.unsupported || !state.used) {
      layoutViewport();
      return Promise.resolve();
    }

    return Excel.run(function (context) {
      var win = context.workbook.getActiveWindow();
      win.load(["zoom", "scrollRow", "scrollColumn"]);
      var visible = win.visibleRange;
      visible.load(["rowIndex", "columnIndex", "rowCount", "columnCount"]);
      return context.sync().then(function () {
        state.zoom = win.zoom;
        state.visible = {
          row: visible.rowIndex,
          col: visible.columnIndex,
          rowCount: Math.max(1, visible.rowCount),
          colCount: Math.max(1, visible.columnCount)
        };
        layoutViewport();
        updateMeta();
      });
    });
  }

  function scrollToCell(row, col) {
    return Excel.run(function (context) {
      var win = context.workbook.getActiveWindow();
      win.scrollRow = Math.max(1, row + 1); // 1-based in the API
      win.scrollColumn = Math.max(1, col + 1);
      return context.sync();
    });
  }

  // ---------------------------------------------------------------------------
  // Drawing
  // ---------------------------------------------------------------------------

  function drawMap() {
    var canvas = ui.canvas;
    var stage = ui.stage;
    var dpr = window.devicePixelRatio || 1;
    var cssW = stage.clientWidth;
    var cssH = stage.clientHeight;
    if (cssW < 2 || cssH < 2) { return; }

    canvas.width = Math.floor(cssW * dpr);
    canvas.height = Math.floor(cssH * dpr);
    canvas.style.width = cssW + "px";
    canvas.style.height = cssH + "px";

    var ctx = canvas.getContext("2d");
    ctx.setTransform(dpr, 0, 0, dpr, 0, 0);
    ctx.fillStyle = "#2c3138";
    ctx.fillRect(0, 0, cssW, cssH);

    if (!state.sample || !state.used) {
      ui.empty.classList.remove("hidden");
      ui.viewport.hidden = true;
      state.mapRect = null;
      return;
    }
    ui.empty.classList.add("hidden");

    var sample = state.sample;
    var fit = fitRect(sample.cols, sample.rows, cssW - 8, cssH - 8);
    var ox = 4 + fit.x;
    var oy = 4 + fit.y;
    state.mapRect = { x: ox, y: oy, w: fit.w, h: fit.h };

    var cellW = fit.w / sample.cols;
    var cellH = fit.h / sample.rows;
    for (var r = 0; r < sample.rows; r++) {
      for (var c = 0; c < sample.cols; c++) {
        ctx.fillStyle = sample.colors[r][c];
        ctx.fillRect(
          ox + c * cellW,
          oy + r * cellH,
          Math.ceil(cellW) + 0.5,
          Math.ceil(cellH) + 0.5
        );
      }
    }

    // Outer frame
    ctx.strokeStyle = "rgba(255,255,255,0.12)";
    ctx.strokeRect(ox + 0.5, oy + 0.5, fit.w - 1, fit.h - 1);
  }

  function fitRect(contentW, contentH, boxW, boxH) {
    var scale = Math.min(boxW / contentW, boxH / contentH);
    var w = Math.max(1, contentW * scale);
    var h = Math.max(1, contentH * scale);
    return {
      x: (boxW - w) / 2,
      y: (boxH - h) / 2,
      w: w,
      h: h
    };
  }

  function layoutViewport() {
    if (!state.mapRect || !state.used || !state.visible) {
      ui.viewport.hidden = true;
      return;
    }

    var used = state.used;
    var vis = state.visible;
    var map = state.mapRect;

    // Intersection of visible range with used range, in used-local coords.
    var left = clamp(vis.col - used.col, 0, used.colCount);
    var top = clamp(vis.row - used.row, 0, used.rowCount);
    var right = clamp(vis.col + vis.colCount - used.col, 0, used.colCount);
    var bottom = clamp(vis.row + vis.rowCount - used.row, 0, used.rowCount);

    if (right <= left || bottom <= top) {
      // Viewport completely outside used range — still show a thin indicator at the edge.
      left = clamp(vis.col - used.col, 0, used.colCount - 1);
      top = clamp(vis.row - used.row, 0, used.rowCount - 1);
      right = Math.min(used.colCount, left + 1);
      bottom = Math.min(used.rowCount, top + 1);
    }

    var x = map.x + (left / used.colCount) * map.w;
    var y = map.y + (top / used.rowCount) * map.h;
    var w = Math.max(8, ((right - left) / used.colCount) * map.w);
    var h = Math.max(8, ((bottom - top) / used.rowCount) * map.h);

    // Cap to map bounds
    if (x + w > map.x + map.w) { w = map.x + map.w - x; }
    if (y + h > map.y + map.h) { h = map.y + map.h - y; }

    ui.viewport.hidden = false;
    ui.viewport.style.left = x + "px";
    ui.viewport.style.top = y + "px";
    ui.viewport.style.width = w + "px";
    ui.viewport.style.height = h + "px";
  }

  // ---------------------------------------------------------------------------
  // Pointer / drag
  // ---------------------------------------------------------------------------

  function onViewportPointerDown(ev) {
    if (!state.mapRect || !state.used || !state.visible) { return; }
    ev.preventDefault();
    ev.stopPropagation();
    ui.viewport.setPointerCapture(ev.pointerId);
    var rect = ui.viewport.getBoundingClientRect();
    var stageRect = ui.stage.getBoundingClientRect();
    drag = {
      mode: "move",
      pointerId: ev.pointerId,
      offsetX: ev.clientX - rect.left,
      offsetY: ev.clientY - rect.top,
      stageLeft: stageRect.left,
      stageTop: stageRect.top,
      boxW: rect.width,
      boxH: rect.height
    };
    state.dragging = true;
  }

  function onCanvasPointerDown(ev) {
    if (!state.mapRect || !state.used) { return; }
    ev.preventDefault();
    var stageRect = ui.stage.getBoundingClientRect();
    var localX = ev.clientX - stageRect.left;
    var localY = ev.clientY - stageRect.top;
    var center = mapPointToCell(localX, localY);
    if (!center) { return; }

    // Center the current viewport size on the click (or 20×10 if unknown).
    var vc = state.visible ? state.visible.colCount : 20;
    var vr = state.visible ? state.visible.rowCount : 10;
    var targetCol = clamp(center.col - Math.floor(vc / 2), state.used.col, state.used.col + state.used.colCount - 1);
    var targetRow = clamp(center.row - Math.floor(vr / 2), state.used.row, state.used.row + state.used.rowCount - 1);

    state.dragging = true;
    scrollToCell(targetRow, targetCol)
      .then(function () { return syncViewport(); })
      .finally(function () { state.dragging = false; });
  }

  function onPointerMove(ev) {
    if (!drag || !state.mapRect || !state.used) { return; }
    if (drag.mode !== "move") { return; }

    var boxW = drag.boxW;
    var boxH = drag.boxH;
    var map = state.mapRect;
    var left = ev.clientX - drag.stageLeft - drag.offsetX;
    var top = ev.clientY - drag.stageTop - drag.offsetY;

    left = clamp(left, map.x, map.x + map.w - boxW);
    top = clamp(top, map.y, map.y + map.h - boxH);

    ui.viewport.style.left = left + "px";
    ui.viewport.style.top = top + "px";

    // Top-left of the viewport box maps to scroll position (VS Code-style).
    var cell = mapPointToCell(left, top);
    if (!cell) { return; }
    var targetCol = cell.col;
    var targetRow = cell.row;

    // Throttle Excel scrolls while dragging.
    if (!drag.pending) {
      drag.pending = true;
      drag.targetRow = targetRow;
      drag.targetCol = targetCol;
      scrollToCell(targetRow, targetCol).finally(function () {
        drag.pending = false;
        if (drag && (drag.targetRow !== targetRow || drag.targetCol !== targetCol)) {
          // A newer position was requested during the sync — apply latest.
          scrollToCell(drag.targetRow, drag.targetCol).catch(function () {});
        }
      });
    } else {
      drag.targetRow = targetRow;
      drag.targetCol = targetCol;
    }
  }

  function onPointerUp() {
    if (!drag) { return; }
    drag = null;
    state.dragging = false;
    syncViewport().catch(function () {});
  }

  function mapPointToCell(localX, localY) {
    var map = state.mapRect;
    var used = state.used;
    if (!map || !used) { return null; }
    var u = (localX - map.x) / map.w;
    var v = (localY - map.y) / map.h;
    if (u < 0 || v < 0 || u > 1 || v > 1) { return null; }
    var col = used.col + Math.min(used.colCount - 1, Math.floor(u * used.colCount));
    var row = used.row + Math.min(used.rowCount - 1, Math.floor(v * used.rowCount));
    return { row: row, col: col };
  }

  // ---------------------------------------------------------------------------
  // Helpers
  // ---------------------------------------------------------------------------

  function updateMeta() {
    if (!state.used) {
      ui.metaText.textContent = "Empty sheet";
      return;
    }
    var parts = [
      state.used.sheetName,
      state.used.address.replace(/^[^!]*!/, ""),
      state.used.rowCount + "×" + state.used.colCount + " cells"
    ];
    if (state.zoom) {
      parts.push(state.zoom + "% zoom");
    }
    ui.metaText.textContent = parts.join(" · ");
  }

  function clamp(n, lo, hi) {
    return Math.max(lo, Math.min(hi, n));
  }

  function setStatus(text) {
    ui.statusText.textContent = text;
  }
})();
