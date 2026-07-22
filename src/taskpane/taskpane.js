/* global Office, Excel */

(function () {
  "use strict";

  var VIEW_POLL_MS = 250;
  var MAP_REFRESH_MS = 12000;
  /** Cap drawn grid cells; above this we still keep full extent but downsample drawing. */
  var MAX_DRAW_CELLS = 8000;

  var state = {
    /** Map extent from A1: { row:0, col:0, rowCount, colCount, width, height, xStops, yStops, ... } */
    map: null,
    /** What's on screen in Excel. */
    visible: null,
    zoom: 100,
    colors: null,
    mapRect: null,
    dragging: false,
    building: false
  };

  var ui = {};
  var drag = null;
  /** Bumps on sheet switch so a slow rebuild for the old sheet can't overwrite the new one. */
  var mapGeneration = 0;

  Office.onReady(function (info) {
    if (info.host !== Office.HostType.Excel) {
      return;
    }
    bindUi();
    setStatus("Building map…");
    rebuild()
      .then(function () {
        setStatus("Green box = what’s on screen. Drag to pan · resize to zoom.");
      })
      .catch(function (err) {
        setStatus("Could not build map: " + err.message);
      });

    setInterval(function () {
      if (state.dragging || state.building) { return; }
      syncViewport().catch(function () {});
    }, VIEW_POLL_MS);

    setInterval(function () {
      if (state.dragging || state.building) { return; }
      rebuild().catch(function () {});
    }, MAP_REFRESH_MS);

    registerSheetEvents();
  });

  function registerSheetEvents() {
    Excel.run(function (context) {
      var sheets = context.workbook.worksheets;
      sheets.onActivated.add(function () {
        // Immediate feedback — don't wait for the periodic rebuild.
        mapGeneration += 1;
        var gen = mapGeneration;
        setStatus("Switching sheet…");
        state.map = null;
        state.visible = null;
        state.colors = null;
        state.building = false;
        drawAll();
        return rebuild(gen)
          .then(function () {
            if (gen !== mapGeneration) { return; }
            setStatus("Green box = what’s on screen. Drag to pan · resize to zoom.");
          })
          .catch(function (err) {
            if (gen !== mapGeneration) { return; }
            setStatus("Could not load sheet map: " + err.message);
          });
      });
      return context.sync();
    }).catch(function () {});
  }
  function bindUi() {
    ui.stage = document.getElementById("map-stage");
    ui.canvas = document.getElementById("map-canvas");
    ui.viewport = document.getElementById("viewport");
    ui.empty = document.getElementById("empty-state");
    ui.metaText = document.getElementById("meta-text");
    ui.statusText = document.getElementById("status-text");
    ui.refreshBtn = document.getElementById("refresh-btn");

    ui.refreshBtn.addEventListener("click", function () {
      setStatus("Refreshing…");
      rebuild()
        .then(function () { setStatus("Map updated."); })
        .catch(function (err) { setStatus("Refresh failed: " + err.message); });
    });

    window.addEventListener("resize", onStageResize);

    // Task pane drag-resize doesn't always fire window.resize in Excel's WebView.
    if (typeof ResizeObserver !== "undefined") {
      var ro = new ResizeObserver(function () { onStageResize(); });
      ro.observe(ui.stage);
    }

    ui.viewport.addEventListener("pointerdown", onViewportPointerDown);
    ui.canvas.addEventListener("pointerdown", onCanvasPointerDown);
    window.addEventListener("pointermove", onPointerMove);
    window.addEventListener("pointerup", onPointerUp);
    window.addEventListener("pointercancel", onPointerUp);
  }

  var resizeTimer = null;
  function onStageResize() {
    // Debounce: pane drag fires many events; Excel's visible range updates as the grid shrinks/grows.
    if (resizeTimer) { clearTimeout(resizeTimer); }
    resizeTimer = setTimeout(function () {
      resizeTimer = null;
      drawAll();
      syncViewport().catch(function () {});
    }, 80);
  }

  /**
   * Map = A1 → farthest of (used range, visible window, active cell).
   * Viewport rect = visible window ∩ map (fills the map when the whole map is on screen).
   */
  function rebuild(expectedGen) {
    var gen = expectedGen != null ? expectedGen : mapGeneration;
    state.building = true;
    ui.refreshBtn.disabled = true;

    return Excel.run(function (context) {
      var sheet = context.workbook.worksheets.getActiveWorksheet();
      sheet.load("name");

      var used = sheet.getUsedRangeOrNullObject(false);
      used.load(["rowIndex", "columnIndex", "rowCount", "columnCount"]);

      var active = context.workbook.getActiveCell();
      active.load(["rowIndex", "columnIndex"]);

      var win = context.application.activeWindow;
      win.load(["zoom", "scrollRow", "scrollColumn"]);

      return context.sync().then(function () {
        if (gen !== mapGeneration) { return; }

        var visible = win.visibleRange;
        visible.load(["rowIndex", "columnIndex", "rowCount", "columnCount", "address"]);
        return context.sync().then(function () {
          if (gen !== mapGeneration) { return; }

          state.zoom = win.zoom;
          state.visible = {
            row: visible.rowIndex,
            col: visible.columnIndex,
            rowCount: Math.max(1, visible.rowCount),
            colCount: Math.max(1, visible.columnCount),
            address: visible.address
          };

          var endRow = 0;
          var endCol = 0;

          if (!used.isNullObject && used.rowCount > 0 && used.columnCount > 0) {
            endRow = Math.max(endRow, used.rowIndex + used.rowCount - 1);
            endCol = Math.max(endCol, used.columnIndex + used.columnCount - 1);
          }

          // Include what's currently on screen (so map matches A1:visible BR when that's the working area).
          endRow = Math.max(endRow, visible.rowIndex + visible.rowCount - 1);
          endCol = Math.max(endCol, visible.columnIndex + visible.columnCount - 1);

          endRow = Math.max(endRow, active.rowIndex);
          endCol = Math.max(endCol, active.columnIndex);

          var rowCount = endRow + 1;
          var colCount = endCol + 1;
          var mapRange = sheet.getRangeByIndexes(0, 0, rowCount, colCount);
          mapRange.load(["address", "width", "height", "left", "top"]);

          // Sample values for “has content” marks (fast: one sync).
          var valuesRange = mapRange;
          if (rowCount * colCount > MAX_DRAW_CELLS) {
            // Still load address/size; values sampled later sparsely.
            valuesRange = null;
          } else {
            mapRange.load("values");
          }

          return context.sync().then(function () {
            if (gen !== mapGeneration) { return; }

            var width = Math.max(mapRange.width, 1);
            var height = Math.max(mapRange.height, 1);
            var colWidths = uniformSizes(colCount, width);
            var rowHeights = uniformSizes(rowCount, height);
            var xStops = cumulative(colWidths);
            var yStops = cumulative(rowHeights);

            state.map = {
              row: 0,
              col: 0,
              rowCount: rowCount,
              colCount: colCount,
              address: mapRange.address,
              sheetName: sheet.name,
              left: mapRange.left,
              top: mapRange.top,
              width: xStops[xStops.length - 1],
              height: yStops[yStops.length - 1],
              colWidths: colWidths,
              rowHeights: rowHeights,
              xStops: xStops,
              yStops: yStops
            };

            state.colors = buildColorsFromValues(
              valuesRange ? mapRange.values : null,
              rowCount,
              colCount
            );

            drawAll();
            updateMeta();
          });
        });
      });
    }).catch(function (err) {
      if (gen !== mapGeneration) { return; }
      // If activeWindow/visibleRange fails, still build from used + active only.
      return rebuildWithoutVisible().then(function () {
        if (gen !== mapGeneration) { return; }
        setStatus("Map built (viewport limited): " + err.message);
      });
    }).finally(function () {
      if (gen === mapGeneration) {
        state.building = false;
        ui.refreshBtn.disabled = false;
      }
    });
  }

  function rebuildWithoutVisible() {
    return Excel.run(function (context) {
      var sheet = context.workbook.worksheets.getActiveWorksheet();
      sheet.load("name");
      var used = sheet.getUsedRangeOrNullObject(false);
      used.load(["rowIndex", "columnIndex", "rowCount", "columnCount"]);
      var active = context.workbook.getActiveCell();
      active.load(["rowIndex", "columnIndex"]);
      return context.sync().then(function () {
        var endRow = Math.max(0, active.rowIndex);
        var endCol = Math.max(0, active.columnIndex);
        if (!used.isNullObject && used.rowCount > 0) {
          endRow = Math.max(endRow, used.rowIndex + used.rowCount - 1);
          endCol = Math.max(endCol, used.columnIndex + used.columnCount - 1);
        }
        var rowCount = endRow + 1;
        var colCount = endCol + 1;
        var mapRange = sheet.getRangeByIndexes(0, 0, rowCount, colCount);
        mapRange.load(["address", "width", "height", "left", "top", "values"]);
        return context.sync().then(function () {
          var colWidths = uniformSizes(colCount, Math.max(mapRange.width, 1));
          var rowHeights = uniformSizes(rowCount, Math.max(mapRange.height, 1));
          var xStops = cumulative(colWidths);
          var yStops = cumulative(rowHeights);
          state.map = {
            row: 0,
            col: 0,
            rowCount: rowCount,
            colCount: colCount,
            address: mapRange.address,
            sheetName: sheet.name,
            left: mapRange.left,
            top: mapRange.top,
            width: xStops[xStops.length - 1],
            height: yStops[yStops.length - 1],
            colWidths: colWidths,
            rowHeights: rowHeights,
            xStops: xStops,
            yStops: yStops
          };
          state.colors = buildColorsFromValues(mapRange.values, rowCount, colCount);
          // Assume view covers whole map until we can sync.
          state.visible = {
            row: 0,
            col: 0,
            rowCount: rowCount,
            colCount: colCount,
            address: mapRange.address
          };
          drawAll();
          updateMeta();
        });
      });
    });
  }

  function syncViewport() {
    if (!state.map) { return Promise.resolve(); }

    return Excel.run(function (context) {
      var win = context.application.activeWindow;
      win.load(["zoom"]);
      return context.sync().then(function () {
        var visible = win.visibleRange;
        visible.load(["rowIndex", "columnIndex", "rowCount", "columnCount", "address"]);
        return context.sync().then(function () {
          state.zoom = win.zoom;
          state.visible = {
            row: visible.rowIndex,
            col: visible.columnIndex,
            rowCount: Math.max(1, visible.rowCount),
            colCount: Math.max(1, visible.columnCount),
            address: visible.address
          };

          // If the window now extends past the map, grow the map on next rebuild;
          // for now just draw the viewport clipped to the map.
          drawAll();
          updateMeta();
        });
      });
    });
  }

  function scrollToCell(row, col) {
    return Excel.run(function (context) {
      var win = context.application.activeWindow;
      win.scrollRow = Math.max(1, row + 1);
      win.scrollColumn = Math.max(1, col + 1);
      return context.sync();
    });
  }

  /**
   * Apply a viewport box on the map to Excel: scroll to its top-left and zoom
   * so the on-screen cell span roughly matches the box size.
   */
  function applyViewportBoxToExcel(box) {
    if (!state.mapRect || !state.map || !state.visible) {
      return Promise.resolve();
    }
    var mr = state.mapRect;
    var map = state.map;
    var vis = state.visible;
    var zoom = state.zoom || 100;

    var origin = mapPointToCell(box.x, box.y);
    if (!origin) { return Promise.resolve(); }

    var desiredCols = Math.max(1, (box.w / mr.w) * map.colCount);
    var desiredRows = Math.max(1, (box.h / mr.h) * map.rowCount);
    var scaleCols = vis.colCount / desiredCols;
    var scaleRows = vis.rowCount / desiredRows;
    // Uniform zoom — average the two axes so free resize still feels right.
    var scale = (scaleCols + scaleRows) / 2;
    var newZoom = Math.round(clamp(zoom * scale, 10, 400));

    return Excel.run(function (context) {
      var win = context.application.activeWindow;
      win.zoom = newZoom;
      win.scrollRow = Math.max(1, origin.row + 1);
      win.scrollColumn = Math.max(1, origin.col + 1);
      return context.sync();
    }).then(function () {
      state.zoom = newZoom;
      return syncViewport();
    });
  }

  // ---------------------------------------------------------------------------
  // Draw map + viewport
  // ---------------------------------------------------------------------------

  function drawAll() {
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
    ctx.fillStyle = "#E1DFDD";
    ctx.fillRect(0, 0, cssW, cssH);

    if (!state.map) {
      ui.empty.classList.remove("hidden");
      ui.viewport.hidden = true;
      state.mapRect = null;
      return;
    }
    ui.empty.classList.add("hidden");

    var map = state.map;
    // Fill the stage with the map's physical aspect (letterbox only if needed).
    var fit = fitRect(map.width, map.height, cssW - 4, cssH - 4);
    var ox = 2 + fit.x;
    var oy = 2 + fit.y;
    state.mapRect = { x: ox, y: oy, w: fit.w, h: fit.h };

    var sx = fit.w / map.width;
    var sy = fit.h / map.height;

    // White sheet background
    ctx.fillStyle = "#FFFFFF";
    ctx.fillRect(ox, oy, fit.w, fit.h);

    // Cell content marks + optional fills
    var stepR = 1;
    var stepC = 1;
    var total = map.rowCount * map.colCount;
    if (total > MAX_DRAW_CELLS) {
      var scale = Math.sqrt(total / MAX_DRAW_CELLS);
      stepR = Math.max(1, Math.ceil(scale));
      stepC = Math.max(1, Math.ceil(scale));
    }

    for (var r = 0; r < map.rowCount; r += stepR) {
      var y0 = oy + map.yStops[r] * sy;
      var y1 = oy + map.yStops[Math.min(r + stepR, map.rowCount)] * sy;
      for (var c = 0; c < map.colCount; c += stepC) {
        var x0 = ox + map.xStops[c] * sx;
        var x1 = ox + map.xStops[Math.min(c + stepC, map.colCount)] * sx;
        var color = state.colors && state.colors[r] ? state.colors[r][c] : "#FFFFFF";
        if (color && color !== "#FFFFFF") {
          ctx.fillStyle = color;
          ctx.fillRect(x0, y0, Math.max(x1 - x0, 0.5), Math.max(y1 - y0, 0.5));
        }
      }
    }

    // Grid lines
    ctx.strokeStyle = "#C8C6C4";
    ctx.lineWidth = 1;
    ctx.beginPath();
    var gridStepC = map.colCount > 80 ? Math.ceil(map.colCount / 80) : 1;
    var gridStepR = map.rowCount > 80 ? Math.ceil(map.rowCount / 80) : 1;
    for (var c2 = 0; c2 <= map.colCount; c2 += gridStepC) {
      var gx = Math.round(ox + map.xStops[c2] * sx) + 0.5;
      ctx.moveTo(gx, oy);
      ctx.lineTo(gx, oy + fit.h);
    }
    // Always draw right edge
    var gxR = Math.round(ox + map.xStops[map.colCount] * sx) + 0.5;
    ctx.moveTo(gxR, oy);
    ctx.lineTo(gxR, oy + fit.h);
    for (var r2 = 0; r2 <= map.rowCount; r2 += gridStepR) {
      var gy = Math.round(oy + map.yStops[r2] * sy) + 0.5;
      ctx.moveTo(ox, gy);
      ctx.lineTo(ox + fit.w, gy);
    }
    var gyB = Math.round(oy + map.yStops[map.rowCount] * sy) + 0.5;
    ctx.moveTo(ox, gyB);
    ctx.lineTo(ox + fit.w, gyB);
    ctx.stroke();

    ctx.strokeStyle = "#605E5C";
    ctx.lineWidth = 1;
    ctx.strokeRect(ox + 0.5, oy + 0.5, fit.w - 1, fit.h - 1);

    // Viewport = visible ∩ map (fills entire map when view covers the whole map).
    var box = computeViewportRect();
    if (box) {
      ctx.fillStyle = "rgba(33, 115, 70, 0.22)";
      ctx.strokeStyle = "#217346";
      ctx.lineWidth = 3;
      ctx.fillRect(box.x, box.y, box.w, box.h);
      ctx.strokeRect(box.x + 1.5, box.y + 1.5, Math.max(box.w - 3, 1), Math.max(box.h - 3, 1));

      ui.viewport.hidden = false;
      ui.viewport.removeAttribute("hidden");
      ui.viewport.style.display = "block";
      ui.viewport.style.left = box.x + "px";
      ui.viewport.style.top = box.y + "px";
      ui.viewport.style.width = box.w + "px";
      ui.viewport.style.height = box.h + "px";
    } else {
      ui.viewport.hidden = true;
    }
  }

  /**
   * Visible window clipped to the map.
   * If the window covers the entire map, this returns the full map rect.
   */
  function computeViewportRect() {
    if (!state.mapRect || !state.map || !state.visible) { return null; }

    var map = state.map;
    var vis = state.visible;
    var mr = state.mapRect;

    // Intersection in cell indexes (map is always A1-based: row/col start at 0).
    var c0 = clamp(vis.col, 0, map.colCount);
    var r0 = clamp(vis.row, 0, map.rowCount);
    var c1 = clamp(vis.col + vis.colCount, 0, map.colCount);
    var r1 = clamp(vis.row + vis.rowCount, 0, map.rowCount);

    if (c1 <= c0 || r1 <= r0) {
      // Visible is completely outside map — shouldn't happen often; show a thin marker.
      c0 = clamp(vis.col, 0, map.colCount - 1);
      r0 = clamp(vis.row, 0, map.rowCount - 1);
      c1 = Math.min(map.colCount, c0 + 1);
      r1 = Math.min(map.rowCount, r0 + 1);
    }

    var x0 = map.xStops[c0];
    var y0 = map.yStops[r0];
    var x1 = map.xStops[c1];
    var y1 = map.yStops[r1];

    return {
      x: mr.x + (x0 / map.width) * mr.w,
      y: mr.y + (y0 / map.height) * mr.h,
      w: Math.max(12, ((x1 - x0) / map.width) * mr.w),
      h: Math.max(12, ((y1 - y0) / map.height) * mr.h)
    };
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

  // ---------------------------------------------------------------------------
  // Pointer — move viewport (scroll) or resize (zoom)
  // ---------------------------------------------------------------------------

  function onViewportPointerDown(ev) {
    if (!state.mapRect || !state.map || !state.visible) { return; }
    ev.preventDefault();
    ev.stopPropagation();
    ui.viewport.setPointerCapture(ev.pointerId);

    var handleEl = ev.target.closest ? ev.target.closest(".vp-handle") : null;
    var handle = handleEl ? handleEl.getAttribute("data-handle") : null;
    var stageRect = ui.stage.getBoundingClientRect();
    var left = parseFloat(ui.viewport.style.left) || 0;
    var top = parseFloat(ui.viewport.style.top) || 0;
    var boxW = parseFloat(ui.viewport.style.width) || ui.viewport.offsetWidth;
    var boxH = parseFloat(ui.viewport.style.height) || ui.viewport.offsetHeight;

    drag = {
      mode: handle ? "resize" : "move",
      handle: handle,
      startX: ev.clientX,
      startY: ev.clientY,
      origLeft: left,
      origTop: top,
      origW: boxW,
      origH: boxH,
      stageLeft: stageRect.left,
      stageTop: stageRect.top,
      offsetX: ev.clientX - (stageRect.left + left),
      offsetY: ev.clientY - (stageRect.top + top),
      aspect: boxW / Math.max(boxH, 1)
    };
    state.dragging = true;
  }

  function onCanvasPointerDown(ev) {
    if (!state.mapRect || !state.map) { return; }
    ev.preventDefault();
    var stageRect = ui.stage.getBoundingClientRect();
    var cell = mapPointToCell(ev.clientX - stageRect.left, ev.clientY - stageRect.top);
    if (!cell) { return; }
    var vc = state.visible ? state.visible.colCount : 10;
    var vr = state.visible ? state.visible.rowCount : 20;
    var targetCol = clamp(cell.col - Math.floor(vc / 2), 0, state.map.colCount - 1);
    var targetRow = clamp(cell.row - Math.floor(vr / 2), 0, state.map.rowCount - 1);
    state.dragging = true;
    scrollToCell(targetRow, targetCol)
      .then(function () { return syncViewport(); })
      .finally(function () { state.dragging = false; });
  }

  function onPointerMove(ev) {
    if (!drag || !state.mapRect || !state.map) { return; }
    var mr = state.mapRect;

    if (drag.mode === "move") {
      var left = clamp(ev.clientX - drag.stageLeft - drag.offsetX, mr.x, mr.x + mr.w - drag.origW);
      var top = clamp(ev.clientY - drag.stageTop - drag.offsetY, mr.y, mr.y + mr.h - drag.origH);
      ui.viewport.style.left = left + "px";
      ui.viewport.style.top = top + "px";

      var cell = mapPointToCell(left, top);
      if (!cell) { return; }

      if (!drag.pending) {
        drag.pending = true;
        drag.targetRow = cell.row;
        drag.targetCol = cell.col;
        var sendRow = cell.row;
        var sendCol = cell.col;
        scrollToCell(sendRow, sendCol).finally(function () {
          if (!drag) { return; }
          drag.pending = false;
          if (drag.targetRow !== sendRow || drag.targetCol !== sendCol) {
            scrollToCell(drag.targetRow, drag.targetCol).catch(function () {});
          }
        });
      } else {
        drag.targetRow = cell.row;
        drag.targetCol = cell.col;
      }
      return;
    }

    // Resize — keep Excel-like aspect when dragging corners; edges free one axis.
    var dx = ev.clientX - drag.startX;
    var dy = ev.clientY - drag.startY;
    var l = drag.origLeft;
    var t = drag.origTop;
    var w = drag.origW;
    var h = drag.origH;
    var minW = 24;
    var minH = 24;
    var handle = drag.handle;

    if (handle.indexOf("e") >= 0) { w = drag.origW + dx; }
    if (handle.indexOf("w") >= 0) { w = drag.origW - dx; l = drag.origLeft + dx; }
    if (handle.indexOf("s") >= 0) { h = drag.origH + dy; }
    if (handle.indexOf("n") >= 0) { h = drag.origH - dy; t = drag.origTop + dy; }

    // Corner resize: lock aspect to the current Excel view proportions.
    if (handle.length === 2) {
      if (Math.abs(dx) * drag.aspect > Math.abs(dy)) {
        h = w / drag.aspect;
        if (handle.indexOf("n") >= 0) { t = drag.origTop + drag.origH - h; }
      } else {
        w = h * drag.aspect;
        if (handle.indexOf("w") >= 0) { l = drag.origLeft + drag.origW - w; }
      }
    }

    if (w < minW) {
      if (handle.indexOf("w") >= 0) { l -= (minW - w); }
      w = minW;
    }
    if (h < minH) {
      if (handle.indexOf("n") >= 0) { t -= (minH - h); }
      h = minH;
    }

    // Clamp inside map.
    if (l < mr.x) { w -= (mr.x - l); l = mr.x; }
    if (t < mr.y) { h -= (mr.y - t); t = mr.y; }
    if (l + w > mr.x + mr.w) { w = mr.x + mr.w - l; }
    if (t + h > mr.y + mr.h) { h = mr.y + mr.h - t; }
    w = Math.max(minW, w);
    h = Math.max(minH, h);

    drag.box = { x: l, y: t, w: w, h: h };
    ui.viewport.style.left = l + "px";
    ui.viewport.style.top = t + "px";
    ui.viewport.style.width = w + "px";
    ui.viewport.style.height = h + "px";
  }

  function onPointerUp() {
    if (!drag) { return; }
    var finished = drag;
    drag = null;
    state.dragging = false;

    if (finished.mode === "resize" && finished.box) {
      setStatus("Zooming Excel to match viewport…");
      applyViewportBoxToExcel(finished.box)
        .then(function () {
          setStatus("Green box = what’s on screen. Drag to pan · resize to zoom.");
        })
        .catch(function (err) {
          setStatus("Could not apply zoom: " + err.message);
          syncViewport().catch(function () {});
        });
      return;
    }

    syncViewport().catch(function () {});
  }

  function mapPointToCell(localX, localY) {
    var mr = state.mapRect;
    var map = state.map;
    if (!mr || !map) { return null; }
    var px = ((localX - mr.x) / mr.w) * map.width;
    var py = ((localY - mr.y) / mr.h) * map.height;
    if (px < 0 || py < 0 || px > map.width || py > map.height) { return null; }
    return {
      row: indexFromStops(map.yStops, py),
      col: indexFromStops(map.xStops, px)
    };
  }

  function indexFromStops(stops, pos) {
    var lo = 0;
    var hi = stops.length - 2;
    while (lo < hi) {
      var mid = (lo + hi + 1) >> 1;
      if (stops[mid] <= pos) { lo = mid; } else { hi = mid - 1; }
    }
    return clamp(lo, 0, stops.length - 2);
  }

  // ---------------------------------------------------------------------------
  // Helpers
  // ---------------------------------------------------------------------------

  function buildColorsFromValues(values, rowCount, colCount) {
    var colors = new Array(rowCount);
    for (var r = 0; r < rowCount; r++) {
      colors[r] = new Array(colCount);
      for (var c = 0; c < colCount; c++) {
        var v = values && values[r] ? values[r][c] : null;
        var has = v !== null && v !== undefined && v !== "";
        colors[r][c] = has ? "#D0E7D8" : "#FFFFFF";
      }
    }
    return colors;
  }

  function uniformSizes(count, total) {
    var each = total / Math.max(count, 1);
    var arr = new Array(count);
    for (var i = 0; i < count; i++) { arr[i] = each; }
    return arr;
  }

  function cumulative(sizes) {
    var stops = [0];
    var sum = 0;
    for (var i = 0; i < sizes.length; i++) {
      sum += sizes[i];
      stops.push(sum);
    }
    return stops;
  }

  function updateMeta() {
    if (!state.map) {
      ui.metaText.textContent = "Empty sheet";
      return;
    }
    var addr = state.map.address.replace(/^[^!]*!/, "");
    var end = addr.indexOf(":") >= 0 ? addr.split(":")[1] : addr;
    var parts = [
      state.map.sheetName,
      "map A1:" + end,
      state.map.rowCount + "×" + state.map.colCount
    ];
    if (state.zoom) { parts.push(state.zoom + "%"); }
    if (state.visible && state.visible.address) {
      parts.push("view " + state.visible.address.replace(/^[^!]*!/, ""));
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
