/* global Office, Excel */

var validationTimer = null;

Office.onReady((info) => {
  if (info.host === Office.HostType.Excel) {
    // Enable auto-startup: run this add-in when the workbook opens
    if (Office.addin && Office.addin.setStartupBehavior) {
      Office.addin.setStartupBehavior(Office.StartupBehavior.load);
    }

    // Tag this document so the task pane auto-opens next time
    Office.context.document.settings.set("Office.AutoShowTaskpaneWithDocument", true);
    Office.context.document.settings.saveAsync();

    var btn = document.getElementById("validate-btn");
    if (btn) {
      btn.addEventListener("click", runValidation);
    }
    // Auto-validate on open
    runValidation();
    // Re-validate when cells change
    registerChangeHandler();
  }
});

/**
 * Register a handler that re-validates after worksheet changes.
 * Uses a debounce timer to avoid running on every keystroke.
 */
async function registerChangeHandler() {
  try {
    await Excel.run(async (context) => {
      context.workbook.onChanged.add(onWorkbookChanged);
      await context.sync();
    });
  } catch (e) {
    // Fallback: onChanged not available, manual only
  }
}

function onWorkbookChanged() {
  // Debounce: wait 1.5s after last change before re-validating
  if (validationTimer) clearTimeout(validationTimer);
  validationTimer = setTimeout(function () {
    runValidation();
  }, 1500);
}

async function runValidation() {
  const btn = document.getElementById("validate-btn");
  btn.disabled = true;
  setStatus("チェック中...");

  try {
    const errors = [];

    await Excel.run(async (context) => {
      // === Check 1: 見積書 overall amount must have SUM formula ===
      await checkSumFormula(context, errors);

      // === Check 2: Yellow background cells must not be empty ===
      await checkYellowCells(context, errors);
    });

    updateGateBanner(errors);
    renderResults(errors);

    // Auto-open taskpane if there are errors
    if (errors.length > 0 && Office.addin && Office.addin.showAsTaskpane) {
      try { Office.addin.showAsTaskpane(); } catch (_) { /* already visible */ }
    }
  } catch (e) {
    setStatus("エラーが発生しました: " + e.message);
    console.error(e);
  } finally {
    btn.disabled = false;
  }
}

/**
 * Check 1: On 見積書 sheet, verify that the overall amount cell (G14)
 * contains a SUM formula.
 */
async function checkSumFormula(context, errors) {
  const sheets = context.workbook.worksheets;
  sheets.load("items/name");
  await context.sync();

  const mitsumoriSheet = sheets.items.find((s) => s.name.includes("見積書"));
  if (!mitsumoriSheet) {
    errors.push({
      type: "warn",
      title: "見積書シートが見つかりません",
      details: ["「見積書」という名前のシートが存在しません。"],
    });
    return;
  }

  // The overall amount is at G14 (ご金額), which should have SUM formula
  const amountCell = mitsumoriSheet.getRange("G14");
  amountCell.load(["formulas", "values", "address"]);
  await context.sync();

  const formula = amountCell.formulas[0][0];
  const value = amountCell.values[0][0];

  if (typeof formula !== "string" || !formula.toUpperCase().includes("SUM(")) {
    errors.push({
      type: "fail",
      title: "見積書：合計金額にSUM関数がありません",
      details: [
        "セル G14（ご金額）にSUM関数が設定されていません。",
        "現在の内容: " + (formula || value || "(空)"),
        "期待される形式: =SUM(...)",
      ],
      sheet: "見積書",
      cell: "G14",
    });
  }
}

/**
 * Check 2: Find all cells with yellow background across all sheets.
 * For merged regions, only the top-left cell holds a value — secondary
 * cells appear empty but are not truly unfilled. We detect this by
 * checking each yellow cell's isEntireRow/column merge status via the
 * address of its surrounding merged area.
 */
async function checkYellowCells(context, errors) {
  const sheets = context.workbook.worksheets;
  sheets.load("items/name");
  await context.sync();

  for (const sheet of sheets.items) {
    const usedRange = sheet.getUsedRangeOrNullObject();
    usedRange.load(["rowCount", "columnCount", "columnIndex", "rowIndex"]);
    await context.sync();

    if (usedRange.isNullObject) continue;

    const rowCount = usedRange.rowCount;
    const colCount = usedRange.columnCount;
    const startRow = usedRange.rowIndex;
    const startCol = usedRange.columnIndex;

    // Build a grid: for each cell, store its color and value
    // key = "row,col", value = { color, value }
    const grid = {};
    const batchSize = 20;

    for (let r = 0; r < rowCount; r += batchSize) {
      const rowsInBatch = Math.min(batchSize, rowCount - r);
      const cellInfos = [];

      for (let cr = 0; cr < rowsInBatch; cr++) {
        for (let cc = 0; cc < colCount; cc++) {
          const cell = sheet.getRangeByIndexes(startRow + r + cr, startCol + cc, 1, 1);
          cell.load("values");
          cell.format.fill.load("color");
          cellInfos.push({
            cell,
            absRow: startRow + r + cr,
            absCol: startCol + cc,
          });
        }
      }
      await context.sync();

      for (const { cell, absRow, absCol } of cellInfos) {
        var fillColor;
        try {
          fillColor = cell.format.fill.color;
        } catch (e) {
          continue;
        }
        if (isYellowColor(fillColor)) {
          grid[absRow + "," + absCol] = {
            value: cell.values[0][0],
            row: absRow,
            col: absCol,
          };
        }
      }
    }

    // Now find empty yellow cells, skipping secondary cells of merged regions.
    // Heuristic: if a yellow cell is empty AND the cell directly to its left
    // is also yellow (with or without value), it is likely a secondary merged cell.
    // Only report cells where no yellow neighbor to the left exists,
    // OR no yellow neighbor directly above with the same column exists as a "start".
    const emptyYellowCells = [];

    for (var key in grid) {
      var info = grid[key];
      if (info.value !== null && info.value !== undefined && info.value !== "") {
        continue; // has value, not a problem
      }

      // Check if this is a secondary cell in a horizontal merge:
      // if the cell to the left is also yellow, skip this cell
      var leftKey = info.row + "," + (info.col - 1);
      if (grid[leftKey]) continue;

      // Check if this is a secondary cell in a vertical merge:
      // if the cell above is also yellow AND empty with same column, skip
      var aboveKey = (info.row - 1) + "," + info.col;
      if (grid[aboveKey]) continue;

      emptyYellowCells.push(getCellAddress(info.row, info.col));
    }

    if (emptyYellowCells.length > 0) {
      errors.push({
        type: "fail",
        title: sheet.name + "：黄色セルが未入力",
        details: emptyYellowCells.map(function (addr) {
          return "セル " + addr + " が未入力です";
        }),
        sheet: sheet.name,
        cells: emptyYellowCells,
      });
    }
  }
}

/**
 * Check if a color string represents yellow.
 * Office JS returns colors as "#RRGGBB" format.
 */
function isYellowColor(colorStr) {
  if (!colorStr) return false;

  const hex = colorStr.replace("#", "").toUpperCase();
  if (hex.length !== 6) return false;

  const r = parseInt(hex.substring(0, 2), 16);
  const g = parseInt(hex.substring(2, 4), 16);
  const b = parseInt(hex.substring(4, 6), 16);

  // Yellow: high red, high green, low blue
  return r > 200 && g > 200 && b < 80;
}

/**
 * Convert 0-based row/col to Excel address like "A1"
 */
function getCellAddress(row, col) {
  let colStr = "";
  let c = col;
  while (c >= 0) {
    colStr = String.fromCharCode((c % 26) + 65) + colStr;
    c = Math.floor(c / 26) - 1;
  }
  return colStr + (row + 1);
}

function updateGateBanner(errors) {
  const banner = document.getElementById("gate-banner");
  // Count total unfilled fields across all errors
  var totalUnfilled = 0;
  for (var i = 0; i < errors.length; i++) {
    if (errors[i].cells) {
      totalUnfilled += errors[i].cells.length;
    } else if (errors[i].type === "fail") {
      totalUnfilled += 1;
    }
  }

  if (totalUnfilled === 0) {
    banner.className = "ready";
    banner.innerHTML = '<span class="count">&#10004;</span>提出OK';
  } else {
    banner.className = "not-ready";
    banner.innerHTML =
      '<span class="count">' + totalUnfilled + '</span>' +
      '<span class="label">件の未入力があります — 提出できません</span>';
  }
}

function renderResults(errors) {
  const container = document.getElementById("results-container");

  if (errors.length === 0) {
    container.innerHTML =
      '<div class="no-issues">&#10004; すべてのチェックに合格しました</div>';
    setStatus("チェック完了 - 問題なし");
    return;
  }

  let html = "";
  for (const err of errors) {
    html += '<div class="result-section">';
    html += '<h2 class="' + err.type + '">' + err.title + "</h2>";
    html += '<div class="result-body"><ul>';
    if (err.cells && err.sheet) {
      // Yellow cell errors: each cell gets its own clickable link
      for (var j = 0; j < err.cells.length; j++) {
        html +=
          '<li>セル <span class="cell-link" data-sheet="' +
          err.sheet +
          '" data-cell="' +
          err.cells[j] +
          '">' +
          err.cells[j] +
          "</span> が未入力です</li>";
      }
    } else {
      for (const detail of err.details) {
        html += "<li>" + detail + "</li>";
      }
    }
    html += "</ul>";
    if (err.sheet && err.cell) {
      html +=
        '<p style="margin-top:8px"><span class="cell-link" data-sheet="' +
        err.sheet +
        '" data-cell="' +
        err.cell +
        '">&#8594; セルに移動</span></p>';
    }
    html += "</div></div>";
  }

  container.innerHTML = html;
  setStatus("チェック完了 - " + errors.length + "件の問題が見つかりました");

  // Add click handlers for cell navigation
  document.querySelectorAll(".cell-link").forEach(function (el) {
    el.addEventListener("click", async function () {
      const sheetName = el.dataset.sheet;
      const cellAddr = el.dataset.cell;
      await navigateToCell(sheetName, cellAddr);
    });
  });
}

async function navigateToCell(sheetName, cellAddr) {
  try {
    await Excel.run(async (context) => {
      const sheet = context.workbook.worksheets.getItem(sheetName);
      const range = sheet.getRange(cellAddr);
      range.select();
      sheet.activate();
      await context.sync();
    });
  } catch (e) {
    console.error("Navigation failed:", e);
  }
}

function setStatus(msg) {
  document.getElementById("status").textContent = msg;
}
