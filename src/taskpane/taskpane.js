/* global Office, Excel */

Office.onReady((info) => {
  if (info.host === Office.HostType.Excel) {
    document.getElementById("validate-btn").addEventListener("click", runValidation);
    registerSaveHandler();
  }
});

/**
 * Register a handler that runs validation after the workbook is saved.
 * Uses Workbook.onSaved (ExcelApi 1.12+), with graceful fallback.
 */
async function registerSaveHandler() {
  try {
    await Excel.run(async (context) => {
      context.workbook.onSaved.add(onWorkbookSaved);
      await context.sync();
      setStatus("保存時に自動チェックが有効です");
    });
  } catch (e) {
    setStatus("手動チェックモード（保存時の自動チェックはこのバージョンでは利用できません）");
  }
}

async function onWorkbookSaved() {
  await runValidation();
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

    renderResults(errors);
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
 * Only report the top-left cell of each merged region (secondary merged
 * cells inherit the color but have no independent value).
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

    // First, collect merged areas to know which cells are secondary (non-top-left)
    const mergedCells = new Set();
    try {
      // Load all cells' merge info by checking each cell's mergeArea
      // This is expensive, so we'll do it alongside color checking
    } catch (e) {
      // Merge detection not critical
    }

    const emptyYellowCells = [];
    // Process rows in batches to limit context.sync() calls
    const batchSize = 20;

    for (let r = 0; r < rowCount; r += batchSize) {
      const rowsInBatch = Math.min(batchSize, rowCount - r);
      const cellInfos = [];

      for (let cr = 0; cr < rowsInBatch; cr++) {
        for (let cc = 0; cc < colCount; cc++) {
          const cell = sheet.getRangeByIndexes(startRow + r + cr, startCol + cc, 1, 1);
          cell.load("values");
          cell.format.fill.load("color");
          // Load mergeArea to detect merged cells
          const mergeArea = cell.getMergeAreasOrNullObject();
          mergeArea.load("address");
          cellInfos.push({
            cell,
            mergeArea,
            absRow: startRow + r + cr,
            absCol: startCol + cc,
          });
        }
      }
      await context.sync();

      // Track which merge areas we've already checked (by address)
      const checkedMergeAreas = new Set();

      for (const { cell, mergeArea, absRow, absCol } of cellInfos) {
        const fillColor = cell.format.fill.color;
        if (!isYellowColor(fillColor)) continue;

        // For merged cells, only check the top-left cell of the merge area
        if (mergeArea && !mergeArea.isNullObject) {
          const mergeAddr = mergeArea.address;
          if (checkedMergeAreas.has(mergeAddr)) continue;
          checkedMergeAreas.add(mergeAddr);

          // Extract top-left address from merge address (e.g., "Sheet!A1:C3" -> check A1)
          const topLeftAddr = mergeAddr.split("!").pop().split(":")[0];
          const thisCellAddr = getCellAddress(absRow, absCol);
          if (topLeftAddr !== thisCellAddr) continue;
        }

        const cellValue = cell.values[0][0];
        if (cellValue === null || cellValue === undefined || cellValue === "") {
          emptyYellowCells.push(getCellAddress(absRow, absCol));
        }
      }
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
    for (const detail of err.details) {
      html += "<li>" + detail + "</li>";
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
