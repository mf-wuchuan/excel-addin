/* global Office, Excel */

Office.onReady(() => {
  Office.actions.associate("validateWorkbook", validateWorkbook);
});

/**
 * Ribbon button handler — runs validation and shows a dialog with results.
 */
async function validateWorkbook(event) {
  try {
    const errors = [];

    await Excel.run(async (context) => {
      await checkSumFormula(context, errors);
      await checkYellowCells(context, errors);
    });

    if (errors.length === 0) {
      // No issues found — no notification needed
    } else {
      const messages = errors.map(function (e) {
        return "【" + e.title + "】\n" + e.details.join("\n");
      });

      // Show dialog with validation results
      await showValidationDialog(messages.join("\n\n"));
    }
  } catch (e) {
    console.error("Validation error:", e);
  }

  event.completed();
}

function showValidationDialog(message) {
  return new Promise((resolve) => {
    // Encode message for the dialog
    const encodedMsg = encodeURIComponent(message);
    const dialogHtml =
      "data:text/html," +
      encodeURIComponent(
        '<!DOCTYPE html><html><head><meta charset="UTF-8"><style>' +
          "body{font-family:'Yu Gothic',sans-serif;padding:20px;}" +
          "h1{font-size:16px;color:#c00;margin-bottom:12px;}" +
          "pre{white-space:pre-wrap;font-size:12px;line-height:1.6;}" +
          "</style></head><body>" +
          "<h1>入力チェック - 問題が見つかりました</h1>" +
          "<pre>" +
          message.replace(/</g, "&lt;") +
          "</pre>" +
          "</body></html>"
      );

    try {
      Office.context.ui.displayDialogAsync(
        dialogHtml,
        { height: 40, width: 30 },
        function (result) {
          resolve();
        }
      );
    } catch (e) {
      // Dialog not supported, fall through
      resolve();
    }
  });
}

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

  const amountCell = mitsumoriSheet.getRange("G14");
  amountCell.load(["formulas", "values"]);
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
      ],
    });
  }
}

async function checkYellowCells(context, errors) {
  const sheets = context.workbook.worksheets;
  sheets.load("items/name");
  await context.sync();

  for (const sheet of sheets.items) {
    const usedRange = sheet.getUsedRangeOrNullObject();
    usedRange.load(["rowCount", "columnCount", "rowIndex", "columnIndex"]);
    await context.sync();

    if (usedRange.isNullObject) continue;

    const rowCount = usedRange.rowCount;
    const colCount = usedRange.columnCount;
    const startRow = usedRange.rowIndex;
    const startCol = usedRange.columnIndex;
    const grid = {};
    const batchSize = 20;

    for (let r = 0; r < rowCount; r += batchSize) {
      const rowsInBatch = Math.min(batchSize, rowCount - r);
      const cellInfos = [];

      for (let cr = 0; cr < rowsInBatch; cr++) {
        for (let cc = 0; cc < colCount; cc++) {
          const cell = sheet.getRangeByIndexes(
            startRow + r + cr, startCol + cc, 1, 1
          );
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
        try { fillColor = cell.format.fill.color; } catch (e) { continue; }
        if (isYellowColor(fillColor)) {
          grid[absRow + "," + absCol] = {
            value: cell.values[0][0],
            row: absRow,
            col: absCol,
          };
        }
      }
    }

    // Find empty yellow cells, skipping secondary cells of merged regions.
    // Heuristic: if the cell to the left or above is also yellow, it is
    // likely a continuation of a merged region — skip it.
    const emptyYellowCells = [];
    for (var key in grid) {
      var info = grid[key];
      if (info.value !== null && info.value !== undefined && info.value !== "") continue;
      if (grid[info.row + "," + (info.col - 1)]) continue;
      if (grid[(info.row - 1) + "," + info.col]) continue;
      emptyYellowCells.push(getCellAddress(info.row, info.col));
    }

    if (emptyYellowCells.length > 0) {
      errors.push({
        type: "fail",
        title: sheet.name + "：黄色セルが未入力",
        details: emptyYellowCells.map(function (addr) {
          return "セル " + addr + " が未入力です";
        }),
      });
    }
  }
}

function isYellowColor(colorStr) {
  if (!colorStr) return false;
  const hex = colorStr.replace("#", "").toUpperCase();
  if (hex.length !== 6) return false;
  const r = parseInt(hex.substring(0, 2), 16);
  const g = parseInt(hex.substring(2, 4), 16);
  const b = parseInt(hex.substring(4, 6), 16);
  return r > 200 && g > 200 && b < 80;
}

function getCellAddress(row, col) {
  let colStr = "";
  let c = col;
  while (c >= 0) {
    colStr = String.fromCharCode((c % 26) + 65) + colStr;
    c = Math.floor(c / 26) - 1;
  }
  return colStr + (row + 1);
}
