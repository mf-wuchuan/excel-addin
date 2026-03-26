const fs = require("fs");
const path = require("path");
const ExcelJS = require("exceljs");

// --- Configuration ---
const TARGET_DIR = process.argv[2] || ".";
const REPORT_DIR = process.argv[3] || path.join(__dirname, "reports");

async function main() {
  const xlsxFiles = findXlsxFiles(TARGET_DIR);
  if (xlsxFiles.length === 0) {
    console.log("No .xlsx files found in: " + TARGET_DIR);
    return;
  }

  console.log("Found " + xlsxFiles.length + " xlsx file(s) in: " + TARGET_DIR);

  const allResults = [];

  for (const filePath of xlsxFiles) {
    console.log("Checking: " + filePath);
    try {
      const errors = await validateWorkbook(filePath);
      allResults.push({ file: filePath, errors });
    } catch (e) {
      allResults.push({
        file: filePath,
        errors: [{ type: "error", title: "Failed to read file", details: [e.message] }],
      });
    }
  }

  writeReport(allResults);
}

// Recursively find all .xlsx files (skip temp files starting with ~$)
function findXlsxFiles(dir) {
  const results = [];
  let entries;
  try {
    entries = fs.readdirSync(dir, { withFileTypes: true });
  } catch (e) {
    console.warn("Cannot read directory: " + dir + " (" + e.message + ")");
    return results;
  }
  for (const entry of entries) {
    const fullPath = path.join(dir, entry.name);
    if (entry.isDirectory()) {
      results.push(...findXlsxFiles(fullPath));
    } else if (
      entry.isFile() &&
      entry.name.endsWith(".xlsx") &&
      !entry.name.startsWith("~$")
    ) {
      results.push(fullPath);
    }
  }
  return results;
}

async function validateWorkbook(filePath) {
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.readFile(filePath);

  const errors = [];

  // Check 1: SUM formula in G14 on sheet containing "見積書"
  checkSumFormula(workbook, errors);

  // Check 2: Yellow cells must not be empty
  checkYellowCells(workbook, errors);

  return errors;
}

function checkSumFormula(workbook, errors) {
  const sheet = workbook.worksheets.find((s) => s.name.includes("見積書"));
  if (!sheet) {
    errors.push({
      type: "warn",
      title: "見積書シートが見つかりません",
      details: ["「見積書」という名前のシートが存在しません。"],
    });
    return;
  }

  const cell = sheet.getCell("G14");
  const formula = cell.formula || "";
  const value = cell.value;

  if (!formula.toUpperCase().includes("SUM(")) {
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

function checkYellowCells(workbook, errors) {
  for (const sheet of workbook.worksheets) {
    const emptyYellowCells = [];
    const yellowCells = new Set(); // track "row,col" of yellow cells

    sheet.eachRow({ includeEmpty: false }, (row, rowNumber) => {
      row.eachCell({ includeEmpty: true }, (cell, colNumber) => {
        const fill = cell.style && cell.style.fill;
        if (!fill) return;

        const color = extractColor(fill);
        if (!isYellowColor(color)) return;

        yellowCells.add(rowNumber + "," + colNumber);
      });
    });

    // Second pass: find empty yellow cells, skip merged secondary cells
    for (const key of yellowCells) {
      const [row, col] = key.split(",").map(Number);
      const cell = sheet.getCell(row, col);
      const value = cell.value;

      if (value !== null && value !== undefined && value !== "") continue;

      // Skip secondary cells in horizontal merge (left neighbor is also yellow)
      if (yellowCells.has(row + "," + (col - 1))) continue;

      // Skip secondary cells in vertical merge (above neighbor is also yellow)
      if (yellowCells.has((row - 1) + "," + col)) continue;

      emptyYellowCells.push(getCellAddress(row, col));
    }

    if (emptyYellowCells.length > 0) {
      errors.push({
        type: "fail",
        title: sheet.name + "：黄色セルが未入力",
        details: emptyYellowCells.map((addr) => "セル " + addr + " が未入力です"),
        sheet: sheet.name,
        cells: emptyYellowCells,
      });
    }
  }
}

function extractColor(fill) {
  if (!fill || fill.type !== "pattern" || !fill.fgColor) return null;

  const fg = fill.fgColor;

  // Theme/indexed colors with tint — hard to resolve without theme XML, skip
  if (fg.argb) return fg.argb;
  if (fg.theme !== undefined) return null;

  return null;
}

function isYellowColor(colorStr) {
  if (!colorStr) return false;

  // ARGB format: "FFRRGGBB" or "RRGGBB"
  const hex = colorStr.replace(/^FF/, "").toUpperCase();
  if (hex.length !== 6) return false;

  const r = parseInt(hex.substring(0, 2), 16);
  const g = parseInt(hex.substring(2, 4), 16);
  const b = parseInt(hex.substring(4, 6), 16);

  return r > 200 && g > 200 && b < 80;
}

function getCellAddress(row, col) {
  let colStr = "";
  let c = col - 1; // exceljs uses 1-based columns
  while (c >= 0) {
    colStr = String.fromCharCode((c % 26) + 65) + colStr;
    c = Math.floor(c / 26) - 1;
  }
  return colStr + row;
}

function writeReport(allResults) {
  if (!fs.existsSync(REPORT_DIR)) {
    fs.mkdirSync(REPORT_DIR, { recursive: true });
  }

  const now = new Date();
  const timestamp = now.getFullYear() +
    String(now.getMonth() + 1).padStart(2, "0") +
    String(now.getDate()).padStart(2, "0") + "_" +
    String(now.getHours()).padStart(2, "0") +
    String(now.getMinutes()).padStart(2, "0") +
    String(now.getSeconds()).padStart(2, "0");

  const reportPath = path.join(REPORT_DIR, "report_" + timestamp + ".txt");

  const lines = [];
  lines.push("=".repeat(60));
  lines.push("見積書バリデーション レポート");
  lines.push("実行日時: " + now.toLocaleString("ja-JP"));
  lines.push("対象フォルダ: " + path.resolve(TARGET_DIR));
  lines.push("=".repeat(60));
  lines.push("");

  let totalFiles = allResults.length;
  let filesWithErrors = 0;
  let totalErrors = 0;

  for (const result of allResults) {
    const relPath = path.relative(TARGET_DIR, result.file) || result.file;
    const hasErrors = result.errors.length > 0;
    if (hasErrors) filesWithErrors++;

    lines.push("-".repeat(60));
    lines.push("ファイル: " + relPath);

    if (!hasErrors) {
      lines.push("  OK - 問題なし");
    } else {
      for (const err of result.errors) {
        totalErrors++;
        const icon = err.type === "fail" ? "NG" : err.type === "warn" ? "WARN" : "ERR";
        lines.push("  [" + icon + "] " + err.title);
        for (const detail of err.details) {
          lines.push("       " + detail);
        }
      }
    }
    lines.push("");
  }

  lines.push("=".repeat(60));
  lines.push("合計: " + totalFiles + "ファイル / " +
    filesWithErrors + "ファイルにエラー / " +
    totalErrors + "件の問題");
  lines.push("=".repeat(60));

  const report = lines.join("\n");
  fs.writeFileSync(reportPath, report, "utf-8");

  console.log("");
  console.log("Report written to: " + reportPath);
  console.log(totalFiles + " files checked, " + filesWithErrors + " with errors, " + totalErrors + " issues total");
}

main().catch((e) => {
  console.error("Fatal error:", e);
  process.exit(1);
});
