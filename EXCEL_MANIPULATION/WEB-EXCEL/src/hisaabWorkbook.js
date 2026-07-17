import * as XLSX from "xlsx";
import hisaabWorkbookUrl from "./assets/Hisaab.xlsx?url";

const PREMADE_SHEET_NAME = "Sheet4";

function cleanNumberish(value) {
  return String(value ?? "")
    .replaceAll(",", "")
    .replaceAll("\u20B9", "")
    .replaceAll("$", "")
    .trim();
}

function toFormulaValue(value) {
  if (typeof value === "number") return Number.isFinite(value) ? value : 0;
  if (typeof value === "boolean") return value ? 1 : 0;
  if (value == null || value === "") return 0;
  const numeric = Number(cleanNumberish(value));
  return Number.isFinite(numeric) ? numeric : value;
}

function normalizeComputedValue(value) {
  if (typeof value === "number") return Number.isFinite(value) ? value : 0;
  if (value == null) return "";
  return value;
}

function formatCellDisplay(cell) {
  if (!cell) return "";
  if (cell.value == null || cell.value === "") return "";
  if (typeof cell.value === "number" && cell.format) {
    try {
      return XLSX.SSF.format(cell.format, cell.value);
    } catch {
      return String(cell.value);
    }
  }
  return String(cell.value);
}

function formatEditableValue(cell) {
  if (!cell) return "";
  if (cell.value == null || cell.value === "") return "";
  return typeof cell.value === "number" ? String(cell.value) : String(cell.value);
}

function parseEditableValue(input) {
  const text = String(input ?? "").trim();
  if (!text) return "";
  const numeric = Number(cleanNumberish(text));
  return Number.isFinite(numeric) && /^[-+]?[\$\u20B9,\d.\s]+$/.test(text) ? numeric : text;
}

function compileFormula(formula, currentSheet) {
  let expression = String(formula || "").trim();
  if (expression.startsWith("=")) expression = expression.slice(1);
  const placeholders = [];
  const keep = (code) => {
    const token = `__FORMULA_TOKEN_${placeholders.length}__`;
    placeholders.push(code);
    return token;
  };

  expression = expression.replace(/(SUM|sum)\(\s*'([^']+)'!([A-Z]{1,3}\d+:[A-Z]{1,3}\d+)\s*\)/g, (_, __, sheetName, range) => keep(`sumRange(${JSON.stringify(sheetName)}, ${JSON.stringify(range)})`));
  expression = expression.replace(/(SUM|sum)\(\s*([A-Z]{1,3}\d+:[A-Z]{1,3}\d+)\s*\)/g, (_, __, range) => keep(`sumRange(${JSON.stringify(currentSheet)}, ${JSON.stringify(range)})`));
  expression = expression.replace(/'([^']+)'!([A-Z]{1,3}\d+)/g, (_, sheetName, address) => keep(`cellRef(${JSON.stringify(sheetName)}, ${JSON.stringify(address)})`));
  expression = expression.replace(/\b(INT|int)\s*\(/g, "intFn(");
  expression = expression.replace(/\b([A-Z]{1,3}\d+)\b/g, (_, address) => keep(`cellRef(${JSON.stringify(currentSheet)}, ${JSON.stringify(address)})`));
  expression = expression.replace(/__FORMULA_TOKEN_(\d+)__/g, (_, index) => placeholders[Number(index)]);

  return new Function("cellRef", "sumRange", "intFn", `return (${expression});`);
}

function createCell(address, rowIndex, columnIndex, source) {
  return {
    address,
    rowIndex,
    columnIndex,
    value: source?.v ?? "",
    formula: source?.f || "",
    format: source?.z || null,
    editable: !source?.f,
    formulaEvaluator: source?.f ? compileFormula(source.f, PREMADE_SHEET_NAME) : null,
    error: "",
  };
}

function buildSheet(sheetName, worksheet) {
  const range = XLSX.utils.decode_range(worksheet["!ref"] || "A1:A1");
  const rows = [];
  const cellMap = {};
  const formulaAddresses = [];
  const columnLabels = [];

  for (let columnIndex = range.s.c; columnIndex <= range.e.c; columnIndex += 1) {
    columnLabels.push(XLSX.utils.encode_col(columnIndex));
  }

  for (let rowIndex = range.s.r; rowIndex <= range.e.r; rowIndex += 1) {
    const row = [];
    for (let columnIndex = range.s.c; columnIndex <= range.e.c; columnIndex += 1) {
      const address = XLSX.utils.encode_cell({ r: rowIndex, c: columnIndex });
      const source = worksheet[address];
      const cell = createCell(address, rowIndex, columnIndex, source);
      if (cell.formula) {
        cell.formulaEvaluator = compileFormula(cell.formula, sheetName);
        formulaAddresses.push(address);
      }
      row.push(cell);
      cellMap[address] = cell;
    }
    rows.push(row);
  }

  return {
    name: sheetName,
    range,
    rows,
    cellMap,
    formulaAddresses,
    columnLabels,
  };
}

function cloneWorkbook(workbook) {
  const sheetNames = [...workbook.sheetNames];
  const sheets = sheetNames.map((sheetName) => {
    const source = workbook.sheets[sheetName];
    const rows = source.rows.map((row) => row.map((cell) => ({ ...cell })));
    const cellMap = {};
    for (const row of rows) {
      for (const cell of row) cellMap[cell.address] = cell;
    }
    return [sheetName, { ...source, rows, cellMap, formulaAddresses: [...source.formulaAddresses], columnLabels: [...source.columnLabels] }];
  });
  return {
    ...workbook,
    sheetNames,
    sheets: Object.fromEntries(sheets),
  };
}

function evaluateWorkbook(workbook) {
  const next = cloneWorkbook(workbook);
  const sheetNames = next.sheetNames;

  const cellRef = (sheetName, address) => {
    const cell = next.sheets[sheetName]?.cellMap[address];
    return toFormulaValue(cell?.value ?? 0);
  };

  const sumRange = (sheetName, rangeText) => {
    const range = XLSX.utils.decode_range(rangeText);
    let total = 0;
    for (let rowIndex = range.s.r; rowIndex <= range.e.r; rowIndex += 1) {
      for (let columnIndex = range.s.c; columnIndex <= range.e.c; columnIndex += 1) {
        const address = XLSX.utils.encode_cell({ r: rowIndex, c: columnIndex });
        const value = toFormulaValue(next.sheets[sheetName]?.cellMap[address]?.value ?? 0);
        total += typeof value === "number" ? value : 0;
      }
    }
    return total;
  };

  const intFn = (value) => Math.trunc(Number(value) || 0);

  for (let pass = 0; pass < 4; pass += 1) {
    for (const sheetName of sheetNames) {
      const sheet = next.sheets[sheetName];
      for (const address of sheet.formulaAddresses) {
        const cell = sheet.cellMap[address];
        try {
          cell.value = normalizeComputedValue(cell.formulaEvaluator(cellRef, sumRange, intFn));
          cell.error = "";
        } catch (error) {
          cell.error = error instanceof Error ? error.message : "Formula error";
        }
      }
    }
  }

  return next;
}

function buildWorkbook(workbookFile, title = "Workbook.xlsx") {
  const sheetNames = [...workbookFile.SheetNames];
  const sheets = Object.fromEntries(sheetNames.map((sheetName) => [sheetName, buildSheet(sheetName, workbookFile.Sheets[sheetName])]));
  return evaluateWorkbook({
    title,
    sheetNames,
    sheets,
  });
}

export async function loadPremadeWorkbook() {
  const response = await fetch(hisaabWorkbookUrl);
  const arrayBuffer = await response.arrayBuffer();
  const workbookFile = XLSX.read(arrayBuffer, { type: "array", cellFormula: true, cellNF: true, cellStyles: true, cellText: true });
  return buildWorkbook(workbookFile, "Hisaab.xlsx");
}

export async function loadHisaabWorkbookFile(file) {
  const workbookFile = XLSX.read(await file.arrayBuffer(), { type: "array", cellFormula: true, cellNF: true, cellStyles: true, cellText: true });
  return buildWorkbook(workbookFile, file.name || "Workbook.xlsx");
}

export function updatePremadeWorkbookCell(workbook, sheetName, address, inputValue) {
  const next = cloneWorkbook(workbook);
  const cell = next.sheets[sheetName]?.cellMap[address];
  if (!cell || !cell.editable) return workbook;
  cell.value = parseEditableValue(inputValue);
  cell.error = "";
  return evaluateWorkbook(next);
}

export function getPremadeSheet(workbook, sheetName) {
  return workbook?.sheets?.[sheetName] || null;
}

export function exportHisaabWorkbook(workbook) {
  const output = XLSX.utils.book_new();
  for (const sheetName of workbook.sheetNames) {
    const sheet = workbook.sheets[sheetName];
    const worksheet = {};
    for (const row of sheet.rows) {
      for (const cell of row) {
        if (cell.value === "" && !cell.formula) continue;
        worksheet[cell.address] = cell.formula
          ? { t: typeof cell.value === "number" ? "n" : "s", v: cell.value, f: cell.formula, z: cell.format || undefined }
          : { t: typeof cell.value === "number" ? "n" : "s", v: cell.value, z: cell.format || undefined };
      }
    }
    worksheet["!ref"] = XLSX.utils.encode_range(sheet.range);
    XLSX.utils.book_append_sheet(output, worksheet, sheetName.slice(0, 31));
  }
  const data = XLSX.write(output, { bookType: "xlsx", type: "array" });
  return { fileName: String(workbook.title || "Hisaab-edited.xlsx").replace(/\.xlsx$/i, "") + "-edited.xlsx", blob: new Blob([data], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" }) };
}

export { formatCellDisplay, formatEditableValue, PREMADE_SHEET_NAME };
