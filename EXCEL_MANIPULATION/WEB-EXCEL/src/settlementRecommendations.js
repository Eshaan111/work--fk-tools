import * as XLSX from "xlsx";

export const SETTLEMENT_RECOMMENDATION_MODE = "settlementRecommendations";
export const SETTLEMENT_RECOMMENDATION_COLUMNS = {
  sku: "SKU",
  title: "Product Name",
  currentSettlement: "Current Settlement",
  recommendedSettlement: "Recommended Settlement",
  recommendedRange: "Recommended Settlement Range",
};
export const SETTLEMENT_RECOMMENDATION_EXPORT_COLUMNS = [
  "SKU",
  "Product Name",
  "Kind",
  "Current Settlement",
  "Recommended Settlement",
  "Recommended Settlement Range",
  "Accept / Reject",
  "Output Settlement",
];

function parseRangeMax(value, cleanNumeric) {
  const text = String(value || "").trim();
  if (!text) return null;
  const parts = text.split("-").map((item) => cleanNumeric(item));
  const maxValue = parts[parts.length - 1];
  return Number.isFinite(maxValue) ? maxValue : null;
}

export function extractSettlementRecommendationRows(workbook) {
  const sheetName = workbook.SheetNames[0];
  const sheet = workbook.Sheets[sheetName];
  if (!sheet) return null;
  const rows = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: "" });
  const headerRow = rows[2] || [];
  const required = Object.values(SETTLEMENT_RECOMMENDATION_COLUMNS);
  if (!required.every((header) => headerRow.includes(header))) return null;
  return rows.slice(4)
    .filter((row) => row.some((value) => String(value || "").trim()))
    .map((row) => Object.fromEntries(headerRow.map((header, index) => [String(header || "").trim(), row[index] ?? ""])));
}

export function decorateSettlementRecommendationRow(row, helpers) {
  const { cleanNumeric, detectSize, jeansType, listingType, overrides, thresholds } = helpers;
  const sku = row[SETTLEMENT_RECOMMENDATION_COLUMNS.sku];
  const title = row[SETTLEMENT_RECOMMENDATION_COLUMNS.title];
  const kind = jeansType(sku, title);
  const threshold = Number(thresholds[String(kind)] || 0);
  const recommended = cleanNumeric(row[SETTLEMENT_RECOMMENDATION_COLUMNS.recommendedSettlement]);
  const rangeMax = parseRangeMax(row[SETTLEMENT_RECOMMENDATION_COLUMNS.recommendedRange], cleanNumeric);
  const acceptedValue = recommended != null && recommended >= threshold
    ? recommended
    : rangeMax != null && rangeMax >= threshold
      ? rangeMax
      : null;

  return {
    ...row,
    __locked: false,
    Kind: kind,
    Size: detectSize(sku, overrides),
    "Listing Type": listingType(title),
    "Jeans Type": kind,
    [SETTLEMENT_RECOMMENDATION_COLUMNS.currentSettlement]: cleanNumeric(row[SETTLEMENT_RECOMMENDATION_COLUMNS.currentSettlement]),
    [SETTLEMENT_RECOMMENDATION_COLUMNS.recommendedSettlement]: recommended,
    "Accept / Reject": acceptedValue != null ? "ACCEPT" : "REJECT",
    "Output Settlement": acceptedValue != null ? acceptedValue : "",
  };
}

export function buildSettlementRecommendationRows(rows, helpers) {
  return rows.map((row) => decorateSettlementRecommendationRow(row, helpers));
}
