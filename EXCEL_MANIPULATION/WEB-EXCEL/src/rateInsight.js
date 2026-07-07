import * as XLSX from "xlsx";
import detectionConfig from "./rateInsightDetections.json";
import {
  buildSettlementRecommendationRows,
  decorateSettlementRecommendationRow,
  extractSettlementRecommendationRows,
  SETTLEMENT_RECOMMENDATION_COLUMNS,
  SETTLEMENT_RECOMMENDATION_EXPORT_COLUMNS,
  SETTLEMENT_RECOMMENDATION_MODE,
} from "./settlementRecommendations";

const SIZE_OVERRIDE_KEY = "rate_insight_size_overrides_v1";
const TITLE_FLAGS = ["Dark Blue"];
const SKU_FLAGS = [];
const FALLBACK_SIZE_VALUES = ["26", "28", "30", "32", "34", "36"];
const ORDER_CSV_VALUE_COLUMNS = ["Price inc. FKMP Contribution & Subsidy", "Invoice Amount", "Selling Price Per Item"];
const FALLBACK_KIND_DEFINITIONS = [
  { key: "ICE", label: "ICE final price floor", defaultThreshold: "469", rules: [{ skuIncludesAny: ["ice", "blue"] }] },
  { key: "BEIGE", label: "BEIGE final price floor", defaultThreshold: "459", rules: [{ skuIncludesAny: ["beige", "cream"] }] },
  { key: "WHITE", label: "WHITE final price floor", defaultThreshold: "389", rules: [{ skuIncludesAny: ["white"] }] },
  { key: "BLACK-BAGGY", label: "BLACK BAGGY floor", defaultThreshold: "429", rules: [{ skuIncludesAny: ["baggy"] }, { skuIncludesAll: ["black"], titleIncludesAny: ["relaxed"] }] },
  { key: "BLACK-PLAIN", label: "BLACK PLAIN floor", defaultThreshold: "399", rules: [{ skuIncludesAny: ["black"] }] },
  { key: "MIX", label: "MIX floor", defaultThreshold: "469", match: {} }
];

function normalizeStringList(values) {
  return [...new Set((Array.isArray(values) ? values : []).map((value) => String(value || "").trim()).filter(Boolean))];
}

const configuredSizeValues = normalizeStringList(detectionConfig.sizeOptions);
const SIZE_VALUES = configuredSizeValues.length ? configuredSizeValues : FALLBACK_SIZE_VALUES;
const KIND_DEFINITIONS = (Array.isArray(detectionConfig.kindDefinitions) ? detectionConfig.kindDefinitions : [])
  .map((entry) => ({
    key: String(entry?.key || "").trim(),
    label: String(entry?.label || "").trim(),
    defaultThreshold: String(entry?.defaultThreshold ?? "").trim(),
    match: entry?.match || {},
    rules: Array.isArray(entry?.rules) ? entry.rules : []
  }))
  .filter((entry) => entry.key);

export { SIZE_VALUES };
export const THRESHOLD_KEYS = (KIND_DEFINITIONS.length ? KIND_DEFINITIONS : FALLBACK_KIND_DEFINITIONS).map((entry) => entry.key);
export const THRESHOLD_LABELS = Object.fromEntries((KIND_DEFINITIONS.length ? KIND_DEFINITIONS : FALLBACK_KIND_DEFINITIONS).map((entry) => [entry.key, entry.label || `${entry.key} floor`]));
export const THRESHOLD_DEFAULTS = Object.fromEntries((KIND_DEFINITIONS.length ? KIND_DEFINITIONS : FALLBACK_KIND_DEFINITIONS).map((entry) => [entry.key, entry.defaultThreshold || "0"]));

const SIZE_DETECTION_RULES = Array.isArray(detectionConfig.sizeDetectionRules) ? detectionConfig.sizeDetectionRules : [];
const UNDETECTED_SIZE = String(detectionConfig.undetectedSize || "UNDETECTED").trim() || "UNDETECTED";
const INACTIVE_SIZES = normalizeStringList(detectionConfig.inactiveSizes);
const ACTIVE_KIND_DEFINITIONS = KIND_DEFINITIONS.length ? KIND_DEFINITIONS : FALLBACK_KIND_DEFINITIONS;
const FALLBACK_KIND = THRESHOLD_KEYS.includes("MIX") ? "MIX" : THRESHOLD_KEYS[THRESHOLD_KEYS.length - 1] || "UNKNOWN";

function cloneRows(rows) {
  return rows.map((row) => ({ ...row }));
}

function normalizeText(value) {
  return String(value || "").toLowerCase();
}

function testRegex(pattern, value) {
  if (!pattern) return true;
  try {
    return new RegExp(String(pattern), "i").test(value);
  } catch {
    return false;
  }
}

function includesAny(text, values) {
  const items = normalizeStringList(values).map((value) => value.toLowerCase());
  return !items.length || items.some((value) => text.includes(value));
}

function includesAll(text, values) {
  const items = normalizeStringList(values).map((value) => value.toLowerCase());
  return !items.length || items.every((value) => text.includes(value));
}

function matchesRule(match, context) {
  const rule = match || {};
  const hasCondition = [
    rule.skuRegex,
    rule.titleRegex,
    ...(Array.isArray(rule.skuIncludesAny) ? rule.skuIncludesAny : []),
    ...(Array.isArray(rule.skuIncludesAll) ? rule.skuIncludesAll : []),
    ...(Array.isArray(rule.titleIncludesAny) ? rule.titleIncludesAny : []),
    ...(Array.isArray(rule.titleIncludesAll) ? rule.titleIncludesAll : [])
  ].some(Boolean);
  if (!hasCondition) return false;

  return testRegex(rule.skuRegex, context.skuRaw)
    && testRegex(rule.titleRegex, context.titleRaw)
    && includesAny(context.skuText, rule.skuIncludesAny)
    && includesAll(context.skuText, rule.skuIncludesAll)
    && includesAny(context.titleText, rule.titleIncludesAny)
    && includesAll(context.titleText, rule.titleIncludesAll);
}

function matchesDefinition(definition, context) {
  const rules = Array.isArray(definition?.rules) && definition.rules.length ? definition.rules : [definition?.match || {}];
  return rules.some((rule) => matchesRule(rule, context));
}

function getModeColumns(mode) {
  if (mode === "orderCsv") {
    return { sku: "SKU", title: "Product", settlement: ORDER_CSV_VALUE_COLUMNS[0] };
  }
  if (mode === SETTLEMENT_RECOMMENDATION_MODE) {
    return {
      sku: SETTLEMENT_RECOMMENDATION_COLUMNS.sku,
      title: SETTLEMENT_RECOMMENDATION_COLUMNS.title,
      settlement: SETTLEMENT_RECOMMENDATION_COLUMNS.currentSettlement,
    };
  }
  return mode === "offer"
    ? { sku: "SKU ID", title: "FSN", settlement: "Selling Price(Rs)" }
    : { sku: "Seller SKU Id", title: "Product Title", settlement: "Bank Settlement" };
}

function resolveModeColumns(datasetOrMode, selectedValueColumn) {
  const mode = typeof datasetOrMode === "string" ? datasetOrMode : datasetOrMode?.mode;
  const base = getModeColumns(mode);
  if (mode !== "orderCsv") return base;
  const candidate = typeof datasetOrMode === "string"
    ? selectedValueColumn
    : datasetOrMode?.selectedValueColumn;
  return {
    ...base,
    settlement: ORDER_CSV_VALUE_COLUMNS.includes(candidate) ? candidate : ORDER_CSV_VALUE_COLUMNS[0],
  };
}

function valueColumnOptionsForMode(mode) {
  return mode === "orderCsv" ? [...ORDER_CSV_VALUE_COLUMNS] : [];
}

function defaultValueColumnForMode(mode) {
  return valueColumnOptionsForMode(mode)[0] || null;
}

function cleanNumeric(value) {
  if (typeof value === "number") return Number.isFinite(value) ? value : null;
  const text = String(value ?? "")
    .replaceAll(",", "")
    .replaceAll("?", "")
    .replaceAll("₹", "")
    .replace(/^Rs\s*/i, "")
    .trim();
  if (!text) return null;
  const numeric = Number(text);
  return Number.isFinite(numeric) ? numeric : null;
}

function detectAccount(fileName) {
  const lowered = String(fileName || "").toLowerCase();
  if (lowered.includes("84f77")) return "Prabhu";
  if (lowered.includes("946b8")) return "Seema";
  return "Unknown";
}

function getFileType(fileName) {
  const ext = String(fileName || "").split(".").pop()?.toLowerCase();
  return ext || "xlsx";
}

function listingType(title) {
  const owners = ["Starvielle", "Genz Vane", "INDIVANE", "FADEVIELLE", "FLEECRANE"];
  const text = normalizeText(title);
  return owners.some((item) => text.includes(item.toLowerCase())) ? "Owner" : "Latched";
}

function jeansType(skuValue, titleValue) {
  const context = {
    skuRaw: String(skuValue || ""),
    titleRaw: String(titleValue || ""),
    skuText: normalizeText(skuValue),
    titleText: normalizeText(titleValue)
  };
  for (const definition of ACTIVE_KIND_DEFINITIONS) {
    if (matchesDefinition(definition, context)) return definition.key;
  }
  return FALLBACK_KIND;
}

function loadSizeOverrides() {
  try {
    return JSON.parse(window.localStorage.getItem(SIZE_OVERRIDE_KEY) || "{}") || {};
  } catch {
    return {};
  }
}

function persistSizeOverrides(overrides) {
  window.localStorage.setItem(SIZE_OVERRIDE_KEY, JSON.stringify(overrides));
}

function detectSize(skuValue, overrides) {
  const sku = String(skuValue || "").trim();
  if (!sku) return UNDETECTED_SIZE;
  if (overrides[sku]) return overrides[sku];
  for (const rule of SIZE_DETECTION_RULES) {
    const size = String(rule?.size || "").trim();
    if (size && matchesRule(rule, { skuRaw: sku, titleRaw: "", skuText: sku.toLowerCase(), titleText: "" })) return size;
  }
  for (const size of SIZE_VALUES) {
    if (new RegExp(`-${size}-|-${size}_|_${size}_|_${size}$`, "i").test(sku)) return size;
  }
  for (const size of FALLBACK_SIZE_VALUES) {
    if (SIZE_VALUES.includes(size)) continue;
    if (new RegExp(`-${size}-|-${size}_|_${size}_|_${size}$`, "i").test(sku)) return size;
  }
  return UNDETECTED_SIZE;
}

function applyFlags(rows, columns, overrides) {
  const titleKeywords = TITLE_FLAGS;
  const skuKeywords = SKU_FLAGS;
  const inactiveSizes = INACTIVE_SIZES;

  return rows.map((row) => {
    const next = { ...row };
    const title = String(next[columns.title] || "").toLowerCase();
    const sku = String(next[columns.sku] || "").toLowerCase();
    const size = String(next.Size || detectSize(next[columns.sku], overrides)).trim();
    const reasons = [];
    const titleMatches = titleKeywords.filter((item) => item && title.includes(String(item).toLowerCase()));
    const skuMatches = skuKeywords.filter((item) => item && sku.includes(String(item).toLowerCase()));

    if (inactiveSizes.includes(size)) reasons.push(`SIZE: ${size}`);
    if (titleMatches.length) reasons.push(`TITLE: ${titleMatches.join(" | ")}`);
    if (skuMatches.length) reasons.push(`SKU: ${skuMatches.join(" | ")}`);

    next["Auto Flag"] = reasons.join(" | ");
    if (next["Auto Flag"]) next["Listing Status"] = "INACTIVE";
    return next;
  });
}

function syncStock(rows) {
  return rows.map((row) => {
    const next = { ...row };
    if (!("Your Stock Count" in next) || !("Listing Status" in next)) return next;
    const status = String(next["Listing Status"] || "").trim().toUpperCase();
    if (status === "ACTIVE") next["Your Stock Count"] = 250;
    if (status === "INACTIVE") next["Your Stock Count"] = 0;
    return next;
  });
}

function rebuildRows(rows, mode, overrides) {
  const columns = resolveModeColumns(mode);
  const base = cloneRows(rows).map((row) => {
    const next = { ...row };
    next[columns.settlement] = cleanNumeric(next[columns.settlement]);
    next["Listing Type"] = listingType(next[columns.title]);
    next["Jeans Type"] = jeansType(next[columns.sku], next[columns.title]);
    next.Size = detectSize(next[columns.sku], overrides);
    if ("Listing Status" in next) next["Listing Status"] = String(next["Listing Status"] || "").trim().toUpperCase();
    return next;
  });
  return syncStock(applyFlags(base, columns, overrides));
}

function mm(values) {
  const clean = values.filter((value) => typeof value === "number" && Number.isFinite(value));
  if (!clean.length) return "NA / NA";
  return `${Math.min(...clean).toFixed(2)} / ${Math.max(...clean).toFixed(2)}`;
}

function computeSelectionLabel(values) {
  const clean = values.filter((value) => typeof value === "number" && Number.isFinite(value));
  if (!clean.length) return "No Selection";
  const sorted = [...clean].sort((a, b) => a - b);
  const mean = clean.reduce((sum, value) => sum + value, 0) / clean.length;
  const mid = Math.floor(sorted.length / 2);
  const median = sorted.length % 2 ? sorted[mid] : (sorted[mid - 1] + sorted[mid]) / 2;
  const counts = new Map();
  let modeValue = "NA";
  let best = 0;
  for (const value of clean) {
    const next = (counts.get(value) || 0) + 1;
    counts.set(value, next);
    if (next > best) {
      best = next;
      modeValue = value;
    }
  }
  return `Count: ${clean.length}  Mean: ${mean.toFixed(2)}  Median: ${median.toFixed(2)}  Mode: ${modeValue}`;
}

function normalizeFilters(dataset, raw) {
  const base = defaultFiltersForDataset(dataset);
  const data = { ...base, ...(raw || {}) };
  data.search = String(data.search || "").trim();
  for (const key of ["listing", "jeans", "size", "status"]) data[key] = String(data[key] || "All");
  data.settlementMax = Number(data.settlementMax ?? base.settlementMax) || 0;
  if (Array.isArray(data.selectedRange) && data.selectedRange.length === 2) {
    const low = Number(data.selectedRange[0]);
    const high = Number(data.selectedRange[1]);
    data.selectedRange = [Math.min(low, high), Math.max(low, high)];
  } else {
    data.selectedRange = null;
  }
  return data;
}

function filteredRows(dataset, rawFilters, includeSelection) {
  const filters = normalizeFilters(dataset, rawFilters);
  const columns = resolveModeColumns(dataset);
  let rows = dataset.rows.filter((row) => !row.__locked && row[columns.settlement] != null && Number(row[columns.settlement]) <= filters.settlementMax);

  if (filters.search) {
    const search = filters.search.toLowerCase();
    rows = rows.filter((row) => String(row[columns.sku] || "").toLowerCase().includes(search) || String(row[columns.title] || "").toLowerCase().includes(search));
  }
  if (filters.listing !== "All") rows = rows.filter((row) => row["Listing Type"] === filters.listing);
  if (filters.jeans !== "All") rows = rows.filter((row) => row["Jeans Type"] === filters.jeans);
  if (filters.size !== "All") rows = rows.filter((row) => String(row.Size || "") === filters.size);
  if (filters.status !== "All") rows = rows.filter((row) => String(row["Listing Status"] || "").toUpperCase() === filters.status);
  if (includeSelection && filters.selectedRange) {
    const [low, high] = filters.selectedRange;
    rows = rows.filter((row) => {
      const value = Number(row[columns.settlement]);
      return Number.isFinite(value) && value >= low && value <= high;
    });
  }
  return rows;
}

function buildOptions(dataset) {
  const columns = resolveModeColumns(dataset);
  const settlementValues = dataset.rows.map((row) => Number(row[columns.settlement])).filter((value) => Number.isFinite(value));
  const sizeSet = new Set(dataset.rows.map((row) => String(row.Size || "")).filter(Boolean));
  const statusSet = new Set(dataset.rows.map((row) => String(row["Listing Status"] || "").toUpperCase()).filter(Boolean));
  const sizeOptions = SIZE_VALUES.filter((size) => sizeSet.has(size));
  if (sizeSet.has(UNDETECTED_SIZE)) sizeOptions.push(UNDETECTED_SIZE);
  return {
    listing: [...new Set(dataset.rows.map((row) => row["Listing Type"]).filter(Boolean))].sort(),
    jeans: [...new Set(dataset.rows.map((row) => row["Jeans Type"]).filter(Boolean))].sort(),
    size: sizeOptions,
    status: ["ACTIVE", "INACTIVE"].filter((status) => statusSet.has(status)),
    settlementMax: settlementValues.length ? Math.max(...settlementValues) : 0,
  };
}
function buildListingTypeRatio(rows) {
  const counts = rows.reduce((acc, row) => {
    const key = String(row["Listing Type"] || "Unknown");
    acc[key] = (acc[key] || 0) + 1;
    return acc;
  }, {});
  return [
    { label: "Owner", value: counts.Owner || 0 },
    { label: "Latched", value: counts.Latched || 0 },
  ];
}

function buildSizeColorStatusRatio(rows) {
  const groupMap = new Map();
  for (const row of rows) {
    const size = String(row.Size || UNDETECTED_SIZE);
    const color = String(row["Jeans Type"] || "MIX");
    const status = String(row["Listing Status"] || "").toUpperCase();
    const key = `${size} | ${color}`;
    const current = groupMap.get(key) || { size, color, active: 0, inactive: 0, total: 0 };
    current.total += 1;
    if (status === "ACTIVE") current.active += 1;
    if (status === "INACTIVE") current.inactive += 1;
    groupMap.set(key, current);
  }

  const statusCounts = { ACTIVE: 0, INACTIVE: 0, BLEND: 0 };
  const breakdown = [...groupMap.values()].map((entry) => {
    const state = entry.active > 0 && entry.inactive > 0 ? "BLEND" : entry.active > 0 ? "ACTIVE" : "INACTIVE";
    statusCounts[state] += entry.total;
    return {
      label: `${entry.size} / ${entry.color}`,
      state,
      count: entry.total,
    };
  }).sort((a, b) => b.count - a.count || a.label.localeCompare(b.label));

  return {
    ratio: [
      { label: "Active groups", value: statusCounts.ACTIVE },
      { label: "Inactive groups", value: statusCounts.INACTIVE },
      { label: "Blend groups", value: statusCounts.BLEND },
    ],
    breakdown,
  };
}

function appendLog(dataset, action, beforeValues, afterValues, rowCount, extra = "") {
  const next = { ...dataset };
  next.changeLog = [
    ...next.changeLog,
    { action, rows: rowCount, before: mm(beforeValues), after: mm(afterValues), extra },
  ].slice(-12);
  return next;
}

function withUndo(dataset) {
  return { ...dataset, undoRows: cloneRows(dataset.rows) };
}

function updateRows(dataset, rowIndexes, updater) {
  const nextRows = cloneRows(dataset.rows);
  for (const index of rowIndexes) nextRows[index] = updater({ ...nextRows[index] }, index);
  return nextRows;
}

export async function loadWorkbook(file) {
  const workbook = XLSX.read(await file.arrayBuffer(), { type: "array" });
  const sizeOverrides = loadSizeOverrides();
  const settlementRecommendationRows = extractSettlementRecommendationRows(workbook);

  if (settlementRecommendationRows?.length) {
    const preparedRows = buildSettlementRecommendationRows(
      settlementRecommendationRows.map((row, index) => ({ ...row, __orig_index: index })),
      {
        cleanNumeric,
        detectSize,
        jeansType,
        listingType,
        overrides: sizeOverrides,
        thresholds: THRESHOLD_DEFAULTS,
      }
    );
    return {
      fileName: file.name,
      fileType: getFileType(file.name),
      accountName: detectAccount(file.name),
      availableModes: { offer: false, normal: false, orderCsv: false, settlementRecommendations: true },
      mode: SETTLEMENT_RECOMMENDATION_MODE,
      selectedValueColumn: null,
      rows: preparedRows,
      sizeOverrides,
      undoRows: null,
      changeLog: [],
    };
  }

  const sheetName = workbook.SheetNames[0];
  const rows = XLSX.utils.sheet_to_json(workbook.Sheets[sheetName], { defval: "" });
  const columns = rows[0] ? Object.keys(rows[0]) : [];
  const availableModes = {
    offer: ["SKU ID", "FSN", "Selling Price(Rs)"].every((key) => columns.includes(key)),
    normal: ["Seller SKU Id", "Product Title", "Bank Settlement"].every((key) => columns.includes(key)),
    orderCsv: ["SKU", "Product", ...ORDER_CSV_VALUE_COLUMNS].every((key) => columns.includes(key)),
    settlementRecommendations: false,
  };
  if (!availableModes.offer && !availableModes.normal && !availableModes.orderCsv) {
    throw new Error("Unknown file format. Required settlement columns were not found.");
  }
  const mode = availableModes.normal ? "normal" : availableModes.offer ? "offer" : "orderCsv";
  const preparedRows = rows.map((row, index) => ({ ...row, __orig_index: index, __locked: false }));
  return {
    fileName: file.name,
    fileType: getFileType(file.name),
    accountName: detectAccount(file.name),
    availableModes,
    mode,
    selectedValueColumn: defaultValueColumnForMode(mode),
    rows: rebuildRows(preparedRows, mode, sizeOverrides),
    sizeOverrides,
    undoRows: null,
    changeLog: [],
  };
}

export function defaultFiltersForDataset(dataset) {
  if (!dataset) return { search: "", listing: "All", jeans: "All", size: "All", status: "All", settlementMax: 0, selectedRange: null };
  const columns = resolveModeColumns(dataset);
  const max = Math.max(0, ...dataset.rows.map((row) => Number(row[columns.settlement]) || 0));
  return { search: "", listing: "All", jeans: "All", size: "All", status: "All", settlementMax: max, selectedRange: null };
}

export function createSnapshot(dataset, rawFilters) {
  const filters = normalizeFilters(dataset, rawFilters);
  const columns = resolveModeColumns(dataset);
  const visible = filteredRows(dataset, filters, false);
  const exported = filteredRows(dataset, filters, true);
  const values = visible.map((row) => Number(row[columns.settlement])).filter((value) => Number.isFinite(value));
  const frequency = new Map();
  for (const value of values) frequency.set(value, (frequency.get(value) || 0) + 1);
  const chart = [...frequency.entries()].sort((a, b) => a[0] - b[0]).map(([value, count]) => ({
    value,
    count,
    selected: filters.selectedRange ? value >= filters.selectedRange[0] && value <= filters.selectedRange[1] : false,
  }));
  const statusCounts = visible.reduce((acc, row) => {
    const status = String(row["Listing Status"] || "").toUpperCase();
    if (status === "ACTIVE") acc.active += 1;
    if (status === "INACTIVE") acc.inactive += 1;
    return acc;
  }, { active: 0, inactive: 0 });
  const matrix = ["ACTIVE", "INACTIVE"].map((status) => ({
    status,
    count: exported.filter((row) => String(row["Listing Status"] || "").toUpperCase() === status).length,
    sizes: [...new Set(exported.filter((row) => String(row["Listing Status"] || "").toUpperCase() === status).map((row) => String(row.Size || "")).filter(Boolean))].sort(),
  }));
  const selectionValues = exported.map((row) => Number(row[columns.settlement])).filter((value) => Number.isFinite(value));
  const listingTypeRatio = buildListingTypeRatio(exported);
  const sizeColorStatus = buildSizeColorStatusRatio(exported);
  const rowColumns = dataset.mode === SETTLEMENT_RECOMMENDATION_MODE ? SETTLEMENT_RECOMMENDATION_EXPORT_COLUMNS : [...new Set(dataset.rows.flatMap((row) => Object.keys(row)))].filter((key) => !["__locked", "__orig_index"].includes(key));

  return {
    fileName: dataset.fileName,
    accountName: dataset.accountName,
    mode: dataset.mode,
    availableModes: dataset.availableModes,
    selectedValueColumn: dataset.selectedValueColumn,
    availableValueColumns: valueColumnOptionsForMode(dataset.mode),
    filters,
    filterOptions: buildOptions(dataset),
    columns: rowColumns,
    rows: exported.slice(0, 400),
    visibleCount: visible.length,
    exportCount: exported.length,
    chart,
    metrics: {
      loaded: dataset.rows.length,
      visible: visible.length,
      export: exported.length,
      active: statusCounts.active,
      inactive: statusCounts.inactive,
      activePct: visible.length ? Number(((statusCounts.active / visible.length) * 100).toFixed(1)) : 0,
      inactivePct: visible.length ? Number(((statusCounts.inactive / visible.length) * 100).toFixed(1)) : 0,
    },
    summary: {
      rowCountText: `Loaded: ${dataset.rows.length} | Visible: ${visible.length} | Export: ${exported.length}`,
      selectionRange: filters.selectedRange,
      selectionStats: { label: filters.selectedRange ? computeSelectionLabel(selectionValues) : "No Selection" },
      statusMatrix: matrix,
      listingTypeRatio,
      sizeColorStatusRatio: sizeColorStatus.ratio,
      sizeColorBreakdown: sizeColorStatus.breakdown,
    },
    changeLog: dataset.changeLog,
  };
}

export function setMode(dataset, mode) {
  if (!dataset.availableModes[mode]) throw new Error(`${mode} mode is not available for this file.`);
  const nextValueColumn = valueColumnOptionsForMode(mode).includes(dataset.selectedValueColumn)
    ? dataset.selectedValueColumn
    : defaultValueColumnForMode(mode);
  return { ...dataset, mode, selectedValueColumn: nextValueColumn, rows: rebuildRows(dataset.rows, mode, dataset.sizeOverrides) };
}

export function setValueColumn(dataset, valueColumn) {
  if (dataset.mode !== "orderCsv") return dataset;
  if (!ORDER_CSV_VALUE_COLUMNS.includes(valueColumn)) throw new Error("Invalid value column");
  return { ...dataset, selectedValueColumn: valueColumn };
}

export function bulkEdit(dataset, filters, modeName, value, capMode, capValue) {
  if (dataset.mode === SETTLEMENT_RECOMMENDATION_MODE) throw new Error("Bulk editing is disabled for settlement recommendations files");
  const columns = resolveModeColumns(dataset);
  const selected = new Set(filteredRows(dataset, filters, true).map((row) => row.__orig_index));
  const rowIndexes = dataset.rows.map((row, index) => (selected.has(row.__orig_index) ? index : -1)).filter((index) => index >= 0);
  if (!rowIndexes.length) throw new Error("No visible unlocked rows selected");
  const before = rowIndexes.map((index) => Number(dataset.rows[index][columns.settlement]));
  let next = withUndo(dataset);
  next.rows = updateRows(next, rowIndexes, (row) => {
    const current = Number(row[columns.settlement]) || 0;
    if (modeName === "Add") row[columns.settlement] = current + value;
    else if (modeName === "Multiply") {
      let result = current * value;
      if (capValue != null && Number.isFinite(capValue)) result = capMode === "Min" ? Math.min(result, capValue) : Math.max(result, capValue);
      row[columns.settlement] = result;
    } else row[columns.settlement] = value;
    return row;
  });
  const after = rowIndexes.map((index) => Number(next.rows[index][columns.settlement]));
  return appendLog(next, "Bulk Edit", before, after, rowIndexes.length, `Mode: ${modeName}, Value: ${value}`);
}

export function setStatus(dataset, filters, status) {
  if (dataset.mode === SETTLEMENT_RECOMMENDATION_MODE) throw new Error("Status changes are disabled for settlement recommendations files");
  const selected = new Set(filteredRows(dataset, filters, true).map((row) => row.__orig_index));
  const rowIndexes = dataset.rows.map((row, index) => (selected.has(row.__orig_index) ? index : -1)).filter((index) => index >= 0);
  if (!rowIndexes.length) throw new Error("No visible unlocked rows to update");
  let next = withUndo(dataset);
  const columns = resolveModeColumns(next);
  const before = rowIndexes.map((index) => Number(next.rows[index][columns.settlement]));
  next.rows = syncStock(updateRows(next, rowIndexes, (row) => {
    row["Listing Status"] = status;
    return row;
  }));
  const after = rowIndexes.map((index) => Number(next.rows[index][columns.settlement]));
  return appendLog(next, "Set Status", before, after, rowIndexes.length, `Status: ${status}`);
}

export function computeDiscount(dataset, filters, yPct, xPct, cap) {
  if (dataset.mode === SETTLEMENT_RECOMMENDATION_MODE) throw new Error("Offer calculations are disabled for settlement recommendations files");
  const columns = resolveModeColumns(dataset);
  const selected = new Set(filteredRows(dataset, filters, true).map((row) => row.__orig_index));
  const rowIndexes = dataset.rows.map((row, index) => (selected.has(row.__orig_index) ? index : -1)).filter((index) => index >= 0);
  if (!rowIndexes.length) throw new Error("No visible unlocked rows selected");
  const before = rowIndexes.map((index) => Number(dataset.rows[index][columns.settlement]));
  let next = withUndo(dataset);
  next.rows = updateRows(next, rowIndexes, (row) => {
    const settlement = Number(row[columns.settlement]) || 0;
    const base = (yPct / 100) * settlement;
    const discount = Math.min((xPct / 100) * base, cap);
    row.Discount = Number(discount.toFixed(2));
    row["Final Price"] = Number((settlement - discount).toFixed(2));
    return row;
  });
  const after = rowIndexes.map((index) => Number(next.rows[index][columns.settlement]));
  return appendLog(next, "Compute Discount", before, after, rowIndexes.length);
}

export function applyDecision(dataset, filters, thresholds) {
  if (dataset.mode === SETTLEMENT_RECOMMENDATION_MODE) throw new Error("Offer decisions are disabled for settlement recommendations files");
  if (!dataset.rows.some((row) => row["Final Price"] != null && row["Final Price"] !== "")) throw new Error("Compute discount first");
  const selected = new Set(filteredRows(dataset, filters, true).map((row) => row.__orig_index));
  const rowIndexes = dataset.rows.map((row, index) => (selected.has(row.__orig_index) ? index : -1)).filter((index) => index >= 0);
  if (!rowIndexes.length) throw new Error("No visible unlocked rows selected");
  const columns = resolveModeColumns(dataset);
  const before = rowIndexes.map((index) => Number(dataset.rows[index][columns.settlement]));
  let next = withUndo(dataset);
  next.rows = updateRows(next, rowIndexes, (row) => {
    const threshold = Number(thresholds[String(row["Jeans Type"])]) || 0;
    row.Decision = Number(row["Final Price"] || 0) >= threshold ? "ACCEPT" : "REJECT";
    return row;
  });
  const after = rowIndexes.map((index) => Number(next.rows[index][columns.settlement]));
  return appendLog(next, "Apply Decision", before, after, rowIndexes.length);
}

export function saveSizeOverride(dataset, filters, sku, size) {
  if (dataset.mode === SETTLEMENT_RECOMMENDATION_MODE) throw new Error("Size overrides are disabled for settlement recommendations files");
  const trimmedSku = String(sku || "").trim();
  if (!trimmedSku) throw new Error("Enter a SKU");
  if (!SIZE_VALUES.includes(size)) throw new Error("Invalid size");
  const nextOverrides = { ...dataset.sizeOverrides, [trimmedSku]: size };
  persistSizeOverrides(nextOverrides);
  let next = withUndo(dataset);
  next.sizeOverrides = nextOverrides;
  next.rows = rebuildRows(next.rows, next.mode, nextOverrides);
  const columns = resolveModeColumns(next);
  const selected = filteredRows(next, filters, true).filter((row) => String(row[columns.sku] || "").trim() === trimmedSku);
  const values = selected.map((row) => Number(row[columns.settlement])).filter((value) => Number.isFinite(value));
  return appendLog(next, "Save Size", values, values, selected.length, `SKU: ${trimmedSku} -> ${size}`);
}

export function undoDataset(dataset) {
  if (!dataset.undoRows) return dataset;
  return { ...dataset, rows: cloneRows(dataset.undoRows), undoRows: null };
}

export function exportDataset(dataset) {
  const rows = cloneRows(dataset.rows)
    .sort((a, b) => a.__orig_index - b.__orig_index)
    .map(({ __orig_index, __locked, ...row }) => row);
  const exportRows = dataset.mode === SETTLEMENT_RECOMMENDATION_MODE
    ? rows.map((row) => Object.fromEntries(SETTLEMENT_RECOMMENDATION_EXPORT_COLUMNS.map((column) => [column, row[column] ?? ""])))
    : rows;
  const sheet = XLSX.utils.json_to_sheet(exportRows);

  if (dataset.fileType === "csv") {
    const csv = XLSX.utils.sheet_to_csv(sheet);
    return {
      fileName: "PROGRAM_OUTPUTTED.csv",
      blob: new Blob([csv], { type: "text/csv;charset=utf-8" }),
    };
  }

  const workbook = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(workbook, sheet, "RateInsight");
  const output = XLSX.write(workbook, { bookType: "xlsx", type: "array" });
  return {
    fileName: dataset.mode === SETTLEMENT_RECOMMENDATION_MODE ? "SETTLEMENT_RECOMMENDATIONS_OUTPUT.xlsx" : "PROGRAM_OUTPUTTED.xlsx",
    blob: new Blob([output], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" }),
  };
}










