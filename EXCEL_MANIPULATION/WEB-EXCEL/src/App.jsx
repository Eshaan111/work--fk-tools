import "./styles.css";

import React, { useEffect, useMemo, useRef, useState } from "react";
import HisaabPage from "./HisaabPage";
import {
  THRESHOLD_DEFAULTS,
  THRESHOLD_KEYS,
  THRESHOLD_LABELS,
  SETTLEMENT_THRESHOLD_DEFAULTS,
  SETTLEMENT_THRESHOLD_LABELS,
  SIZE_VALUES,
  applyDecision,
  bulkEdit,
  computeDiscount,
  createSnapshot,
  defaultFiltersForDataset,
  exportDataset,
  loadWorkbook,
  saveSizeOverride,
  setMode,
  setStatus,
  setValueColumn,
  undoDataset,
} from "./rateInsight";

const defaultFilters = { search: "", listing: "All", jeans: "All", size: "All", status: "All", settlementMax: 0, selectedRange: null };
const PRICE_THRESHOLD_MODE = "price";
const SETTLEMENT_THRESHOLD_MODE = "settlement";

const LISTING_COLUMN_VIEWS = {
  overview: {
    label: "Overview",
    columns: ["Product Title", "Product", "Product Name", "FSN", "Seller SKU Id", "SKU ID", "SKU", "Bank Settlement", "Current Settlement", "Recommended Settlement", "Recommended Settlement Range", "Accept / Reject", "Output Settlement", "Selling Price(Rs)", "Price inc. FKMP Contribution & Subsidy", "Invoice Amount", "Selling Price Per Item", "Listing Status", "Size", "Jeans Type", "Listing Type", "Auto Flag", "Kind"],
  },
  pricing: {
    label: "Pricing",
    columns: ["Product Title", "Product", "Product Name", "FSN", "Seller SKU Id", "SKU ID", "SKU", "Bank Settlement", "Current Settlement", "Recommended Settlement", "Recommended Settlement Range", "Accept / Reject", "Output Settlement", "Selling Price(Rs)", "Price inc. FKMP Contribution & Subsidy", "Invoice Amount", "Selling Price Per Item", "Discount", "Final Price", "Decision", "MRP (?)", "Your Selling Price (Rs)", "Kind"],
  },
  status: {
    label: "Status & Flags",
    columns: ["Product Title", "Product", "Product Name", "FSN", "Seller SKU Id", "SKU ID", "SKU", "Listing Status", "Accept / Reject", "Output Settlement", "Your Stock Count", "Size", "Jeans Type", "Listing Type", "Auto Flag", "Kind"],
  },
  all: {
    label: "All Columns",
    columns: null,
  },
};

function visibleListingColumns(allColumns, viewKey) {
  const config = LISTING_COLUMN_VIEWS[viewKey] || LISTING_COLUMN_VIEWS.overview;
  if (!config.columns) return allColumns;
  const picked = config.columns.filter((column) => allColumns.includes(column));
  const fallback = allColumns.filter((column) => ["Product Title", "FSN", "Seller SKU Id", "SKU ID", "Bank Settlement", "Selling Price(Rs)", "Listing Status"].includes(column));
  return (picked.length ? picked : fallback.length ? fallback : allColumns).concat(allColumns.filter((column) => picked.includes(column) === false && fallback.includes(column) === false && ["__orig_index"].includes(column) === false).slice(0, 0));
}
function HelpTip({ text }) {
  return (
    <span className="help-tip" tabIndex={0} aria-label="More information">
      ?
      <span className="help-bubble">{text}</span>
    </span>
  );
}

function LabelWithHelp({ label, help }) {
  return (
    <span className="label-with-help">
      <span>{label}</span>
      {help ? <HelpTip text={help} /> : null}
    </span>
  );
}

function Modal({ title, children, onClose, actions, size = "normal" }) {
  return (
    <div className="modal-backdrop" role="presentation" onMouseDown={(event) => event.target === event.currentTarget && onClose()}>
      <section className={`modal-card modal-${size}`} role="dialog" aria-modal="true" aria-labelledby="modal-title">
        <div className="modal-head"><h2 id="modal-title">{title}</h2><button className="icon-button" onClick={onClose} aria-label="Close dialog">×</button></div>
        <div className="modal-body">{children}</div>
        <div className="modal-actions">{actions}</div>
      </section>
    </div>
  );
}

function Notice({ notice, onClose }) {
  if (!notice) return null;
  return <div className={`notice notice-${notice.type || "success"}`} role="status"><span>{notice.message}</span><button onClick={onClose} aria-label="Dismiss notification">×</button></div>;
}

function metricCards(metrics) {
  return [
    { key: "loaded", label: "Total Loaded", value: metrics?.loaded ?? 0, detail: "Rows in file" },
    { key: "visible", label: "Visible Listings", value: metrics?.visible ?? 0, detail: "Rows after filters" },
    { key: "export", label: "Exported Rows", value: metrics?.export ?? 0, detail: "Rows in selection" },
    { key: "active", label: "Active Listings", value: metrics?.active ?? 0, detail: `${metrics?.activePct ?? 0}% of visible` },
    { key: "inactive", label: "Inactive Listings", value: metrics?.inactive ?? 0, detail: `${metrics?.inactivePct ?? 0}% of visible` },
  ];
}

function thresholdDefaultsForMode(mode) {
  return mode === SETTLEMENT_THRESHOLD_MODE ? SETTLEMENT_THRESHOLD_DEFAULTS : THRESHOLD_DEFAULTS;
}

function thresholdLabelsForMode(mode) {
  return mode === SETTLEMENT_THRESHOLD_MODE ? SETTLEMENT_THRESHOLD_LABELS : THRESHOLD_LABELS;
}

function buildThresholdMap(defaults, current = {}) {
  return Object.fromEntries(THRESHOLD_KEYS.map((key) => [key, current?.[key] ?? defaults[key] ?? "0"]));
}

function modeMeta(mode, selectedValueColumn, thresholdMode = PRICE_THRESHOLD_MODE) {
  if (mode === "offer") {
    return {
      modeLabel: "Offer File",
      valueLabel: "Selling Price(Rs)",
      valueShort: "selling price",
      thresholdBasis: "Final Price",
    };
  }
  if (mode === "settlementRecommendations") {
    return {
      modeLabel: "Settlement Recommendations File",
      valueLabel: "Current Settlement",
      valueShort: "current settlement",
      thresholdBasis: thresholdMode === SETTLEMENT_THRESHOLD_MODE ? "Settlement Threshold" : "Price Threshold",
    };
  }
  if (mode === "orderCsv") {
    const label = selectedValueColumn || "Price inc. FKMP Contribution & Subsidy";
    return {
      modeLabel: "Order CSV",
      valueLabel: label,
      valueShort: label.toLowerCase(),
      thresholdBasis: "Final Price",
    };
  }
  return {
    modeLabel: "Standard File",
    valueLabel: "Bank Settlement",
    valueShort: "bank settlement",
    thresholdBasis: "Final Price",
  };
}

function MiniLineChart({ data, selectedRange, onRangeChange }) {
  const [draftRange, setDraftRange] = useState(selectedRange || null);
  const [dragRange, setDragRange] = useState(null);
  const svgRef = useRef(null);

  useEffect(() => {
    setDraftRange(selectedRange || null);
  }, [selectedRange]);

  if (!data.length) {
    return <div className="empty-state chart-empty">Load a workbook to see settlement distribution.</div>;
  }

  const width = 900;
  const height = 320;
  const padding = { top: 24, right: 28, bottom: 50, left: 56 };
  const plotWidth = width - padding.left - padding.right;
  const plotHeight = height - padding.top - padding.bottom;
  const values = data.map((point) => Number(point.value || 0));
  const counts = data.map((point) => Number(point.count || 0));
  const minX = Math.min(...values);
  const maxX = Math.max(...values);
  const maxY = Math.max(...counts, 1);
  const xSpan = Math.max(maxX - minX, 1);

  function xScale(value) {
    return padding.left + ((value - minX) / xSpan) * plotWidth;
  }

  function yScale(value) {
    return padding.top + plotHeight - (value / maxY) * plotHeight;
  }

  function clampValue(value) {
    return Math.max(minX, Math.min(maxX, value));
  }

  function commitRange(nextMin, nextMax) {
    const safeMin = clampValue(Math.min(nextMin, nextMax));
    const safeMax = clampValue(Math.max(nextMin, nextMax));
    const next = [safeMin, safeMax];
    setDraftRange(next);
    onRangeChange(next);
  }

  function pointerToValue(event) {
    const svg = svgRef.current;
    if (!svg) return minX;
    const rect = svg.getBoundingClientRect();
    const rawX = ((event.clientX - rect.left) / rect.width) * width;
    const clampedX = Math.max(padding.left, Math.min(width - padding.right, rawX));
    return minX + ((clampedX - padding.left) / plotWidth) * xSpan;
  }

  function handlePointerDown(event) {
    const startValue = pointerToValue(event);
    setDragRange({ start: startValue, end: startValue });
    event.currentTarget.setPointerCapture?.(event.pointerId);
  }

  function handlePointerMove(event) {
    if (!dragRange) return;
    setDragRange((current) => current ? { ...current, end: pointerToValue(event) } : current);
  }

  function finishDrag(event) {
    if (!dragRange) return;
    const endValue = pointerToValue(event);
    const startValue = dragRange.start;
    const delta = Math.abs(endValue - startValue);
    setDragRange(null);
    if (delta < xSpan * 0.01) {
      setDraftRange(null);
      onRangeChange(null);
      return;
    }
    commitRange(startValue, endValue);
  }

  const polyline = data.map((point) => `${xScale(Number(point.value || 0)).toFixed(1)},${yScale(Number(point.count || 0)).toFixed(1)}`).join(" ");
  const tickCount = Math.min(6, data.length);
  const xTicks = Array.from({ length: tickCount }, (_, index) => {
    if (data.length === 1) return data[0].value;
    const ratio = index / Math.max(tickCount - 1, 1);
    return Math.round(minX + ratio * xSpan);
  });
  const yTicks = Array.from({ length: 5 }, (_, index) => Math.round((maxY * (4 - index)) / 4));
  const effectiveRange = dragRange ? [Math.min(dragRange.start, dragRange.end), Math.max(dragRange.start, dragRange.end)] : draftRange;
  const [selectedMin, selectedMax] = effectiveRange || [];
  const selectionX1 = effectiveRange ? xScale(selectedMin) : null;
  const selectionX2 = effectiveRange ? xScale(selectedMax) : null;

  return (
    <div className="mini-chart-shell">
      <svg ref={svgRef} viewBox={`0 0 ${width} ${height}`} className="mini-chart" role="img" aria-label="Settlement distribution chart">
        <rect x="0" y="0" width={width} height={height} fill="#ffffff" />
        {yTicks.map((tick) => {
          const y = yScale(tick);
          return (
            <g key={`y-${tick}`}>
              <line x1={padding.left} y1={y} x2={width - padding.right} y2={y} stroke="#e5edf8" strokeWidth="1" />
              <text x={padding.left - 12} y={y + 4} textAnchor="end" className="chart-axis-label">{tick}</text>
            </g>
          );
        })}
        {xTicks.map((tick) => {
          const x = xScale(tick);
          return (
            <g key={`x-${tick}`}>
              <line x1={x} y1={padding.top} x2={x} y2={height - padding.bottom} stroke="#f1f5f9" strokeWidth="1" />
              <text x={x} y={height - padding.bottom + 22} textAnchor="middle" className="chart-axis-label">{tick}</text>
            </g>
          );
        })}
        {effectiveRange ? <rect x={Math.min(selectionX1, selectionX2)} y={padding.top} width={Math.abs(selectionX2 - selectionX1)} height={plotHeight} fill="rgba(99, 102, 241, 0.12)" stroke="#6366f1" strokeDasharray="6 4" /> : null}
        <line x1={padding.left} y1={height - padding.bottom} x2={width - padding.right} y2={height - padding.bottom} stroke="#94a3b8" strokeWidth="1.5" />
        <line x1={padding.left} y1={padding.top} x2={padding.left} y2={height - padding.bottom} stroke="#94a3b8" strokeWidth="1.5" />
        <polyline fill="none" stroke="#6366f1" strokeWidth="3" points={polyline} />
        {data.map((point) => {
          const x = xScale(Number(point.value || 0));
          const y = yScale(Number(point.count || 0));
          const inSelection = !effectiveRange || (Number(point.value || 0) >= selectedMin && Number(point.value || 0) <= selectedMax);
          return <circle key={`${point.value}-${point.count}`} cx={x} cy={y} r="4.5" fill={inSelection ? "#6366f1" : "#c7d2fe"}><title>{`${point.value}: ${point.count}`}</title></circle>;
        })}
        <rect
          x={padding.left}
          y={padding.top}
          width={plotWidth}
          height={plotHeight}
          fill="transparent"
          className="chart-drag-zone"
          onPointerDown={handlePointerDown}
          onPointerMove={handlePointerMove}
          onPointerUp={finishDrag}
          onPointerLeave={(event) => {
            if (dragRange && (event.buttons & 1) === 0) finishDrag(event);
          }}
        />
      </svg>
      <div className="chart-help">Drag across the chart to select a settlement range. Click and release without dragging to clear selection.</div>
      <div className="chart-range-row">
        <label className="field-block"><span>Selection Min</span><input type="number" value={draftRange ? draftRange[0] : ""} placeholder={String(minX)} onChange={(e) => setDraftRange([Number(e.target.value || minX), draftRange ? draftRange[1] : maxX])} onBlur={() => draftRange && commitRange(draftRange[0], draftRange[1])} /></label>
        <label className="field-block"><span>Selection Max</span><input type="number" value={draftRange ? draftRange[1] : ""} placeholder={String(maxX)} onChange={(e) => setDraftRange([draftRange ? draftRange[0] : minX, Number(e.target.value || maxX)])} onBlur={() => draftRange && commitRange(draftRange[0], draftRange[1])} /></label>
        <button className="ghost-button" onClick={() => commitRange(minX, maxX)}>Select All</button>
        <button className="ghost-button" onClick={() => { setDraftRange(null); setDragRange(null); onRangeChange(null); }}>Clear</button>
      </div>
    </div>
  );
}

function DonutChart({ data, totalLabel = "Rows", palette, centerLabel = "Total" }) {
  const [activeIndex, setActiveIndex] = useState(null);
  const chartData = data.map((item, index) => ({ ...item, color: palette[index] || palette[palette.length - 1] }));
  const total = chartData.reduce((sum, item) => sum + Number(item.value || 0), 0);
  const hovered = activeIndex != null ? chartData[activeIndex] : null;
  const radius = 54;
  const circumference = 2 * Math.PI * radius;
  let offset = 0;

  if (!total) {
    return <div className="empty-state">No rows available for this breakdown.</div>;
  }

  return (
    <div className="donut-layout">
      <div className="donut-visual">
        <svg viewBox="0 0 160 160" className="donut-chart" role="img" aria-label={centerLabel}>
          <circle cx="80" cy="80" r={radius} fill="none" stroke="#e5edf8" strokeWidth="20" />
          {chartData.map((item, index) => {
            const segment = (Number(item.value || 0) / total) * circumference;
            const node = (
              <circle
                key={item.label}
                cx="80"
                cy="80"
                r={radius}
                fill="none"
                stroke={item.color}
                strokeWidth={activeIndex === index ? 24 : 20}
                strokeLinecap="round"
                strokeDasharray={`${segment} ${Math.max(circumference - segment, 0)}`}
                strokeDashoffset={-offset}
                transform="rotate(-90 80 80)"
                onMouseEnter={() => setActiveIndex(index)}
                onMouseLeave={() => setActiveIndex(null)}
              />
            );
            offset += segment;
            return node;
          })}
        </svg>
        <div className="donut-center">
          <div className="donut-center-label">{hovered ? hovered.label : centerLabel}</div>
          <div className="donut-center-value">{hovered ? hovered.value : total}</div>
          <div className="donut-center-subtle">{hovered ? `${Math.round((hovered.value / total) * 100)}%` : totalLabel}</div>
        </div>
      </div>
      <div className="donut-legend">{chartData.map((item, index) => <div className="donut-legend-item" key={item.label} onMouseEnter={() => setActiveIndex(index)} onMouseLeave={() => setActiveIndex(null)}><span className="donut-swatch" style={{ background: item.color }}></span><div className="donut-legend-copy"><div className="donut-legend-label">{item.label}</div><div className="donut-legend-meta">{item.value} rows</div></div></div>)}</div>
    </div>
  );
}
function App() {
  const [page, setPage] = useState("Dashboard");
  const [dataset, setDataset] = useState(null);
  const [dashboard, setDashboard] = useState(null);
  const [filters, setFilters] = useState(defaultFilters);
  const [bulkEditState, setBulkEditState] = useState({ mode: "Add", value: "", capMode: "Min", capValue: "" });
  const [offerConfig, setOfferConfig] = useState({
    yPct: "15",
    xPct: "20",
    cap: "500",
    thresholdMode: PRICE_THRESHOLD_MODE,
    thresholdSets: {
      [PRICE_THRESHOLD_MODE]: { ...THRESHOLD_DEFAULTS },
      [SETTLEMENT_THRESHOLD_MODE]: { ...SETTLEMENT_THRESHOLD_DEFAULTS },
    },
  });
  const [sizeOverride, setSizeOverrideState] = useState({ sku: "", size: SIZE_VALUES[0] || "32" });
  const [listingView, setListingView] = useState("overview");
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState("");
  const [notice, setNotice] = useState(null);
  const [selectedIds, setSelectedIds] = useState([]);
  const [confirmAction, setConfirmAction] = useState(null);
  const [exportOpen, setExportOpen] = useState(false);
  const [exportOptions, setExportOptions] = useState({ scope: "all", format: "xlsx", fileName: "WEB_EXCEL_OUTPUT", sheetName: "Web Excel", columns: [] });
  const [sort, setSort] = useState({ column: "", direction: "asc" });
  const [mobileNavOpen, setMobileNavOpen] = useState(false);
  const fileInputRef = useRef(null);
  const selectRangeMatchesRef = useRef(false);

  useEffect(() => {
    if (!dataset) {
      setDashboard(null);
      return;
    }
    try {
      setDashboard(createSnapshot(dataset, filters));
    } catch (err) {
      setError(err.message);
    }
  }, [dataset, filters]);

  useEffect(() => {
    if (!dataset) {
      setSelectedIds([]);
      selectRangeMatchesRef.current = false;
      return;
    }

    if (selectRangeMatchesRef.current && filters.selectedRange) {
      const matchingIds = createSnapshot(dataset, filters).matchingRowIds || [];
      setSelectedIds(matchingIds);
      setNotice({
        type: "success",
        message: `${matchingIds.length.toLocaleString()} rows in the selected price range are ready in Review & Edit.`,
      });
    } else {
      setSelectedIds([]);
    }
    selectRangeMatchesRef.current = false;
  }, [dataset, filters]);

  useEffect(() => {
    setOfferConfig((current) => ({
      ...current,
      thresholdSets: {
        [PRICE_THRESHOLD_MODE]: buildThresholdMap(THRESHOLD_DEFAULTS, current.thresholdSets?.[PRICE_THRESHOLD_MODE]),
        [SETTLEMENT_THRESHOLD_MODE]: buildThresholdMap(SETTLEMENT_THRESHOLD_DEFAULTS, current.thresholdSets?.[SETTLEMENT_THRESHOLD_MODE]),
      },
    }));
  }, []);

  const cards = useMemo(() => metricCards(dashboard?.metrics), [dashboard]);
  const filterOptions = dashboard?.filterOptions || { listing: [], jeans: [], size: [], status: [], settlementMax: 0 };
  const rows = dashboard?.rows || [];
  const columns = dashboard?.columns || [];
  const displayedColumns = useMemo(() => visibleListingColumns(columns, listingView), [columns, listingView]);
  const selectedSet = useMemo(() => new Set(selectedIds), [selectedIds]);
  const sortedRows = useMemo(() => {
    if (!sort.column) return rows;
    return [...rows].sort((a, b) => {
      const left = a[sort.column];
      const right = b[sort.column];
      const numeric = Number(left) - Number(right);
      const comparison = Number.isFinite(numeric) && left !== "" && right !== "" ? numeric : String(left ?? "").localeCompare(String(right ?? ""), undefined, { numeric: true });
      return sort.direction === "asc" ? comparison : -comparison;
    });
  }, [rows, sort]);
  const chartData = dashboard?.chart || [];
  const activeThresholdMode = offerConfig.thresholdMode || PRICE_THRESHOLD_MODE;
  const activeThresholdDefaults = thresholdDefaultsForMode(activeThresholdMode);
  const activeThresholdLabels = thresholdLabelsForMode(activeThresholdMode);
  const activeThresholds = offerConfig.thresholdSets?.[activeThresholdMode] || activeThresholdDefaults;
  const activeMode = modeMeta(dashboard?.mode, dashboard?.selectedValueColumn, activeThresholdMode);

  async function handleUpload(event) {
    const file = event.target.files?.[0];
    if (!file) return;
    try {
      setLoading(true);
      setError("");
      const nextDataset = await loadWorkbook(file);
      const nextFilters = defaultFiltersForDataset(nextDataset);
      setDataset(nextDataset);
      setFilters(nextFilters);
      setSelectedIds([]);
      setExportOptions((current) => ({ ...current, format: nextDataset.fileType === "csv" ? "csv" : "xlsx", columns: [] }));
      setOfferConfig((current) => ({
        ...current,
        thresholdMode: nextDataset.mode === "settlementRecommendations" ? SETTLEMENT_THRESHOLD_MODE : PRICE_THRESHOLD_MODE,
        thresholdSets: {
          [PRICE_THRESHOLD_MODE]: buildThresholdMap(THRESHOLD_DEFAULTS, current.thresholdSets?.[PRICE_THRESHOLD_MODE]),
          [SETTLEMENT_THRESHOLD_MODE]: buildThresholdMap(SETTLEMENT_THRESHOLD_DEFAULTS, current.thresholdSets?.[SETTLEMENT_THRESHOLD_MODE]),
        },
      }));
      setPage("Dashboard");
      setNotice({ type: "success", message: `${file.name} loaded successfully (${nextDataset.rows.length.toLocaleString()} rows).` });
    } catch (err) {
      setError(err.message);
    } finally {
      setLoading(false);
    }
  }

  function runAction(work, successMessage = "Changes applied successfully.") {
    try {
      setLoading(true);
      setError("");
      const nextDataset = work();
      setDataset(nextDataset);
      setNotice({ type: "success", message: successMessage });
    } catch (err) {
      setError(err.message);
    } finally {
      setLoading(false);
    }
  }

  function handleExport() {
    if (!dataset) return;
    setExportOpen(true);
  }

  function completeExport() {
    try {
      const exported = exportDataset(dataset, { ...exportOptions, filters, selectedIds, columns: exportOptions.columns.length ? exportOptions.columns : null });
      const url = URL.createObjectURL(exported.blob);
      const link = document.createElement("a");
      link.href = url;
      link.download = exported.fileName;
      link.click();
      URL.revokeObjectURL(url);
      setExportOpen(false);
      setNotice({ type: "success", message: `${exported.fileName} exported successfully.` });
    } catch (err) {
      setError(err.message);
    }
  }

  function requestAction({ title, message, detail, work, successMessage }) {
    setConfirmAction({ title, message, detail, work, successMessage });
  }

  function confirmRequestedAction() {
    if (!confirmAction) return;
    const action = confirmAction;
    setConfirmAction(null);
    runAction(action.work, action.successMessage);
  }

  function resetFilters() {
    if (!dataset) return;
    setFilters(defaultFiltersForDataset(dataset));
    setSelectedIds([]);
    setNotice({ type: "info", message: "Filters and row selection cleared." });
  }

  function toggleRow(rowId) {
    setSelectedIds((current) => current.includes(rowId) ? current.filter((id) => id !== rowId) : [...current, rowId]);
  }

  function toggleVisibleRows() {
    const visibleIds = rows.map((row) => row.__orig_index);
    const allSelected = visibleIds.length > 0 && visibleIds.every((id) => selectedSet.has(id));
    setSelectedIds((current) => allSelected ? current.filter((id) => !visibleIds.includes(id)) : [...new Set([...current, ...visibleIds])]);
  }

  function toggleSort(column) {
    setSort((current) => current.column === column ? { column, direction: current.direction === "asc" ? "desc" : "asc" } : { column, direction: "asc" });
  }

  function handleModeChange(mode) {
    if (!dataset) return;
    runAction(() => {
      const nextDataset = setMode(dataset, mode);
      setFilters(defaultFiltersForDataset(nextDataset));
      return nextDataset;
    });
  }

  function handleRangeChange(nextRange) {
    selectRangeMatchesRef.current = Boolean(nextRange);
    setFilters((current) => ({ ...current, selectedRange: nextRange ? [Math.min(nextRange[0], nextRange[1]), Math.max(nextRange[0], nextRange[1])] : null }));
  }

  function handleValueColumnChange(valueColumn) {
    if (!dataset) return;
    runAction(() => {
      const nextDataset = setValueColumn(dataset, valueColumn);
      const nextDefaults = defaultFiltersForDataset(nextDataset);
      setFilters((current) => ({
        ...current,
        settlementMax: nextDefaults.settlementMax,
        selectedRange: null,
      }));
      return nextDataset;
    });
  }

  function handleThresholdModeChange(nextMode) {
    setOfferConfig((current) => ({
      ...current,
      thresholdMode: nextMode,
      thresholdSets: {
        [PRICE_THRESHOLD_MODE]: buildThresholdMap(THRESHOLD_DEFAULTS, current.thresholdSets?.[PRICE_THRESHOLD_MODE]),
        [SETTLEMENT_THRESHOLD_MODE]: buildThresholdMap(SETTLEMENT_THRESHOLD_DEFAULTS, current.thresholdSets?.[SETTLEMENT_THRESHOLD_MODE]),
      },
    }));
  }

  function renderFilterPanel() {
    return (
      <section className="panel filter-panel">
        <div className="panel-head compact-head"><div><div className="panel-title">Find and filter rows</div><div className="panel-subtitle">Filters change the dashboard and visible table. Select rows separately before applying changes.</div></div><button className="ghost-button" onClick={resetFilters}>Reset filters</button></div>
        {dashboard?.availableValueColumns?.length ? <div className="toggle-row" style={{ marginBottom: 14 }}>{dashboard.availableValueColumns.map((column) => <button key={column} className={dashboard.selectedValueColumn === column ? "toggle-button toggle-button-active" : "toggle-button"} onClick={() => handleValueColumnChange(column)}>{column}</button>)}</div> : null}
        <div className="filter-grid">
          <label className="field-block field-span-2"><LabelWithHelp label="Search rows" help="Searches the visible workbook rows by SKU and title at the same time." /><div className="search-field"><input value={filters.search} onChange={(e) => setFilters((c) => ({ ...c, search: e.target.value }))} placeholder="Try SKU, title, or keyword" /></div></label>
          {[ ["Listing type", "listing", filterOptions.listing, "Filters owner vs latched listing rows."], ["Jeans family", "jeans", filterOptions.jeans, "Filters the classified jeans bucket used by offer mode thresholds."], ["Detected size", "size", filterOptions.size, "Uses SKU-based size detection plus any saved overrides."], ["Listing status", "status", filterOptions.status, "Shows only ACTIVE or INACTIVE rows when needed."] ].map(([label, key, options, help]) => (
            <label className="field-block" key={key}><LabelWithHelp label={label} help={help} /><select value={filters[key]} onChange={(e) => setFilters((c) => ({ ...c, [key]: e.target.value }))}><option value="All">All</option>{options.map((option) => <option key={option} value={option}>{option}</option>)}</select></label>
          ))}
          <label className="field-block field-span-2"><LabelWithHelp label={`${activeMode.valueLabel} ceiling`} help={`Only rows with ${activeMode.valueShort} at or below this value stay visible.`} /><div className="slider-block"><input type="range" min="0" max={filterOptions.settlementMax || 0} value={filters.settlementMax || 0} onChange={(e) => setFilters((c) => ({ ...c, settlementMax: Number(e.target.value) }))} /><div className="mono-value">{Number(filters.settlementMax || 0).toLocaleString()}</div></div></label>
        </div>
        <div className="filter-summary" aria-live="polite"><strong>{dashboard?.exportCount || 0}</strong> rows match the current filters and range · <strong>{selectedIds.length}</strong> explicitly selected</div>
      </section>
    );
  }

  function renderDashboardPage() {
    return (
      <>
        {renderFilterPanel()}
        <section className="metrics-grid">{cards.map((card) => <div className="metric-card" key={card.key}><div className="metric-label">{card.label}</div><div className="metric-value">{card.value}</div><div className="metric-detail">{card.detail}</div></div>)}</section>
        <section className="content-grid">
          <div className="panel chart-panel"><div className="panel-head"><div><div className="panel-title">{activeMode.valueLabel} Distribution <HelpTip text={`Drag across the graph to select a ${activeMode.valueShort} band. Every matching row is selected automatically in Review & Edit.`} /></div><div className="panel-subtitle">{`Interactive view of the currently visible ${activeMode.valueShort} values.`}</div></div><button className="ghost-button" disabled={!dataset} onClick={() => handleRangeChange(null)}>Reset Selection</button></div><div className="chart-wrap"><MiniLineChart data={chartData} selectedRange={filters.selectedRange} onRangeChange={handleRangeChange} /></div></div>
          <div className="panel"><div className="panel-head"><div><div className="panel-title">{activeMode.valueLabel} Summary</div><div className="panel-subtitle">{`Counts, ${activeMode.valueShort} selection stats, and active inactive matrix.`}</div></div></div><div className="summary-stack"><div className="summary-box">{dashboard?.summary?.rowCountText || "Loaded: 0 | Visible: 0 | Export: 0"}</div><div className="summary-box">Account: {dashboard?.accountName || "Unknown"}</div><div className="summary-box">Mode value source: {activeMode.valueLabel}</div><div className="summary-box">{`${activeMode.valueLabel} selection range: ${dashboard?.summary?.selectionRange ? `${dashboard.summary.selectionRange[0].toFixed(2)} to ${dashboard.summary.selectionRange[1].toFixed(2)}` : "None"}`}</div><div className="summary-box">{dashboard?.summary?.selectionStats?.label || "No Selection"}</div><table className="summary-table"><thead><tr><th>Status</th><th>Count</th><th>Sizes</th></tr></thead><tbody>{(dashboard?.summary?.statusMatrix || []).map((row) => <tr key={row.status}><td>{row.status}</td><td>{row.count}</td><td>{row.sizes?.length ? row.sizes.join(", ") : "-"}</td></tr>)}</tbody></table></div></div>
        </section>
        <section className="pie-grid">
          <div className="panel"><div className="panel-head"><div><div className="panel-title">Listing Type Ratio <HelpTip text="Shows how the current working set is split between Owner and Latched listings." /></div><div className="panel-subtitle">Owner versus latched mix in the filtered workset.</div></div></div><DonutChart data={dashboard?.summary?.listingTypeRatio || []} centerLabel="Listings" totalLabel="visible rows" palette={["#1e293b", "#6366f1"]} /></div>
          <div className="panel"><div className="panel-head"><div><div className="panel-title">Size-Color Status Mix <HelpTip text="Each size and jeans-color group is classified as Active, Inactive, or Blend depending on whether that group's rows are all active, all inactive, or mixed." /></div><div className="panel-subtitle">Status quality across size and color combinations.</div></div></div><DonutChart data={dashboard?.summary?.sizeColorStatusRatio || []} centerLabel="Groups" totalLabel="classified rows" palette={["#10b981", "#f43f5e", "#f59e0b"]} /><div className="group-breakdown">{(dashboard?.summary?.sizeColorBreakdown || []).slice(0, 8).map((entry) => <div className="group-breakdown-item" key={entry.label}><span className={`group-state-badge group-state-${String(entry.state || "").toLowerCase()}`}>{entry.state}</span><div className="group-breakdown-copy"><div className="group-breakdown-label">{entry.label}</div><div className="group-breakdown-meta">{entry.count} rows</div></div></div>)}</div></div>
        </section>
      </>
    );
  }

  function renderListingsPage() {
    const visibleIds = rows.map((row) => row.__orig_index);
    const allVisibleSelected = visibleIds.length > 0 && visibleIds.every((id) => selectedSet.has(id));
    const selectedPreview = rows.filter((row) => selectedSet.has(row.__orig_index)).slice(0, 3);

    const selectionDetail = (
      <div className="confirmation-preview">
        <div className="preview-stat"><strong>{selectedIds.length}</strong><span>rows will be changed</span></div>
        {selectedPreview.map((row) => <div className="preview-row" key={row.__orig_index}>{String(row[displayedColumns[0]] || row[displayedColumns[1]] || `Row ${row.__orig_index + 1}`)}</div>)}
        {selectedIds.length > 3 ? <div className="subtle">and {selectedIds.length - 3} more selected rows</div> : null}
      </div>
    );

    const requireSelection = (config) => {
      if (!selectedIds.length) {
        setNotice({ type: "warning", message: "Select one or more rows in the table before applying a change." });
        return;
      }
      requestAction({ ...config, detail: selectionDetail });
    };

    return (
      <>
        {renderFilterPanel()}
        <section className="selection-bar" aria-live="polite">
          <div><strong>{selectedIds.length.toLocaleString()} selected</strong><span>Only checked rows can be changed.</span></div>
          <div className="selection-actions">
            <button className="ghost-button" onClick={toggleVisibleRows}>{allVisibleSelected ? "Clear visible selection" : `Select ${rows.length} visible rows`}</button>
            <button className="text-button" disabled={!selectedIds.length} onClick={() => setSelectedIds([])}>Clear all</button>
          </div>
        </section>

        <section className="edit-layout">
          {dashboard?.mode !== "settlementRecommendations" ? (
            <div className="panel">
              <div className="panel-head"><div><div className="panel-title">Change selected values</div><div className="panel-subtitle">Preview and confirm every bulk edit before it is applied.</div></div></div>
              <div className="stack-grid">
                <div className="mini-form-row">
                  <label className="field-block"><span>Edit method</span><select value={bulkEditState.mode} onChange={(e) => setBulkEditState((c) => ({ ...c, mode: e.target.value }))}><option>Add</option><option>Multiply</option><option>Replace</option></select></label>
                  <label className="field-block"><span>Value</span><input type="number" value={bulkEditState.value} onChange={(e) => setBulkEditState((c) => ({ ...c, value: e.target.value }))} placeholder="Example: 25 or 1.05" /></label>
                </div>
                <div className="mini-form-row">
                  <label className="field-block"><span>Optional limit type</span><select value={bulkEditState.capMode} onChange={(e) => setBulkEditState((c) => ({ ...c, capMode: e.target.value }))}><option value="Min">Maximum value</option><option value="Max">Minimum value</option></select></label>
                  <label className="field-block"><span>Optional limit</span><input type="number" value={bulkEditState.capValue} onChange={(e) => setBulkEditState((c) => ({ ...c, capValue: e.target.value }))} placeholder="Leave empty for no limit" /></label>
                </div>
                <button className="primary-button" disabled={!selectedIds.length || bulkEditState.value === ""} onClick={() => requireSelection({
                  title: "Confirm bulk value change",
                  message: `${bulkEditState.mode} ${bulkEditState.value} using ${activeMode.valueLabel}?`,
                  work: () => bulkEdit(dataset, filters, bulkEditState.mode, Number(bulkEditState.value), bulkEditState.capValue ? bulkEditState.capMode : null, bulkEditState.capValue ? Number(bulkEditState.capValue) : null, selectedIds),
                  successMessage: `${selectedIds.length} selected rows updated.`,
                })}>Preview value change</button>
                <div className="status-button-row">
                  <button className="status-chip status-chip-active" disabled={!selectedIds.length} onClick={() => requireSelection({ title: "Mark rows active?", message: "Stock values may also be updated by the workbook rules.", work: () => setStatus(dataset, filters, "ACTIVE", selectedIds), successMessage: `${selectedIds.length} rows marked active.` })}>Mark selected active</button>
                  <button className="status-chip status-chip-inactive" disabled={!selectedIds.length} onClick={() => requireSelection({ title: "Mark rows inactive?", message: "Stock values may also be updated by the workbook rules.", work: () => setStatus(dataset, filters, "INACTIVE", selectedIds), successMessage: `${selectedIds.length} rows marked inactive.` })}>Mark selected inactive</button>
                </div>
              </div>
            </div>
          ) : null}

          <div className="panel">
            <div className="panel-head"><div><div className="panel-title">{dashboard?.mode === "settlementRecommendations" ? "Recommendation rules" : "Offer calculation"}</div><div className="panel-subtitle">Advanced controls for the current supported file type.</div></div></div>
            <div className="stack-grid">
              {dashboard?.mode === "settlementRecommendations" ? (
                <div className="toggle-row"><button className={activeThresholdMode === SETTLEMENT_THRESHOLD_MODE ? "toggle-button toggle-button-active" : "toggle-button"} onClick={() => handleThresholdModeChange(SETTLEMENT_THRESHOLD_MODE)}>Use settlement thresholds</button><button className={activeThresholdMode === PRICE_THRESHOLD_MODE ? "toggle-button toggle-button-active" : "toggle-button"} onClick={() => handleThresholdModeChange(PRICE_THRESHOLD_MODE)}>Use price thresholds</button></div>
              ) : (
                <div className="mini-form-row three-inputs">
                  <label className="field-block"><span>Base percentage</span><input type="number" value={offerConfig.yPct} onChange={(e) => setOfferConfig((c) => ({ ...c, yPct: e.target.value }))} /></label>
                  <label className="field-block"><span>Discount percentage</span><input type="number" value={offerConfig.xPct} onChange={(e) => setOfferConfig((c) => ({ ...c, xPct: e.target.value }))} /></label>
                  <label className="field-block"><span>Maximum discount</span><input type="number" value={offerConfig.cap} onChange={(e) => setOfferConfig((c) => ({ ...c, cap: e.target.value }))} /></label>
                </div>
              )}
              <details className="advanced-details"><summary>Review or change thresholds</summary><div className="threshold-grid">{THRESHOLD_KEYS.map((key) => <label className="threshold-card" key={key}><span>{activeThresholdLabels[key]}</span><input type="number" value={activeThresholds[key]} onChange={(e) => setOfferConfig((c) => ({ ...c, thresholdSets: { ...c.thresholdSets, [activeThresholdMode]: { ...activeThresholds, [key]: e.target.value } } }))} /></label>)}</div></details>
              <div className="status-button-row">
                {dashboard?.mode === "offer" ? <button className="primary-button" disabled={!selectedIds.length} onClick={() => requireSelection({ title: "Calculate discounts?", message: `Base ${offerConfig.yPct}%, discount ${offerConfig.xPct}%, capped at ${offerConfig.cap}.`, work: () => computeDiscount(dataset, filters, Number(offerConfig.yPct || 0), Number(offerConfig.xPct || 0), Number(offerConfig.cap || 0), selectedIds), successMessage: `Discount calculated for ${selectedIds.length} rows.` })}>Preview discount calculation</button> : null}
                <button className="ghost-button" disabled={!selectedIds.length || !["offer", "settlementRecommendations"].includes(dashboard?.mode || "")} onClick={() => requireSelection({ title: "Apply accept/reject rules?", message: `The selected ${activeThresholdMode} thresholds will be applied.`, work: () => applyDecision(dataset, filters, Object.fromEntries(Object.entries(activeThresholds).map(([key, value]) => [key, Number(value || 0)])), activeThresholdMode, selectedIds), successMessage: `Decision rules applied to ${selectedIds.length} rows.` })}>Preview decision rules</button>
              </div>
            </div>
          </div>

          {dashboard?.mode !== "settlementRecommendations" ? (
            <div className="panel">
              <div className="panel-head"><div><div className="panel-title">Correct a detected size</div><div className="panel-subtitle">Saved on this browser and reused for the same SKU.</div></div></div>
              <div className="stack-grid">
                <label className="field-block"><span>Exact SKU</span><input value={sizeOverride.sku} onChange={(e) => setSizeOverrideState((c) => ({ ...c, sku: e.target.value }))} placeholder="Paste an SKU" /></label>
                <label className="field-block"><span>Correct size</span><select value={sizeOverride.size} onChange={(e) => setSizeOverrideState((c) => ({ ...c, size: e.target.value }))}>{SIZE_VALUES.map((size) => <option key={size}>{size}</option>)}</select></label>
                <button className="primary-button" disabled={!sizeOverride.sku.trim()} onClick={() => runAction(() => saveSizeOverride(dataset, filters, sizeOverride.sku, sizeOverride.size), `Saved size ${sizeOverride.size} for ${sizeOverride.sku}.`)}>Save correction</button>
              </div>
            </div>
          ) : null}
        </section>

        <section className="panel">
          <div className="panel-head"><div><div className="panel-title">Review workbook rows</div><div className="panel-subtitle">Check rows to select them. Click a column heading to sort the current view.</div></div><div className="mono-value">{rows.length} of {dashboard?.exportCount || 0} matching rows shown</div></div>
          <div className="listing-toolbar"><div className="listing-view-tabs">{Object.entries(LISTING_COLUMN_VIEWS).map(([key, config]) => <button key={key} className={listingView === key ? "listing-view-tab listing-view-tab-active" : "listing-view-tab"} onClick={() => setListingView(key)}>{config.label}</button>)}</div><div className="listing-toolbar-note">{displayedColumns.length} columns visible</div></div>
          <div className="table-scroll" tabIndex="0" aria-label="Workbook rows">
            <table className="listing-table">
              <thead><tr>
                <th className="select-column"><input type="checkbox" checked={allVisibleSelected} onChange={toggleVisibleRows} aria-label="Select all visible rows" /></th>
                {displayedColumns.map((column, index) => <th key={column} className={index === 0 ? "sticky-first" : index === 1 ? "sticky-second" : ""}><button className="sort-button" onClick={() => toggleSort(column)}>{column}<span aria-hidden="true">{sort.column === column ? (sort.direction === "asc" ? " ↑" : " ↓") : ""}</span></button></th>)}
              </tr></thead>
              <tbody>{sortedRows.map((row) => {
                const status = String(row["Listing Status"] || "").toUpperCase();
                const selected = selectedSet.has(row.__orig_index);
                return <tr key={row.__orig_index} className={`${status === "ACTIVE" ? "row-active" : status === "INACTIVE" ? "row-inactive" : ""} ${selected ? "row-selected" : ""}`}>
                  <td className="select-column"><input type="checkbox" checked={selected} onChange={() => toggleRow(row.__orig_index)} aria-label={`Select row ${row.__orig_index + 1}`} /></td>
                  {displayedColumns.map((column, index) => {
                    const value = row[column];
                    const stickyClass = index === 0 ? "sticky-first-cell" : index === 1 ? "sticky-second-cell" : "";
                    if (column === "Listing Status") return <td key={column} className={stickyClass}><span className={`table-badge ${status === "ACTIVE" ? "table-badge-active" : status === "INACTIVE" ? "table-badge-inactive" : "table-badge-neutral"}`}>{status || "-"}</span></td>;
                    if (column === "Auto Flag" && value) return <td key={column} className={`flag-cell ${stickyClass}`.trim()}><span className="flag-inline">! {String(value)}</span></td>;
                    return <td key={column} className={stickyClass} title={value == null ? "" : String(value)}>{value == null || value === "" ? "-" : String(value)}</td>;
                  })}
                </tr>;
              })}</tbody>
            </table>
          </div>
        </section>
      </>
    );
  }

  function renderInsightsPage() {
    return (
      <section className="panel"><div className="panel-head"><div><div className="panel-title">Recent Changes</div><div className="panel-subtitle">Change log, file context, and workbook mode.</div></div></div><div className="change-log-grid">{(dashboard?.changeLog || []).length ? dashboard.changeLog.map((entry, index) => <article className="change-card" key={`${entry.action}-${index}`}><div className="change-title">{entry.action}</div><div className="change-meta">Rows changed: {entry.rows}</div><div><strong>Before Min/Max:</strong> {entry.before}</div><div><strong>After Min/Max:</strong> {entry.after}</div>{entry.extra ? <div className="change-copy subtle">{entry.extra}</div> : null}</article>) : <div className="empty-state">Load a file, then your local edits and workbook decisions will show here.</div>}<article className="change-card"><div className="change-title">Session Context</div><div className="change-meta">Mode: {dashboard?.mode || "normal"}</div><div><strong>File:</strong> {dashboard?.fileName || "No file loaded"}</div><div><strong>Account:</strong> {dashboard?.accountName || "Unknown"}</div><div className="change-copy subtle">Using {dataset ? "browser-side JS workbook logic" : "no dataset yet"} inspired by <code>RateInsight.jsx</code>.</div></article></div></section>
    );
  }

  const navigation = [
    { page: "Dashboard", label: "Overview", step: "1" },
    { page: "Listings", label: "Review & Edit", step: "2" },
    { page: "Insights", label: "Change History", step: "3" },
    { page: "Hisaab", label: "Hisaab", step: "H" },
  ];

  const navigate = (nextPage) => {
    setPage(nextPage);
    setMobileNavOpen(false);
  };

  return (
    <div className="app-shell">
      <aside className={`sidebar ${mobileNavOpen ? "sidebar-open" : ""}`}>
        <div className="sidebar-brand"><span className="brand-mark">WE</span><span>Web Excel</span></div>
        <nav className="sidebar-nav" aria-label="Main navigation">{navigation.map((item) => <button key={item.page} className={page === item.page ? "nav-item nav-item-active" : "nav-item"} onClick={() => navigate(item.page)}><span className="nav-step">{item.step}</span>{item.label}</button>)}</nav>
        <div className="sidebar-help"><strong>Simple workflow</strong><span>Load a supported workbook, review it, select rows, apply changes, and export.</span></div>
        <div className="sidebar-user"><div className="avatar">WE</div><div><div className="user-name">Local workspace</div><div className="user-role">Files stay in this browser</div></div></div>
      </aside>
      {mobileNavOpen ? <button className="nav-scrim" aria-label="Close navigation" onClick={() => setMobileNavOpen(false)} /> : null}

      <main className="main-shell">
        <header className="topbar">
          <button className="mobile-menu-button" onClick={() => setMobileNavOpen(true)} aria-label="Open navigation">Menu</button>
          <div className="topbar-title"><h1>{navigation.find((item) => item.page === page)?.label || "Web Excel"}</h1><p>{dataset ? dataset.fileName : "Load a supported workbook to begin"}</p></div>
          <div className="topbar-actions">
            <button className={dataset ? "ghost-button" : "primary-button load-button-thematic"} onClick={() => fileInputRef.current?.click()}><span className="load-button-mark">+</span>{dataset ? "Replace workbook" : "Load workbook"}</button>
            <button className="primary-button" disabled={!dataset} onClick={handleExport}>Export</button>
            <button className="ghost-button" disabled={!dataset?.history?.length} onClick={() => runAction(() => undoDataset(dataset), "Last change undone.")}>Undo <span className="button-count">{dataset?.history?.length || 0}</span></button>
            <input hidden ref={fileInputRef} type="file" accept=".xlsx,.xls,.csv" onChange={handleUpload} />
          </div>
        </header>

        {dataset ? <div className="workbook-status" role="status"><span><strong>{dataset.fileName}</strong></span><span>{dataset.rows.length.toLocaleString()} rows loaded</span><span>{dashboard?.exportCount || 0} matching</span><span>{selectedIds.length} selected</span><span>{dataset.history?.length || 0} undo steps</span></div> : null}

        <div className="page-scroll">
          {dataset && page !== "Hisaab" ? <nav className="workflow-strip" aria-label="Workbook workflow">
            <button className={page === "Dashboard" ? "workflow-step workflow-step-active" : "workflow-step"} onClick={() => navigate("Dashboard")}><span>1</span><div><strong>Understand</strong><small>Review totals and trends</small></div></button>
            <button className={page === "Listings" ? "workflow-step workflow-step-active" : "workflow-step"} onClick={() => navigate("Listings")}><span>2</span><div><strong>Select and change</strong><small>Choose rows before editing</small></div></button>
            <button className={page === "Insights" ? "workflow-step workflow-step-active" : "workflow-step"} onClick={() => navigate("Insights")}><span>3</span><div><strong>Check history</strong><small>Review completed operations</small></div></button>
            <button className="workflow-step" onClick={handleExport}><span>4</span><div><strong>Export</strong><small>Choose rows and columns</small></div></button>
          </nav> : null}

          {error ? <div className="error-banner" role="alert"><span>{error}</span><button className="text-button" onClick={() => setError("")}>Dismiss</button></div> : null}

          {!dataset && page !== "Hisaab" ? <section className="panel empty-upload-state"><div className="empty-upload-art"><div className="empty-upload-icon">WE</div><div className="empty-upload-lines"><span></span><span></span><span></span></div></div><div className="empty-upload-copy"><div className="eyebrow">Start here</div><h2>Open a supported workbook</h2><p>Your file is processed locally in the browser. Web Excel supports the existing listing, offer, order, and settlement recommendation formats.</p><div className="empty-upload-tags"><span>.xlsx</span><span>.xls</span><span>.csv</span></div><button className="primary-button empty-upload-button" onClick={() => fileInputRef.current?.click()}><span className="load-button-mark">+</span>Select workbook</button><p className="privacy-note">No file is uploaded to a server.</p></div></section> : null}
          {dataset && page === "Dashboard" ? renderDashboardPage() : null}
          {dataset && page === "Listings" ? renderListingsPage() : null}
          {dataset && page === "Insights" ? renderInsightsPage() : null}
          {page === "Hisaab" ? <HisaabPage /> : null}
        </div>
      </main>

      <Notice notice={notice} onClose={() => setNotice(null)} />

      {confirmAction ? <Modal title={confirmAction.title} onClose={() => setConfirmAction(null)} actions={<><button className="ghost-button" onClick={() => setConfirmAction(null)}>Cancel</button><button className="primary-button" onClick={confirmRequestedAction}>Confirm change</button></>}>
        <p>{confirmAction.message}</p>
        {confirmAction.detail}
        <div className="warning-box">You can undo this operation after it is applied.</div>
      </Modal> : null}

      {exportOpen ? <Modal title="Export workbook" size="wide" onClose={() => setExportOpen(false)} actions={<><button className="ghost-button" onClick={() => setExportOpen(false)}>Cancel</button><button className="primary-button" onClick={completeExport}>Download export</button></>}>
        <div className="export-grid">
          <fieldset className="field-group"><legend>Rows to export</legend>
            <label className="radio-card"><input type="radio" name="scope" checked={exportOptions.scope === "all"} onChange={() => setExportOptions((c) => ({ ...c, scope: "all" }))} /><span><strong>All rows</strong><small>{dataset?.rows.length || 0} rows</small></span></label>
            <label className="radio-card"><input type="radio" name="scope" checked={exportOptions.scope === "filtered"} onChange={() => setExportOptions((c) => ({ ...c, scope: "filtered" }))} /><span><strong>Matching rows</strong><small>{dashboard?.exportCount || 0} rows after filters and range</small></span></label>
            <label className="radio-card"><input type="radio" name="scope" disabled={!selectedIds.length} checked={exportOptions.scope === "selected"} onChange={() => setExportOptions((c) => ({ ...c, scope: "selected" }))} /><span><strong>Selected rows</strong><small>{selectedIds.length} checked rows</small></span></label>
          </fieldset>
          <div className="stack-grid">
            <div className="mini-form-row"><label className="field-block"><span>File name</span><input value={exportOptions.fileName} onChange={(e) => setExportOptions((c) => ({ ...c, fileName: e.target.value }))} /></label><label className="field-block"><span>Format</span><select value={exportOptions.format} onChange={(e) => setExportOptions((c) => ({ ...c, format: e.target.value }))}><option value="xlsx">Excel (.xlsx)</option><option value="csv">CSV (.csv)</option></select></label></div>
            {exportOptions.format === "xlsx" ? <label className="field-block"><span>Sheet name</span><input value={exportOptions.sheetName} onChange={(e) => setExportOptions((c) => ({ ...c, sheetName: e.target.value }))} maxLength="31" /></label> : null}
            <div className="field-block"><span>Columns</span><div className="column-picker-actions"><button className="text-button" onClick={() => setExportOptions((c) => ({ ...c, columns: [...columns] }))}>Select all</button><button className="text-button" onClick={() => setExportOptions((c) => ({ ...c, columns: [] }))}>Use default columns</button></div><div className="column-picker">{columns.map((column) => <label key={column}><input type="checkbox" checked={exportOptions.columns.includes(column)} onChange={() => setExportOptions((c) => ({ ...c, columns: c.columns.includes(column) ? c.columns.filter((item) => item !== column) : [...c.columns, column] }))} />{column}</label>)}</div><small className="subtle">{exportOptions.columns.length ? `${exportOptions.columns.length} custom columns selected` : "All default output columns will be included"}</small></div>
          </div>
        </div>
      </Modal> : null}
    </div>
  );

}

export default App;









