import "./styles.css";

import React, { useEffect, useMemo, useRef, useState } from "react";
import {
  THRESHOLD_KEYS,
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
  undoDataset,
} from "./rateInsight";

const defaultFilters = { search: "", listing: "All", jeans: "All", size: "All", status: "All", settlementMax: 0, selectedRange: null };
const DEFAULT_THRESHOLDS = {
  ICE: "469",
  BEIGE: "459",
  WHITE: "389",
  "BLACK-BAGGY": "429",
  "BLACK-PLAIN": "399",
  MIX: "469",
};
const THRESHOLD_LABELS = {
  ICE: "ICE final price floor",
  BEIGE: "BEIGE final price floor",
  WHITE: "WHITE final price floor",
  "BLACK-BAGGY": "BLACK BAGGY floor",
  "BLACK-PLAIN": "BLACK PLAIN floor",
  MIX: "MIX floor",
};

const LISTING_COLUMN_VIEWS = {
  overview: {
    label: "Overview",
    columns: ["Product Title", "FSN", "Seller SKU Id", "SKU ID", "Bank Settlement", "Selling Price(Rs)", "Listing Status", "Size", "Jeans Type", "Listing Type", "Auto Flag"],
  },
  pricing: {
    label: "Pricing",
    columns: ["Product Title", "FSN", "Seller SKU Id", "SKU ID", "Bank Settlement", "Selling Price(Rs)", "Discount", "Final Price", "Decision", "MRP (?)", "Your Selling Price (Rs)"],
  },
  status: {
    label: "Status & Flags",
    columns: ["Product Title", "FSN", "Seller SKU Id", "SKU ID", "Listing Status", "Your Stock Count", "Size", "Jeans Type", "Listing Type", "Auto Flag"],
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

function metricCards(metrics) {
  return [
    { key: "loaded", label: "Total Loaded", value: metrics?.loaded ?? 0, detail: "Rows in file" },
    { key: "visible", label: "Visible Listings", value: metrics?.visible ?? 0, detail: "Rows after filters" },
    { key: "export", label: "Exported Rows", value: metrics?.export ?? 0, detail: "Rows in selection" },
    { key: "active", label: "Active Listings", value: metrics?.active ?? 0, detail: `${metrics?.activePct ?? 0}% of visible` },
    { key: "inactive", label: "Inactive Listings", value: metrics?.inactive ?? 0, detail: `${metrics?.inactivePct ?? 0}% of visible` },
  ];
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
  const [offerConfig, setOfferConfig] = useState({ yPct: "15", xPct: "20", cap: "500", thresholds: { ...DEFAULT_THRESHOLDS } });
  const [sizeOverride, setSizeOverrideState] = useState({ sku: "", size: "32" });
  const [listingView, setListingView] = useState("overview");
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState("");
  const fileInputRef = useRef(null);

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

  const cards = useMemo(() => metricCards(dashboard?.metrics), [dashboard]);
  const filterOptions = dashboard?.filterOptions || { listing: [], jeans: [], size: [], status: [], settlementMax: 0 };
  const rows = dashboard?.rows || [];
  const columns = dashboard?.columns || [];
  const displayedColumns = useMemo(() => visibleListingColumns(columns, listingView), [columns, listingView]);
  const chartData = dashboard?.chart || [];

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
      setPage("Dashboard");
    } catch (err) {
      setError(err.message);
    } finally {
      setLoading(false);
    }
  }

  function runAction(work) {
    try {
      setLoading(true);
      setError("");
      const nextDataset = work();
      setDataset(nextDataset);
    } catch (err) {
      setError(err.message);
    } finally {
      setLoading(false);
    }
  }

  function handleExport() {
    if (!dataset) return;
    const exported = exportDataset(dataset);
    const url = URL.createObjectURL(exported.blob);
    const link = document.createElement("a");
    link.href = url;
    link.download = exported.fileName;
    link.click();
    URL.revokeObjectURL(url);
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
    setFilters((current) => ({ ...current, selectedRange: nextRange ? [Math.min(nextRange[0], nextRange[1]), Math.max(nextRange[0], nextRange[1])] : null }));
  }

  function renderFilterPanel() {
    return (
      <section className="panel filter-panel">
        <div className="filter-grid">
          <label className="field-block field-span-2"><LabelWithHelp label="Search rows" help="Searches the visible workbook rows by SKU and title at the same time." /><div className="search-field"><input value={filters.search} onChange={(e) => setFilters((c) => ({ ...c, search: e.target.value }))} placeholder="Try SKU, title, or keyword" /></div></label>
          {[ ["Listing type", "listing", filterOptions.listing, "Filters owner vs latched listing rows."], ["Jeans family", "jeans", filterOptions.jeans, "Filters the classified jeans bucket used by offer mode thresholds."], ["Detected size", "size", filterOptions.size, "Uses SKU-based size detection plus any saved overrides."], ["Listing status", "status", filterOptions.status, "Shows only ACTIVE or INACTIVE rows when needed."] ].map(([label, key, options, help]) => (
            <label className="field-block" key={key}><LabelWithHelp label={label} help={help} /><select value={filters[key]} onChange={(e) => setFilters((c) => ({ ...c, [key]: e.target.value }))}><option value="All">All</option>{options.map((option) => <option key={option} value={option}>{option}</option>)}</select></label>
          ))}
          <label className="field-block field-span-2"><LabelWithHelp label="Settlement ceiling" help="Only rows with settlement at or below this value stay visible." /><div className="slider-block"><input type="range" min="0" max={filterOptions.settlementMax || 0} value={filters.settlementMax || 0} onChange={(e) => setFilters((c) => ({ ...c, settlementMax: Number(e.target.value) }))} /><div className="mono-value">{Number(filters.settlementMax || 0).toLocaleString()}</div></div></label>
        </div>
      </section>
    );
  }

  function renderDashboardPage() {
    return (
      <>
        {renderFilterPanel()}
        <section className="metrics-grid">{cards.map((card) => <div className="metric-card" key={card.key}><div className="metric-label">{card.label}</div><div className="metric-value">{card.value}</div><div className="metric-detail">{card.detail}</div></div>)}</section>
        <section className="content-grid">
          <div className="panel chart-panel"><div className="panel-head"><div><div className="panel-title">Settlement Distribution <HelpTip text="Drag across the graph to select a settlement band. The dashboard summary and listing table update to that band." /></div><div className="panel-subtitle">Interactive view of the currently visible settlement values.</div></div><button className="ghost-button" disabled={!dataset} onClick={() => setFilters((c) => ({ ...c, selectedRange: null }))}>Reset Selection</button></div><div className="chart-wrap"><MiniLineChart data={chartData} selectedRange={filters.selectedRange} onRangeChange={handleRangeChange} /></div></div>
          <div className="panel"><div className="panel-head"><div><div className="panel-title">Settlement Summary</div><div className="panel-subtitle">Counts, selection stats, and active inactive matrix.</div></div></div><div className="summary-stack"><div className="summary-box">{dashboard?.summary?.rowCountText || "Loaded: 0 | Visible: 0 | Export: 0"}</div><div className="summary-box">Account: {dashboard?.accountName || "Unknown"}</div><div className="summary-box">Selection range: {dashboard?.summary?.selectionRange ? `${dashboard.summary.selectionRange[0].toFixed(2)} to ${dashboard.summary.selectionRange[1].toFixed(2)}` : "None"}</div><div className="summary-box">{dashboard?.summary?.selectionStats?.label || "No Selection"}</div><table className="summary-table"><thead><tr><th>Status</th><th>Count</th><th>Sizes</th></tr></thead><tbody>{(dashboard?.summary?.statusMatrix || []).map((row) => <tr key={row.status}><td>{row.status}</td><td>{row.count}</td><td>{row.sizes?.length ? row.sizes.join(", ") : "-"}</td></tr>)}</tbody></table></div></div>
        </section>
        <section className="pie-grid">
          <div className="panel"><div className="panel-head"><div><div className="panel-title">Listing Type Ratio <HelpTip text="Shows how the current working set is split between Owner and Latched listings." /></div><div className="panel-subtitle">Owner versus latched mix in the filtered workset.</div></div></div><DonutChart data={dashboard?.summary?.listingTypeRatio || []} centerLabel="Listings" totalLabel="visible rows" palette={["#1e293b", "#6366f1"]} /></div>
          <div className="panel"><div className="panel-head"><div><div className="panel-title">Size-Color Status Mix <HelpTip text="Each size and jeans-color group is classified as Active, Inactive, or Blend depending on whether that group's rows are all active, all inactive, or mixed." /></div><div className="panel-subtitle">Status quality across size and color combinations.</div></div></div><DonutChart data={dashboard?.summary?.sizeColorStatusRatio || []} centerLabel="Groups" totalLabel="classified rows" palette={["#10b981", "#f43f5e", "#f59e0b"]} /><div className="group-breakdown">{(dashboard?.summary?.sizeColorBreakdown || []).slice(0, 8).map((entry) => <div className="group-breakdown-item" key={entry.label}><span className={`group-state-badge group-state-${String(entry.state || "").toLowerCase()}`}>{entry.state}</span><div className="group-breakdown-copy"><div className="group-breakdown-label">{entry.label}</div><div className="group-breakdown-meta">{entry.count} rows</div></div></div>)}</div></div>
        </section>
      </>
    );
  }

  function renderListingsPage() {
    return (
      <>
        {renderFilterPanel()}
        <section className="three-up-grid">
          <div className="panel"><div className="panel-head"><div><div className="panel-title">Bulk Actions <HelpTip text="Apply a settlement rule to the currently filtered and selected rows." /></div><div className="panel-subtitle">Quick pricing edits without leaving the browser.</div></div></div><div className="stack-grid"><div className="mini-form-row"><label className="field-block"><LabelWithHelp label="Edit rule" help="Add increases by a fixed amount, Multiply scales the settlement, Replace sets one value for all selected rows." /><select value={bulkEditState.mode} onChange={(e) => setBulkEditState((c) => ({ ...c, mode: e.target.value }))}><option>Add</option><option>Multiply</option><option>Replace</option></select></label><label className="field-block"><LabelWithHelp label="Rule value" help="Examples: 25 for Add, 1.05 for Multiply, or a final settlement amount for Replace." /><input value={bulkEditState.value} onChange={(e) => setBulkEditState((c) => ({ ...c, value: e.target.value }))} placeholder="Examples: 25, 1.05, 499" /></label></div><div className="mini-form-row"><label className="field-block"><LabelWithHelp label="Clamp result" help="Min keeps the result from going above the cap. Max keeps the result from going below it." /><select value={bulkEditState.capMode} onChange={(e) => setBulkEditState((c) => ({ ...c, capMode: e.target.value }))}><option>Min</option><option>Max</option></select></label><label className="field-block"><LabelWithHelp label="Cap value" help="Optional. Leave blank if you do not want a floor or ceiling." /><input value={bulkEditState.capValue} onChange={(e) => setBulkEditState((c) => ({ ...c, capValue: e.target.value }))} placeholder="Optional floor or ceiling" /></label></div><button className="primary-button" disabled={!dataset} onClick={() => runAction(() => bulkEdit(dataset, filters, bulkEditState.mode, Number(bulkEditState.value || 0), bulkEditState.capValue ? bulkEditState.capMode : null, bulkEditState.capValue ? Number(bulkEditState.capValue) : null))}>Apply to Selected Rows</button><div className="status-button-row"><button className="status-chip status-chip-active" disabled={!dataset} onClick={() => runAction(() => setStatus(dataset, filters, "ACTIVE"))}>Mark Active</button><button className="status-chip status-chip-inactive" disabled={!dataset} onClick={() => runAction(() => setStatus(dataset, filters, "INACTIVE"))}>Mark Inactive</button></div></div></div>
          <div className="panel"><div className="panel-head"><div><div className="panel-title">Offer Mode <HelpTip text="Offer mode uses offer-file columns to calculate discount, final price, and ACCEPT or REJECT decisions." /></div><div className="panel-subtitle">Clearer defaults taken from the desktop workflow.</div></div>{dashboard?.availableModes?.normal && dashboard?.availableModes?.offer ? <div className="toggle-row"><button className={dashboard.mode === "normal" ? "toggle-button toggle-button-active" : "toggle-button"} onClick={() => handleModeChange("normal")}>Standard File</button><button className={dashboard.mode === "offer" ? "toggle-button toggle-button-active" : "toggle-button"} onClick={() => handleModeChange("offer")}>Offer File</button></div> : null}</div><div className="stack-grid"><div className="hint-box">Formula: discount = min((x% of (y% of settlement)), cap). Final price = settlement minus discount.</div><div className="mini-form-row three-inputs"><label className="field-block"><LabelWithHelp label="Base percent (y)" help="First percentage applied to the settlement value before the discount calculation." /><input value={offerConfig.yPct} onChange={(e) => setOfferConfig((c) => ({ ...c, yPct: e.target.value }))} placeholder="Default 15" /></label><label className="field-block"><LabelWithHelp label="Discount percent (x)" help="Second percentage applied on top of the y percent amount." /><input value={offerConfig.xPct} onChange={(e) => setOfferConfig((c) => ({ ...c, xPct: e.target.value }))} placeholder="Default 20" /></label><label className="field-block"><LabelWithHelp label="Discount cap (Rs)" help="Maximum discount allowed for each selected row." /><input value={offerConfig.cap} onChange={(e) => setOfferConfig((c) => ({ ...c, cap: e.target.value }))} placeholder="Default 500" /></label></div><div className="threshold-grid">{THRESHOLD_KEYS.map((key) => <label className="threshold-card" key={key}><LabelWithHelp label={THRESHOLD_LABELS[key]} help={`Rows in ${key} are marked ACCEPT when final price is at or above this floor.`} /><input value={offerConfig.thresholds[key]} onChange={(e) => setOfferConfig((c) => ({ ...c, thresholds: { ...c.thresholds, [key]: e.target.value } }))} placeholder={`Default ${DEFAULT_THRESHOLDS[key]}`} /></label>)}</div><div className="status-button-row"><button className="primary-button" disabled={!dataset || dashboard?.mode !== "offer"} onClick={() => runAction(() => computeDiscount(dataset, filters, Number(offerConfig.yPct || 0), Number(offerConfig.xPct || 0), Number(offerConfig.cap || 0)))}>Calculate Discount + Final Price</button><button className="ghost-button" disabled={!dataset || dashboard?.mode !== "offer"} onClick={() => runAction(() => applyDecision(dataset, filters, Object.fromEntries(Object.entries(offerConfig.thresholds).map(([key, value]) => [key, Number(value || 0)]))))}>Apply Accept / Reject</button></div></div></div>
          <div className="panel"><div className="panel-head"><div><div className="panel-title">Size Tools <HelpTip text="Use this when the auto-detected size is wrong for a SKU. The saved override is reused next time in this browser." /></div><div className="panel-subtitle">Saved in local browser storage.</div></div></div><div className="stack-grid"><label className="field-block"><LabelWithHelp label="SKU to override" help="Enter the exact SKU text from the workbook row you want to correct." /><input value={sizeOverride.sku} onChange={(e) => setSizeOverrideState((c) => ({ ...c, sku: e.target.value }))} placeholder="Paste exact SKU here" /></label><label className="field-block"><LabelWithHelp label="Correct size" help="Choose the size that should be assigned to this SKU." /><select value={sizeOverride.size} onChange={(e) => setSizeOverrideState((c) => ({ ...c, size: e.target.value }))}>{SIZE_VALUES.map((size) => <option key={size}>{size}</option>)}</select></label><button className="primary-button" disabled={!dataset} onClick={() => runAction(() => saveSizeOverride(dataset, filters, sizeOverride.sku, sizeOverride.size))}>Save Size Override</button><div className="hint-box">This React build stores overrides in <code>localStorage</code>, not a Python sidecar.</div></div></div>
        </section>
        <section className="panel"><div className="panel-head"><div><div className="panel-title">Listing Overview</div><div className="panel-subtitle">Focused column views keep the important fields visible without pushing everything into a wide spreadsheet.</div></div><div className="mono-value">{rows.length} / {dashboard?.exportCount || 0} rows shown</div></div><div className="listing-toolbar"><div className="listing-view-tabs">{Object.entries(LISTING_COLUMN_VIEWS).map(([key, config]) => <button key={key} className={listingView === key ? "listing-view-tab listing-view-tab-active" : "listing-view-tab"} onClick={() => setListingView(key)}>{config.label}</button>)}</div><div className="listing-toolbar-note">{displayedColumns.length} columns visible</div></div><div className="table-scroll"><table className="listing-table"><thead><tr>{displayedColumns.map((column, index) => <th key={column} className={index === 0 ? "sticky-first" : index === 1 ? "sticky-second" : ""}>{column}</th>)}</tr></thead><tbody>{rows.map((row) => { const status = String(row["Listing Status"] || "").toUpperCase(); const rowClass = status === "ACTIVE" ? "row-active" : status === "INACTIVE" ? "row-inactive" : ""; return <tr key={`${row.__orig_index}-${row[displayedColumns[0]] || "row"}`} className={rowClass}>{displayedColumns.map((column, index) => { const value = row[column]; const stickyClass = index === 0 ? "sticky-first-cell" : index === 1 ? "sticky-second-cell" : ""; if (column === "Listing Status") return <td key={column} className={stickyClass}><span className={`table-badge ${status === "ACTIVE" ? "table-badge-active" : status === "INACTIVE" ? "table-badge-inactive" : "table-badge-neutral"}`}>{status || "-"}</span></td>; if (column === "Auto Flag" && value) return <td key={column} className={`flag-cell ${stickyClass}`.trim()}><span className="flag-inline">! {String(value)}</span></td>; return <td key={column} className={stickyClass}>{value == null || value === "" ? "-" : String(value)}</td>; })}</tr>; })}</tbody></table></div></section>
      </>
    );
  }

  function renderInsightsPage() {
    return (
      <section className="panel"><div className="panel-head"><div><div className="panel-title">Recent Changes</div><div className="panel-subtitle">Change log, file context, and workbook mode.</div></div></div><div className="change-log-grid">{(dashboard?.changeLog || []).length ? dashboard.changeLog.map((entry, index) => <article className="change-card" key={`${entry.action}-${index}`}><div className="change-title">{entry.action}</div><div className="change-meta">Rows changed: {entry.rows}</div><div><strong>Before Min/Max:</strong> {entry.before}</div><div><strong>After Min/Max:</strong> {entry.after}</div>{entry.extra ? <div className="change-copy subtle">{entry.extra}</div> : null}</article>) : <div className="empty-state">Load a file, then your local edits and workbook decisions will show here.</div>}<article className="change-card"><div className="change-title">Session Context</div><div className="change-meta">Mode: {dashboard?.mode || "normal"}</div><div><strong>File:</strong> {dashboard?.fileName || "No file loaded"}</div><div><strong>Account:</strong> {dashboard?.accountName || "Unknown"}</div><div className="change-copy subtle">Using {dataset ? "browser-side JS workbook logic" : "no dataset yet"} inspired by <code>RateInsight.jsx</code>.</div></article></div></section>
    );
  }

  return (
    <div className="app-shell">
      <aside className="sidebar">
        <div className="sidebar-brand">Rate Insight</div>
        <nav className="sidebar-nav">{["Dashboard", "Listings", "Insights"].map((item) => <button key={item} className={page === item ? "nav-item nav-item-active" : "nav-item"} onClick={() => setPage(item)}>{item}</button>)}</nav>
        <div className="sidebar-user"><div className="avatar">RI</div><div><div className="user-name">Local Operator</div><div className="user-role">React Workbook Flow</div></div></div>
      </aside>
      <main className="main-shell">
        <header className="topbar">
          <div><h1>Rate Insight Dashboard</h1><p>Simple React project with workbook logic running locally in the browser.</p></div>
          <div className="topbar-actions"><button className={dataset ? "ghost-button" : "primary-button load-button-thematic"} onClick={() => fileInputRef.current?.click()}><span className="load-button-mark">+</span>{dataset ? "Load Excel" : "Load Workbook"}</button><button className="primary-button" disabled={!dataset} onClick={handleExport}>Export</button><button className="ghost-button" disabled={!dataset?.undoRows} onClick={() => runAction(() => undoDataset(dataset))}>Undo</button><input hidden ref={fileInputRef} type="file" accept=".xlsx,.xls,.csv" onChange={handleUpload} /></div>
        </header>
        <div className="page-scroll">
          <section className="hero-panel"><div><div className="eyebrow">Standalone React App</div><h2>Unified settlement review across dashboard, listing management, and decision insights.</h2><p>Review settlement trends, refine listings, apply pricing actions, and export results from one streamlined workspace.</p></div><div className="hero-statuses"><div className="hero-pill">{dashboard?.accountName || "No file loaded"}</div><div className="hero-pill">Mode: {dashboard?.mode || "normal"}</div>{loading ? <div className="hero-pill hero-pill-accent">Working…</div> : null}</div></section>
          {error ? <div className="error-banner">{error}</div> : null}
          {!dataset ? <section className="panel empty-upload-state"><div className="empty-upload-art"><div className="empty-upload-icon">RI</div><div className="empty-upload-lines"><span></span><span></span><span></span></div></div><div className="empty-upload-copy"><div className="eyebrow">Get Started</div><h3>Load your settlement workbook</h3><p>Start with an Excel or CSV export to unlock dashboard metrics, listing review, offer calculations, and final export.</p><div className="empty-upload-tags"><span>.xlsx</span><span>.xls</span><span>.csv</span></div><button className="primary-button empty-upload-button" onClick={() => fileInputRef.current?.click()}><span className="load-button-mark">+</span>Select Workbook</button></div></section> : null}
          {dataset && page === "Dashboard" ? renderDashboardPage() : null}
          {dataset && page === "Listings" ? renderListingsPage() : null}
          {dataset && page === "Insights" ? renderInsightsPage() : null}
        </div>
      </main>
    </div>
  );
}

export default App;













