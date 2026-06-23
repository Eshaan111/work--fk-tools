import { useState, useMemo } from "react";
import { BarChart, Bar, XAxis, YAxis, CartesianGrid, Tooltip, ResponsiveContainer, Cell } from "recharts";

// ─── Design Tokens ────────────────────────────────────────────────────────────
const C = {
  primary: "#091426",
  primaryContainer: "#1e293b",
  onPrimaryContainer: "#8590a6",
  surface: "#f7f9fb",
  surfaceLowest: "#ffffff",
  surfaceContainerLow: "#f2f4f6",
  surfaceContainer: "#eceef0",
  surfaceContainerHigh: "#e6e8ea",
  onSurface: "#191c1e",
  onSurfaceVariant: "#45474c",
  outlineVariant: "#c5c6cd",
  outline: "#75777d",
  emerald: "#10B981",
  rose: "#F43F5E",
  amber: "#F59E0B",
  lockedGray: "#94A3B8",
  indigo: "#6366F1",
  tertiaryContainer: "#002c42",
  onTertiary: "#0099d9",
  primaryFixed: "#d8e3fb",
};

// ─── Mock Data ─────────────────────────────────────────────────────────────────
const MOCK_SETTLEMENTS = [
  { id: "SKU-8924-A", description: "Logistics Module Base Unit", value: 1245.00, status: "Active", updated: "Oct 24, 10:42 AM", locked: false, flag: null },
  { id: "SKU-119X-F", description: "Sensor Array V2 (Deprecated)", value: 85.50, status: "Inactive", updated: "Oct 23, 04:15 PM", locked: false, flag: null },
  { id: "SKU-7742-R", description: "Transmission Hub Controller", value: 4120.00, status: "Review", updated: "Oct 24, 09:12 AM", locked: false, flag: "warning" },
  { id: "SKU-9910-Z", description: "Core Processor Unit Alpha", value: 12050.00, status: "Locked", updated: "Oct 20, 11:00 AM", locked: true, flag: null },
  { id: "SKU-2234-B", description: "Secondary Power Supply", value: 340.25, status: "Active", updated: "Oct 24, 08:30 AM", locked: false, flag: null },
  { id: "SKU-5512-C", description: "Network Interface Module", value: 678.00, status: "Active", updated: "Oct 24, 07:15 AM", locked: false, flag: null },
  { id: "SKU-3301-D", description: "Cooling Fan Assembly", value: 45.00, status: "Inactive", updated: "Oct 22, 03:00 PM", locked: false, flag: null },
  { id: "SKU-8801-E", description: "Power Regulator v3", value: 2200.00, status: "Active", updated: "Oct 24, 06:50 AM", locked: false, flag: "warning" },
];

const CHART_DATA = [
  { date: "Oct 1", volume: 26000 },
  { date: "Oct 8", volume: 44000 },
  { date: "", volume: 36000 },
  { date: "Oct 15", volume: 82000 },
  { date: "", volume: 75000 },
  { date: "Oct 22", volume: 64000 },
  { date: "", volume: 71000 },
  { date: "Oct 30", volume: 50000, accent: true },
  { date: "", volume: 84000 },
];

const CATEGORIES = [
  { name: "Jeans - Ice Wash", base: 1200, offset: 50, final: 1250, status: "Active" },
  { name: "Jeans - Beige", base: 1050, offset: -20, final: 1030, status: "Active" },
  { name: "Jeans - Raw Denim", base: 1500, offset: 0, final: 1500, status: "Locked", locked: true },
  { name: "Jackets - Denim", base: 2200, offset: 100, final: 2300, status: "Flagged" },
];

const ACCEPTANCE_DATA = [
  { range: "<₹500", count: 120 },
  { range: "₹500", count: 210 },
  { range: "₹1000", count: 340 },
  { range: "₹1500", count: 290 },
  { range: ">₹2000", count: 80 },
];

const CHANGE_LOG = [
  { id: 1, time: "2 mins ago", action: "Bulk Edit", detail: "Mode: Multiply, Value: 1.05", rows: 24, color: C.onTertiary },
  { id: 2, time: "14 mins ago", action: "Set Status → ACTIVE", detail: "Jeans: ICE | Size: 30, 32", rows: 8, color: C.emerald },
  { id: 3, time: "31 mins ago", action: "Freeze", detail: "Column contains: Raw Denim", rows: 3, color: C.lockedGray },
];

// ─── Shared Components ─────────────────────────────────────────────────────────
const Icon = ({ name, size = 20, style = {} }) => (
  <span className="material-symbols-outlined" style={{ fontSize: size, lineHeight: 1, ...style }}>{name}</span>
);

const StatusChip = ({ status }) => {
  const map = {
    Active:   { bg: `${C.emerald}18`, color: C.emerald, icon: "check_circle" },
    Inactive: { bg: `${C.rose}18`,    color: C.rose,    icon: "cancel" },
    Review:   { bg: `${C.amber}18`,   color: C.amber,   icon: "flag" },
    Locked:   { bg: `${C.lockedGray}18`, color: C.lockedGray, icon: "lock" },
    Flagged:  { bg: `${C.amber}18`,   color: C.amber,   icon: "flag" },
    Settled:  { bg: `${C.emerald}18`, color: C.emerald, icon: "check_circle" },
    Pending:  { bg: `${C.amber}18`,   color: C.amber,   icon: "schedule" },
    Disputed: { bg: `${C.rose}18`,    color: C.rose,    icon: "cancel" },
  };
  const s = map[status] || { bg: "#eee", color: "#555", icon: "help" };
  return (
    <span style={{
      display: "inline-flex", alignItems: "center", gap: 4,
      background: s.bg, color: s.color,
      padding: "2px 8px", borderRadius: 999, fontSize: 11, fontWeight: 600,
      fontFamily: "'JetBrains Mono', monospace",
    }}>
      <Icon name={s.icon} size={12} />
      {status}
    </span>
  );
};

const Btn = ({ children, variant = "primary", onClick, disabled, small, icon, style = {} }) => {
  const styles = {
    primary: { background: C.primaryContainer, color: "#fff", border: "none" },
    outline: { background: "transparent", color: C.primary, border: `1px solid ${C.outlineVariant}` },
    ghost:   { background: "transparent", color: C.primary, border: "none" },
    danger:  { background: C.rose, color: "#fff", border: "none" },
    teal:    { background: C.onTertiary, color: "#fff", border: "none" },
  };
  return (
    <button onClick={onClick} disabled={disabled} style={{
      ...styles[variant],
      borderRadius: 8, padding: small ? "5px 12px" : "7px 16px",
      fontSize: 13, fontWeight: 600, fontFamily: "'JetBrains Mono', monospace",
      cursor: disabled ? "not-allowed", opacity: disabled ? 0.5 : 1,
      display: "inline-flex", alignItems: "center", gap: 6,
      transition: "opacity .15s", whiteSpace: "nowrap",
      ...style,
    }}>
      {icon && <Icon name={icon} size={15} />}
      {children}
    </button>
  );
};

const Card = ({ children, style = {}, pad = 16 }) => (
  <div style={{
    background: C.surfaceLowest, border: `1px solid ${C.outlineVariant}`,
    borderRadius: 12, padding: pad, ...style,
  }}>{children}</div>
);

// ─── Top App Bar ────────────────────────────────────────────────────────────────
const TopBar = ({ page, setPage, offerMode }) => (
  <header style={{
    height: 44, background: C.surfaceLowest, borderBottom: `1px solid ${C.outlineVariant}`,
    display: "flex", alignItems: "center", justifyContent: "space-between",
    padding: "0 16px", position: "sticky", top: 0, zIndex: 100,
    flexShrink: 0,
  }}>
    <div style={{ display: "flex", alignItems: "center", gap: 16 }}>
      <Icon name="menu" size={22} style={{ color: C.onSurfaceVariant, cursor: "pointer" }} />
      <span style={{ fontFamily: "'Hanken Grotesk', sans-serif", fontWeight: 700, fontSize: 18, color: C.primary }}>
        Rate Insight
      </span>
    </div>
    <nav style={{ display: "flex", gap: 4 }}>
      {["Dashboard", "Listings", "Insights"].map(p => (
        <button key={p} onClick={() => setPage(p)} style={{
          background: "none", border: "none", cursor: "pointer",
          padding: "0 12px", height: 44,
          fontFamily: "'JetBrains Mono', monospace", fontSize: 13, fontWeight: 500,
          color: page === p ? C.primary : C.onSurfaceVariant,
          borderBottom: page === p ? `2px solid ${C.primary}` : "2px solid transparent",
          transition: "all .15s",
        }}>{p}</button>
      ))}
    </nav>
    <div style={{ display: "flex", alignItems: "center", gap: 10 }}>
      <Icon name="search" size={20} style={{ color: C.onSurfaceVariant }} />
      <div style={{
        width: 32, height: 32, borderRadius: "50%",
        background: C.primaryContainer, color: "#fff",
        display: "flex", alignItems: "center", justifyContent: "center",
        fontWeight: 700, fontSize: 13, cursor: "pointer",
        fontFamily: "'Hanken Grotesk', sans-serif",
      }}>UA</div>
    </div>
  </header>
);

// ─── KPI Card ──────────────────────────────────────────────────────────────────
const KpiCard = ({ label, value, delta, deltaPositive, icon }) => (
  <Card style={{ flex: 1, minWidth: 160 }}>
    <div style={{ display: "flex", justifyContent: "space-between", alignItems: "flex-start" }}>
      <div>
        <div style={{ fontSize: 11, fontWeight: 600, color: C.onSurfaceVariant, textTransform: "uppercase", letterSpacing: "0.08em", fontFamily: "'JetBrains Mono', monospace", marginBottom: 8 }}>
          {label}
        </div>
        <div style={{ fontFamily: "'Hanken Grotesk', sans-serif", fontSize: 28, fontWeight: 700, color: C.primary, lineHeight: 1 }}>
          {value}
        </div>
        <div style={{ marginTop: 6, fontSize: 12, color: deltaPositive === false ? C.rose : C.emerald, display: "flex", alignItems: "center", gap: 3 }}>
          {deltaPositive === true && <Icon name="trending_up" size={14} />}
          {deltaPositive === false && <Icon name="trending_up" size={14} style={{ transform: "rotate(180deg)" }} />}
          {deltaPositive === null && <Icon name="remove" size={14} style={{ color: C.onSurfaceVariant }} />}
          <span style={{ color: deltaPositive === null ? C.onSurfaceVariant : undefined }}>{delta}</span>
        </div>
      </div>
      <div style={{ color: C.onSurfaceVariant }}><Icon name={icon} size={22} /></div>
    </div>
  </Card>
);

// ─── DASHBOARD PAGE ────────────────────────────────────────────────────────────
const DashboardPage = () => {
  const [activeRange, setActiveRange] = useState("30D");
  const [filters, setFilters] = useState({
    fcl: true, lcl: true, air: false,
    settled: false, pending: false, disputed: false, flagged: false,
    region: "All Regions",
  });

  const recentSettlements = [
    { sku: "LDG-8921-A", date: "Oct 30, 2023", type: "FCL Import", amount: "$12,450.00", status: "Settled" },
    { sku: "LDG-8922-B", date: "Oct 30, 2023", type: "LCL Export", amount: "$3,120.50", status: "Flagged" },
    { sku: "LDG-8919-X", date: "Oct 29, 2023", type: "Air Expedited", amount: "$8,900.00", status: "Locked", locked: true },
  ];

  return (
    <div style={{ display: "flex", height: "100%", overflow: "hidden" }}>
      {/* Left Sidebar */}
      <nav style={{ width: 220, background: C.surfaceLowest, borderRight: `1px solid ${C.outlineVariant}`, flexShrink: 0, padding: "16px 0" }}>
        <div style={{ padding: "4px 16px 12px", fontSize: 11, fontWeight: 700, color: C.onSurfaceVariant, letterSpacing: "0.1em", textTransform: "uppercase", fontFamily: "'JetBrains Mono', monospace" }}>Navigation</div>
        {[
          { icon: "dashboard", label: "Dashboard", active: true },
          { icon: "list_alt", label: "Listings" },
          { icon: "insights", label: "Insights" },
        ].map(item => (
          <div key={item.label} style={{
            display: "flex", alignItems: "center", gap: 10, padding: "10px 16px",
            background: item.active ? `${C.primary}10` : "transparent",
            borderRadius: item.active ? "0 8px 8px 0" : 0,
            cursor: "pointer", marginRight: 8,
            color: item.active ? C.primary : C.onSurfaceVariant,
            fontFamily: "'Inter', sans-serif", fontSize: 14, fontWeight: item.active ? 600 : 400,
          }}>
            <Icon name={item.icon} size={18} />
            {item.label}
          </div>
        ))}
        <div style={{ position: "absolute", bottom: 16, left: 0, width: 220, padding: "0 16px", borderTop: `1px solid ${C.outlineVariant}`, paddingTop: 12 }}>
          <div style={{ display: "flex", alignItems: "center", gap: 10 }}>
            <div style={{ width: 32, height: 32, borderRadius: "50%", background: C.primaryContainer, color: "#fff", display: "flex", alignItems: "center", justifyContent: "center", fontWeight: 700, fontSize: 12 }}>JS</div>
            <div>
              <div style={{ fontSize: 13, fontWeight: 600, color: C.primary }}>Jane Smith</div>
              <div style={{ fontSize: 11, color: C.onSurfaceVariant }}>Logistics Manager</div>
            </div>
          </div>
        </div>
      </nav>

      {/* Main Content */}
      <div style={{ flex: 1, display: "flex", flexDirection: "column", overflow: "hidden" }}>
        {/* Top Controls */}
        <div style={{ padding: "12px 20px", background: C.surfaceLowest, borderBottom: `1px solid ${C.outlineVariant}`, display: "flex", justifyContent: "space-between", alignItems: "center", flexShrink: 0 }}>
          <div style={{ flex: 1, maxWidth: 380 }}>
            <div style={{ position: "relative" }}>
              <Icon name="search" size={16} style={{ position: "absolute", left: 10, top: "50%", transform: "translateY(-50%)", color: C.onSurfaceVariant }} />
              <input placeholder="Search shipments, SKUs, or locations..." style={{
                width: "100%", padding: "7px 10px 7px 32px", border: `1px solid ${C.outlineVariant}`,
                borderRadius: 8, fontSize: 13, background: C.surface, color: C.onSurface,
                fontFamily: "'Inter', sans-serif", outline: "none", boxSizing: "border-box",
              }} />
            </div>
          </div>
          <div style={{ display: "flex", gap: 8 }}>
            <Btn variant="outline" icon="upload_file">Load Excel</Btn>
            <Btn variant="primary" icon="download">Export</Btn>
          </div>
        </div>

        {/* KPI Cards */}
        <div style={{ padding: "16px 20px", display: "flex", gap: 12, flexShrink: 0 }}>
          <KpiCard label="Total Loaded" value="1,245,890" delta="+12.5% this month" deltaPositive={true} icon="database" />
          <KpiCard label="Visible Listings" value="892,104" delta="+4.2% this month" deltaPositive={true} icon="visibility" />
          <KpiCard label="Active Routes" value="342" delta="No change" deltaPositive={null} icon="route" />
          <KpiCard label="Inactive / Errors" value="1,402" delta="+2.1% spike today" deltaPositive={false} icon="error" />
        </div>

        {/* Chart + Refine Panel */}
        <div style={{ flex: 1, display: "flex", gap: 12, padding: "0 20px 16px", overflow: "hidden", minHeight: 0 }}>
          <div style={{ flex: 1, display: "flex", flexDirection: "column", gap: 12, overflow: "hidden" }}>
            {/* Settlement Volume Chart */}
            <Card style={{ flex: "0 0 auto" }}>
              <div style={{ display: "flex", justifyContent: "space-between", alignItems: "flex-start", marginBottom: 4 }}>
                <div>
                  <div style={{ fontFamily: "'Hanken Grotesk', sans-serif", fontSize: 18, fontWeight: 700, color: C.onSurface }}>Settlement Volume</div>
                  <div style={{ fontSize: 12, color: C.onSurfaceVariant, marginTop: 2 }}>30-day historical trend across top regions.</div>
                </div>
                <div style={{ display: "flex", gap: 4 }}>
                  {["7D", "30D", "YTD"].map(r => (
                    <button key={r} onClick={() => setActiveRange(r)} style={{
                      padding: "4px 10px", borderRadius: 6, fontSize: 12, fontWeight: 600,
                      fontFamily: "'JetBrains Mono', monospace", cursor: "pointer",
                      background: activeRange === r ? C.primary : "transparent",
                      color: activeRange === r ? "#fff" : C.onSurfaceVariant,
                      border: `1px solid ${activeRange === r ? C.primary : C.outlineVariant}`,
                    }}>{r}</button>
                  ))}
                </div>
              </div>
              <div style={{ height: 220, marginTop: 8 }}>
                <ResponsiveContainer width="100%" height="100%">
                  <BarChart data={CHART_DATA} barGap={2}>
                    <CartesianGrid vertical={false} stroke={C.outlineVariant} strokeOpacity={0.5} />
                    <XAxis dataKey="date" tick={{ fontSize: 11, fill: C.onSurfaceVariant, fontFamily: "Inter" }} axisLine={false} tickLine={false} />
                    <YAxis tick={{ fontSize: 11, fill: C.onSurfaceVariant, fontFamily: "Inter" }} axisLine={false} tickLine={false} tickFormatter={v => v >= 1000 ? `${v/1000}k` : v} />
                    <Tooltip formatter={v => [`${v.toLocaleString()}`, "Volume"]} contentStyle={{ fontSize: 12, borderRadius: 8, border: `1px solid ${C.outlineVariant}` }} />
                    <Bar dataKey="volume" radius={[3, 3, 0, 0]}>
                      {CHART_DATA.map((d, i) => (
                        <Cell key={i} fill={d.accent ? C.indigo : `${C.indigo}55`} />
                      ))}
                    </Bar>
                  </BarChart>
                </ResponsiveContainer>
              </div>
            </Card>

            {/* Recent Settlements */}
            <Card style={{ flex: 1, overflow: "hidden", display: "flex", flexDirection: "column" }}>
              <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 12 }}>
                <div style={{ fontFamily: "'Hanken Grotesk', sans-serif", fontSize: 16, fontWeight: 700, color: C.onSurface }}>Recent Settlements</div>
                <button style={{ background: "none", border: "none", color: C.onTertiary, fontSize: 13, cursor: "pointer", fontFamily: "'Inter', sans-serif", display: "flex", alignItems: "center", gap: 4 }}>
                  View Full Ledger <Icon name="arrow_forward" size={16} />
                </button>
              </div>
              <table style={{ width: "100%", borderCollapse: "collapse" }}>
                <thead>
                  <tr style={{ borderBottom: `1px solid ${C.outlineVariant}` }}>
                    {["SKU / ID", "Date", "Type", "Amount", "Status"].map(h => (
                      <th key={h} style={{ padding: "6px 12px", textAlign: "left", fontSize: 11, fontWeight: 600, color: C.onSurfaceVariant, fontFamily: "'JetBrains Mono', monospace", textTransform: "uppercase", letterSpacing: "0.06em" }}>{h}</th>
                    ))}
                  </tr>
                </thead>
                <tbody>
                  {recentSettlements.map((row, i) => (
                    <tr key={i} style={{ borderBottom: `1px solid ${C.outlineVariant}`, opacity: row.locked ? 0.55 : 1, background: row.locked ? C.surfaceContainer : "transparent" }}>
                      <td style={{ padding: "9px 12px", fontFamily: "'JetBrains Mono', monospace", fontSize: 13, color: row.locked ? C.lockedGray : C.primary, display: "flex", alignItems: "center", gap: 6 }}>
                        {row.locked && <Icon name="lock" size={13} style={{ color: C.lockedGray }} />}
                        {row.sku}
                      </td>
                      <td style={{ padding: "9px 12px", fontSize: 13, color: C.onSurfaceVariant }}>{row.date}</td>
                      <td style={{ padding: "9px 12px", fontSize: 13, color: C.onSurfaceVariant }}>{row.type}</td>
                      <td style={{ padding: "9px 12px", fontSize: 13, fontFamily: "'JetBrains Mono', monospace", color: row.locked ? C.lockedGray : C.onSurface }}>{row.amount}</td>
                      <td style={{ padding: "9px 12px" }}><StatusChip status={row.status} /></td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </Card>
          </div>

          {/* Refine Data Sidebar */}
          <Card style={{ width: 280, flexShrink: 0, display: "flex", flexDirection: "column", gap: 16, overflowY: "auto" }}>
            <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center" }}>
              <div style={{ fontFamily: "'Hanken Grotesk', sans-serif", fontSize: 16, fontWeight: 700, color: C.onSurface }}>Refine Data</div>
              <button style={{ background: "none", border: "none", color: C.onSurfaceVariant, fontSize: 12, cursor: "pointer", display: "flex", alignItems: "center", gap: 4, fontFamily: "'Inter', sans-serif" }}>
                <Icon name="filter_alt_off" size={14} /> Clear
              </button>
            </div>

            <div>
              <div style={{ fontSize: 13, fontWeight: 600, color: C.onSurface, marginBottom: 10 }}>Listing Type</div>
              {[{ label: "Standard Freight (FCL)", key: "fcl" }, { label: "Less than Container (LCL)", key: "lcl" }, { label: "Air Freight Expedited", key: "air" }].map(({ label, key }) => (
                <label key={key} style={{ display: "flex", alignItems: "center", gap: 8, marginBottom: 8, cursor: "pointer", fontSize: 13, color: C.onSurfaceVariant }}>
                  <input type="checkbox" checked={filters[key]} onChange={() => setFilters(f => ({ ...f, [key]: !f[key] }))}
                    style={{ accentColor: C.primary, width: 15, height: 15 }} />
                  {label}
                </label>
              ))}
            </div>

            <hr style={{ border: "none", borderTop: `1px solid ${C.outlineVariant}` }} />

            <div>
              <div style={{ fontSize: 13, fontWeight: 600, color: C.onSurface, marginBottom: 10 }}>Settlement Status</div>
              <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 8 }}>
                {["Settled", "Pending", "Disputed", "Flagged"].map(s => (
                  <button key={s} onClick={() => setFilters(f => ({ ...f, [s.toLowerCase()]: !f[s.toLowerCase()] }))}
                    style={{
                      padding: "6px 10px", borderRadius: 8, fontSize: 12, fontWeight: 600,
                      fontFamily: "'JetBrains Mono', monospace", cursor: "pointer",
                      display: "flex", alignItems: "center", gap: 5,
                      border: `1px solid ${filters[s.toLowerCase()] ? C.primary : C.outlineVariant}`,
                      background: filters[s.toLowerCase()] ? `${C.primary}10` : C.surfaceLowest,
                      color: filters[s.toLowerCase()] ? C.primary : C.onSurfaceVariant,
                    }}>
                    <StatusChip status={s} />
                  </button>
                ))}
              </div>
            </div>

            <hr style={{ border: "none", borderTop: `1px solid ${C.outlineVariant}` }} />

            <div>
              <div style={{ fontSize: 13, fontWeight: 600, color: C.onSurface, marginBottom: 10 }}>Origin Regions</div>
              <select value={filters.region} onChange={e => setFilters(f => ({ ...f, region: e.target.value }))}
                style={{ width: "100%", padding: "8px 12px", border: `1px solid ${C.outlineVariant}`, borderRadius: 8, fontSize: 13, background: C.surfaceLowest, color: C.onSurface, outline: "none", fontFamily: "'Inter', sans-serif" }}>
                {["All Regions", "Asia Pacific", "North America", "Europe", "Middle East"].map(r => <option key={r}>{r}</option>)}
              </select>
            </div>

            <Btn variant="primary" style={{ width: "100%", justifyContent: "center", marginTop: "auto" }}>Apply Filters</Btn>
          </Card>
        </div>
      </div>
    </div>
  );
};

// ─── LISTINGS PAGE ─────────────────────────────────────────────────────────────
const ListingsPage = () => {
  const [search, setSearch] = useState("");
  const [selected, setSelected] = useState(new Set());
  const [page, setPage] = useState(1);

  const toggleSelect = (id) => {
    setSelected(prev => {
      const next = new Set(prev);
      next.has(id) ? next.delete(id) : next.add(id);
      return next;
    });
  };

  const filtered = useMemo(() =>
    MOCK_SETTLEMENTS.filter(r =>
      r.id.toLowerCase().includes(search.toLowerCase()) ||
      r.description.toLowerCase().includes(search.toLowerCase())
    ), [search]);

  return (
    <div style={{ display: "flex", height: "100%", overflow: "hidden" }}>
      {/* Main Data Area */}
      <div style={{ flex: 1, display: "flex", flexDirection: "column", overflow: "hidden" }}>
        {/* Header */}
        <div style={{ padding: "16px 24px", background: C.surfaceLowest, borderBottom: `1px solid ${C.outlineVariant}`, flexShrink: 0 }}>
          <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center" }}>
            <div>
              <h2 style={{ fontFamily: "'Hanken Grotesk', sans-serif", fontSize: 26, fontWeight: 700, color: C.primary, margin: 0 }}>Settlement Data</h2>
              <p style={{ fontSize: 12, color: C.onSurfaceVariant, margin: "4px 0 0" }}>Manage and audit SKU-level settlement values.</p>
            </div>
            <div style={{ display: "flex", alignItems: "center", gap: 20 }}>
              <div style={{ textAlign: "right" }}>
                <div style={{ fontSize: 11, fontWeight: 600, color: C.onSurfaceVariant, textTransform: "uppercase", letterSpacing: "0.08em", fontFamily: "'JetBrains Mono', monospace" }}>Total Active</div>
                <div style={{ fontFamily: "'Hanken Grotesk', sans-serif", fontSize: 24, fontWeight: 700, color: C.primary }}>12,450</div>
              </div>
              <div style={{ width: 1, height: 32, background: C.outlineVariant }} />
              <div style={{ textAlign: "right" }}>
                <div style={{ fontSize: 11, fontWeight: 600, color: C.onSurfaceVariant, textTransform: "uppercase", letterSpacing: "0.08em", fontFamily: "'JetBrains Mono', monospace" }}>Total Inactive</div>
                <div style={{ fontFamily: "'Hanken Grotesk', sans-serif", fontSize: 24, fontWeight: 700, color: C.rose }}>892</div>
              </div>
            </div>
          </div>
          <div style={{ display: "flex", gap: 10, marginTop: 14, alignItems: "center" }}>
            <div style={{ position: "relative", flex: 1, maxWidth: 380 }}>
              <Icon name="search" size={16} style={{ position: "absolute", left: 10, top: "50%", transform: "translateY(-50%)", color: C.onSurfaceVariant }} />
              <input value={search} onChange={e => setSearch(e.target.value)}
                placeholder="Search SKU or Description..."
                style={{ width: "100%", padding: "8px 10px 8px 32px", border: `1px solid ${C.outlineVariant}`, borderRadius: 8, fontSize: 13, background: C.surface, fontFamily: "'Inter', sans-serif", outline: "none", boxSizing: "border-box" }} />
            </div>
            <Btn variant="outline" icon="filter_list">Filters</Btn>
            <Btn variant="outline" icon="download" style={{ marginLeft: "auto" }}>Export</Btn>
          </div>
        </div>

        {/* Table */}
        <div style={{ flex: 1, overflow: "auto" }}>
          <table style={{ width: "100%", borderCollapse: "collapse" }}>
            <thead style={{ position: "sticky", top: 0, background: "#F1F5F9", zIndex: 10 }}>
              <tr>
                <th style={{ width: 48, padding: "10px 16px", borderBottom: `1px solid ${C.outlineVariant}` }}>
                  <input type="checkbox" style={{ accentColor: C.primary, width: 15, height: 15 }} />
                </th>
                {["SKU ID", "Description", "Settlement Value", "Status", "Last Updated", ""].map(h => (
                  <th key={h} style={{ padding: "10px 16px", borderBottom: `1px solid ${C.outlineVariant}`, textAlign: h === "Settlement Value" ? "right" : "left", fontSize: 11, fontWeight: 600, color: C.onSurfaceVariant, fontFamily: "'JetBrains Mono', monospace", textTransform: "uppercase", letterSpacing: "0.06em", whiteSpace: "nowrap" }}>{h}</th>
                ))}
              </tr>
            </thead>
            <tbody>
              {filtered.map(row => (
                <tr key={row.id} style={{ borderBottom: `1px solid ${C.outlineVariant}`, background: row.locked ? `${C.surfaceContainer}88` : "white", transition: "background .1s" }}
                  onMouseEnter={e => { if (!row.locked) e.currentTarget.style.background = "#F1F5F9"; }}
                  onMouseLeave={e => { e.currentTarget.style.background = row.locked ? `${C.surfaceContainer}88` : "white"; }}>
                  <td style={{ padding: "7px 16px" }}>
                    {row.locked
                      ? <Icon name="lock" size={15} style={{ color: C.lockedGray }} />
                      : <input type="checkbox" checked={selected.has(row.id)} onChange={() => toggleSelect(row.id)} style={{ accentColor: C.primary, width: 15, height: 15 }} />
                    }
                  </td>
                  <td style={{ padding: "7px 16px", fontFamily: "'JetBrains Mono', monospace", fontSize: 13, color: row.locked ? C.lockedGray : row.flag === "warning" ? C.amber : C.primary }}>
                    <div style={{ display: "flex", alignItems: "center", gap: 5 }}>
                      {row.flag === "warning" && <Icon name="warning" size={13} style={{ color: C.amber }} />}
                      {row.id}
                    </div>
                  </td>
                  <td style={{ padding: "7px 16px", fontSize: 13, color: row.locked ? C.lockedGray : C.onSurfaceVariant, maxWidth: 220, overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }}>{row.description}</td>
                  <td style={{ padding: "7px 16px", textAlign: "right", fontFamily: "'JetBrains Mono', monospace", fontSize: 13, color: row.locked ? C.lockedGray : C.onSurface }}>${row.value.toLocaleString("en-US", { minimumFractionDigits: 2 })}</td>
                  <td style={{ padding: "7px 16px" }}><StatusChip status={row.status} /></td>
                  <td style={{ padding: "7px 16px", fontSize: 12, color: row.locked ? C.lockedGray : C.onSurfaceVariant }}>{row.updated}</td>
                  <td style={{ padding: "7px 16px", textAlign: "right" }}>
                    <button style={{ background: "none", border: "none", cursor: "pointer", color: C.onSurfaceVariant }}>
                      <Icon name="more_vert" size={18} />
                    </button>
                  </td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>

        {/* Pagination */}
        <div style={{ height: 44, background: C.surfaceLowest, borderTop: `1px solid ${C.outlineVariant}`, display: "flex", alignItems: "center", justifyContent: "space-between", padding: "0 16px", flexShrink: 0 }}>
          <span style={{ fontSize: 12, color: C.onSurfaceVariant, fontFamily: "'Inter', sans-serif" }}>Showing 1-{filtered.length} of 13,342 entries</span>
          <div style={{ display: "flex", alignItems: "center", gap: 12 }}>
            <button onClick={() => setPage(p => Math.max(1, p - 1))} style={{ background: "none", border: "none", cursor: "pointer", color: C.onSurfaceVariant }}>
              <Icon name="chevron_left" size={20} />
            </button>
            <span style={{ fontSize: 12, fontFamily: "'JetBrains Mono', monospace", color: C.onSurface }}>Page {page} of 2669</span>
            <button onClick={() => setPage(p => p + 1)} style={{ background: "none", border: "none", cursor: "pointer", color: C.onSurfaceVariant }}>
              <Icon name="chevron_right" size={20} />
            </button>
          </div>
        </div>
      </div>

      {/* Right Sidebar */}
      <aside style={{ width: 220, background: C.surfaceLowest, borderLeft: `1px solid ${C.outlineVariant}`, display: "flex", flexDirection: "column", padding: 20, gap: 24, overflowY: "auto", flexShrink: 0 }}>
        <section>
          <div style={{ fontSize: 11, fontWeight: 600, color: C.onSurfaceVariant, textTransform: "uppercase", letterSpacing: "0.1em", fontFamily: "'JetBrains Mono', monospace", marginBottom: 12, display: "flex", alignItems: "center", gap: 6 }}>
            <Icon name="straighten" size={14} /> Size Tools
          </div>
          {[{ icon: "cloud_download", label: "Download Undetected", sub: "Export missing dim data" },
            { icon: "upload_file", label: "Upload Sizes", sub: "Import CSV bulk dims" }].map(btn => (
            <button key={btn.label} style={{ width: "100%", display: "flex", alignItems: "center", gap: 10, padding: "10px 12px", background: C.surfaceContainer, border: `1px solid ${C.outlineVariant}`, borderRadius: 10, cursor: "pointer", textAlign: "left", marginBottom: 8, transition: "all .15s" }}
              onMouseEnter={e => { e.currentTarget.style.borderColor = C.primary; e.currentTarget.style.background = C.surfaceContainerHigh; }}
              onMouseLeave={e => { e.currentTarget.style.borderColor = C.outlineVariant; e.currentTarget.style.background = C.surfaceContainer; }}>
              <div style={{ width: 32, height: 32, borderRadius: 6, background: C.surfaceLowest, border: `1px solid ${C.outlineVariant}`, display: "flex", alignItems: "center", justifyContent: "center", flexShrink: 0 }}>
                <Icon name={btn.icon} size={16} style={{ color: C.onSurfaceVariant }} />
              </div>
              <div>
                <div style={{ fontSize: 12, fontWeight: 600, color: C.primary, fontFamily: "'JetBrains Mono', monospace" }}>{btn.label}</div>
                <div style={{ fontSize: 11, color: C.onSurfaceVariant, marginTop: 2 }}>{btn.sub}</div>
              </div>
            </button>
          ))}
        </section>

        <hr style={{ border: "none", borderTop: `1px solid ${C.outlineVariant}` }} />

        <section>
          <div style={{ fontSize: 11, fontWeight: 600, color: C.onSurfaceVariant, textTransform: "uppercase", letterSpacing: "0.1em", fontFamily: "'JetBrains Mono', monospace", marginBottom: 12, display: "flex", alignItems: "center", gap: 6 }}>
            <Icon name="edit_square" size={14} /> Bulk Actions
          </div>
          <Card pad={14}>
            <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 12 }}>
              <span style={{ fontSize: 13, color: C.onSurfaceVariant }}>Selected Rows:</span>
              <span style={{ fontSize: 13, fontWeight: 700, fontFamily: "'JetBrains Mono', monospace", color: C.primary, background: C.primaryFixed, padding: "2px 8px", borderRadius: 4 }}>{selected.size}</span>
            </div>
            <div style={{ display: "flex", flexDirection: "column", gap: 6 }}>
              <Btn variant="outline" icon="edit" disabled={selected.size === 0} style={{ width: "100%", justifyContent: "center", background: selected.size > 0 ? C.primaryContainer : undefined, color: selected.size > 0 ? "#fff" : undefined, border: "none", opacity: selected.size === 0 ? 0.45 : 1 }}>Edit Selected</Btn>
              <Btn variant="outline" icon="ac_unit" disabled={selected.size === 0} style={{ width: "100%", justifyContent: "center", opacity: selected.size === 0 ? 0.45 : 1 }}>Freeze Selected</Btn>
            </div>
          </Card>
        </section>

        <section style={{ marginTop: "auto" }}>
          <Btn variant="danger" icon="lock_open" style={{ width: "100%", justifyContent: "center" }}>Unfreeze All</Btn>
        </section>
      </aside>
    </div>
  );
};

// ─── INSIGHTS PAGE (Offer Mode Configuration) ──────────────────────────────────
const InsightsPage = () => {
  const [margin, setMargin] = useState("15.5");
  const [discount, setDiscount] = useState("20.0");
  const [cap, setCap] = useState("500");
  const [offsets, setOffsets] = useState({ 0: 50, 1: -20, 2: 0, 3: 100 });
  const [lastRun, setLastRun] = useState("2 mins ago");

  const categories = useMemo(() => CATEGORIES.map((c, i) => ({
    ...c,
    offset: offsets[i] ?? 0,
    final: c.base + (offsets[i] ?? 0),
  })), [offsets]);

  const handleRunSim = () => setLastRun("just now");

  return (
    <div style={{ display: "flex", height: "100%", overflow: "hidden" }}>
      {/* Left Panel */}
      <div style={{ flex: 1, padding: 24, overflowY: "auto", borderRight: `1px solid ${C.outlineVariant}` }}>
        <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 24 }}>
          <h2 style={{ fontFamily: "'Hanken Grotesk', sans-serif", fontSize: 22, fontWeight: 700, color: C.primary, margin: 0 }}>Offer Mode Configuration</h2>
          <Btn variant="teal" icon="play_arrow" onClick={handleRunSim}>Run Simulation</Btn>
        </div>

        {/* Global Parameters */}
        <Card pad={20} style={{ marginBottom: 16 }}>
          <div style={{ display: "flex", alignItems: "center", gap: 8, marginBottom: 16 }}>
            <Icon name="tune" size={18} style={{ color: C.onSurfaceVariant }} />
            <span style={{ fontFamily: "'Hanken Grotesk', sans-serif", fontSize: 16, fontWeight: 700, color: C.onSurface }}>Global Parameters</span>
          </div>
          <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr 1fr", gap: 16 }}>
            {[{ label: "Margin (y%)", value: margin, set: setMargin, suffix: "%" },
              { label: "Discount (x%)", value: discount, set: setDiscount, suffix: "%" },
              { label: "Cap (Rs)", value: cap, set: setCap, prefix: "₹" }].map(({ label, value, set, suffix, prefix }) => (
              <div key={label}>
                <div style={{ fontSize: 12, color: C.onSurfaceVariant, marginBottom: 6, fontFamily: "'Inter', sans-serif" }}>{label}</div>
                <div style={{ position: "relative" }}>
                  {prefix && <span style={{ position: "absolute", left: 10, top: "50%", transform: "translateY(-50%)", color: C.onSurfaceVariant, fontSize: 13 }}>{prefix}</span>}
                  <input value={value} onChange={e => set(e.target.value)}
                    style={{ width: "100%", padding: `8px ${suffix ? "32px" : "10px"} 8px ${prefix ? "22px" : "10px"}`, border: `1px solid ${C.outlineVariant}`, borderRadius: 8, fontSize: 14, fontFamily: "'JetBrains Mono', monospace", outline: "none", boxSizing: "border-box", background: C.surfaceLowest }} />
                  {suffix && <span style={{ position: "absolute", right: 10, top: "50%", transform: "translateY(-50%)", color: C.onSurfaceVariant, fontSize: 13 }}>{suffix}</span>}
                </div>
              </div>
            ))}
          </div>
        </Card>

        {/* Category Thresholds Table */}
        <Card pad={20}>
          <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 16 }}>
            <div style={{ display: "flex", alignItems: "center", gap: 8 }}>
              <Icon name="category" size={18} style={{ color: C.onSurfaceVariant }} />
              <span style={{ fontFamily: "'Hanken Grotesk', sans-serif", fontSize: 16, fontWeight: 700, color: C.onSurface }}>Category Thresholds</span>
            </div>
            <div style={{ display: "flex", gap: 8 }}>
              <Btn variant="outline" small onClick={() => setOffsets({ 0: 0, 1: 0, 2: 0, 3: 0 })}>Reset</Btn>
              <Btn variant="primary" small>Apply Offset</Btn>
            </div>
          </div>
          <table style={{ width: "100%", borderCollapse: "collapse" }}>
            <thead>
              <tr style={{ borderBottom: `1px solid ${C.outlineVariant}` }}>
                {["Category / Type", "Base Threshold", "Offset Adjustment", "Final Value", "Status"].map(h => (
                  <th key={h} style={{ padding: "8px 12px", textAlign: "left", fontSize: 11, fontWeight: 600, color: C.onSurfaceVariant, fontFamily: "'JetBrains Mono', monospace", textTransform: "uppercase", letterSpacing: "0.06em" }}>{h}</th>
                ))}
              </tr>
            </thead>
            <tbody>
              {categories.map((cat, i) => (
                <tr key={i} style={{ borderBottom: `1px solid ${C.outlineVariant}`, background: cat.locked ? C.surfaceContainerLow : "transparent", opacity: cat.locked ? 0.7 : 1 }}>
                  <td style={{ padding: "10px 12px", fontSize: 14, color: cat.locked ? C.lockedGray : C.onSurface, display: "flex", alignItems: "center", gap: 6 }}>
                    {cat.locked && <Icon name="lock" size={14} style={{ color: C.lockedGray }} />}
                    {cat.name}
                  </td>
                  <td style={{ padding: "10px 12px", fontFamily: "'JetBrains Mono', monospace", fontSize: 13, color: cat.locked ? C.lockedGray : C.onSurface }}>₹ {cat.base.toLocaleString()}</td>
                  <td style={{ padding: "10px 12px" }}>
                    <input value={offsets[i] ?? 0} disabled={cat.locked}
                      onChange={e => setOffsets(p => ({ ...p, [i]: Number(e.target.value) }))}
                      style={{ width: 80, padding: "5px 8px", border: `1px solid ${C.outlineVariant}`, borderRadius: 6, fontFamily: "'JetBrains Mono', monospace", fontSize: 13, textAlign: "center", background: cat.locked ? C.surfaceContainer : C.surfaceLowest, color: cat.locked ? C.lockedGray : C.onSurface, outline: "none" }} />
                  </td>
                  <td style={{ padding: "10px 12px", fontFamily: "'JetBrains Mono', monospace", fontSize: 13, color: cat.locked ? C.lockedGray : C.onSurface }}>₹ {cat.final.toLocaleString()}</td>
                  <td style={{ padding: "10px 12px" }}>
                    {cat.locked ? (
                      <span style={{ display: "inline-flex", alignItems: "center", gap: 4, fontSize: 12, color: C.lockedGray, fontFamily: "'JetBrains Mono', monospace" }}>
                        <Icon name="lock" size={12} /> Locked
                      </span>
                    ) : <StatusChip status={cat.status} />}
                  </td>
                </tr>
              ))}
            </tbody>
          </table>
        </Card>
      </div>

      {/* Right Panel: Simulation Results */}
      <div style={{ width: 340, flexShrink: 0, padding: 24, overflowY: "auto", background: C.surface }}>
        <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 20 }}>
          <div style={{ display: "flex", alignItems: "center", gap: 8 }}>
            <Icon name="bar_chart" size={18} style={{ color: C.onTertiary }} />
            <span style={{ fontFamily: "'Hanken Grotesk', sans-serif", fontSize: 16, fontWeight: 700, color: C.onSurface }}>Simulation Results</span>
          </div>
          <span style={{ fontSize: 11, color: C.onSurfaceVariant }}>Last run: {lastRun}</span>
        </div>

        {/* Projected KPIs */}
        <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 10, marginBottom: 16 }}>
          <Card style={{ borderLeft: `3px solid ${C.emerald}` }} pad={14}>
            <div style={{ fontSize: 11, fontWeight: 600, color: C.onSurfaceVariant, textTransform: "uppercase", fontFamily: "'JetBrains Mono', monospace", letterSpacing: "0.06em", marginBottom: 6 }}>Projected Accepted</div>
            <div style={{ fontFamily: "'Hanken Grotesk', sans-serif", fontSize: 24, fontWeight: 700, color: C.primary }}>12,450</div>
            <div style={{ fontSize: 12, color: C.emerald, marginTop: 4, display: "flex", alignItems: "center", gap: 3 }}>
              <Icon name="trending_up" size={13} /> +5.2% vs baseline
            </div>
          </Card>
          <Card style={{ borderLeft: `3px solid ${C.rose}` }} pad={14}>
            <div style={{ fontSize: 11, fontWeight: 600, color: C.onSurfaceVariant, textTransform: "uppercase", fontFamily: "'JetBrains Mono', monospace", letterSpacing: "0.06em", marginBottom: 6 }}>Projected Rejected</div>
            <div style={{ fontFamily: "'Hanken Grotesk', sans-serif", fontSize: 24, fontWeight: 700, color: C.primary }}>3,120</div>
            <div style={{ fontSize: 12, color: C.rose, marginTop: 4, display: "flex", alignItems: "center", gap: 3 }}>
              <Icon name="trending_down" size={13} /> -1.8% vs baseline
            </div>
          </Card>
        </div>

        {/* Acceptance Distribution Chart */}
        <Card pad={16} style={{ marginBottom: 16 }}>
          <div style={{ fontSize: 11, fontWeight: 600, color: C.onSurfaceVariant, textTransform: "uppercase", fontFamily: "'JetBrains Mono', monospace", letterSpacing: "0.08em", marginBottom: 12 }}>Acceptance Distribution</div>
          <div style={{ height: 120 }}>
            <ResponsiveContainer width="100%" height="100%">
              <BarChart data={ACCEPTANCE_DATA} barGap={2}>
                <XAxis dataKey="range" tick={{ fontSize: 10, fill: C.onSurfaceVariant }} axisLine={false} tickLine={false} />
                <YAxis hide />
                <Bar dataKey="count" radius={[3, 3, 0, 0]}>
                  {ACCEPTANCE_DATA.map((d, i) => (
                    <Cell key={i} fill={i === ACCEPTANCE_DATA.length - 1 ? `${C.rose}88` : `${C.indigo}88`} />
                  ))}
                </Bar>
              </BarChart>
            </ResponsiveContainer>
          </div>
        </Card>

        {/* Simulation Insights */}
        <Card pad={16}>
          <div style={{ fontSize: 11, fontWeight: 600, color: C.onSurfaceVariant, textTransform: "uppercase", fontFamily: "'JetBrains Mono', monospace", letterSpacing: "0.08em", marginBottom: 12 }}>Simulation Insights</div>
          {[
            { icon: "warning", color: C.amber, title: "Margin impact alert", detail: "Current settings reduce overall margin by 1.2% in the 'Jeans' category." },
            { icon: "check_circle", color: C.emerald, title: "Optimal Cap Reached", detail: "Cap of ₹500 aligns with historical high-conversion bands." },
          ].map((ins, i) => (
            <div key={i} style={{ display: "flex", gap: 10, marginBottom: 12 }}>
              <Icon name={ins.icon} size={18} style={{ color: ins.color, flexShrink: 0, marginTop: 1 }} />
              <div>
                <div style={{ fontSize: 13, fontWeight: 600, color: C.onSurface, marginBottom: 2 }}>{ins.title}</div>
                <div style={{ fontSize: 12, color: C.onSurfaceVariant, lineHeight: 1.5 }}>{ins.detail}</div>
              </div>
            </div>
          ))}
        </Card>

        {/* Change Log */}
        <div style={{ marginTop: 16 }}>
          <div style={{ fontSize: 11, fontWeight: 600, color: C.onSurfaceVariant, textTransform: "uppercase", fontFamily: "'JetBrains Mono', monospace", letterSpacing: "0.08em", marginBottom: 10 }}>Recent Changes</div>
          {CHANGE_LOG.map(log => (
            <div key={log.id} style={{ borderLeft: `3px solid ${log.color}`, background: C.surfaceLowest, border: `1px solid ${C.outlineVariant}`, borderLeftWidth: 3, borderLeftColor: log.color, borderRadius: 8, padding: "10px 12px", marginBottom: 8 }}>
              <div style={{ fontSize: 12, fontWeight: 700, color: C.onSurface }}>{log.action}</div>
              <div style={{ fontSize: 11, color: C.onSurfaceVariant, marginTop: 2 }}>{log.detail}</div>
              <div style={{ fontSize: 11, color: C.onSurfaceVariant, marginTop: 4 }}>Rows: {log.rows} · {log.time}</div>
            </div>
          ))}
        </div>
      </div>
    </div>
  );
};

// ─── ROOT APP ──────────────────────────────────────────────────────────────────
export default function App() {
  const [page, setPage] = useState("Dashboard");

  return (
    <div style={{ height: "100vh", display: "flex", flexDirection: "column", background: C.surface, fontFamily: "'Inter', sans-serif", overflow: "hidden" }}>
      <link href="https://fonts.googleapis.com/css2?family=Hanken+Grotesk:wght@600;700&family=Inter:wght@400;500;600&family=JetBrains+Mono:wght@500&display=swap" rel="stylesheet" />
      <link href="https://fonts.googleapis.com/css2?family=Material+Symbols+Outlined:wght,FILL@400,0&display=swap" rel="stylesheet" />
      <style>{`
        * { box-sizing: border-box; }
        ::-webkit-scrollbar { width: 6px; height: 6px; }
        ::-webkit-scrollbar-track { background: transparent; }
        ::-webkit-scrollbar-thumb { background: #CBD5E1; border-radius: 4px; }
        input, button, select { box-sizing: border-box; }
        .material-symbols-outlined { font-family: 'Material Symbols Outlined'; font-variation-settings: 'FILL' 0, 'wght' 400; }
      `}</style>
      <TopBar page={page} setPage={setPage} />
      <div style={{ flex: 1, overflow: "hidden", position: "relative" }}>
        {page === "Dashboard" && <DashboardPage />}
        {page === "Listings"  && <ListingsPage />}
        {page === "Insights"  && <InsightsPage />}
      </div>
    </div>
  );
}
