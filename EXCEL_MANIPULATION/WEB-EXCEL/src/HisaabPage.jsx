import React, { useEffect, useState } from "react";
import {
  PREMADE_SHEET_NAME,
  formatCellDisplay,
  formatEditableValue,
  getPremadeSheet,
  loadPremadeWorkbook,
  updatePremadeWorkbookCell,
} from "./hisaabWorkbook";

function WorkbookCell({ sheetName, cell, sticky, onCommit }) {
  const className = [
    "hisaab-cell",
    cell.formula ? "hisaab-cell-formula" : "",
    sticky ? "hisaab-sticky-cell" : "",
  ].filter(Boolean).join(" ");

  if (cell.formula) {
    return (
      <td className={className} title={cell.formula}>
        <div>{formatCellDisplay(cell) || "-"}</div>
        <div className="hisaab-formula-tag">fx</div>
      </td>
    );
  }

  const defaultValue = formatEditableValue(cell);
  return (
    <td className={className}>
      <input
        key={`${sheetName}-${cell.address}-${defaultValue}`}
        className="hisaab-input"
        defaultValue={defaultValue}
        onBlur={(event) => onCommit(sheetName, cell.address, event.target.value)}
        onKeyDown={(event) => {
          if (event.key === "Enter") {
            event.preventDefault();
            event.currentTarget.blur();
          }
        }}
      />
    </td>
  );
}

export default function HisaabPage() {
  const [workbook, setWorkbook] = useState(null);
  const [loading, setLoading] = useState(true);
  const [error, setError] = useState("");

  async function loadWorkbook() {
    try {
      setLoading(true);
      setError("");
      const nextWorkbook = await loadPremadeWorkbook();
      setWorkbook(nextWorkbook);
    } catch (nextError) {
      setError(nextError instanceof Error ? nextError.message : "Failed to load workbook");
    } finally {
      setLoading(false);
    }
  }

  useEffect(() => {
    loadWorkbook();
  }, []);

  const sheet = getPremadeSheet(workbook, PREMADE_SHEET_NAME);

  function commitCell(sheetName, address, value) {
    setWorkbook((current) => (current ? updatePremadeWorkbookCell(current, sheetName, address, value) : current));
  }

  return (
    <section className="panel hisaab-shell">
      <div className="panel-head">
        <div>
          <div className="panel-title">Premade Hisaab Table</div>
          <div className="panel-subtitle">Showing only <code>{PREMADE_SHEET_NAME}</code> from bundled <code>Hisaab.xlsx</code>. Edits are temporary and reset on refresh or reset.</div>
        </div>
        <div className="topbar-actions">
          <button className="ghost-button" onClick={loadWorkbook} disabled={loading}>Reset Workbook</button>
        </div>
      </div>
      {error ? <div className="error-banner">{error}</div> : null}
      {loading ? <div className="empty-state">Loading workbook...</div> : null}
      {!loading && workbook ? (
        sheet ? (
          <div className="table-scroll hisaab-table-scroll">
            <table className="listing-table hisaab-table">
              <thead>
                <tr>
                  {sheet.columnLabels.map((label, index) => (
                    <th key={`${sheet.name}-${label}`} className={index === 0 ? "hisaab-sticky-head" : ""}>{label}</th>
                  ))}
                </tr>
              </thead>
              <tbody>
                {sheet.rows.map((row, rowIndex) => (
                  <tr key={`${sheet.name}-row-${rowIndex}`}>
                    {row.map((cell, cellIndex) => (
                      <WorkbookCell
                        key={cell.address}
                        sheetName={sheet.name}
                        cell={cell}
                        sticky={cellIndex === 0}
                        onCommit={commitCell}
                      />
                    ))}
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        ) : <div className="empty-state">Sheet4 was not found in the bundled workbook.</div>
      ) : null}
    </section>
  );
}
