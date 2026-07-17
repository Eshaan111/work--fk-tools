import React, { useEffect, useMemo, useRef, useState } from "react";
import {
  PREMADE_SHEET_NAME,
  exportHisaabWorkbook,
  formatCellDisplay,
  formatEditableValue,
  getPremadeSheet,
  loadHisaabWorkbookFile,
  loadPremadeWorkbook,
  updatePremadeWorkbookCell,
} from "./hisaabWorkbook";

const HISAAB_EDITS_KEY = "web_excel_hisaab_edits_v2";

function WorkbookCell({ sheetName, cell, onCommit }) {
  if (cell.formula) {
    return <td className={`hisaab-cell hisaab-cell-formula ${cell.error ? "hisaab-cell-error" : ""}`} title={cell.error || cell.formula}><div>{cell.error ? "#ERROR" : formatCellDisplay(cell) || "-"}</div><div className="hisaab-formula-tag">fx</div></td>;
  }
  const defaultValue = formatEditableValue(cell);
  return <td className="hisaab-cell"><input key={`${sheetName}-${cell.address}-${defaultValue}`} className="hisaab-input" defaultValue={defaultValue} aria-label={`${sheetName} ${cell.address}`} onBlur={(event) => onCommit(sheetName, cell.address, event.target.value)} onKeyDown={(event) => { if (event.key === "Enter") { event.preventDefault(); event.currentTarget.blur(); } }} /></td>;
}

function readSavedEdits() {
  try { return JSON.parse(window.localStorage.getItem(HISAAB_EDITS_KEY) || "{}"); } catch { return {}; }
}

function restoreEdits(workbook) {
  const saved = readSavedEdits()[workbook.title] || {};
  return Object.entries(saved).reduce((current, [key, value]) => {
    const separator = key.indexOf("!");
    return separator > 0 ? updatePremadeWorkbookCell(current, key.slice(0, separator), key.slice(separator + 1), value) : current;
  }, workbook);
}

export default function HisaabPage() {
  const [workbook, setWorkbook] = useState(null);
  const [activeSheet, setActiveSheet] = useState(PREMADE_SHEET_NAME);
  const [loading, setLoading] = useState(true);
  const [error, setError] = useState("");
  const [notice, setNotice] = useState("");
  const fileInputRef = useRef(null);

  async function loadBundledWorkbook({ clearSaved = false } = {}) {
    try {
      setLoading(true);
      setError("");
      if (clearSaved) {
        const saved = readSavedEdits();
        delete saved["Hisaab.xlsx"];
        window.localStorage.setItem(HISAAB_EDITS_KEY, JSON.stringify(saved));
      }
      const loaded = await loadPremadeWorkbook();
      const nextWorkbook = clearSaved ? loaded : restoreEdits(loaded);
      setWorkbook(nextWorkbook);
      setActiveSheet(nextWorkbook.sheetNames.includes(PREMADE_SHEET_NAME) ? PREMADE_SHEET_NAME : nextWorkbook.sheetNames[0]);
      setNotice(clearSaved ? "Hisaab reset to its original values." : "Hisaab workbook ready.");
    } catch (nextError) {
      setError(nextError instanceof Error ? nextError.message : "Failed to load workbook");
    } finally { setLoading(false); }
  }

  useEffect(() => { loadBundledWorkbook(); }, []);

  async function handleFile(event) {
    const file = event.target.files?.[0];
    if (!file) return;
    try {
      setLoading(true);
      setError("");
      const loaded = restoreEdits(await loadHisaabWorkbookFile(file));
      setWorkbook(loaded);
      setActiveSheet(loaded.sheetNames[0]);
      setNotice(`${file.name} opened with ${loaded.sheetNames.length} sheet(s).`);
    } catch (nextError) {
      setError(nextError instanceof Error ? nextError.message : "The workbook could not be opened in Hisaab.");
    } finally {
      setLoading(false);
      event.target.value = "";
    }
  }

  const sheet = getPremadeSheet(workbook, activeSheet);
  const formulaErrors = useMemo(() => sheet ? sheet.rows.flat().filter((cell) => cell.error).length : 0, [sheet]);

  function commitCell(sheetName, address, value) {
    setWorkbook((current) => {
      if (!current) return current;
      const next = updatePremadeWorkbookCell(current, sheetName, address, value);
      const saved = readSavedEdits();
      saved[current.title] = { ...(saved[current.title] || {}), [`${sheetName}!${address}`]: value };
      window.localStorage.setItem(HISAAB_EDITS_KEY, JSON.stringify(saved));
      return next;
    });
    setNotice(`${address} saved locally.`);
  }

  function downloadWorkbook() {
    if (!workbook) return;
    const exported = exportHisaabWorkbook(workbook);
    const url = URL.createObjectURL(exported.blob);
    const link = document.createElement("a");
    link.href = url;
    link.download = exported.fileName;
    link.click();
    URL.revokeObjectURL(url);
    setNotice(`${exported.fileName} downloaded.`);
  }

  return (
    <section className="panel hisaab-shell">
      <div className="panel-head">
        <div><div className="eyebrow">Editable workbook</div><div className="panel-title">{workbook?.title || "Hisaab"}</div><div className="panel-subtitle">Editable values are recovered on this browser. Formula cells are recalculated and marked with fx.</div></div>
        <div className="topbar-actions"><button className="ghost-button" onClick={() => fileInputRef.current?.click()} disabled={loading}>Open workbook</button><button className="primary-button" onClick={downloadWorkbook} disabled={!workbook || loading}>Download edited workbook</button><button className="ghost-button" onClick={() => loadBundledWorkbook({ clearSaved: true })} disabled={loading}>Reset Hisaab</button><input ref={fileInputRef} hidden type="file" accept=".xlsx,.xls" onChange={handleFile} /></div>
      </div>
      {error ? <div className="error-banner" role="alert"><span>{error}</span><button className="text-button" onClick={() => setError("")}>Dismiss</button></div> : null}
      {notice ? <div className="success-banner" role="status"><span>{notice}</span><button className="text-button" onClick={() => setNotice("")}>Dismiss</button></div> : null}
      {loading ? <div className="empty-state">Opening workbook...</div> : null}
      {!loading && workbook ? <>
        <div className="hisaab-toolbar"><div className="hisaab-tabs" role="tablist" aria-label="Workbook sheets">{workbook.sheetNames.map((name) => <button role="tab" aria-selected={activeSheet === name} key={name} className={activeSheet === name ? "listing-view-tab listing-view-tab-active" : "listing-view-tab"} onClick={() => setActiveSheet(name)}>{name}</button>)}</div><div className={formulaErrors ? "formula-status formula-status-error" : "formula-status"}>{formulaErrors ? `${formulaErrors} formula errors` : "Formulas calculated"}</div></div>
        {sheet ? <div className="table-scroll hisaab-table-scroll" tabIndex="0" aria-label={`${activeSheet} worksheet`}><table className="listing-table hisaab-table"><thead><tr><th className="hisaab-row-number">#</th>{sheet.columnLabels.map((label) => <th key={`${sheet.name}-${label}`}>{label}</th>)}</tr></thead><tbody>{sheet.rows.map((row, rowIndex) => <tr key={`${sheet.name}-row-${rowIndex}`}><th scope="row" className="hisaab-row-number">{rowIndex + 1}</th>{row.map((cell) => <WorkbookCell key={cell.address} sheetName={sheet.name} cell={cell} onCommit={commitCell} />)}</tr>)}</tbody></table></div> : <div className="empty-state">The selected sheet is unavailable.</div>}
      </> : null}
    </section>
  );
}
