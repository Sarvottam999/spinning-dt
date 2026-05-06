"use client";
import React, { useState, useMemo } from "react";
import * as XLSX from "xlsx";

type RowData = Record<string, any>;
const columns = ["NGD", "BCK", "HRR", "VIL-1", "VIL-2", "IBR", "TRC"];

export default function ExcelUploader() {
  const [rawRows, setRawRows] = useState<RowData[]>([]);

  // Filter states
//   const [headReason, setHeadReason] = useState("");
//   const [section, setSection] = useState("");
//   const [department, setDepartment] = useState("");
//   const [subHeadReason, setSubHeadReason] = useState("");
  const [dateFrom, setDateFrom] = useState("");
  const [dateTo, setDateTo] = useState("");
//   const [lossCapacity, setLossCapacity] = useState("");

  const [headReason, setHeadReason] = useState<string[]>([]);
const [section, setSection] = useState<string[]>([]);
const [department, setDepartment] = useState<string[]>([]);
const [subHeadReason, setSubHeadReason] = useState<string[]>([]);
const [lossCapacity, setLossCapacity] = useState<string[]>([]);

  function extractPlantAndLine(fl: string) {
    const parts = fl?.split("-") || [];
    return { plant: parts[0] || "", line: parts[3] || "" };
  }

  function mapPlantToUnit(plant: string) {
    const map: Record<string, string> = {
      "1101": "NGD",
      "1201": "BCK",
      "2101": "HRR",
      "3101": "VIL-1",
      "3201": "VIL-2",
      "4101": "IBR",
      "5101": "TRC",
    };
    return map[plant] || null;
  }

  function parseExcelDate(excelDate: any): Date | null {
    if (!excelDate) return null;
    if (typeof excelDate === "string") {
      const [d, m, y] = excelDate.split("/");
      return new Date(`${y}-${m}-${d}`);
    }
    const date = XLSX.SSF.parse_date_code(excelDate);
    if (!date) return null;
    return new Date(date.y, date.m - 1, date.d);
  }

  function formatExcelDate(excelDate: any): string {
    const d = parseExcelDate(excelDate);
    if (!d) return "";
    return `${String(d.getDate()).padStart(2, "0")}/${String(
      d.getMonth() + 1
    ).padStart(2, "0")}/${d.getFullYear()}`;
  }

  function formatCell(row: RowData) {
    const date = formatExcelDate(row["Date"]);
    const dt = row["Total Down Time(Hrs)"];
    const sub = row["Sub Head reason"];
    const loss = row["Loss Capacity"];
    return `${date} (${dt}) – ${sub} – ${loss}`;
  }

  // Unique options for dropdowns
  const headReasonOptions = useMemo(
    () => [...new Set(rawRows.map((r) => r["Head reason"]).filter(Boolean))],
    [rawRows]
  );
  const sectionOptions = useMemo(
    () => [...new Set(rawRows.map((r) => r["Section"]).filter(Boolean))],
    [rawRows]
  );
  const departmentOptions = useMemo(
    () => [...new Set(rawRows.map((r) => r["Department"]).filter(Boolean))],
    [rawRows]
  );
  const subHeadOptions = useMemo(
    () => [
      ...new Set(rawRows.map((r) => r["Sub Head reason"]).filter(Boolean)),
    ],
    [rawRows]
  );
  const lossOptions = useMemo(
    () => [...new Set(rawRows.map((r) => r["Loss Capacity"]).filter(Boolean))],
    [rawRows]
  );

  const filteredRows = useMemo(() => {
    return rawRows.filter((row) => {
    //   if (headReason && row["Head reason"] !== headReason) return false;
    //   if (section && row["Section"] !== section) return false;
    //   if (department && row["Department"] !== department) return false;
    //   if (subHeadReason && row["Sub Head reason"] !== subHeadReason)
    //     return false;
    //   if (lossCapacity && String(row["Loss Capacity"]) !== lossCapacity)
    //     return false;
    if (headReason.length && !headReason.includes(row["Head reason"])) return false;
if (section.length && !section.includes(row["Section"])) return false;
if (department.length && !department.includes(row["Department"])) return false;
if (subHeadReason.length && !subHeadReason.includes(row["Sub Head reason"])) return false;
if (lossCapacity.length && !lossCapacity.includes(String(row["Loss Capacity"]))) return false;
      if (dateFrom || dateTo) {
        const d = parseExcelDate(row["Date"]);
        if (!d) return false;
        if (dateFrom && d < new Date(dateFrom)) return false;
        if (dateTo && d > new Date(dateTo)) return false;
      }
      return true;
    });
  }, [
    rawRows,
    headReason,
    section,
    department,
    subHeadReason,
    lossCapacity,
    dateFrom,
    dateTo,
  ]);

  function processData(rows: RowData[]) {
    const result: any = {};
    rows.forEach((row) => {
      const { plant, line } = extractPlantAndLine(row["Functional Location"]);
      const unit = mapPlantToUnit(plant);
      if (!unit || !line) return;
      if (!result[line]) {
        result[line] = {};
        columns.forEach((c) => (result[line][c] = []));
      }
      result[line][unit].push(formatCell(row));
    });
    return result;
  }

  const tableData = useMemo(() => processData(filteredRows), [filteredRows]);

  // Frequency and hours summaries
  const freqRow: Record<string, number> = {};
  const hrsRow: Record<string, number> = {};
  columns.forEach((col) => {
    freqRow[col] = 0;
    hrsRow[col] = 0;
    filteredRows.forEach((row) => {
      const { plant } = extractPlantAndLine(row["Functional Location"]);
      const unit = mapPlantToUnit(plant);
      if (unit === col) {
        freqRow[col]++;
        hrsRow[col] += parseFloat(row["Total Down Time(Hrs)"]) || 0;
      }
    });
  });

  const handleFile = (file: File) => {
    const reader = new FileReader();
    reader.onload = (e) => {
      const wb = XLSX.read(e.target?.result, { type: "binary" });
      const sheet = wb.Sheets[wb.SheetNames[0]];
      setRawRows(XLSX.utils.sheet_to_json(sheet));
    };
    reader.readAsBinaryString(file);
  };

  const handleDrop = (e: React.DragEvent<HTMLDivElement>) => {
    e.preventDefault();
    const file = e.dataTransfer.files[0];
    if (file) handleFile(file);
  };

  const selectStyle: React.CSSProperties = {
    border: "1px solid #d1d5db",
    borderRadius: 4,
    padding: "4px 8px",
    fontSize: 12,
    background: "white",
    minWidth: 120,
  };
  const inputStyle: React.CSSProperties = {
    border: "1px solid #d1d5db",
    borderRadius: 4,
    padding: "4px 8px",
    fontSize: 12,
    background: "white",
  };

  function exportToExcel(tableData: any, freqRow: any, hrsRow: any) {
    const wsData: any[][] = [];
    wsData.push(["Unit", ...columns]);
    Object.entries(tableData).forEach(([line, row]: any) => {
      const maxLen = Math.max(...columns.map((c) => row[c]?.length || 0));
      for (let i = 0; i < Math.max(1, maxLen); i++) {
        wsData.push([
          i === 0 ? line : "",
          ...columns.map((c) => row[c]?.[i] || ""),
        ]);
      }
    });
    wsData.push(["DT – (Freq.)", ...columns.map((c) => freqRow[c] || "")]);
    wsData.push([
      "DT – (Hrs.)",
      ...columns.map((c) => (hrsRow[c] ? hrsRow[c].toFixed(2) : "")),
    ]);

    const ws = XLSX.utils.aoa_to_sheet(wsData);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Breakdown");
    XLSX.writeFile(wb, "breakdown_report.xlsx");
  }

  function copyAsExcel(tableData: any, freqRow: any, hrsRow: any) {
    const rows: string[][] = [];
    rows.push(["Unit", ...columns]);
    Object.entries(tableData).forEach(([line, row]: any) => {
      const maxLen = Math.max(...columns.map((c) => row[c]?.length || 0));
      for (let i = 0; i < Math.max(1, maxLen); i++) {
        rows.push([
          i === 0 ? line : "",
          ...columns.map((c) => row[c]?.[i] || ""),
        ]);
      }
    });
    rows.push(["DT – (Freq.)", ...columns.map((c) => freqRow[c] || "")]);
    rows.push([
      "DT – (Hrs.)",
      ...columns.map((c) => (hrsRow[c] ? hrsRow[c].toFixed(2) : "")),
    ]);

    const text = rows.map((r) => r.join("\t")).join("\n");
    navigator.clipboard
      .writeText(text)
      .then(() => alert("Copied! Paste directly into Excel."));
  }

  function MultiCheckDropdown({ label, options, selected, onChange }: {
    label: string;
    options: string[];
    selected: string[];
    onChange: (val: string[]) => void;
  }) {
    const [open, setOpen] = useState(false);
  
    const toggle = (val: string) => {
      onChange(selected.includes(val) ? selected.filter(v => v !== val) : [...selected, val]);
    };
  
    return (
      <div style={{ position: "relative" }}>
        <button
          onClick={() => setOpen(o => !o)}
          style={{
            border: "1px solid #d1d5db", borderRadius: 4, padding: "5px 10px",
            fontSize: 12, background: "white", cursor: "pointer", minWidth: 130,
            display: "flex", justifyContent: "space-between", alignItems: "center", gap: 6,
          }}
        >
          <span>{selected.length ? `${label} (${selected.length})` : label}</span>
          <span>{open ? "▲" : "▼"}</span>
        </button>
  
        {open && (
          <div style={{
            position: "absolute", top: "110%", left: 0, zIndex: 999,
            background: "white", border: "1px solid #d1d5db", borderRadius: 6,
            boxShadow: "0 4px 12px rgba(0,0,0,0.12)", minWidth: 180, maxHeight: 220,
            overflowY: "auto", padding: "6px 0",
          }}>
            {options.map(o => (
              <label key={o} style={{
                display: "flex", alignItems: "center", gap: 8,
                padding: "5px 12px", cursor: "pointer", fontSize: 12,
                background: selected.includes(o) ? "#fdf0ee" : "white",
              }}>
                <input
                  type="checkbox"
                  checked={selected.includes(o)}
                  onChange={() => toggle(o)}
                  style={{ accentColor: "#c0392b" }}
                />
                {o}
              </label>
            ))}
          </div>
        )}
      </div>
    );
  }

  return (
    <div
      style={{
        fontFamily: "Arial, sans-serif",
        minHeight: "100vh",
        background: "#f5f5f5",
      }}
    >
      {/* Header */}
      {/* <div style={{
        background: "linear-gradient(90deg, #c0392b 60%, #e67e22 100%)",
        padding: "14px 24px", display: "flex", alignItems: "center", justifyContent: "space-between"
      }}>
       
        <div style={{
          background: "white", borderRadius: 6, padding: "4px 10px",
          display: "flex", flexDirection: "column", alignItems: "center"
        }}>
          <span style={{ color: "#c0392b", fontWeight: "bold", fontSize: 13, letterSpacing: 1 }}>ADITYA BIRLA</span>
          <span style={{ color: "#e67e22", fontWeight: "bold", fontSize: 16, letterSpacing: 2 }}>GRASIM</span>
        </div>
      </div> */}

      <div
        style={{
          margin: "24px auto",
          padding: "0 16px",
        }}
      >
        {/* Upload */}
        {rawRows.length === 0 && (
          <div
            onDrop={handleDrop}
            onDragOver={(e) => e.preventDefault()}
            style={{
              border: "2px dashed #c0392b",
              borderRadius: 8,
              padding: 40,
              textAlign: "center",
              color: "#c0392b",
              marginBottom: 16,
              background: "white",
              cursor: "pointer",
            }}
            onClick={() => {
              const inp = document.createElement("input");
              inp.type = "file";
              inp.accept = ".xlsx,.xls";
              inp.onchange = (e: any) => {
                if (e.target.files[0]) handleFile(e.target.files[0]);
              };
              inp.click();
            }}
          >
            Drag & Drop Excel File or Click to Upload
          </div>
        )}

        {rawRows.length > 0 && (
          <>
            <div
              style={{
                display: "flex",
                justifyContent: "flex-end",
                marginBottom: 10,
              }}
            >
              <button
                onClick={() => {
                  setRawRows([]);
                //   setHeadReason("");
                //   setSection("");
                //   setDepartment("");
                //   setSubHeadReason("");
                //   setLossCapacity("");

                  setDateFrom("");
                  setDateTo("");
                }}
                style={{
                  background: "#c0392b",
                  color: "white",
                  border: "none",
                  borderRadius: 4,
                  padding: "6px 16px",
                  fontSize: 13,
                  cursor: "pointer",
                }}
              >
                {" "}
                Upload New Excel
              </button>
            </div>
            {/* Filters */}
            <div
              style={{
                background: "white",
                borderRadius: 8,
                padding: "12px 16px",
                marginBottom: 14,
                display: "flex",
                flexWrap: "wrap",
                gap: 10,
                alignItems: "center",
                boxShadow: "0 1px 4px rgba(0,0,0,0.07)",
              }}
            >
              <strong
                style={{ color: "#c0392b", fontSize: 13, marginRight: 4 }}
              >
                Filters:
              </strong>

              <MultiCheckDropdown label="Head Reason" options={headReasonOptions as string[]} selected={headReason} onChange={setHeadReason} />
<MultiCheckDropdown label="Section" options={sectionOptions as string[]} selected={section} onChange={setSection} />
<MultiCheckDropdown label="Sub Head Reason" options={subHeadOptions as string[]} selected={subHeadReason} onChange={setSubHeadReason} />
            {/* <select multiple style={{...selectStyle, height: 60, minWidth: 130}} value={headReason}
  onChange={e => setHeadReason([...e.target.selectedOptions].map(o => o.value))}>
  {headReasonOptions.map(o => <option key={o} value={o}>{o}</option>)}
</select> */}

{/* <select multiple style={{...selectStyle, height: 60, minWidth: 130}} value={section}
  onChange={e => setSection([...e.target.selectedOptions].map(o => o.value))}>
  {sectionOptions.map(o => <option key={o} value={o}>{o}</option>)}
</select> */}

{/* <select multiple style={{...selectStyle, height: 60, minWidth: 130}} value={department}
  onChange={e => setDepartment([...e.target.selectedOptions].map(o => o.value))}>
  {departmentOptions.map(o => <option key={o} value={o}>{o}</option>)}
</select> */}

{/* <select multiple style={{...selectStyle, height: 60, minWidth: 130}} value={subHeadReason}
  onChange={e => setSubHeadReason([...e.target.selectedOptions].map(o => o.value))}>
  {subHeadOptions.map(o => <option key={o} value={o}>{o}</option>)}
</select> */}

{/* <select multiple style={{...selectStyle, height: 60, minWidth: 130}} value={lossCapacity}
  onChange={e => setLossCapacity([...e.target.selectedOptions].map(o => o.value))}>
  {lossOptions.map(o => <option key={o} value={o}>{o}</option>)}
</select> */}

              <div style={{ display: "flex", alignItems: "center", gap: 4 }}>
                <span style={{ fontSize: 12, color: "#666" }}>From:</span>
                <input
                  type="date"
                  style={inputStyle}
                  value={dateFrom}
                  onChange={(e) => setDateFrom(e.target.value)}
                />
              </div>
              <div style={{ display: "flex", alignItems: "center", gap: 4 }}>
                <span style={{ fontSize: 12, color: "#666" }}>To:</span>
                <input
                  type="date"
                  style={inputStyle}
                  value={dateTo}
                  onChange={(e) => setDateTo(e.target.value)}
                />
              </div>

              <button
             onClick={() => { setHeadReason([]); setSection([]); setDepartment([]); setSubHeadReason([]); setLossCapacity([]); setDateFrom(""); setDateTo(""); }}
                style={{
                  background: "#c0392b",
                  color: "white",
                  border: "none",
                  borderRadius: 4,
                  padding: "4px 12px",
                  fontSize: 12,
                  cursor: "pointer",
                }}
              >
                Clear
              </button>

              {/* spacer + buttons */}
              <div style={{ marginLeft: "auto", display: "flex", gap: 8 }}>
                <button
                  onClick={() => exportToExcel(tableData, freqRow, hrsRow)}
                  style={{
                    background: "#1e7e34",
                    color: "white",
                    border: "none",
                    borderRadius: 4,
                    padding: "4px 12px",
                    fontSize: 12,
                    cursor: "pointer",
                  }}
                >
                  Export Excel
                </button>

                <button
                  onClick={() => copyAsExcel(tableData, freqRow, hrsRow)}
                  style={{
                    background: "#0056b3",
                    color: "white",
                    border: "none",
                    borderRadius: 4,
                    padding: "4px 12px",
                    fontSize: 12,
                    cursor: "pointer",
                  }}
                >
                  Copy as Excel
                </button>
              </div>
            </div>

            {/* Table */}
            <div
              style={{
                overflowX: "auto",
                background: "white",
                borderRadius: 8,
                boxShadow: "0 1px 4px rgba(0,0,0,0.07)",
              }}
            >
              <table
                style={{
                  minWidth: "100%",
                  borderCollapse: "collapse",
                  fontSize: 12,
                }}
              >
                <thead>
                  <tr style={{ background: "#f2e0dc" }}>
                    <th
                      style={{
                        border: "1px solid #ddd",
                        padding: "8px 12px",
                        textAlign: "left",
                        fontWeight: "bold",
                      }}
                    >
                      Unit
                    </th>
                    {columns.map((col) => (
                      <th
                        key={col}
                        style={{
                          border: "1px solid #ddd",
                          padding: "8px 12px",
                          textAlign: "center",
                          fontWeight: "bold",
                        }}
                      >
                        {col}
                      </th>
                    ))}
                  </tr>
                </thead>
                <tbody>
                  {Object.entries(tableData).map(([line, row]: any, idx) => (
                    <tr
                      key={line}
                      style={{
                        background: idx % 2 === 0 ? "white" : "#fdf6f5",
                      }}
                    >
                      <td
                        style={{
                          border: "1px solid #ddd",
                          padding: "6px 12px",
                          fontWeight: "bold",
                          color: "#333",
                        }}
                      >
                        {line}
                      </td>
                      {columns.map((col) => (
                        <td
                          key={col}
                          style={{
                            border: "1px solid #ddd",
                            padding: "6px 12px",
                            verticalAlign: "top",
                          }}
                        >
                          {row[col]?.map((item: string, i: number) => (
                            <div
                              key={i}
                              style={{
                                marginBottom: 2,
                                fontWeight: i === 0 ? "600" : "normal",
                              }}
                            >
                              {item}
                            </div>
                          ))}
                        </td>
                      ))}
                    </tr>
                  ))}

                  {/* DT Frequency row */}
                  <tr style={{ background: "#f0d6f5", fontWeight: "bold" }}>
                    <td
                      style={{ border: "1px solid #ddd", padding: "6px 12px" }}
                    >
                      DT – (Freq.)
                    </td>
                    {columns.map((col) => (
                      <td
                        key={col}
                        style={{
                          border: "1px solid #ddd",
                          padding: "6px 12px",
                          textAlign: "center",
                        }}
                      >
                        {freqRow[col] || ""}
                      </td>
                    ))}
                  </tr>

                  {/* DT Hours row */}
                  <tr style={{ background: "#f0d6f5", fontWeight: "bold" }}>
                    <td
                      style={{ border: "1px solid #ddd", padding: "6px 12px" }}
                    >
                      DT – (Hrs.)
                    </td>
                    {columns.map((col) => (
                      <td
                        key={col}
                        style={{
                          border: "1px solid #ddd",
                          padding: "6px 12px",
                          textAlign: "center",
                        }}
                      >
                        {hrsRow[col] ? hrsRow[col].toFixed(2) : ""}
                      </td>
                    ))}
                  </tr>
                </tbody>
              </table>
            </div>

            <div style={{ marginTop: 8, fontSize: 11, color: "#999" }}>
              Sensitivity: General
            </div>
          </>
        )}
      </div>
    </div>
  );
}
