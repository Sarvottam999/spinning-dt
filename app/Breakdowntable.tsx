"use client";

import React, { useMemo, useState } from "react";
import * as XLSX from "xlsx";

type CellValue = string | number | boolean | Date | null | undefined;
type RowData = Record<string, CellValue>;
type TableData = Record<string, Record<string, string[]>>;
type SummaryRow = Record<string, number>;

type ExcelUploaderProps = {
  title?: string;
  reportName?: string;
  extractLine?: (functionalLocation: string) => string;
  lineDetails?: Record<string, string[]>;
  lineMode?: "sequential" | "actual";
};

const columns = ["NGD", "BCK", "HRR", "VIL-1", "VIL-2", "IBR", "TRC"];

function extractBreakdownLine(functionalLocation: string) {
  const parts = functionalLocation.split("-");
  return parts[3] || "";
}

function mapPlantToUnit(plant: string) {
  const map: Record<string, string> = {
    "1101": "NGD",
    "1201": "BCK",
    "1301": "HRR",
    "1601": "VIL-1",
    "1602": "VIL-2",
    "2111": "IBR",
    "4101": "TRC",
  };
  return map[plant] || null;
}

function parseExcelDate(excelDate: CellValue): Date | null {
  if (!excelDate) return null;

  if (typeof excelDate === "string") {
    const [d, m, y] = excelDate.split("/");
    if (!d || !m || !y) return null;
    return new Date(`${y}-${m}-${d}`);
  }

  if (typeof excelDate !== "number") return null;

  const date = XLSX.SSF.parse_date_code(excelDate);
  if (!date) return null;
  return new Date(date.y, date.m - 1, date.d);
}

function formatExcelDate(excelDate: CellValue) {
  const d = parseExcelDate(excelDate);
  if (!d) return "";

  return `${String(d.getDate()).padStart(2, "0")}/${String(
    d.getMonth() + 1
  ).padStart(2, "0")}/${d.getFullYear()}`;
}

function toTitleCase(text: string) {
  return text
    .toLowerCase()
    .split(" ")
    .map((word) => word.charAt(0).toUpperCase() + word.slice(1))
    .join(" ");
}

function formatCell(row: RowData) {
  const date = formatExcelDate(row["Date"]);
  const dt = row["Total Down Time(Hrs)"];
  const sub = String(row["Sub Head reason"] || "");
  const loss = row["Loss Capacity"];

  return `${date} (${dt} Hrs.) – ${toTitleCase(sub)} (${loss} T) `;
}

function getPlantAndLine(
  functionalLocation: CellValue,
  extractLine: (functionalLocation: string) => string
) {
  const text = String(functionalLocation || "");
  const [plant = ""] = text.split("-");

  return {
    plant,
    line: extractLine(text),
  };
}

function getStringOptions(rows: RowData[], key: string) {
  return [
    ...new Set(rows.map((row) => row[key]).filter(Boolean).map(String)),
  ].sort((a, b) => a.localeCompare(b, undefined, { numeric: true }));
}

function sortLineKeys(keys: string[]) {
  return [...keys].sort((a, b) =>
    a.localeCompare(b, undefined, { numeric: true, sensitivity: "base" })
  );
}

function getSequentialLineKeys(tableData: TableData) {
  const maxLine = Math.max(
    0,
    ...Object.keys(tableData).map((line) => {
      const match = line.match(/\d+/);
      return match ? Number(match[0]) : 0;
    })
  );

  return Array.from(
    { length: maxLine },
    (_, index) => `L${String(index + 1).padStart(2, "0")}`
  );
}

function MultiCheckDropdown({
  label,
  options,
  selected,
  onChange,
}: {
  label: string;
  options: string[];
  selected: string[];
  onChange: (value: string[]) => void;
}) {
  const [open, setOpen] = useState(false);

  const toggle = (value: string) => {
    onChange(
      selected.includes(value)
        ? selected.filter((item) => item !== value)
        : [...selected, value]
    );
  };

  return (
    <div style={{ position: "relative" }}>
      <button
        onClick={() => setOpen((current) => !current)}
        style={{
          border: "1px solid #d1d5db",
          borderRadius: 4,
          padding: "5px 10px",
          fontSize: 12,
          background: "white",
          cursor: "pointer",
          minWidth: 130,
          display: "flex",
          justifyContent: "space-between",
          alignItems: "center",
          gap: 6,
        }}
      >
        <span>{selected.length ? `${label} (${selected.length})` : label}</span>
        <span>{open ? "▲" : "▼"}</span>
      </button>

      {open && (
        <div
          style={{
            position: "absolute",
            top: "110%",
            left: 0,
            zIndex: 999,
            background: "white",
            border: "1px solid #d1d5db",
            borderRadius: 6,
            boxShadow: "0 4px 12px rgba(0,0,0,0.12)",
            minWidth: 180,
            maxHeight: 220,
            overflowY: "auto",
            padding: "6px 0",
          }}
        >
          {options.map((option) => (
            <label
              key={option}
              style={{
                display: "flex",
                alignItems: "center",
                gap: 8,
                padding: "5px 12px",
                cursor: "pointer",
                fontSize: 12,
                background: selected.includes(option) ? "#fdf0ee" : "white",
              }}
            >
              <input
                type="checkbox"
                checked={selected.includes(option)}
                onChange={() => toggle(option)}
                style={{ accentColor: "#c0392b" }}
              />
              {option}
            </label>
          ))}
        </div>
      )}
    </div>
  );
}

export default function ExcelUploader({
  title = "Spinning Analysis",
  reportName = "breakdown",
  extractLine = extractBreakdownLine,
  lineDetails = {},
  lineMode = "sequential",
}: ExcelUploaderProps) {
  const [rawRows, setRawRows] = useState<RowData[]>([]);
  const [dateFrom, setDateFrom] = useState("");
  const [dateTo, setDateTo] = useState("");
  const [headReason, setHeadReason] = useState<string[]>([]);
  const [section, setSection] = useState<string[]>([]);
  const [subHeadReason, setSubHeadReason] = useState<string[]>([]);
  const [reasonDesc, setReasonDesc] = useState<string[]>([]);
  const [lineFilter, setLineFilter] = useState<string[]>([]);
  const [unitFilter, setUnitFilter] = useState<string[]>([]);

  const headReasonOptions = useMemo(
    () => getStringOptions(rawRows, "Head reason"),
    [rawRows]
  );
  const sectionOptions = useMemo(
    () => getStringOptions(rawRows, "Section"),
    [rawRows]
  );
  const subHeadOptions = useMemo(
    () => getStringOptions(rawRows, "Sub Head reason"),
    [rawRows]
  );
  const reasonDescOptions = useMemo(
    () => getStringOptions(rawRows, "Reason Desc"),
    [rawRows]
  );
  const lineOptions = useMemo(() => {
    const lines = rawRows
      .map((row) => getPlantAndLine(row["Functional Location"], extractLine).line)
      .filter(Boolean);

    return sortLineKeys([...new Set(lines)]);
  }, [rawRows, extractLine]);
  const unitOptions = useMemo(() => {
    const units = rawRows
      .map((row) => {
        const { plant } = getPlantAndLine(row["Functional Location"], extractLine);
        return mapPlantToUnit(plant);
      })
      .filter((unit): unit is string => Boolean(unit));

    return columns.filter((column) => units.includes(column));
  }, [rawRows, extractLine]);
  const visibleColumns = unitFilter.length
    ? columns.filter((column) => unitFilter.includes(column))
    : columns;

  const filteredRows = useMemo(() => {
    return rawRows.filter((row) => {
      const { plant, line } = getPlantAndLine(row["Functional Location"], extractLine);
      const unit = mapPlantToUnit(plant);

      if (lineFilter.length && !lineFilter.includes(line)) {
        return false;
      }

      if (unitFilter.length && (!unit || !unitFilter.includes(unit))) {
        return false;
      }

      if (
        headReason.length &&
        !headReason.includes(String(row["Head reason"] || ""))
      ) {
        return false;
      }

      if (section.length && !section.includes(String(row["Section"] || ""))) {
        return false;
      }

      if (
        subHeadReason.length &&
        !subHeadReason.includes(String(row["Sub Head reason"] || ""))
      ) {
        return false;
      }

      if (
        reasonDesc.length &&
        !reasonDesc.includes(String(row["Reason Desc"] || ""))
      ) {
        return false;
      }

      if (dateFrom || dateTo) {
        const date = parseExcelDate(row["Date"]);
        if (!date) return false;
        if (dateFrom && date < new Date(dateFrom)) return false;
        if (dateTo && date > new Date(dateTo)) return false;
      }

      return true;
    });
  }, [
    rawRows,
    extractLine,
    lineFilter,
    unitFilter,
    headReason,
    section,
    subHeadReason,
    reasonDesc,
    dateFrom,
    dateTo,
  ]);

  const tableData = useMemo(() => {
    const result: TableData = {};

    filteredRows.forEach((row) => {
      const { plant, line } = getPlantAndLine(
        row["Functional Location"],
        extractLine
      );
      const unit = mapPlantToUnit(plant);

      if (!unit || !line) return;

      if (!result[line]) {
        result[line] = {};
        columns.forEach((column) => {
          result[line][column] = [];
        });
      }

      result[line][unit].push(formatCell(row));
    });

    return result;
  }, [filteredRows, extractLine]);

  const { freqRow, hrsRow } = useMemo(() => {
    const freq: SummaryRow = {};
    const hrs: SummaryRow = {};

    columns.forEach((column) => {
      freq[column] = 0;
      hrs[column] = 0;
    });

    filteredRows.forEach((row) => {
      const { plant } = getPlantAndLine(row["Functional Location"], extractLine);
      const unit = mapPlantToUnit(plant);

      if (!unit) return;

      freq[unit]++;
      hrs[unit] += parseFloat(String(row["Total Down Time(Hrs)"] || "")) || 0;
    });

    return { freqRow: freq, hrsRow: hrs };
  }, [filteredRows, extractLine]);

  const lineKeys = useMemo(() => {
    if (lineFilter.length) {
      return sortLineKeys(Object.keys(tableData));
    }

    if (lineMode === "sequential") {
      return getSequentialLineKeys(tableData);
    }

    return sortLineKeys(Object.keys(tableData));
  }, [tableData, lineMode, lineFilter]);

  const resetFilters = () => {
    setHeadReason([]);
    setSection([]);
    setSubHeadReason([]);
    setReasonDesc([]);
    setLineFilter([]);
    setUnitFilter([]);
    setDateFrom("");
    setDateTo("");
  };

  const handleFile = (file: File) => {
    const reader = new FileReader();
    reader.onload = (event) => {
      const wb = XLSX.read(event.target?.result, { type: "binary" });
      const sheet = wb.Sheets[wb.SheetNames[0]];
      setRawRows(XLSX.utils.sheet_to_json<RowData>(sheet));
      resetFilters();
    };
    reader.readAsBinaryString(file);
  };

  const handleDrop = (event: React.DragEvent<HTMLDivElement>) => {
    event.preventDefault();
    const file = event.dataTransfer.files[0];
    if (file) handleFile(file);
  };

  const inputStyle: React.CSSProperties = {
    border: "1px solid #d1d5db",
    borderRadius: 4,
    padding: "4px 8px",
    fontSize: 12,
    background: "white",
  };

  function buildExportRows() {
    const rows: string[][] = [];
    rows.push(["Unit", ...visibleColumns]);

    lineKeys.forEach((line) => {
      const row = tableData[line];
      const maxLen = Math.max(
        ...visibleColumns.map((column) => row?.[column]?.length || 0)
      );

      for (let index = 0; index < Math.max(1, maxLen); index++) {
        rows.push([
          index === 0
            ? [line, ...(lineDetails[line] || [])].join("\n")
            : "",
          ...visibleColumns.map((column) => row?.[column]?.[index] || ""),
        ]);
      }
    });

    rows.push([
      "DT – (Freq.)",
      ...visibleColumns.map((column) => String(freqRow[column] || "")),
    ]);
    rows.push([
      "DT – (Hrs.)",
      ...visibleColumns.map((column) =>
        hrsRow[column] ? hrsRow[column].toFixed(2) : ""
      ),
    ]);

    return rows;
  }

  function exportToExcel() {
    const ws = XLSX.utils.aoa_to_sheet(buildExportRows());
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, reportName);
    XLSX.writeFile(wb, `${reportName}_report.xlsx`);
  }

  function copyAsExcel() {
    const text = buildExportRows()
      .map((row) => row.join("\t"))
      .join("\n");

    navigator.clipboard
      .writeText(text)
      .then(() => alert("Copied! Paste directly into Excel."));
  }

  return (
    <div
      style={{
        fontFamily: "Arial, sans-serif",
        minHeight: "100vh",
        background: "#f5f5f5",
        width: "100%",
      }}
    >
      <div
        style={{
          margin: "24px auto",
          padding: "0 16px",
        }}
      >
        <h1
          style={{
            fontSize: 22,
            fontWeight: 700,
            color: "#333",
            marginBottom: 16,
          }}
        >
          {title}
        </h1>

        {rawRows.length === 0 && (
          <div
            onDrop={handleDrop}
            onDragOver={(event) => event.preventDefault()}
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
              const input = document.createElement("input");
              input.type = "file";
              input.accept = ".xlsx,.xls";
              input.onchange = (event) => {
                const target = event.target as HTMLInputElement;
                if (target.files?.[0]) handleFile(target.files[0]);
              };
              input.click();
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
                  resetFilters();
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
                Upload New Excel
              </button>
            </div>

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

              <MultiCheckDropdown
                label="Head Reason"
                options={headReasonOptions}
                selected={headReason}
                onChange={setHeadReason}
              />
              <MultiCheckDropdown
                label="Section"
                options={sectionOptions}
                selected={section}
                onChange={setSection}
              />
              <MultiCheckDropdown
                label="Sub Head Reason"
                options={subHeadOptions}
                selected={subHeadReason}
                onChange={setSubHeadReason}
              />
              <MultiCheckDropdown
                label="Reason Desc"
                options={reasonDescOptions}
                selected={reasonDesc}
                onChange={setReasonDesc}
              />
              <MultiCheckDropdown
                label="Line"
                options={lineOptions}
                selected={lineFilter}
                onChange={setLineFilter}
              />
              <MultiCheckDropdown
                label="Unit"
                options={unitOptions}
                selected={unitFilter}
                onChange={setUnitFilter}
              />

              <div style={{ display: "flex", alignItems: "center", gap: 4 }}>
                <span style={{ fontSize: 12, color: "#666" }}>From:</span>
                <input
                  type="date"
                  style={inputStyle}
                  value={dateFrom}
                  onChange={(event) => setDateFrom(event.target.value)}
                />
              </div>
              <div style={{ display: "flex", alignItems: "center", gap: 4 }}>
                <span style={{ fontSize: 12, color: "#666" }}>To:</span>
                <input
                  type="date"
                  style={inputStyle}
                  value={dateTo}
                  onChange={(event) => setDateTo(event.target.value)}
                />
              </div>

              <button
                onClick={resetFilters}
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

              <div style={{ marginLeft: "auto", display: "flex", gap: 8 }}>
                <button
                  onClick={exportToExcel}
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
                  onClick={copyAsExcel}
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
                    {visibleColumns.map((column) => (
                      <th
                        key={column}
                        style={{
                          border: "1px solid #ddd",
                          padding: "8px 12px",
                          textAlign: "center",
                          fontWeight: "bold",
                        }}
                      >
                        {column}
                      </th>
                    ))}
                  </tr>
                </thead>
                <tbody>
                  {lineKeys.map((lineKey, index) => {
                    const row = tableData[lineKey];

                    return (
                      <tr
                        key={lineKey}
                        style={{
                          background: index % 2 === 0 ? "white" : "#fdf6f5",
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
                          <div>{lineKey}</div>
                          {lineDetails[lineKey]?.map((detail) => (
                            <div
                              key={detail}
                              style={{
                                color: "#666",
                                fontSize: 11,
                                fontWeight: 400,
                                marginTop: 3,
                                whiteSpace: "nowrap",
                              }}
                            >
                              {detail}
                            </div>
                          ))}
                        </td>
                        {visibleColumns.map((column) => (
                          <td
                            key={column}
                            style={{
                              border: "1px solid #ddd",
                              padding: "6px 12px",
                              verticalAlign: "top",
                            }}
                          >
                            {row?.[column]?.map((item, itemIndex) => (
                              <div key={itemIndex} style={{ marginBottom: 2 }}>
                                {item}
                              </div>
                            ))}
                          </td>
                        ))}
                      </tr>
                    );
                  })}

                  <tr style={{ background: "#f0d6f5", fontWeight: "bold" }}>
                    <td
                      style={{ border: "1px solid #ddd", padding: "6px 12px" }}
                    >
                      DT – (Freq.)
                    </td>
                    {visibleColumns.map((column) => (
                      <td
                        key={column}
                        style={{
                          border: "1px solid #ddd",
                          padding: "6px 12px",
                          textAlign: "center",
                        }}
                      >
                        {freqRow[column] || ""}
                      </td>
                    ))}
                  </tr>

                  <tr style={{ background: "#f0d6f5", fontWeight: "bold" }}>
                    <td
                      style={{ border: "1px solid #ddd", padding: "6px 12px" }}
                    >
                      DT – (Hrs.)
                    </td>
                    {visibleColumns.map((column) => (
                      <td
                        key={column}
                        style={{
                          border: "1px solid #ddd",
                          padding: "6px 12px",
                          textAlign: "center",
                        }}
                      >
                        {hrsRow[column] ? hrsRow[column].toFixed(2) : ""}
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
