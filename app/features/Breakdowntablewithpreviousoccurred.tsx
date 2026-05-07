"use client";

import React, { useEffect, useMemo, useState } from "react";
import * as XLSX from "xlsx";

type CellValue = string | number | boolean | Date | null | undefined;
type RowData = Record<string, CellValue>;
type TableData = Record<string, Record<string, FormattedCell[]>>;
type SummaryRow = Record<string, number>;
type FilterName =
  | "headReason"
  | "section"
  | "subHeadReason"
  | "reasonDesc"
  | "line"
  | "unit";

/** Enriched cell: current entry + optional previous occurrence */
type FormattedCell = {
  current: string;
  previous: string | null; // null = no previous occurrence
};

type ExcelUploaderProps = {
  title?: string;
  reportName?: string;
  extractLine?: (functionalLocation: string) => string;
  lineDetails?: Record<string, string[]>;
  lineMode?: "sequential" | "actual";
};

const columns = ["NGD", "BCK", "HRR", "VIL-1", "VIL-2", "IBR", "TRC"];

// ─── helpers ────────────────────────────────────────────────────────────────

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

/** Format the "current" part of the cell (same as original) */
function formatCurrentCell(row: RowData) {
  const date = formatExcelDate(row["Date"]);
  const dt = row["Total Down Time(Hrs)"];
  const sub = String(row["Sub Head reason"] || "");
  const loss = row["Loss Capacity"];
  return `${date} (${dt} Hrs.) – ${toTitleCase(sub)} (${loss} T)`;
}

function getPlantAndLine(
  functionalLocation: CellValue,
  extractLine: (fl: string) => string
) {
  const text = String(functionalLocation || "");
  const [plant = ""] = text.split("-");
  return { plant, line: extractLine(text) };
}

function getStringOptions(rows: RowData[], key: string) {
  return [
    ...new Set(rows.map((row) => row[key]).filter(Boolean).map(String)),
  ].sort((a, b) => a.localeCompare(b, undefined, { numeric: true }));
}

function mergeStringOptions(options: string[], selected: string[]) {
  return [...new Set([...options, ...selected])].sort((a, b) =>
    a.localeCompare(b, undefined, { numeric: true })
  );
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
    (_, i) => `L${String(i + 1).padStart(2, "0")}`
  );
}

// ─── MultiCheckDropdown ──────────────────────────────────────────────────────

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
        onClick={() => setOpen((c) => !c)}
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

// ─── PreviousOccurrenceTag ───────────────────────────────────────────────────

/**
 * Small inline badge showing "Prev: DD/MM/YYYY (X Hrs.) – SubReason (Y T)"
 * Styled differently so it's visually distinct from the current entry.
 */
function PreviousOccurrenceTag({ text }: { text: string }) {
  return (
    <div
      style={{
        display: "inline-flex",
        alignItems: "center",
        gap: 4,
        marginTop: 3,
        padding: "2px 7px",
        borderRadius: 3,
        background: "#fff7e6",
        border: "1px solid #f0b429",
        fontSize: 10.5,
        color: "#7a4f00",
        lineHeight: 1.4,
      }}
    >
      {/* small clock icon via unicode */}
      <span style={{ fontSize: 10, opacity: 0.7 }}>🕐</span>
      <span>
        <strong style={{ color: "#b36b00" }}>Prev:</strong> {text}
      </span>
    </div>
  );
}

// ─── Main component ──────────────────────────────────────────────────────────

export default function BreakdownTableWithPreviousOccurred({
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

  // ── Build a lookup: for each (line, unit, subHeadReason) key → sorted dates
  // This is built from ALL raw rows (unfiltered) so we can always find a
  // previous occurrence even if it falls outside the current filter window.
  const previousOccurrenceMap = useMemo(() => {
    /**
     * Key: `${line}||${unit}||${subHeadReasonNormalized}`
     * Value: array of { date: Date; formatted: string } sorted ascending by date
     */
    const map = new Map<string, { date: Date; formatted: string }[]>();

    rawRows.forEach((row) => {
      const { plant, line } = getPlantAndLine(
        row["Functional Location"],
        extractLine
      );
      const unit = mapPlantToUnit(plant);
      if (!unit || !line) return;

      const sub = String(row["Sub Head reason"] || "").trim().toLowerCase();
      if (!sub) return;

      const date = parseExcelDate(row["Date"]);
      if (!date) return;

      const key = `${line}||${unit}||${sub}`;
      const entry = {
        date,
        formatted: formatCurrentCell(row),
      };

      if (!map.has(key)) {
        map.set(key, []);
      }
      map.get(key)!.push(entry);
    });

    // Sort each bucket ascending by date
    map.forEach((entries) => {
      entries.sort((a, b) => a.date.getTime() - b.date.getTime());
    });

    return map;
  }, [rawRows, extractLine]);

  /**
   * Given a current row, find the most-recent occurrence of the same
   * (line, unit, subHeadReason) that happened STRICTLY BEFORE this row's date.
   * Returns the formatted string or null.
   */
  const findPreviousOccurrence = (row: RowData): string | null => {
    const { plant, line } = getPlantAndLine(
      row["Functional Location"],
      extractLine
    );
    const unit = mapPlantToUnit(plant);
    if (!unit || !line) return null;

    const sub = String(row["Sub Head reason"] || "").trim().toLowerCase();
    if (!sub) return null;

    const currentDate = parseExcelDate(row["Date"]);
    if (!currentDate) return null;

    const key = `${line}||${unit}||${sub}`;
    const bucket = previousOccurrenceMap.get(key);
    if (!bucket) return null;

    // Find all entries strictly before currentDate, pick the latest one
    const earlier = bucket.filter(
      (e) => e.date.getTime() < currentDate.getTime()
    );
    if (earlier.length === 0) return null;

    return earlier[earlier.length - 1].formatted;
  };

  // ── Filtering (identical logic to original) ──────────────────────────────

  const {
    filteredRows,
    headReasonOptions,
    sectionOptions,
    subHeadOptions,
    reasonDescOptions,
    lineOptions,
    unitOptions,
  } = useMemo(() => {
    const matchesFilters = (row: RowData, ignoredFilters: FilterName[] = []) => {
      const { plant, line } = getPlantAndLine(
        row["Functional Location"],
        extractLine
      );
      const unit = mapPlantToUnit(plant);

      if (
        !ignoredFilters.includes("line") &&
        lineFilter.length &&
        !lineFilter.includes(line)
      )
        return false;

      if (
        !ignoredFilters.includes("unit") &&
        unitFilter.length &&
        (!unit || !unitFilter.includes(unit))
      )
        return false;

      if (
        !ignoredFilters.includes("headReason") &&
        headReason.length &&
        !headReason.includes(String(row["Head reason"] || ""))
      )
        return false;

      if (
        !ignoredFilters.includes("section") &&
        section.length &&
        !section.includes(String(row["Section"] || ""))
      )
        return false;

      if (
        !ignoredFilters.includes("subHeadReason") &&
        subHeadReason.length &&
        !subHeadReason.includes(String(row["Sub Head reason"] || ""))
      )
        return false;

      if (
        !ignoredFilters.includes("reasonDesc") &&
        reasonDesc.length &&
        !reasonDesc.includes(String(row["Reason Desc"] || ""))
      )
        return false;

      if (dateFrom || dateTo) {
        const date = parseExcelDate(row["Date"]);
        if (!date) return false;
        if (dateFrom && date < new Date(dateFrom)) return false;
        if (dateTo && date > new Date(dateTo)) return false;
      }

      return true;
    };

    const getRowsForOptions = (ignoredFilters: FilterName[]) =>
      rawRows.filter((row) => matchesFilters(row, ignoredFilters));

    const getLineOptions = (rows: RowData[]) => {
      const lines = rows
        .map(
          (row) =>
            getPlantAndLine(row["Functional Location"], extractLine).line
        )
        .filter(Boolean);
      return sortLineKeys([...new Set([...lines, ...lineFilter])]);
    };

    const getUnitOptions = (rows: RowData[]) => {
      const units = rows
        .map((row) => {
          const { plant } = getPlantAndLine(
            row["Functional Location"],
            extractLine
          );
          return mapPlantToUnit(plant);
        })
        .filter((u): u is string => Boolean(u));
      return columns.filter((c) => units.includes(c) || unitFilter.includes(c));
    };

    return {
      filteredRows: rawRows.filter((row) => matchesFilters(row)),
      headReasonOptions: mergeStringOptions(
        getStringOptions(getRowsForOptions(["headReason"]), "Head reason"),
        headReason
      ),
      sectionOptions: mergeStringOptions(
        getStringOptions(getRowsForOptions(["section"]), "Section"),
        section
      ),
      subHeadOptions: mergeStringOptions(
        getStringOptions(getRowsForOptions(["subHeadReason"]), "Sub Head reason"),
        subHeadReason
      ),
      reasonDescOptions: mergeStringOptions(
        getStringOptions(getRowsForOptions(["reasonDesc"]), "Reason Desc"),
        reasonDesc
      ),
      lineOptions: getLineOptions(getRowsForOptions(["line"])),
      unitOptions: getUnitOptions(getRowsForOptions(["unit"])),
    };
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

  const visibleColumns = unitFilter.length
    ? columns.filter((c) => unitFilter.includes(c))
    : columns;

  // ── Build table data with previous occurrence enrichment ─────────────────

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
        columns.forEach((c) => {
          result[line][c] = [];
        });
      }

      const current = formatCurrentCell(row);
      const previous = findPreviousOccurrence(row);

      result[line][unit].push({ current, previous });
    });

    return result;
  }, [filteredRows, extractLine, findPreviousOccurrence]);

  // ── Summary rows ─────────────────────────────────────────────────────────

  const { freqRow, hrsRow } = useMemo(() => {
    const freq: SummaryRow = {};
    const hrs: SummaryRow = {};
    columns.forEach((c) => {
      freq[c] = 0;
      hrs[c] = 0;
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
    if (lineFilter.length) return sortLineKeys(Object.keys(tableData));
    if (lineMode === "sequential") return getSequentialLineKeys(tableData);
    return sortLineKeys(Object.keys(tableData));
  }, [tableData, lineMode, lineFilter]);

  // ── Handlers ─────────────────────────────────────────────────────────────

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

  // ── Export helpers (export only the current text, no previous tag) ───────

  function buildExportRows() {
    const rows: string[][] = [];
    rows.push(["Unit", ...visibleColumns]);

    lineKeys.forEach((line) => {
      const row = tableData[line];
      const maxLen = Math.max(
        ...visibleColumns.map((c) => row?.[c]?.length || 0)
      );

      for (let i = 0; i < Math.max(1, maxLen); i++) {
        rows.push([
          i === 0
            ? [line, ...(lineDetails[line] || [])].join("\n")
            : "",
          ...visibleColumns.map((c) => {
            const cell = row?.[c]?.[i];
            if (!cell) return "";
            // Append previous info as plain text in export
            return cell.previous
              ? `${cell.current}\n(Prev: ${cell.previous})`
              : cell.current;
          }),
        ]);
      }
    });

    rows.push([
      "DT – (Freq.)",
      ...visibleColumns.map((c) => String(freqRow[c] || "")),
    ]);
    rows.push([
      "DT – (Hrs.)",
      ...visibleColumns.map((c) =>
        hrsRow[c] ? hrsRow[c].toFixed(2) : ""
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

  const inputStyle: React.CSSProperties = {
    border: "1px solid #d1d5db",
    borderRadius: 4,
    padding: "4px 8px",
    fontSize: 12,
    background: "white",
  };

  // ── Render ────────────────────────────────────────────────────────────────

  return (
    <div
      style={{
        fontFamily: "Arial, sans-serif",
        minHeight: "100vh",
        background: "#f5f5f5",
        width: "100%",
      }}
    >
      <div style={{ margin: "24px auto", padding: "0 16px" }}>
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

        {/* ── Upload zone ── */}
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
              const input = document.createElement("input");
              input.type = "file";
              input.accept = ".xlsx,.xls";
              input.onchange = (e) => {
                const target = e.target as HTMLInputElement;
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
            {/* ── Re-upload button ── */}
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

            {/* ── Filter bar ── */}
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
              <strong style={{ color: "#c0392b", fontSize: 13, marginRight: 4 }}>
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

            {/* ── Legend ── */}
            <div
              style={{
                display: "flex",
                alignItems: "center",
                gap: 8,
                marginBottom: 8,
                fontSize: 11,
                color: "#666",
              }}
            >
              <div
                style={{
                  display: "inline-flex",
                  alignItems: "center",
                  gap: 4,
                  padding: "2px 7px",
                  borderRadius: 3,
                  background: "#fff7e6",
                  border: "1px solid #f0b429",
                  fontSize: 10.5,
                  color: "#7a4f00",
                }}
              >
                🕐 <strong style={{ color: "#b36b00" }}>Prev:</strong> previous
                occurrence (same Line + Unit + Sub Head Reason)
              </div>
            </div>

            {/* ── Table ── */}
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
                    {visibleColumns.map((col) => (
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
                  {lineKeys.map((lineKey, idx) => {
                    const row = tableData[lineKey];

                    return (
                      <tr
                        key={lineKey}
                        style={{
                          background: idx % 2 === 0 ? "white" : "#fdf6f5",
                        }}
                      >
                        {/* Line cell */}
                        <td
                          style={{
                            border: "1px solid #ddd",
                            padding: "6px 12px",
                            fontWeight: "bold",
                            color: "#333",
                            verticalAlign: "top",
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

                        {/* Data cells */}
                        {visibleColumns.map((col) => (
                          <td
                            key={col}
                            style={{
                              border: "1px solid #ddd",
                              padding: "6px 12px",
                              verticalAlign: "top",
                            }}
                          >
                            {row?.[col]?.map((cell, cellIdx) => (
                              <div
                                key={cellIdx}
                                style={{
                                  marginBottom: cell.previous ? 6 : 2,
                                }}
                              >
                                {/* Current breakdown entry */}
                                <div style={{ lineHeight: 1.5 }}>
                                  {cell.current}
                                </div>

                                {/* Previous occurrence badge (only if exists) */}
                                {cell.previous && (
                                  <PreviousOccurrenceTag text={cell.previous} />
                                )}
                              </div>
                            ))}
                          </td>
                        ))}
                      </tr>
                    );
                  })}

                  {/* Summary: Frequency */}
                  <tr style={{ background: "#f0d6f5", fontWeight: "bold" }}>
                    <td style={{ border: "1px solid #ddd", padding: "6px 12px" }}>
                      DT – (Freq.)
                    </td>
                    {visibleColumns.map((col) => (
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

                  {/* Summary: Hours */}
                  <tr style={{ background: "#f0d6f5", fontWeight: "bold" }}>
                    <td style={{ border: "1px solid #ddd", padding: "6px 12px" }}>
                      DT – (Hrs.)
                    </td>
                    {visibleColumns.map((col) => (
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