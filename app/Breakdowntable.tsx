"use client"
import React, { useState } from "react";
import * as XLSX from "xlsx";

type RowData = Record<string, any>;

const columns = ["NGD", "BCK", "HRR", "VIL-1", "VIL-2", "IBR", "TRC"];

export default function ExcelUploader() {
  const [tableData, setTableData] = useState<any>({});

  // ✅ Extract Plant + Line
  function extractPlantAndLine(functionalLocation: string) {
    const parts = functionalLocation?.split("-") || [];
    return {
      plant: parts[0] || "",
      line: parts[3] || "",
    };
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

  function formatExcelDate(excelDate: any) {
    if (!excelDate) return "";
  
    // If already string (like 01/04/2026), return as is
    if (typeof excelDate === "string") return excelDate;
  
    // Convert Excel serial → JS date
    const date = XLSX.SSF.parse_date_code(excelDate);
  
    if (!date) return "";
  
    const d = String(date.d).padStart(2, "0");
    const m = String(date.m).padStart(2, "0");
    const y = date.y;
  
    return `${d}/${m}/${y}`;
  }
  // ✅ Format cell
  function formatCell(row: RowData) {
    const date = formatExcelDate(row["Date"]);
    const dt = row["Total Down Time(Hrs)"];
    const sub = row["Sub Head reason"];
    const loss = row["Loss Capacity"];
  
    return `${date} (${dt}) – ${sub} – ${loss}`;
  }

  // ✅ Process Data → Pivot Table
  function processData(rows: RowData[]) {
    const result: any = {};

    rows.forEach((row) => {
      const { plant, line } = extractPlantAndLine(
        row["Functional Location"]
      );

      const unit = mapPlantToUnit(plant);
      if (!unit || !line) return;

      if (!result[line]) {
        result[line] = {};
        columns.forEach((col) => (result[line][col] = []));
      }

      result[line][unit].push(formatCell(row));
    });

    return result;
  }

  // ✅ Handle file
  const handleFile = (file: File) => {
    const reader = new FileReader();

    reader.onload = (e) => {
      const binaryStr = e.target?.result;
      const workbook = XLSX.read(binaryStr, { type: "binary" });

      const sheet = workbook.Sheets[workbook.SheetNames[0]];
      const jsonData: RowData[] = XLSX.utils.sheet_to_json(sheet);

      const finalTable = processData(jsonData);
      console.log(finalTable);

      setTableData(finalTable);
    };

    reader.readAsBinaryString(file);
  };

  // ✅ Drag drop
  const handleDrop = (e: React.DragEvent<HTMLDivElement>) => {
    e.preventDefault();
    const file = e.dataTransfer.files[0];
    if (file) handleFile(file);
  };

  return (
    <div className="p-6">
      {/* Upload */}
      <div
        onDrop={handleDrop}
        onDragOver={(e) => e.preventDefault()}
        className="border-2 border-dashed border-gray-400 p-10 text-center rounded-lg"
      >
        <p className="text-gray-600">Drag & Drop Excel File</p>
      </div>

      {/* Table */}
      <div className="mt-6 overflow-auto">
        <table className="min-w-full border text-sm">
          <thead>
            <tr className="bg-gray-100">
              <th className="border px-2 py-1">Line</th>
              {columns.map((col) => (
                <th key={col} className="border px-2 py-1">
                  {col}
                </th>
              ))}
            </tr>
          </thead>

          <tbody>
            {Object.entries(tableData).map(([line, row]: any) => (
              <tr key={line}>
                <td className="border px-2 py-1 font-semibold">
                  {line}
                </td>

                {columns.map((col) => (
                  <td key={col} className="border px-2 py-1 align-top">
                    {row[col]?.map((item: string, i: number) => (
                      <div key={i}>{item}</div>
                    ))}
                  </td>
                ))}
              </tr>
            ))}
          </tbody>
        </table>
      </div>
    </div>
  );
}