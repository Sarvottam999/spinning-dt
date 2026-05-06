"use client";

import ExcelUploader from "../Breakdowntable";

function extractViscoseLine(functionalLocation: string) {
  const [, ...lineParts] = functionalLocation?.split("-") || [];
  return lineParts.join("-");
}

export default function ViscoseAnalysis() {
  return (
    <ExcelUploader
      title="Viscose Analysis"
      reportName="viscose"
      extractLine={extractViscoseLine}
      lineMode="actual"
    />
  );
}
