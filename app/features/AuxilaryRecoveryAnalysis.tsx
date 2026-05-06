"use client";

import ExcelUploader from "../Breakdowntable";

function extractAuxilaryRecoveryLine(functionalLocation: string) {
  const [, ...lineParts] = functionalLocation?.split("-") || [];
  return lineParts.join("-");
}

export default function AuxilaryRecoveryAnalysis() {
  return (
    <ExcelUploader
      title="Auxilary&Recovery Analysis"
      reportName="auxilary_recovery"
      extractLine={extractAuxilaryRecoveryLine}
      lineMode="actual"
    />
  );
}
