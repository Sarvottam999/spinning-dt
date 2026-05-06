"use client";

import ExcelUploader from "../Breakdowntable";

const acidPlantLineGroups: Record<string, string[]> = {
  "AP # 1": [
    "P3-ACD-AP1",
    "MP-ANC-ACD-ACP1",
    "MP-ANC-ACD-AC01",
  ],

  "AP # 2": [
    "P3-ACD-AP2",
    "MP-ACD-AC2",
    "MP-CSA-AP2",
    "MP-ANC-ACD-ACP2",
    "MP-ANC-ACD-AC02",
  ],

  "AP # 3": [
    "P3-ACD-AP3",
    "MP-ANC-ACD-ACP3",
    "MP-ANC-ACD-AC03",
  ],

  "AP # 4": [
    "P3-ACD-AP4",
  ],

  AP: [
    "MP-ACD",
  ],

  WSA2: [
    "MP-ANC-WSA-WSA2",
  ],

  
};

const acidPlantLineLookup = Object.entries(acidPlantLineGroups).reduce<
  Record<string, string>
>((lookup, [groupName, itemCodes]) => {
  itemCodes.forEach((itemCode) => {
    lookup[itemCode] = groupName;
  });

  return lookup;
}, {});

function extractAcidPlantLine(functionalLocation: string) {
  const [, ...lineParts] = functionalLocation?.split("-") || [];
  const line = lineParts.join("-");

  return acidPlantLineLookup[line] || line;
}

export default function AcidPlantAnalysis() {
  return (
    <ExcelUploader
      title="Acid Plant Analysis"
      reportName="acid_plant"
      extractLine={extractAcidPlantLine}
      lineDetails={acidPlantLineGroups}
      lineMode="actual"
    />
  );
}

