"use client";

import ExcelUploader from "../Breakdowntable";

const cs2PlantLineGroups: Record<string, string[]> = {
  FR01: ["P3-CS2-FUR-FR01", "MP-CS2-FUR-FR01"],
  FR02: ["P3-CS2-FUR-FR02", "MP-CS2-FUR-FR02"],
  FR03: ["P3-CS2-FUR-FR03", "MP-CS2-FUR-FR03"],
  FR04: ["P3-CS2-FUR-FR04", "MP-CS2-FUR-FR04"],
  FR05: ["MP-CS2-FUR-FR05", "P3-CS2-FUR-FR05"],
  FR06: ["P3-CS2-FUR-FR06", "MP-CS2-FUR-FR06"],
  FR07: ["P3-CS2-FUR-FR07", "MP-CS2-FUR-FR07"],
  FR08: ["P3-CS2-FUR-FR08", "MP-CS2-FUR-FR08"],
  FR09: ["P3-CS2-FUR-FR09", "MP-CS2-FUR-FR09"],
  FR10: ["P3-CS2-FUR-FR10", "MP-CS2-FUR-FR10"],
  FR11: ["MP-CS2-FUR-FR11"],
  FR12: ["MP-CS2-FUR-FR12"],
  FR13: ["MP-CS2-FUR-FR13"],
  FR16: ["P3-CS2-FUR-FR16"],

  FRN1: ["MP-CSA-FUR-FRN1"],
  FRN2: ["MP-CSA-FUR-FRN2"],
  FRN3: ["MP-CSA-FUR-FRN3"],
  FRN4: ["MP-CSA-FUR-FRN4"],
  FRN5: ["MP-CSA-FUR-FRN5"],
  FRN6: ["MP-CSA-FUR-FRN6"],
  FRN7: ["MP-CSA-FUR-FRN7"],
  FRN8: ["MP-CSA-FUR-FRN8"],
  FRN9: ["MP-CSA-FUR-FRN9"],

  CS2: ["MP-CS2"],
};

const cs2PlantLineLookup = Object.entries(cs2PlantLineGroups).reduce<
  Record<string, string>
>((lookup, [groupName, itemCodes]) => {
  itemCodes.forEach((itemCode) => {
    lookup[itemCode] = groupName;
  });

  return lookup;
}, {});

function extractCs2Line(functionalLocation: string) {
  const [, ...lineParts] = functionalLocation?.split("-") || [];
  const line = lineParts.join("-");

  return cs2PlantLineLookup[line] || line;
}

export default function Cs2Analysis() {
  return (
    <ExcelUploader
      title="CS2 Analysis"
      reportName="cs2"
      extractLine={extractCs2Line}
      lineDetails={cs2PlantLineGroups}
      lineMode="actual"
    />
  );
}
