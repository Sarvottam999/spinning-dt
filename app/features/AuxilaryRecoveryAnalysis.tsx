"use client";

import ExcelUploader from "../Breakdowntable";

const auxPlantLineGroups: Record<string, string[]> = {
  AAC1: [
    "P2-AUX-CRY-AAC1",
    "MP-AUX-CRS-AAC1",
    "MP-AUX-CRS-AACS-AAC1",
  ],
  AAC2: [
    "P2-AUX-CRY-AAC2",
    "MP-AUX-CRS-AAC2",
    "MP-AUX-CRS-AACS-AAC2",
  ],
  AAC3: ["MP-AUX-CRS-AAC3", "MP-AUX-CRS-AACS-AAC3"],
  AAC4: ["MP-AUX-CRS-AAC4", "MP-AUX-CRS-AACS-AAC4"],
  AAC5: ["MP-AUX-CRS-AACS-AAC5"],
  AAC6: ["MP-AUX-CRS-AACS-AAC6"],
  AAC7: ["MP-AUX-CRS-AACS-AAC7"],
  AAC8: ["MP-AUX-CRS-AACS-AAC8"],

  CR01: ["MP-AUX-CRY-CR01"],
  CR02: ["MP-AUX-CRY-CR02"],
  CR03: ["MP-AUX-CRY-CR03"],
  CR04: ["MP-AUX-CRY-CR04"],
  CR05: ["MP-AUX-CRY-CR05"],
  CR06: ["MP-AUX-CRY-CR06"],
  CR07: ["MP-AUX-CRY-CR07"],
  CR08: ["MP-AUX-CRY-CR08"],

  SALT_AACR: [
    "MP-AUX-LP1-SALT-AACR",
    "MP-AUX-LP2-SALT-AACR",
    "MP-AUX-L01-SALT-AACR",
    "MP-AUX-L02-SALT-AACR",
  ],
  TEES: ["MP-AUX-L01-SALT-TEES", "MP-AUX-L02-SALT-TEES"],
  TEE1: ["MP-AUX-L03-SALT-TEE1"],
  TEE2: ["MP-AUX-L03-SALT-TEE2"],
  TEE3: ["MP-AUX-L04-SALT-TEE3"],

  APL1: ["MP-AUX-APL-APL1"],
  APL2: ["MP-AUX-APL-APL2"],
  APL3: ["MP-AUX-APL-APL3"],
  APL4: ["MP-AUX-APL-APL4"],
  APL5: ["MP-AUX-APL-APL5"],
  APL6: ["MP-AUX-APL-APL6"],

  AEAA: ["P2-AUX-ANH-AEAA"],
  AEAB: ["P2-AUX-ANH-AEAB"],

  CAL1: ["MP-AUX-CAL-ANHE-CAL1"],
  CAL2: ["MP-AUX-CAL-ANHE-CAL2"],
  CAL3: ["MP-AUX-CAL-ANHE-CAL3"],
  CAL4: ["MP-AUX-CAL-ANHE-CAL4"],
  CAL5: ["MP-AUX-CAL-ANHE-CAL5"],
  CAL6: ["MP-AUX-CAL-ANHE-CAL6"],

  MS02: ["MP-AUX-EVA-MS02"],
  MS04: ["MP-AUX-EVA-MS04"],
  MS05: ["MP-AUX-EVA-MS05"],
  MS06: ["MP-AUX-EVA-MS06"],
  MS07: ["MP-AUX-EVA-MS07"],
  MS08: ["P2-AUX-EVA-MS08", "MP-AUX-EVA-MS08"],
  MS09: ["P2-AUX-EVA-MS09", "MP-AUX-EVA-MS09"],
  MS10: ["P2-AUX-EVA-MS10", "MP-AUX-EVA-MS10"],
  MS11: ["P2-AUX-EVA-MS11", "MP-AUX-EVA-MS11"],
  MS12: ["P2-AUX-EVA-MS12", "MP-AUX-EVA-MS12"],
  MS13: ["MP-AUX-EVA-MS13"],

  MSF1: ["MP-AUX-L01-MSFE-MSF1", "MP-AUX-L05-MSFE-MSF1"],
  MSF2: [
    "MP-AUX-L01-MSFE-MSF2",
    "MP-AUX-L03-MSFE-MSF2",
    "MP-AUX-L05-MSFE-MSF2",
  ],
  MSF3: [
    "MP-AUX-L02-MSFE-MSF3",
    "MP-AUX-L04-MSFE-MSF3",
    "MP-AUX-L05-MSFE-MSF3",
    "MP-AUX-EVA-MSFE-MSF3",
  ],
  MSF4: [
    "MP-AUX-L02-MSFE-MSF4",
    "MP-AUX-L04-MSFE-MSF4",
    "MP-AUX-L05-MSFE-MSF4",
  ],
  MSF5: ["MP-AUX-COM-MSFE-MSF5", "MP-AUX-EVA-MSFE-MSF5"],
  MSF6: ["MP-AUX-L06-MSFE-MSF6", "MP-AUX-EVA-MSFE-MSF6"],
  MSF7: ["MP-AUX-L06-MSFE-MSF7"],
  MSF8: ["MP-AUX-L06-MSFE-MSF8"],
  MSF9: ["MP-AUX-L06-MSFE-MSF9"],
};

const auxPlantLineLookup = Object.entries(auxPlantLineGroups).reduce<
  Record<string, string>
>((lookup, [groupName, itemCodes]) => {
  itemCodes.forEach((itemCode) => {
    lookup[itemCode] = groupName;
  });

  return lookup;
}, {});

function extractAuxilaryRecoveryLine(functionalLocation: string) {
  const [, ...lineParts] = functionalLocation?.split("-") || [];
  const line = lineParts.join("-");

  return auxPlantLineLookup[line] || line;
}

export default function AuxilaryRecoveryAnalysis() {
  return (
    <ExcelUploader
      title="Auxilary&Recovery Analysis"
      reportName="auxilary_recovery"
      extractLine={extractAuxilaryRecoveryLine}
      lineDetails={auxPlantLineGroups}
      lineMode="actual"
    />
  );
}
