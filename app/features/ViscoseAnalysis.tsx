"use client";

import ExcelUploader from "../Breakdowntable";

const viscosePlantLineGroups: Record<string, string[]> = {
  TPA0: ["P1-VIS-PRE-TPA0"],
  TPB0: ["P1-VIS-PRE-TPB0"],
  TPC0: ["P1-VIS-PRE-TPC0"],
  TPC2: ["P1-VIS-PRE-TPC2"],
  TPD1: ["P1-VIS-PRE-TPD1"],
  TPD2: ["P1-VIS-PRE-TPD2"],
  TPE3: ["P1-VIS-PRE-TPE3"],
  TPW0: ["P2-VIS-PRE-TPW0"],

  PRSA: ["MP-VIS-PRS-PRSA"],
  PRSB: ["MP-VIS-PRS-PRSB"],
  PRSC: ["MP-VIS-PRS-PRSC"],
  PRSG: ["MP-VIS-PRS-PRSG"],
  PRSH: ["MP-VIS-PRS-PRSH"],
  PRSZ: ["MP-VIS-PRS-PRSZ"],
  TRPD: ["MP-VIS-PRS-TRPD"],
  TRPE: ["MP-VIS-PRS-TRPE"],

  SP1A: ["MP-VIS-LP1-SLUR-TWRP-SP1A"],
  SP1C: ["MP-VIS-LP1-SLUR-TWRP-SP1C"],
  SP1D: ["MP-VIS-LP1-SLUR-TWRP-SP1D"],
  SP2A: ["MP-VIS-LP1-SLUR-TWRP-SP2A"],
  SP2C: ["MP-VIS-LP1-SLUR-TWRP-SP2C"],
  SP3A: ["MP-VIS-LN3-SLUR-TWRP-SP3A"],
  SP3C: ["MP-VIS-LN3-SLUR-TWRP-SP3C"],
  SP4A: ["MP-VIS-LN4-SLUR-TWRP-SP4A"],
  SP4D: ["MP-VIS-LN4-SLUR-TWRP-SP4D"],

  RE1A: ["MP-VIS-L05-XNTN-CHRN-RE1A"],
  RE1B: ["MP-VIS-L05-XNTN-CHRN-RE1B", "MP-VIS-L06-XNTN-CHRN-RE1B"],
  RE1D: ["MP-VIS-L05-XNTN-CHRN-RE1D"],

  GDP1: ["MP-VIS-L01-GDPR-GDP1"],
  GDP2: ["MP-VIS-L01-GDPR-GDP2", "MP-VIS-L02-GDPR-GDP2"],
  GDP3: ["MP-VIS-L01-GDPR-GDP3"],
  GDP4: ["MP-VIS-L01-GDPR-GDP4", "MP-VIS-L02-GDPR-GDP4"],
  GDP5: ["MP-VIS-L01-GDPR-GDP5", "MP-VIS-L02-GDPR-GDP5"],
  GDP6: ["MP-VIS-L01-GDPR-GDP6", "MP-VIS-L02-GDPR-GDP6"],
  GDP7: ["MP-VIS-L02-GDPR-GDP7"],
  GDP8: ["MP-VIS-L01-GDPR-GDP8", "MP-VIS-L02-GDPR-GDP8"],
  GDP9: ["MP-VIS-L01-GDPR-GDP9", "MP-VIS-L02-GDPR-GDP9"],
  GD10: ["MP-VIS-L01-GDPR-GD10", "MP-VIS-L02-GDPR-GD10"],

  N001: ["MP-VIS-L01-SIMX-VESL-N001"],
  N003: ["MP-VIS-L01-SIMX-VESL-N003"],
  N004: ["MP-VIS-L01-SIMX-VESL-N004"],
  N005: ["MP-VIS-L01-SIMX-VESL-N005"],
  N007: ["MP-VIS-L01-SIMX-VESL-N007"],
  N008: ["MP-VIS-L01-SIMX-VESL-N008"],
  N015: ["MP-VIS-L01-SIMX-VESL-N015"],
  N016: ["MP-VIS-L01-SIMX-VESL-N016"],
  N018: ["MP-VIS-L01-SIMX-VESL-N018"],

  SL02: ["MP-VIS-SLP-SL02"],
  SL03: ["MP-VIS-SLP-SL03"],
  SL04: ["MP-VIS-SLP-SL04"],
  SL05: ["MP-VIS-SLP-SL05"],
  SL06: ["MP-VIS-SLP-SL06"],
  SL10: ["MP-VIS-SLP-SL10"],
  SL11: ["MP-VIS-SLP-SL11"],
  SL12: ["MP-VIS-SLP-SL12"],
  SL13: ["MP-VIS-SLP-SL13"],
  SL14: ["MP-VIS-SLP-SL14"],
  SL15: ["MP-VIS-SLP-SL15"],
  SL16: ["MP-VIS-SLP-SL16"],

  XA05: ["MP-VIS-XAN-XNTR-XA05"],
  XA14: ["MP-VIS-XAN-XNTR-XA14"],
  XA16: ["MP-VIS-XAN-XNTR-XA16"],
  XA17: ["MP-VIS-XAN-XNTR-XA17"],
  XA19: ["MP-VIS-XAN-XNTR-XA19"],

  CHC1: ["MP-VIS-CHU-CHC1"],
  CHC2: ["MP-VIS-CHU-CHC2"],
  CHC3: ["MP-VIS-CHU-CHC3"],
};

const viscosePlantLineLookup = Object.entries(viscosePlantLineGroups).reduce<
  Record<string, string>
>((lookup, [groupName, itemCodes]) => {
  itemCodes.forEach((itemCode) => {
    lookup[itemCode] = groupName;
  });

  return lookup;
}, {});

function extractViscoseLine(functionalLocation: string) {
  const [, ...lineParts] = functionalLocation?.split("-") || [];
  const line = lineParts.join("-");

  return viscosePlantLineLookup[line] || line;
}

export default function ViscoseAnalysis() {
  return (
    <ExcelUploader
      title="Viscose Analysis"
      reportName="viscose"
      extractLine={extractViscoseLine}
      lineDetails={viscosePlantLineGroups}
      lineMode="actual"
    />
  );
}
