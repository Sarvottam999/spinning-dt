"use client";

import { useState, useCallback, CSSProperties, ReactNode } from "react";
import * as XLSX from "xlsx";

// ─── Types ────────────────────────────────────────────────────────────────────

type SectionKey =
  | "spinning" | "dryer" | "bailingPress"
  | "simplex" | "churn" | "twinRollPress"
  | "firstStageGCF" | "secondStageGCF" | "thirdStageGCF" | "rejectGCF" | "msfe"
  | "aac" | "anhydrousEvaporator" | "acidPlant"
  | "cs2Furnace" | "cs2Plant" | "wsa";

type HeadReason = "IDLE" | "PLANNED" | "UNPLANNED" | "OTHERS";

interface Unit {
  name: string;
  plant: number;
}

interface MonthMeta {
  label: string;
  y: number;
  m: number;
  days: number;
}

interface ParsedRow {
  _d: unknown;
  _plant: number;
  _sec: string;
  _hr: string;
  _shr: string;
  _rd: string;
  _down: number;
}

interface YM {
  y: number;
  m: number;
}

type MonthMap   = Record<string, number>;
type SkMap      = Record<string, MonthMap>;
type PlantMap   = Record<number, SkMap>;
type HrMap      = PlantMap;
type ShrMap     = Record<string, Record<string, PlantMap>>;
type ReasonMap  = Record<string, Record<string, Record<string, PlantMap>>>;

interface AnalysisMaps {
  idleMap:     HrMap;
  hrMap:       Record<string, HrMap>;
  shrMap:      ShrMap;
  reasonMap:   ReasonMap;
  months:      MonthMeta[];
  activePks:   SectionKey[];
  activeUnits: Unit[];
}

interface HrColorDef {
  bg:    string;
  text:  string;
  badge: string;
}

// ─── Static Data ─────────────────────────────────────────────────────────────

const UNITS: Unit[] = [
  { name: "NGD",   plant: 1101 },
  { name: "BCK",   plant: 1201 },
  { name: "HRR",   plant: 1301 },
  { name: "VIL-1", plant: 1601 },
  { name: "VIL-2", plant: 1602 },
  { name: "IBR",   plant: 2111 },
  { name: "TRC",   plant: 4101 },
];

const MC: Record<string, Record<SectionKey, number>> = {
  NGD:    { spinning:11,dryer:11,bailingPress:13,simplex:22,churn:0,twinRollPress:15,firstStageGCF:42,secondStageGCF:31,thirdStageGCF:32,rejectGCF:4,msfe:17,aac:5,anhydrousEvaporator:6,acidPlant:4,cs2Furnace:16,cs2Plant:0,wsa:0 },
  BCK:    { spinning:4,dryer:4,bailingPress:5,simplex:12,churn:0,twinRollPress:14,firstStageGCF:38,secondStageGCF:26,thirdStageGCF:27,rejectGCF:8,msfe:9,aac:4,anhydrousEvaporator:5,acidPlant:2,cs2Furnace:13,cs2Plant:0,wsa:0 },
  HRR:    { spinning:3,dryer:3,bailingPress:4,simplex:9,churn:2,twinRollPress:9,firstStageGCF:20,secondStageGCF:16,thirdStageGCF:16,rejectGCF:6,msfe:8,aac:3,anhydrousEvaporator:4,acidPlant:1,cs2Furnace:9,cs2Plant:0,wsa:0 },
  "VIL-1":{ spinning:4,dryer:4,bailingPress:5,simplex:0,churn:8,twinRollPress:17,firstStageGCF:41,secondStageGCF:34,thirdStageGCF:34,rejectGCF:11,msfe:10,aac:7,anhydrousEvaporator:6,acidPlant:1,cs2Furnace:0,cs2Plant:1,wsa:0 },
  "VIL-2":{ spinning:2,dryer:2,bailingPress:5,simplex:0,churn:10,twinRollPress:16,firstStageGCF:28,secondStageGCF:24,thirdStageGCF:24,rejectGCF:8,msfe:10,aac:4,anhydrousEvaporator:5,acidPlant:1,cs2Furnace:0,cs2Plant:1,wsa:1 },
  IBR:    { spinning:5,dryer:5,bailingPress:7,simplex:18,churn:4,twinRollPress:20,firstStageGCF:43,secondStageGCF:30,thirdStageGCF:31,rejectGCF:8,msfe:13,aac:8,anhydrousEvaporator:6,acidPlant:3,cs2Furnace:0,cs2Plant:1,wsa:2 },
  TRC:    { spinning:4,dryer:4,bailingPress:7,simplex:19,churn:3,twinRollPress:16,firstStageGCF:39,secondStageGCF:25,thirdStageGCF:29,rejectGCF:10,msfe:10,aac:7,anhydrousEvaporator:6,acidPlant:3,cs2Furnace:0,cs2Plant:1,wsa:0 },
};

const ALL_KEYS: SectionKey[] = [
  "spinning","dryer","bailingPress",
  "simplex","churn","twinRollPress","firstStageGCF","secondStageGCF","thirdStageGCF","rejectGCF","msfe",
  "aac","anhydrousEvaporator","acidPlant",
  "cs2Furnace","cs2Plant","wsa",
];

const PL: Record<SectionKey, string> = {
  spinning:"Spinning", dryer:"Dryer", bailingPress:"Bailing Press",
  simplex:"Simplex", churn:"Churn", twinRollPress:"Twin Roll Press",
  firstStageGCF:"1st Stage GCF", secondStageGCF:"2nd Stage GCF", thirdStageGCF:"3rd Stage GCF",
  rejectGCF:"Reject GCF", msfe:"MSFE", aac:"AAC",
  anhydrousEvaporator:"Anhydrous Evaporator", acidPlant:"Acid Plant",
  cs2Furnace:"CS2 Furnace", cs2Plant:"CS2 Plant", wsa:"WSA",
};

const SEC_MAP: Record<string, SectionKey> = {
  "SPINNING MACHINE":"spinning","SPINNING":"spinning","SPIN":"spinning",
  "DRYER":"dryer","DRIER":"dryer",
  "BAILING PRESS":"bailingPress","BALING PRESS":"bailingPress","BAILING":"bailingPress","BALING":"bailingPress",
  "SIMPLEX":"simplex","CHURN":"churn",
  "TWIN ROLL PRESS":"twinRollPress","TWIN ROLL":"twinRollPress","TWINROLL":"twinRollPress",
  "1ST STAGE GCF":"firstStageGCF","FIRST STAGE GCF":"firstStageGCF","STAGE 1 GCF":"firstStageGCF",
  "2ND STAGE GCF":"secondStageGCF","SECOND STAGE GCF":"secondStageGCF","STAGE 2 GCF":"secondStageGCF",
  "3RD STAGE GCF":"thirdStageGCF","THIRD STAGE GCF":"thirdStageGCF","STAGE 3 GCF":"thirdStageGCF",
  "REJECT GCF":"rejectGCF","GCF REJECT":"rejectGCF","MSFE":"msfe","AAC":"aac",
  "ANHYDROUS EVAPORATOR":"anhydrousEvaporator","ANHYDROUS":"anhydrousEvaporator",
  "ACID PLANT":"acidPlant","ACID":"acidPlant",
  "CS2 FURNACE":"cs2Furnace","CS2FURNACE":"cs2Furnace","CS-2 FURNACE":"cs2Furnace",
  "CS2 PLANT":"cs2Plant","CS2PLANT":"cs2Plant","CS-2 PLANT":"cs2Plant","WSA":"wsa",
};

const HR_ORDER: HeadReason[] = ["IDLE","PLANNED","UNPLANNED","OTHERS"];

const HR_COLOR: Record<HeadReason, HrColorDef> = {
  IDLE:      { bg:"#fce4d6", text:"#7b2800", badge:"#f97316" },
  PLANNED:   { bg:"#d1fae5", text:"#065f46", badge:"#059669" },
  UNPLANNED: { bg:"#fef3c7", text:"#78350f", badge:"#d97706" },
  OTHERS:    { bg:"#f1f5f9", text:"#334155", badge:"#64748b" },
};

const HR_STYLE: Record<HeadReason, { headerBg: string; headerText: string; totalBg: string; totalText: string; subBg: string }> = {
  IDLE:      { headerBg:"#f97316", headerText:"#fff",    totalBg:"#fce4d6", totalText:"#7b2800", subBg:"#fff7f3" },
  PLANNED:   { headerBg:"#059669", headerText:"#fff",    totalBg:"#d1fae5", totalText:"#065f46", subBg:"#f0fdf9" },
  UNPLANNED: { headerBg:"#d97706", headerText:"#fff",    totalBg:"#fef3c7", totalText:"#78350f", subBg:"#fffbeb" },
  OTHERS:    { headerBg:"#64748b", headerText:"#fff",    totalBg:"#f1f5f9", totalText:"#334155", subBg:"#f8fafc" },
};

// ─── Helpers ──────────────────────────────────────────────────────────────────

function trunc3(v: number): number {
  return Math.trunc(v * 1000) / 1000;
}

function normCol(s: string): string {
  return s.toLowerCase().replace(/[^a-z0-9]/g, " ").replace(/\s+/g, " ").trim();
}

function daysInMonth(y: number, m: number): number {
  return new Date(y, m, 0).getDate();
}

function parseDate(v: unknown): YM | null {
  if (!v) return null;
  if (typeof v === "number") {
    const dt = new Date((Math.floor(v) - 25569) * 86400000);
    return { y: dt.getUTCFullYear(), m: dt.getUTCMonth() + 1 };
  }
  if (v instanceof Date) return { y: v.getFullYear(), m: v.getMonth() + 1 };
  if (typeof v === "string") {
    const d = v.trim();
    const m1 = d.match(/^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{4})$/);
    if (m1) return { y: +m1[3], m: +m1[2] };
    const m2 = d.match(/^(\d{4})-(\d{2})-(\d{2})/);
    if (m2) return { y: +m2[1], m: +m2[2] };
  }
  return null;
}

// ─── Excel Parser ─────────────────────────────────────────────────────────────

function parseExcel(file: File): Promise<ParsedRow[]> {
  return new Promise((resolve) => {
    const reader = new FileReader();
    reader.onload = (e: ProgressEvent<FileReader>) => {
      const result = e.target?.result;
      if (!result) { resolve([]); return; }

      const wb   = XLSX.read(result, { type: "array", raw: true });
      const ws   = wb.Sheets[wb.SheetNames[0]];
      const json = XLSX.utils.sheet_to_json<Record<string, unknown>>(ws, { defval: "", raw: true });

      const rows: ParsedRow[] = [];

      json.forEach((row) => {
        const r: ParsedRow = { _d: null, _plant: 0, _sec: "", _hr: "", _shr: "", _rd: "", _down: 0 };

        Object.entries(row).forEach(([k, v]) => {
          const n = normCol(k);
          if (n === "date")
            r._d = v;
          else if (n === "plant")
            r._plant = Number(v) || 0;
          else if (n === "section")
            r._sec = String(v).toUpperCase().replace(/\s+/g, " ").trim();
          else if (n === "head reason" || n === "headreason")
            r._hr = String(v).toUpperCase().replace(/\s+/g, " ").trim();
          else if (n === "sub head reason" || n === "subheadreason" || n === "sub reason")
            r._shr = String(v).trim();
          else if (n === "reason desc" || n === "reasondesc" || n === "reason description")
            r._rd = String(v).trim();
          else if (n.includes("total down") || n === "totaldowntime")
            r._down = Number(v) || 0;
        });

        if (r._plant && r._sec) rows.push(r);
      });

      resolve(rows);
    };
    reader.readAsArrayBuffer(file);
  });
}

// ─── Build Analysis Maps ──────────────────────────────────────────────────────

function buildMaps(rows: ParsedRow[]): AnalysisMaps {
  const idleMap:   HrMap                        = {};
  const hrMap:     Record<string, HrMap>        = {};
  const shrMap:    ShrMap                       = {};
  const reasonMap: ReasonMap                    = {};
  const seenPks    = new Set<string>();
  const seenPlants = new Set<number>();
  const seenMonths = new Map<string, MonthMeta>();

  function addDown(target: Record<string, unknown>, keys: (string | number)[], down: number): void {
    let cur: Record<string, unknown> = target;
    keys.slice(0, -1).forEach((k) => {
      const ks = String(k);
      if (!cur[ks]) cur[ks] = {};
      cur = cur[ks] as Record<string, unknown>;
    });
    const last = String(keys[keys.length - 1]);
    cur[last] = ((cur[last] as number) || 0) + down;
  }

  rows.forEach((r) => {
    const sk = SEC_MAP[r._sec];
    if (sk) seenPks.add(sk);
    seenPlants.add(r._plant);

    const ymd = parseDate(r._d);
    if (!ymd) return;

    const lbl = `${String(ymd.m).padStart(2, "0")}/${String(ymd.y).slice(-2)}`;
    if (!seenMonths.has(lbl))
      seenMonths.set(lbl, { label: lbl, y: ymd.y, m: ymd.m, days: daysInMonth(ymd.y, ymd.m) });

    if (!sk) return;

    const p   = r._plant;
    const hr  = r._hr;
    const shr = r._shr || "Unknown";
    const rd  = r._rd  || "Unknown";
    const d   = r._down;

    if (hr === "IDLE") {
      addDown(idleMap as Record<string, unknown>, [p, sk, lbl], d);
      addDown(shrMap    as Record<string, unknown>, [hr, shr, p, sk, lbl],      d);
      addDown(reasonMap as Record<string, unknown>, [hr, shr, rd, p, sk, lbl], d);
    } else if (hr === "PLANNED" || hr === "UNPLANNED" || hr === "OTHERS") {
      if (!hrMap[hr]) hrMap[hr] = {};
      addDown(hrMap[hr]  as Record<string, unknown>, [p, sk, lbl],              d);
      addDown(shrMap     as Record<string, unknown>, [hr, shr, p, sk, lbl],     d);
      addDown(reasonMap  as Record<string, unknown>, [hr, shr, rd, p, sk, lbl], d);
    }
  });

  const months      = [...seenMonths.values()].sort((a, b) => a.y !== b.y ? a.y - b.y : a.m - b.m);
  const activePks   = ALL_KEYS.filter((k) => seenPks.has(k));
  const activeUnits = UNITS.filter((u) => seenPlants.has(u.plant));

  return { idleMap, hrMap, shrMap, reasonMap, months, activePks, activeUnits };
}

// ─── Pct Calculation ─────────────────────────────────────────────────────────

function calcPct(
  plant:    number,
  unitName: string,
  sk:       SectionKey,
  months:   MonthMeta[],
  getDown:  (sk: SectionKey, monthLabel: string) => number,
  idleMap:  HrMap,
): number | null {
  const mc = MC[unitName]?.[sk] ?? 0;
  if (mc === 0) return null;

  let num = 0;
  let den = 0;

  months.forEach((mk) => {
    const total = mc * mk.days * 24;
    const idle  = (idleMap[plant]?.[sk]?.[mk.label] as number | undefined) ?? 0;
    const denom = total - idle;
    const down  = getDown(sk, mk.label);
    if (down === 0) return;
    num += down;
    den += denom;
  });

  if (num === 0 || den === 0) return null;
  return trunc3((num / den) * 100);
}

// ─── FY grouping helper ───────────────────────────────────────────────────────

function fyLabel(m: MonthMeta): string {
  const fy = m.m >= 4 ? m.y + 1 : m.y;
  return `FY${String(fy).slice(-2)}`;
}

function fyGroups(months: MonthMeta[]): string[] {
  const seen = new Set<string>();
  months.forEach((m) => seen.add(fyLabel(m)));
  return [...seen].sort();
}

// ─── Cross-Unit Table (with expandable drill-down) ────────────────────────────

interface CrossUnitTableProps {
  sk:          SectionKey;
  activeUnits: Unit[];
  activePks:   SectionKey[];
  months:      MonthMeta[];
  idleMap:     HrMap;
  hrMap:       Record<string, HrMap>;
  shrMap:      ShrMap;
  reasonMap:   ReasonMap;
}

function CrossUnitTable({
  sk, activeUnits, activePks, months, idleMap, hrMap, shrMap, reasonMap,
}: CrossUnitTableProps): ReactNode {
  // Default: all HR and SHR rows expanded
  const initHR = new Set<string>(HR_ORDER);
  const initSHR = (): Set<string> => {
    const s = new Set<string>();
    HR_ORDER.forEach((hr) => {
      Object.keys(shrMap[hr] ?? {}).forEach((shr) => s.add(`${hr}|${shr}`));
    });
    return s;
  };

  const [expandedHR,  setExpandedHR]  = useState<Set<string>>(initHR);
  const [expandedSHR, setExpandedSHR] = useState<Set<string>>(initSHR);

  if (!activePks.includes(sk)) return null;

  const toggleHR = (hr: string): void =>
    setExpandedHR((prev) => { const n = new Set(prev); n.has(hr) ? n.delete(hr) : n.add(hr); return n; });

  const toggleSHR = (hr: string, shr: string): void => {
    const key = `${hr}|${shr}`;
    setExpandedSHR((prev) => { const n = new Set(prev); n.has(key) ? n.delete(key) : n.add(key); return n; });
  };

  const fys    = fyGroups(months);
  const lastFy = fys[fys.length - 1];
  const monthlyCols = months;
  const colsPerUnit = monthlyCols.length + 1;

  const getMonthDown = (
    unit: Unit,
    hr: HeadReason,
    shr: string | null,
    rd: string | null,
    ml: string,
  ): number => {
    if (rd !== null && shr !== null) {
      return (reasonMap[hr]?.[shr]?.[rd]?.[unit.plant]?.[sk]?.[ml] as number | undefined) ?? 0;
    }
    if (shr !== null) {
      return (shrMap[hr]?.[shr]?.[unit.plant]?.[sk]?.[ml] as number | undefined) ?? 0;
    }
    return hr === "IDLE"
      ? (idleMap[unit.plant]?.[sk]?.[ml] as number | undefined) ?? 0
      : (hrMap[hr]?.[unit.plant]?.[sk]?.[ml] as number | undefined) ?? 0;
  };

  const toPct = (
    unit: Unit,
    forMonths: MonthMeta[],
    hr: HeadReason,
    shr: string | null,
    rd: string | null = null,
  ): number | null => {
    const mc = MC[unit.name]?.[sk] ?? 0;
    if (mc === 0) return null;
    let num = 0;
    let den = 0;
    forMonths.forEach((mk) => {
      const down = getMonthDown(unit, hr, shr, rd, mk.label);
      if (down === 0) return;
      const total = mc * mk.days * 24;
      const idle  = (idleMap[unit.plant]?.[sk]?.[mk.label] as number | undefined) ?? 0;
      num += down;
      den += total - idle;
    });
    if (num === 0 || den === 0) return null;
    return trunc3((num / den) * 100);
  };

  const fmtV = (v: number | null): string => {
    if (v === null) return "-";
    const s = v.toFixed(3);
    return v > 100 ? `⚠${s} %` : `${s} %`;
  };

  const baseCell: CSSProperties = {
    padding: "5px 8px", fontSize: 11, textAlign: "right",
    borderBottom: "0.5px solid #e2e8f0", borderRight: "0.5px solid #e2e8f0",
    whiteSpace: "nowrap", minWidth: 52,
  };

  const valStyle = (v: number | null, bg?: string): CSSProperties => ({
    ...baseCell,
    background: bg ?? "transparent",
    color: v === null ? "#cbd5e1" : v > 100 ? "#dc2626" : v > 10 ? "#b45309" : "#1e293b",
    fontWeight: v !== null && v > 0 ? 500 : 400,
  });

  const shrsByHr: Record<HeadReason, string[]> = {
    IDLE:      Object.keys(shrMap["IDLE"]      ?? {}).sort(),
    PLANNED:   Object.keys(shrMap["PLANNED"]   ?? {}).sort(),
    UNPLANNED: Object.keys(shrMap["UNPLANNED"] ?? {}).sort(),
    OTHERS:    Object.keys(shrMap["OTHERS"]    ?? {}).sort(),
  };

  // Render monthly + UTD cells for a given unit/hr/shr/rd combo
  const renderUnitCells = (
    unit: Unit,
    hr: HeadReason,
    shr: string | null,
    rd: string | null,
    rowBg?: string,
  ): ReactNode[] => {
    const cells: ReactNode[] = [];
    monthlyCols.forEach((mk) => {
      const v = toPct(unit, [mk], hr, shr, rd);
      cells.push(<td key={`${unit.name}-m-${mk.label}`} style={valStyle(v, rowBg)}>{fmtV(v)}</td>);
    });
    const vAll = toPct(unit, months, hr, shr, rd);
    cells.push(
      <td key={`${unit.name}-utd`} style={{ ...valStyle(vAll, rowBg), fontWeight: 600, borderRight: "2px solid #cbd5e1" }}>
        {fmtV(vAll)}
      </td>,
    );
    return cells;
  };

  const rows: ReactNode[] = [];

  HR_ORDER.forEach((hr) => {
    const shrs    = shrsByHr[hr];
    const style   = HR_STYLE[hr];
    const hrOpen  = expandedHR.has(hr);
    if (shrs.length === 0) return;

    // ── HR header row (clickable) ──────────────────────────────────────────
    rows.push(
      <tr
        key={`hr-hdr-${hr}`}
        style={{ cursor: "pointer" }}
        onClick={() => toggleHR(hr)}
      >
        <td style={{
          padding: "6px 10px", fontSize: 11, fontWeight: 700,
          background: style.headerBg, color: style.headerText,
          borderBottom: "0.5px solid #e2e8f0", borderRight: "0.5px solid #e2e8f0",
          letterSpacing: "0.06em",
          userSelect: "none",
        }}>
          {hr}
        </td>
        {activeUnits.map((u) => (
          <td key={u.name} colSpan={colsPerUnit}
              style={{
                background: style.headerBg, borderBottom: "0.5px solid #e2e8f0",
                borderRight: "2px solid #cbd5e1",
              }} />
        ))}
      </tr>,
    );

    if (!hrOpen) return;

    shrs.forEach((shr) => {
      const shrKey  = `${hr}|${shr}`;
      const shrOpen = expandedSHR.has(shrKey);
      const rds     = Object.keys(reasonMap[hr]?.[shr] ?? {}).sort();

      // ── SHR row (clickable if has reason descriptions) ─────────────────
      rows.push(
        <tr
          key={`shr-${hr}-${shr}`}
          style={{ background: style.subBg, cursor: rds.length > 0 ? "pointer" : "default" }}
          onClick={() => rds.length > 0 && toggleSHR(hr, shr)}
        >
          <td style={{
            padding: "5px 10px 5px 22px", fontSize: 11, fontWeight: 400,
            color: "#374151", borderBottom: "0.5px solid #e2e8f0",
            borderRight: "0.5px solid #e2e8f0", whiteSpace: "nowrap",
            userSelect: "none",
          }}>
            <span style={{
              width: 5, height: 5, borderRadius: "50%", background: "#94a3b8",
              display: "inline-block", marginRight: 7, verticalAlign: "middle",
            }} />
            {shr}
          </td>
          {activeUnits.flatMap((u) => renderUnitCells(u, hr, shr, null, style.subBg))}
        </tr>,
      );

      // ── Reason Description rows (shown when SHR expanded) ──────────────
      if (shrOpen && rds.length > 0) {
        rds.forEach((rd) => {
          rows.push(
            <tr key={`rd-${hr}-${shr}-${rd}`} style={{ background: "#fff" }}>
              <td style={{
                padding: "5px 10px 5px 38px", fontSize: 11, color: "#64748b",
                borderBottom: "0.5px solid #f1f5f9", borderRight: "0.5px solid #e2e8f0",
                whiteSpace: "nowrap",
              }}>
                <span style={{
                  width: 5, height: 5, borderRadius: "50%", background: "#cbd5e1",
                  display: "inline-block", marginRight: 6, verticalAlign: "middle",
                }} />
                {rd}
              </td>
              {activeUnits.flatMap((u) => renderUnitCells(u, hr, shr, rd, "#fff"))}
            </tr>,
          );
        });
      }
    });

    // ── HR total row ────────────────────────────────────────────────────────
    rows.push(
      <tr key={`hr-total-${hr}`} style={{ background: style.totalBg }}>
        <td style={{
          padding: "6px 10px", fontSize: 11, fontWeight: 700,
          color: style.totalText, borderBottom: "1px solid #cbd5e1",
          borderRight: "0.5px solid #e2e8f0", whiteSpace: "nowrap",
        }}>
          {hr.charAt(0) + hr.slice(1).toLowerCase()} Total
        </td>
        {activeUnits.flatMap((u) => renderUnitCells(u, hr, null, null, style.totalBg))}
      </tr>,
    );
  });

  // ── Grand Total row ─────────────────────────────────────────────────────────
  rows.push(
    <tr key="total-downtime" style={{ background: "#1e293b" }}>
      <td style={{
        padding: "7px 10px", fontSize: 12, fontWeight: 700,
        color: "#f1f5f9", borderBottom: "0.5px solid #334155",
        borderRight: "0.5px solid #334155", whiteSpace: "nowrap",
      }}>
        ▣ Total Downtime
      </td>
      {activeUnits.flatMap((unit) => {
        const cells: ReactNode[] = [];

        const toPctTotal = (forMonths: MonthMeta[]): number | null => {
          const mc = MC[unit.name]?.[sk] ?? 0;
          if (mc === 0) return null;
          let num = 0;
          let den = 0;
          forMonths.forEach((mk) => {
            const down = HR_ORDER.reduce((s, hr) => s + getMonthDown(unit, hr, null, null, mk.label), 0);
            if (down === 0) return;
            const total = mc * mk.days * 24;
            const idle  = (idleMap[unit.plant]?.[sk]?.[mk.label] as number | undefined) ?? 0;
            num += down;
            den += total - idle;
          });
          if (num === 0 || den === 0) return null;
          return trunc3((num / den) * 100);
        };

        monthlyCols.forEach((mk) => {
          const v = toPctTotal([mk]);
          cells.push(
            <td key={`${unit.name}-m-${mk.label}`} style={{
              ...baseCell, background: "#1e293b",
              color: v === null ? "#475569" : v > 100 ? "#fca5a5" : "#86efac",
              fontWeight: 600,
            }}>
              {fmtV(v)}
            </td>,
          );
        });

        const vAll = toPctTotal(months);
        cells.push(
          <td key={`${unit.name}-utd`} style={{
            ...baseCell, background: "#0f172a",
            color: vAll === null ? "#475569" : vAll > 100 ? "#fca5a5" : "#4ade80",
            fontWeight: 700, borderRight: "2px solid #334155",
          }}>
            {fmtV(vAll)}
          </td>,
        );

        return cells;
      })}
    </tr>,
  );

  const unitHeaderRow = (
    <tr>
      <th rowSpan={2} style={{
        padding: "8px 10px", fontSize: 11, fontWeight: 700,
        color: "#94a3b8", background: "#0f172a", textAlign: "left",
        borderBottom: "1px solid #334155", borderRight: "0.5px solid #334155",
        minWidth: 220, position: "sticky", left: 0, zIndex: 2,
      }}>
        Head Reason / Sub Head Reason
      </th>
      {activeUnits.map((u) => (
        <th key={u.name} colSpan={colsPerUnit} style={{
          padding: "7px 10px", fontSize: 11, fontWeight: 700,
          color: "#f1f5f9", background: "#1e293b", textAlign: "center",
          borderBottom: "0.5px solid #334155", borderRight: "2px solid #475569",
          letterSpacing: "0.04em",
        }}>
          {u.name}
        </th>
      ))}
    </tr>
  );

  const subHeaderRow = (
    <tr>
      {activeUnits.map((u) => (
        <>
          {monthlyCols.map((mk) => (
            <th key={`${u.name}-${mk.label}`} style={{
              padding: "5px 8px", fontSize: 10, fontWeight: 600,
              color: "#7dd3fc", background: "#1e293b", textAlign: "right",
              borderBottom: "1px solid #334155", borderRight: "0.5px solid #334155",
              whiteSpace: "nowrap",
            }}>
              {mk.label}
            </th>
          ))}
          <th key={`${u.name}-utd`} style={{
            padding: "5px 8px", fontSize: 10, fontWeight: 700,
            color: "#4ade80", background: "#0f172a", textAlign: "right",
            borderBottom: "1px solid #334155", borderRight: "2px solid #475569",
            whiteSpace: "nowrap",
          }}>
            {lastFy} UTD Avg
          </th>
        </>
      ))}
    </tr>
  );

  return (
    <div style={{ marginBottom: 32, border: "0.5px solid #e2e8f0", borderRadius: 8, overflow: "hidden" }}>
      <div style={{ background: "#0f172a", padding: "10px 14px", display: "flex", alignItems: "center", gap: 8 }}>
        <span style={{ fontSize: 10, fontWeight: 700, color: "#64748b", letterSpacing: "0.1em", textTransform: "uppercase" }}>
          Cross-Unit View
        </span>
        <span style={{ fontSize: 14, fontWeight: 600, color: "#f1f5f9" }}>{PL[sk]}</span>
        <span style={{ fontSize: 10, color: "#475569", marginLeft: "auto" }}>
          Click a row to expand / collapse
        </span>
      </div>

      <div style={{ overflowX: "auto" }}>
        <table style={{ borderCollapse: "collapse", width: "100%", tableLayout: "auto" }}>
          <thead>
            {unitHeaderRow}
            {subHeaderRow}
          </thead>
          <tbody>{rows}</tbody>
        </table>
      </div>
    </div>
  );
}

// ─── Section Table ────────────────────────────────────────────────────────────

interface SectionTableProps {
  sk:          SectionKey;
  activeUnits: Unit[];
  activePks:   SectionKey[];
  months:      MonthMeta[];
  idleMap:     HrMap;
  hrMap:       Record<string, HrMap>;
  shrMap:      ShrMap;
  reasonMap:   ReasonMap;
}

function SectionTable({
  sk, activeUnits, activePks, months, idleMap, hrMap, shrMap, reasonMap,
}: SectionTableProps): ReactNode {
  const [expandedHR,  setExpandedHR]  = useState<Set<string>>(new Set());
  const [expandedSHR, setExpandedSHR] = useState<Set<string>>(new Set());

  if (!activePks.includes(sk)) return null;

  const toggleHR = (hr: string): void => {
    setExpandedHR((prev) => { const n = new Set(prev); n.has(hr) ? n.delete(hr) : n.add(hr); return n; });
  };

  const toggleSHR = (hr: string, shr: string): void => {
    const key = `${hr}|${shr}`;
    setExpandedSHR((prev) => { const n = new Set(prev); n.has(key) ? n.delete(key) : n.add(key); return n; });
  };

  const getPct = (unit: Unit, getDown: (sk: SectionKey, ml: string) => number): number | null =>
    calcPct(unit.plant, unit.name, sk, months, getDown, idleMap);

  const fmtPct = (v: number | null): string => v === null ? "-" : v > 100 ? `⚠ ${v.toFixed(3)} %` : `${v.toFixed(3)} %`;

  const cellStyle = (v: number | null): CSSProperties => ({
    textAlign: "right", padding: "6px 10px", fontSize: 12,
    borderBottom: "0.5px solid #e2e8f0", borderRight: "0.5px solid #e2e8f0",
    color: v === null ? "#cbd5e1" : v > 100 ? "#dc2626" : v > 10 ? "#b45309" : "#1e293b",
    fontWeight: v !== null && v > 0 ? 500 : 400,
    background: "transparent", whiteSpace: "nowrap", minWidth: 64,
  });

  const tableRows: ReactNode[] = [];

  HR_ORDER.forEach((hr) => {
    const hrOpen = expandedHR.has(hr);
    const shrs   = Object.keys(shrMap[hr] ?? {}).sort();
    if (shrs.length === 0) return;

    const hrVals = activeUnits.map((u) =>
      getPct(u, (_sk, ml) =>
        hr === "IDLE"
          ? (idleMap[u.plant]?.[_sk]?.[ml] as number | undefined) ?? 0
          : (hrMap[hr]?.[u.plant]?.[_sk]?.[ml] as number | undefined) ?? 0,
      ),
    );

    tableRows.push(
      <tr key={`hr-${hr}`} style={{ background: HR_COLOR[hr].bg, cursor: "pointer" }} onClick={() => toggleHR(hr)}>
        <td style={{
          padding: "7px 10px", fontSize: 12, fontWeight: 600, color: HR_COLOR[hr].text,
          borderBottom: "0.5px solid #e2e8f0", borderRight: "0.5px solid #e2e8f0", whiteSpace: "nowrap",
        }}>
          <span style={{ fontSize: 10, opacity: 0.7, width: 12, display: "inline-block", marginRight: 4 }}>
            {hrOpen ? "▼" : "▶"}
          </span>
          <span style={{
            background: HR_COLOR[hr].badge, color: "#fff", fontSize: 10,
            padding: "1px 6px", borderRadius: 3, fontWeight: 700, letterSpacing: "0.04em", marginRight: 6,
          }}>
            {hr}
          </span>
          {hr.charAt(0) + hr.slice(1).toLowerCase()} downtime
        </td>
        {hrVals.map((v, i) => <td key={i} style={cellStyle(v)}>{fmtPct(v)}</td>)}
      </tr>,
    );

    if (!hrOpen) return;

    shrs.forEach((shr) => {
      const shrKey  = `${hr}|${shr}`;
      const shrOpen = expandedSHR.has(shrKey);
      const rds     = Object.keys(reasonMap[hr]?.[shr] ?? {}).sort();

      const shrVals = activeUnits.map((u) =>
        getPct(u, (_sk, ml) => (shrMap[hr]?.[shr]?.[u.plant]?.[_sk]?.[ml] as number | undefined) ?? 0),
      );

      tableRows.push(
        <tr key={`shr-${hr}-${shr}`} style={{ background: "#fff", cursor: rds.length > 0 ? "pointer" : "default" }}
            onClick={() => rds.length > 0 && toggleSHR(hr, shr)}>
          <td style={{
            padding: "6px 10px 6px 28px", fontSize: 12, fontWeight: 500, color: "#374151",
            borderBottom: "0.5px solid #e2e8f0", borderRight: "0.5px solid #e2e8f0", whiteSpace: "nowrap",
          }}>
            {rds.length > 0 && (
              <span style={{ fontSize: 9, opacity: 0.5, width: 10, display: "inline-block", marginRight: 4 }}>
                {shrOpen ? "▼" : "▶"}
              </span>
            )}
            {!rds.length && <span style={{ width: 14, display: "inline-block" }} />}
            {shr}
          </td>
          {shrVals.map((v, i) => <td key={i} style={{ ...cellStyle(v), background: "#fafafa" }}>{fmtPct(v)}</td>)}
        </tr>,
      );

      if (!shrOpen) return;

      rds.forEach((rd) => {
        const rdVals = activeUnits.map((u) =>
          getPct(u, (_sk, ml) => (reasonMap[hr]?.[shr]?.[rd]?.[u.plant]?.[_sk]?.[ml] as number | undefined) ?? 0),
        );
        tableRows.push(
          <tr key={`rd-${hr}-${shr}-${rd}`} style={{ background: "#fff" }}>
            <td style={{
              padding: "5px 10px 5px 46px", fontSize: 11, color: "#64748b",
              borderBottom: "0.5px solid #f1f5f9", borderRight: "0.5px solid #e2e8f0", whiteSpace: "nowrap",
            }}>
              <span style={{
                width: 6, height: 6, borderRadius: "50%", background: "#cbd5e1",
                display: "inline-block", marginRight: 6, verticalAlign: "middle",
              }} />
              {rd}
            </td>
            {rdVals.map((v, i) => <td key={i} style={{ ...cellStyle(v), background: "#fff", fontSize: 11 }}>{fmtPct(v)}</td>)}
          </tr>,
        );
      });
    });
  });

  return (
    <div style={{ marginBottom: 24, border: "0.5px solid #e2e8f0", borderRadius: 8, overflow: "hidden" }}>
      <div style={{ background: "#1e293b", padding: "10px 14px", display: "flex", alignItems: "center", gap: 8 }}>
        <span style={{ fontSize: 11, fontWeight: 700, color: "#94a3b8", letterSpacing: "0.08em", textTransform: "uppercase" }}>
          Section
        </span>
        <span style={{ fontSize: 14, fontWeight: 600, color: "#f1f5f9" }}>{PL[sk]}</span>
      </div>
      <div style={{ overflowX: "auto" }}>
        <table style={{ borderCollapse: "collapse", width: "100%", tableLayout: "auto" }}>
          <thead>
            <tr style={{ background: "#f8fafc" }}>
              <th style={{
                padding: "7px 10px", fontSize: 11, fontWeight: 600, color: "#64748b",
                textAlign: "left", borderBottom: "1px solid #e2e8f0",
                borderRight: "0.5px solid #e2e8f0", whiteSpace: "nowrap", minWidth: 200,
              }}>
                Head / Sub / Reason
              </th>
              {activeUnits.map((u) => (
                <th key={u.name} style={{
                  padding: "7px 10px", fontSize: 11, fontWeight: 600, color: "#64748b",
                  textAlign: "right", borderBottom: "1px solid #e2e8f0",
                  borderRight: "0.5px solid #e2e8f0", whiteSpace: "nowrap", minWidth: 64,
                }}>
                  {u.name}
                </th>
              ))}
            </tr>
          </thead>
          <tbody>{tableRows}</tbody>
        </table>
      </div>
    </div>
  );
}

// ─── Main App ─────────────────────────────────────────────────────────────────

export default function DowntimeConsolidated(): ReactNode {
  const [data,     setData]     = useState<AnalysisMaps | null>(null);
  const [fileName, setFileName] = useState<string>("");
  const [loading,  setLoading]  = useState<boolean>(false);
  const [drag,     setDrag]     = useState<boolean>(false);
  const [copyMsg,  setCopyMsg]  = useState<string>("");

  const load = useCallback(async (file: File): Promise<void> => {
    setLoading(true);
    setFileName(file.name);
    const rows = await parseExcel(file);
    const maps = buildMaps(rows);
    setData(maps);
    setLoading(false);
  }, []);

  const onDrop = (e: React.DragEvent<HTMLDivElement>): void => {
    e.preventDefault(); setDrag(false);
    const f = e.dataTransfer.files[0];
    if (f) void load(f);
  };

  const onFileChange = (e: React.ChangeEvent<HTMLInputElement>): void => {
    const f = e.target.files?.[0];
    if (f) void load(f);
  };

  const onReset = (): void => { setData(null); setFileName(""); };

  const onCopy = (): void => {
    const tables = document.querySelectorAll<HTMLTableElement>("#crossUnitArea table");
    const tsvParts: string[] = [];

    tables.forEach((tbl) => {
      const allRows = tbl.querySelectorAll("tr");
      allRows.forEach((row) => {
        const cells = row.querySelectorAll("th, td");
        const isSubHeader =
          row.closest("thead") !== null &&
          row.rowIndex === 1 &&
          !Array.from(cells).some((c) => parseInt(c.getAttribute("rowspan") ?? "1", 10) > 1);

        const parts: string[] = isSubHeader ? [""] : [];
        cells.forEach((c) => {
          const span = parseInt(c.getAttribute("colspan") ?? "1", 10);
          const text = (c as HTMLElement).innerText.replace(/\n/g, " ").trim();
          parts.push(text);
          for (let i = 1; i < span; i++) parts.push("");
        });
        tsvParts.push(parts.join("\t"));
      });
      tsvParts.push("");
    });

    navigator.clipboard.writeText(tsvParts.join("\n")).then(() => {
      setCopyMsg("Copied!");
      setTimeout(() => setCopyMsg(""), 2000);
    }).catch(() => {
      setCopyMsg("Failed");
      setTimeout(() => setCopyMsg(""), 2000);
    });
  };

  return (
    <div style={{ fontFamily: "'DM Sans', system-ui, sans-serif", minHeight: "100vh", background: "#f8fafc", padding: 0 }}>

      {/* Top bar */}
      <div style={{ padding: "12px 20px", display: "flex", alignItems: "center", gap: 12, flexWrap: "wrap" }}>
        {fileName && <span style={{ fontSize: 11, color: "#64748b" }}>{fileName}</span>}

        {data && (
          <>
            <button
              onClick={onCopy}
              style={{
                marginLeft: 8, fontSize: 11, padding: "4px 14px",
                background: copyMsg ? "#059669" : "#1e293b",
                border: "none", borderRadius: 4,
                color: "#f1f5f9", cursor: "pointer", transition: "background 0.2s",
                fontWeight: 500,
              }}
            >
              {copyMsg || "📋 Copy Table"}
            </button>
            <button
              onClick={onReset}
              style={{
                marginLeft: "auto", fontSize: 11, padding: "4px 12px",
                background: "transparent", border: "0.5px solid #475569",
                borderRadius: 4, color: "#94a3b8", cursor: "pointer",
              }}
            >
              ✕ Reset
            </button>
          </>
        )}
      </div>

      <div style={{ padding: 20 }}>

        {/* Upload zone */}
        {!data && !loading && (
          <div
            onDragOver={(e) => { e.preventDefault(); setDrag(true); }}
            onDragLeave={() => setDrag(false)}
            onDrop={onDrop}
            onClick={() => document.getElementById("fileInput")?.click()}
            style={{
              border: `2px dashed ${drag ? "#10b981" : "#cbd5e1"}`,
              background: drag ? "#ecfdf5" : "#fff",
              borderRadius: 10, padding: "56px 0", textAlign: "center",
              cursor: "pointer", transition: "all 0.15s",
            }}
          >
            <input id="fileInput" type="file" accept=".xlsx,.xls,.csv"
              style={{ display: "none" }} onChange={onFileChange} />
            <div style={{ fontSize: 32, marginBottom: 10 }}>📊</div>
            <div style={{ fontSize: 14, fontWeight: 600, color: "#1e293b", marginBottom: 6 }}>
              Drop your Excel file here or click to browse
            </div>
            <div style={{ fontSize: 12, color: "#94a3b8" }}>.xlsx · .xls · .csv</div>
          </div>
        )}

        {loading && (
          <div style={{ textAlign: "center", padding: "60px 0", color: "#64748b", fontSize: 13 }}>
            Parsing file...
          </div>
        )}

        {/* Cross-Unit tables first, then Original Section tables */}
        <div id="tableArea">
          <div id="crossUnitArea">
            {data && data.activePks.map((sk) => (
              <CrossUnitTable
                key={`cu-${sk}`} sk={sk}
                activeUnits={data.activeUnits}
                activePks={data.activePks}
                months={data.months}
                idleMap={data.idleMap}
                hrMap={data.hrMap}
                shrMap={data.shrMap}
                reasonMap={data.reasonMap}
              />
            ))}
          </div>
          {data && data.activePks.map((sk) => (
            <SectionTable
              key={`sec-${sk}`} sk={sk}
              activeUnits={data.activeUnits}
              activePks={data.activePks}
              months={data.months}
              idleMap={data.idleMap}
              hrMap={data.hrMap}
              shrMap={data.shrMap}
              reasonMap={data.reasonMap}
            />
          ))}
        </div>
      </div>
    </div>
  );
}