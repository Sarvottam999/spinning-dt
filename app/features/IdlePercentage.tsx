"use client";
import { useState, useCallback, useRef, useMemo, useEffect, Fragment } from "react";
import * as XLSX from "xlsx";

// ─── Types ─────────────────────────────────────────────────────────────────────

interface Unit { name: string; plant: number; }
interface SectionDef { label: string; keys: string[]; }
interface ActiveSection extends SectionDef { activeKeys: string[]; }
interface YMD { y: number; m: number; d: number; }
interface MonthMeta { y: number; m: number; label: string; days: number; }

interface ParsedRow {
  _d: unknown; _plant: number; _floc: string;
  _sec: string; _hr: string; _shr: string; _down: number;
}

// map[plant][pk][monthLabel] = totalDown
type HrMap = Record<number, Record<string, Record<string, number>>>;

// shrMap[hr][shr][plant][pk][monthLabel] = totalDown
type ShrMap = Record<string, Record<string, Record<number, Record<string, Record<string, number>>>>>;

// ─── Static Data ───────────────────────────────────────────────────────────────

const UNITS: Unit[] = [
  { name: "NGD",   plant: 1101 },
  { name: "BCK",   plant: 1201 },
  { name: "HRR",   plant: 1301 },
  { name: "VIL-1", plant: 1601 },
  { name: "VIL-2", plant: 1602 },
  { name: "IBR",   plant: 2111 },
  { name: "TRC",   plant: 4101 },
];

const MC: Record<string, Record<string, number>> = {
  NGD:    { spinning:11,dryer:11,bailingPress:13,simplex:22,churn:0, twinRollPress:15,firstStageGCF:42,secondStageGCF:31,thirdStageGCF:32,rejectGCF:4, msfe:17,aac:5, anhydrousEvaporator:6,acidPlant:4,cs2Furnace:16,cs2Plant:0,wsa:0 },
  BCK:    { spinning:4, dryer:4, bailingPress:5, simplex:12,churn:0, twinRollPress:14,firstStageGCF:38,secondStageGCF:26,thirdStageGCF:27,rejectGCF:8, msfe:9, aac:4, anhydrousEvaporator:5,acidPlant:2,cs2Furnace:13,cs2Plant:0,wsa:0 },
  HRR:    { spinning:3, dryer:3, bailingPress:4, simplex:9, churn:2, twinRollPress:9, firstStageGCF:20,secondStageGCF:16,thirdStageGCF:16,rejectGCF:6, msfe:8, aac:3, anhydrousEvaporator:4,acidPlant:1,cs2Furnace:9, cs2Plant:0,wsa:0 },
  "VIL-1":{ spinning:4, dryer:4, bailingPress:5, simplex:0, churn:8, twinRollPress:17,firstStageGCF:41,secondStageGCF:34,thirdStageGCF:34,rejectGCF:11,msfe:10,aac:7, anhydrousEvaporator:6,acidPlant:1,cs2Furnace:0, cs2Plant:1,wsa:0 },
  "VIL-2":{ spinning:2, dryer:2, bailingPress:5, simplex:0, churn:10,twinRollPress:16,firstStageGCF:28,secondStageGCF:24,thirdStageGCF:24,rejectGCF:8, msfe:10,aac:4, anhydrousEvaporator:5,acidPlant:1,cs2Furnace:0, cs2Plant:1,wsa:1 },
  IBR:    { spinning:5, dryer:5, bailingPress:7, simplex:18,churn:4, twinRollPress:20,firstStageGCF:43,secondStageGCF:30,thirdStageGCF:31,rejectGCF:8, msfe:13,aac:8, anhydrousEvaporator:6,acidPlant:3,cs2Furnace:0, cs2Plant:1,wsa:2 },
  TRC:    { spinning:4, dryer:4, bailingPress:7, simplex:19,churn:3, twinRollPress:16,firstStageGCF:39,secondStageGCF:25,thirdStageGCF:29,rejectGCF:10,msfe:10,aac:7, anhydrousEvaporator:6,acidPlant:3,cs2Furnace:0, cs2Plant:1,wsa:0 },
};

const SECTIONS: SectionDef[] = [
  { label: "Spinning",  keys: ["spinning","dryer","bailingPress"] },
  { label: "Viscose",   keys: ["simplex","churn","twinRollPress","firstStageGCF","secondStageGCF","thirdStageGCF","rejectGCF","msfe"] },
  { label: "Auxiliary", keys: ["aac","anhydrousEvaporator","acidPlant"] },
  { label: "Ancillary", keys: ["cs2Furnace","cs2Plant","wsa"] },
];

const PL: Record<string, string> = {
  spinning:"Spinning", dryer:"Dryer", bailingPress:"Bailing Press",
  simplex:"Simplex", churn:"Churn", twinRollPress:"Twin Roll Press",
  firstStageGCF:"1st Stage GCF", secondStageGCF:"2nd Stage GCF", thirdStageGCF:"3rd Stage GCF",
  rejectGCF:"Reject GCF", msfe:"MSFE",
  aac:"AAC", anhydrousEvaporator:"Anhydrous Evaporator", acidPlant:"Acid Plant",
  cs2Furnace:"CS2 Furnace", cs2Plant:"CS2 Plant", wsa:"WSA",
};

const ALL_KEYS: string[] = SECTIONS.flatMap(s => s.keys);

const SEC_MAP: Record<string, string> = {
  "SPINNING MACHINE":"spinning","SPINNING":"spinning","SPIN":"spinning",
  "DRYER":"dryer","DRIER":"dryer",
  "BAILING PRESS":"bailingPress","BALING PRESS":"bailingPress","BAILING":"bailingPress","BALING":"bailingPress",
  "SIMPLEX":"simplex",
  "CHURN":"churn",
  "TWIN ROLL PRESS":"twinRollPress","TWIN ROLL":"twinRollPress","TWINROLL PRESS":"twinRollPress","TWINROLL":"twinRollPress",
  "1ST STAGE GCF":"firstStageGCF","FIRST STAGE GCF":"firstStageGCF","1ST GCF":"firstStageGCF","STAGE 1 GCF":"firstStageGCF",
  "2ND STAGE GCF":"secondStageGCF","SECOND STAGE GCF":"secondStageGCF","2ND GCF":"secondStageGCF","STAGE 2 GCF":"secondStageGCF",
  "3RD STAGE GCF":"thirdStageGCF","THIRD STAGE GCF":"thirdStageGCF","3RD GCF":"thirdStageGCF","STAGE 3 GCF":"thirdStageGCF",
  "REJECT GCF":"rejectGCF","GCF REJECT":"rejectGCF",
  "MSFE":"msfe",
  "AAC":"aac",
  "ANHYDROUS EVAPORATOR":"anhydrousEvaporator","ANHYDROUS":"anhydrousEvaporator",
  "ACID PLANT":"acidPlant","ACID":"acidPlant",
  "CS2 FURNACE":"cs2Furnace","CS2FURNACE":"cs2Furnace","CS-2 FURNACE":"cs2Furnace",
  "CS2 PLANT":"cs2Plant","CS2PLANT":"cs2Plant","CS-2 PLANT":"cs2Plant",
  "WSA":"wsa",
};

const HR_CONFIG: Record<string, { label: string; bg: string; sub: string; pctBg: string; pctSub: string; }> = {
  PLANNED:   { label:"Planned",   bg:"#e2efda", sub:"#a9d18e", pctBg:"#f0f7eb", pctSub:"#c6e3b1" },
  UNPLANNED: { label:"Unplanned", bg:"#fff2cc", sub:"#ffd966", pctBg:"#fffae5", pctSub:"#ffe99a" },
  OTHERS:    { label:"Others",    bg:"#ededed", sub:"#bfbfbf", pctBg:"#f5f5f5", pctSub:"#d6d6d6" },
};
const HR_ORDER = ["PLANNED", "UNPLANNED", "OTHERS"] as const;

// Consolidated table colour scheme per HR bucket
const CONSOL_HR_STYLE: Record<string, { rowBg: string; rowText: string; }> = {
  // ── NEW: IDLE row style (salmon/orange) ──
  IDLE:      { rowBg: "#fce4d6", rowText: "#7b2800" },
  PLANNED:   { rowBg: "#c6efce", rowText: "#1a4d1a" },
  UNPLANNED: { rowBg: "#ffeb9c", rowText: "#7b5800" },
  OTHERS:    { rowBg: "#bfbfbf", rowText: "#222" },
};
const CONSOL_COMBINED_STYLE = { rowBg: "#dce6f1", rowText: "#1e3a5f" };
const CONSOL_TOTAL_STYLE    = { rowBg: "#1e3a5f", rowText: "#fff" };

// ─── FY helpers ────────────────────────────────────────────────────────────────

function getFYLabel(mk: MonthMeta): string {
  const endYear = mk.m >= 4 ? mk.y + 1 : mk.y;
  return `FY${String(endYear).slice(-2)}`;
}

// ─── Helpers ───────────────────────────────────────────────────────────────────
function trunc3(v: number): string {
    return (Math.trunc(v * 1000) / 1000).toString();
  }
function daysInMonth(y: number, m: number): number { return new Date(y, m, 0).getDate(); }

function parseDateOnly(v: unknown): YMD | null {
  if (v === null || v === undefined || v === "") return null;
  if (typeof v === "number") {
    const ms = (Math.floor(v) - 25569) * 86400000;
    const dt = new Date(ms);
    return { y: dt.getUTCFullYear(), m: dt.getUTCMonth() + 1, d: dt.getUTCDate() };
  }
  if (typeof v === "string") {
    const s = v.trim();
    const dmy = s.match(/^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{4})$/);
    if (dmy) return { d: +dmy[1], m: +dmy[2], y: +dmy[3] };
    const iso = s.match(/^(\d{4})-(\d{2})-(\d{2})/);
    if (iso) return { y: +iso[1], m: +iso[2], d: +iso[3] };
    return null;
  }
  if (v instanceof Date) return { y: v.getFullYear(), m: v.getMonth() + 1, d: v.getDate() };
  return null;
}

function getMonths(rows: ParsedRow[]): MonthMeta[] {
  const seen = new Map<string, MonthMeta>();
  rows.forEach(r => {
    const ymd = parseDateOnly(r._d);
    if (!ymd) return;
    const label = `${String(ymd.m).padStart(2, "0")}/${String(ymd.y).slice(-2)}`;
    if (!seen.has(label))
      seen.set(label, { y: ymd.y, m: ymd.m, label, days: daysInMonth(ymd.y, ymd.m) });
  });
  return [...seen.values()].sort((a, b) => a.y !== b.y ? a.y - b.y : a.m - b.m);
}

function getActiveSections(pks: string[]): ActiveSection[] {
  return SECTIONS
    .map(s => ({ ...s, activeKeys: s.keys.filter(k => pks.includes(k)) }))
    .filter(s => s.activeKeys.length > 0);
}

// ─── Excel Parser ──────────────────────────────────────────────────────────────

async function parseExcel(file: File): Promise<ParsedRow[]> {
  const buf  = await file.arrayBuffer();
  const wb   = XLSX.read(buf, { type: "array", cellDates: false, raw: false });
  const ws   = wb.Sheets[wb.SheetNames[0]];
  const json = XLSX.utils.sheet_to_json<Record<string, unknown>>(ws, { defval: "", raw: true });
  const norm = (s: string) =>
    s.toLowerCase().replace(/[^a-z0-9]/g, " ").replace(/\s+/g, " ").trim();

  return json.map(row => {
    const r: Partial<ParsedRow> = {};
    Object.entries(row).forEach(([k, v]) => {
      const n = norm(k);
      if (n === "date")                                                                     r._d    = v;
      if (n === "plant")                                                                    r._plant= Number(v) || 0;
      if (["functional location","functional loc","funcloc","func location"].includes(n))  r._floc = String(v).trim();
      if (n === "section")                                                                  r._sec  = String(v).toUpperCase().replace(/\s+/g, " ").trim();
      if (n === "head reason" || n === "headreason")                                       r._hr   = String(v).toUpperCase().replace(/\s+/g, " ").trim();
      if (n === "sub head reason" || n === "subheadreason" || n === "sub reason")          r._shr  = String(v).trim();
      if (n.includes("total down") || n === "totaldowntime")                               r._down = Number(v) || 0;
    });
    if (!r._floc) {
      const u = UNITS.find(u => u.plant === r._plant);
      r._floc = u ? u.name : String(r._plant ?? "");
    }
    return r as ParsedRow;
  }).filter(r => r._plant && r._sec);
}

// ─── Multi-Select Dropdown ─────────────────────────────────────────────────────

interface MultiSelectProps {
  label: string; options: string[]; selected: string[];
  onChange: (vals: string[]) => void; disabled?: boolean;
}

function MultiSelect({ label, options, selected, onChange, disabled = false }: MultiSelectProps) {
  const [open, setOpen] = useState(false);
  const ref = useRef<HTMLDivElement>(null);

  useEffect(() => {
    if (!open) return;
    const handler = (e: MouseEvent) => {
      if (ref.current && !ref.current.contains(e.target as Node)) setOpen(false);
    };
    document.addEventListener("mousedown", handler);
    return () => document.removeEventListener("mousedown", handler);
  }, [open]);

  const allSel    = options.length > 0 && selected.length === options.length;
  const toggle    = (v: string) => onChange(selected.includes(v) ? selected.filter(x => x !== v) : [...selected, v]);
  const toggleAll = () => onChange(allSel ? [] : [...options]);
  const display   = selected.length === 0 || allSel ? `All ${label}` : selected.length === 1 ? selected[0] : `${selected.length} selected`;

  return (
    <div ref={ref} style={{ position: "relative", display: "inline-block", minWidth: 140 }}>
      <div style={{ fontSize: 10, color: "#6b7280", marginBottom: 3, fontWeight: 600, textTransform: "uppercase", letterSpacing: "0.05em" }}>
        {label}
      </div>
      <button
        disabled={disabled}
        onClick={() => !disabled && setOpen(o => !o)}
        style={{
          width: "100%", padding: "5px 26px 5px 8px", border: "1px solid #d1d5db",
          borderRadius: 4, background: disabled ? "#f9fafb" : "#fff",
          cursor: disabled ? "not-allowed" : "pointer", textAlign: "left",
          fontSize: 11, fontFamily: "inherit", color: disabled ? "#9ca3af" : "#1f2937",
          whiteSpace: "nowrap", overflow: "hidden", textOverflow: "ellipsis",
          position: "relative", boxShadow: "0 1px 2px rgba(0,0,0,0.05)",
        }}
      >
        {display}
        <span style={{ position: "absolute", right: 7, top: "50%", transform: "translateY(-50%)", fontSize: 8, color: "#6b7280", pointerEvents: "none" }}>▼</span>
      </button>
      {open && !disabled && (
        <div style={{
          position: "absolute", top: "calc(100% + 2px)", left: 0, zIndex: 9999,
          background: "#fff", border: "1px solid #d1d5db", borderRadius: 4,
          boxShadow: "0 4px 12px rgba(0,0,0,0.12)", minWidth: "100%", maxHeight: 200, overflowY: "auto",
        }}>
          <div onClick={toggleAll} style={{
            padding: "5px 10px", cursor: "pointer", fontSize: 11,
            background: allSel ? "#eff6ff" : "transparent", fontWeight: 600,
            display: "flex", alignItems: "center", gap: 6,
            borderBottom: "1px solid #f3f4f6", userSelect: "none",
          }}>
            <input type="checkbox" readOnly checked={allSel} style={{ margin: 0, cursor: "pointer" }} />
            {allSel ? "Deselect All" : "Select All"}
          </div>
          {options.map(o => (
            <div key={o} onClick={() => toggle(o)} style={{
              padding: "4px 10px", cursor: "pointer", fontSize: 11,
              background: selected.includes(o) ? "#eff6ff" : "transparent",
              display: "flex", alignItems: "center", gap: 6, userSelect: "none",
            }}>
              <input type="checkbox" readOnly checked={selected.includes(o)} style={{ margin: 0 }} />
              {o}
            </div>
          ))}
        </div>
      )}
    </div>
  );
}

// ─── Shared Table Styles ───────────────────────────────────────────────────────

const BORDER   = "1px solid #d1d5db";
const baseCell: React.CSSProperties = { border: BORDER, padding: "3px 6px", fontSize: 11, whiteSpace: "nowrap", fontFamily: "inherit" };
const numCell:  React.CSSProperties = { ...baseCell, textAlign: "right" };
const hdrCell:  React.CSSProperties = { ...baseCell, textAlign: "center", fontWeight: 700 };

interface TableHead4RowProps {
  months: MonthMeta[]; activeSections: ActiveSection[];
  bgColor: string; subColor: string;
}

function TableHead4Row({ months, activeSections, bgColor, subColor }: TableHead4RowProps) {
  const totalCols = activeSections.reduce((s, sec) => s + sec.activeKeys.length, 0);
  return (
    <thead>
      <tr>
        <th style={{ ...hdrCell, background: bgColor, minWidth: 65 }} rowSpan={4}>Unit</th>
        <th style={{ ...hdrCell, background: bgColor, minWidth: 48 }} rowSpan={4}>Plant</th>
        {months.map(mk => (
          <th key={mk.label} style={{ ...hdrCell, background: bgColor }} colSpan={totalCols}>{mk.label}</th>
        ))}
      </tr>
      <tr>
        {months.map(mk => (
          <th key={mk.label} style={{ ...hdrCell, background: subColor, fontSize: 10 }} colSpan={totalCols}>{mk.days} Days</th>
        ))}
      </tr>
      <tr>
        {months.map(mk =>
          activeSections.map(s => (
            <th key={`${mk.label}-${s.label}`} style={{ ...hdrCell, background: bgColor }} colSpan={s.activeKeys.length}>{s.label}</th>
          ))
        )}
      </tr>
      <tr>
        {months.map(mk =>
          activeSections.flatMap(s =>
            s.activeKeys.map(k => (
              <th key={`${mk.label}-${k}`} style={{ ...hdrCell, background: subColor, fontSize: 10, minWidth: 52 }}>{PL[k]}</th>
            ))
          )
        )}
      </tr>
    </thead>
  );
}

function SectionTitle({ label, color }: { label: string; color: string }) {
  return (
    <div style={{
      fontWeight: 700, fontSize: 12, margin: "16px 0 5px", color: "#1e3a5f",
      borderLeft: `3px solid ${color}`, paddingLeft: 8, lineHeight: "1.4",
    }}>
      {label}
    </div>
  );
}

// ─── Main Component ────────────────────────────────────────────────────────────

export default function DowntimeAnalysis() {
  const [allRows,  setAllRows]  = useState<ParsedRow[]>([]);
  const [fileName, setFN]       = useState<string>("");
  const [loaded,   setLoaded]   = useState<boolean>(false);
  const [drag,     setDrag]     = useState<boolean>(false);
  const fileRef = useRef<HTMLInputElement>(null);

  const [selFloc,    setSelFloc]    = useState<string[]>([]);
  const [selPlant,   setSelPlant]   = useState<string[]>([]);
  const [selSection, setSelSection] = useState<string[]>([]);
  const [selHR,      setSelHR]      = useState<string[]>([]);
  const [selSHR,     setSelSHR]     = useState<string[]>([]);

  const load = useCallback(async (file: File) => {
    setFN(file.name);
    const rows = await parseExcel(file);
    setAllRows(rows);
    setSelFloc([]); setSelPlant([]); setSelSection([]); setSelHR([]); setSelSHR([]);
    setLoaded(true);
  }, []);

  const reset = () => {
    setLoaded(false); setAllRows([]); setFN("");
    setSelFloc([]); setSelPlant([]); setSelSection([]); setSelHR([]); setSelSHR([]);
  };

  const onDrop = (e: React.DragEvent<HTMLDivElement>) => {
    e.preventDefault(); setDrag(false);
    const f = e.dataTransfer.files[0]; if (f) load(f);
  };

  // ── Cascading filter options ──────────────────────────────────────────────────
  const allFlocs = useMemo(() =>
    [...new Set(allRows.map(r => r._floc).filter(Boolean))].sort(), [allRows]);

  const rowsAfterFloc = useMemo(() =>
    selFloc.length === 0 ? allRows : allRows.filter(r => selFloc.includes(r._floc)), [allRows, selFloc]);

  const allPlants = useMemo(() =>
    [...new Set(rowsAfterFloc.map(r => String(r._plant)).filter(Boolean))].sort((a, b) => +a - +b), [rowsAfterFloc]);

  const rowsAfterPlant = useMemo(() =>
    selPlant.length === 0 ? rowsAfterFloc : rowsAfterFloc.filter(r => selPlant.includes(String(r._plant))), [rowsAfterFloc, selPlant]);

  const allSections = useMemo(() =>
    [...new Set(rowsAfterPlant.map(r => r._sec).filter(Boolean))].sort(), [rowsAfterPlant]);

  const rowsAfterSection = useMemo(() =>
    selSection.length === 0 ? rowsAfterPlant : rowsAfterPlant.filter(r => selSection.includes(r._sec)), [rowsAfterPlant, selSection]);

  const allHRs = useMemo(() =>
    [...new Set(rowsAfterSection.map(r => r._hr).filter(Boolean))].sort(), [rowsAfterSection]);

  const rowsAfterHR = useMemo(() =>
    selHR.length === 0 ? rowsAfterSection : rowsAfterSection.filter(r => selHR.includes(r._hr)), [rowsAfterSection, selHR]);

  const allSHRs = useMemo(() =>
    [...new Set(rowsAfterHR.map(r => r._shr).filter(Boolean))].sort(), [rowsAfterHR]);

  // ── Final filtered rows ───────────────────────────────────────────────────────
  const filteredRows = useMemo<ParsedRow[]>(() => {
    let rows = allRows;
    if (selFloc.length > 0)    rows = rows.filter(r => selFloc.includes(r._floc));
    if (selPlant.length > 0)   rows = rows.filter(r => selPlant.includes(String(r._plant)));
    if (selSection.length > 0) rows = rows.filter(r => selSection.includes(r._sec));
    if (selHR.length > 0)      rows = rows.filter(r => selHR.includes(r._hr));
    if (selSHR.length > 0)     rows = rows.filter(r => selSHR.includes(r._shr));
    return rows;
  }, [allRows, selFloc, selPlant, selSection, selHR, selSHR]);

  const months = useMemo(() => getMonths(filteredRows), [filteredRows]);

  // ── Derived maps ──────────────────────────────────────────────────────────────
  const { idleMap, hrMaps, shrMap, activePks, activeUnits } = useMemo(() => {
    const idleM: HrMap = {};
    const hrM: Record<string, HrMap> = { PLANNED: {}, UNPLANNED: {}, OTHERS: {} };
    // shrMap[bucket][shr][plant][pk][month] = down  (IDLE bucket included)
    const shrM: ShrMap = { IDLE: {}, PLANNED: {}, UNPLANNED: {}, OTHERS: {} };

    const seenPks    = new Set<string>();
    const seenPlants = new Set<number>();

    filteredRows.forEach(r => {
      const pk = SEC_MAP[r._sec];
      if (pk) seenPks.add(pk);
      seenPlants.add(r._plant);

      const ymd = parseDateOnly(r._d); if (!ymd) return;
      const lbl = `${String(ymd.m).padStart(2, "0")}/${String(ymd.y).slice(-2)}`;
      if (!pk) return;
      const p   = r._plant;
      const hr  = r._hr;
      const shr = r._shr || "Unknown";

      const addToHrMap = (map: HrMap) => {
        if (!map[p])     map[p]     = {};
        if (!map[p][pk]) map[p][pk] = {};
        map[p][pk][lbl] = (map[p][pk][lbl] ?? 0) + r._down;
      };

      const addToShrMap = (bucket: string) => {
        if (!shrM[bucket])           shrM[bucket]           = {};
        if (!shrM[bucket][shr])      shrM[bucket][shr]      = {};
        if (!shrM[bucket][shr][p])   shrM[bucket][shr][p]   = {};
        if (!shrM[bucket][shr][p][pk]) shrM[bucket][shr][p][pk] = {};
        shrM[bucket][shr][p][pk][lbl] = (shrM[bucket][shr][p][pk][lbl] ?? 0) + r._down;
      };

      if (hr === "IDLE") {
        addToHrMap(idleM); addToShrMap("IDLE");
      } else if (hr === "PLANNED") {
        addToHrMap(hrM.PLANNED); addToShrMap("PLANNED");
      } else if (hr === "UNPLANNED") {
        addToHrMap(hrM.UNPLANNED); addToShrMap("UNPLANNED");
      } else {
        addToHrMap(hrM.OTHERS); addToShrMap("OTHERS");
      }
    });

    const orderedPks = ALL_KEYS.filter(k => seenPks.has(k));

    const activeU = UNITS.filter(u => {
      if (selFloc.length > 0  && !selFloc.includes(u.name))           return false;
      if (selPlant.length > 0 && !selPlant.includes(String(u.plant))) return false;
      if (selFloc.length === 0 && selPlant.length === 0)               return seenPlants.has(u.plant);
      return true;
    });

    return { idleMap: idleM, hrMaps: hrM, shrMap: shrM, activePks: orderedPks, activeUnits: activeU };
  }, [filteredRows, selFloc, selPlant]);

  // Filter handlers
  const handleFlocChange    = (v: string[]) => { setSelFloc(v);    setSelPlant([]); setSelSection([]); setSelHR([]); setSelSHR([]); };
  const handlePlantChange   = (v: string[]) => { setSelPlant(v);   setSelSection([]); setSelHR([]); setSelSHR([]); };
  const handleSectionChange = (v: string[]) => { setSelSection(v); setSelHR([]); setSelSHR([]); };
  const handleHRChange      = (v: string[]) => { setSelHR(v);      setSelSHR([]); };
  const clearAll = () => { setSelFloc([]); setSelPlant([]); setSelSection([]); setSelHR([]); setSelSHR([]); };
  const hasFilters = selFloc.length > 0 || selPlant.length > 0 || selSection.length > 0 || selHR.length > 0 || selSHR.length > 0;

  // ─── Core pct calculation helpers ────────────────────────────────────────────

  // Total machine-hours for a unit across all pks and given months
  function totalHrs(unitName: string, mks: MonthMeta[]): number {
    return activePks.reduce((s, pk) => {
      return s + mks.reduce((ms, mk) => ms + (MC[unitName]?.[pk] ?? 0) * mk.days * 24, 0);
    }, 0);
  }

  // Idle downtime for a unit across all pks and given months
  function idleHrs(plant: number, mks: MonthMeta[]): number {
    return activePks.reduce((s, pk) => {
      return s + mks.reduce((ms, mk) => ms + (idleMap[plant]?.[pk]?.[mk.label] ?? 0), 0);
    }, 0);
  }

  // idleRemaining = totalHrs - idleHrs
  function idleRemainingHrs(unitName: string, plant: number, mks: MonthMeta[]): number {
    return totalHrs(unitName, mks) - idleHrs(plant, mks);
  }

  // Downtime for a given HR bucket, unit, and months (sum across all pks)
  function hrDownHrs(bucket: string, plant: number, mks: MonthMeta[]): number {
    const map = bucket === "IDLE" ? idleMap : (hrMaps[bucket] ?? {});
    return activePks.reduce((s, pk) => {
      return s + mks.reduce((ms, mk) => ms + (map[plant]?.[pk]?.[mk.label] ?? 0), 0);
    }, 0);
  }

  // SHR downtime for a specific SHR in a given bucket
  function shrDownHrs(bucket: string, shr: string, plant: number, mks: MonthMeta[]): number {
    const sm = shrMap[bucket]?.[shr] ?? {};
    return activePks.reduce((s, pk) => {
      return s + mks.reduce((ms, mk) => ms + (sm[plant]?.[pk]?.[mk.label] ?? 0), 0);
    }, 0);
  }

  // % for non-IDLE buckets = down / idleRemaining * 100
  // % for IDLE bucket      = down / totalHrs * 100
  function calcPct(down: number, unitName: string, plant: number, mks: MonthMeta[], isIdle = false): number | null {
    if (isIdle) {
      const tot = totalHrs(unitName, mks);
      if (tot <= 0 || down === 0) return null;
      return (down / tot) * 100;
    }
    const rem = idleRemainingHrs(unitName, plant, mks);
    if (rem <= 0 || down === 0) return null;
    return (down / rem) * 100;
  }

  // Format pct value for display
  function fmtPct(v: number | null): string {
    if (v === null) return "-";
    return (Math.trunc(v * 1000) / 1000).toString();
}

  // ─── Table Renderers (existing) ────────────────────────────────────────────────

  function renderMachineCountTable(): React.ReactNode {
    const displayPks = activePks.length > 0 ? activePks : ALL_KEYS;
    const activeSecs = getActiveSections(displayPks);
    return (
      <div style={{ overflowX: "auto", marginBottom: 4 }}>
        <table style={{ borderCollapse: "collapse" }}>
          <thead>
            <tr>
              <th style={{ ...hdrCell, background: "#dce6f1", minWidth: 65 }} rowSpan={2}>Unit Name</th>
              <th style={{ ...hdrCell, background: "#dce6f1", minWidth: 48 }} rowSpan={2}>Plant</th>
              {activeSecs.map(s => (
                <th key={s.label} style={{ ...hdrCell, background: "#dce6f1" }} colSpan={s.activeKeys.length}>{s.label}</th>
              ))}
            </tr>
            <tr>
              {activeSecs.flatMap(s => s.activeKeys.map(k => (
                <th key={k} style={{ ...hdrCell, background: "#bdd7ee", fontSize: 10, minWidth: 52 }}>{PL[k]}</th>
              )))}
            </tr>
          </thead>
          <tbody>
            {activeUnits.map((u, i) => (
              <tr key={u.name} style={{ background: i % 2 === 1 ? "#f8fafc" : "#fff" }}>
                <td style={{ ...baseCell, fontWeight: 600 }}>{u.name}</td>
                <td style={{ ...numCell, textAlign: "center", color: "#6b7280" }}>{u.plant}</td>
                {activeSecs.flatMap(s => s.activeKeys.map(k => (
                  <td key={k} style={numCell}>{MC[u.name]?.[k] ?? 0}</td>
                )))}
              </tr>
            ))}
          </tbody>
        </table>
      </div>
    );
  }

  function renderTotalHrsTable(): React.ReactNode {
    if (!activePks.length || !months.length) return null;
    const activeSecs = getActiveSections(activePks);
    return (
      <div style={{ overflowX: "auto", marginBottom: 4 }}>
        <table style={{ borderCollapse: "collapse" }}>
          <TableHead4Row months={months} activeSections={activeSecs} bgColor="#dce6f1" subColor="#bdd7ee" />
          <tbody>
            {activeUnits.map((u, i) => (
              <tr key={u.name} style={{ background: i % 2 === 1 ? "#f8fafc" : "#fff" }}>
                <td style={{ ...baseCell, fontWeight: 600 }}>{u.name}</td>
                <td style={{ ...numCell, textAlign: "center", color: "#6b7280" }}>{u.plant}</td>
                {months.map(mk =>
                  activeSecs.flatMap(s => s.activeKeys.map(k => {
                    const total = (MC[u.name]?.[k] ?? 0) * mk.days * 24;
                    return (
                      <td key={`${mk.label}-${k}`} style={{ ...numCell, color: total === 0 ? "#d1d5db" : "#111" }}>
                        {total === 0 ? "-" : total}
                      </td>
                    );
                  }))
                )}
              </tr>
            ))}
          </tbody>
        </table>
      </div>
    );
  }

  function renderIdleHrsTable(): React.ReactNode {
    if (!activePks.length || !months.length) return null;
    const activeSecs = getActiveSections(activePks);
    return (
      <div style={{ overflowX: "auto", marginBottom: 4 }}>
        <table style={{ borderCollapse: "collapse" }}>
          <TableHead4Row months={months} activeSections={activeSecs} bgColor="#fce4d6" subColor="#f4b183" />
          <tbody>
            {activeUnits.map((u, i) => (
              <tr key={u.name} style={{ background: i % 2 === 1 ? "#f8fafc" : "#fff" }}>
                <td style={{ ...baseCell, fontWeight: 600 }}>{u.name}</td>
                <td style={{ ...numCell, textAlign: "center", color: "#6b7280" }}>{u.plant}</td>
                {months.map(mk =>
                  activeSecs.flatMap(s => s.activeKeys.map(k => {
                    const v = idleMap[u.plant]?.[k]?.[mk.label] ?? 0;
                    return (
                      <td key={`${mk.label}-${k}`} style={{ ...numCell, color: v === 0 ? "#d1d5db" : "#111" }}>
                        {v === 0 ? "-" :trunc3(v)}
                      </td>
                    );
                  }))
                )}
              </tr>
            ))}
          </tbody>
        </table>
      </div>
    );
  }

  function renderRemainingHrsTable(): React.ReactNode {
    if (!activePks.length || !months.length) return null;
    const activeSecs = getActiveSections(activePks);
    return (
      <div style={{ overflowX: "auto", marginBottom: 4 }}>
        <table style={{ borderCollapse: "collapse" }}>
          <TableHead4Row months={months} activeSections={activeSecs} bgColor="#fdf3ee" subColor="#f9cdb0" />
          <tbody>
            {activeUnits.map((u, i) => (
              <tr key={u.name} style={{ background: i % 2 === 1 ? "#f8fafc" : "#fff" }}>
                <td style={{ ...baseCell, fontWeight: 600 }}>{u.name}</td>
                <td style={{ ...numCell, textAlign: "center", color: "#6b7280" }}>{u.plant}</td>
                {months.map(mk =>
                  activeSecs.flatMap(s => s.activeKeys.map(k => {
                    const mc        = MC[u.name]?.[k] ?? 0;
                    const total     = mc * mk.days * 24;
                    const idleDown  = idleMap[u.plant]?.[k]?.[mk.label] ?? 0;
                    const remaining = total - idleDown;
                    const isNeg     = remaining < 0;
                    return (
                      <td key={`${mk.label}-${k}`} style={{ ...numCell, color: total === 0 ? "#d1d5db" : isNeg ? "#dc2626" : "#111", fontWeight: isNeg ? 700 : "normal" }}>
                        {total === 0 ? "-" : trunc3(remaining)}
                      </td>
                    );
                  }))
                )}
              </tr>
            ))}
          </tbody>
        </table>
      </div>
    );
  }

  function renderPctIdleTable(): React.ReactNode {
    if (!activePks.length || !months.length) return null;
    const activeSecs = getActiveSections(activePks);
    return (
      <div style={{ overflowX: "auto", marginBottom: 4 }}>
        <table style={{ borderCollapse: "collapse" }}>
          <TableHead4Row months={months} activeSections={activeSecs} bgColor="#fce9df" subColor="#f7c4a0" />
          <tbody>
            {activeUnits.map((u, i) => (
              <tr key={u.name} style={{ background: i % 2 === 1 ? "#f8fafc" : "#fff" }}>
                <td style={{ ...baseCell, fontWeight: 600 }}>{u.name}</td>
                <td style={{ ...numCell, textAlign: "center", color: "#6b7280" }}>{u.plant}</td>
                {months.map(mk =>
                  activeSecs.flatMap(s => s.activeKeys.map(k => {
                    const mc        = MC[u.name]?.[k] ?? 0;
                    const total     = mc * mk.days * 24;
                    const idleDown  = idleMap[u.plant]?.[k]?.[mk.label] ?? 0;
                    const remaining = total - idleDown;
                    const pct       = total === 0 || remaining === 0 || idleDown === 0 ? null : (idleDown / remaining) * 100;
                    const isHigh    = pct !== null && pct > 100;
                    return (
                      <td key={`${mk.label}-${k}`} style={{ ...numCell, color: pct === null ? "#d1d5db" : isHigh ? "#dc2626" : "#111", fontWeight: isHigh ? 700 : "normal" }}>
{pct === null ? "-" : `${Math.trunc(pct * 1000) / 1000}%`}
</td>
                    );
                  }))
                )}
              </tr>
            ))}
          </tbody>
        </table>
      </div>
    );
  }

  function renderHrHrsTable(hrKey: string, bgColor: string, subColor: string): React.ReactNode {
    if (!activePks.length || !months.length) return null;
    const activeSecs = getActiveSections(activePks);
    const map = hrMaps[hrKey] ?? {};
    return (
      <div style={{ overflowX: "auto", marginBottom: 4 }}>
        <table style={{ borderCollapse: "collapse" }}>
          <TableHead4Row months={months} activeSections={activeSecs} bgColor={bgColor} subColor={subColor} />
          <tbody>
            {activeUnits.map((u, i) => (
              <tr key={u.name} style={{ background: i % 2 === 1 ? "#f8fafc" : "#fff" }}>
                <td style={{ ...baseCell, fontWeight: 600 }}>{u.name}</td>
                <td style={{ ...numCell, textAlign: "center", color: "#6b7280" }}>{u.plant}</td>
                {months.map(mk =>
                  activeSecs.flatMap(s => s.activeKeys.map(k => {
                    const v = map[u.plant]?.[k]?.[mk.label] ?? 0;
                    return (
                      <td key={`${mk.label}-${k}`} style={{ ...numCell, color: v === 0 ? "#d1d5db" : "#111" }}>
                        {v === 0 ? "-" : trunc3(v)}
                      </td>
                    );
                  }))
                )}
              </tr>
            ))}
          </tbody>
        </table>
      </div>
    );
  }

  function renderPctOfIdleRemainingTable(hrKey: string, bgColor: string, subColor: string): React.ReactNode {
    if (!activePks.length || !months.length) return null;
    const activeSecs = getActiveSections(activePks);
    const map = hrMaps[hrKey] ?? {};
    return (
      <div style={{ overflowX: "auto", marginBottom: 4 }}>
        <table style={{ borderCollapse: "collapse" }}>
          <TableHead4Row months={months} activeSections={activeSecs} bgColor={bgColor} subColor={subColor} />
          <tbody>
            {activeUnits.map((u, i) => (
              <tr key={u.name} style={{ background: i % 2 === 1 ? "#f8fafc" : "#fff" }}>
                <td style={{ ...baseCell, fontWeight: 600 }}>{u.name}</td>
                <td style={{ ...numCell, textAlign: "center", color: "#6b7280" }}>{u.plant}</td>
                {months.map(mk =>
                  activeSecs.flatMap(s => s.activeKeys.map(k => {
                    const mc       = MC[u.name]?.[k] ?? 0;
                    const total    = mc * mk.days * 24;
                    const idleDown = idleMap[u.plant]?.[k]?.[mk.label] ?? 0;
                    const idleRem  = total - idleDown;
                    const hrDown   = map[u.plant]?.[k]?.[mk.label] ?? 0;
                    const pct      = total === 0 || idleRem <= 0 || hrDown === 0 ? null : (hrDown / idleRem) * 100;
                    const isHigh   = pct !== null && pct > 100;
                    return (
                      <td key={`${mk.label}-${k}`} style={{ ...numCell, color: pct === null ? "#d1d5db" : isHigh ? "#dc2626" : "#111", fontWeight: isHigh ? 700 : "normal" }}>
                        {pct === null ? "-" : `${ Math.trunc(pct * 1000) / 1000}%`}
                      </td>
                    );
                  }))
                )}
              </tr>
            ))}
          </tbody>
        </table>
      </div>
    );
  }

  // ─── Consolidated HR Summary Table ────────────────────────────────────────────
  //
  // Row order: IDLE SHRs → IDLE total (salmon) → PLANNED SHRs → PLANNED total (green)
  //            → UNPLANNED SHRs → UNPLANNED total (yellow) → OTHERS SHRs → OTHERS total (grey)
  //            → Combined rows (blue)
  //
  // % denominator:
  //   IDLE rows  → down / Total Hrs * 100
  //   Other rows → down / Idle Remaining Hrs * 100
  //
  function renderConsolidatedTable(): React.ReactNode {
    if (!activePks.length || !months.length || !activeUnits.length) return null;

    // ── Build FY groups ──────────────────────────────────────────────────────────
    const fyOrder: string[] = [];
    const fyMonths: Record<string, MonthMeta[]> = {};
    months.forEach(mk => {
      const fy = getFYLabel(mk);
      if (!fyMonths[fy]) { fyMonths[fy] = []; fyOrder.push(fy); }
      fyMonths[fy].push(mk);
    });

    const completeFYs  = fyOrder.slice(0, -1);
    const currentFY    = fyOrder[fyOrder.length - 1];
    const currentMks   = fyMonths[currentFY];

    const colsPerUnit = completeFYs.length + currentMks.length + 1;

    // ── Styles ────────────────────────────────────────────────────────────────────
    const B = BORDER;
    const lbl: React.CSSProperties  = { border: B, padding: "3px 7px", fontSize: 11, whiteSpace: "nowrap", fontFamily: "inherit", textAlign: "left" };
    const num2: React.CSSProperties = { border: B, padding: "3px 6px", fontSize: 11, whiteSpace: "nowrap", fontFamily: "inherit", textAlign: "right", minWidth: 48 };
    const hdr2: React.CSSProperties = { border: B, padding: "3px 6px", fontSize: 11, whiteSpace: "nowrap", fontFamily: "inherit", textAlign: "center", fontWeight: 700 };

    // ── Render one data cell ──────────────────────────────────────────────────────
    function cell(
      down: number,
      unitName: string,
      plant: number,
      mks: MonthMeta[],
      style: React.CSSProperties,
      key: string,
      isIdle: boolean
    ): React.ReactNode {
      const pct  = calcPct(down, unitName, plant, mks, isIdle);
      const high = pct !== null && pct > 100;
      return (
        <td key={key} style={{
          ...num2, ...style,
          color: pct === null ? "#d1d5db" : high ? "#dc2626" : undefined,
          fontWeight: high ? 700 : (style.fontWeight ?? "normal"),
        }}>
          {fmtPct(pct)}
        </td>
      );
    }

    // ── Render one unit's columns ──────────────────────────────────────────────
    function unitCols(
      getDown: (mks: MonthMeta[]) => number,
      unitName: string,
      plant: number,
      rowStyle: React.CSSProperties,
      keyPrefix: string,
      isIdle: boolean
    ): React.ReactNode {
      const nodes: React.ReactNode[] = [];
      completeFYs.forEach(fy => {
        const mks  = fyMonths[fy];
        const down = getDown(mks);
        nodes.push(cell(down, unitName, plant, mks, rowStyle, `${keyPrefix}-${fy}`, isIdle));
      });
      currentMks.forEach(mk => {
        const down = getDown([mk]);
        nodes.push(cell(down, unitName, plant, [mk], rowStyle, `${keyPrefix}-${mk.label}`, isIdle));
      });
      const utdDown = getDown(currentMks);
      nodes.push(cell(utdDown, unitName, plant, currentMks, { ...rowStyle, background: (rowStyle.background ?? "#f0f4fb") }, `${keyPrefix}-utd`, isIdle));
      return nodes;
    }

    // ── Collect all SHRs per bucket in sorted order ────────────────────────────
    const bucketSHRs: Record<string, string[]> = {};
    // Include IDLE in the SHR collection
    ["IDLE", ...HR_ORDER].forEach(bucket => {
      bucketSHRs[bucket] = Object.keys(shrMap[bucket] ?? {}).sort();
    });

    // ── Build row definitions ─────────────────────────────────────────────────
    interface ConsolRow {
      type: "shr" | "hr" | "combined";
      label: string;
      rowNo?: string;
      bucket?: string;
      shr?: string;
      buckets?: string[];
      style: React.CSSProperties;
      isIdle: boolean;
    }

    const rows: ConsolRow[] = [];
    let rowNo = 1;

    // ── IDLE block first (salmon) ──────────────────────────────────────────────
    const idleSHRs = bucketSHRs["IDLE"];
    idleSHRs.forEach(shr => {
      rows.push({ type: "shr", label: shr, bucket: "IDLE", shr, style: { background: "#fff" }, isIdle: true });
    });
    rows.push({
      type: "hr", label: "Idle", rowNo: String(rowNo++), bucket: "IDLE",
      style: { background: CONSOL_HR_STYLE.IDLE.rowBg, color: CONSOL_HR_STYLE.IDLE.rowText, fontWeight: 700 },
      isIdle: true,
    });

    // ── PLANNED / UNPLANNED / OTHERS blocks ───────────────────────────────────
    HR_ORDER.forEach(bucket => {
      const cfg  = CONSOL_HR_STYLE[bucket];
      const shrs = bucketSHRs[bucket];

      shrs.forEach(shr => {
        rows.push({ type: "shr", label: shr, bucket, shr, style: { background: "#fff" }, isIdle: false });
      });
      rows.push({
        type: "hr", label: HR_CONFIG[bucket].label, rowNo: String(rowNo++), bucket,
        style: { background: cfg.rowBg, color: cfg.rowText, fontWeight: 700 },
        isIdle: false,
      });
    });

    // Combined rows (always non-IDLE denominator)
    rows.push({
      type: "combined", label: "Unplanned + Others (2+3)",
      rowNo: String(rowNo++), buckets: ["UNPLANNED", "OTHERS"],
      style: { ...CONSOL_COMBINED_STYLE, fontWeight: 700 },
      isIdle: false,
    });
    rows.push({
      type: "combined", label: "Planned + Unplanned + Others (1+2+3)",
      rowNo: String(rowNo++), buckets: ["PLANNED", "UNPLANNED", "OTHERS"],
      style: { ...CONSOL_COMBINED_STYLE, fontWeight: 700 },
      isIdle: false,
    });

    // ── Render ──────────────────────────────────────────────────────────────────
    return (
      <div style={{ overflowX: "auto", marginBottom: 4 }}>
        <table style={{ borderCollapse: "collapse" }}>
          <thead>
            <tr>
              <th style={{ ...hdr2, background: "#dce6f1", minWidth: 200 }} rowSpan={3}>Head / Sub Head Reason</th>
              {activeUnits.map(u => (
                <th key={u.name} style={{ ...hdr2, background: "#dce6f1" }} colSpan={colsPerUnit}>{u.name}</th>
              ))}
            </tr>
            <tr>
              {activeUnits.map(u => (
                <Fragment key={u.name}>
                  {completeFYs.map(fy => (
                    <th key={fy} style={{ ...hdr2, background: "#bdd7ee", minWidth: 48 }}>{fy}</th>
                  ))}
                  {currentMks.map(mk => (
                    <th key={mk.label} style={{ ...hdr2, background: "#bdd7ee", minWidth: 48 }}>
                      {new Date(mk.y, mk.m - 1).toLocaleString("en-GB", { month: "short" })}-{String(mk.y).slice(-2)}
                    </th>
                  ))}
                  <th style={{ ...hdr2, background: "#dce6f1", minWidth: 52 }}>{currentFY} UTD Avg</th>
                </Fragment>
              ))}
            </tr>
            <tr>
              {activeUnits.map(u => (
                <Fragment key={u.name}>
                  {completeFYs.map(fy => (
                    <th key={fy} style={{ ...hdr2, background: "#dce6f1", fontSize: 10 }}>{fy}</th>
                  ))}
                  {currentMks.map(mk => (
                    <th key={mk.label} style={{ ...hdr2, background: "#bdd7ee", fontSize: 10 }}>
                      {new Date(mk.y, mk.m - 1).toLocaleString("en-GB", { month: "short" })}-{String(mk.y).slice(-2)}
                    </th>
                  ))}
                  <th style={{ ...hdr2, background: "#c5d9f1", fontSize: 10 }}>UTD Avg</th>
                </Fragment>
              ))}
            </tr>
          </thead>
          <tbody>
            {rows.map((row, ri) => (
              <tr key={ri} style={{ background: (row.style.background as string) ?? "#fff" }}>
                <td style={{
                  ...lbl,
                  background: (row.style.background as string) ?? "#fff",
                  color: (row.style.color as string) ?? "#111",
                  fontWeight: row.style.fontWeight ?? "normal",
                  paddingLeft: row.type === "shr" ? 20 : 8,
                }}>
                  {row.rowNo ? <span style={{ color: "#6b7280", marginRight: 5 }}>{row.rowNo}</span> : null}
                  {row.label}
                </td>

                {activeUnits.map(u => {
                  const getDown = (mks: MonthMeta[]): number => {
                    if (row.type === "shr") {
                      return shrDownHrs(row.bucket!, row.shr!, u.plant, mks);
                    } else if (row.type === "hr") {
                      return hrDownHrs(row.bucket!, u.plant, mks);
                    } else {
                      return (row.buckets ?? []).reduce((s, b) => s + hrDownHrs(b, u.plant, mks), 0);
                    }
                  };
                  return (
                    <Fragment key={u.name}>
                      {unitCols(getDown, u.name, u.plant, row.style, `${ri}-${u.name}`, row.isIdle)}
                    </Fragment>
                  );
                })}
              </tr>
            ))}
          </tbody>
        </table>
      </div>
    );
  }

  // ─── Render ───────────────────────────────────────────────────────────────────

  return (
    <div style={{ padding: 16, background: "#fff", fontFamily: "'Segoe UI', system-ui, sans-serif", fontSize: 11, minHeight: "100vh" }}>

      {/* Header */}
      <div style={{ display: "flex", alignItems: "center", gap: 10, marginBottom: 14, paddingBottom: 10, borderBottom: "2px solid #e5e7eb" }}>
        <div style={{ fontWeight: 800, fontSize: 15, color: "#1e3a5f", letterSpacing: "-0.02em" }}>
          Downtime Analysis
        </div>
        {loaded && <span style={{ color: "#6b7280", fontSize: 11, fontStyle: "italic" }}>— {fileName}</span>}
        {loaded && (
          <button onClick={reset} style={{ marginLeft: "auto", fontSize: 11, padding: "3px 12px", cursor: "pointer", background: "#fff", border: "1px solid #e5e7eb", borderRadius: 4, color: "#6b7280", fontWeight: 500 }}>
            ✕ Reset
          </button>
        )}
      </div>

      {/* Upload Zone */}
      {!loaded && (
        <div
          onDragOver={e => { e.preventDefault(); setDrag(true); }}
          onDragLeave={() => setDrag(false)}
          onDrop={onDrop}
          onClick={() => fileRef.current?.click()}
          style={{
            border: `2px dashed ${drag ? "#3b82f6" : "#d1d5db"}`,
            background: drag ? "#eff6ff" : "#f9fafb",
            borderRadius: 8, padding: "60px 0", textAlign: "center",
            cursor: "pointer", transition: "all 0.15s",
          }}
        >
          <input ref={fileRef} type="file" accept=".xlsx,.xls,.csv" style={{ display: "none" }}
            onChange={e => { const f = e.target.files?.[0]; if (f) load(f); }} />
          <div style={{ fontSize: 36, marginBottom: 8 }}>📊</div>
          <div style={{ fontWeight: 700, fontSize: 13, color: "#374151", marginBottom: 4 }}>
            Drop your Excel file here or click to browse
          </div>
          <div style={{ color: "#9ca3af", fontSize: 11 }}>.xlsx · .xls · .csv</div>
        </div>
      )}

      {/* Main Content */}
      {loaded && (
        <>
          {/* Filter Bar */}
          <div style={{
            background: "#f8fafc", border: "1px solid #e2e8f0", borderRadius: 6,
            padding: "10px 12px", marginBottom: 14,
            display: "flex", flexWrap: "wrap", gap: 12, alignItems: "flex-end",
          }}>
            <MultiSelect label="Functional Location" options={allFlocs}    selected={selFloc}    onChange={handleFlocChange} />
            <MultiSelect label="Plant"                options={allPlants}   selected={selPlant}   onChange={handlePlantChange}   disabled={allPlants.length === 0} />
            <MultiSelect label="Section"              options={allSections} selected={selSection} onChange={handleSectionChange} disabled={allSections.length === 0} />
            <MultiSelect label="Head Reason"          options={allHRs}      selected={selHR}      onChange={handleHRChange}      disabled={allHRs.length === 0} />
            <MultiSelect label="Sub Head Reason"      options={allSHRs}     selected={selSHR}     onChange={v => setSelSHR(v)}   disabled={allSHRs.length === 0} />
            {hasFilters && (
              <button onClick={clearAll} style={{ fontSize: 11, padding: "5px 12px", cursor: "pointer", background: "#fff", border: "1px solid #fca5a5", borderRadius: 4, color: "#dc2626", fontWeight: 600, alignSelf: "flex-end" }}>
                ✕ Clear All
              </button>
            )}
          </div>

          <div style={{ overflowX: "auto" }}>

            {/* 1. Machine Count */}
            <SectionTitle label="Machine Count" color="#3b82f6" />
            {renderMachineCountTable()}

            {/* 2. Total Hrs */}
            <SectionTitle label="Total Hrs." color="#3b82f6" />
            {renderTotalHrsTable()}

            {/* 3. Idle Hrs */}
            <SectionTitle label="Idle Hrs." color="#f97316" />
            {renderIdleHrsTable()}

            {/* 4. Remaining Hrs (after Idle) */}
            <SectionTitle label="Remaining Hrs. (after Idle)" color="#f97316" />
            {renderRemainingHrsTable()}

            {/* 5. % Idle Hrs / Total Hrs */}
            <SectionTitle label="% Idle Hrs. / Total Hrs." color="#f97316" />
            {renderPctIdleTable()}

            {/* 6–11. Per HR blocks */}
            {HR_ORDER.map(hrKey => {
              const cfg     = HR_CONFIG[hrKey];
              const hasData = Object.keys(hrMaps[hrKey] ?? {}).length > 0;
              if (!hasData) return null;
              return (
                <div key={hrKey}>
                  <SectionTitle label={`${cfg.label} Hrs.`} color={cfg.sub} />
                  {renderHrHrsTable(hrKey, cfg.bg, cfg.sub)}

                  <SectionTitle label={`% ${cfg.label} Hrs. / Idle Remaining Hrs.`} color={cfg.pctSub} />
                  {renderPctOfIdleRemainingTable(hrKey, cfg.pctBg, cfg.pctSub)}
                </div>
              );
            })}

            {/* 12. Consolidated HR Summary */}
            <SectionTitle label="Consolidated HR Summary" color="#1e3a5f" />
            {renderConsolidatedTable()}

          </div>
        </>
      )}
    </div>
  );
}