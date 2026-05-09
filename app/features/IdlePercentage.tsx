"use client";
import React, { useState, useCallback, useRef, useMemo } from "react";
import * as XLSX from "xlsx";

// ─── Static Data ──────────────────────────────────────────────────────────────

const UNITS = [
  { name: "NGD",    plant: 1101 },
  { name: "BCK",    plant: 1201 },
  { name: "HRR",    plant: 1301 },
  { name: "VIL-1",  plant: 1601 },
  { name: "VIL-2",  plant: 1602 },
  { name: "IBR",    plant: 2111 },
  { name: "TRC",    plant: 4101 },
];

const MC: Record<string, Record<string, number>> = {
  NGD:    { spinning:11,dryer:11,bailingPress:13,simplex:22,churn:0, twinRollPress:15,firstStageGCF:42,secondStageGCF:31,thirdStageGCF:32,rejectGCF:4, msfe:17,aac:5,anhydrousEvaporator:6,acidPlant:4,cs2Furnace:16,cs2Plant:0,wsa:0 },
  BCK:    { spinning:4, dryer:4, bailingPress:5, simplex:12,churn:0, twinRollPress:14,firstStageGCF:38,secondStageGCF:26,thirdStageGCF:27,rejectGCF:8, msfe:9, aac:4,anhydrousEvaporator:5,acidPlant:2,cs2Furnace:13,cs2Plant:0,wsa:0 },
  HRR:    { spinning:3, dryer:3, bailingPress:4, simplex:9, churn:2, twinRollPress:9, firstStageGCF:20,secondStageGCF:16,thirdStageGCF:16,rejectGCF:6, msfe:8, aac:3,anhydrousEvaporator:4,acidPlant:1,cs2Furnace:9, cs2Plant:0,wsa:0 },
  "VIL-1":{ spinning:4, dryer:4, bailingPress:5, simplex:0, churn:8, twinRollPress:17,firstStageGCF:41,secondStageGCF:34,thirdStageGCF:34,rejectGCF:11,msfe:10,aac:7,anhydrousEvaporator:6,acidPlant:1,cs2Furnace:0, cs2Plant:1,wsa:0 },
  "VIL-2":{ spinning:2, dryer:2, bailingPress:5, simplex:0, churn:10,twinRollPress:16,firstStageGCF:28,secondStageGCF:24,thirdStageGCF:24,rejectGCF:8, msfe:10,aac:4,anhydrousEvaporator:5,acidPlant:1,cs2Furnace:0, cs2Plant:1,wsa:1 },
  IBR:    { spinning:5, dryer:5, bailingPress:7, simplex:18,churn:4, twinRollPress:20,firstStageGCF:43,secondStageGCF:30,thirdStageGCF:31,rejectGCF:8, msfe:13,aac:8,anhydrousEvaporator:6,acidPlant:3,cs2Furnace:0, cs2Plant:1,wsa:2 },
  TRC:    { spinning:4, dryer:4, bailingPress:7, simplex:19,churn:3, twinRollPress:16,firstStageGCF:39,secondStageGCF:25,thirdStageGCF:29,rejectGCF:10,msfe:10,aac:7,anhydrousEvaporator:6,acidPlant:3,cs2Furnace:0, cs2Plant:1,wsa:0 },
};

const SECTIONS = [
  { label:"Spinning",  keys:["spinning","dryer","bailingPress"] },
  { label:"Viscose",   keys:["simplex","churn","twinRollPress","firstStageGCF","secondStageGCF","thirdStageGCF","rejectGCF","msfe"] },
  { label:"Auxiliary", keys:["aac","anhydrousEvaporator","acidPlant"] },
  { label:"Ancillary", keys:["cs2Furnace","cs2Plant","wsa"] },
];

const PL: Record<string,string> = {
  spinning:"Spinning",dryer:"Dryer",bailingPress:"Bailing Press",
  simplex:"Simplex",churn:"Churn",twinRollPress:"Twin Roll Press",
  firstStageGCF:"1st Stage GCF",secondStageGCF:"2nd Stage GCF",thirdStageGCF:"3rd Stage GCF",
  rejectGCF:"Reject GCF",msfe:"MSFE",
  aac:"AAC",anhydrousEvaporator:"Anhydrous Evaporator",acidPlant:"Acid Plant",
  cs2Furnace:"CS2 Furnace",cs2Plant:"CS2 Plant",wsa:"WSA",
};

const ALL_KEYS = SECTIONS.flatMap(s => s.keys);

const SEC_MAP: Record<string,string> = {
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

const HEAD_REASON_ORDER = ["IDLE","PLANNED","UNPLANNED","OTHERS"];

const HR_COLORS: Record<string,{bg:string,sub:string,label:string,remBg:string,remSub:string,remLabel:string,pctBg:string,pctSub:string,pctLabel:string}> = {
  IDLE:      { bg:"#fce4d6", sub:"#f4b183", label:"Idle Hrs.",      remBg:"#fdf3ee", remSub:"#f9cdb0", remLabel:"Remaining Hrs. (after Idle)",      pctBg:"#fce9df", pctSub:"#f7c4a0", pctLabel:"% Idle Hrs. / Total Hrs."      },
  PLANNED:   { bg:"#e2efda", sub:"#a9d18e", label:"Planned Hrs.",   remBg:"#f0f7eb", remSub:"#c6e3b1", remLabel:"Remaining Hrs. (after Planned)",   pctBg:"#e8f3e1", pctSub:"#bcdba4", pctLabel:"% Planned Hrs. / Total Hrs."   },
  UNPLANNED: { bg:"#fff2cc", sub:"#ffd966", label:"Unplanned Hrs.", remBg:"#fffae5", remSub:"#ffe99a", remLabel:"Remaining Hrs. (after Unplanned)", pctBg:"#fff6d6", pctSub:"#ffdf80", pctLabel:"% Unplanned Hrs. / Total Hrs." },
  OTHERS:    { bg:"#ededed", sub:"#bfbfbf", label:"Others Hrs.",    remBg:"#f5f5f5", remSub:"#d6d6d6", remLabel:"Remaining Hrs. (after Others)",    pctBg:"#f0f0f0", pctSub:"#cccccc", pctLabel:"% Others Hrs. / Total Hrs."    },
};

// ─── Date Helpers ─────────────────────────────────────────────────────────────

function dim(y: number, m: number) { return new Date(y, m, 0).getDate(); }

interface YMD { y: number; m: number; d: number }

function parseDateOnly(v: any): YMD | null {
  if (v === null || v === undefined || v === "") return null;
  if (typeof v === "number") {
    const serial = Math.floor(v);
    const ms = (serial - 25569) * 86400000;
    const y = new Date(ms).getUTCFullYear();
    const m = new Date(ms).getUTCMonth() + 1;
    const d = new Date(ms).getUTCDate();
    return { y, m, d };
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

interface MM { y:number; m:number; label:string; days:number }

function getMonths(rows: any[]): MM[] {
  const seen = new Map<string,MM>();
  rows.forEach(r => {
    const ymd = parseDateOnly(r._d); if (!ymd) return;
    const label = `${String(ymd.m).padStart(2,"0")}/${String(ymd.y).slice(-2)}`;
    if (!seen.has(label)) seen.set(label, { y: ymd.y, m: ymd.m, label, days: dim(ymd.y, ymd.m) });
  });
  return [...seen.values()].sort((a,b) => a.y!==b.y ? a.y-b.y : a.m-b.m);
}

// ─── Excel Parser ─────────────────────────────────────────────────────────────

async function parseExcel(file: File) {
  const buf = await file.arrayBuffer();
  const wb = XLSX.read(buf, { type:"array", cellDates:false, raw:false });
  const ws = wb.Sheets[wb.SheetNames[0]];
  const json: any[] = XLSX.utils.sheet_to_json(ws, { defval:"", raw:true });

  const norm = (s: string) => s.toLowerCase().replace(/[^a-z0-9]/g," ").replace(/\s+/g," ").trim();

  return json.map((row) => {
    const r: any = {};
    Object.entries(row).forEach(([k,v]) => {
      const n = norm(k);
      if (n === "date")                                   r._d    = v;
      if (n === "plant")                                  r._plant= Number(v)||0;
      if (["functional location","functional loc","funcloc","func location"].includes(n))
                                                          r._floc = String(v).trim();
      if (n === "section")                                r._sec  = String(v).toUpperCase().replace(/\s+/g," ").trim();
      if (n === "head reason" || n === "headreason")      r._hr   = String(v).toUpperCase().replace(/\s+/g," ").trim();
      if (n === "sub head reason" || n === "subheadreason" || n === "sub reason") r._shr = String(v).trim();
      if (n.includes("total down") || n === "totaldowntime") r._down= Number(v)||0;
    });
    if (!r._floc) {
      const u = UNITS.find(u => u.plant === r._plant);
      r._floc = u ? u.name : String(r._plant);
    }
    return r;
  }).filter(r => r._plant && r._sec);
}

// ─── Multi-Select Dropdown ────────────────────────────────────────────────────

interface MultiSelectProps {
  label: string;
  options: string[];
  selected: string[];
  onChange: (vals: string[]) => void;
  disabled?: boolean;
}

function MultiSelect({ label, options, selected, onChange, disabled }: MultiSelectProps) {
  const [open, setOpen] = useState(false);
  const ref = useRef<HTMLDivElement>(null);

  React.useEffect(() => {
    if (!open) return;
    const handler = (e: MouseEvent) => {
      if (ref.current && !ref.current.contains(e.target as Node)) setOpen(false);
    };
    document.addEventListener("mousedown", handler);
    return () => document.removeEventListener("mousedown", handler);
  }, [open]);

  const allSelected = options.length > 0 && selected.length === options.length;

  const toggle = (v: string) => {
    onChange(selected.includes(v) ? selected.filter(x => x !== v) : [...selected, v]);
  };

  const toggleAll = () => onChange(allSelected ? [] : [...options]);

  const displayText =
    selected.length === 0 || allSelected
      ? `All ${label}`
      : selected.length === 1
      ? selected[0]
      : `${selected.length} selected`;

  return (
    <div ref={ref} style={{ position:"relative", display:"inline-block", minWidth:140, fontFamily:"Calibri,Arial,sans-serif", fontSize:11 }}>
      <div style={{ fontSize:10, color:"#555", marginBottom:2, fontWeight:600 }}>{label}</div>
      <button
        disabled={disabled}
        onClick={() => !disabled && setOpen(o => !o)}
        style={{
          width:"100%", padding:"3px 22px 3px 6px", border:"1px solid #aaa",
          background: disabled ? "#f5f5f5" : "#fff",
          cursor: disabled ? "default" : "pointer",
          textAlign:"left", fontSize:11, fontFamily:"Calibri,Arial,sans-serif",
          whiteSpace:"nowrap", overflow:"hidden", textOverflow:"ellipsis",
          color: disabled ? "#aaa" : "#222", borderRadius:2, position:"relative",
        }}
      >
        {displayText}
        <span style={{ position:"absolute", right:5, top:"50%", transform:"translateY(-50%)", pointerEvents:"none", fontSize:9 }}>▼</span>
      </button>
      {open && !disabled && (
        <div style={{
          position:"absolute", top:"100%", left:0, zIndex:9999,
          background:"#fff", border:"1px solid #aaa", boxShadow:"0 2px 8px rgba(0,0,0,0.15)",
          minWidth:"100%", maxHeight:220, overflowY:"auto", borderRadius:2,
        }}>
          <div
            onClick={toggleAll}
            style={{ padding:"3px 8px", cursor:"pointer", background: allSelected ? "#dce6f1" : "transparent",
              display:"flex", alignItems:"center", gap:5, fontSize:11, fontWeight:600,
              borderBottom:"1px solid #eee", userSelect:"none" }}
          >
            <input type="checkbox" readOnly checked={allSelected} style={{ margin:0 }} />
            <span>{allSelected ? "Deselect All" : "Select All"}</span>
          </div>
          {options.map(o => (
            <div
              key={o}
              onClick={() => toggle(o)}
              style={{ padding:"3px 8px", cursor:"pointer",
                background: selected.includes(o) ? "#dce6f1" : "transparent",
                display:"flex", alignItems:"center", gap:5, fontSize:11, userSelect:"none" }}
            >
              <input type="checkbox" readOnly checked={selected.includes(o)} style={{ margin:0 }} />
              <span>{o}</span>
            </div>
          ))}
        </div>
      )}
    </div>
  );
}

// ─── Types ────────────────────────────────────────────────────────────────────

// dm[hr][plant][pk][month] = totalDown
type DM = Record<string, Record<number, Record<string, Record<string, number>>>>;

// shrMap[hr][shr][plant][pk][month] = down  ← NOW PER PK
type SHRM = Record<string, Record<string, Record<number, Record<string, Record<string, number>>>>>;

interface ParsedRow {
  _d: any; _plant: number; _floc: string; _sec: string; _hr: string; _shr: string; _down: number;
}

// ─── Main Component ───────────────────────────────────────────────────────────

export default function IdleAnalysis() {
  const [allRows, setAllRows] = useState<ParsedRow[]>([]);
  const [fileName, setFN]     = useState("");
  const [loaded, setLoaded]   = useState(false);
  const [drag, setDrag]       = useState(false);
  const fileRef = useRef<HTMLInputElement>(null);

  const [selFloc,    setSelFloc]    = useState<string[]>([]);
  const [selPlant,   setSelPlant]   = useState<string[]>([]);
  const [selSection, setSelSection] = useState<string[]>([]);
  const [selHR,      setSelHR]      = useState<string[]>([]);
  const [selSHR,     setSelSHR]     = useState<string[]>([]);

  const load = useCallback(async (file: File) => {
    setFN(file.name);
    const rows = await parseExcel(file) as ParsedRow[];
    setAllRows(rows);
    setSelFloc([]); setSelPlant([]); setSelSection([]); setSelHR([]); setSelSHR([]);
    setLoaded(true);
  }, []);

  const reset = () => {
    setLoaded(false); setAllRows([]); setFN("");
    setSelFloc([]); setSelPlant([]); setSelSection([]); setSelHR([]); setSelSHR([]);
  };

  const onDrop = (e: React.DragEvent) => {
    e.preventDefault(); setDrag(false);
    const f = e.dataTransfer.files[0]; if (f) load(f);
  };

  // ── Cascading filter options ──────────────────────────────────────────────

  const allFlocs = useMemo(() =>
    [...new Set(allRows.map(r => r._floc).filter(Boolean))].sort(),
  [allRows]);

  const rowsAfterFloc = useMemo(() =>
    selFloc.length === 0 ? allRows : allRows.filter(r => selFloc.includes(r._floc)),
  [allRows, selFloc]);

  const allPlants = useMemo(() =>
    [...new Set(rowsAfterFloc.map(r => String(r._plant)).filter(Boolean))].sort((a,b) => +a - +b),
  [rowsAfterFloc]);

  const rowsAfterPlant = useMemo(() =>
    selPlant.length === 0 ? rowsAfterFloc : rowsAfterFloc.filter(r => selPlant.includes(String(r._plant))),
  [rowsAfterFloc, selPlant]);

  const allSections = useMemo(() =>
    [...new Set(rowsAfterPlant.map(r => r._sec).filter(Boolean))].sort(),
  [rowsAfterPlant]);

  const rowsAfterSection = useMemo(() =>
    selSection.length === 0 ? rowsAfterPlant : rowsAfterPlant.filter(r => selSection.includes(r._sec)),
  [rowsAfterPlant, selSection]);

  const allHRs = useMemo(() => {
    const seen = new Set(rowsAfterSection.map(r => r._hr).filter(Boolean));
    return [
      ...HEAD_REASON_ORDER.filter(h => seen.has(h)),
      ...[...seen].filter(h => !HEAD_REASON_ORDER.includes(h)).sort(),
    ];
  }, [rowsAfterSection]);

  const rowsAfterHR = useMemo(() =>
    selHR.length === 0 ? rowsAfterSection : rowsAfterSection.filter(r => selHR.includes(r._hr)),
  [rowsAfterSection, selHR]);

  const allSHRs = useMemo(() =>
    [...new Set(rowsAfterHR.map(r => r._shr).filter(Boolean))].sort(),
  [rowsAfterHR]);

  // ── Final filtered rows ───────────────────────────────────────────────────
  const filteredRows = useMemo(() => {
    let rows = allRows;
    if (selFloc.length > 0)    rows = rows.filter(r => selFloc.includes(r._floc));
    if (selPlant.length > 0)   rows = rows.filter(r => selPlant.includes(String(r._plant)));
    if (selSection.length > 0) rows = rows.filter(r => selSection.includes(r._sec));
    if (selHR.length > 0)      rows = rows.filter(r => selHR.includes(r._hr));
    if (selSHR.length > 0)     rows = rows.filter(r => selSHR.includes(r._shr));
    return rows;
  }, [allRows, selFloc, selPlant, selSection, selHR, selSHR]);

  const months = useMemo(() => getMonths(filteredRows), [filteredRows]);

  const { dm, shrMap, activePks, activeHRs } = useMemo(() => {
    const map: DM   = {};
    // shrMap[hr][shr][plant][pk][month] = down
    const smap: SHRM = {};

    filteredRows.forEach((r) => {
      const ymd = parseDateOnly(r._d); if (!ymd) return;
      const lbl = `${String(ymd.m).padStart(2,"0")}/${String(ymd.y).slice(-2)}`;
      const pk  = SEC_MAP[r._sec]; if (!pk) return;
      const hr  = r._hr || "OTHERS";
      const p   = r._plant;

      // ── dm: unchanged ──
      if (!map[hr])        map[hr]        = {};
      if (!map[hr][p])     map[hr][p]     = {};
      if (!map[hr][p][pk]) map[hr][p][pk] = {};
      map[hr][p][pk][lbl] = (map[hr][p][pk][lbl] || 0) + r._down;

      // ── shrMap: now keyed per pk ──
      const shr = r._shr || "Unknown";
      if (!smap[hr])                   smap[hr]                   = {};
      if (!smap[hr][shr])              smap[hr][shr]              = {};
      if (!smap[hr][shr][p])           smap[hr][shr][p]           = {};
      if (!smap[hr][shr][p][pk])       smap[hr][shr][p][pk]       = {};
      smap[hr][shr][p][pk][lbl] = (smap[hr][shr][p][pk][lbl] || 0) + r._down;
    });

    const seenPks = new Set(filteredRows.map((r: any) => SEC_MAP[r._sec]).filter(Boolean));
    const orderedPks = ALL_KEYS.filter(k => seenPks.has(k));

    const seenHRs = new Set(filteredRows.map((r: any) => r._hr).filter(Boolean));
    const orderedHRs = [
      ...HEAD_REASON_ORDER.filter(h => seenHRs.has(h)),
      ...[...seenHRs].filter(h => !HEAD_REASON_ORDER.includes(h)).sort(),
    ];

    return { dm: map, shrMap: smap, activePks: orderedPks, activeHRs: orderedHRs };
  }, [filteredRows]);

  const activeUnits = useMemo(() => {
    if (selFloc.length === 0 && selPlant.length === 0) return UNITS;
    return UNITS.filter(u => {
      if (selFloc.length > 0 && !selFloc.includes(u.name)) return false;
      if (selPlant.length > 0 && !selPlant.includes(String(u.plant))) return false;
      return true;
    });
  }, [selFloc, selPlant]);

  const handleFlocChange    = (v: string[]) => { setSelFloc(v);    setSelPlant([]); setSelSection([]); setSelHR([]); setSelSHR([]); };
  const handlePlantChange   = (v: string[]) => { setSelPlant(v);   setSelSection([]); setSelHR([]); setSelSHR([]); };
  const handleSectionChange = (v: string[]) => { setSelSection(v); setSelHR([]); setSelSHR([]); };
  const handleHRChange      = (v: string[]) => { setSelHR(v);      setSelSHR([]); };
  const clearAll = () => { setSelFloc([]); setSelPlant([]); setSelSection([]); setSelHR([]); setSelSHR([]); };

  // ─── Table style helpers ──────────────────────────────────────────────────
  const border = "1px solid #b0b0b0";
  const base: React.CSSProperties = { border, padding:"2px 5px", fontSize:11, whiteSpace:"nowrap", fontFamily:"Calibri,Arial,sans-serif" };
  const num:  React.CSSProperties = { ...base, textAlign:"right" };
  const hdr:  React.CSSProperties = { ...base, textAlign:"center", fontWeight:700 };

  function getActiveSections(pks: string[]) {
    return SECTIONS.map(s => ({ ...s, activeKeys: s.keys.filter(k => pks.includes(k)) }))
                   .filter(s => s.activeKeys.length > 0);
  }

  // ── Machine Count Table ───────────────────────────────────────────────────
  function renderMachineCountTable() {
    const displayPks = activePks.length > 0 ? activePks : ALL_KEYS;
    const activeSections = getActiveSections(displayPks);
    return (
      <table style={{ borderCollapse:"collapse", marginBottom:16 }}>
        <thead>
          <tr>
            <th style={{...hdr, background:"#dce6f1", minWidth:65}} rowSpan={2}>Unit Name</th>
            <th style={{...hdr, background:"#dce6f1", minWidth:48}} rowSpan={2}>Plant</th>
            {activeSections.map(s => (
              <th key={s.label} style={{...hdr, background:"#dce6f1"}} colSpan={s.activeKeys.length}>{s.label}</th>
            ))}
          </tr>
          <tr>
            {activeSections.flatMap(s => s.activeKeys.map(k => (
              <th key={k} style={{...hdr, background:"#bdd7ee", fontSize:10, minWidth:52}}>{PL[k]}</th>
            )))}
          </tr>
        </thead>
        <tbody>
          {activeUnits.map((u,i) => (
            <tr key={u.name} style={{background:i%2===1?"#f2f2f2":"#fff"}}>
              <td style={{...base, textAlign:"left", fontWeight:600}}>{u.name}</td>
              <td style={{...num, textAlign:"center", color:"#555"}}>{u.plant}</td>
              {activeSections.flatMap(s => s.activeKeys.map(k => (
                <td key={k} style={num}>{MC[u.name]?.[k]??0}</td>
              )))}
            </tr>
          ))}
        </tbody>
      </table>
    );
  }

  // ── Total Hrs Table ───────────────────────────────────────────────────────
  function renderTotalHrsTable(pks: string[]) {
    const activeSections = getActiveSections(pks);
    const totalCols = pks.length;
    return (
      <table style={{ borderCollapse:"collapse", marginBottom:16 }}>
        <thead>
          <tr>
            <th style={{...hdr, background:"#dce6f1", minWidth:65}} rowSpan={4}>Unit</th>
            <th style={{...hdr, background:"#dce6f1", minWidth:48}} rowSpan={4}>Plant</th>
            {months.map(mk => (
              <th key={mk.label} style={{...hdr, background:"#dce6f1"}} colSpan={totalCols}>{mk.label}</th>
            ))}
          </tr>
          <tr>
            {months.map(mk => (
              <th key={mk.label} style={{...hdr, background:"#bdd7ee", fontSize:10}} colSpan={totalCols}>{mk.days} Days</th>
            ))}
          </tr>
          <tr>
            {months.map(mk =>
              activeSections.map(s => (
                <th key={`${mk.label}-${s.label}`} style={{...hdr, background:"#dce6f1"}} colSpan={s.activeKeys.length}>{s.label}</th>
              ))
            )}
          </tr>
          <tr>
            {months.map(mk =>
              activeSections.flatMap(s =>
                s.activeKeys.map(k => (
                  <th key={`${mk.label}-${k}`} style={{...hdr, background:"#bdd7ee", fontSize:10, minWidth:52}}>{PL[k]}</th>
                ))
              )
            )}
          </tr>
        </thead>
        <tbody>
          {activeUnits.map((u, i) => (
            <tr key={u.name} style={{background:i%2===1?"#f2f2f2":"#fff"}}>
              <td style={{...base, textAlign:"left", fontWeight:600}}>{u.name}</td>
              <td style={{...num, textAlign:"center", color:"#555"}}>{u.plant}</td>
              {months.map(mk =>
                activeSections.flatMap(s =>
                  s.activeKeys.map(k => {
                    const mc = MC[u.name]?.[k] ?? 0;
                    const total = mc * mk.days * 24;
                    return (
                      <td key={`${mk.label}-${k}`} style={{...num, color:total===0?"#bbb":"#000"}}>
                        {total===0?"-":total}
                      </td>
                    );
                  })
                )
              )}
            </tr>
          ))}
        </tbody>
      </table>
    );
  }

  // ── Consolidated HR Table ─────────────────────────────────────────────────
  function renderConsolidatedHRTable(hr: string, pks: string[], bgColor: string, subColor: string) {
    const activeSections = getActiveSections(pks);
    const totalCols = pks.length;
    const dataMap = dm[hr] ?? {};
    return (
      <table style={{ borderCollapse:"collapse", marginBottom:16 }}>
        <thead>
          <tr>
            <th style={{...hdr, background:bgColor, minWidth:65}} rowSpan={4}>Unit</th>
            <th style={{...hdr, background:bgColor, minWidth:48}} rowSpan={4}>Plant</th>
            {months.map(mk => (
              <th key={mk.label} style={{...hdr, background:bgColor}} colSpan={totalCols}>{mk.label}</th>
            ))}
          </tr>
          <tr>
            {months.map(mk => (
              <th key={mk.label} style={{...hdr, background:subColor, fontSize:10}} colSpan={totalCols}>{mk.days} Days</th>
            ))}
          </tr>
          <tr>
            {months.map(mk =>
              activeSections.map(s => (
                <th key={`${mk.label}-${s.label}`} style={{...hdr, background:bgColor}} colSpan={s.activeKeys.length}>{s.label}</th>
              ))
            )}
          </tr>
          <tr>
            {months.map(mk =>
              activeSections.flatMap(s =>
                s.activeKeys.map(k => (
                  <th key={`${mk.label}-${k}`} style={{...hdr, background:subColor, fontSize:10, minWidth:52}}>{PL[k]}</th>
                ))
              )
            )}
          </tr>
        </thead>
        <tbody>
          {activeUnits.map((u, i) => (
            <tr key={u.name} style={{background:i%2===1?"#f2f2f2":"#fff"}}>
              <td style={{...base, textAlign:"left", fontWeight:600}}>{u.name}</td>
              <td style={{...num, textAlign:"center", color:"#555"}}>{u.plant}</td>
              {months.map(mk =>
                activeSections.flatMap(s =>
                  s.activeKeys.map(k => {
                    const v = dataMap[u.plant]?.[k]?.[mk.label] ?? 0;
                    return (
                      <td key={`${mk.label}-${k}`} style={{...num, color:v===0?"#bbb":"#000"}}>
                        {v===0?"-":v.toFixed(2)}
                      </td>
                    );
                  })
                )
              )}
            </tr>
          ))}
        </tbody>
      </table>
    );
  }

  // ── Remaining Hrs Table ───────────────────────────────────────────────────
  function renderRemainingTables(hr: string, pks: string[], bgColor: string, subColor: string) {
    const activeSections = getActiveSections(pks);
    const totalCols = pks.length;
    return (
      <table style={{ borderCollapse:"collapse", marginBottom:16 }}>
        <thead>
          <tr>
            <th style={{...hdr, background:bgColor, minWidth:65}} rowSpan={4}>Unit</th>
            <th style={{...hdr, background:bgColor, minWidth:48}} rowSpan={4}>Plant</th>
            {months.map(mk => (
              <th key={mk.label} style={{...hdr, background:bgColor}} colSpan={totalCols}>{mk.label}</th>
            ))}
          </tr>
          <tr>
            {months.map(mk => (
              <th key={mk.label} style={{...hdr, background:subColor, fontSize:10}} colSpan={totalCols}>{mk.days} Days</th>
            ))}
          </tr>
          <tr>
            {months.map(mk =>
              activeSections.map(s => (
                <th key={`${mk.label}-${s.label}`} style={{...hdr, background:bgColor}} colSpan={s.activeKeys.length}>{s.label}</th>
              ))
            )}
          </tr>
          <tr>
            {months.map(mk =>
              activeSections.flatMap(s =>
                s.activeKeys.map(k => (
                  <th key={`${mk.label}-${k}`} style={{...hdr, background:subColor, fontSize:10, minWidth:52}}>{PL[k]}</th>
                ))
              )
            )}
          </tr>
        </thead>
        <tbody>
          {activeUnits.map((u, i) => (
            <tr key={u.name} style={{background:i%2===1?"#f2f2f2":"#fff"}}>
              <td style={{...base, textAlign:"left", fontWeight:600}}>{u.name}</td>
              <td style={{...num, textAlign:"center", color:"#555"}}>{u.plant}</td>
              {months.map(mk =>
                activeSections.flatMap(s =>
                  s.activeKeys.map(k => {
                    const mc        = MC[u.name]?.[k] ?? 0;
                    const total     = mc * mk.days * 24;
                    const downForHR = dm[hr]?.[u.plant]?.[k]?.[mk.label] ?? 0;
                    const remaining = total - downForHR;
                    const isNeg     = remaining < 0;
                    return (
                      <td key={`${mk.label}-${k}`} style={{...num,
                        color: total===0?"#bbb":isNeg?"#c00":"#000",
                        fontWeight: isNeg?700:"normal"
                      }}>
                        {total===0?"-":remaining.toFixed(2)}
                      </td>
                    );
                  })
                )
              )}
            </tr>
          ))}
        </tbody>
      </table>
    );
  }

  // ── Percentage Table ──────────────────────────────────────────────────────
  function renderPercentageTables(hr: string, pks: string[], bgColor: string, subColor: string) {
    const activeSections = getActiveSections(pks);
    const totalCols = pks.length;
    return (
      <table style={{ borderCollapse:"collapse", marginBottom:16 }}>
        <thead>
          <tr>
            <th style={{...hdr, background:bgColor, minWidth:65}} rowSpan={4}>Unit</th>
            <th style={{...hdr, background:bgColor, minWidth:48}} rowSpan={4}>Plant</th>
            {months.map(mk => (
              <th key={mk.label} style={{...hdr, background:bgColor}} colSpan={totalCols}>{mk.label}</th>
            ))}
          </tr>
          <tr>
            {months.map(mk => (
              <th key={mk.label} style={{...hdr, background:subColor, fontSize:10}} colSpan={totalCols}>{mk.days} Days</th>
            ))}
          </tr>
          <tr>
            {months.map(mk =>
              activeSections.map(s => (
                <th key={`${mk.label}-${s.label}`} style={{...hdr, background:bgColor}} colSpan={s.activeKeys.length}>{s.label}</th>
              ))
            )}
          </tr>
          <tr>
            {months.map(mk =>
              activeSections.flatMap(s =>
                s.activeKeys.map(k => (
                  <th key={`${mk.label}-${k}`} style={{...hdr, background:subColor, fontSize:10, minWidth:52}}>{PL[k]}</th>
                ))
              )
            )}
          </tr>
        </thead>
        <tbody>
          {activeUnits.map((u, i) => (
            <tr key={u.name} style={{background:i%2===1?"#f2f2f2":"#fff"}}>
              <td style={{...base, textAlign:"left", fontWeight:600}}>{u.name}</td>
              <td style={{...num, textAlign:"center", color:"#555"}}>{u.plant}</td>
              {months.map(mk =>
                activeSections.flatMap(s =>
                  s.activeKeys.map(k => {
                    const mc        = MC[u.name]?.[k] ?? 0;
                    const total     = mc * mk.days * 24;
                    const downForHR = dm[hr]?.[u.plant]?.[k]?.[mk.label] ?? 0;
                    const remaining = total - downForHR;
                    const pct       = remaining === 0 ? null : (downForHR / remaining) * 100;
                    const isHigh    = pct !== null && pct > 100;
                    const isEmpty   = pct === null || pct === 0;
                    return (
                      <td key={`${mk.label}-${k}`} style={{...num,
                        color: isEmpty?"#bbb":isHigh?"#c00":"#000",
                        fontWeight: isHigh?700:"normal"
                      }}>
                        {isEmpty ? "-" : `${pct!.toFixed(3)}%`}
                      </td>
                    );
                  })
                )
              )}
            </tr>
          ))}
        </tbody>
      </table>
    );
  }

  // ─────────────────────────────────────────────────────────────────────────
  // ── SHR Summary Table ─────────────────────────────────────────────────────
  // ─────────────────────────────────────────────────────────────────────────

  function getFY(mk: MM): string {
    return mk.m >= 4 ? `FY${(mk.y+1).toString().slice(-2)}` : `FY${mk.y.toString().slice(-2)}`;
  }

  // ── KEY FIX: pctCell now matches renderPercentageTables exactly ──────────
  //
  // Formula per pk per month (same as renderPercentageTables):
  //   pct_pk = shrDown_pk / (total_pk - hrDown_pk) * 100
  //
  // For summary we aggregate:
  //   numerator   = sum over pks where shrDown_pk > 0 of shrDown_pk
  //   denominator = sum over SAME pks of (total_pk - hrDown_pk)
  //
  // By restricting denominator to only pks that have shrDown,
  // a single-pk case gives exactly the same result as the % table.
  //
  function pctCell(
    hr: string,
    shr: string | null,   // null = HR total row
    plant: number,
    unitName: string,
    mks: MM[]
  ): string {
    let numerator   = 0;
    let denominator = 0;

    activePks.forEach(pk => {
      mks.forEach(mk => {
        const mc    = MC[unitName]?.[pk] ?? 0;
        const total = mc * mk.days * 24;
        if (total === 0) return; // machine doesn't exist at this unit

        // shrDown for this pk/month
        const shrDown = shr === null
          // HR total row: sum across all SHRs for this pk/month
          ? Object.values(shrMap[hr] ?? {}).reduce(
              (s, pkMap) => s + (pkMap[plant]?.[pk]?.[mk.label] ?? 0), 0
            )
          // specific SHR row
          : (shrMap[hr]?.[shr]?.[plant]?.[pk]?.[mk.label] ?? 0);

        // ── CRITICAL: only include this pk in denominator if it has shrDown ──
        // This ensures the summary % matches the per-pk % table exactly
        if (shrDown === 0) return;

        const hrDown = dm[hr]?.[plant]?.[pk]?.[mk.label] ?? 0;

        numerator   += shrDown;
        denominator += (total - hrDown);
      });
    });

    if (numerator === 0)   return "-";
    if (denominator === 0) return "-";

    const pct = (numerator / denominator) * 100;
    return `${pct.toFixed(3)}%`;
  }

  function renderSHRTable() {
    if (!months.length) return null;

    const fySet: string[] = [];
    const fyMonths: Record<string, MM[]> = {};
    months.forEach(mk => {
      const fy = getFY(mk);
      if (!fyMonths[fy]) { fyMonths[fy] = []; fySet.push(fy); }
      fyMonths[fy].push(mk);
    });

    const fullFYs   = fySet.slice(0, -1);
    const currentFY = fySet[fySet.length - 1];
    const curMonths = fyMonths[currentFY];
    const colsPerUnit = fullFYs.length + curMonths.length + 1;

    const HR_ROW_COLORS: Record<string,{bg:string,text:string}> = {
      IDLE:      { bg:"#f4b183", text:"#7b3200" },
      PLANNED:   { bg:"#a9d18e", text:"#1e4d1e" },
      UNPLANNED: { bg:"#ffd966", text:"#7b5800" },
      OTHERS:    { bg:"#bfbfbf", text:"#333"    },
    };

    const cellStyle:    React.CSSProperties = { ...base, textAlign:"right", minWidth:48 };
    const labelStyle:   React.CSSProperties = { ...base, textAlign:"left", minWidth:160, paddingLeft:6 };
    const hrLabelStyle: React.CSSProperties = { ...base, textAlign:"left", fontWeight:700, minWidth:160, paddingLeft:4 };
    const normsStyle:   React.CSSProperties = { ...base, textAlign:"center", minWidth:40 };

    return (
      <table style={{ borderCollapse:"collapse", marginBottom:16 }}>
        <thead>
          <tr>
            <th style={{...hdr, background:"#dce6f1", minWidth:160}} rowSpan={3}>Head / Sub Head Reason</th>
            <th style={{...hdr, background:"#dce6f1", minWidth:40}}  rowSpan={3}>Norms</th>
            {activeUnits.map(u => (
              <th key={u.name} style={{...hdr, background:"#dce6f1"}} colSpan={colsPerUnit}>{u.name}</th>
            ))}
          </tr>
          <tr>
            {activeUnits.map(u => (
              <React.Fragment key={u.name}>
                {fullFYs.map(fy => (
                  <th key={fy} style={{...hdr, background:"#bdd7ee", minWidth:48}}>{fy}</th>
                ))}
                <th style={{...hdr, background:"#dce6f1"}} colSpan={curMonths.length + 1}>{currentFY}</th>
              </React.Fragment>
            ))}
          </tr>
          <tr>
            {activeUnits.map(u => (
              <React.Fragment key={u.name}>
                {fullFYs.map(fy => (
                  <th key={fy} style={{...hdr, background:"#bdd7ee", fontSize:10}}>{fy}</th>
                ))}
                {curMonths.map(mk => (
                  <th key={mk.label} style={{...hdr, background:"#bdd7ee", fontSize:10, minWidth:48}}>{mk.label}</th>
                ))}
                <th style={{...hdr, background:"#dae3f3", fontSize:10, minWidth:52}}>UTD Avg</th>
              </React.Fragment>
            ))}
          </tr>
        </thead>
        <tbody>
          {activeHRs.map(hr => {
            const shrs = Object.keys(shrMap[hr] ?? {}).sort();
            const hrColor = HR_ROW_COLORS[hr] ?? { bg:"#bfbfbf", text:"#333" };
            const hrLabel = hr.charAt(0) + hr.slice(1).toLowerCase();

            return (
              <React.Fragment key={hr}>
                {/* ── SHR rows ── */}
                {shrs.map((shr, si) => (
                  <tr key={`${hr}-${shr}`} style={{background: si%2===0?"#fff":"#f5f5f5"}}>
                    <td style={{...labelStyle, paddingLeft:14, color:"#333"}}>{shr}</td>
                    <td style={normsStyle}>-</td>
                    {activeUnits.map(u => (
                      <React.Fragment key={u.name}>
                        {fullFYs.map(fy => (
                          <td key={fy} style={cellStyle}>
                            {pctCell(hr, shr, u.plant, u.name, fyMonths[fy])}
                          </td>
                        ))}
                        {curMonths.map(mk => (
                          <td key={mk.label} style={cellStyle}>
                            {pctCell(hr, shr, u.plant, u.name, [mk])}
                          </td>
                        ))}
                        <td style={{...cellStyle, background:"#eef3fb"}}>
                          {pctCell(hr, shr, u.plant, u.name, curMonths)}
                        </td>
                      </React.Fragment>
                    ))}
                  </tr>
                ))}

                {/* ── HR total row ── */}
                <tr style={{background: hrColor.bg}}>
                  <td style={{...hrLabelStyle, color: hrColor.text}}>{hrLabel}</td>
                  <td style={{...normsStyle, color: hrColor.text}}>-</td>
                  {activeUnits.map(u => (
                    <React.Fragment key={u.name}>
                      {fullFYs.map(fy => (
                        <td key={fy} style={{...cellStyle, fontWeight:700, color: hrColor.text}}>
                          {pctCell(hr, null, u.plant, u.name, fyMonths[fy])}
                        </td>
                      ))}
                      {curMonths.map(mk => (
                        <td key={mk.label} style={{...cellStyle, fontWeight:700, color: hrColor.text}}>
                          {pctCell(hr, null, u.plant, u.name, [mk])}
                        </td>
                      ))}
                      <td style={{...cellStyle, fontWeight:700, background: hrColor.bg, color: hrColor.text}}>
                        {pctCell(hr, null, u.plant, u.name, curMonths)}
                      </td>
                    </React.Fragment>
                  ))}
                </tr>
              </React.Fragment>
            );
          })}
        </tbody>
      </table>
    );
  }

  function SectionTitle({ label, borderColor }: { label: string; borderColor: string }) {
    return (
      <div style={{ fontWeight:700, fontSize:12, margin:"14px 0 4px", color:"#1f3864",
        borderBottom:`2px solid ${borderColor}`, paddingBottom:2, display:"inline-block" }}>
        {label}
      </div>
    );
  }

  const hasFilters = selFloc.length > 0 || selPlant.length > 0 || selSection.length > 0 || selHR.length > 0 || selSHR.length > 0;

  return (
    <div style={{ padding:12, background:"#fff", fontFamily:"Calibri,Arial,sans-serif", fontSize:11 }}>

      {/* Header */}
      <div style={{ display:"flex", alignItems:"center", gap:10, marginBottom:10 }}>
        <b style={{ fontSize:13, color:"#1f3864" }}>Downtime Analysis</b>
        {loaded && <span style={{ color:"#555", fontSize:11 }}>— {fileName}</span>}
        {loaded && (
          <button onClick={reset} style={{ marginLeft:"auto", fontSize:11, padding:"2px 10px", cursor:"pointer" }}>✕ Reset</button>
        )}
      </div>

      {/* Upload */}
      {!loaded && (
        <div
          onDragOver={e=>{e.preventDefault();setDrag(true)}}
          onDragLeave={()=>setDrag(false)}
          onDrop={onDrop}
          onClick={()=>fileRef.current?.click()}
          style={{ border:`2px dashed ${drag?"#4472c4":"#bbb"}`, background:drag?"#eef3fb":"#fafafa",
            padding:"48px 0", textAlign:"center", cursor:"pointer", borderRadius:3, marginBottom:12 }}
        >
          <input ref={fileRef} type="file" accept=".xlsx,.xls,.csv" style={{display:"none"}}
            onChange={e=>{ const f=e.target.files?.[0]; if(f) load(f); }} />
          <div style={{fontSize:28,marginBottom:6}}>📊</div>
          <b style={{fontSize:12}}>Drop Excel file here or click to browse</b>
          <div style={{color:"#888",fontSize:11,marginTop:4}}>.xlsx · .xls · .csv</div>
        </div>
      )}

      {loaded && (
        <>
          {/* ── Filter Bar ── */}
          <div style={{
            background:"#f4f7fc", border:"1px solid #c8d8ee", borderRadius:3,
            padding:"8px 10px", marginBottom:12,
            display:"flex", flexWrap:"wrap", gap:12, alignItems:"flex-end",
          }}>
            <MultiSelect label="Functional Location" options={allFlocs}    selected={selFloc}    onChange={handleFlocChange} />
            <MultiSelect label="Plant"                options={allPlants}   selected={selPlant}   onChange={handlePlantChange}   disabled={allPlants.length === 0} />
            <MultiSelect label="Section"              options={allSections} selected={selSection} onChange={handleSectionChange} disabled={allSections.length === 0} />
            <MultiSelect label="Head Reason"          options={allHRs}      selected={selHR}      onChange={handleHRChange}      disabled={allHRs.length === 0} />
            <MultiSelect label="Sub Head Reason"      options={allSHRs}     selected={selSHR}     onChange={v => setSelSHR(v)}   disabled={allSHRs.length === 0} />
            {hasFilters && (
              <button onClick={clearAll} style={{
                fontSize:11, padding:"3px 10px", cursor:"pointer",
                background:"#fff", border:"1px solid #aaa", borderRadius:2,
                color:"#c00", fontWeight:600, alignSelf:"flex-end",
              }}>✕ Clear All</button>
            )}
          </div>

          <div style={{ overflowX:"auto" }}>

            {/* ══ MACHINE COUNT ══ */}
            <SectionTitle label="Machine Count" borderColor="#4472c4" />
            {renderMachineCountTable()}

            {/* ══ SUB HEAD REASON SUMMARY ══ */}
            <SectionTitle label="Sub Head Reason Summary" borderColor="#4472c4" />
            {renderSHRTable()}

            {/* ══ TOTAL HRS ══ */}
            <SectionTitle label="Total Hrs." borderColor="#4472c4" />
            {renderTotalHrsTable(activePks)}

            {/* ══ PER HEAD REASON BLOCKS ══ */}
            {activeHRs.map(hr => {
              const hrName = hr.charAt(0) + hr.slice(1).toLowerCase();
              const c = HR_COLORS[hr] ?? {
                bg:"#ededed", sub:"#bfbfbf", label:`${hrName} Hrs.`,
                remBg:"#f5f5f5", remSub:"#d6d6d6", remLabel:`Remaining Hrs. (after ${hrName})`,
                pctBg:"#f0f0f0", pctSub:"#cccccc", pctLabel:`% ${hrName} Hrs. / Total Hrs.`,
              };
              return (
                <div key={hr}>
                  <SectionTitle label={c.label} borderColor={c.sub} />
                  {renderConsolidatedHRTable(hr, activePks, c.bg, c.sub)}

                  <SectionTitle label={c.remLabel} borderColor={c.remSub} />
                  {renderRemainingTables(hr, activePks, c.remBg, c.remSub)}

                  <SectionTitle label={c.pctLabel} borderColor={c.pctSub} />
                  {renderPercentageTables(hr, activePks, c.pctBg, c.pctSub)}
                </div>
              );
            })}

          </div>
        </>
      )}
    </div>
  );
}