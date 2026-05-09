"use client";
import { useState, useCallback, useRef } from "react";
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

  if (v instanceof Date) {
    return { y: v.getFullYear(), m: v.getMonth() + 1, d: v.getDate() };
  }

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
      if (n === "date")                                  r._d    = v;
      if (n==="plant")                                   r._plant= Number(v)||0;
      if (n==="section")                                 r._sec  = String(v).toUpperCase().replace(/\s+/g," ").trim();
      if (n==="head reason"||n==="headreason")           r._hr   = String(v).toUpperCase().replace(/\s+/g," ").trim();
      if (n.includes("total down")||n==="totaldowntime") r._down = Number(v)||0;
    });
    return r;
  }).filter(r => r._plant && r._sec);
}

// ─── Component ────────────────────────────────────────────────────────────────

type DM = Record<string, Record<number, Record<string, Record<string, number>>>>;

export default function IdleAnalysis() {
  const [months,    setMonths]    = useState<MM[]>([]);
  const [dm,        setDm]        = useState<DM>({});
  const [activePks, setActivePks] = useState<string[]>([]);
  const [activeHRs, setActiveHRs] = useState<string[]>([]);
  const [fileName,  setFN]        = useState("");
  const [loaded,    setLoaded]    = useState(false);
  const [drag,      setDrag]      = useState(false);
  const fileRef = useRef<HTMLInputElement>(null);

  const load = useCallback(async (file: File) => {
    setFN(file.name);
    const rows = await parseExcel(file);
    const ms   = getMonths(rows);

    const map: DM = {};
    rows.forEach((r) => {
      const ymd = parseDateOnly(r._d); if (!ymd) return;
      const lbl = `${String(ymd.m).padStart(2,"0")}/${String(ymd.y).slice(-2)}`;
      const pk = SEC_MAP[r._sec]; if (!pk) return;
      const hr = r._hr || "OTHERS";
      const p  = r._plant;
      if (!map[hr])         map[hr]         = {};
      if (!map[hr][p])      map[hr][p]      = {};
      if (!map[hr][p][pk])  map[hr][p][pk]  = {};
      map[hr][p][pk][lbl] = (map[hr][p][pk][lbl]||0) + r._down;
    });

    const seenPks = new Set(rows.map((r:any) => SEC_MAP[r._sec]).filter(Boolean));
    const orderedPks = ALL_KEYS.filter(k => seenPks.has(k));

    const seenHRs = new Set(rows.map((r:any) => r._hr).filter(Boolean));
    const orderedHRs = [
      ...HEAD_REASON_ORDER.filter(h => seenHRs.has(h)),
      ...[...seenHRs].filter(h => !HEAD_REASON_ORDER.includes(h)).sort(),
    ];

    setMonths(ms);
    setDm(map);
    setActivePks(orderedPks);
    setActiveHRs(orderedHRs);
    setLoaded(true);
  }, []);

  const reset = () => { setLoaded(false); setMonths([]); setDm({}); setActivePks([]); setActiveHRs([]); setFN(""); };
  const onDrop = (e: React.DragEvent) => { e.preventDefault(); setDrag(false); const f=e.dataTransfer.files[0]; if(f) load(f); };

  const border = "1px solid #b0b0b0";
  const base: React.CSSProperties = { border, padding:"2px 5px", fontSize:11, whiteSpace:"nowrap", fontFamily:"Calibri,Arial,sans-serif" };
  const num:  React.CSSProperties = { ...base, textAlign:"right" };
  const hdr:  React.CSSProperties = { ...base, textAlign:"center", fontWeight:700 };

  // ── Sections filtered to only active keys ────────────────────────────────
  function getActiveSections(pks: string[]) {
    return SECTIONS.map(s => ({
      ...s,
      activeKeys: s.keys.filter(k => pks.includes(k)),
    })).filter(s => s.activeKeys.length > 0);
  }

  // ── TOTAL HRS consolidated table ─────────────────────────────────────────
  function renderTotalHrsTable(pks: string[]) {
    const activeSections = getActiveSections(pks);
    const totalCols = pks.length; // number of machine columns

    return (
      <table style={{ borderCollapse:"collapse", marginBottom:16 }}>
        <thead>
          {/* Row 1: Month labels, each spanning all machine cols */}
          <tr>
            <th style={{...hdr, background:"#dce6f1", minWidth:65}} rowSpan={4}>Unit</th>
            <th style={{...hdr, background:"#dce6f1", minWidth:48}} rowSpan={4}>Plant</th>
            {months.map(mk => (
              <th key={mk.label} style={{...hdr, background:"#dce6f1"}} colSpan={totalCols}>
                {mk.label}
              </th>
            ))}
          </tr>

          {/* Row 2: Days count per month */}
          <tr>
            {months.map(mk => (
              <th key={mk.label} style={{...hdr, background:"#bdd7ee", fontSize:10}} colSpan={totalCols}>
                {mk.days} Days
              </th>
            ))}
          </tr>

          {/* Row 3: Section group headers repeated per month */}
          <tr>
            {months.map(mk =>
              activeSections.map(s => (
                <th key={`${mk.label}-${s.label}`} style={{...hdr, background:"#dce6f1"}} colSpan={s.activeKeys.length}>
                  {s.label}
                </th>
              ))
            )}
          </tr>

          {/* Row 4: Individual machine headers repeated per month */}
          <tr>
            {months.map(mk =>
              activeSections.flatMap(s =>
                s.activeKeys.map(k => (
                  <th key={`${mk.label}-${k}`} style={{...hdr, background:"#bdd7ee", fontSize:10, minWidth:52}}>
                    {PL[k]}
                  </th>
                ))
              )
            )}
          </tr>
        </thead>
        <tbody>
          {UNITS.map((u, i) => (
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

  // ── Per-machine-row tables (for HR sections) ──────────────────────────────
  function renderTableShell(pk: string, bgColor: string, subColor: string, rows: React.ReactNode) {
    return (
      <table key={pk} style={{ borderCollapse:"collapse", marginBottom:8 }}>
        <thead>
          <tr>
            <td style={{...hdr, background:bgColor, textAlign:"left", minWidth:65}} colSpan={2}>{PL[pk]}</td>
            {months.map(mk => (
              <th key={mk.label} style={{...hdr, background:bgColor, minWidth:52}}>{mk.label}</th>
            ))}
          </tr>
          <tr>
            <td style={{...base, background:subColor}} colSpan={2}></td>
            {months.map(mk => (
              <td key={mk.label} style={{...num, background:subColor, textAlign:"center", color:"#555", fontSize:10}}>{mk.days}</td>
            ))}
          </tr>
        </thead>
        <tbody>{rows}</tbody>
      </table>
    );
  }

  // Consolidated HR table — same 4-row header layout as Total Hrs.
  function renderConsolidatedHRTable(
    hr: string,
    pks: string[],
    bgColor: string,
    subColor: string
  ) {
    const activeSections = getActiveSections(pks);
    const totalCols = pks.length;
    const dataMap = dm[hr] ?? {};

    return (
      <table style={{ borderCollapse:"collapse", marginBottom:16 }}>
        <thead>
          {/* Row 1: Month labels */}
          <tr>
            <th style={{...hdr, background:bgColor, minWidth:65}} rowSpan={4}>Unit</th>
            <th style={{...hdr, background:bgColor, minWidth:48}} rowSpan={4}>Plant</th>
            {months.map(mk => (
              <th key={mk.label} style={{...hdr, background:bgColor}} colSpan={totalCols}>{mk.label}</th>
            ))}
          </tr>
          {/* Row 2: Days */}
          <tr>
            {months.map(mk => (
              <th key={mk.label} style={{...hdr, background:subColor, fontSize:10}} colSpan={totalCols}>{mk.days} Days</th>
            ))}
          </tr>
          {/* Row 3: Section groups */}
          <tr>
            {months.map(mk =>
              activeSections.map(s => (
                <th key={`${mk.label}-${s.label}`} style={{...hdr, background:bgColor}} colSpan={s.activeKeys.length}>{s.label}</th>
              ))
            )}
          </tr>
          {/* Row 4: Machine names */}
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
          {UNITS.map((u, i) => (
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

  // Legacy per-machine tables — kept for Remaining & Percentage sections
  function renderProcessTables(
    dataMap: Record<number, Record<string, Record<string, number>>>,
    pks: string[],
    bgColor: string,
    subColor: string
  ) {
    return pks.map(pk =>
      renderTableShell(pk, bgColor, subColor,
        UNITS.map((u,i) => (
          <tr key={u.name} style={{background:i%2===1?"#f2f2f2":"#fff"}}>
            <td style={{...base, textAlign:"left", fontWeight:600, minWidth:55}}>{u.name}</td>
            <td style={{...num, textAlign:"center", color:"#555", minWidth:45}}>{u.plant}</td>
            {months.map(mk => {
              const v = dataMap[u.plant]?.[pk]?.[mk.label] ?? 0;
              return <td key={mk.label} style={{...num, color:v===0?"#bbb":"#000"}}>{v===0?"-":v.toFixed(2)}</td>;
            })}
          </tr>
        ))
      )
    );
  }

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
          {UNITS.map((u, i) => (
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
          {UNITS.map((u, i) => (
            <tr key={u.name} style={{background:i%2===1?"#f2f2f2":"#fff"}}>
              <td style={{...base, textAlign:"left", fontWeight:600}}>{u.name}</td>
              <td style={{...num, textAlign:"center", color:"#555"}}>{u.plant}</td>
              {months.map(mk =>
                activeSections.flatMap(s =>
                  s.activeKeys.map(k => {
                    const mc        = MC[u.name]?.[k] ?? 0;
                    const total     = mc * mk.days * 24;
                    const downForHR = dm[hr]?.[u.plant]?.[k]?.[mk.label] ?? 0;
                    const pct       = total === 0 ? null : (downForHR / total) * 100;
                    const isHigh    = pct !== null && pct > 100;
                    const isEmpty   = pct === null || pct === 0;
                    return (
                      <td key={`${mk.label}-${k}`} style={{...num,
                        color: isEmpty?"#bbb":isHigh?"#c00":"#000",
                        fontWeight: isHigh?700:"normal"
                      }}>
                        {isEmpty ? "-" : `${pct!.toFixed(2)}%`}
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

  function SectionTitle({ label, borderColor }: { label: string; borderColor: string }) {
    return (
      <div style={{ fontWeight:700, fontSize:12, margin:"14px 0 4px", color:"#1f3864",
        borderBottom:`2px solid ${borderColor}`, paddingBottom:2, display:"inline-block" }}>
        {label}
      </div>
    );
  }

  return (
    <div style={{ padding:12, background:"#fff", fontFamily:"Calibri,Arial,sans-serif", fontSize:11 }}>

      <div style={{ display:"flex", alignItems:"center", gap:10, marginBottom:10 }}>
        <b style={{ fontSize:13, color:"#1f3864" }}>Downtime Analysis</b>
        {loaded && <span style={{ color:"#555", fontSize:11 }}>— {fileName}</span>}
        {loaded && (
          <button onClick={reset} style={{ marginLeft:"auto", fontSize:11, padding:"2px 10px", cursor:"pointer" }}>✕ Reset</button>
        )}
      </div>

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
        <div style={{ overflowX:"auto" }}>

          {/* ══ MACHINE COUNT TABLE ══ */}
          <table style={{ borderCollapse:"collapse", marginBottom:16 }}>
            <thead>
              <tr>
                <th style={{...hdr, background:"#dce6f1", minWidth:65}} rowSpan={2}>Unit Name</th>
                <th style={{...hdr, background:"#dce6f1", minWidth:48}} rowSpan={2}>Plant</th>
                {SECTIONS.map(s => (
                  <th key={s.label} style={{...hdr, background:"#dce6f1"}} colSpan={s.keys.length}>{s.label}</th>
                ))}
              </tr>
              <tr>
                {SECTIONS.flatMap(s => s.keys.map(k => (
                  <th key={k} style={{...hdr, background:"#bdd7ee", fontSize:10, minWidth:52}}>{PL[k]}</th>
                )))}
              </tr>
            </thead>
            <tbody>
              {UNITS.map((u,i) => (
                <tr key={u.name} style={{background:i%2===1?"#f2f2f2":"#fff"}}>
                  <td style={{...base, textAlign:"left", fontWeight:600}}>{u.name}</td>
                  <td style={{...num, textAlign:"center", color:"#555"}}>{u.plant}</td>
                  {ALL_KEYS.map(k => (
                    <td key={k} style={num}>{MC[u.name]?.[k]??0}</td>
                  ))}
                </tr>
              ))}
            </tbody>
          </table>

          {/* ══ TOTAL HRS — consolidated multi-header table ══ */}
          <SectionTitle label="Total Hrs." borderColor="#4472c4" />
          {renderTotalHrsTable(activePks)}

          {/* ══ ONE BLOCK PER HEAD REASON ══ */}
          {activeHRs.map(hr => {
            const hrName = hr.charAt(0) + hr.slice(1).toLowerCase();
            const c = HR_COLORS[hr] ?? {
              bg:"#ededed",  sub:"#bfbfbf",  label:`${hrName} Hrs.`,
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
      )}
    </div>
  );
}