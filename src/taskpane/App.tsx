/// <reference types="office-js" />
import * as React from "react";
import { useState, useEffect, useCallback, useRef } from "react";
import {
  AreaChart, Area, XAxis, YAxis, CartesianGrid, Tooltip,
  Legend, ResponsiveContainer, Line,
} from "recharts";

/* ─── constants ─────────────────────────────────────────────────────────────── */
const APP_VERSION = "v0.4.0";
const STAGES = ["ESL", "CTO", "Assembly", "Testing", "FAT"];
const STATUS_LIST = [
  "Backlog","FROM OI TO KO","ANTICIPATION - ENGINEERING ONLY + PROCUREMENT",
  "ANTICIPATION - ENGINEERING ONLY","ON HAND","FORECAST","BACKUP","OUTLOOK","OTHER",
];
const STATUS_STACK_ORDER = [
  "Backlog","FROM OI TO KO","ANTICIPATION - ENGINEERING ONLY + PROCUREMENT",
  "ANTICIPATION - ENGINEERING ONLY","ON HAND","FORECAST","BACKUP","OUTLOOK","OTHER",
];
const STATUS_SHORT: Record<string,string> = {
  "Backlog":"Backlog","FROM OI TO KO":"OI→KO",
  "ANTICIPATION - ENGINEERING ONLY + PROCUREMENT":"Antic+Proc",
  "ANTICIPATION - ENGINEERING ONLY":"Antic Eng",
  "ON HAND":"On Hand","FORECAST":"Forecast","BACKUP":"Backup","OUTLOOK":"Outlook","OTHER":"Other",
};
const STATUS_COLORS: Record<string,string> = {
  "Backlog":"#1e293b","FROM OI TO KO":"#0f4c81",
  "ANTICIPATION - ENGINEERING ONLY + PROCUREMENT":"#1a6b3c",
  "ANTICIPATION - ENGINEERING ONLY":"#2d9c5f",
  "ON HAND":"#166534","FORECAST":"#854d0e","BACKUP":"#6b21a8","OUTLOOK":"#374151","OTHER":"#9ca3af",
};
const CHANGE_COLORS: Record<string,string> = {
  "New":"#16a34a","Changed":"#ca8a04","Removed":"#dc2626","Unchanged":"#94a3b8",
};
const SPECULATIVE_STATUSES = new Set(["BACKUP","OUTLOOK","OTHER"]);

type Tab = "graph"|"heatmap"|"demand"|"adjustments";
type Mode = "constrained"|"unconstrained";

/* ─── types ─────────────────────────────────────────────────────────────────── */
type ResultRow = {
  plant:string;stage:string;week:string;status:string;
  backlog:number;planned:number;total:number;cap:number;util:number;overload:number;
};
type SchedRow = {
  oppId:string;lineId:string;equip:string;plant:string;
  oiDatePlanned:number;oiDate:number;status:string;
  kom:number;ikom:number;bom1:number;frozen:number;finalBom:number;
  asmStart:number;asmFinish:number;tstStart:number;tstFinish:number;
  pdi:number;fat:number;eop:number;fca:number;lt:number;
};
type DemandRow = {
  rowIdx:number;oppId:string;lineId:string;equip:string;routingKey:string;plant:string;
  oiDate:number;oiDatePlanned:number;status:string;priority:string;
  changeFlag:string;changedFields:string;alignFlag:string;
  customer:string;region:string;subRegion:string;country:string;planSpeculative:string;
};
type AdjRow = { rowIdx:number;oppId:string;lineId:string;stage:string;loadOv:string;startWk:string; };
type CapOvRow = { rowIdx:number;plant:string;stage:string;week:string;cap:string; };

/* ─── helpers ───────────────────────────────────────────────────────────────── */
// European date format dd-mm-yyyy
const excelDateToStr = (s:number):string => {
  if (!s) return "";
  const d = new Date((s-25569)*86400000);
  const dd = String(d.getUTCDate()).padStart(2,"0");
  const mm = String(d.getUTCMonth()+1).padStart(2,"0");
  const yyyy = d.getUTCFullYear();
  return `${dd}-${mm}-${yyyy}`;
};
const strToExcelDate = (s:string):number => {
  if (!s) return 0;
  const d = new Date(s+"T00:00:00Z");
  return isNaN(d.getTime()) ? 0 : Math.round(d.getTime()/86400000+25569);
};
function isoWeekLabel(serial:number):string {
  const d = new Date((serial-25569)*86400000);
  const day = (d.getUTCDay()+6)%7;
  d.setUTCDate(d.getUTCDate()-day+3);
  const isoYear = d.getUTCFullYear();
  const jan1 = new Date(Date.UTC(isoYear,0,1));
  const week = Math.floor((d.getTime()-jan1.getTime())/86400000/7)+1;
  return `${isoYear}-${String(week).padStart(2,"0")}`;
}
// Compact week label: "2026-08"
const weekLabel = (serial:number):string => serial ? isoWeekLabel(serial) : "";

const utilColor = (u:number,ov:number) => ov>0?"#dc2626":u>=90?"#ea580c":u>=70?"#ca8a04":u>0?"#16a34a":"#e2e8f0";
const utilText  = (u:number,ov:number) => (ov>0||u>=90)?"#fff":u>0?"#14532d":"#94a3b8";

/* ─── Office.js helpers ─────────────────────────────────────────────────────── */
async function readNamedTable(name:string):Promise<{headers:string[];rows:(string|number)[][]}> {
  return Excel.run(async ctx => {
    const tbl = ctx.workbook.tables.getItem(name);
    const hdrRange = tbl.getHeaderRowRange();
    const bodyRange = tbl.getDataBodyRange();
    hdrRange.load("values"); bodyRange.load("values");
    await ctx.sync();
    const headers = (hdrRange.values[0] as string[]).map(h=>String(h).trim());
    const rows = (bodyRange.values as (string|number)[][]).filter(r=>r[0]!==""&&r[0]!==null);
    return {headers,rows};
  });
}
async function writeCellInTable(tableName:string,rowIdx:number,colName:string,value:string|number,headers:string[]):Promise<void> {
  const colIdx = headers.indexOf(colName);
  if (colIdx<0) throw new Error(`Column "${colName}" not found in ${tableName}`);
  await Excel.run(async ctx => {
    const body = ctx.workbook.tables.getItem(tableName).getDataBodyRange();
    body.getCell(rowIdx,colIdx).values=[[value]];
    await ctx.sync();
  });
}
async function appendRowToTable(tableName:string,row:(string|number)[]):Promise<void> {
  await Excel.run(async ctx => { ctx.workbook.tables.getItem(tableName).rows.add(-1,[row]); await ctx.sync(); });
}
async function updateRowInTable(tableName:string,rowIdx:number,values:(string|number)[]):Promise<void> {
  await Excel.run(async ctx => {
    const body = ctx.workbook.tables.getItem(tableName).getDataBodyRange();
    body.load("rowCount"); await ctx.sync();
    body.getRow(rowIdx).values=[values];
    await ctx.sync();
  });
}
async function deleteRowFromTable(tableName:string,rowIdx:number):Promise<void> {
  await Excel.run(async ctx => {
    const tbl = ctx.workbook.tables.getItem(tableName);
    tbl.getDataBodyRange().getRow(rowIdx).delete(Excel.DeleteShiftDirection.up);
    await ctx.sync();
  });
}
// Batch-write a single column value to multiple rows in one Excel.run
async function batchWriteColumn(tableName:string,rowIndices:number[],colName:string,value:string|number,headers:string[]):Promise<void> {
  const colIdx = headers.indexOf(colName);
  if (colIdx<0) throw new Error(`Column "${colName}" not found in ${tableName}`);
  await Excel.run(async ctx => {
    const body = ctx.workbook.tables.getItem(tableName).getDataBodyRange();
    for (const ri of rowIndices) body.getCell(ri,colIdx).values=[[value]];
    await ctx.sync();
  });
}

async function downloadChartAsPng(containerEl:HTMLElement,filename:string):Promise<void> {
  const svg = containerEl.querySelector("svg");
  if (!svg) throw new Error("No chart found.");
  const bbox = svg.getBoundingClientRect();
  if (bbox.width===0||bbox.height===0) throw new Error("Chart has zero size.");
  const svgClone = svg.cloneNode(true) as SVGSVGElement;
  svgClone.setAttribute("xmlns","http://www.w3.org/2000/svg");
  svgClone.setAttribute("width",String(bbox.width));
  svgClone.setAttribute("height",String(bbox.height));
  svgClone.setAttribute("viewBox",`0 0 ${bbox.width} ${bbox.height}`);
  svgClone.style.fontFamily="Arial, sans-serif";
  const bg = document.createElementNS("http://www.w3.org/2000/svg","rect");
  bg.setAttribute("width","100%"); bg.setAttribute("height","100%"); bg.setAttribute("fill","#ffffff");
  svgClone.insertBefore(bg,svgClone.firstChild);
  const svg64 = btoa(unescape(encodeURIComponent(new XMLSerializer().serializeToString(svgClone))));
  return new Promise((resolve,reject) => {
    const img = new Image();
    img.onload = () => {
      try {
        const scale=2, canvas=document.createElement("canvas");
        canvas.width=bbox.width*scale; canvas.height=bbox.height*scale;
        const c=canvas.getContext("2d");
        if(!c){reject(new Error("Canvas not supported.")); return;}
        c.scale(scale,scale); c.drawImage(img,0,0,bbox.width,bbox.height);
        const link=document.createElement("a");
        link.href=canvas.toDataURL("image/png"); link.download=filename;
        document.body.appendChild(link); link.click(); document.body.removeChild(link);
        resolve();
      } catch(e){reject(e);}
    };
    img.onerror=()=>reject(new Error("Could not render chart."));
    img.src=`data:image/svg+xml;base64,${svg64}`;
  });
}

/* ─── Gantt Modal ────────────────────────────────────────────────────────────── */
function GanttModal({ sched, onClose }: { sched: SchedRow; onClose: () => void }) {
  const phases = [
    { label:"OI",         start:sched.oiDatePlanned||sched.oiDate, end:sched.oiDatePlanned||sched.oiDate, color:"#0f2942", milestone:true },
    { label:"KOM",        start:sched.kom,       end:sched.ikom,      color:"#0f4c81" },
    { label:"1st BOM",    start:sched.bom1,      end:sched.bom1,      color:"#1a6b3c", milestone:true },
    { label:"Frozen",     start:sched.frozen,    end:sched.frozen,    color:"#2d9c5f", milestone:true },
    { label:"Final BOM",  start:sched.finalBom,  end:sched.finalBom,  color:"#166534", milestone:true },
    { label:"Procurement",start:sched.finalBom,  end:sched.asmStart,  color:"#854d0e" },
    { label:"Assembly",   start:sched.asmStart,  end:sched.asmFinish, color:"#0284c7" },
    { label:"Testing",    start:sched.tstStart,  end:sched.tstFinish, color:"#7c3aed" },
    { label:"PDI",        start:sched.pdi,       end:sched.pdi,       color:"#64748b", milestone:true },
    { label:"FAT",        start:sched.fat,       end:sched.fat,       color:"#ca8a04", milestone:true },
    { label:"EOP",        start:sched.eop,       end:sched.eop,       color:"#dc2626", milestone:true },
    { label:"FCA",        start:sched.fca,       end:sched.fca,       color:"#dc2626", milestone:true },
  ].filter(p=>p.start>0);

  const minSerial = Math.min(...phases.map(p=>p.start));
  const maxSerial = Math.max(...phases.map(p=>p.end||p.start));
  const totalDays = Math.max(1, maxSerial - minSerial);

  const pct = (s:number) => Math.max(0,Math.min(100,((s-minSerial)/totalDays)*100));
  const pctW = (s:number,e:number) => Math.max(0.5,((e-s)/totalDays)*100);

  // Generate week axis ticks every ~4 weeks
  const ticks:number[] = [];
  let t = minSerial;
  while(t<=maxSerial){ticks.push(t);t+=28;}
  if(ticks[ticks.length-1]<maxSerial)ticks.push(maxSerial);

  return (
    <div style={gStyles.overlay} onClick={onClose}>
      <div style={gStyles.modal} onClick={e=>e.stopPropagation()}>
        <div style={gStyles.header}>
          <div>
            <div style={gStyles.title}>{sched.equip}</div>
            <div style={gStyles.sub}>{sched.oppId} · {sched.plant} · OI: {excelDateToStr(sched.oiDatePlanned||sched.oiDate)} · LT: {sched.lt}w</div>
          </div>
          <button style={gStyles.closeBtn} onClick={onClose}>✕</button>
        </div>
        {/* Week axis */}
        <div style={gStyles.axisRow}>
          <div style={gStyles.labelCol}/>
          <div style={gStyles.barArea}>
            {ticks.map((tk,i)=>(
              <div key={i} style={{...gStyles.tick,left:`${pct(tk)}%`}}>
                <div style={gStyles.tickLine}/>
                <div style={gStyles.tickLabel}>{weekLabel(tk)}</div>
              </div>
            ))}
          </div>
        </div>
        {/* Phase rows */}
        <div style={gStyles.phaseList}>
          {phases.map((p,i)=>(
            <div key={i} style={gStyles.phaseRow}>
              <div style={gStyles.labelCol}><span style={gStyles.phaseLabel}>{p.label}</span></div>
              <div style={gStyles.barArea}>
                {(p as {milestone?:boolean}).milestone ? (
                  <div style={{...gStyles.diamond,left:`${pct(p.start)}%`,background:p.color}} title={excelDateToStr(p.start)}/>
                ) : (
                  <div style={{...gStyles.bar,left:`${pct(p.start)}%`,width:`${pctW(p.start,p.end)}%`,background:p.color}} title={`${excelDateToStr(p.start)} → ${excelDateToStr(p.end)}`}/>
                )}
              </div>
              <div style={gStyles.dateCol}>
                {(p as {milestone?:boolean}).milestone ? excelDateToStr(p.start) : `${excelDateToStr(p.start)} → ${excelDateToStr(p.end)}`}
              </div>
            </div>
          ))}
        </div>
      </div>
    </div>
  );
}

const gStyles = {
  overlay:{position:"fixed" as const,inset:0,background:"rgba(0,0,0,0.55)",zIndex:1000,display:"flex" as const,alignItems:"center" as const,justifyContent:"center" as const,padding:16},
  modal:{background:"#fff",borderRadius:12,width:"100%",maxWidth:720,maxHeight:"90vh",overflowY:"auto" as const,boxShadow:"0 8px 32px rgba(0,0,0,0.25)"},
  header:{background:"#0f2942",color:"#fff",padding:"12px 16px",borderRadius:"12px 12px 0 0",display:"flex" as const,justifyContent:"space-between" as const,alignItems:"flex-start" as const},
  title:{fontSize:16,fontWeight:700},
  sub:{fontSize:13,color:"#94a3b8",marginTop:2},
  closeBtn:{background:"none",border:"none",color:"#fff",fontSize:18,cursor:"pointer" as const,padding:"0 4px",lineHeight:1},
  axisRow:{display:"flex" as const,alignItems:"flex-end" as const,padding:"8px 16px 0",borderBottom:"1px solid #e2e8f0"},
  phaseList:{padding:"8px 16px 16px"},
  phaseRow:{display:"flex" as const,alignItems:"center" as const,gap:8,marginBottom:10},
  labelCol:{width:90,flexShrink:0},
  phaseLabel:{fontSize:13,fontWeight:600,color:"#475569"},
  barArea:{flex:1,position:"relative" as const,height:22,background:"#f8fafc",borderRadius:4,overflow:"visible" as const},
  bar:{position:"absolute" as const,height:"100%",borderRadius:4,minWidth:4,cursor:"default" as const},
  diamond:{position:"absolute" as const,width:12,height:12,borderRadius:2,transform:"rotate(45deg) translate(-50%,-25%)",top:"50%",cursor:"default" as const},
  tick:{position:"absolute" as const,top:0,transform:"translateX(-50%)",display:"flex" as const,flexDirection:"column" as const,alignItems:"center" as const},
  tickLine:{width:1,height:8,background:"#cbd5e1"},
  tickLabel:{fontSize:10,color:"#94a3b8",whiteSpace:"nowrap" as const,marginTop:1},
  dateCol:{width:200,flexShrink:0,fontSize:12,color:"#64748b",textAlign:"right" as const},
};

/* ─── Stage Adjustment Row ───────────────────────────────────────────────────── */
function StageAdjRow({
  stage, std, ov, exAdj, pendingDelete,
  onChange, onPendingDelete, onDeleteConfirm, onDeleteCancel, dSaving,
  deleteBtnHard, cancelBtn, inpSm, stageName, stageStd, stageDeleteSoft,
}: {
  stage:string; std:number;
  ov:{loadOv:string;startWk:string};
  exAdj:AdjRow|undefined;
  pendingDelete:boolean;
  onChange:(field:"loadOv"|"startWk",val:string)=>void;
  onPendingDelete:()=>void;
  onDeleteConfirm:()=>void;
  onDeleteCancel:()=>void;
  dSaving:boolean;
  deleteBtnHard:React.CSSProperties; cancelBtn:React.CSSProperties;
  inpSm:React.CSSProperties; stageName:React.CSSProperties;
  stageStd:React.CSSProperties; stageDeleteSoft:React.CSSProperties;
}) {
  return (
    <div style={{borderBottom:"1px solid #f1f5f9",paddingBottom:8,marginBottom:8,background:pendingDelete?"#fef2f2":"inherit",borderRadius:4,padding:"6px 4px"}}>
      <div style={{display:"flex" as const,alignItems:"center" as const,gap:8,flexWrap:"wrap" as const}}>
        <div style={stageName}>{stage}</div>
        <div style={stageStd}>std: {std}</div>
        <input style={inpSm} type="number" placeholder={`load (std ${std})`} value={ov.loadOv} onChange={e=>onChange("loadOv",e.target.value)}/>
        <input style={inpSm} placeholder="start wk (YYYY-WW)" value={ov.startWk} onChange={e=>onChange("startWk",e.target.value)}/>
        {exAdj&&!pendingDelete&&(
          <button style={stageDeleteSoft} title="Delete this stage adjustment" onClick={onPendingDelete}>✕</button>
        )}
      </div>
      {exAdj&&pendingDelete&&(
        <div style={{display:"flex" as const,gap:4,alignItems:"center" as const,marginTop:4}}>
          <span style={{fontSize:13,color:"#dc2626"}}>Delete {stage} adjustment?</span>
          <button style={deleteBtnHard} onClick={onDeleteConfirm} disabled={dSaving}>Confirm</button>
          <button style={cancelBtn} onClick={onDeleteCancel}>Cancel</button>
        </div>
      )}
    </div>
  );
}

/* ─── main ──────────────────────────────────────────────────────────────────── */
export default function App() {
  const [tab,setTab] = useState<Tab>("graph");
  const [mode,setMode] = useState<Mode>("unconstrained"); // default unconstrained
  const [loadMsg,setLoadMsg] = useState("");
  const [errMsg,setErrMsg] = useState("");
  const [loading,setLoading] = useState(false);
  const [results,setResults] = useState<ResultRow[]>([]);
  const [schedule,setSchedule] = useState<SchedRow[]>([]);
  const [demand,setDemand] = useState<DemandRow[]>([]);
  const [demandHeaders,setDemandHeaders] = useState<string[]>([]);
  const [adjRows,setAdjRows] = useState<AdjRow[]>([]);
  const [adjHeaders,setAdjHeaders] = useState<string[]>([]);
  const [capOvRows,setCapOvRows] = useState<CapOvRow[]>([]);
  const [plants,setPlants] = useState<string[]>([]);
  const [weekLabelToIndex,setWeekLabelToIndex] = useState<Record<string,number>>({});
  const [indexToWeekLabel,setIndexToWeekLabel] = useState<Record<string,string>>({});
  const [routingKeys,setRoutingKeys] = useState<string[]>([]);
  const [routingLoads,setRoutingLoads] = useState<Record<string,Record<string,number>>>({});
  // graph
  const [gPlants,setGPlants] = useState<string[]>([]);
  const [gStage,setGStage] = useState("Assembly");
  const [gStatuses,setGStatuses] = useState<string[]>(["ON HAND","FORECAST","Backlog"]);
  const [gFrom,setGFrom] = useState("");
  const [gTo,setGTo] = useState("");
  // heatmap
  const [hPlant,setHPlant] = useState("");
  const [hStage,setHStage] = useState("Assembly");
  const [hSelWeek,setHSelWeek] = useState<string>("");
  // demand — default show all
  const [dFilter,setDFilter] = useState<"unassigned"|"all">("all");
  const [dChangeFilter,setDChangeFilter] = useState("");
  const [dStatusFilter,setDStatusFilter] = useState("");
  const [dPlantFilter,setDPlantFilter] = useState("");
  const [dRegionFilter,setDRegionFilter] = useState("");
  const [dCustomerFilter,setDCustomerFilter] = useState("");
  const [dSearch,setDSearch] = useState("");
  const [dEditing,setDEditing] = useState<DemandRow|null>(null);
  const [dRk,setDRk] = useState("");
  const [dPlant,setDPlant] = useState("");
  const [dPriority,setDPriority] = useState("");
  const [dAlignFlag,setDAlignFlag] = useState(false);
  const [dPlanSpec,setDPlanSpec] = useState(false);
  const [dOiPlanned,setDOiPlanned] = useState("");
  const [dOiSpread,setDOiSpread] = useState(true);
  const [dSaving,setDSaving] = useState(false);
  // stage adjustments — dropdown add pattern
  const [dStageOv,setDStageOv] = useState<Record<string,{loadOv:string;startWk:string}>>({});
  const [dActiveStages,setDActiveStages] = useState<string[]>([]);
  const [dStageToAdd,setDStageToAdd] = useState("Assembly");
  const [dStagePendingDelete,setDStagePendingDelete] = useState<string|null>(null);
  // gantt
  const [ganttSched,setGanttSched] = useState<SchedRow|null>(null);
  // adjustments tab
  const [aOppId,setAOppId] = useState("");
  const [aLineId,setALineId] = useState("");
  const [aStage,setAStage] = useState("Assembly");
  const [aLoadOv,setALoadOv] = useState("");
  const [aStartWk,setAStartWk] = useState("");
  const [aSaving,setASaving] = useState(false);
  const [aEditingRow,setAEditingRow] = useState<AdjRow|null>(null);
  const [aPendingDelete,setAPendingDelete] = useState(false);
  const [coEditingRow,setCoEditingRow] = useState<CapOvRow|null>(null);
  const [coPendingDelete,setCoPendingDelete] = useState(false);
  const [coPlant,setCoPlant] = useState("");
  const [coStage,setCoStage] = useState("Assembly");
  const [coWeek,setCoWeek] = useState("");
  const [coCap,setCoCap] = useState("");
  const [coSaving,setCoSaving] = useState(false);
  const [actionNote,setActionNote] = useState("");
  const chartRef = useRef<HTMLDivElement>(null);
  const [shotBusy,setShotBusy] = useState(false);

  const resolveWeekInput = (raw:string):string|number => {
    if (raw==="") return "";
    const t=raw.trim();
    if (/^\d+$/.test(t)) return Number(t);
    const idx=weekLabelToIndex[t];
    return idx!==undefined?idx:"";
  };
  const formatWeek = (raw:string):string => {
    if (raw==="") return "";
    return indexToWeekLabel[raw.trim()]??raw;
  };

  const load = useCallback(async () => {
    setLoading(true); setErrMsg(""); setLoadMsg("Reading tables…");
    try {
      const rTbl = mode==="constrained"?"Results":"ResultsUnconstrained";
      const sTbl = mode==="constrained"?"Schedule":"ScheduleUnconstrained";
      const [rData,sData,dData,aData,cData,rtData,cbData,calData] = await Promise.all([
        readNamedTable(rTbl),readNamedTable(sTbl),readNamedTable("Demand"),
        readNamedTable("Adjustments"),readNamedTable("CapacityOverride"),
        readNamedTable("Routings"),readNamedTable("CapacityBase"),readNamedTable("Calendar"),
      ]);
      setResults(rData.rows.map(r=>({
        plant:String(r[0]),stage:String(r[1]),week:String(r[2]),status:String(r[3]),
        backlog:Number(r[4]),planned:Number(r[5]),total:Number(r[6]),
        cap:Number(r[7]),util:Number(r[8]),overload:Number(r[9]),
      })));
      // Schedule col layout (21 cols):
      // 0=OppID,1=LineID,2=Equip,3=Plant,4=OIDatePlanned,5=OIDate,6=Status,
      // 7=KOM,8=iKOM,9=1stBOM,10=Frozen,11=FinalBOM,
      // 12=AsmStart,13=AsmFinish,14=TstStart,15=TstFinish,
      // 16=PDI,17=FAT,18=EOP,19=FCA,20=LT
      setSchedule(sData.rows.map(r=>({
        oppId:String(r[0]),lineId:String(r[1]),equip:String(r[2]),plant:String(r[3]),
        oiDatePlanned:Number(r[4]),oiDate:Number(r[5]),status:String(r[6]),
        kom:Number(r[7]),ikom:Number(r[8]),bom1:Number(r[9]),
        frozen:Number(r[10]),finalBom:Number(r[11]),
        asmStart:Number(r[12]),asmFinish:Number(r[13]),
        tstStart:Number(r[14]),tstFinish:Number(r[15]),
        pdi:Number(r[16]),fat:Number(r[17]),eop:Number(r[18]),
        fca:Number(r[19]),lt:Number(r[20]),
      })));
      const dh=dData.headers; setDemandHeaders(dh);
      setDemand(dData.rows.map((r,i)=>({
        rowIdx:i,
        oppId:String(r[dh.indexOf("OppID")]??""),
        lineId:String(r[dh.indexOf("LineID")]??""),
        equip:String(r[dh.indexOf("MachineDescription")]??""),
        routingKey:String(r[dh.indexOf("RoutingKey")]??""),
        plant:String(r[dh.indexOf("Plant")]??""),
        oiDate:Number(r[dh.indexOf("OIDate")]??0),
        oiDatePlanned:Number(r[dh.indexOf("OIDatePlanned")]??0),
        status:String(r[dh.indexOf("Status")]??""),
        priority:String(r[dh.indexOf("Priority")]??""),
        changeFlag:String(r[dh.indexOf("ChangeFlag")]??""),
        changedFields:String(r[dh.indexOf("ChangedFields")]??""),
        alignFlag:String(r[dh.indexOf("AlignFlag")]??""),
        customer:String(r[dh.indexOf("Customer")]??""),
        region:String(r[dh.indexOf("Region")]??""),
        subRegion:String(r[dh.indexOf("SubRegion")]??""),
        country:String(r[dh.indexOf("Country")]??""),
        planSpeculative:String(r[dh.indexOf("PlanSpeculative")]??""),
      })));
      const ah=aData.headers; setAdjHeaders(ah);
      setAdjRows(aData.rows.map((r,i)=>({
        rowIdx:i,oppId:String(r[ah.indexOf("OppID")]??""),
        lineId:String(r[ah.indexOf("LineID")]??""),
        stage:String(r[ah.indexOf("Stage")]??""),
        loadOv:String(r[ah.indexOf("LoadOverride")]??""),
        startWk:String(r[ah.indexOf("StartWeek")]??""),
      })));
      const ch=cData.headers;
      setCapOvRows(cData.rows.map((r,i)=>({
        rowIdx:i,plant:String(r[ch.indexOf("Plant")]??""),
        stage:String(r[ch.indexOf("Stage")]??""),
        week:String(r[ch.indexOf("Week")]??""),
        cap:String(r[ch.indexOf("Cap")]??""),
      })));
      const rkIdx=rtData.headers.indexOf("Routing Key");
      setRoutingKeys(Array.from(new Set(rtData.rows.map(r=>String(r[rkIdx]??"")).filter(Boolean))).sort());
      const STAGE_COL:Record<string,string>={ESL:"ESL Load",CTO:"CTO Load",Assembly:"Assembly Load",Testing:"Testing Load",FAT:"FAT Load"};
      const rl:Record<string,Record<string,number>>={};
      for (const r of rtData.rows) {
        const key=String(r[rkIdx]??""); if(!key)continue;
        rl[key]={};
        for (const st of STAGES){const ci=rtData.headers.indexOf(STAGE_COL[st]);rl[key][st]=ci>=0?Number(r[ci]??0):0;}
      }
      setRoutingLoads(rl);
      const ps=new Set<string>();
      cbData.rows.forEach(r=>{const p=String(r[cbData.headers.indexOf("Plant")]??"");if(p)ps.add(p);});
      const pl=Array.from(ps).sort(); setPlants(pl);
      if(!hPlant&&pl.length)setHPlant(pl[0]);
      if(!gPlants.length&&pl.length)setGPlants([pl[0]]);
      if(!coPlant&&pl.length)setCoPlant(pl[0]);
      const wIdx=calData.headers.indexOf("WeekIndex"),wStart=calData.headers.indexOf("WeekStart");
      const l2i:Record<string,number>={},i2l:Record<string,string>={};
      if(wIdx>=0&&wStart>=0){
        for(const r of calData.rows){
          const idx=Number(r[wIdx]),start=Number(r[wStart]);
          if(idx>=1&&start){const lbl=isoWeekLabel(start);l2i[lbl]=idx;i2l[String(idx)]=lbl;}
        }
      }
      setWeekLabelToIndex(l2i); setIndexToWeekLabel(i2l);
      setLoadMsg(`Loaded ${rData.rows.length} capacity rows · ${sData.rows.length} schedule lines`);
    }catch(e:unknown){setErrMsg(String(e));setLoadMsg("");}
    setLoading(false);
  },[mode]);

  useEffect(()=>{load();},[load]);

  /* ── graph data ── */
  const graphData=(()=>{
    const f=results.filter(r=>(gPlants.length===0||gPlants.includes(r.plant))&&r.stage===gStage&&(gStatuses.length===0||gStatuses.includes(r.status))&&(!gFrom||r.week>=gFrom)&&(!gTo||r.week<=gTo));
    const wm:Record<string,Record<string,number>>={};
    for(const r of f){if(!wm[r.week])wm[r.week]={};wm[r.week][r.status]=(wm[r.week][r.status]??0)+r.total;}
    const capByWeek:Record<string,number>={};
    const capSeen=new Set<string>();
    for(const r of f){const k=`${r.week}||${r.plant}`;if(!capSeen.has(k)){capSeen.add(k);capByWeek[r.week]=(capByWeek[r.week]??0)+r.cap;}}
    return Object.keys(wm).sort().map(wk=>{
      const base:Record<string,number|string>={week:wk};
      for(const st of STATUS_STACK_ORDER)base[st]=wm[wk][st]??0;
      base["_cap"]=capByWeek[wk]??0;
      return base;
    });
  })();
  const activeStatuses=gStatuses.length?gStatuses:STATUS_LIST;

  /* ── heatmap ── */
  const heatRows=results.filter(r=>r.plant===hPlant&&r.stage===hStage);
  const weekCapMap:Record<string,number>={};
  for(const r of heatRows)weekCapMap[r.week]=r.cap;
  const hwm:Record<string,{total:number;overload:number}>={};
  for(const r of heatRows){if(!hwm[r.week])hwm[r.week]={total:0,overload:0};hwm[r.week].total+=r.total;hwm[r.week].overload+=r.overload;}
  const heatWeeks=Object.keys(hwm).sort();

  /* ── demand + schedule join ── */
  const schedByLine:Record<string,SchedRow>={};
  for(const s of schedule)schedByLine[s.lineId]=s;

  const alignedOppIds=new Set<string>();
  for(const d of demand){const f=d.alignFlag;if(f!==""&&f!=="0")alignedOppIds.add(d.oppId);}

  // Driver line per aligned deal: max LT among planned lines
  const driverByDeal:Record<string,string>={};
  for(const oppId of Array.from(alignedOppIds)){
    let maxLt=-1,driverLineId="";
    for(const d of demand){
      if(d.oppId!==oppId)continue;
      const s=schedByLine[d.lineId];
      if(s&&s.lt>maxLt){maxLt=s.lt;driverLineId=d.lineId;}
    }
    if(driverLineId)driverByDeal[oppId]=driverLineId;
  }

  const allRegions=Array.from(new Set(demand.map(d=>d.region).filter(Boolean))).sort();
  const allCustomers=Array.from(new Set(demand.map(d=>d.customer).filter(Boolean))).sort();

  const demandFiltered=demand.filter(d=>{
    if(!dChangeFilter&&d.changeFlag==="Removed")return false;
    if(dChangeFilter&&d.changeFlag!==dChangeFilter)return false;
    if(dFilter==="unassigned"&&d.routingKey&&d.plant)return false;
    if(dPlantFilter&&d.plant!==dPlantFilter)return false;
    if(dStatusFilter&&d.status!==dStatusFilter)return false;
    if(dRegionFilter&&d.region!==dRegionFilter)return false;
    if(dCustomerFilter&&d.customer!==dCustomerFilter)return false;
    if(dSearch&&!d.oppId.toLowerCase().includes(dSearch.toLowerCase())&&!d.equip.toLowerCase().includes(dSearch.toLowerCase()))return false;
    return true;
  }).map(d=>({...d,sched:schedByLine[d.lineId]}));

  // Batch-write flag to all deal lines
  const batchWriteFlag = async (oppId:string, colName:string, value:string) => {
    const dealLines = demand.filter(d=>d.oppId===oppId);
    const rowIndices = dealLines.map(d=>d.rowIdx);
    await batchWriteColumn("Demand",rowIndices,colName,value,demandHeaders);
  };

  const toggleAlignFlag=async(d:DemandRow)=>{
    const newVal=(d.alignFlag!==""&&d.alignFlag!=="0")?"":"1";
    try{
      await batchWriteFlag(d.oppId,"AlignFlag",newVal);
      setActionNote(newVal==="1"?`Deal ${d.oppId} flagged for alignment.`:`Alignment flag cleared for deal ${d.oppId}.`);
      await load();
    }catch(e){setActionNote(`Error: ${String(e)}`);}
  };

  const togglePlanSpec=async(d:DemandRow)=>{
    const newVal=(d.planSpeculative!==""&&d.planSpeculative!=="0")?"":"1";
    try{
      await batchWriteFlag(d.oppId,"PlanSpeculative",newVal);
      setActionNote(newVal==="1"?`Deal ${d.oppId} opted in to speculative planning.`:`Speculative flag cleared for deal ${d.oppId}.`);
      await load();
    }catch(e){setActionNote(`Error: ${String(e)}`);}
  };

  const saveDemandLine=async()=>{
    if(!dEditing)return;
    for(const st of dActiveStages){
      const ov=dStageOv[st];
      if(!ov||ov.startWk==="")continue;
      if(resolveWeekInput(ov.startWk)===""){setActionNote(`Could not resolve "${ov.startWk}" (${st}).`);return;}
    }
    setDSaving(true); setActionNote("");
    try{
      const h=demandHeaders;
      if(dRk)await writeCellInTable("Demand",dEditing.rowIdx,"RoutingKey",dRk,h);
      if(dPlant)await writeCellInTable("Demand",dEditing.rowIdx,"Plant",dPlant,h);
      await writeCellInTable("Demand",dEditing.rowIdx,"Priority",dPriority===""?"":Number(dPriority),h);
      // Flags — batch write to all deal lines
      await batchWriteFlag(dEditing.oppId,"AlignFlag",dAlignFlag?"1":"");
      await batchWriteFlag(dEditing.oppId,"PlanSpeculative",dPlanSpec?"1":"");
      // OIDatePlanned
      const serial=dOiPlanned!==""?strToExcelDate(dOiPlanned):0;
      if(dOiPlanned!==""&&serial>0){
        if(dOiSpread){
          const dealLines=demand.filter(d=>d.oppId===dEditing.oppId);
          const oiColIdx=demandHeaders.indexOf("OIDatePlanned");
          if(oiColIdx>=0){
            await Excel.run(async ctx=>{
              const body=ctx.workbook.tables.getItem("Demand").getDataBodyRange();
              for(const dl of dealLines)body.getCell(dl.rowIdx,oiColIdx).values=[[serial]];
              await ctx.sync();
            });
          }
        }else{
          await writeCellInTable("Demand",dEditing.rowIdx,"OIDatePlanned",serial,h);
        }
      }else if(dOiPlanned===""){
        await writeCellInTable("Demand",dEditing.rowIdx,"OIDatePlanned","",h);
      }
      // Stage adjustments
      for(const st of dActiveStages){
        const ov=dStageOv[st];
        if(!ov||(ov.loadOv===""&&ov.startWk===""))continue;
        const ex=adjRows.find(r=>r.oppId===dEditing.oppId&&r.lineId===dEditing.lineId&&r.stage===st);
        const row=[dEditing.oppId,dEditing.lineId,st,ov.loadOv===""?"":Number(ov.loadOv),resolveWeekInput(ov.startWk)] as (string|number)[];
        if(ex)await updateRowInTable("Adjustments",ex.rowIdx,row);
        else await appendRowToTable("Adjustments",row);
      }
      setActionNote(`Saved ${dEditing.oppId}`);
      setDEditing(null); setDStageOv({}); setDActiveStages([]); setDStagePendingDelete(null);
      await load();
    }catch(e){setActionNote(`Error: ${String(e)}`);}
    setDSaving(false);
  };

  const deleteAdjForStage=async(stage:string)=>{
    if(!dEditing)return;
    const ex=adjRows.find(r=>r.oppId===dEditing.oppId&&r.lineId===dEditing.lineId&&r.stage===stage);
    if(!ex)return;
    setDSaving(true); setActionNote("");
    try{
      await deleteRowFromTable("Adjustments",ex.rowIdx);
      setActionNote(`Deleted ${stage} adjustment for ${dEditing.oppId}.`);
      setDStageOv(prev=>({...prev,[stage]:{loadOv:"",startWk:""}}));
      setDActiveStages(prev=>prev.filter(s=>s!==stage));
      setDStagePendingDelete(null);
      await load();
    }catch(e){setActionNote(`Error: ${String(e)}`);}
    setDSaving(false);
  };

  const saveAdj=async()=>{
    if(!aOppId||!aLineId||!aStage){setActionNote("Fill OppID, LineID and Stage.");return;}
    const wk=resolveWeekInput(aStartWk);
    if(aStartWk!==""&&wk===""){setActionNote(`Could not resolve "${aStartWk}".`);return;}
    setASaving(true); setActionNote("");
    try{
      const ex=aEditingRow??adjRows.find(r=>r.oppId===aOppId&&r.lineId===aLineId&&r.stage===aStage);
      const row=[aOppId,aLineId,aStage,aLoadOv===""?"":Number(aLoadOv),wk] as (string|number)[];
      if(ex){await updateRowInTable("Adjustments",ex.rowIdx,row);setActionNote(`Updated ${aOppId}/${aStage}`);}
      else{await appendRowToTable("Adjustments",row);setActionNote(`Added ${aOppId}/${aStage}`);}
      setAEditingRow(null); setAPendingDelete(false); await load();
    }catch(e){setActionNote(`Error: ${String(e)}`);}
    setASaving(false);
  };

  const deleteAdj=async()=>{
    if(!aEditingRow)return;
    setASaving(true); setActionNote("");
    try{
      await deleteRowFromTable("Adjustments",aEditingRow.rowIdx);
      setActionNote(`Deleted ${aEditingRow.oppId}/${aEditingRow.stage}.`);
      setAOppId(""); setALineId(""); setALoadOv(""); setAStartWk("");
      setAEditingRow(null); setAPendingDelete(false); await load();
    }catch(e){setActionNote(`Error: ${String(e)}`);}
    setASaving(false);
  };

  const saveCapOv=async()=>{
    if(!coPlant||!coStage||!coWeek||!coCap){setActionNote("Fill all override fields.");return;}
    const wk=resolveWeekInput(coWeek);
    if(wk===""){setActionNote(`Could not resolve "${coWeek}".`);return;}
    setCoSaving(true); setActionNote("");
    try{
      const ex=coEditingRow??capOvRows.find(r=>r.plant===coPlant&&r.stage===coStage&&r.week===String(wk));
      const row=[coPlant,coStage,wk,Number(coCap)] as (string|number)[];
      if(ex){await updateRowInTable("CapacityOverride",ex.rowIdx,row);setActionNote(`Updated ${coPlant}/${coStage}/W${wk}`);}
      else{await appendRowToTable("CapacityOverride",row);setActionNote(`Added ${coPlant}/${coStage}/W${wk}`);}
      setCoEditingRow(null); setCoPendingDelete(false); await load();
    }catch(e){setActionNote(`Error: ${String(e)}`);}
    setCoSaving(false);
  };

  const deleteCapOv=async()=>{
    if(!coEditingRow)return;
    setCoSaving(true); setActionNote("");
    try{
      await deleteRowFromTable("CapacityOverride",coEditingRow.rowIdx);
      setActionNote(`Deleted ${coEditingRow.plant}/${coEditingRow.stage}/W${coEditingRow.week}.`);
      setCoPlant(plants[0]??""); setCoStage("Assembly"); setCoWeek(""); setCoCap("");
      setCoEditingRow(null); setCoPendingDelete(false); await load();
    }catch(e){setActionNote(`Error: ${String(e)}`);}
    setCoSaving(false);
  };

  const s=styles;
  return (
    <div style={s.shell}>
      {ganttSched&&<GanttModal sched={ganttSched} onClose={()=>setGanttSched(null)}/>}
      <div style={s.header}>
        <div style={s.headerTitle}>S&OP Planning <span style={s.versionTag}>{APP_VERSION}</span></div>
        <div style={s.modeToggle}>
          <button style={mode==="constrained"?s.modeActive:s.modeBtn} onClick={()=>setMode("constrained")}>Constrained</button>
          <button style={mode==="unconstrained"?s.modeActive:s.modeBtn} onClick={()=>setMode("unconstrained")}>Unconstrained</button>
        </div>
      </div>
      <div style={s.tabs}>
        {([["graph","Load"],["heatmap","Heat"],["demand","Demand"],["adjustments","Adjust"]] as [Tab,string][]).map(([t,label])=>(
          <button key={t} style={tab===t?s.tabActive:s.tab} onClick={()=>setTab(t)}>{label}</button>
        ))}
        <button style={s.refreshBtn} onClick={load} disabled={loading} title="Refresh">{loading?"…":"↻"}</button>
      </div>
      {(loadMsg||errMsg)&&<div style={errMsg?s.errorBar:s.statusBar}>{errMsg||loadMsg}</div>}
      {actionNote&&<div style={s.actionBar}>{actionNote}</div>}

      {/* ── GRAPH ── */}
      {tab==="graph"&&(
        <div style={s.content}>
          <div style={s.filterGrid2}>
            <div>
              <div style={s.sectionLbl}>Plants</div>
              <div style={s.multiBox}>{plants.map(p=>(
                <label key={p} style={s.checkRow}>
                  <input type="checkbox" checked={gPlants.includes(p)} onChange={e=>setGPlants(e.target.checked?[...gPlants,p]:gPlants.filter(x=>x!==p))}/>{p}
                </label>
              ))}</div>
              <div style={{...s.sectionLbl,marginTop:10}}>Status</div>
              <div style={s.multiBox}>{STATUS_LIST.map(st=>(
                <label key={st} style={s.checkRow}>
                  <input type="checkbox" checked={gStatuses.includes(st)} onChange={e=>setGStatuses(e.target.checked?[...gStatuses,st]:gStatuses.filter(x=>x!==st))}/>
                  <span style={{...s.dot,background:STATUS_COLORS[st]}}/>{STATUS_SHORT[st]}
                </label>
              ))}</div>
            </div>
            <div>
              <div style={s.sectionLbl}>Stage</div>
              <select style={s.sel} value={gStage} onChange={e=>setGStage(e.target.value)}>{STAGES.map(st=><option key={st}>{st}</option>)}</select>
              <div style={{...s.sectionLbl,marginTop:10}}>From week</div>
              <input style={s.inp} placeholder="2026-01" value={gFrom} onChange={e=>setGFrom(e.target.value)}/>
              <div style={{...s.sectionLbl,marginTop:8}}>To week</div>
              <input style={s.inp} placeholder="2028-52" value={gTo} onChange={e=>setGTo(e.target.value)}/>
            </div>
          </div>
          {graphData.length===0?<div style={s.empty}>No data for current filters.</div>:(<>
            <div style={{display:"flex",justifyContent:"flex-end",marginBottom:6}}>
              <button style={s.shotBtn} disabled={shotBusy} onClick={async()=>{
                if(!chartRef.current)return; setShotBusy(true); setActionNote("");
                try{await downloadChartAsPng(chartRef.current,`sop-${gStage}-${new Date().toISOString().slice(0,10)}.png`);setActionNote("Chart saved.");}
                catch(e){setActionNote(`Error: ${String(e)}`);}
                setShotBusy(false);
              }}>{shotBusy?"Saving…":"Save as PNG"}</button>
            </div>
            <div ref={chartRef}>
              <ResponsiveContainer width="100%" height={300}>
                <AreaChart data={graphData} margin={{top:4,right:8,left:0,bottom:64}}>
                  <CartesianGrid strokeDasharray="3 3" stroke="#e2e8f0"/>
                  <XAxis dataKey="week" tick={{fontSize:10}} angle={-60} textAnchor="end" interval="preserveStartEnd"/>
                  <YAxis tick={{fontSize:11}}/>
                  <Tooltip contentStyle={{fontSize:12}}/>
                  <Legend wrapperStyle={{fontSize:11}}/>
                  {STATUS_STACK_ORDER.filter(st=>activeStatuses.includes(st)&&graphData.some(d=>(d as Record<string,number>)[st]!=null)).map(st=>(
                    <Area key={st} type="monotone" dataKey={st} name={STATUS_SHORT[st]} stackId="1"
                      stroke={STATUS_COLORS[st]} fill={STATUS_COLORS[st]} fillOpacity={0.75} strokeWidth={1} dot={false} connectNulls/>
                  ))}
                  <Line type="monotone" dataKey="_cap" name="Capacity" stroke="#dc2626" strokeWidth={2} dot={false} connectNulls legendType="line" isAnimationActive={false}/>
                </AreaChart>
              </ResponsiveContainer>
            </div>
          </>)}
          <div style={s.coInlineCard}>
            <div style={s.sectionLbl}>Quick capacity override</div>
            <div style={s.filterRow}>
              <label style={s.fieldLbl}>Plant<select style={s.sel} value={coPlant} onChange={e=>setCoPlant(e.target.value)}>{plants.map(p=><option key={p}>{p}</option>)}</select></label>
              <label style={s.fieldLbl}>Stage<select style={s.sel} value={coStage} onChange={e=>setCoStage(e.target.value)}>{STAGES.map(st=><option key={st}>{st}</option>)}</select></label>
              <label style={s.fieldLbl}>Week<input style={s.inp} value={coWeek} onChange={e=>setCoWeek(e.target.value)} placeholder="e.g. 2026-14"/></label>
              <label style={s.fieldLbl}>Capacity<input style={s.inp} type="number" value={coCap} onChange={e=>setCoCap(e.target.value)}/></label>
              <button style={s.saveBtnInline} onClick={saveCapOv} disabled={coSaving}>{coSaving?"Saving…":"Save"}</button>
            </div>
          </div>
        </div>
      )}

      {/* ── HEATMAP ── */}
      {tab==="heatmap"&&(
        <div style={s.content}>
          <div style={s.filterRow}>
            <label style={s.fieldLbl}>Plant<select style={s.sel} value={hPlant} onChange={e=>setHPlant(e.target.value)}>{plants.map(p=><option key={p}>{p}</option>)}</select></label>
            <label style={s.fieldLbl}>Stage<select style={s.sel} value={hStage} onChange={e=>setHStage(e.target.value)}>{STAGES.map(st=><option key={st}>{st}</option>)}</select></label>
          </div>
          <div style={s.legend}>{[["#16a34a","<70%"],["#ca8a04","70–90%"],["#ea580c","≥90%"],["#dc2626","Overload"]].map(([c,l])=>(
            <span key={l} style={s.legendItem}><span style={{...s.dot,background:c}}/>{l}</span>
          ))}</div>
          {heatWeeks.length===0?<div style={s.empty}>No data for {hPlant}·{hStage}.</div>:(
            <div style={s.heatGrid}>{heatWeeks.map(wk=>{
              const w=hwm[wk],cap=weekCapMap[wk]??0;
              const util=cap?Math.round(w.total/cap*100):0;
              const bg=utilColor(util,w.overload),fg=utilText(util,w.overload);
              const selected=hSelWeek===wk;
              return(
                <div key={wk} style={{...s.cell,background:bg,cursor:"pointer",outline:selected?"2px solid #0f2942":"none",outlineOffset:1}}
                  onClick={()=>{setHSelWeek(selected?"":wk);if(!selected){setCoPlant(hPlant);setCoStage(hStage);setCoWeek(wk);setCoCap(String(cap));setCoEditingRow(null);setCoPendingDelete(false);}}}
                  title={heatRows.filter(r=>r.week===wk).map(b=>`${b.status}:${Math.round(b.total)}`).join("\n")+`\nTotal:${Math.round(w.total)}/${cap}`}>
                  <div style={{...s.cellWk,color:fg}}>{wk}</div>
                  <div style={{...s.cellUtil,color:fg}}>{util}%</div>
                  {w.overload>0&&<div style={s.cellOver}>+{Math.round(w.overload)}</div>}
                </div>
              );
            })}</div>
          )}
          {hSelWeek&&(
            <div style={s.coInlineCard}>
              <div style={s.sectionLbl}>Adjust capacity — {hPlant} · {hStage} · {hSelWeek}</div>
              <div style={s.filterRow}>
                <label style={s.fieldLbl}>New capacity<input style={s.inp} type="number" value={coCap} onChange={e=>setCoCap(e.target.value)}/></label>
                <button style={s.saveBtnInline} onClick={saveCapOv} disabled={coSaving}>{coSaving?"Saving…":"Save"}</button>
                <button style={s.cancelBtn} onClick={()=>setHSelWeek("")}>Clear</button>
              </div>
            </div>
          )}
          {heatWeeks.length>0&&(
            <div style={{marginTop:14}}>
              <div style={s.sectionLbl}>Breakdown{hSelWeek?` — ${hSelWeek}`:""}</div>
              <div style={s.tableScroll}>
                <table style={s.table}><thead><tr>
                  <th style={s.th}>Week</th><th style={s.th}>Status</th>
                  <th style={s.thR}>Load</th><th style={s.thR}>Cap</th><th style={s.thR}>Util%</th>
                </tr></thead><tbody>
                  {heatRows.filter(r=>!hSelWeek||r.week===hSelWeek).sort((a,b)=>a.week.localeCompare(b.week)||a.status.localeCompare(b.status)).slice(0,200).map((r,i)=>(
                    <tr key={i} style={i%2===0?s.trEven:s.trOdd}>
                      <td style={s.td}>{r.week}</td>
                      <td style={s.td}><span style={{...s.badge,background:STATUS_COLORS[r.status]??"#6b7280"}}>{STATUS_SHORT[r.status]??r.status}</span></td>
                      <td style={s.tdR}>{Math.round(r.total)}</td><td style={s.tdR}>{r.cap}</td>
                      <td style={{...s.tdR,color:r.overload>0?"#dc2626":"inherit",fontWeight:r.overload>0?700:400}}>{r.util}%</td>
                    </tr>
                  ))}
                </tbody></table>
              </div>
            </div>
          )}
        </div>
      )}

      {/* ── DEMAND ── */}
      {tab==="demand"&&(
        <div style={s.content}>
          <div style={s.filterRow}>
            <label style={s.fieldLbl}>Show
              <select style={s.sel} value={dFilter} onChange={e=>setDFilter(e.target.value as "unassigned"|"all")}>
                <option value="all">All lines</option>
                <option value="unassigned">Unassigned only</option>
              </select>
            </label>
            <label style={s.fieldLbl}>Plant
              <select style={s.sel} value={dPlantFilter} onChange={e=>setDPlantFilter(e.target.value)}>
                <option value="">All plants</option>{plants.map(p=><option key={p}>{p}</option>)}
              </select>
            </label>
            <label style={s.fieldLbl}>Change
              <select style={s.sel} value={dChangeFilter} onChange={e=>setDChangeFilter(e.target.value)}>
                <option value="">All (hide Removed)</option><option value="New">New</option>
                <option value="Changed">Changed</option><option value="Unchanged">Unchanged</option><option value="Removed">Removed</option>
              </select>
            </label>
            <label style={s.fieldLbl}>Status
              <select style={s.sel} value={dStatusFilter} onChange={e=>setDStatusFilter(e.target.value)}>
                <option value="">All statuses</option>{STATUS_LIST.filter(st=>st!=="Backlog").map(st=><option key={st} value={st}>{STATUS_SHORT[st]}</option>)}
              </select>
            </label>
            <label style={s.fieldLbl}>Region
              <select style={s.sel} value={dRegionFilter} onChange={e=>setDRegionFilter(e.target.value)}>
                <option value="">All regions</option>{allRegions.map(r=><option key={r}>{r}</option>)}
              </select>
            </label>
            <label style={s.fieldLbl}>Customer
              <select style={s.sel} value={dCustomerFilter} onChange={e=>setDCustomerFilter(e.target.value)}>
                <option value="">All customers</option>{allCustomers.map(c=><option key={c}>{c}</option>)}
              </select>
            </label>
            <label style={s.fieldLbl}>Search
              <input style={s.inp} placeholder="OppID or equipment…" value={dSearch} onChange={e=>setDSearch(e.target.value)}/>
            </label>
          </div>

          {dEditing&&(
            <div style={s.editCard}>
              <div style={s.editTitle}>{dEditing.oppId} · {dEditing.equip}</div>
              {(dEditing.customer||dEditing.region)&&(
                <div style={s.contextNote}>{[dEditing.customer,dEditing.region,dEditing.subRegion,dEditing.country].filter(Boolean).join(" · ")}</div>
              )}
              {dEditing.changedFields&&(
                <div style={s.changedNote}>
                  <span style={{...s.badge,background:CHANGE_COLORS[dEditing.changeFlag]??"#6b7280",marginRight:6}}>{dEditing.changeFlag}</span>
                  {dEditing.changedFields}
                </div>
              )}
              <div style={s.filterRow}>
                <label style={s.fieldLbl}>Routing Key
                  <select style={s.sel} value={dRk} onChange={e=>setDRk(e.target.value)}>
                    <option value="">— select —</option>{routingKeys.map(k=><option key={k}>{k}</option>)}
                  </select>
                </label>
                <label style={s.fieldLbl}>Plant
                  <select style={s.sel} value={dPlant} onChange={e=>setDPlant(e.target.value)}>
                    <option value="">— select —</option>{plants.map(p=><option key={p}>{p}</option>)}
                  </select>
                </label>
                <label style={s.fieldLbl}>Priority
                  <input style={s.inp} type="number" placeholder="blank = auto" value={dPriority} onChange={e=>setDPriority(e.target.value)}/>
                </label>
              </div>

              {/* OI Planned */}
              <div style={s.oiRow}>
                <div style={s.oiBlock}>
                  <div style={s.sectionLbl}>OI Live</div>
                  <div style={s.oiValue}>{excelDateToStr(dEditing.oiDate)}</div>
                </div>
                <div style={s.oiBlock}>
                  <div style={s.sectionLbl}>OI Planned <span style={s.oiHint}>(blank = inherit live OI)</span></div>
                  <input style={{...s.inp,borderColor:dOiPlanned&&dOiPlanned!==excelDateToStr(dEditing.oiDate)?"#ca8a04":"#cbd5e1",background:dOiPlanned&&dOiPlanned!==excelDateToStr(dEditing.oiDate)?"#fffbeb":"#fff"}}
                    type="date" value={dOiPlanned} onChange={e=>setDOiPlanned(e.target.value)}/>
                  <label style={s.oiSpreadLabel}>
                    <input type="checkbox" checked={dOiSpread} onChange={e=>setDOiSpread(e.target.checked)} style={{marginRight:5}}/>
                    <span>Apply to all equipment in deal <strong>{dEditing.oppId}</strong></span>
                  </label>
                </div>
              </div>

              {/* Flags */}
              <div style={s.alignFlagRow}>
                <label style={s.alignFlagLabel}>
                  <input type="checkbox" checked={dAlignFlag} onChange={e=>setDAlignFlag(e.target.checked)} style={{marginRight:6}}/>
                  <span style={s.alignFlagText}>Align deal LT</span>
                  <span style={s.alignFlagHint}> — all equipment in deal {dEditing.oppId} finishes within the same window</span>
                </label>
              </div>
              {SPECULATIVE_STATUSES.has(dEditing.status.toUpperCase())&&(
                <div style={{...s.alignFlagRow,borderColor:"#f9a8d4",background:"#fdf2f8"}}>
                  <label style={s.alignFlagLabel}>
                    <input type="checkbox" checked={dPlanSpec} onChange={e=>setDPlanSpec(e.target.checked)} style={{marginRight:6}}/>
                    <span style={{...s.alignFlagText,color:"#9d174d"}}>Include in planning</span>
                    <span style={s.alignFlagHint}> — speculative status ({STATUS_SHORT[dEditing.status]??dEditing.status}); off by default</span>
                  </label>
                </div>
              )}

              {/* Stage adjustments — dropdown add pattern */}
              <div style={{...s.sectionLbl,marginTop:12}}>Stage adjustments</div>
              {dActiveStages.map(st=>{
                const std=dRk?(routingLoads[dRk]?.[st]??0):(dEditing.routingKey?(routingLoads[dEditing.routingKey]?.[st]??0):0);
                const ov=dStageOv[st]??{loadOv:"",startWk:""};
                const exAdj=adjRows.find(r=>r.oppId===dEditing.oppId&&r.lineId===dEditing.lineId&&r.stage===st);
                return(
                  <StageAdjRow key={st} stage={st} std={std} ov={ov} exAdj={exAdj}
                    pendingDelete={dStagePendingDelete===st}
                    onChange={(field,val)=>setDStageOv({...dStageOv,[st]:{...ov,[field]:val}})}
                    onPendingDelete={()=>setDStagePendingDelete(st)}
                    onDeleteConfirm={()=>deleteAdjForStage(st)}
                    onDeleteCancel={()=>setDStagePendingDelete(null)}
                    dSaving={dSaving}
                    deleteBtnHard={s.deleteBtnHard} cancelBtn={s.cancelBtn}
                    inpSm={s.inpSm} stageName={s.stageName} stageStd={s.stageStd}
                    stageDeleteSoft={s.stageDeleteSoft}
                  />
                );
              })}
              <div style={{display:"flex",gap:8,alignItems:"center",marginTop:6,marginBottom:10}}>
                <select style={{...s.sel,flex:1}} value={dStageToAdd} onChange={e=>setDStageToAdd(e.target.value)}>
                  {STAGES.filter(st=>!dActiveStages.includes(st)).map(st=><option key={st}>{st}</option>)}
                </select>
                <button style={s.addStageBtn} onClick={()=>{
                  if(!dActiveStages.includes(dStageToAdd)){
                    setDActiveStages([...dActiveStages,dStageToAdd]);
                    // pre-fill from existing adj if present
                    const ex=adjRows.find(r=>r.oppId===dEditing.oppId&&r.lineId===dEditing.lineId&&r.stage===dStageToAdd);
                    setDStageOv({...dStageOv,[dStageToAdd]:{loadOv:ex?.loadOv??"",startWk:ex?.startWk?formatWeek(ex.startWk):""}});
                  }
                  // pick next available stage for the dropdown
                  const remaining=STAGES.filter(st=>st!==dStageToAdd&&![...dActiveStages,dStageToAdd].includes(st));
                  if(remaining.length)setDStageToAdd(remaining[0]);
                }}>+ Add stage</button>
              </div>

              <div style={{display:"flex",gap:8,marginTop:4}}>
                <button style={s.saveBtn} onClick={saveDemandLine} disabled={dSaving}>{dSaving?"Saving…":"Save"}</button>
                <button style={s.cancelBtn} onClick={()=>{setDEditing(null);setDStageOv({});setDActiveStages([]);setDStagePendingDelete(null);}}>Cancel</button>
              </div>
            </div>
          )}

          <div style={s.sectionLbl}>
            {demandFiltered.length} lines
            {alignedOppIds.size>0&&<span style={s.alignLegend}>· <span style={s.alignDot}>⟳</span> = aligned · <span style={{color:"#ca8a04",fontWeight:700}}>●</span> = LT driver</span>}
          </div>
          <div style={s.tableScrollFlex}>
            <table style={s.table}><thead><tr>
              <th style={s.th}>OppID</th><th style={s.th}>Equipment</th>
              <th style={s.th}>Customer</th><th style={s.th}>Region</th>
              <th style={s.th}>Routing</th><th style={s.th}>Plant</th><th style={s.thR}>Priority</th>
              <th style={s.th}>Change</th><th style={s.th}>Status</th>
              <th style={s.thR}>OI Planned</th><th style={s.thR}>OI Live</th>
              <th style={s.thR}>Asm Start</th><th style={s.thR}>Asm Finish</th><th style={s.thR}>FCA</th>
              <th style={s.thR}>LT</th>
              <th style={s.th}>Align</th><th style={s.th}>Plan?</th>
              <th style={s.th}></th>
            </tr></thead><tbody>
              {demandFiltered.map((d,i)=>{
                const isAligned=d.alignFlag!==""&&d.alignFlag!=="0";
                const dealAligned=alignedOppIds.has(d.oppId);
                const isPlanSpec=d.planSpeculative!==""&&d.planSpeculative!=="0";
                const isSpecStatus=SPECULATIVE_STATUSES.has(d.status.toUpperCase());
                const oiDiverged=d.oiDatePlanned>0&&d.oiDatePlanned!==d.oiDate;
                const isDriver=dealAligned&&driverByDeal[d.oppId]===d.lineId;
                return(
                  <tr key={i} style={dEditing?.lineId===d.lineId?s.trSelected:i%2===0?s.trEven:s.trOdd}>
                    <td style={s.tdMono}>
                      {dealAligned&&<span style={s.alignIndicator} title="Deal alignment active">⟳ </span>}
                      {d.oppId}
                    </td>
                    <td style={{...s.td,maxWidth:90,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}} title={d.equip}>{d.equip}</td>
                    <td style={{...s.td,maxWidth:80,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}} title={d.customer}>{d.customer||"—"}</td>
                    <td style={s.td}>{d.region||"—"}</td>
                    <td style={{...s.td,color:d.routingKey?"inherit":"#dc2626"}}>{d.routingKey||"—"}</td>
                    <td style={{...s.td,color:d.plant?"inherit":"#dc2626"}}>{d.plant||"—"}</td>
                    <td style={s.tdR}>{d.priority||""}</td>
                    <td style={s.td}>{d.changeFlag&&<span style={{...s.badge,background:CHANGE_COLORS[d.changeFlag]??"#6b7280"}}>{d.changeFlag}</span>}</td>
                    <td style={s.td}>{d.status?<span style={{...s.badge,background:STATUS_COLORS[d.status]??"#6b7280"}}>{STATUS_SHORT[d.status]??d.status.split(" ")[0]}</span>:<span style={{color:"#94a3b8"}}>—</span>}</td>
                    <td style={{...s.tdR,background:oiDiverged?"#fffbeb":"inherit",color:oiDiverged?"#92400e":"inherit",fontWeight:oiDiverged?700:400}} title={oiDiverged?`Live: ${excelDateToStr(d.oiDate)}`:"Matches live OI"}>
                      {d.oiDatePlanned?excelDateToStr(d.oiDatePlanned):excelDateToStr(d.oiDate)}{oiDiverged&&" ⚠"}
                    </td>
                    <td style={s.tdR}>{excelDateToStr(d.oiDate)}</td>
                    {/* Assembly Start/Finish as week labels */}
                    <td style={s.tdR}>{d.sched?weekLabel(d.sched.asmStart):""}</td>
                    <td style={s.tdR}>{d.sched?weekLabel(d.sched.asmFinish):""}</td>
                    <td style={s.tdR}>{d.sched?excelDateToStr(d.sched.fca):""}</td>
                    {/* LT — yellow highlight for driver */}
                    <td style={{...s.tdR,fontWeight:700,background:isDriver?"#fef08a":"inherit",color:isDriver?"#713f12":"inherit"}} title={isDriver?"LT driver for this deal":undefined}>
                      {d.sched?.lt??""}{isDriver&&" ●"}
                    </td>
                    <td style={s.td}>
                      <button style={isAligned?s.alignBtnOn:s.alignBtnOff} onClick={()=>toggleAlignFlag(d)}>
                        {isAligned?"⟳ On":"Off"}
                      </button>
                    </td>
                    <td style={s.td}>
                      {isSpecStatus?(
                        <button style={isPlanSpec?s.planSpecBtnOn:s.planSpecBtnOff} onClick={()=>togglePlanSpec(d)}>
                          {isPlanSpec?"✓ Yes":"No"}
                        </button>
                      ):<span style={{color:"#cbd5e1",fontSize:13}}>—</span>}
                    </td>
                    <td style={s.td}>
                      <div style={{display:"flex",gap:4}}>
                        <button style={s.editBtn} onClick={()=>{
                          setDEditing(d);setDRk(d.routingKey);setDPlant(d.plant);setDPriority(d.priority);
                          setDAlignFlag(d.alignFlag!==""&&d.alignFlag!=="0");
                          setDPlanSpec(d.planSpeculative!==""&&d.planSpeculative!=="0");
                          setDOiPlanned(d.oiDatePlanned?excelDateToStr(d.oiDatePlanned):"");
                          setDOiSpread(true); setDStagePendingDelete(null);
                          // Seed active stages from existing adj rows
                          const existingStages=adjRows.filter(r=>r.oppId===d.oppId&&r.lineId===d.lineId).map(r=>r.stage);
                          setDActiveStages(existingStages);
                          const seed:Record<string,{loadOv:string;startWk:string}>={};
                          for(const st of existingStages){
                            const ex=adjRows.find(r=>r.oppId===d.oppId&&r.lineId===d.lineId&&r.stage===st);
                            seed[st]={loadOv:ex?.loadOv??"",startWk:ex?.startWk?formatWeek(ex.startWk):""};
                          }
                          setDStageOv(seed);
                          const remaining=STAGES.filter(st=>!existingStages.includes(st));
                          if(remaining.length)setDStageToAdd(remaining[0]);
                        }}>Edit</button>
                        {d.sched&&(
                          <button style={s.ganttBtn} title="View Gantt" onClick={()=>setGanttSched(d.sched!)}>📅</button>
                        )}
                      </div>
                    </td>
                  </tr>
                );
              })}
            </tbody></table>
          </div>
        </div>
      )}

      {/* ── ADJUSTMENTS ── */}
      {tab==="adjustments"&&(
        <div style={s.content}>
          <div style={s.sectionLbl}>Line adjustment</div>
          <div style={s.editCard}>
            <div style={s.filterRow}>
              <label style={s.fieldLbl}>OppID<input style={s.inp} value={aOppId} onChange={e=>setAOppId(e.target.value)} placeholder="e.g. 4711"/></label>
              <label style={s.fieldLbl}>LineID<input style={s.inp} value={aLineId} onChange={e=>setALineId(e.target.value)} placeholder="e.g. L001"/></label>
              <label style={s.fieldLbl}>Stage<select style={s.sel} value={aStage} onChange={e=>setAStage(e.target.value)}>{STAGES.map(st=><option key={st}>{st}</option>)}</select></label>
            </div>
            <div style={s.filterRow}>
              <label style={s.fieldLbl}>Load override<input style={s.inp} type="number" value={aLoadOv} onChange={e=>setALoadOv(e.target.value)} placeholder="blank = routing default"/></label>
              <label style={s.fieldLbl}>Start week<input style={s.inp} value={aStartWk} onChange={e=>setAStartWk(e.target.value)} placeholder="blank = computed"/></label>
            </div>
            <div style={{display:"flex",gap:8,alignItems:"center",flexWrap:"wrap" as const}}>
              <button style={s.saveBtn} onClick={saveAdj} disabled={aSaving}>{aSaving?"Saving…":"Save adjustment"}</button>
              {aEditingRow&&!aPendingDelete&&(<button style={s.deleteBtnSoft} onClick={()=>setAPendingDelete(true)}>Delete</button>)}
              {aEditingRow&&aPendingDelete&&(<><button style={s.deleteBtnHard} onClick={deleteAdj} disabled={aSaving}>Confirm delete</button><button style={s.cancelBtn} onClick={()=>setAPendingDelete(false)}>Cancel</button></>)}
              {aEditingRow&&<button style={s.cancelBtn} onClick={()=>{setAEditingRow(null);setAPendingDelete(false);setAOppId("");setALineId("");setALoadOv("");setAStartWk("");}}>Clear</button>}
            </div>
          </div>
          {adjRows.length>0&&(<>
            <div style={{...s.sectionLbl,marginTop:14}}>Existing ({adjRows.length})</div>
            <div style={s.tableScroll}><table style={s.table}><thead><tr>
              <th style={s.th}>OppID</th><th style={s.th}>LineID</th><th style={s.th}>Stage</th>
              <th style={s.thR}>LoadOv</th><th style={s.thR}>StartWk</th><th style={s.th}></th>
            </tr></thead><tbody>
              {adjRows.map((r,i)=>(
                <tr key={i} style={aEditingRow?.rowIdx===r.rowIdx?s.trSelected:i%2===0?s.trEven:s.trOdd}>
                  <td style={s.tdMono}>{r.oppId}</td><td style={s.tdMono}>{r.lineId}</td>
                  <td style={s.td}>{r.stage}</td><td style={s.tdR}>{r.loadOv||"—"}</td>
                  <td style={s.tdR}>{r.startWk?formatWeek(r.startWk):"—"}</td>
                  <td style={s.td}><button style={s.editBtn} onClick={()=>{setAOppId(r.oppId);setALineId(r.lineId);setAStage(r.stage);setALoadOv(r.loadOv);setAStartWk(formatWeek(r.startWk));setAEditingRow(r);setAPendingDelete(false);}}>Edit</button></td>
                </tr>
              ))}
            </tbody></table></div>
          </>)}

          <div style={{...s.sectionLbl,marginTop:20}}>Capacity override</div>
          <div style={s.editCard}>
            <div style={s.filterRow}>
              <label style={s.fieldLbl}>Plant<select style={s.sel} value={coPlant} onChange={e=>setCoPlant(e.target.value)}>{plants.map(p=><option key={p}>{p}</option>)}</select></label>
              <label style={s.fieldLbl}>Stage<select style={s.sel} value={coStage} onChange={e=>setCoStage(e.target.value)}>{STAGES.map(st=><option key={st}>{st}</option>)}</select></label>
              <label style={s.fieldLbl}>Week<input style={s.inp} value={coWeek} onChange={e=>setCoWeek(e.target.value)} placeholder="e.g. 2026-14"/></label>
              <label style={s.fieldLbl}>Capacity<input style={s.inp} type="number" value={coCap} onChange={e=>setCoCap(e.target.value)}/></label>
            </div>
            <div style={{display:"flex",gap:8,alignItems:"center",flexWrap:"wrap" as const}}>
              <button style={s.saveBtn} onClick={saveCapOv} disabled={coSaving}>{coSaving?"Saving…":"Save override"}</button>
              {coEditingRow&&!coPendingDelete&&(<button style={s.deleteBtnSoft} onClick={()=>setCoPendingDelete(true)}>Delete</button>)}
              {coEditingRow&&coPendingDelete&&(<><button style={s.deleteBtnHard} onClick={deleteCapOv} disabled={coSaving}>Confirm delete</button><button style={s.cancelBtn} onClick={()=>setCoPendingDelete(false)}>Cancel</button></>)}
              {coEditingRow&&<button style={s.cancelBtn} onClick={()=>{setCoEditingRow(null);setCoPendingDelete(false);setCoWeek("");setCoCap("");}}>Clear</button>}
            </div>
          </div>
          {capOvRows.length>0&&(<>
            <div style={{...s.sectionLbl,marginTop:14}}>Existing overrides ({capOvRows.length})</div>
            <div style={s.tableScroll}><table style={s.table}><thead><tr>
              <th style={s.th}>Plant</th><th style={s.th}>Stage</th><th style={s.thR}>Week</th><th style={s.thR}>Cap</th><th style={s.th}></th>
            </tr></thead><tbody>
              {capOvRows.map((r,i)=>(
                <tr key={i} style={coEditingRow?.rowIdx===r.rowIdx?s.trSelected:i%2===0?s.trEven:s.trOdd}>
                  <td style={s.td}>{r.plant}</td><td style={s.td}>{r.stage}</td>
                  <td style={s.tdR}>{r.week?formatWeek(r.week):"—"}</td><td style={s.tdR}>{r.cap}</td>
                  <td style={s.td}><button style={s.editBtn} onClick={()=>{setCoPlant(r.plant);setCoStage(r.stage);setCoWeek(formatWeek(r.week));setCoCap(r.cap);setCoEditingRow(r);setCoPendingDelete(false);}}>Edit</button></td>
                </tr>
              ))}
            </tbody></table></div>
          </>)}
        </div>
      )}
    </div>
  );
}

const styles = {
  shell:{fontFamily:"'Segoe UI',system-ui,sans-serif",fontSize:16,color:"#1e293b",background:"#f8fafc",minHeight:"100vh",display:"flex" as const,flexDirection:"column" as const},
  header:{background:"#0f2942",color:"#fff",padding:"10px 14px",display:"flex" as const,alignItems:"center" as const,justifyContent:"space-between" as const,flexShrink:0},
  headerTitle:{fontSize:22,fontWeight:700,letterSpacing:"-0.01em"},
  modeToggle:{display:"flex" as const,gap:4},
  modeBtn:{fontSize:14,padding:"4px 10px",borderRadius:6,border:"1px solid #334d6e",background:"transparent",color:"#94a3b8",cursor:"pointer" as const},
  modeActive:{fontSize:14,padding:"4px 10px",borderRadius:6,border:"1px solid #38bdf8",background:"#0284c7",color:"#fff",cursor:"pointer" as const,fontWeight:600},
  tabs:{display:"flex" as const,background:"#fff",borderBottom:"2px solid #e2e8f0",padding:"0 8px",flexShrink:0,overflowX:"auto" as const},
  tab:{fontSize:14,padding:"9px 12px",border:"none",background:"none",color:"#64748b",cursor:"pointer" as const,borderBottom:"2px solid transparent",marginBottom:-2,whiteSpace:"nowrap" as const},
  tabActive:{fontSize:14,padding:"9px 12px",border:"none",background:"none",color:"#0f2942",cursor:"pointer" as const,borderBottom:"2px solid #0284c7",marginBottom:-2,fontWeight:700,whiteSpace:"nowrap" as const},
  refreshBtn:{marginLeft:"auto",fontSize:18,padding:"6px 10px",border:"none",background:"none",color:"#64748b",cursor:"pointer" as const},
  statusBar:{fontSize:14,padding:"5px 14px",background:"#eff6ff",color:"#1d4ed8",borderBottom:"1px solid #bfdbfe",flexShrink:0},
  errorBar:{fontSize:14,padding:"5px 14px",background:"#fef2f2",color:"#dc2626",borderBottom:"1px solid #fecaca",flexShrink:0},
  actionBar:{fontSize:14,padding:"5px 14px",background:"#f0fdf4",color:"#15803d",borderBottom:"1px solid #bbf7d0",flexShrink:0},
  content:{padding:12,flex:1,overflowY:"auto" as const},
  filterGrid2:{display:"grid" as const,gridTemplateColumns:"1.1fr 1fr",gap:14,marginBottom:12},
  filterRow:{display:"flex" as const,gap:10,flexWrap:"wrap" as const,marginBottom:10},
  fieldLbl:{fontSize:14,color:"#64748b",fontWeight:600,display:"flex" as const,flexDirection:"column" as const,gap:3,flex:1,minWidth:90},
  sectionLbl:{fontSize:13,fontWeight:700,color:"#64748b",textTransform:"uppercase" as const,letterSpacing:"0.05em",marginBottom:6},
  sel:{fontSize:16,padding:"6px 8px",borderRadius:7,border:"1px solid #cbd5e1",background:"#fff",minWidth:90},
  inp:{fontSize:16,padding:"6px 8px",borderRadius:7,border:"1px solid #cbd5e1",background:"#fff",width:"100%",boxSizing:"border-box" as const},
  multiBox:{maxHeight:110,overflowY:"auto" as const,border:"1px solid #e2e8f0",borderRadius:7,padding:"4px 6px",background:"#fff"},
  checkRow:{display:"flex" as const,alignItems:"center" as const,gap:5,fontSize:14,padding:"2px 0",cursor:"pointer" as const},
  dot:{width:10,height:10,borderRadius:2,display:"inline-block" as const,flexShrink:0},
  legend:{display:"flex" as const,gap:10,flexWrap:"wrap" as const,marginBottom:8},
  legendItem:{display:"flex" as const,alignItems:"center" as const,gap:4,fontSize:14,color:"#64748b"},
  heatGrid:{display:"flex" as const,flexWrap:"wrap" as const,gap:4},
  cell:{borderRadius:7,padding:"5px 7px",minWidth:68,cursor:"default" as const,border:"1px solid rgba(0,0,0,0.06)"},
  cellWk:{fontSize:13,fontWeight:600,marginBottom:1},
  cellUtil:{fontSize:16,fontWeight:700},
  cellOver:{fontSize:13,color:"#fef2f2",marginTop:1},
  tableScroll:{overflowX:"auto" as const,overflowY:"auto" as const,maxHeight:360},
  tableScrollFlex:{overflowX:"auto" as const,overflowY:"auto" as const,maxHeight:"60vh"},
  table:{width:"100%",borderCollapse:"collapse" as const,fontSize:14},
  th:{padding:"6px 8px",background:"#0f2942",color:"#fff",fontWeight:600,textAlign:"left" as const,whiteSpace:"nowrap" as const,position:"sticky" as const,top:0,zIndex:1},
  thR:{padding:"6px 8px",background:"#0f2942",color:"#fff",fontWeight:600,textAlign:"right" as const,whiteSpace:"nowrap" as const,position:"sticky" as const,top:0,zIndex:1},
  td:{padding:"5px 8px",borderBottom:"1px solid #f1f5f9",verticalAlign:"middle" as const},
  tdR:{padding:"5px 8px",borderBottom:"1px solid #f1f5f9",textAlign:"right" as const,verticalAlign:"middle" as const},
  tdMono:{padding:"5px 8px",borderBottom:"1px solid #f1f5f9",fontFamily:"monospace",fontSize:14},
  trEven:{background:"#fff"},
  trOdd:{background:"#f8fafc"},
  trSelected:{background:"#eff6ff"},
  badge:{fontSize:13,color:"#fff",padding:"2px 6px",borderRadius:10,display:"inline-block" as const,whiteSpace:"nowrap" as const,maxWidth:100,overflow:"hidden" as const,textOverflow:"ellipsis" as const},
  editCard:{background:"#fff",border:"1px solid #e2e8f0",borderRadius:10,padding:12,marginBottom:10},
  editTitle:{fontSize:16,fontWeight:700,color:"#0f2942",marginBottom:10},
  changedNote:{fontSize:14,color:"#854d0e",background:"#fffbeb",border:"1px solid #fde68a",borderRadius:7,padding:"6px 10px",marginBottom:10,lineHeight:1.5},
  contextNote:{fontSize:14,color:"#475569",background:"#f1f5f9",border:"1px solid #e2e8f0",borderRadius:7,padding:"6px 10px",marginBottom:10},
  saveBtn:{fontSize:14,padding:"8px 16px",borderRadius:8,border:"none",background:"#0f2942",color:"#fff",cursor:"pointer" as const,fontWeight:600},
  cancelBtn:{fontSize:14,padding:"8px 16px",borderRadius:8,border:"1px solid #cbd5e1",background:"#fff",color:"#64748b",cursor:"pointer" as const},
  editBtn:{fontSize:13,padding:"4px 10px",borderRadius:6,border:"1px solid #0284c7",background:"#eff6ff",color:"#0284c7",cursor:"pointer" as const,fontWeight:600},
  ganttBtn:{fontSize:13,padding:"4px 8px",borderRadius:6,border:"1px solid #e2e8f0",background:"#f8fafc",cursor:"pointer" as const},
  deleteBtnSoft:{fontSize:13,padding:"8px 14px",borderRadius:8,border:"1px solid #fca5a5",background:"#fff",color:"#dc2626",cursor:"pointer" as const,fontWeight:600},
  deleteBtnHard:{fontSize:13,padding:"8px 14px",borderRadius:8,border:"none",background:"#dc2626",color:"#fff",cursor:"pointer" as const,fontWeight:700},
  empty:{color:"#94a3b8",fontSize:16,textAlign:"center" as const,padding:"32px 0",lineHeight:1.8},
  shotBtn:{fontSize:14,padding:"6px 12px",borderRadius:7,border:"1px solid #cbd5e1",background:"#fff",color:"#374151",cursor:"pointer" as const,fontWeight:600},
  coInlineCard:{background:"#fffbeb",border:"1px solid #fde68a",borderRadius:10,padding:10,marginBottom:14},
  saveBtnInline:{fontSize:14,padding:"7px 14px",borderRadius:7,border:"none",background:"#854d0e",color:"#fff",cursor:"pointer" as const,fontWeight:600,alignSelf:"flex-end" as const,height:36},
  stageName:{fontSize:14,fontWeight:600,color:"#0f2942",width:70,flexShrink:0},
  stageStd:{fontSize:13,color:"#94a3b8",width:64,flexShrink:0},
  inpSm:{fontSize:14,padding:"5px 7px",borderRadius:6,border:"1px solid #cbd5e1",background:"#fff",flex:1,minWidth:0},
  stageDeleteSoft:{fontSize:12,padding:"2px 7px",borderRadius:5,border:"1px solid #fca5a5",background:"#fff",color:"#dc2626",cursor:"pointer" as const,fontWeight:600,flexShrink:0},
  addStageBtn:{fontSize:14,padding:"7px 14px",borderRadius:7,border:"1px solid #0284c7",background:"#eff6ff",color:"#0284c7",cursor:"pointer" as const,fontWeight:600,whiteSpace:"nowrap" as const},
  versionTag:{fontSize:13,fontWeight:400,color:"#94a3b8",marginLeft:6},
  alignFlagRow:{background:"#f0f9ff",border:"1px solid #bae6fd",borderRadius:8,padding:"10px 12px",marginBottom:10,marginTop:4},
  alignFlagLabel:{display:"flex" as const,alignItems:"flex-start" as const,cursor:"pointer" as const,fontSize:15},
  alignFlagText:{fontWeight:700,color:"#0284c7",whiteSpace:"nowrap" as const},
  alignFlagHint:{fontSize:13,color:"#64748b",marginLeft:6,lineHeight:1.5},
  alignBtnOn:{fontSize:13,padding:"3px 8px",borderRadius:6,border:"1px solid #0284c7",background:"#0284c7",color:"#fff",cursor:"pointer" as const,fontWeight:700,whiteSpace:"nowrap" as const},
  alignBtnOff:{fontSize:13,padding:"3px 8px",borderRadius:6,border:"1px solid #cbd5e1",background:"#f8fafc",color:"#94a3b8",cursor:"pointer" as const,whiteSpace:"nowrap" as const},
  planSpecBtnOn:{fontSize:13,padding:"3px 8px",borderRadius:6,border:"1px solid #9d174d",background:"#9d174d",color:"#fff",cursor:"pointer" as const,fontWeight:700,whiteSpace:"nowrap" as const},
  planSpecBtnOff:{fontSize:13,padding:"3px 8px",borderRadius:6,border:"1px solid #cbd5e1",background:"#f8fafc",color:"#94a3b8",cursor:"pointer" as const,whiteSpace:"nowrap" as const},
  alignIndicator:{color:"#0284c7",fontWeight:700,fontSize:14},
  alignLegend:{fontSize:13,color:"#64748b",marginLeft:6,fontWeight:400,textTransform:"none" as const,letterSpacing:0},
  alignDot:{color:"#0284c7",fontWeight:700},
  oiRow:{display:"flex" as const,gap:16,marginBottom:10,flexWrap:"wrap" as const},
  oiBlock:{display:"flex" as const,flexDirection:"column" as const,gap:4,flex:1,minWidth:140},
  oiValue:{fontSize:16,fontWeight:600,color:"#1e293b",padding:"6px 0"},
  oiHint:{fontSize:12,fontWeight:400,color:"#94a3b8",textTransform:"none" as const,letterSpacing:0},
  oiSpreadLabel:{display:"flex" as const,alignItems:"center" as const,fontSize:13,color:"#64748b",marginTop:5,cursor:"pointer" as const},
};