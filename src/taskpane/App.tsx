/// <reference types="office-js" />
/// <reference types="office-js" />
import * as React from "react";
import { useState, useEffect, useCallback, useRef } from "react";
import {
  AreaChart, Area, XAxis, YAxis, CartesianGrid, Tooltip,
  Legend, ResponsiveContainer, Line,
} from "recharts";

/* ─── constants ─────────────────────────────────────────────────────────────── */
// DIFF: bumped to v0.3.0 — OIDatePlanned, context columns, PlanSpeculative, delete actions
const APP_VERSION = "v0.3.0";
const STAGES = ["ESL", "CTO", "Assembly", "Testing", "FAT"];
const STATUS_LIST = [
  "Backlog","FROM OI TO KO","ANTICIPATION - ENGINEERING ONLY + PROCUREMENT",
  "ANTICIPATION - ENGINEERING ONLY","ON HAND","FORECAST","BACKUP","OUTLOOK","OTHER",
];
const STATUS_STACK_ORDER = [
  "Backlog","FROM OI TO KO","ANTICIPATION - ENGINEERING ONLY + PROCUREMENT",
  "ANTICIPATION - ENGINEERING ONLY","ON HAND","FORECAST","BACKUP","OUTLOOK","OTHER",
];
const STATUS_SHORT: Record<string, string> = {
  "Backlog":"Backlog","FROM OI TO KO":"OI→KO",
  "ANTICIPATION - ENGINEERING ONLY + PROCUREMENT":"Antic+Proc",
  "ANTICIPATION - ENGINEERING ONLY":"Antic Eng",
  "ON HAND":"On Hand","FORECAST":"Forecast","BACKUP":"Backup","OUTLOOK":"Outlook","OTHER":"Other",
};
const STATUS_COLORS: Record<string, string> = {
  "Backlog":"#1e293b","FROM OI TO KO":"#0f4c81",
  "ANTICIPATION - ENGINEERING ONLY + PROCUREMENT":"#1a6b3c",
  "ANTICIPATION - ENGINEERING ONLY":"#2d9c5f",
  "ON HAND":"#166534","FORECAST":"#854d0e","BACKUP":"#6b21a8","OUTLOOK":"#374151","OTHER":"#9ca3af",
};
const CHANGE_COLORS: Record<string, string> = {
  "New":"#16a34a","Changed":"#ca8a04","Removed":"#dc2626","Unchanged":"#94a3b8",
};
// DIFF: speculative statuses that require PlanSpeculative flag to be planned
const SPECULATIVE_STATUSES = new Set(["BACKUP","OUTLOOK","OTHER"]);

type Tab = "graph"|"heatmap"|"demand"|"adjustments";
type Mode = "constrained"|"unconstrained";

/* ─── types ─────────────────────────────────────────────────────────────────── */
type ResultRow = { plant:string;stage:string;week:string;status:string;backlog:number;planned:number;total:number;cap:number;util:number;overload:number; };
// DIFF: oiDatePlanned + live oiDate both present; schedule col indices shifted +1
type SchedRow = { oppId:string;lineId:string;equip:string;plant:string;oiDatePlanned:number;oiDate:number;status:string;kom:number;fat:number;fca:number;lt:number; };
// DIFF: added oiDatePlanned, customer, region, subRegion, country, planSpeculative, alignFlag
type DemandRow = {
  rowIdx:number;oppId:string;lineId:string;equip:string;routingKey:string;plant:string;
  oiDate:number;oiDatePlanned:number;status:string;priority:string;
  changeFlag:string;changedFields:string;alignFlag:string;
  customer:string;region:string;subRegion:string;country:string;planSpeculative:string;
};
type AdjRow = { rowIdx:number;oppId:string;lineId:string;stage:string;loadOv:string;startWk:string; };
type CapOvRow = { rowIdx:number;plant:string;stage:string;week:string;cap:string; };

/* ─── helpers ───────────────────────────────────────────────────────────────── */
const excelDateToStr = (s:number) => s ? new Date((s-25569)*86400000).toISOString().slice(0,10) : "";
// DIFF: parse "YYYY-MM-DD" string back to Excel serial for writing OIDatePlanned
const strToExcelDate = (s:string): number => {
  if (!s) return 0;
  const d = new Date(s + "T00:00:00Z");
  return isNaN(d.getTime()) ? 0 : Math.round(d.getTime()/86400000 + 25569);
};

function isoWeekLabel(serial: number): string {
  const d = new Date((serial - 25569) * 86400000);
  const day = (d.getUTCDay() + 6) % 7;
  d.setUTCDate(d.getUTCDate() - day + 3);
  const isoYear = d.getUTCFullYear();
  const jan1 = new Date(Date.UTC(isoYear, 0, 1));
  const week = Math.floor((d.getTime() - jan1.getTime()) / 86400000 / 7) + 1;
  return `${isoYear}-${String(week).padStart(2, "0")}`;
}
const utilColor = (u:number,ov:number) => ov>0?"#dc2626":u>=90?"#ea580c":u>=70?"#ca8a04":u>0?"#16a34a":"#e2e8f0";
const utilText = (u:number,ov:number) => (ov>0||u>=90)?"#fff":u>0?"#14532d":"#94a3b8";

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
    body.getCell(rowIdx,colIdx).values = [[value]];
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
    body.getRow(rowIdx).values = [values];
    await ctx.sync();
  });
}
// DIFF: delete a row from a table by its 0-based data row index
async function deleteRowFromTable(tableName:string, rowIdx:number):Promise<void> {
  await Excel.run(async ctx => {
    const tbl = ctx.workbook.tables.getItem(tableName);
    tbl.rows.getItemAt(rowIdx).delete();
    await ctx.sync();
  });
}

async function downloadChartAsPng(containerEl: HTMLElement, filename: string): Promise<void> {
  const svg = containerEl.querySelector("svg");
  if (!svg) throw new Error("No chart found to capture — wait for the chart to render first.");
  const bbox = svg.getBoundingClientRect();
  if (bbox.width === 0 || bbox.height === 0) throw new Error("Chart has zero size.");
  const svgClone = svg.cloneNode(true) as SVGSVGElement;
  svgClone.setAttribute("xmlns","http://www.w3.org/2000/svg");
  svgClone.setAttribute("xmlns:xlink","http://www.w3.org/1999/xlink");
  svgClone.setAttribute("width",String(bbox.width));
  svgClone.setAttribute("height",String(bbox.height));
  svgClone.setAttribute("viewBox",`0 0 ${bbox.width} ${bbox.height}`);
  svgClone.style.fontFamily="Arial, sans-serif";
  const bgRect = document.createElementNS("http://www.w3.org/2000/svg","rect");
  bgRect.setAttribute("width","100%"); bgRect.setAttribute("height","100%"); bgRect.setAttribute("fill","#ffffff");
  svgClone.insertBefore(bgRect,svgClone.firstChild);
  const svgData = new XMLSerializer().serializeToString(svgClone);
  const svg64 = btoa(unescape(encodeURIComponent(svgData)));
  const dataUrl = `data:image/svg+xml;base64,${svg64}`;
  return new Promise((resolve,reject) => {
    const img = new Image();
    img.onload = () => {
      try {
        const scale=2;
        const canvas=document.createElement("canvas");
        canvas.width=bbox.width*scale; canvas.height=bbox.height*scale;
        const c=canvas.getContext("2d");
        if(!c){reject(new Error("Canvas not supported.")); return;}
        c.scale(scale,scale); c.drawImage(img,0,0,bbox.width,bbox.height);
        let pngUrl:string;
        try { pngUrl=canvas.toDataURL("image/png"); }
        catch(err){reject(new Error(`Canvas export blocked (${String(err)}).`)); return;}
        const link=document.createElement("a");
        link.href=pngUrl; link.download=filename;
        document.body.appendChild(link); link.click(); document.body.removeChild(link);
        resolve();
      } catch(err){reject(err instanceof Error?err:new Error(String(err)));}
    };
    img.onerror=()=>reject(new Error("Could not render chart as image."));
    img.src=dataUrl;
  });
}

/* ─── main ──────────────────────────────────────────────────────────────────── */
export default function App() {
  const [tab,setTab] = useState<Tab>("graph");
  const [mode,setMode] = useState<Mode>("constrained");
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
  // demand
  const [dFilter,setDFilter] = useState<"unassigned"|"all">("unassigned");
  const [dChangeFilter,setDChangeFilter] = useState("");
  const [dStatusFilter,setDStatusFilter] = useState("");
  const [dPlantFilter,setDPlantFilter] = useState("");
  const [dSearch,setDSearch] = useState("");
  // DIFF: new context column filters
  const [dRegionFilter,setDRegionFilter] = useState("");
  const [dCustomerFilter,setDCustomerFilter] = useState("");
  const [dEditing,setDEditing] = useState<DemandRow|null>(null);
  const [dRk,setDRk] = useState("");
  const [dPlant,setDPlant] = useState("");
  const [dPriority,setDPriority] = useState("");
  const [dAlignFlag,setDAlignFlag] = useState(false);
  // DIFF: PlanSpeculative edit state
  const [dPlanSpec,setDPlanSpec] = useState(false);
  // DIFF: OIDatePlanned edit state (string for date input)
  const [dOiPlanned,setDOiPlanned] = useState("");
  const [dSaving,setDSaving] = useState(false);
  const [dStageOv,setDStageOv] = useState<Record<string,{loadOv:string;startWk:string}>>({});
  // adjustments
  const [aOppId,setAOppId] = useState("");
  const [aLineId,setALineId] = useState("");
  const [aStage,setAStage] = useState("Assembly");
  const [aLoadOv,setALoadOv] = useState("");
  const [aStartWk,setAStartWk] = useState("");
  const [aSaving,setASaving] = useState(false);
  // DIFF: delete confirmation states for adj and cap override
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

  const resolveWeekInput = (raw: string): string | number => {
    if (raw === "") return "";
    const t = raw.trim();
    if (/^\d+$/.test(t)) return Number(t);
    const idx = weekLabelToIndex[t];
    return idx !== undefined ? idx : "";
  };
  const formatWeek = (raw: string): string => {
    if (raw === "") return "";
    return indexToWeekLabel[raw.trim()] ?? raw;
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
      // DIFF: Schedule col indices shifted +1 after live OIDate insertion at col 5
      // New layout: 0=OppID,1=LineID,2=Equip,3=Plant,4=OIDatePlanned,5=OIDate,6=Status,
      //             7=KOM,...,17=FAT,19=FCA,20=LT
      setSchedule(sData.rows.map(r=>({
        oppId:String(r[0]),lineId:String(r[1]),equip:String(r[2]),plant:String(r[3]),
        oiDatePlanned:Number(r[4]),oiDate:Number(r[5]),status:String(r[6]),
        kom:Number(r[7]),fat:Number(r[17]),fca:Number(r[19]),lt:Number(r[20]),
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
        // DIFF: read OIDatePlanned; blank-safe
        oiDatePlanned:Number(r[dh.indexOf("OIDatePlanned")]??0),
        status:String(r[dh.indexOf("Status")]??""),
        priority:String(r[dh.indexOf("Priority")]??""),
        changeFlag:String(r[dh.indexOf("ChangeFlag")]??""),
        changedFields:String(r[dh.indexOf("ChangedFields")]??""),
        alignFlag:String(r[dh.indexOf("AlignFlag")]??""),
        // DIFF: context columns
        customer:String(r[dh.indexOf("Customer")]??""),
        region:String(r[dh.indexOf("Region")]??""),
        subRegion:String(r[dh.indexOf("SubRegion")]??""),
        country:String(r[dh.indexOf("Country")]??""),
        // DIFF: PlanSpeculative flag
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
      const STAGE_COL: Record<string,string> = { ESL:"ESL Load",CTO:"CTO Load",Assembly:"Assembly Load",Testing:"Testing Load",FAT:"FAT Load" };
      const rl: Record<string,Record<string,number>> = {};
      for (const r of rtData.rows) {
        const key=String(r[rkIdx]??""); if(!key) continue;
        rl[key]={};
        for (const st of STAGES) {
          const ci=rtData.headers.indexOf(STAGE_COL[st]);
          rl[key][st]=ci>=0?Number(r[ci]??0):0;
        }
      }
      setRoutingLoads(rl);
      const ps=new Set<string>();
      cbData.rows.forEach(r=>{const p=String(r[cbData.headers.indexOf("Plant")]??"");if(p)ps.add(p);});
      const pl=Array.from(ps).sort(); setPlants(pl);
      if(!hPlant&&pl.length){setHPlant(pl[0]);}
      if(!gPlants.length&&pl.length){setGPlants([pl[0]]);}
      if(!coPlant&&pl.length){setCoPlant(pl[0]);}
      const wIdx=calData.headers.indexOf("WeekIndex"),wStart=calData.headers.indexOf("WeekStart");
      const l2i: Record<string,number>={};
      const i2l: Record<string,string>={};
      if(wIdx>=0&&wStart>=0){
        for(const r of calData.rows){
          const idx=Number(r[wIdx]),start=Number(r[wStart]);
          if(idx>=1&&start){const label=isoWeekLabel(start);l2i[label]=idx;i2l[String(idx)]=label;}
        }
      }
      setWeekLabelToIndex(l2i); setIndexToWeekLabel(i2l);
      setLoadMsg(`Loaded ${rData.rows.length} capacity rows · ${sData.rows.length} schedule lines`);
    } catch(e:unknown){setErrMsg(String(e));setLoadMsg("");}
    setLoading(false);
  },[mode]);

  useEffect(()=>{load();},[load]);

  /* graph data */
  const graphData = (()=>{
    const f=results.filter(r=>
      (gPlants.length===0||gPlants.includes(r.plant))&&r.stage===gStage&&
      (gStatuses.length===0||gStatuses.includes(r.status))&&
      (!gFrom||r.week>=gFrom)&&(!gTo||r.week<=gTo)
    );
    const wm:Record<string,Record<string,number>>={};
    for(const r of f){if(!wm[r.week])wm[r.week]={};wm[r.week][r.status]=(wm[r.week][r.status]??0)+r.total;}
    const capByWeek:Record<string,number>={};
    for(const r of f){if(!capByWeek[r.week])capByWeek[r.week]=0;}
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

  /* heatmap */
  const heatRows=results.filter(r=>r.plant===hPlant&&r.stage===hStage);
  const weekCapMap:Record<string,number>={};
  for(const r of heatRows)weekCapMap[r.week]=r.cap;
  const hwm:Record<string,{total:number;overload:number}>={};
  for(const r of heatRows){if(!hwm[r.week])hwm[r.week]={total:0,overload:0};hwm[r.week].total+=r.total;hwm[r.week].overload+=r.overload;}
  const heatWeeks=Object.keys(hwm).sort();

  /* demand + schedule join */
  const schedByLine:Record<string,SchedRow>={};
  for(const s of schedule)schedByLine[s.lineId]=s;

  const alignedOppIds=new Set<string>();
  for(const d of demand){const f=d.alignFlag;if(f!==""&&f!=="0")alignedOppIds.add(d.oppId);}

  // DIFF: derive unique regions and customers for filter dropdowns
  const allRegions=Array.from(new Set(demand.map(d=>d.region).filter(Boolean))).sort();
  const allCustomers=Array.from(new Set(demand.map(d=>d.customer).filter(Boolean))).sort();

  const demandFiltered=demand.filter(d=>{
    if(!dChangeFilter&&d.changeFlag==="Removed")return false;
    if(dChangeFilter&&d.changeFlag!==dChangeFilter)return false;
    if(dFilter==="unassigned"&&d.routingKey&&d.plant)return false;
    if(dPlantFilter&&d.plant!==dPlantFilter)return false;
    if(dStatusFilter&&d.status!==dStatusFilter)return false;
    if(dSearch&&!d.oppId.toLowerCase().includes(dSearch.toLowerCase())&&!d.equip.toLowerCase().includes(dSearch.toLowerCase()))return false;
    // DIFF: context column filters
    if(dRegionFilter&&d.region!==dRegionFilter)return false;
    if(dCustomerFilter&&d.customer!==dCustomerFilter)return false;
    return true;
  }).map(d=>({...d,sched:schedByLine[d.lineId]}));

  const toggleAlignFlag=async(d:DemandRow)=>{
    const newVal=(d.alignFlag!==""&&d.alignFlag!=="0")?"":"1";
    try{
      await writeCellInTable("Demand",d.rowIdx,"AlignFlag",newVal,demandHeaders);
      setActionNote(newVal==="1"?`Deal ${d.oppId} flagged for alignment.`:`Alignment flag cleared for ${d.oppId}/${d.lineId}.`);
      await load();
    }catch(e){setActionNote(`Error: ${String(e)}`);}
  };

  // DIFF: toggle PlanSpeculative directly from table row
  const togglePlanSpec=async(d:DemandRow)=>{
    const newVal=(d.planSpeculative!==""&&d.planSpeculative!=="0")?"":"1";
    try{
      await writeCellInTable("Demand",d.rowIdx,"PlanSpeculative",newVal,demandHeaders);
      setActionNote(newVal==="1"?`Deal ${d.oppId} opted in to speculative planning.`:`Speculative flag cleared for ${d.oppId}/${d.lineId}.`);
      await load();
    }catch(e){setActionNote(`Error: ${String(e)}`);}
  };

  const saveDemandLine=async()=>{
    if(!dEditing)return;
    for(const st of STAGES){
      const ov=dStageOv[st];
      if(!ov||ov.startWk==="")continue;
      if(resolveWeekInput(ov.startWk)===""){setActionNote(`Could not resolve "${ov.startWk}" (${st} start week).`);return;}
    }
    setDSaving(true); setActionNote("");
    try{
      const h=demandHeaders;
      if(dRk)await writeCellInTable("Demand",dEditing.rowIdx,"RoutingKey",dRk,h);
      if(dPlant)await writeCellInTable("Demand",dEditing.rowIdx,"Plant",dPlant,h);
      await writeCellInTable("Demand",dEditing.rowIdx,"Priority",dPriority===""?"":Number(dPriority),h);
      await writeCellInTable("Demand",dEditing.rowIdx,"AlignFlag",dAlignFlag?"1":"",h);
      await writeCellInTable("Demand",dEditing.rowIdx,"PlanSpeculative",dPlanSpec?"1":"",h);
      // DIFF: persist OIDatePlanned — write serial or blank to inherit from OIDate
      if(dOiPlanned!==""){
        const serial=strToExcelDate(dOiPlanned);
        if(serial>0)await writeCellInTable("Demand",dEditing.rowIdx,"OIDatePlanned",serial,h);
      } else {
        // blank → clears the override, Sync will refill from OIDate on next run
        await writeCellInTable("Demand",dEditing.rowIdx,"OIDatePlanned","",h);
      }
      for(const st of STAGES){
        const ov=dStageOv[st];
        if(!ov||(ov.loadOv===""&&ov.startWk===""))continue;
        const ex=adjRows.find(r=>r.oppId===dEditing.oppId&&r.lineId===dEditing.lineId&&r.stage===st);
        const row=[dEditing.oppId,dEditing.lineId,st,ov.loadOv===""?"":Number(ov.loadOv),resolveWeekInput(ov.startWk)] as (string|number)[];
        if(ex)await updateRowInTable("Adjustments",ex.rowIdx,row);
        else await appendRowToTable("Adjustments",row);
      }
      setActionNote(`Saved ${dEditing.oppId}`); setDEditing(null); setDStageOv({}); await load();
    }catch(e){setActionNote(`Error: ${String(e)}`);}
    setDSaving(false);
  };

  const saveAdj=async()=>{
    if(!aOppId||!aLineId||!aStage){setActionNote("Fill OppID, LineID and Stage.");return;}
    const wk=resolveWeekInput(aStartWk);
    if(aStartWk!==""&&wk===""){setActionNote(`Could not resolve "${aStartWk}" to a week.`);return;}
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

  // DIFF: delete adjustment row
  const deleteAdj=async()=>{
    if(!aEditingRow)return;
    setASaving(true); setActionNote("");
    try{
      await deleteRowFromTable("Adjustments",aEditingRow.rowIdx);
      setActionNote(`Deleted adjustment ${aEditingRow.oppId}/${aEditingRow.stage}.`);
      setAOppId(""); setALineId(""); setALoadOv(""); setAStartWk("");
      setAEditingRow(null); setAPendingDelete(false); await load();
    }catch(e){setActionNote(`Error: ${String(e)}`);}
    setASaving(false);
  };

  const saveCapOv=async()=>{
    if(!coPlant||!coStage||!coWeek||!coCap){setActionNote("Fill all override fields.");return;}
    const wk=resolveWeekInput(coWeek);
    if(wk===""){setActionNote(`Could not resolve "${coWeek}" to a week.`);return;}
    setCoSaving(true); setActionNote("");
    try{
      const ex=coEditingRow??capOvRows.find(r=>r.plant===coPlant&&r.stage===coStage&&r.week===String(wk));
      const row=[coPlant,coStage,wk,Number(coCap)] as (string|number)[];
      if(ex){await updateRowInTable("CapacityOverride",ex.rowIdx,row);setActionNote(`Updated override ${coPlant}/${coStage}/W${wk}`);}
      else{await appendRowToTable("CapacityOverride",row);setActionNote(`Added override ${coPlant}/${coStage}/W${wk}`);}
      setCoEditingRow(null); setCoPendingDelete(false); await load();
    }catch(e){setActionNote(`Error: ${String(e)}`);}
    setCoSaving(false);
  };

  // DIFF: delete capacity override row
  const deleteCapOv=async()=>{
    if(!coEditingRow)return;
    setCoSaving(true); setActionNote("");
    try{
      await deleteRowFromTable("CapacityOverride",coEditingRow.rowIdx);
      setActionNote(`Deleted override ${coEditingRow.plant}/${coEditingRow.stage}/W${coEditingRow.week}.`);
      setCoPlant(plants[0]??""); setCoStage("Assembly"); setCoWeek(""); setCoCap("");
      setCoEditingRow(null); setCoPendingDelete(false); await load();
    }catch(e){setActionNote(`Error: ${String(e)}`);}
    setCoSaving(false);
  };

  const s=styles;
  return (
    <div style={s.shell}>
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

      {/* GRAPH */}
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
                if(!chartRef.current)return;
                setShotBusy(true); setActionNote("");
                try{await downloadChartAsPng(chartRef.current,`sop-load-${gStage}-${new Date().toISOString().slice(0,10)}.png`);setActionNote("Chart saved as PNG.");}
                catch(e){setActionNote(`Screenshot error: ${String(e)}`);}
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
                  {STATUS_STACK_ORDER.filter(st=>activeStatuses.includes(st)&&graphData.some(d=>(d as unknown as Record<string,number>)[st]!=null)).map(st=>(
                    <Area key={st} type="monotone" dataKey={st} name={STATUS_SHORT[st]} stackId="1"
                      stroke={STATUS_COLORS[st]} fill={STATUS_COLORS[st]} fillOpacity={0.75} strokeWidth={1} dot={false} connectNulls={true}/>
                  ))}
                  <Line type="monotone" dataKey="_cap" name="Capacity" stroke="#dc2626" strokeWidth={2} dot={false} connectNulls={true} legendType="line" isAnimationActive={false}/>
                </AreaChart>
              </ResponsiveContainer>
            </div>
          </>)}
          <div style={s.coInlineCard}>
            <div style={s.sectionLbl}>Quick capacity override (writes immediately — re-run script to see impact)</div>
            <div style={s.filterRow}>
              <label style={s.fieldLbl}>Plant<select style={s.sel} value={coPlant} onChange={e=>setCoPlant(e.target.value)}>{plants.map(p=><option key={p}>{p}</option>)}</select></label>
              <label style={s.fieldLbl}>Stage<select style={s.sel} value={coStage} onChange={e=>setCoStage(e.target.value)}>{STAGES.map(st=><option key={st}>{st}</option>)}</select></label>
              <label style={s.fieldLbl}>Week<input style={s.inp} value={coWeek} onChange={e=>setCoWeek(e.target.value)} placeholder="e.g. 2026-14"/></label>
              <label style={s.fieldLbl}>New capacity<input style={s.inp} type="number" value={coCap} onChange={e=>setCoCap(e.target.value)}/></label>
              <button style={s.saveBtnInline} onClick={saveCapOv} disabled={coSaving}>{coSaving?"Saving…":"Save"}</button>
            </div>
          </div>
        </div>
      )}

      {/* HEATMAP */}
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
              const w=hwm[wk];const cap=weekCapMap[wk]??0;
              const util=cap?Math.round(w.total/cap*100):0;
              const bg=utilColor(util,w.overload);const fg=utilText(util,w.overload);
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

      {/* DEMAND */}
      {tab==="demand"&&(
        <div style={s.content}>
          <div style={s.sectionLbl}>Run scripts</div>
          <div style={s.runRow}>
            {[{label:"Sync demand",sheet:"Demand"},{label:"Rebalance (constrained)",sheet:"Results"},{label:"Rebalance (unconstrained)",sheet:"ResultsUnconstrained"}].map(({label,sheet})=>(
              <button key={label} style={s.runBtn} onClick={async()=>{
                try{await Excel.run(async ctx=>{ctx.workbook.worksheets.getItem(sheet).activate();await ctx.sync();});setActionNote(`Switched to "${sheet}" — run "${label}" from the Automate tab, then refresh.`);}
                catch{setActionNote(`Could not find sheet "${sheet}".`);}
              }}>{label} →</button>
            ))}
          </div>
          <div style={s.filterRow}>
            <label style={s.fieldLbl}>Show
              <select style={s.sel} value={dFilter} onChange={e=>setDFilter(e.target.value as "unassigned"|"all")}>
                <option value="unassigned">Unassigned</option><option value="all">All lines</option>
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
            {/* DIFF: Region and Customer filters */}
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
              {/* DIFF: show customer/region context in edit card */}
              {(dEditing.customer||dEditing.region)&&(
                <div style={s.contextNote}>
                  {[dEditing.customer,dEditing.region,dEditing.subRegion,dEditing.country].filter(Boolean).join(" · ")}
                </div>
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
                <label style={s.fieldLbl}>Priority (optional)
                  <input style={s.inp} type="number" placeholder="blank = auto" value={dPriority} onChange={e=>setDPriority(e.target.value)}/>
                </label>
              </div>

              {/* DIFF: OIDatePlanned edit row — shows divergence when it differs from live OI */}
              <div style={s.oiRow}>
                <div style={s.oiBlock}>
                  <div style={s.sectionLbl}>OI (live)</div>
                  <div style={s.oiValue}>{excelDateToStr(dEditing.oiDate)}</div>
                </div>
                <div style={s.oiBlock}>
                  <div style={s.sectionLbl}>OI Planned <span style={s.oiHint}>(edit to override; blank = inherit live OI)</span></div>
                  <input style={{
                    ...s.inp,
                    borderColor: dOiPlanned && dOiPlanned !== excelDateToStr(dEditing.oiDate) ? "#ca8a04" : "#cbd5e1",
                    background: dOiPlanned && dOiPlanned !== excelDateToStr(dEditing.oiDate) ? "#fffbeb" : "#fff",
                  }} type="date" value={dOiPlanned} onChange={e=>setDOiPlanned(e.target.value)}/>
                </div>
              </div>

              {/* DIFF: AlignFlag checkbox */}
              <div style={s.alignFlagRow}>
                <label style={s.alignFlagLabel}>
                  <input type="checkbox" checked={dAlignFlag} onChange={e=>setDAlignFlag(e.target.checked)} style={{marginRight:6}}/>
                  <span style={s.alignFlagText}>Align deal LT</span>
                  <span style={s.alignFlagHint}> — flags all equipment in deal {dEditing.oppId} to finish within the same window as the slowest line</span>
                </label>
              </div>

              {/* DIFF: PlanSpeculative checkbox — only shown for speculative statuses */}
              {SPECULATIVE_STATUSES.has(dEditing.status.toUpperCase())&&(
                <div style={{...s.alignFlagRow,borderColor:"#f9a8d4",background:"#fdf2f8"}}>
                  <label style={s.alignFlagLabel}>
                    <input type="checkbox" checked={dPlanSpec} onChange={e=>setDPlanSpec(e.target.checked)} style={{marginRight:6}}/>
                    <span style={{...s.alignFlagText,color:"#9d174d"}}>Include in planning</span>
                    <span style={s.alignFlagHint}> — this deal has speculative status ({STATUS_SHORT[dEditing.status]??dEditing.status}); off by default</span>
                  </label>
                </div>
              )}

              <div style={{...s.sectionLbl,marginTop:12}}>Stage adjustments (blank = use standard)</div>
              {STAGES.map(st=>{
                const std=dRk?(routingLoads[dRk]?.[st]??0):(dEditing.routingKey?(routingLoads[dEditing.routingKey]?.[st]??0):0);
                const ov=dStageOv[st]??{loadOv:"",startWk:""};
                return(
                  <div key={st} style={s.stageRow}>
                    <div style={s.stageName}>{st}</div>
                    <div style={s.stageStd}>std: {std}</div>
                    <input style={s.inpSm} type="number" placeholder={`load (std ${std})`} value={ov.loadOv} onChange={e=>setDStageOv({...dStageOv,[st]:{...ov,loadOv:e.target.value}})}/>
                    <input style={s.inpSm} placeholder="start wk (YYYY-WW)" value={ov.startWk} onChange={e=>setDStageOv({...dStageOv,[st]:{...ov,startWk:e.target.value}})}/>
                  </div>
                );
              })}
              <div style={{display:"flex",gap:8,marginTop:10}}>
                <button style={s.saveBtn} onClick={saveDemandLine} disabled={dSaving}>{dSaving?"Saving…":"Save"}</button>
                <button style={s.cancelBtn} onClick={()=>{setDEditing(null);setDStageOv({});}}>Cancel</button>
              </div>
            </div>
          )}

          <div style={s.sectionLbl}>
            {demandFiltered.length} lines
            {alignedOppIds.size>0&&<span style={s.alignLegend}>· <span style={s.alignDot}>⟳</span> = deal aligned ({alignedOppIds.size} deal{alignedOppIds.size!==1?"s":""})</span>}
          </div>
          <div style={s.tableScrollFlex}>
            <table style={s.table}><thead><tr>
              <th style={s.th}>OppID</th><th style={s.th}>Equipment</th>
              <th style={s.th}>Customer</th><th style={s.th}>Region</th>
              <th style={s.th}>Routing</th><th style={s.th}>Plant</th><th style={s.thR}>Priority</th>
              <th style={s.th}>Change</th><th style={s.th}>Status</th>
              {/* DIFF: OI Planned + live OI side by side */}
              <th style={s.thR}>OI Planned</th><th style={s.thR}>OI Live</th>
              <th style={s.thR}>KOM</th><th style={s.thR}>FAT</th><th style={s.thR}>FCA</th><th style={s.thR}>LT</th>
              <th style={s.th} title="Align deal LT">Align</th>
              <th style={s.th} title="Include speculative deal in planning">Plan?</th>
              <th style={s.th}></th>
            </tr></thead><tbody>
              {demandFiltered.map((d,i)=>{
                const isAligned=d.alignFlag!==""&&d.alignFlag!=="0";
                const dealAligned=alignedOppIds.has(d.oppId);
                const isPlanSpec=d.planSpeculative!==""&&d.planSpeculative!=="0";
                const isSpecStatus=SPECULATIVE_STATUSES.has(d.status.toUpperCase());
                // DIFF: divergence cue — amber when planned OI differs from live OI
                const oiDiverged=d.oiDatePlanned>0&&d.oiDatePlanned!==d.oiDate;
                return(
                  <tr key={i} style={dEditing?.lineId===d.lineId?s.trSelected:i%2===0?s.trEven:s.trOdd}>
                    <td style={s.tdMono}>
                      {dealAligned&&<span style={s.alignIndicator} title="Deal alignment active">⟳ </span>}
                      {d.oppId}
                    </td>
                    <td style={{...s.td,maxWidth:90,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}} title={d.equip}>{d.equip}</td>
                    {/* DIFF: customer and region columns */}
                    <td style={{...s.td,maxWidth:80,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}} title={d.customer}>{d.customer||"—"}</td>
                    <td style={s.td}>{d.region||"—"}</td>
                    <td style={{...s.td,color:d.routingKey?"inherit":"#dc2626"}}>{d.routingKey||"—"}</td>
                    <td style={{...s.td,color:d.plant?"inherit":"#dc2626"}}>{d.plant||"—"}</td>
                    <td style={s.tdR}>{d.priority||""}</td>
                    <td style={s.td}>{d.changeFlag&&<span style={{...s.badge,background:CHANGE_COLORS[d.changeFlag]??"#6b7280"}}>{d.changeFlag}</span>}</td>
                    <td style={s.td}>{d.status?<span style={{...s.badge,background:STATUS_COLORS[d.status]??"#6b7280"}}>{STATUS_SHORT[d.status]??d.status.split(" ")[0]}</span>:<span style={{color:"#94a3b8"}}>—</span>}</td>
                    {/* DIFF: OI Planned with amber divergence cue */}
                    <td style={{...s.tdR,background:oiDiverged?"#fffbeb":"inherit",color:oiDiverged?"#92400e":"inherit",fontWeight:oiDiverged?700:400}} title={oiDiverged?`Live OI: ${excelDateToStr(d.oiDate)}`:"Matches live OI"}>
                      {d.oiDatePlanned?excelDateToStr(d.oiDatePlanned):excelDateToStr(d.oiDate)}
                      {oiDiverged&&" ⚠"}
                    </td>
                    <td style={s.tdR}>{excelDateToStr(d.oiDate)}</td>
                    <td style={s.tdR}>{d.sched?excelDateToStr(d.sched.kom):""}</td>
                    <td style={s.tdR}>{d.sched?excelDateToStr(d.sched.fat):""}</td>
                    <td style={s.tdR}>{d.sched?excelDateToStr(d.sched.fca):""}</td>
                    <td style={{...s.tdR,fontWeight:600}}>{d.sched?.lt??""}</td>
                    <td style={s.td}>
                      <button style={isAligned?s.alignBtnOn:s.alignBtnOff} title={isAligned?"Remove alignment flag":"Flag deal for LT alignment"} onClick={()=>toggleAlignFlag(d)}>
                        {isAligned?"⟳ On":"Off"}
                      </button>
                    </td>
                    {/* DIFF: PlanSpeculative toggle — only meaningful for speculative statuses */}
                    <td style={s.td}>
                      {isSpecStatus?(
                        <button style={isPlanSpec?s.planSpecBtnOn:s.planSpecBtnOff} title={isPlanSpec?"Remove from planning":"Include in planning"} onClick={()=>togglePlanSpec(d)}>
                          {isPlanSpec?"✓ Yes":"No"}
                        </button>
                      ):<span style={{color:"#cbd5e1",fontSize:13}}>—</span>}
                    </td>
                    <td style={s.td}>
                      <button style={s.editBtn} onClick={()=>{
                        setDEditing(d);setDRk(d.routingKey);setDPlant(d.plant);setDPriority(d.priority);
                        setDAlignFlag(d.alignFlag!==""&&d.alignFlag!=="0");
                        setDPlanSpec(d.planSpeculative!==""&&d.planSpeculative!=="0");
                        setDOiPlanned(d.oiDatePlanned?excelDateToStr(d.oiDatePlanned):"");
                        const seed:Record<string,{loadOv:string;startWk:string}>={};
                        for(const st of STAGES){
                          const ex=adjRows.find(r=>r.oppId===d.oppId&&r.lineId===d.lineId&&r.stage===st);
                          seed[st]={loadOv:ex?.loadOv??"",startWk:ex?.startWk?formatWeek(ex.startWk):""};
                        }
                        setDStageOv(seed);
                      }}>Edit</button>
                    </td>
                  </tr>
                );
              })}
            </tbody></table>
          </div>
        </div>
      )}

      {/* ADJUSTMENTS */}
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
              <label style={s.fieldLbl}>Start week (index or YYYY-WW)<input style={s.inp} value={aStartWk} onChange={e=>setAStartWk(e.target.value)} placeholder="blank = computed"/></label>
            </div>
            <div style={{display:"flex",gap:8,alignItems:"center",flexWrap:"wrap" as const}}>
              <button style={s.saveBtn} onClick={saveAdj} disabled={aSaving}>{aSaving?"Saving…":"Save adjustment"}</button>
              {/* DIFF: delete button inside edit card — only when editing an existing row */}
              {aEditingRow&&!aPendingDelete&&(
                <button style={s.deleteBtnSoft} onClick={()=>setAPendingDelete(true)}>Delete this adjustment</button>
              )}
              {aEditingRow&&aPendingDelete&&(
                <>
                  <button style={s.deleteBtnHard} onClick={deleteAdj} disabled={aSaving}>Confirm delete</button>
                  <button style={s.cancelBtn} onClick={()=>setAPendingDelete(false)}>Cancel</button>
                </>
              )}
              {aEditingRow&&<button style={s.cancelBtn} onClick={()=>{setAEditingRow(null);setAPendingDelete(false);setAOppId("");setALineId("");setALoadOv("");setAStartWk("");}}>Clear form</button>}
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
                  <td style={s.td}>
                    {/* DIFF: Edit button loads row into form for edit-or-delete */}
                    <button style={s.editBtn} onClick={()=>{
                      setAOppId(r.oppId);setALineId(r.lineId);setAStage(r.stage);
                      setALoadOv(r.loadOv);setAStartWk(formatWeek(r.startWk));
                      setAEditingRow(r);setAPendingDelete(false);
                    }}>Edit</button>
                  </td>
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
              {/* DIFF: delete button inside edit card */}
              {coEditingRow&&!coPendingDelete&&(
                <button style={s.deleteBtnSoft} onClick={()=>setCoPendingDelete(true)}>Delete this override</button>
              )}
              {coEditingRow&&coPendingDelete&&(
                <>
                  <button style={s.deleteBtnHard} onClick={deleteCapOv} disabled={coSaving}>Confirm delete</button>
                  <button style={s.cancelBtn} onClick={()=>setCoPendingDelete(false)}>Cancel</button>
                </>
              )}
              {coEditingRow&&<button style={s.cancelBtn} onClick={()=>{setCoEditingRow(null);setCoPendingDelete(false);setCoWeek("");setCoCap("");}}>Clear form</button>}
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
                  <td style={s.tdR}>{r.week?formatWeek(r.week):"—"}</td>
                  <td style={s.tdR}>{r.cap}</td>
                  <td style={s.td}>
                    {/* DIFF: Edit button loads row for edit-or-delete */}
                    <button style={s.editBtn} onClick={()=>{
                      setCoPlant(r.plant);setCoStage(r.stage);
                      setCoWeek(formatWeek(r.week));setCoCap(r.cap);
                      setCoEditingRow(r);setCoPendingDelete(false);
                    }}>Edit</button>
                  </td>
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
  runRow:{display:"flex" as const,gap:8,flexWrap:"wrap" as const,marginBottom:14},
  runBtn:{fontSize:14,padding:"8px 14px",borderRadius:8,border:"1px solid #0284c7",background:"#eff6ff",color:"#0284c7",cursor:"pointer" as const,fontWeight:600},
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
  // DIFF: context note style
  contextNote:{fontSize:14,color:"#475569",background:"#f1f5f9",border:"1px solid #e2e8f0",borderRadius:7,padding:"6px 10px",marginBottom:10},
  saveBtn:{fontSize:14,padding:"8px 16px",borderRadius:8,border:"none",background:"#0f2942",color:"#fff",cursor:"pointer" as const,fontWeight:600},
  cancelBtn:{fontSize:14,padding:"8px 16px",borderRadius:8,border:"1px solid #cbd5e1",background:"#fff",color:"#64748b",cursor:"pointer" as const},
  editBtn:{fontSize:13,padding:"4px 10px",borderRadius:6,border:"1px solid #0284c7",background:"#eff6ff",color:"#0284c7",cursor:"pointer" as const,fontWeight:600},
  // DIFF: delete button styles — soft (first click) and hard (confirm)
  deleteBtnSoft:{fontSize:13,padding:"8px 14px",borderRadius:8,border:"1px solid #fca5a5",background:"#fff",color:"#dc2626",cursor:"pointer" as const,fontWeight:600},
  deleteBtnHard:{fontSize:13,padding:"8px 14px",borderRadius:8,border:"none",background:"#dc2626",color:"#fff",cursor:"pointer" as const,fontWeight:700},
  empty:{color:"#94a3b8",fontSize:16,textAlign:"center" as const,padding:"32px 0",lineHeight:1.8},
  shotBtn:{fontSize:14,padding:"6px 12px",borderRadius:7,border:"1px solid #cbd5e1",background:"#fff",color:"#374151",cursor:"pointer" as const,fontWeight:600},
  coInlineCard:{background:"#fffbeb",border:"1px solid #fde68a",borderRadius:10,padding:10,marginBottom:14},
  saveBtnInline:{fontSize:14,padding:"7px 14px",borderRadius:7,border:"none",background:"#854d0e",color:"#fff",cursor:"pointer" as const,fontWeight:600,alignSelf:"flex-end" as const,height:36},
  stageRow:{display:"flex" as const,alignItems:"center" as const,gap:8,padding:"5px 0",borderBottom:"1px solid #f1f5f9"},
  stageName:{fontSize:14,fontWeight:600,color:"#0f2942",width:70,flexShrink:0},
  stageStd:{fontSize:13,color:"#94a3b8",width:64,flexShrink:0},
  inpSm:{fontSize:14,padding:"5px 7px",borderRadius:6,border:"1px solid #cbd5e1",background:"#fff",flex:1,minWidth:0},
  versionTag:{fontSize:13,fontWeight:400,color:"#94a3b8",marginLeft:6},
  alignFlagRow:{background:"#f0f9ff",border:"1px solid #bae6fd",borderRadius:8,padding:"10px 12px",marginBottom:10,marginTop:4},
  alignFlagLabel:{display:"flex" as const,alignItems:"flex-start" as const,cursor:"pointer" as const,fontSize:15},
  alignFlagText:{fontWeight:700,color:"#0284c7",whiteSpace:"nowrap" as const},
  alignFlagHint:{fontSize:13,color:"#64748b",marginLeft:6,lineHeight:1.5},
  alignBtnOn:{fontSize:13,padding:"3px 8px",borderRadius:6,border:"1px solid #0284c7",background:"#0284c7",color:"#fff",cursor:"pointer" as const,fontWeight:700,whiteSpace:"nowrap" as const},
  alignBtnOff:{fontSize:13,padding:"3px 8px",borderRadius:6,border:"1px solid #cbd5e1",background:"#f8fafc",color:"#94a3b8",cursor:"pointer" as const,fontWeight:400,whiteSpace:"nowrap" as const},
  // DIFF: PlanSpeculative toggle button styles
  planSpecBtnOn:{fontSize:13,padding:"3px 8px",borderRadius:6,border:"1px solid #9d174d",background:"#9d174d",color:"#fff",cursor:"pointer" as const,fontWeight:700,whiteSpace:"nowrap" as const},
  planSpecBtnOff:{fontSize:13,padding:"3px 8px",borderRadius:6,border:"1px solid #cbd5e1",background:"#f8fafc",color:"#94a3b8",cursor:"pointer" as const,fontWeight:400,whiteSpace:"nowrap" as const},
  alignIndicator:{color:"#0284c7",fontWeight:700,fontSize:14},
  alignLegend:{fontSize:13,color:"#64748b",marginLeft:6,fontWeight:400,textTransform:"none" as const,letterSpacing:0},
  alignDot:{color:"#0284c7",fontWeight:700},
  // DIFF: OIDatePlanned row in edit card
  oiRow:{display:"flex" as const,gap:16,marginBottom:10,flexWrap:"wrap" as const},
  oiBlock:{display:"flex" as const,flexDirection:"column" as const,gap:4,flex:1,minWidth:140},
  oiValue:{fontSize:16,fontWeight:600,color:"#1e293b",padding:"6px 0"},
  oiHint:{fontSize:12,fontWeight:400,color:"#94a3b8",textTransform:"none" as const,letterSpacing:0},
};