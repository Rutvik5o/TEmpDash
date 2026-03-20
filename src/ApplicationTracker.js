/* eslint-disable */
import { useState, useCallback, useMemo } from "react";
import {
  BarChart, Bar, PieChart, Pie, Cell, AreaChart, Area, FunnelChart, Funnel,
  XAxis, YAxis, CartesianGrid, Tooltip, ResponsiveContainer, ComposedChart, Line, LabelList
} from "recharts";
import Papa from "papaparse";
import * as XLSX from "xlsx";

/* ══════════════════════════════════════════════
   TOKENS
══════════════════════════════════════════════ */
const C = {
  bg:"#0e1117",panel:"#151821",card:"#1a1f2e",border:"#252b3d",hover:"#1f2536",
  blue:"#2b7de9",teal:"#00c9a7",amber:"#f5a623",red:"#e8445a",purple:"#8b5cf6",
  green:"#22c55e",orange:"#fb923c",pink:"#f472b6",
  text:"#e2e8f0",sub:"#8892a4",dim:"#2d3448",white:"#ffffff",
};
const PAL=[C.blue,C.teal,C.amber,C.red,C.purple,C.green,C.orange,C.pink,"#38bdf8","#a3e635"];

/* ══════════════════════════════════════════════
   NLP ENGINE
══════════════════════════════════════════════ */
const SENTIMENT = {
  positive: ["interested","confirmed","agreed","submitted","approved","positive","willing","ready","cooperative","happy","great","good","done","completed","received","verified","okay","ok","fine","accepted","responsive","called back","available","connected","reachable","success","clear","resolved","excellent","supportive"],
  negative:  ["refused","rejected","issue","problem","unreachable","not picking","not responding","disconnected","wrong number","busy","error","missing","incomplete","failed","denied","delay","pending","unclear","confusion","dispute","complaint","no response","absent","unavailable","declined","lost","invalid","not interested","dropout","left","quit","withdrawn","cancelled"],
  neutral:   ["called","discussed","informed","asked","told","said","mentioned","shared","explained","forwarded","noted","recorded","updated","transferred","scheduled","reminder","follow","waiting","review","check","verify","contact"]
};

const THEMES = [
  { key:"fee_issue",       label:"Fee / Financial Issue",   keywords:["fee","fees","cost","payment","money","fund","financial","afford","expensive","loan","scholarship amount","stipend","installment"] },
  { key:"document",        label:"Document Problem",        keywords:["document","certificate","marksheet","id proof","aadhar","id","proof","form","upload","submit","missing doc","incomplete form","verification"] },
  { key:"eligibility",     label:"Eligibility Query",       keywords:["eligible","eligibility","criteria","qualify","qualification","merit","percentage","marks","score","cutoff","requirement","condition"] },
  { key:"no_response",     label:"No Response / Unreachable",keywords:["not picking","no response","unreachable","not reachable","busy","disconnected","switched off","wrong number","not available","absent"] },
  { key:"family",          label:"Family / Personal Issue", keywords:["family","parent","father","mother","personal","health","medical","home","house","marriage","relative","emergency"] },
  { key:"college",         label:"College / Course Issue",  keywords:["college","institute","course","admission","transfer","change course","dropout","semester","department","university"] },
  { key:"process",         label:"Process / Status Query",  keywords:["status","process","update","when","timeline","pending","how long","application","stage","progress","review","approval"] },
  { key:"positive_close",  label:"Positively Closed",       keywords:["confirmed","done","submitted","completed","approved","verified","accepted","received","resolved","closed","success"] },
];

function analyzeSentiment(text) {
  if (!text || text.trim() === "" || text === "-" || text === "N/A") return "neutral";
  const t = text.toLowerCase();
  let pos = 0, neg = 0;
  SENTIMENT.positive.forEach(w => { if (t.includes(w)) pos++; });
  SENTIMENT.negative.forEach(w => { if (t.includes(w)) neg++; });
  if (pos === 0 && neg === 0) return "neutral";
  if (pos > neg) return "positive";
  if (neg > pos) return "negative";
  return "mixed";
}

function extractThemes(text) {
  if (!text || text.trim() === "" || text === "-") return [];
  const t = text.toLowerCase();
  return THEMES.filter(theme => theme.keywords.some(kw => t.includes(kw))).map(th => th.label);
}

function extractKeywords(texts) {
  const stopWords = new Set(["the","a","an","is","in","on","at","to","for","of","and","or","but","not","with","this","that","was","are","be","have","has","had","been","will","would","can","could","should","do","did","does","it","its","by","from","as","we","he","she","they","i","you","our","his","her","their","my","your","all","also","so","if","then","when","what","how","who","which","were","no","yes","said","told","asked","call","called"]);
  const freq = {};
  texts.forEach(t => {
    if (!t || t === "-") return;
    t.toLowerCase().replace(/[^a-z\s]/g,"").split(/\s+/).forEach(w => {
      if (w.length > 3 && !stopWords.has(w)) freq[w] = (freq[w]||0) + 1;
    });
  });
  return Object.entries(freq).sort((a,b)=>b[1]-a[1]).slice(0,20).map(([word,count])=>({word,count}));
}

/* ══════════════════════════════════════════════
   DATA CLEANING — MODE 2
══════════════════════════════════════════════ */
function cleanMode2(rawRows) {
  return rawRows.map((row, idx) => {
    const r = {};
    Object.keys(row).forEach(k => { r[k.trim()] = typeof row[k]==="string" ? row[k].trim() : (row[k]||""); });

    // Form Status normalize — use Status col as fallback since Form Status is often blank in Excel
    const rawStatus = (r["Form Status"]||r["Status (Pending/Submitted)"]||"").toString();
    const fs = rawStatus.toLowerCase();
    if (fs.includes("approv")||fs.includes("accept")||fs.includes("select")) r["Form Status"] = "Approved";
    else if (fs.includes("reject")||fs.includes("declin")) r["Form Status"] = "Rejected";
    else if (fs.includes("pend")) r["Form Status"] = "Pending";
    else if (fs.includes("submit")) r["Form Status"] = "Submitted";
    else if (fs.includes("review")||fs.includes("progress")||fs.includes("incomplet")) r["Form Status"] = "Under Review";
    else if (fs.includes("inprocess")||fs.includes("in process")) r["Form Status"] = "In Process";
    else r["Form Status"] = "Unknown";

    // Submission status
    const ss = (r["Status (Pending/Submitted)"]||"").toLowerCase();
    if (ss.includes("submit")) r["Status (Pending/Submitted)"] = "Submitted";
    else if (ss.includes("pend")) r["Status (Pending/Submitted)"] = "Pending";
    else if (!ss) r["Status (Pending/Submitted)"] = "Unknown";

    // Stage normalize
    const stage = (r["Stage"]||"").trim();
    r["_Stage"] = stage || "Not Assigned";

    // Reviewer normalize
    r["_Reviewer"] = (r["Reviewer"]||"").trim() || "Unassigned";

    // Mode of application
    const mode = (r["Mode of Application"]||"").toLowerCase();
    if (mode.includes("online")) r["Mode of Application"] = "Online";
    else if (mode.includes("offline") || mode.includes("physical")) r["Mode of Application"] = "Offline";
    else if (!mode) r["Mode of Application"] = "Unknown";

    // Fix duplicate "Date" headers — XLSX renames them Date, Date_1, Date_2...
    // So we remap by checking all possible key variants
    const callKeys = Object.keys(r);
    const callColMap = {};
    ["1st Call","2nd Call","3rd Call","4th Call","5th Call"].forEach(name => {
      // direct match
      if (r[name] !== undefined) { callColMap[name] = r[name]; return; }
      // try case-insensitive
      const found = callKeys.find(k => k.toLowerCase().replace(/\s/g,"") === name.toLowerCase().replace(/\s/g,""));
      callColMap[name] = found ? r[found] : "";
    });

    // Analyze each call
    const CALLS = ["1st Call","2nd Call","3rd Call","4th Call","5th Call"];
    let callsMade = 0;
    let allCallText = [];
    r["_CallSentiments"] = [];
    r["_CallThemes"] = [];

    CALLS.forEach((col, i) => {
      const txt = (callColMap[col] || r[col] || "").toString().trim();
      const hasContent = txt && txt !== "-" && txt !== "N/A" && txt.trim() !== "";
      if (hasContent) {
        callsMade++;
        allCallText.push(txt);
        r["_CallSentiments"].push({ call: col, sentiment: analyzeSentiment(txt), text: txt });
        extractThemes(txt).forEach(th => { if (!r["_CallThemes"].includes(th)) r["_CallThemes"].push(th); });
      } else {
        r["_CallSentiments"].push({ call: col, sentiment: "no_call", text: "" });
      }
    });

    r["_CallsMade"] = callsMade;
    r["_AllCallText"] = allCallText.join(" ");

    // Final remarks NLP
    const remarks = r["Final Remarks"] || "";
    r["_RemarkSentiment"] = analyzeSentiment(remarks);
    r["_RemarkThemes"] = extractThemes(remarks);

    // Overall sentiment = last non-empty call OR remarks
    const lastCallText = allCallText[allCallText.length-1] || remarks;
    r["_OverallSentiment"] = analyzeSentiment(lastCallText || remarks);

    // Rejection date present?
    r["_IsRejected"] = !!(r["Rejected Date"] && r["Rejected Date"].trim() && r["Rejected Date"] !== "-");

    r["_idx"] = idx;
    r["_id"] = r["S.No"] || r["S.No."] || (idx+1);
    return r;
  });
}

/* ══════════════════════════════════════════════
   HELPERS
══════════════════════════════════════════════ */
const groupBy = (arr,key) => {
  const m={};
  arr.forEach(r=>{const v=r[key]||"Unknown";m[v]=(m[v]||0)+1;});
  return Object.entries(m).sort((a,b)=>b[1]-a[1]).map(([name,value])=>({name,value}));
};
const pct=(n,t)=>t?((n/t)*100).toFixed(1):"0.0";
const sentimentColor = s => s==="positive"?C.green:s==="negative"?C.red:s==="mixed"?C.amber:s==="no_call"?C.dim:C.sub;
const sentimentLabel = s => s==="positive"?"😊 Positive":s==="negative"?"😟 Negative":s==="mixed"?"😐 Mixed":s==="no_call"?"—":"😑 Neutral";

/* ══════════════════════════════════════════════
   TOOLTIP
══════════════════════════════════════════════ */
const DT = ({active,payload,label}) => {
  if(!active||!payload?.length) return null;
  return (
    <div style={{background:C.panel,border:`1px solid ${C.border}`,borderRadius:8,padding:"10px 14px",fontFamily:"'JetBrains Mono',monospace",fontSize:12,boxShadow:"0 8px 24px #00000080"}}>
      {label!==undefined&&<div style={{color:C.sub,marginBottom:6,fontSize:11}}>{label}</div>}
      {payload.map((p,i)=>(
        <div key={i} style={{color:p.color||C.teal,marginBottom:2}}>
          {p.name}: <strong style={{color:C.white}}>{typeof p.value==="number"?p.value.toLocaleString():p.value}</strong>
        </div>
      ))}
    </div>
  );
};

/* ══════════════════════════════════════════════
   COMPONENTS
══════════════════════════════════════════════ */
function KPI({label,value,sub,accent,delay=0}) {
  return (
    <div style={{background:C.card,border:`1px solid ${C.border}`,borderRadius:12,padding:"18px 20px",position:"relative",overflow:"hidden","--accent":accent||C.blue,animationDelay:`${delay}ms`}}
      className="kpi-card fade">
      <div style={{fontSize:11,color:C.sub,fontWeight:600,letterSpacing:.8,textTransform:"uppercase",marginBottom:10}}>{label}</div>
      <div style={{fontSize:28,fontWeight:800,color:accent||C.white,lineHeight:1}}>{value}</div>
      {sub&&<div style={{fontSize:12,color:C.sub,marginTop:6}}>{sub}</div>}
    </div>
  );
}

function HBar({name,value,total,color,rank,sub}) {
  const p=total?(value/total)*100:0;
  return (
    <div style={{marginBottom:12}}>
      <div style={{display:"flex",justifyContent:"space-between",marginBottom:5}}>
        <span style={{fontSize:12,color:C.text,display:"flex",alignItems:"center",gap:6}}>
          {rank&&<span style={{fontSize:10,color:C.sub,width:16}}>#{rank}</span>}{name}
          {sub&&<span style={{fontSize:11,color:C.sub}}>— {sub}</span>}
        </span>
        <span style={{fontSize:12,fontWeight:700,color:color||C.blue,fontFamily:"'JetBrains Mono',monospace"}}>
          {value} <span style={{color:C.sub,fontWeight:400}}>({p.toFixed(1)}%)</span>
        </span>
      </div>
      <div style={{height:6,background:C.dim,borderRadius:3,overflow:"hidden"}}>
        <div style={{height:"100%",borderRadius:3,background:color||C.blue,width:`${p}%`,transition:"width .8s ease"}}/>
      </div>
    </div>
  );
}

/* Sentiment Badge */
function SBadge({sentiment}) {
  const color=sentimentColor(sentiment);
  const label=sentimentLabel(sentiment);
  return <span style={{display:"inline-flex",alignItems:"center",padding:"3px 10px",borderRadius:20,fontSize:11,fontWeight:600,background:`${color}20`,color,border:`1px solid ${color}40`}}>{label}</span>;
}

/* Call dots timeline */
function CallTimeline({callSentiments}) {
  return (
    <div style={{display:"flex",gap:6,alignItems:"center"}}>
      {callSentiments.map((c,i)=>(
        <div key={i} title={`${c.call}: ${c.text||"No call"}`}
          style={{width:20,height:20,borderRadius:"50%",background:sentimentColor(c.sentiment),display:"flex",alignItems:"center",justifyContent:"center",fontSize:9,fontWeight:700,color:C.bg,cursor:"default",flexShrink:0,border:c.sentiment==="no_call"?`1px dashed ${C.dim}`:"none",opacity:c.sentiment==="no_call"?0.3:1}}>
          {i+1}
        </div>
      ))}
    </div>
  );
}

/* ══════════════════════════════════════════════
   MAIN MODE2 DASHBOARD
══════════════════════════════════════════════ */
export default function ApplicationTracker({onBack}) {
  const [data,setData]=useState([]);
  const [fileName,setFileName]=useState("");
  const [tab,setTab]=useState("overview");
  const [filterStatus,setFilterStatus]=useState("All");
  const [filterReviewer,setFilterReviewer]=useState("All");
  const [filterSentiment,setFilterSentiment]=useState("All");
  const [search,setSearch]=useState("");
  const [expandedRow,setExpandedRow]=useState(null);
  const [drag,setDrag]=useState(false);

  const handleFile=useCallback(file=>{
    setFileName(file.name);
    const ext=file.name.split(".").pop().toLowerCase();
    if(ext==="xlsx"||ext==="xls"){
      const reader=new FileReader();
      reader.onload=e=>{
        const wb=XLSX.read(e.target.result,{type:"array"});
        // Use "All Forms" sheet if exists, else first sheet
        const sheetName=wb.SheetNames.includes("All Forms")?"All Forms":wb.SheetNames[0];
        const ws=wb.Sheets[sheetName];
        const rows=XLSX.utils.sheet_to_json(ws,{defval:"",raw:false});
        const cleaned=cleanMode2(rows.map(row=>{
          const r={};
          Object.keys(row).forEach(k=>{
            const v=row[k];
            r[k.trim()]=(v===null||v===undefined)?"":String(v).trim();
          });
          return r;
        }));
        setData(cleaned);
      };
      reader.readAsArrayBuffer(file);
    } else {
      Papa.parse(file,{header:true,skipEmptyLines:true,complete:r=>setData(cleanMode2(r.data))});
    }
  // eslint-disable-next-line
  },[]);

  const onDrop=useCallback(e=>{
    e.preventDefault();setDrag(false);
    const f=e.dataTransfer.files[0];if(f)handleFile(f);
  },[handleFile]);

  const filtered=useMemo(()=>{
    let d=data;
    if(filterStatus!=="All") d=d.filter(r=>r["Form Status"]===filterStatus);
    if(filterReviewer!=="All") d=d.filter(r=>r["_Reviewer"]===filterReviewer);
    if(filterSentiment!=="All") d=d.filter(r=>r["_OverallSentiment"]===filterSentiment.toLowerCase());
    if(search) d=d.filter(r=>(r["Name"]||"").toLowerCase().includes(search.toLowerCase())||String(r["_id"]).includes(search));
    return d;
  },[data,filterStatus,filterReviewer,filterSentiment,search]);

  const total=filtered.length;

  // Aggregations
  const formStatuses=useMemo(()=>groupBy(filtered,"Form Status"),[filtered]);
  const reviewers=useMemo(()=>groupBy(filtered,"_Reviewer").filter(r=>r.name!=="Unassigned"),[filtered]);
  const stages=useMemo(()=>groupBy(filtered,"_Stage"),[filtered]);
  const modes=useMemo(()=>groupBy(filtered,"Mode of Application"),[filtered]);

  const sentimentCounts=useMemo(()=>{
    const m={positive:0,negative:0,neutral:0,mixed:0};
    filtered.forEach(r=>{ if(m[r["_OverallSentiment"]]!==undefined) m[r["_OverallSentiment"]]++; });
    return Object.entries(m).map(([name,value])=>({name,value})).filter(x=>x.value>0);
  },[filtered]);

  const callDepth=useMemo(()=>{
    const m={0:0,1:0,2:0,3:0,4:0,5:0};
    filtered.forEach(r=>{ const n=Math.min(r["_CallsMade"],5); m[n]=(m[n]||0)+1; });
    return Object.entries(m).filter(([,v])=>v>0).map(([calls,count])=>({calls:`${calls} calls`,count}));
  },[filtered]);

  // Theme frequency across all calls + remarks
  const themeFreq=useMemo(()=>{
    const m={};
    filtered.forEach(r=>{
      [...r["_CallThemes"],...r["_RemarkThemes"]].forEach(t=>{m[t]=(m[t]||0)+1;});
    });
    return Object.entries(m).sort((a,b)=>b[1]-a[1]).map(([name,value])=>({name,value}));
  },[filtered]);

  // Keywords from all call text
  const keywords=useMemo(()=>{
    const allTexts=filtered.map(r=>r["_AllCallText"]+" "+(r["Final Remarks"]||""));
    return extractKeywords(allTexts);
  },[filtered]);

  // Reviewer performance
  const reviewerStats=useMemo(()=>{
    return reviewers.map(rv=>{
      const sub=filtered.filter(r=>r["_Reviewer"]===rv.name);
      const approved=sub.filter(r=>r["Form Status"]==="Approved").length;
      const rejected=sub.filter(r=>r["Form Status"]==="Rejected").length;
      const posS=sub.filter(r=>r["_OverallSentiment"]==="positive").length;
      return{name:rv.name,total:rv.value,approved,rejected,pending:rv.value-approved-rejected,approvalRate:rv.value?((approved/rv.value)*100).toFixed(1):0,positivePct:rv.value?((posS/rv.value)*100).toFixed(1):0};
    });
  },[filtered,reviewers]);

  // Stage funnel data
  const stageFunnel=useMemo(()=>{
    const stageOrder=["Received","Under Review","Verification","Interview","Approved","Rejected"];
    return stageOrder.map(s=>({name:s,value:filtered.filter(r=>r["_Stage"]===s||r["Form Status"]===s).length})).filter(x=>x.value>0);
  },[filtered]);

  // Call sentiment per call number
  const callSentimentBreakdown=useMemo(()=>{
    const CALLS=["1st Call","2nd Call","3rd Call","4th Call","5th Call"];
    return CALLS.map(col=>{
      let pos=0,neg=0,neu=0,mix=0,total=0;
      filtered.forEach(r=>{
        const cs=r["_CallSentiments"].find(c=>c.call===col);
        if(cs&&cs.sentiment!=="no_call"){total++;if(cs.sentiment==="positive")pos++;else if(cs.sentiment==="negative")neg++;else if(cs.sentiment==="mixed")mix++;else neu++;}
      });
      return{call:col.replace(" Call",""),Positive:pos,Negative:neg,Neutral:neu,Mixed:mix,total};
    }).filter(x=>x.total>0);
  },[filtered]);

  const allStatuses=["All",...[...new Set(data.map(r=>r["Form Status"]))].filter(Boolean)];
  const allReviewers=["All",...[...new Set(data.map(r=>r["_Reviewer"]))].filter(r=>r!=="Unassigned")];
  const TABS=["overview","call analysis","nlp insights","reviewers","applicants"];
  const statusColors={Approved:C.green,Rejected:C.red,Pending:C.amber,Submitted:C.blue,"Under Review":C.purple,Unknown:C.sub};

  if(!data.length) return (
    <div style={{minHeight:"100vh",background:C.bg,display:"flex",flexDirection:"column",alignItems:"center",justifyContent:"center",padding:24}}>
      <div style={{position:"absolute",inset:0,backgroundImage:`radial-gradient(${C.purple}12 1px,transparent 1px)`,backgroundSize:"32px 32px"}}/>
      <div className="fade" style={{position:"relative",zIndex:1,textAlign:"center",maxWidth:600,width:"100%"}}>
        <button onClick={onBack} style={{position:"absolute",top:-60,left:0,background:"none",border:`1px solid ${C.border}`,color:C.sub,padding:"6px 14px",borderRadius:7,cursor:"pointer",fontSize:12}}>← Mode 1</button>
        <div style={{fontSize:38,fontWeight:800,letterSpacing:-1.5,marginBottom:6}}>
          <span style={{color:C.purple}}>Track</span><span style={{color:C.orange}}>Flow</span>
        </div>
        <p style={{color:C.sub,fontSize:13,fontFamily:"'JetBrains Mono',monospace",marginBottom:48,letterSpacing:1}}>APPLICATION TRACKING · NLP CALL ANALYSIS</p>
        <div onDragOver={e=>{e.preventDefault();setDrag(true);}} onDragLeave={()=>setDrag(false)} onDrop={onDrop}
          style={{border:`2px dashed ${drag?C.purple:C.dim}`,borderRadius:20,background:drag?`${C.purple}08`:C.card,padding:"56px 48px",transition:"all .3s"}}>
          <div style={{fontSize:48,marginBottom:16}}>📞</div>
          <div style={{fontSize:20,fontWeight:700,marginBottom:8}}>Drop your Application Tracking CSV</div>
          <div style={{color:C.sub,fontSize:13,marginBottom:12}}>Expects: S.No, Form Status, Mode, Name, Stage, Reviewer, Call columns, Final Remarks</div>
          <div style={{color:C.sub,fontSize:12,marginBottom:28,fontFamily:"'JetBrains Mono',monospace"}}>NLP will auto-analyze all call notes & remarks</div>
          <label style={{background:`linear-gradient(135deg,${C.purple},${C.orange})`,color:C.white,padding:"12px 32px",borderRadius:10,cursor:"pointer",fontWeight:700,fontSize:14,display:"inline-block"}}>
            CHOOSE FILE<input type="file" accept=".csv,.xlsx,.xls" onChange={e=>e.target.files[0]&&handleFile(e.target.files[0])}/>
          </label>
        </div>
        <div style={{display:"flex",gap:12,justifyContent:"center",marginTop:24,flexWrap:"wrap"}}>
          {["Call sentiment NLP","Theme extraction","Reviewer analytics","Stage funnel","Applicant deep-dive"].map(f=>(
            <span style={{background:C.dim,border:`1px solid ${C.border}`,borderRadius:6,padding:"3px 10px",fontSize:11,color:C.sub}} key={f}>{f}</span>
          ))}
        </div>
      </div>
    </div>
  );

  return (
    <div style={{minHeight:"100vh",background:C.bg}}>
      {/* TOPBAR */}
      <div style={{background:C.panel,borderBottom:`1px solid ${C.border}`,padding:"0 24px",display:"flex",alignItems:"center",gap:16,position:"sticky",top:0,zIndex:100}}>
        <button onClick={onBack} style={{background:"none",border:`1px solid ${C.border}`,color:C.sub,padding:"6px 14px",borderRadius:7,cursor:"pointer",fontSize:12,whiteSpace:"nowrap",flexShrink:0}}>← Mode 1</button>
        <div style={{fontWeight:800,fontSize:18,whiteSpace:"nowrap"}}><span style={{color:C.purple}}>Track</span><span style={{color:C.orange}}>Flow</span></div>
        <div style={{width:1,height:28,background:C.border,flexShrink:0}}/>
        <div style={{fontSize:12,color:C.sub,fontFamily:"'JetBrains Mono',monospace",flex:1,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>{fileName}</div>
        <span style={{background:C.dim,border:`1px solid ${C.border}`,borderRadius:6,padding:"3px 10px",fontSize:11,color:C.sub,flexShrink:0}}>{total.toLocaleString()} records</span>
        <button onClick={()=>{setData([]);setFileName("");}} style={{background:"none",border:`1px solid ${C.border}`,color:C.sub,padding:"6px 14px",borderRadius:7,cursor:"pointer",fontSize:12,whiteSpace:"nowrap",flexShrink:0}}>↑ New File</button>
      </div>

      {/* FILTERS */}
      <div style={{background:C.panel,borderBottom:`1px solid ${C.border}`,padding:"10px 24px",display:"flex",gap:10,alignItems:"center",flexWrap:"wrap"}}>
        <span style={{fontSize:11,color:C.sub,fontWeight:600,letterSpacing:1}}>FILTERS</span>
        <select value={filterStatus} onChange={e=>setFilterStatus(e.target.value)} style={{background:C.dim,border:`1px solid ${C.border}`,color:C.text,padding:"6px 12px",borderRadius:7,fontSize:12,outline:"none",cursor:"pointer"}}>
          {allStatuses.map(o=><option key={o}>{o}</option>)}
        </select>
        <select value={filterReviewer} onChange={e=>setFilterReviewer(e.target.value)} style={{background:C.dim,border:`1px solid ${C.border}`,color:C.text,padding:"6px 12px",borderRadius:7,fontSize:12,outline:"none",cursor:"pointer"}}>
          {allReviewers.map(o=><option key={o}>{o}</option>)}
        </select>
        <select value={filterSentiment} onChange={e=>setFilterSentiment(e.target.value)} style={{background:C.dim,border:`1px solid ${C.border}`,color:C.text,padding:"6px 12px",borderRadius:7,fontSize:12,outline:"none",cursor:"pointer"}}>
          {["All","Positive","Negative","Neutral","Mixed"].map(o=><option key={o}>{o}</option>)}
        </select>
        <input value={search} onChange={e=>setSearch(e.target.value)} placeholder="Search name or ID..." style={{background:C.dim,border:`1px solid ${C.border}`,color:C.text,padding:"6px 12px",borderRadius:7,fontSize:12,outline:"none",fontFamily:"'Plus Jakarta Sans',sans-serif"}}/>
        {(filterStatus!=="All"||filterReviewer!=="All"||filterSentiment!=="All"||search)&&(
          <button onClick={()=>{setFilterStatus("All");setFilterReviewer("All");setFilterSentiment("All");setSearch("");}} style={{background:"none",border:`1px solid ${C.red}60`,color:C.red,padding:"5px 12px",borderRadius:6,cursor:"pointer",fontSize:11}}>✕ Clear</button>
        )}
      </div>

      {/* TABS */}
      <div style={{background:C.panel,borderBottom:`1px solid ${C.border}`,padding:"0 24px",display:"flex",overflowX:"auto"}}>
        {TABS.map(t=>(
          <button key={t} onClick={()=>setTab(t)} style={{background:"none",border:"none",cursor:"pointer",padding:"10px 18px",fontSize:13,fontFamily:"'Plus Jakarta Sans',sans-serif",fontWeight:600,color:tab===t?C.purple:C.sub,borderBottom:`2px solid ${tab===t?C.purple:"transparent"}`,transition:"all .2s",whiteSpace:"nowrap"}}>
            {t.charAt(0).toUpperCase()+t.slice(1)}
          </button>
        ))}
      </div>

      <div style={{padding:"24px",maxWidth:1400,margin:"0 auto"}}>

        {/* ══ OVERVIEW ══ */}
        {tab==="overview"&&(
          <>
            <div style={{display:"grid",gridTemplateColumns:"repeat(auto-fit,minmax(160px,1fr))",gap:14,marginBottom:20}}>
              <KPI label="Total Applications" value={total} accent={C.purple} delay={0}/>
              <KPI label="Approved" value={filtered.filter(r=>r["Form Status"]==="Approved").length} sub={`${pct(filtered.filter(r=>r["Form Status"]==="Approved").length,total)}%`} accent={C.green} delay={60}/>
              <KPI label="Rejected" value={filtered.filter(r=>r["Form Status"]==="Rejected").length} sub={`${pct(filtered.filter(r=>r["Form Status"]==="Rejected").length,total)}%`} accent={C.red} delay={120}/>
              <KPI label="Pending" value={filtered.filter(r=>r["Form Status"]==="Pending").length} accent={C.amber} delay={180}/>
              <KPI label="Online Applications" value={filtered.filter(r=>r["Mode of Application"]==="Online").length} accent={C.blue} delay={240}/>
              <KPI label="Positive Sentiment" value={filtered.filter(r=>r["_OverallSentiment"]==="positive").length} sub={`${pct(filtered.filter(r=>r["_OverallSentiment"]==="positive").length,total)}% of calls`} accent={C.green} delay={300}/>
            </div>

            <div style={{display:"grid",gridTemplateColumns:"1fr 1fr 1fr",gap:16,marginBottom:18}}>
              <div style={{background:C.card,border:`1px solid ${C.border}`,borderRadius:12,padding:"20px 22px"}}>
                <div style={{fontSize:11,fontWeight:700,letterSpacing:1.5,textTransform:"uppercase",color:C.sub,marginBottom:16,display:"flex",alignItems:"center",gap:8}}>Form Status<span style={{flex:1,height:1,background:C.border}}/></div>
                <ResponsiveContainer width="100%" height={170}><PieChart><Pie data={formStatuses} dataKey="value" nameKey="name" cx="50%" cy="50%" outerRadius={70} innerRadius={35}>{formStatuses.map((s,i)=><Cell key={i} fill={statusColors[s.name]||PAL[i]}/>)}</Pie><Tooltip content={<DT/>}/></PieChart></ResponsiveContainer>
                {formStatuses.map((s,i)=>(
                  <div key={i} style={{display:"flex",justifyContent:"space-between",marginBottom:6,alignItems:"center"}}>
                    <div style={{display:"flex",alignItems:"center",gap:6}}><div style={{width:8,height:8,borderRadius:"50%",background:statusColors[s.name]||PAL[i]}}/><span style={{fontSize:12}}>{s.name}</span></div>
                    <span style={{fontSize:12,fontFamily:"'JetBrains Mono',monospace",color:statusColors[s.name]||PAL[i],fontWeight:700}}>{s.value} ({pct(s.value,total)}%)</span>
                  </div>
                ))}
              </div>

              <div style={{background:C.card,border:`1px solid ${C.border}`,borderRadius:12,padding:"20px 22px"}}>
                <div style={{fontSize:11,fontWeight:700,letterSpacing:1.5,textTransform:"uppercase",color:C.sub,marginBottom:16,display:"flex",alignItems:"center",gap:8}}>Overall Sentiment<span style={{flex:1,height:1,background:C.border}}/></div>
                <ResponsiveContainer width="100%" height={170}><PieChart><Pie data={sentimentCounts} dataKey="value" nameKey="name" cx="50%" cy="50%" outerRadius={70} innerRadius={35}>{sentimentCounts.map(s=><Cell key={s.name} fill={sentimentColor(s.name)}/>)}</Pie><Tooltip content={<DT/>}/></PieChart></ResponsiveContainer>
                {sentimentCounts.map((s,i)=>(
                  <div key={i} style={{display:"flex",justifyContent:"space-between",marginBottom:6,alignItems:"center"}}>
                    <div style={{display:"flex",alignItems:"center",gap:6}}><div style={{width:8,height:8,borderRadius:"50%",background:sentimentColor(s.name)}}/><span style={{fontSize:12,textTransform:"capitalize"}}>{s.name}</span></div>
                    <span style={{fontSize:12,fontFamily:"'JetBrains Mono',monospace",color:sentimentColor(s.name),fontWeight:700}}>{s.value} ({pct(s.value,total)}%)</span>
                  </div>
                ))}
              </div>

              <div style={{background:C.card,border:`1px solid ${C.border}`,borderRadius:12,padding:"20px 22px"}}>
                <div style={{fontSize:11,fontWeight:700,letterSpacing:1.5,textTransform:"uppercase",color:C.sub,marginBottom:16,display:"flex",alignItems:"center",gap:8}}>Application Mode<span style={{flex:1,height:1,background:C.border}}/></div>
                {modes.map((m,i)=><HBar key={i} name={m.name} value={m.value} total={total} color={PAL[i]} rank={i+1}/>)}
                <div style={{marginTop:16,fontSize:11,fontWeight:700,letterSpacing:1.5,textTransform:"uppercase",color:C.sub,marginBottom:12}}>Calls Per Applicant</div>
                {callDepth.map((d,i)=><HBar key={i} name={d.calls} value={d.count} total={total} color={PAL[i+2]}/>)}
              </div>
            </div>

            {/* Stage breakdown */}
            <div style={{background:C.card,border:`1px solid ${C.border}`,borderRadius:12,padding:"20px 22px",marginBottom:18}}>
              <div style={{fontSize:11,fontWeight:700,letterSpacing:1.5,textTransform:"uppercase",color:C.sub,marginBottom:16,display:"flex",alignItems:"center",gap:8}}>Stage Distribution<span style={{flex:1,height:1,background:C.border}}/></div>
              <ResponsiveContainer width="100%" height={200}>
                <BarChart data={stages}>
                  <CartesianGrid strokeDasharray="3 3" stroke={C.dim}/>
                  <XAxis dataKey="name" tick={{fontSize:11,fill:C.sub}}/><YAxis tick={{fontSize:10,fill:C.sub}}/>
                  <Tooltip content={<DT/>}/>
                  <Bar dataKey="value" name="Applications" radius={[4,4,0,0]}>{stages.map((_,i)=><Cell key={i} fill={PAL[i]}/>)}</Bar>
                </BarChart>
              </ResponsiveContainer>
            </div>
          </>
        )}

        {/* ══ CALL ANALYSIS ══ */}
        {tab==="call analysis"&&(
          <>
            <div style={{display:"grid",gridTemplateColumns:"repeat(auto-fit,minmax(160px,1fr))",gap:14,marginBottom:20}}>
              <KPI label="Avg Calls Per Case" value={(filtered.reduce((s,r)=>s+r["_CallsMade"],0)/Math.max(total,1)).toFixed(1)} accent={C.purple}/>
              <KPI label="Resolved in 1 Call" value={filtered.filter(r=>r["_CallsMade"]===1).length} sub={`${pct(filtered.filter(r=>r["_CallsMade"]===1).length,total)}%`} accent={C.green}/>
              <KPI label="Needed 3+ Calls" value={filtered.filter(r=>r["_CallsMade"]>=3).length} sub={`${pct(filtered.filter(r=>r["_CallsMade"]>=3).length,total)}%`} accent={C.amber}/>
              <KPI label="No Calls Made" value={filtered.filter(r=>r["_CallsMade"]===0).length} accent={C.red}/>
              <KPI label="Positive Last Call" value={filtered.filter(r=>r["_OverallSentiment"]==="positive").length} accent={C.teal}/>
              <KPI label="Negative Last Call" value={filtered.filter(r=>r["_OverallSentiment"]==="negative").length} accent={C.red}/>
            </div>

            <div style={{background:C.card,border:`1px solid ${C.border}`,borderRadius:12,padding:"20px 22px",marginBottom:18}}>
              <div style={{fontSize:11,fontWeight:700,letterSpacing:1.5,textTransform:"uppercase",color:C.sub,marginBottom:16,display:"flex",alignItems:"center",gap:8}}>Sentiment Per Call Stage<span style={{flex:1,height:1,background:C.border}}/></div>
              <ResponsiveContainer width="100%" height={240}>
                <BarChart data={callSentimentBreakdown}>
                  <CartesianGrid strokeDasharray="3 3" stroke={C.dim}/>
                  <XAxis dataKey="call" tick={{fontSize:12,fill:C.text}}/><YAxis tick={{fontSize:10,fill:C.sub}}/>
                  <Tooltip content={<DT/>}/>
                  <Bar dataKey="Positive" stackId="a" fill={C.green} name="Positive"/>
                  <Bar dataKey="Mixed" stackId="a" fill={C.amber} name="Mixed"/>
                  <Bar dataKey="Neutral" stackId="a" fill={C.sub} name="Neutral"/>
                  <Bar dataKey="Negative" stackId="a" fill={C.red} name="Negative" radius={[4,4,0,0]}/>
                </BarChart>
              </ResponsiveContainer>
              <div style={{display:"flex",gap:16,marginTop:12,justifyContent:"center",flexWrap:"wrap"}}>
                {[["Positive",C.green],["Mixed",C.amber],["Neutral",C.sub],["Negative",C.red]].map(([l,c])=>(
                  <div key={l} style={{display:"flex",alignItems:"center",gap:6,fontSize:12}}><div style={{width:10,height:10,borderRadius:2,background:c}}/>{l}</div>
                ))}
              </div>
            </div>

            <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:16,marginBottom:18}}>
              <div style={{background:C.card,border:`1px solid ${C.border}`,borderRadius:12,padding:"20px 22px"}}>
                <div style={{fontSize:11,fontWeight:700,letterSpacing:1.5,textTransform:"uppercase",color:C.sub,marginBottom:16,display:"flex",alignItems:"center",gap:8}}>Calls Needed to Close<span style={{flex:1,height:1,background:C.border}}/></div>
                <ResponsiveContainer width="100%" height={200}>
                  <BarChart data={callDepth}>
                    <CartesianGrid strokeDasharray="3 3" stroke={C.dim}/>
                    <XAxis dataKey="calls" tick={{fontSize:11,fill:C.sub}}/><YAxis tick={{fontSize:10,fill:C.sub}}/>
                    <Tooltip content={<DT/>}/>
                    <Bar dataKey="count" name="Applicants" radius={[4,4,0,0]}>{callDepth.map((_,i)=><Cell key={i} fill={PAL[i]}/>)}</Bar>
                  </BarChart>
                </ResponsiveContainer>
              </div>
              <div style={{background:C.card,border:`1px solid ${C.border}`,borderRadius:12,padding:"20px 22px"}}>
                <div style={{fontSize:11,fontWeight:700,letterSpacing:1.5,textTransform:"uppercase",color:C.sub,marginBottom:16,display:"flex",alignItems:"center",gap:8}}>Top Discussion Themes<span style={{flex:1,height:1,background:C.border}}/></div>
                {themeFreq.slice(0,8).map((t,i)=><HBar key={i} name={t.name} value={t.value} total={total} color={PAL[i]} rank={i+1}/>)}
              </div>
            </div>
          </>
        )}

        {/* ══ NLP INSIGHTS ══ */}
        {tab==="nlp insights"&&(
          <>
            {/* Sentiment summary */}
            <div style={{display:"grid",gridTemplateColumns:"repeat(4,1fr)",gap:14,marginBottom:20}}>
              {[["Positive",C.green,"positive"],["Negative",C.red,"negative"],["Mixed",C.amber,"mixed"],["Neutral",C.sub,"neutral"]].map(([label,color,key])=>{
                const n=filtered.filter(r=>r["_RemarkSentiment"]===key).length;
                return(
                  <div key={key} style={{background:C.card,border:`1px solid ${C.border}`,borderRadius:12,padding:"18px 20px",borderTop:`2px solid ${color}`}}>
                    <div style={{fontSize:11,color:C.sub,fontWeight:600,letterSpacing:.8,textTransform:"uppercase",marginBottom:10}}>Remarks — {label}</div>
                    <div style={{fontSize:28,fontWeight:800,color,lineHeight:1}}>{n}</div>
                    <div style={{fontSize:12,color:C.sub,marginTop:6}}>{pct(n,total)}% of applicants</div>
                  </div>
                );
              })}
            </div>

            {/* Theme freq chart */}
            <div style={{background:C.card,border:`1px solid ${C.border}`,borderRadius:12,padding:"20px 22px",marginBottom:18}}>
              <div style={{fontSize:11,fontWeight:700,letterSpacing:1.5,textTransform:"uppercase",color:C.sub,marginBottom:16,display:"flex",alignItems:"center",gap:8}}>Discussion Themes Frequency (Calls + Remarks)<span style={{flex:1,height:1,background:C.border}}/></div>
              <ResponsiveContainer width="100%" height={240}>
                <BarChart data={themeFreq} layout="vertical">
                  <CartesianGrid strokeDasharray="3 3" stroke={C.dim} vertical={false}/>
                  <XAxis type="number" tick={{fontSize:10,fill:C.sub}}/>
                  <YAxis dataKey="name" type="category" width={180} tick={{fontSize:11,fill:C.text}}/>
                  <Tooltip content={<DT/>}/>
                  <Bar dataKey="value" name="Occurrences" radius={[0,4,4,0]}>{themeFreq.map((_,i)=><Cell key={i} fill={PAL[i%PAL.length]}/>)}</Bar>
                </BarChart>
              </ResponsiveContainer>
            </div>

            {/* Keyword cloud as bar */}
            <div style={{background:C.card,border:`1px solid ${C.border}`,borderRadius:12,padding:"20px 22px",marginBottom:18}}>
              <div style={{fontSize:11,fontWeight:700,letterSpacing:1.5,textTransform:"uppercase",color:C.sub,marginBottom:16,display:"flex",alignItems:"center",gap:8}}>Top Keywords in Call Notes<span style={{flex:1,height:1,background:C.border}}/></div>
              <div style={{display:"flex",flexWrap:"wrap",gap:10}}>
                {keywords.map((k,i)=>{
                  const size=Math.max(11,Math.min(22,11+(k.count/keywords[0].count)*11));
                  const opacity=0.5+(k.count/keywords[0].count)*0.5;
                  return(
                    <div key={k.word} style={{background:`${PAL[i%PAL.length]}20`,border:`1px solid ${PAL[i%PAL.length]}40`,borderRadius:8,padding:"6px 12px",fontSize:size,color:PAL[i%PAL.length],fontWeight:600,opacity,fontFamily:"'JetBrains Mono',monospace"}}>
                      {k.word} <span style={{fontSize:10,opacity:.7}}>×{k.count}</span>
                    </div>
                  );
                })}
              </div>
            </div>

            {/* Sentiment vs Form Status */}
            <div style={{background:C.card,border:`1px solid ${C.border}`,borderRadius:12,padding:"20px 22px"}}>
              <div style={{fontSize:11,fontWeight:700,letterSpacing:1.5,textTransform:"uppercase",color:C.sub,marginBottom:16,display:"flex",alignItems:"center",gap:8}}>Sentiment → Final Outcome<span style={{flex:1,height:1,background:C.border}}/></div>
              <ResponsiveContainer width="100%" height={220}>
                <BarChart data={["positive","negative","neutral","mixed"].map(s=>{
                  const sub=filtered.filter(r=>r["_OverallSentiment"]===s);
                  return{sentiment:s.charAt(0).toUpperCase()+s.slice(1),Approved:sub.filter(r=>r["Form Status"]==="Approved").length,Rejected:sub.filter(r=>r["Form Status"]==="Rejected").length,Pending:sub.filter(r=>r["Form Status"]==="Pending").length,total:sub.length};
                }).filter(x=>x.total>0)}>
                  <CartesianGrid strokeDasharray="3 3" stroke={C.dim}/>
                  <XAxis dataKey="sentiment" tick={{fontSize:12,fill:C.text}}/><YAxis tick={{fontSize:10,fill:C.sub}}/>
                  <Tooltip content={<DT/>}/>
                  <Bar dataKey="Approved" fill={C.green} stackId="a"/>
                  <Bar dataKey="Pending" fill={C.amber} stackId="a"/>
                  <Bar dataKey="Rejected" fill={C.red} stackId="a" radius={[4,4,0,0]}/>
                </BarChart>
              </ResponsiveContainer>
            </div>
          </>
        )}

        {/* ══ REVIEWERS ══ */}
        {tab==="reviewers"&&(
          <>
            <div style={{background:C.card,border:`1px solid ${C.border}`,borderRadius:12,padding:"20px 22px",marginBottom:18}}>
              <div style={{fontSize:11,fontWeight:700,letterSpacing:1.5,textTransform:"uppercase",color:C.sub,marginBottom:16,display:"flex",alignItems:"center",gap:8}}>Reviewer Performance<span style={{flex:1,height:1,background:C.border}}/></div>
              <div style={{overflowX:"auto"}}>
                <table style={{width:"100%",borderCollapse:"collapse",fontSize:13}}>
                  <thead><tr>
                    {["Reviewer","Total","Approved","Rejected","Pending","Approval %","Positive Sentiment %"].map(h=>(
                      <th key={h} style={{padding:"10px 14px",textAlign:"left",color:C.sub,fontWeight:600,borderBottom:`1px solid ${C.border}`,whiteSpace:"nowrap",fontSize:12}}>{h}</th>
                    ))}
                  </tr></thead>
                  <tbody>
                    {reviewerStats.map((r,i)=>(
                      <tr key={i} style={{borderBottom:`1px solid ${C.border}22`}}>
                        <td style={{padding:"10px 14px",fontWeight:600}}>{r.name}</td>
                        <td style={{padding:"10px 14px",fontFamily:"'JetBrains Mono',monospace"}}>{r.total}</td>
                        <td style={{padding:"10px 14px",color:C.green,fontWeight:600,fontFamily:"'JetBrains Mono',monospace"}}>{r.approved}</td>
                        <td style={{padding:"10px 14px",color:C.red,fontWeight:600,fontFamily:"'JetBrains Mono',monospace"}}>{r.rejected}</td>
                        <td style={{padding:"10px 14px",color:C.amber,fontWeight:600,fontFamily:"'JetBrains Mono',monospace"}}>{r.pending}</td>
                        <td style={{padding:"10px 14px"}}>
                          <div style={{display:"flex",alignItems:"center",gap:8}}>
                            <div style={{height:6,width:80,background:C.dim,borderRadius:3}}><div style={{height:"100%",borderRadius:3,background:r.approvalRate>50?C.green:r.approvalRate>25?C.amber:C.red,width:`${r.approvalRate}%`}}/></div>
                            <span style={{fontFamily:"'JetBrains Mono',monospace",fontSize:12,color:r.approvalRate>50?C.green:r.approvalRate>25?C.amber:C.red}}>{r.approvalRate}%</span>
                          </div>
                        </td>
                        <td style={{padding:"10px 14px"}}>
                          <div style={{display:"flex",alignItems:"center",gap:8}}>
                            <div style={{height:6,width:80,background:C.dim,borderRadius:3}}><div style={{height:"100%",borderRadius:3,background:C.teal,width:`${r.positivePct}%`}}/></div>
                            <span style={{fontFamily:"'JetBrains Mono',monospace",fontSize:12,color:C.teal}}>{r.positivePct}%</span>
                          </div>
                        </td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
            </div>
            <div style={{background:C.card,border:`1px solid ${C.border}`,borderRadius:12,padding:"20px 22px"}}>
              <div style={{fontSize:11,fontWeight:700,letterSpacing:1.5,textTransform:"uppercase",color:C.sub,marginBottom:16,display:"flex",alignItems:"center",gap:8}}>Caseload per Reviewer<span style={{flex:1,height:1,background:C.border}}/></div>
              <ResponsiveContainer width="100%" height={220}>
                <BarChart data={reviewerStats}>
                  <CartesianGrid strokeDasharray="3 3" stroke={C.dim}/>
                  <XAxis dataKey="name" tick={{fontSize:11,fill:C.sub}}/><YAxis tick={{fontSize:10,fill:C.sub}}/>
                  <Tooltip content={<DT/>}/>
                  <Bar dataKey="approved" name="Approved" stackId="a" fill={C.green}/>
                  <Bar dataKey="pending" name="Pending" stackId="a" fill={C.amber}/>
                  <Bar dataKey="rejected" name="Rejected" stackId="a" fill={C.red} radius={[4,4,0,0]}/>
                </BarChart>
              </ResponsiveContainer>
            </div>
          </>
        )}

        {/* ══ APPLICANTS ══ */}
        {tab==="applicants"&&(
          <>
            <div style={{overflowX:"auto"}}>
              <table style={{width:"100%",borderCollapse:"collapse",fontSize:12}}>
                <thead><tr>
                  {["S.No","Name","Form Status","Stage","Reviewer","Calls","Call Progress","Overall Sentiment","Themes","Remarks"].map(h=>(
                    <th key={h} style={{padding:"10px 12px",textAlign:"left",color:C.sub,fontWeight:600,borderBottom:`1px solid ${C.border}`,whiteSpace:"nowrap",fontSize:11,background:C.panel}}>{h}</th>
                  ))}
                </tr></thead>
                <tbody>
                  {filtered.slice(0,100).map((r,i)=>(
                    <>
                      <tr key={i} onClick={()=>setExpandedRow(expandedRow===i?null:i)} style={{borderBottom:`1px solid ${C.border}22`,cursor:"pointer",background:expandedRow===i?C.hover:"transparent"}} onMouseEnter={e=>e.currentTarget.style.background=C.hover} onMouseLeave={e=>e.currentTarget.style.background=expandedRow===i?C.hover:"transparent"}>
                        <td style={{padding:"10px 12px",color:C.sub,fontFamily:"'JetBrains Mono',monospace"}}>{r["_id"]}</td>
                        <td style={{padding:"10px 12px",fontWeight:600,whiteSpace:"nowrap"}}>{r["Name"]||"—"}</td>
                        <td style={{padding:"10px 12px"}}><span style={{padding:"3px 10px",borderRadius:20,fontSize:11,fontWeight:600,background:`${statusColors[r["Form Status"]]||C.sub}20`,color:statusColors[r["Form Status"]]||C.sub}}>{r["Form Status"]}</span></td>
                        <td style={{padding:"10px 12px",color:C.sub,fontSize:11}}>{r["_Stage"]}</td>
                        <td style={{padding:"10px 12px",color:C.sub}}>{r["_Reviewer"]}</td>
                        <td style={{padding:"10px 12px",textAlign:"center",fontFamily:"'JetBrains Mono',monospace",color:C.teal,fontWeight:700}}>{r["_CallsMade"]}</td>
                        <td style={{padding:"10px 12px"}}><CallTimeline callSentiments={r["_CallSentiments"]}/></td>
                        <td style={{padding:"10px 12px"}}><SBadge sentiment={r["_OverallSentiment"]}/></td>
                        <td style={{padding:"10px 12px",maxWidth:180}}>
                          <div style={{display:"flex",flexWrap:"wrap",gap:4}}>
                            {r["_CallThemes"].slice(0,2).map((t,ti)=>(
                              <span key={ti} style={{background:`${PAL[ti]}15`,border:`1px solid ${PAL[ti]}30`,borderRadius:4,padding:"1px 6px",fontSize:10,color:PAL[ti]}}>{t.split(" ")[0]}</span>
                            ))}
                          </div>
                        </td>
                        <td style={{padding:"10px 12px",color:C.sub,fontSize:11,maxWidth:200,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>{r["Final Remarks"]||"—"}</td>
                      </tr>
                      {expandedRow===i&&(
                        <tr key={`exp-${i}`} style={{background:C.hover}}>
                          <td colSpan={10} style={{padding:"16px 20px"}}>
                            <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:20}}>
                              <div>
                                <div style={{fontSize:11,color:C.purple,fontWeight:700,letterSpacing:1,marginBottom:10}}>CALL NOTES</div>
                                {["1st Call","2nd Call","3rd Call","4th Call","5th Call"].map(col=>{
                                  const cs=r["_CallSentiments"].find(c=>c.call===col);
                                  if(!cs||cs.sentiment==="no_call") return null;
                                  return(
                                    <div key={col} style={{marginBottom:10,background:C.card,borderRadius:8,padding:"10px 14px",borderLeft:`3px solid ${sentimentColor(cs.sentiment)}`}}>
                                      <div style={{display:"flex",justifyContent:"space-between",marginBottom:4}}>
                                        <span style={{fontSize:11,fontWeight:700,color:C.text}}>{col}</span>
                                        <SBadge sentiment={cs.sentiment}/>
                                      </div>
                                      <div style={{fontSize:12,color:C.sub,lineHeight:1.6}}>{cs.text}</div>
                                    </div>
                                  );
                                })}
                              </div>
                              <div>
                                <div style={{fontSize:11,color:C.orange,fontWeight:700,letterSpacing:1,marginBottom:10}}>FINAL REMARKS & ANALYSIS</div>
                                <div style={{background:C.card,borderRadius:8,padding:"12px 14px",marginBottom:12,borderLeft:`3px solid ${sentimentColor(r["_RemarkSentiment"])}`}}>
                                  <div style={{fontSize:12,color:C.sub,lineHeight:1.6,marginBottom:8}}>{r["Final Remarks"]||"No remarks"}</div>
                                  <SBadge sentiment={r["_RemarkSentiment"]}/>
                                </div>
                                <div style={{fontSize:11,color:C.sub,fontWeight:700,letterSpacing:1,marginBottom:8}}>THEMES DETECTED</div>
                                <div style={{display:"flex",flexWrap:"wrap",gap:6}}>
                                  {[...r["_CallThemes"],...r["_RemarkThemes"]].filter((v,i,a)=>a.indexOf(v)===i).map((t,ti)=>(
                                    <span key={ti} style={{background:`${PAL[ti%PAL.length]}20`,border:`1px solid ${PAL[ti%PAL.length]}40`,borderRadius:6,padding:"4px 10px",fontSize:11,color:PAL[ti%PAL.length],fontWeight:600}}>{t}</span>
                                  ))}
                                  {[...r["_CallThemes"],...r["_RemarkThemes"]].length===0&&<span style={{color:C.sub,fontSize:12}}>No specific themes detected</span>}
                                </div>
                              </div>
                            </div>
                          </td>
                        </tr>
                      )}
                    </>
                  ))}
                </tbody>
              </table>
              {filtered.length>100&&<div style={{textAlign:"center",padding:20,color:C.sub,fontSize:12}}>Showing 100 of {filtered.length} records — use filters to narrow down</div>}
            </div>
          </>
        )}
      </div>
    </div>
  );
}