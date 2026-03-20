/* eslint-disable */
import { useState, useCallback, useMemo } from "react";
import { BarChart, Bar, PieChart, Pie, Cell, AreaChart, Area, XAxis, YAxis, CartesianGrid, Tooltip, ResponsiveContainer, ComposedChart, Line } from "recharts";
import Papa from "papaparse";
import * as XLSX from "xlsx";
import ApplicationTracker from "./ApplicationTracker";

const C={bg:"#0e1117",panel:"#151821",card:"#1a1f2e",border:"#252b3d",hover:"#1f2536",blue:"#2b7de9",teal:"#00c9a7",amber:"#f5a623",red:"#e8445a",purple:"#8b5cf6",green:"#22c55e",text:"#e2e8f0",sub:"#8892a4",dim:"#2d3448",white:"#ffffff"};
const PAL=[C.blue,C.teal,C.amber,C.red,C.purple,C.green,"#fb923c","#f472b6","#38bdf8","#a3e635"];
const CSS=`
@import url('https://fonts.googleapis.com/css2?family=Plus+Jakarta+Sans:wght@400;500;600;700;800&family=JetBrains+Mono:wght@400;500&display=swap');
*,*::before,*::after{box-sizing:border-box;margin:0;padding:0;}
html,body{background:#0e1117;color:#e2e8f0;font-family:'Plus Jakarta Sans',sans-serif;font-size:14px;}
::-webkit-scrollbar{width:5px;height:5px;}
::-webkit-scrollbar-track{background:#0e1117;}
::-webkit-scrollbar-thumb{background:#2d3448;border-radius:3px;}
input[type=file]{display:none;}
@keyframes fadeIn{from{opacity:0;transform:translateY(16px);}to{opacity:1;transform:translateY(0);}}
.fade{animation:fadeIn .5s ease both;}
.mono{font-family:'JetBrains Mono',monospace;}
.badge{display:inline-flex;align-items:center;padding:2px 9px;border-radius:20px;font-size:11px;font-weight:600;}
.chip{background:#2d3448;border:1px solid #252b3d;border-radius:6px;padding:3px 10px;font-size:11px;color:#8892a4;}
.kpi-card{background:#1a1f2e;border:1px solid #252b3d;border-radius:12px;padding:18px 20px;position:relative;overflow:hidden;transition:border-color .2s;}
.kpi-card:hover{border-color:#2b7de960;}
.kpi-card::after{content:'';position:absolute;top:0;left:0;right:0;height:2px;background:var(--accent,#2b7de9);}
.section{background:#1a1f2e;border:1px solid #252b3d;border-radius:12px;padding:20px 22px;margin-bottom:18px;}
.section-hdr{font-size:11px;font-weight:700;letter-spacing:1.5px;text-transform:uppercase;color:#8892a4;margin-bottom:16px;display:flex;align-items:center;gap:8px;}
.section-hdr::after{content:'';flex:1;height:1px;background:#252b3d;}
.tab-btn{background:none;border:none;cursor:pointer;padding:10px 18px;font-size:13px;font-family:'Plus Jakarta Sans',sans-serif;font-weight:600;color:#8892a4;border-bottom:2px solid transparent;transition:all .2s;white-space:nowrap;}
.tab-btn.active{color:#2b7de9;border-bottom-color:#2b7de9;}
.tab-btn:hover{color:#e2e8f0;}
.filter-select{background:#2d3448;border:1px solid #252b3d;color:#e2e8f0;padding:6px 12px;border-radius:7px;font-size:12px;outline:none;cursor:pointer;}
.progress-bar{height:6px;background:#2d3448;border-radius:3px;overflow:hidden;}
.progress-fill{height:100%;border-radius:3px;transition:width .8s ease;}
.data-table{width:100%;border-collapse:collapse;font-size:12px;}
.data-table th{padding:9px 12px;text-align:left;color:#8892a4;font-weight:600;border-bottom:1px solid #252b3d;white-space:nowrap;}
.data-table td{padding:8px 12px;border-bottom:1px solid #252b3d22;color:#e2e8f0;}
.data-table tr:hover td{background:#1f2536;}
`;

function cleanData(rawRows){return rawRows.map((row,idx)=>{const r={};Object.keys(row).forEach(k=>{r[k.trim()]=typeof row[k]==="string"?row[k].trim():row[k];});const status=(r["Application Status"]||"").toLowerCase();if(!r["Application Status"]||r["Application Status"]==="")r["Application Status"]="Unknown";else if(status.includes("approv")||status.includes("accept"))r["Application Status"]="Approved";else if(status.includes("reject")||status.includes("declin"))r["Application Status"]="Rejected";else if(status.includes("pend")||status.includes("review")||status.includes("process"))r["Application Status"]="Pending";else if(status.includes("incomplet"))r["Application Status"]="Incomplete";const g=(r["Gender"]||"").toLowerCase();if(g==="m"||g==="male"||g==="1")r["Gender"]="Male";else if(g==="f"||g==="female"||g==="2")r["Gender"]="Female";else if(!g)r["Gender"]="Not Specified";let age=parseFloat(r["Age"]);if(isNaN(age)||age<=0||age>100){const dob=r["Date of Birth"]||"";if(dob){const parts=dob.split(/[\/\-\.]/);if(parts.length===3){let yr=parseInt(parts[2]);if(yr<100)yr+=(yr>30?1900:2000);const calc=new Date().getFullYear()-yr;if(calc>0&&calc<100){r["Age"]=calc;age=calc;}}}if(isNaN(age)||age<=0)r["Age"]=null;}else{r["Age"]=age;}const a=r["Age"];if(!a)r["_AgeGroup"]="Unknown";else if(a<18)r["_AgeGroup"]="< 18";else if(a<=20)r["_AgeGroup"]="18-20";else if(a<=22)r["_AgeGroup"]="21-22";else if(a<=25)r["_AgeGroup"]="23-25";else r["_AgeGroup"]="25+";["10%","12%","Graduation %"].forEach(col=>{let v=(r[col]||"").toString().replace(/[%,\s]/g,"");const n=parseFloat(v);r[col]=(!isNaN(n)&&n>=0&&n<=100)?n:null;});const scores=[r["Graduation %"],r["10%"],r["12%"]].filter(v=>v!==null);const meanScore=scores.length?scores.reduce((a,b)=>a+b,0)/scores.length:null;r["_AvgScore"]=meanScore?parseFloat(meanScore.toFixed(1)):null;if(!meanScore)r["_Tier"]="No Data";else if(meanScore>=75)r["_Tier"]="Distinction (75+)";else if(meanScore>=60)r["_Tier"]="First Class (60-75)";else if(meanScore>=50)r["_Tier"]="Second Class (50-60)";else r["_Tier"]="Below 50";let inc=(r["Family Income (INR)"]||"").toString().replace(/[,\s\u20B9]/g,"");const incN=parseFloat(inc);r["_Income"]=!isNaN(incN)&&incN>=0?incN:null;if(!r["_Income"])r["_IncomeGroup"]="Not Provided";else if(r["_Income"]<=100000)r["_IncomeGroup"]="<=1L";else if(r["_Income"]<=300000)r["_IncomeGroup"]="1L-3L";else if(r["_Income"]<=600000)r["_IncomeGroup"]="3L-6L";else r["_IncomeGroup"]="6L+";let fee=(r["Course Fee Per Year"]||r["Total Course Fee"]||"").toString().replace(/[,\s\u20B9]/g,"");r["_FeePY"]=!isNaN(parseFloat(fee))?parseFloat(fee):null;let rent=(r["House Rent"]||"").toString().replace(/[,\s\u20B9]/g,"");r["_Rent"]=!isNaN(parseFloat(rent))?parseFloat(rent):null;const cat=(r["Category"]||"").toUpperCase().trim();if(!cat)r["Category"]="Not Specified";else if(cat==="GEN"||cat==="GENERAL"||cat==="UR")r["Category"]="General";else if(cat.startsWith("OBC"))r["Category"]="OBC";else if(cat==="SC"||cat==="SCHEDULED CASTE")r["Category"]="SC";else if(cat==="ST"||cat==="SCHEDULED TRIBE")r["Category"]="ST";else if(cat==="EWS")r["Category"]="EWS";if(r["State"])r["State"]=r["State"].trim().replace(/\b\w/g,c=>c.toUpperCase());if(r["State of College"])r["State of College"]=r["State of College"].trim().replace(/\b\w/g,c=>c.toUpperCase());const hs=(r["Current House Status"]||"").toLowerCase();if(hs.includes("own"))r["Current House Status"]="Owned";else if(hs.includes("rent"))r["Current House Status"]="Rented";else if(hs.includes("gov"))r["Current House Status"]="Govt Quarters";else if(!hs)r["Current House Status"]="Unknown";if(!r["SR NO."]||r["SR NO."]==="")r["SR NO."]=idx+1;r["_idx"]=idx;return r;});}

const countBy=(arr,key,val)=>arr.filter(r=>r[key]===val).length;
const groupBy=(arr,key)=>{const m={};arr.forEach(r=>{const v=r[key]||"Unknown";m[v]=(m[v]||0)+1;});return Object.entries(m).sort((a,b)=>b[1]-a[1]).map(([name,value])=>({name,value}));};
const avg=(arr,key)=>{const vals=arr.map(r=>r[key]).filter(v=>v!==null&&!isNaN(v));return vals.length?vals.reduce((a,b)=>a+b,0)/vals.length:0;};
const pct=(n,t)=>t?((n/t)*100).toFixed(1):"0.0";
const fmt=n=>n>=10000000?(n/10000000).toFixed(1)+"Cr":n>=100000?(n/100000).toFixed(1)+"L":n>=1000?(n/1000).toFixed(1)+"K":n?.toLocaleString()||"--";
const DT=({active,payload,label})=>{if(!active||!payload?.length)return null;return(<div style={{background:"#151821",border:"1px solid #252b3d",borderRadius:8,padding:"10px 14px",fontFamily:"'JetBrains Mono',monospace",fontSize:12}}>{label!==undefined&&<div style={{color:"#8892a4",marginBottom:6,fontSize:11}}>{label}</div>}{payload.map((p,i)=>(<div key={i} style={{color:p.color||"#00c9a7",marginBottom:2}}>{p.name}: <strong style={{color:"#fff"}}>{typeof p.value==="number"?p.value.toLocaleString():p.value}</strong></div>))}</div>);};

function PasswordScreen({onUnlock}){const [pw,setPw]=useState("");const [err,setErr]=useState(false);const attempt=()=>{if(pw==="Foundation")onUnlock();else{setErr(true);setTimeout(()=>setErr(false),2000);}};return(<div style={{minHeight:"100vh",background:"#0e1117",display:"flex",alignItems:"center",justifyContent:"center"}}><div style={{position:"absolute",inset:0,backgroundImage:"radial-gradient(#2b7de915 1px,transparent 1px)",backgroundSize:"32px 32px"}}/><div className="fade" style={{position:"relative",zIndex:1,textAlign:"center",width:420}}><img src="https://adimaheshsevafoundation.org/images/Adimahesh _Logo.png" alt="AdiMahesh Seva Foundation" style={{width:200,marginBottom:20,background:"white",borderRadius:12,padding:"10px 16px",filter:"drop-shadow(0 0 16px rgba(43,125,233,0.25))"}}/><div style={{fontSize:30,fontWeight:800,letterSpacing:-1,marginBottom:4}}><span style={{color:"#2b7de9"}}>Scholar</span><span style={{color:"#00c9a7"}}>Lens</span></div><div style={{fontSize:11,color:"#8892a4",fontFamily:"'JetBrains Mono',monospace",letterSpacing:1.5,marginBottom:6}}>SCHOLARSHIP ANALYTICS PLATFORM</div><div style={{fontSize:11,color:"#8892a4",marginBottom:36,display:"flex",alignItems:"center",justifyContent:"center",gap:6}}><span style={{color:"#2d3448"}}>--</span><span>Powered by <a href="https://adimaheshsevafoundation.org" target="_blank" rel="noreferrer" style={{color:"#00c9a7",textDecoration:"none",fontWeight:600}}>AdiMahesh Seva Foundation</a></span><span style={{color:"#2d3448"}}>--</span></div><div style={{background:"#1a1f2e",border:`1px solid ${err?"#e8445a":"#252b3d"}`,borderRadius:16,padding:"32px 36px",boxShadow:"0 24px 64px #00000060",transition:"border-color .3s"}}><div style={{fontSize:11,color:"#8892a4",textAlign:"left",marginBottom:10,fontWeight:600,letterSpacing:1}}>ACCESS CODE</div><input type="password" value={pw} onChange={e=>setPw(e.target.value)} onKeyDown={e=>e.key==="Enter"&&attempt()} placeholder="Enter password" autoFocus style={{width:"100%",padding:"13px 16px",background:"#151821",border:`1px solid ${err?"#e8445a":"#2d3448"}`,borderRadius:9,color:"#e2e8f0",fontSize:15,fontFamily:"'JetBrains Mono',monospace",outline:"none",letterSpacing:pw?4:0,marginBottom:8,transition:"border-color .2s"}}/>{err&&<div style={{color:"#e8445a",fontSize:12,marginBottom:10}}>Incorrect password</div>}<button onClick={attempt} style={{width:"100%",padding:"13px",background:"linear-gradient(135deg,#2b7de9,#00c9a7)",border:"none",borderRadius:9,color:"#fff",fontWeight:700,fontSize:14,cursor:"pointer",fontFamily:"'Plus Jakarta Sans',sans-serif",marginTop:4}}>UNLOCK</button></div><div style={{marginTop:20,fontSize:11,color:"#2d3448"}}>AdiMahesh Seva Foundation - Restricted Access</div></div></div>);}

function ModeSelector({onSelectMode}){return(<div style={{minHeight:"100vh",background:"#0e1117",display:"flex",flexDirection:"column",alignItems:"center",justifyContent:"center",padding:24}}><div style={{position:"absolute",inset:0,backgroundImage:"radial-gradient(#2b7de912 1px,transparent 1px)",backgroundSize:"32px 32px"}}/><div className="fade" style={{position:"relative",zIndex:1,textAlign:"center",maxWidth:700,width:"100%"}}><img src="https://adimaheshsevafoundation.org/images/Adimahesh _Logo.png" alt="Logo" style={{width:120,marginBottom:20,background:"white",borderRadius:10,padding:"8px 14px"}}/><div style={{fontSize:32,fontWeight:800,letterSpacing:-1,marginBottom:6}}><span style={{color:"#2b7de9"}}>Scholar</span><span style={{color:"#00c9a7"}}>Lens</span></div><div style={{fontSize:12,color:"#8892a4",fontFamily:"'JetBrains Mono',monospace",marginBottom:48,letterSpacing:1.5}}>SELECT DASHBOARD MODE</div><div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:20}}><div onClick={()=>onSelectMode(1)} style={{background:"#1a1f2e",border:"1px solid #252b3d",borderRadius:16,padding:"32px 28px",cursor:"pointer",transition:"all .25s",textAlign:"left",position:"relative",overflow:"hidden"}} onMouseEnter={e=>{e.currentTarget.style.borderColor="#2b7de9";e.currentTarget.style.transform="translateY(-3px)";}} onMouseLeave={e=>{e.currentTarget.style.borderColor="#252b3d";e.currentTarget.style.transform="translateY(0)";}}><div style={{position:"absolute",top:0,left:0,right:0,height:3,background:"linear-gradient(90deg,#2b7de9,#00c9a7)"}}/><div style={{fontSize:36,marginBottom:16}}>📊</div><div style={{fontSize:20,fontWeight:800,marginBottom:8,color:"#fff"}}>Mode 1</div><div style={{fontSize:15,fontWeight:700,color:"#00c9a7",marginBottom:12}}>Scholarship Analytics</div><div style={{fontSize:13,color:"#8892a4",lineHeight:1.7,marginBottom:20}}>Deep analysis of student applications — academics, income, demographics, geography and approval insights.</div><div style={{display:"flex",flexWrap:"wrap",gap:6}}>{["36-column CSV","Academic tiers","Income analysis","Geographic view","Data cleaning"].map(f=>(<span style={{background:"#2b7de915",border:"1px solid #2b7de930",borderRadius:5,padding:"3px 9px",fontSize:11,color:"#2b7de9"}} key={f}>{f}</span>))}</div></div><div onClick={()=>onSelectMode(2)} style={{background:"#1a1f2e",border:"1px solid #252b3d",borderRadius:16,padding:"32px 28px",cursor:"pointer",transition:"all .25s",textAlign:"left",position:"relative",overflow:"hidden"}} onMouseEnter={e=>{e.currentTarget.style.borderColor="#8b5cf6";e.currentTarget.style.transform="translateY(-3px)";}} onMouseLeave={e=>{e.currentTarget.style.borderColor="#252b3d";e.currentTarget.style.transform="translateY(0)";}}><div style={{position:"absolute",top:0,left:0,right:0,height:3,background:"linear-gradient(90deg,#8b5cf6,#fb923c)"}}/><div style={{fontSize:36,marginBottom:16}}>📞</div><div style={{fontSize:20,fontWeight:800,marginBottom:8,color:"#fff"}}>Mode 2</div><div style={{fontSize:15,fontWeight:700,color:"#8b5cf6",marginBottom:12}}>Application Tracker + NLP</div><div style={{fontSize:13,color:"#8892a4",lineHeight:1.7,marginBottom:20}}>Track follow-up calls, reviewer workload, stage pipeline and analyze call notes using NLP sentiment and theme extraction.</div><div style={{display:"flex",flexWrap:"wrap",gap:6}}>{["Call NLP","Sentiment analysis","Theme extraction","Reviewer stats","Stage funnel"].map(f=>(<span style={{background:"#8b5cf615",border:"1px solid #8b5cf630",borderRadius:5,padding:"3px 9px",fontSize:11,color:"#8b5cf6"}} key={f}>{f}</span>))}</div></div></div></div></div>);}

function KPI({label,value,sub,accent,delay=0}){return(<div className="kpi-card fade" style={{"--accent":accent||C.blue,animationDelay:`${delay}ms`}}><div style={{fontSize:11,color:C.sub,fontWeight:600,letterSpacing:.8,textTransform:"uppercase",marginBottom:10}}>{label}</div><div style={{fontSize:28,fontWeight:800,color:accent||C.white,lineHeight:1}}>{value}</div>{sub&&<div style={{fontSize:12,color:C.sub,marginTop:6}}>{sub}</div>}</div>);}
function HBar({name,value,total,color,rank}){const p=total?(value/total)*100:0;return(<div style={{marginBottom:10}}><div style={{display:"flex",justifyContent:"space-between",marginBottom:4}}><span style={{fontSize:12,color:C.text,display:"flex",alignItems:"center",gap:6}}>{rank&&<span style={{fontSize:10,color:C.sub,width:16}}>#{rank}</span>}{name}</span><span style={{fontSize:12,fontWeight:700,color:color||C.blue,fontFamily:"'JetBrains Mono',monospace"}}>{value} <span style={{color:C.sub,fontWeight:400}}>({p.toFixed(1)}%)</span></span></div><div className="progress-bar"><div className="progress-fill" style={{width:`${p}%`,background:color||C.blue}}/></div></div>);}

function UploadScreen({onFile,onBack}){const [drag,setDrag]=useState(false);const onDrop=useCallback(e=>{e.preventDefault();setDrag(false);const f=e.dataTransfer.files[0];if(f)onFile(f);},[onFile]);return(<div style={{minHeight:"100vh",background:C.bg,display:"flex",flexDirection:"column",alignItems:"center",justifyContent:"center",padding:24}}><div style={{position:"absolute",inset:0,backgroundImage:`radial-gradient(${C.blue}12 1px,transparent 1px)`,backgroundSize:"32px 32px"}}/><div className="fade" style={{position:"relative",zIndex:1,textAlign:"center",maxWidth:600,width:"100%"}}><button onClick={onBack} style={{position:"absolute",top:-60,left:0,background:"none",border:`1px solid ${C.border}`,color:C.sub,padding:"6px 14px",borderRadius:7,cursor:"pointer",fontSize:12}}>Back to Modes</button><div style={{fontSize:38,fontWeight:800,letterSpacing:-1.5,marginBottom:6}}><span style={{color:C.blue}}>Scholar</span><span style={{color:C.teal}}>Lens</span></div><p style={{color:C.sub,fontSize:13,fontFamily:"'JetBrains Mono',monospace",marginBottom:48,letterSpacing:1}}>MODE 1 - SCHOLARSHIP ANALYTICS</p><div onDragOver={e=>{e.preventDefault();setDrag(true);}} onDragLeave={()=>setDrag(false)} onDrop={onDrop} style={{border:`2px dashed ${drag?C.teal:C.dim}`,borderRadius:20,background:drag?`${C.teal}08`:C.card,padding:"56px 48px",transition:"all .3s"}}><div style={{fontSize:48,marginBottom:16}}>📂</div><div style={{fontSize:20,fontWeight:700,marginBottom:8}}>Drop your scholarship CSV here</div><div style={{color:C.sub,fontSize:13,marginBottom:28}}>Auto-cleans all 36 columns</div><label style={{background:`linear-gradient(135deg,${C.blue},${C.teal})`,color:C.white,padding:"12px 32px",borderRadius:10,cursor:"pointer",fontWeight:700,fontSize:14,display:"inline-block"}}>CHOOSE FILE<input type="file" accept=".csv,.xlsx,.xls" onChange={e=>e.target.files[0]&&onFile(e.target.files[0])}/></label></div></div></div>);}

function Dashboard({data,fileName,onReset,onBack}){
  const [tab,setTab]=useState("overview");
  const [fS,setFS]=useState("All");const [fG,setFG]=useState("All");const [fC,setFC]=useState("All");
  const filtered=useMemo(()=>{let d=data;if(fS!=="All")d=d.filter(r=>r["Application Status"]===fS);if(fG!=="All")d=d.filter(r=>r["Gender"]===fG);if(fC!=="All")d=d.filter(r=>r["Category"]===fC);return d;},[data,fS,fG,fC]);
  const total=filtered.length;
  const statuses=useMemo(()=>groupBy(filtered,"Application Status"),[filtered]);
  const genders=useMemo(()=>groupBy(filtered,"Gender"),[filtered]);
  const categories=useMemo(()=>groupBy(filtered,"Category"),[filtered]);
  const ageGroups=useMemo(()=>{const order=["< 18","18-20","21-22","23-25","25+","Unknown"];const g=groupBy(filtered,"_AgeGroup");return order.map(o=>g.find(x=>x.name===o)||{name:o,value:0}).filter(x=>x.value>0);},[filtered]);
  const incomeGroups=useMemo(()=>groupBy(filtered,"_IncomeGroup"),[filtered]);
  const tiers=useMemo(()=>groupBy(filtered,"_Tier"),[filtered]);
  const courses=useMemo(()=>groupBy(filtered,"Course Name").slice(0,8),[filtered]);
  const houseStatus=useMemo(()=>groupBy(filtered,"Current House Status"),[filtered]);
  const streams=useMemo(()=>groupBy(filtered,"Stream").slice(0,8),[filtered]);
  const colleges=useMemo(()=>groupBy(filtered,"Name of College/Institute").slice(0,8),[filtered]);
  const religions=useMemo(()=>groupBy(filtered,"Religion").slice(0,8),[filtered]);
  const approved=countBy(filtered,"Application Status","Approved");
  const rejected=countBy(filtered,"Application Status","Rejected");
  const pending=countBy(filtered,"Application Status","Pending");
  const avgGrad=avg(filtered,"Graduation %");const avg10=avg(filtered,"10%");const avg12=avg(filtered,"12%");
  const avgAge=avg(filtered,"Age");const avgIncome=avg(filtered,"_Income");const avgFee=avg(filtered,"_FeePY");
  const missingGrad=filtered.filter(r=>r["Graduation %"]===null).length;
  const missingIncome=filtered.filter(r=>r["_Income"]===null).length;
  const missingAge=filtered.filter(r=>r["Age"]===null).length;
  const approvedAvgGrad=avg(filtered.filter(r=>r["Application Status"]==="Approved"),"Graduation %");
  const rejectedAvgGrad=avg(filtered.filter(r=>r["Application Status"]==="Rejected"),"Graduation %");
  const genderStatus=["Male","Female","Not Specified"].map(g=>{const s=filtered.filter(r=>r["Gender"]===g);return{name:g,Approved:countBy(s,"Application Status","Approved"),Rejected:countBy(s,"Application Status","Rejected"),Pending:countBy(s,"Application Status","Pending"),total:s.length};}).filter(x=>x.total>0);
  const monthlyTrend=useMemo(()=>{const m={};filtered.forEach(r=>{const d=r["Received Date"]||"";const parts=d.split(/[\/\-\.]/);if(parts.length>=2){const mon=parts[1]?.padStart(2,"0");const yr=parts[2]?.length===4?parts[2]:parts[0]?.length===4?parts[0]:"2024";const key=`${yr}-${mon}`;if(key&&!key.includes("undefined"))m[key]=(m[key]||0)+1;}});return Object.entries(m).sort().map(([name,value])=>({name:name.replace(/^20/,""),value}));},[filtered]);
  const incomeAcademic=useMemo(()=>["<=1L","1L-3L","3L-6L","6L+","Not Provided"].map(g=>{const s=filtered.filter(r=>r["_IncomeGroup"]===g);return{name:g,"Avg Grade":parseFloat(avg(s,"_AvgScore").toFixed(1)),count:s.length};}).filter(x=>x.count>0),[filtered]);
  const cleanReport=useMemo(()=>[{field:"Age",missing:missingAge},{field:"Graduation %",missing:missingGrad},{field:"Family Income",missing:missingIncome},{field:"10th %",missing:filtered.filter(r=>r["10%"]===null).length},{field:"12th %",missing:filtered.filter(r=>r["12%"]===null).length},{field:"Course Fee",missing:filtered.filter(r=>r["_FeePY"]===null).length},].map(r=>({...r,total,pct:pct(r.missing,total)})),[filtered,total,missingAge,missingGrad,missingIncome]);
  const statusColors={Approved:C.green,Rejected:C.red,Pending:C.amber,Incomplete:C.purple,Unknown:C.sub};
  const TABS=["overview","students","academic","financial","geography","data quality"];
  const allStatuses=["All",...[...new Set(data.map(r=>r["Application Status"]))].filter(Boolean)];
  const allGenders=["All",...[...new Set(data.map(r=>r["Gender"]))].filter(Boolean)];
  const allCats=["All",...[...new Set(data.map(r=>r["Category"]))].filter(Boolean)];
  return(<div style={{minHeight:"100vh",background:C.bg}}>
    <div style={{background:C.panel,borderBottom:`1px solid ${C.border}`,padding:"0 24px",display:"flex",alignItems:"center",gap:16,position:"sticky",top:0,zIndex:100}}>
      <button onClick={onBack} style={{background:"none",border:`1px solid ${C.border}`,color:C.sub,padding:"6px 14px",borderRadius:7,cursor:"pointer",fontSize:12,whiteSpace:"nowrap",flexShrink:0}}>Modes</button>
      <div style={{fontWeight:800,fontSize:18,paddingBlock:14,whiteSpace:"nowrap"}}><span style={{color:C.blue}}>Scholar</span><span style={{color:C.teal}}>Lens</span></div>
      <div style={{width:1,height:28,background:C.border}}/>
      <div style={{fontSize:12,color:C.sub,fontFamily:"'JetBrains Mono',monospace",flex:1,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>{fileName}</div>
      <span className="chip">{total.toLocaleString()} records</span>
      <button onClick={onReset} style={{background:"none",border:`1px solid ${C.border}`,color:C.sub,padding:"6px 14px",borderRadius:7,cursor:"pointer",fontSize:12,whiteSpace:"nowrap"}}>New File</button>
    </div>
    <div style={{background:C.panel,borderBottom:`1px solid ${C.border}`,padding:"10px 24px",display:"flex",gap:10,alignItems:"center",flexWrap:"wrap"}}>
      <span style={{fontSize:11,color:C.sub,fontWeight:600,letterSpacing:1}}>FILTERS</span>
      {[["Status",allStatuses,fS,setFS],["Gender",allGenders,fG,setFG],["Category",allCats,fC,setFC]].map(([label,opts,val,set])=>(<select key={label} className="filter-select" value={val} onChange={e=>set(e.target.value)}>{opts.map(o=><option key={o}>{o}</option>)}</select>))}
      {(fS!=="All"||fG!=="All"||fC!=="All")&&(<button onClick={()=>{setFS("All");setFG("All");setFC("All");}} style={{background:"none",border:`1px solid ${C.red}60`,color:C.red,padding:"5px 12px",borderRadius:6,cursor:"pointer",fontSize:11}}>Clear</button>)}
    </div>
    <div style={{background:C.panel,borderBottom:`1px solid ${C.border}`,padding:"0 24px",display:"flex",overflowX:"auto"}}>
      {TABS.map(t=><button key={t} className={`tab-btn${tab===t?" active":""}`} onClick={()=>setTab(t)}>{t.charAt(0).toUpperCase()+t.slice(1)}</button>)}
    </div>
    <div style={{padding:"24px",maxWidth:1400,margin:"0 auto"}}>
      {tab==="overview"&&(<>
        <div style={{display:"grid",gridTemplateColumns:"repeat(auto-fit,minmax(170px,1fr))",gap:14,marginBottom:20}}>
          <KPI label="Total Applications" value={total.toLocaleString()} accent={C.blue} delay={0}/>
          <KPI label="Approved" value={approved} sub={`${pct(approved,total)}% rate`} accent={C.green} delay={80}/>
          <KPI label="Rejected" value={rejected} sub={`${pct(rejected,total)}% rate`} accent={C.red} delay={160}/>
          <KPI label="Pending" value={pending} sub={`${pct(pending,total)}% in queue`} accent={C.amber} delay={240}/>
          <KPI label="Avg Graduation %" value={avgGrad?avgGrad.toFixed(1)+"%":"--"} accent={C.teal} delay={320}/>
          <KPI label="Avg Age" value={avgAge?avgAge.toFixed(1):"--"} sub="years" accent={C.purple} delay={400}/>
        </div>
        <div style={{display:"grid",gridTemplateColumns:"1fr 1fr 1fr",gap:16,marginBottom:18}}>
          <div className="section fade"><div className="section-hdr">Application Status</div><ResponsiveContainer width="100%" height={180}><PieChart><Pie data={statuses} dataKey="value" nameKey="name" cx="50%" cy="50%" outerRadius={75} innerRadius={38}>{statuses.map((s,i)=><Cell key={i} fill={statusColors[s.name]||PAL[i]}/>)}</Pie><Tooltip content={<DT/>}/></PieChart></ResponsiveContainer>{statuses.map((s,i)=>(<div key={i} style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:6}}><div style={{display:"flex",alignItems:"center",gap:6}}><div style={{width:8,height:8,borderRadius:"50%",background:statusColors[s.name]||PAL[i]}}/><span style={{fontSize:12}}>{s.name}</span></div><span style={{fontFamily:"'JetBrains Mono',monospace",fontSize:12,color:statusColors[s.name]||PAL[i],fontWeight:700}}>{s.value} ({pct(s.value,total)}%)</span></div>))}</div>
          <div className="section fade" style={{animationDelay:"100ms"}}><div className="section-hdr">Gender</div><ResponsiveContainer width="100%" height={180}><PieChart><Pie data={genders} dataKey="value" nameKey="name" cx="50%" cy="50%" outerRadius={75} innerRadius={38}>{genders.map((_,i)=><Cell key={i} fill={[C.blue,C.teal,C.sub][i]||PAL[i]}/>)}</Pie><Tooltip content={<DT/>}/></PieChart></ResponsiveContainer>{genders.map((g,i)=>(<div key={i} style={{display:"flex",justifyContent:"space-between",marginBottom:6}}><div style={{display:"flex",alignItems:"center",gap:6}}><div style={{width:8,height:8,borderRadius:"50%",background:[C.blue,C.teal,C.sub][i]||PAL[i]}}/><span style={{fontSize:12}}>{g.name}</span></div><span style={{fontSize:12,fontFamily:"'JetBrains Mono',monospace",color:[C.blue,C.teal,C.sub][i]||PAL[i],fontWeight:700}}>{g.value} ({pct(g.value,total)}%)</span></div>))}</div>
          <div className="section fade" style={{animationDelay:"200ms"}}><div className="section-hdr">Category</div>{categories.map((c,i)=><HBar key={i} name={c.name} value={c.value} total={total} color={PAL[i]} rank={i+1}/>)}</div>
        </div>
        {monthlyTrend.length>1&&(<div className="section fade"><div className="section-hdr">Monthly Trend</div><ResponsiveContainer width="100%" height={200}><AreaChart data={monthlyTrend}><defs><linearGradient id="bg1" x1="0" y1="0" x2="0" y2="1"><stop offset="5%" stopColor={C.blue} stopOpacity={.35}/><stop offset="95%" stopColor={C.blue} stopOpacity={0}/></linearGradient></defs><CartesianGrid strokeDasharray="3 3" stroke={C.dim}/><XAxis dataKey="name" tick={{fontSize:10,fill:C.sub}}/><YAxis tick={{fontSize:10,fill:C.sub}}/><Tooltip content={<DT/>}/><Area type="monotone" dataKey="value" stroke={C.blue} fill="url(#bg1)" strokeWidth={2} dot={{r:3,fill:C.blue}} name="Applications"/></AreaChart></ResponsiveContainer></div>)}
        <div className="section fade"><div className="section-hdr">Approval by Gender</div><ResponsiveContainer width="100%" height={200}><BarChart data={genderStatus} layout="vertical"><CartesianGrid strokeDasharray="3 3" stroke={C.dim} horizontal={false}/><XAxis type="number" tick={{fontSize:10,fill:C.sub}}/><YAxis dataKey="name" type="category" tick={{fontSize:12,fill:C.text}} width={90}/><Tooltip content={<DT/>}/><Bar dataKey="Approved" stackId="a" fill={C.green}/><Bar dataKey="Pending" stackId="a" fill={C.amber}/><Bar dataKey="Rejected" stackId="a" fill={C.red} radius={[0,4,4,0]}/></BarChart></ResponsiveContainer></div>
      </>)}
      {tab==="students"&&(<>
        <div style={{display:"grid",gridTemplateColumns:"repeat(auto-fit,minmax(170px,1fr))",gap:14,marginBottom:20}}>
          <KPI label="Avg Age" value={avgAge?avgAge.toFixed(1):"--"} sub="years" accent={C.purple}/>
          <KPI label="Missing Age" value={missingAge} sub={`${pct(missingAge,total)}%`} accent={C.amber}/>
          <KPI label="Male" value={countBy(filtered,"Gender","Male")} sub={`${pct(countBy(filtered,"Gender","Male"),total)}%`} accent={C.blue}/>
          <KPI label="Female" value={countBy(filtered,"Gender","Female")} sub={`${pct(countBy(filtered,"Gender","Female"),total)}%`} accent={C.teal}/>
          <KPI label="Rented" value={countBy(filtered,"Current House Status","Rented")} accent={C.amber}/>
          <KPI label="Owned" value={countBy(filtered,"Current House Status","Owned")} accent={C.green}/>
        </div>
        <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:16,marginBottom:18}}>
          <div className="section fade"><div className="section-hdr">Age Groups</div><ResponsiveContainer width="100%" height={220}><BarChart data={ageGroups}><CartesianGrid strokeDasharray="3 3" stroke={C.dim}/><XAxis dataKey="name" tick={{fontSize:11,fill:C.sub}}/><YAxis tick={{fontSize:10,fill:C.sub}}/><Tooltip content={<DT/>}/><Bar dataKey="value" name="Students" radius={[4,4,0,0]}>{ageGroups.map((_,i)=><Cell key={i} fill={PAL[i]}/>)}</Bar></BarChart></ResponsiveContainer></div>
          <div className="section fade"><div className="section-hdr">Housing</div><ResponsiveContainer width="100%" height={180}><PieChart><Pie data={houseStatus} dataKey="value" nameKey="name" cx="50%" cy="50%" outerRadius={80} innerRadius={38}>{houseStatus.map((_,i)=><Cell key={i} fill={PAL[i]}/>)}</Pie><Tooltip content={<DT/>}/></PieChart></ResponsiveContainer>{houseStatus.map((h,i)=>(<div key={i} style={{display:"flex",justifyContent:"space-between",fontSize:12,marginBottom:5}}><div style={{display:"flex",alignItems:"center",gap:6}}><div style={{width:8,height:8,borderRadius:"50%",background:PAL[i]}}/>{h.name}</div><span style={{color:PAL[i],fontWeight:700,fontFamily:"'JetBrains Mono',monospace"}}>{h.value} ({pct(h.value,total)}%)</span></div>))}</div>
        </div>
        <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:16}}>
          <div className="section fade"><div className="section-hdr">Religion</div>{religions.map((r,i)=><HBar key={i} name={r.name} value={r.value} total={total} color={PAL[i]} rank={i+1}/>)}</div>
          <div className="section fade"><div className="section-hdr">Stream</div>{streams.map((s,i)=><HBar key={i} name={s.name} value={s.value} total={total} color={PAL[i]} rank={i+1}/>)}</div>
        </div>
      </>)}
      {tab==="academic"&&(<>
        <div style={{display:"grid",gridTemplateColumns:"repeat(auto-fit,minmax(170px,1fr))",gap:14,marginBottom:20}}>
          <KPI label="Avg 10th %" value={avg10?avg10.toFixed(1)+"%":"--"} accent={C.teal}/>
          <KPI label="Avg 12th %" value={avg12?avg12.toFixed(1)+"%":"--"} accent={C.blue}/>
          <KPI label="Avg Graduation %" value={avgGrad?avgGrad.toFixed(1)+"%":"--"} accent={C.purple}/>
          <KPI label="Approved Avg" value={approvedAvgGrad?approvedAvgGrad.toFixed(1)+"%":"--"} accent={C.green} sub="approved only"/>
          <KPI label="Rejected Avg" value={rejectedAvgGrad?rejectedAvgGrad.toFixed(1)+"%":"--"} accent={C.red} sub="rejected only"/>
          <KPI label="Missing Grad %" value={missingGrad} sub={`${pct(missingGrad,total)}%`} accent={C.amber}/>
        </div>
        <div className="section fade" style={{marginBottom:18}}><div className="section-hdr">Score Progression: 10th to 12th to Graduation</div><ResponsiveContainer width="100%" height={220}><BarChart data={[{stage:"10th",score:parseFloat(avg10.toFixed(1))},{stage:"12th",score:parseFloat(avg12.toFixed(1))},{stage:"Graduation",score:parseFloat(avgGrad.toFixed(1))}]}><CartesianGrid strokeDasharray="3 3" stroke={C.dim}/><XAxis dataKey="stage" tick={{fontSize:13,fill:C.text}}/><YAxis domain={[0,100]} tick={{fontSize:10,fill:C.sub}}/><Tooltip content={<DT/>}/><Bar dataKey="score" name="Avg Score" radius={[6,6,0,0]}><Cell fill={C.teal}/><Cell fill={C.blue}/><Cell fill={C.purple}/></Bar></BarChart></ResponsiveContainer></div>
        <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:16,marginBottom:18}}>
          <div className="section fade"><div className="section-hdr">Performance Tier</div>{tiers.map((t,i)=><HBar key={i} name={t.name} value={t.value} total={total} color={PAL[i]} rank={i+1}/>)}</div>
          <div className="section fade"><div className="section-hdr">Top Courses</div>{courses.map((c,i)=><HBar key={i} name={c.name||"Unknown"} value={c.value} total={total} color={PAL[i]} rank={i+1}/>)}</div>
        </div>
        <div className="section fade" style={{marginBottom:18}}><div className="section-hdr">Grade Tier vs Approval Rate</div><ResponsiveContainer width="100%" height={220}><ComposedChart data={tiers.map(t=>{const s=filtered.filter(r=>r["_Tier"]===t.name);const app=countBy(s,"Application Status","Approved");return{name:t.name,Total:t.value,Approved:app,"Approval%":s.length?parseFloat(((app/s.length)*100).toFixed(1)):0};})}><CartesianGrid strokeDasharray="3 3" stroke={C.dim}/><XAxis dataKey="name" tick={{fontSize:10,fill:C.sub}}/><YAxis yAxisId="left" tick={{fontSize:10,fill:C.sub}}/><YAxis yAxisId="right" orientation="right" domain={[0,100]} tick={{fontSize:10,fill:C.sub}} unit="%"/><Tooltip content={<DT/>}/><Bar yAxisId="left" dataKey="Total" fill={C.blue} radius={[4,4,0,0]} opacity={0.7} name="Total"/><Bar yAxisId="left" dataKey="Approved" fill={C.green} radius={[4,4,0,0]} name="Approved"/><Line yAxisId="right" type="monotone" dataKey="Approval%" stroke={C.amber} strokeWidth={2} dot={{r:4,fill:C.amber}} name="Approval%"/></ComposedChart></ResponsiveContainer></div>
        <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:16}}>
          <div className="section fade"><div className="section-hdr">Top Colleges</div>{colleges.map((c,i)=><HBar key={i} name={(c.name||"Unknown").slice(0,32)} value={c.value} total={total} color={PAL[i]} rank={i+1}/>)}</div>
          <div className="section fade"><div className="section-hdr">Income vs Academic Score</div><ResponsiveContainer width="100%" height={200}><BarChart data={incomeAcademic}><CartesianGrid strokeDasharray="3 3" stroke={C.dim}/><XAxis dataKey="name" tick={{fontSize:10,fill:C.sub}}/><YAxis domain={[0,100]} tick={{fontSize:10,fill:C.sub}}/><Tooltip content={<DT/>}/><Bar dataKey="Avg Grade" fill={C.purple} radius={[4,4,0,0]} name="Avg Score"/></BarChart></ResponsiveContainer></div>
        </div>
      </>)}
      {tab==="financial"&&(<>
        <div style={{display:"grid",gridTemplateColumns:"repeat(auto-fit,minmax(170px,1fr))",gap:14,marginBottom:20}}>
          <KPI label="Avg Family Income" value={avgIncome?fmt(avgIncome):"--"} accent={C.green} sub="INR/year"/>
          <KPI label="Avg Course Fee" value={avgFee?fmt(avgFee):"--"} accent={C.blue}/>
          <KPI label="Missing Income" value={missingIncome} sub={`${pct(missingIncome,total)}%`} accent={C.amber}/>
          <KPI label="Renting" value={countBy(filtered,"Current House Status","Rented")} accent={C.red}/>
          <KPI label="Below 1L" value={filtered.filter(r=>r["_IncomeGroup"]==="<=1L").length} accent={C.purple} sub="most vulnerable"/>
          <KPI label="Avg Rent" value={avg(filtered,"_Rent")?fmt(avg(filtered,"_Rent")):"--"} accent={C.teal} sub="INR/month"/>
        </div>
        <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:16,marginBottom:18}}>
          <div className="section fade"><div className="section-hdr">Income Distribution</div><ResponsiveContainer width="100%" height={220}><BarChart data={incomeGroups}><CartesianGrid strokeDasharray="3 3" stroke={C.dim}/><XAxis dataKey="name" tick={{fontSize:11,fill:C.sub}}/><YAxis tick={{fontSize:10,fill:C.sub}}/><Tooltip content={<DT/>}/><Bar dataKey="value" name="Applicants" radius={[4,4,0,0]}>{incomeGroups.map((_,i)=><Cell key={i} fill={PAL[i]}/>)}</Bar></BarChart></ResponsiveContainer></div>
          <div className="section fade"><div className="section-hdr">Approval Rate by Income</div><ResponsiveContainer width="100%" height={220}><BarChart data={incomeGroups.map(g=>{const s=filtered.filter(r=>r["_IncomeGroup"]===g.name);const app=countBy(s,"Application Status","Approved");return{name:g.name,"Approval%":s.length?parseFloat(((app/s.length)*100).toFixed(1)):0};})}><CartesianGrid strokeDasharray="3 3" stroke={C.dim}/><XAxis dataKey="name" tick={{fontSize:11,fill:C.sub}}/><YAxis domain={[0,100]} tick={{fontSize:10,fill:C.sub}} unit="%"/><Tooltip content={<DT/>}/><Bar dataKey="Approval%" fill={C.green} radius={[4,4,0,0]}/></BarChart></ResponsiveContainer></div>
        </div>
        <div className="section fade"><div className="section-hdr">Income Group Deep-Dive</div><table className="data-table"><thead><tr><th>Group</th><th>Total</th><th>Approved</th><th>Rejected</th><th>Pending</th><th>Approval %</th><th>Avg Grad</th></tr></thead><tbody>{incomeGroups.map((g,i)=>{const s=filtered.filter(r=>r["_IncomeGroup"]===g.name);const app=countBy(s,"Application Status","Approved");const rej=countBy(s,"Application Status","Rejected");const pen=countBy(s,"Application Status","Pending");const ap=s.length?((app/s.length)*100).toFixed(1):0;const ag=avg(s,"Graduation %");return(<tr key={i}><td><span className="badge" style={{background:`${PAL[i]}20`,color:PAL[i]}}>{g.name}</span></td><td className="mono">{g.value}</td><td style={{color:C.green,fontWeight:600}} className="mono">{app}</td><td style={{color:C.red,fontWeight:600}} className="mono">{rej}</td><td style={{color:C.amber,fontWeight:600}} className="mono">{pen}</td><td className="mono" style={{color:ap>50?C.green:ap>25?C.amber:C.red}}>{ap}%</td><td className="mono">{ag?ag.toFixed(1)+"%":"--"}</td></tr>);})}</tbody></table></div>
      </>)}
      {tab==="geography"&&(<>
        <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:16,marginBottom:18}}>
          <div className="section fade"><div className="section-hdr">Top States (Students)</div>{groupBy(filtered,"State").slice(0,10).map((s,i)=><HBar key={i} name={s.name||"Unknown"} value={s.value} total={total} color={PAL[i]} rank={i+1}/>)}</div>
          <div className="section fade"><div className="section-hdr">Top States (Colleges)</div>{groupBy(filtered,"State of College").slice(0,10).map((s,i)=><HBar key={i} name={s.name||"Unknown"} value={s.value} total={total} color={PAL[i]} rank={i+1}/>)}</div>
        </div>
        <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:16}}>
          <div className="section fade"><div className="section-hdr">Top Cities</div>{groupBy(filtered,"City").slice(0,10).map((c,i)=><HBar key={i} name={c.name||"Unknown"} value={c.value} total={total} color={PAL[i]} rank={i+1}/>)}</div>
          <div className="section fade"><div className="section-hdr">Institute Status</div>{groupBy(filtered,"Status of Institute").map((s,i)=><HBar key={i} name={s.name||"Unknown"} value={s.value} total={total} color={PAL[i]} rank={i+1}/>)}</div>
        </div>
      </>)}
      {tab==="data quality"&&(<>
        <div style={{display:"grid",gridTemplateColumns:"repeat(auto-fit,minmax(170px,1fr))",gap:14,marginBottom:20}}>
          <KPI label="Total Records" value={data.length} accent={C.blue}/>
          <KPI label="After Filters" value={total} accent={C.teal}/>
          <KPI label="Status OK" value={data.filter(r=>r["Application Status"]!=="Unknown").length} accent={C.green}/>
          <KPI label="Gender OK" value={data.filter(r=>r["Gender"]!=="Not Specified").length} accent={C.purple}/>
          <KPI label="Category OK" value={data.filter(r=>r["Category"]!=="Not Specified").length} accent={C.amber}/>
          <KPI label="Columns" value={Object.keys(data[0]||{}).filter(k=>!k.startsWith("_")).length} accent={C.teal}/>
        </div>
        <div className="section fade"><div className="section-hdr">Missing Data Report</div><table className="data-table"><thead><tr><th>Field</th><th>Total</th><th>Missing</th><th>Completeness</th><th>Status</th></tr></thead><tbody>{cleanReport.map((r,i)=>{const comp=100-parseFloat(r.pct);const color=comp>90?C.green:comp>70?C.amber:C.red;return(<tr key={i}><td style={{fontWeight:600}}>{r.field}</td><td className="mono">{r.total}</td><td className="mono" style={{color:r.missing>0?C.red:C.green}}>{r.missing}</td><td style={{width:200}}><div style={{display:"flex",alignItems:"center",gap:10}}><div className="progress-bar" style={{flex:1}}><div className="progress-fill" style={{width:`${comp}%`,background:color}}/></div><span className="mono" style={{fontSize:11,color,minWidth:40}}>{comp.toFixed(1)}%</span></div></td><td><span className="badge" style={{background:`${color}20`,color}}>{comp>90?"Good":comp>70?"Partial":"Poor"}</span></td></tr>);})}</tbody></table></div>
      </>)}
    </div>
  </div>);}

export default function App(){
  const [unlocked,setUnlocked]=useState(false);
  const [mode,setMode]=useState(null);
  const [data,setData]=useState([]);
  const [fileName,setFileName]=useState("");
  const handleFile=useCallback(file=>{
  setFileName(file.name);
  const ext=file.name.split(".").pop().toLowerCase();
  if(ext==="xlsx"||ext==="xls"){
    const reader=new FileReader();
    reader.onload=e=>{
      const wb=XLSX.read(e.target.result,{type:"array"});
      const ws=wb.Sheets[wb.SheetNames[0]];
      const rows=XLSX.utils.sheet_to_json(ws,{defval:"",raw:false});
      const cleaned=cleanData(rows.map(row=>{const r={};Object.keys(row).forEach(k=>{const v=row[k];r[k.trim()]=(v===null||v===undefined)?"":String(v).trim();});return r;}));
      setData(cleaned);
    };
    reader.readAsArrayBuffer(file);
  } else {
    Papa.parse(file,{header:true,skipEmptyLines:true,complete:r=>setData(cleanData(r.data))});
  }
// eslint-disable-next-line
},[]);
  return(<><style>{CSS}</style>{!unlocked?<PasswordScreen onUnlock={()=>setUnlocked(true)}/>:mode===null?<ModeSelector onSelectMode={setMode}/>:mode===2?<ApplicationTracker onBack={()=>setMode(null)}/>:data.length===0?<UploadScreen onFile={handleFile} onBack={()=>setMode(null)}/>:<Dashboard data={data} fileName={fileName} onReset={()=>{setData([]);setFileName("");}} onBack={()=>setMode(null)}/>}</>);}