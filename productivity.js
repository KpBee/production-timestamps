(() => {
"use strict";
const SHEET="Productivity Log", TABLE="ProductivitySessions";
const HEAD=["Recorded UTC","Employee","Activity","Work Date","Start UTC","End UTC","Hours","Status","Notes"];
const ACTIVITIES=["AR Check Deposits & EFT Payments","AR Client Responses","Bank Feed Clearing","AP Invoice Processing","Payroll Journal - Biweekly","Time Saving","AR Aging Review","AR Follow up","Financial Review","Bank Reconciliation & Uncleared Transactions","Pre-monthly Invoices mail confirmation","Monthly Invoices","EOM","Interest Exp Check","Depreciation Check","Prepaid Exp Amort"];
let active=null;
const div=document.createElement("section");
div.innerHTML='<h2>Productivity Tracker</h2><p style="font-size:12px;color:#52606d">Work sessions are saved in this workbook. No external tracking or uploading.</p><label>Employee <input id="pr-employee" placeholder="Employee name"></label><label>Activity <select id="pr-activity"></select></label><label>Status <select id="pr-status"><option>Completed</option><option>In Progress</option><option>Blocked</option></select></label><label>Notes <input id="pr-notes" placeholder="Optional"></label><div style="display:flex;gap:8px;margin-top:12px"><button id="pr-start">Start</button><button id="pr-stop" disabled>Stop & save</button><button id="pr-refresh">Refresh</button></div><p id="pr-message" aria-live="polite">Ready</p><div id="pr-summary"></div><h3>Average Time by Activity</h3><div id="pr-activity-summary"></div><h3>Employee Workload</h3><div id="pr-employee-summary"></div><h3>Recurring Bottlenecks</h3><div id="pr-bottleneck"></div><h3>Overdue Monthly Work</h3><div id="pr-overdue"></div>';
div.style.cssText="border-top:2px solid #107c41;margin-top:20px;padding-top:10px";
document.body.append(div);
const css=document.createElement("style");
css.textContent="#pr-employee,#pr-notes,#pr-activity,#pr-status{width:100%;margin:4px 0 9px;padding:7px;box-sizing:border-box;border:1px solid #bfc9c5;border-radius:5px}#pr-start,#pr-stop,#pr-refresh{border:0;background:#107c41;color:white;padding:8px 9px;border-radius:5px;cursor:pointer}#pr-stop:disabled{opacity:.4}#pr-summary{padding:10px;background:#edf5f0;border-radius:6px;font-weight:600}#pr-activity-summary table,#pr-employee-summary table{width:100%;border-collapse:collapse;font-size:12px}#pr-activity-summary td,#pr-employee-summary td,#pr-activity-summary th,#pr-employee-summary th{padding:7px 4px;border-bottom:1px solid #ddd;text-align:left}";
document.head.append(css);
const $=x=>document.getElementById(x);
$("pr-activity").innerHTML=ACTIVITIES.map((v,i)=>'<option value="'+i+'">'+v+'</option>').join("");
function message(s,error=false){$("pr-message").textContent=s;$("pr-message").style.color=error?"#b91c1c":"#166534";}
async function ensure(){
  await Excel.run(async c=>{
    let s=c.workbook.worksheets.getItemOrNullObject(SHEET);await c.sync();
    if(s.isNullObject){s=c.workbook.worksheets.add(SHEET);s.getRange("A1:I1").values=[HEAD];const t=s.tables.add("A1:I1",true);t.name=TABLE;s.getRange("A1:I1").format.fill.color="#107c41";s.getRange("A1:I1").format.font.color="#ffffff";s.getRange("A:I").format.columnWidth=130;s.freezePanes.freezeRows(1)}
    else {let h=s.getRange("A1:I1");h.load("values");await c.sync();if(h.values[0].join("|")!==HEAD.join("|"))throw Error("Productivity Log has different headers; nothing was changed.");}
    await c.sync();
  });
}
async function save(start,end,person,activity,status,notes){
  const hours=(end-start)/36e5;
  if(hours<=0||hours>24)throw Error("Session duration must be between 0 and 24 hours.");
  await ensure();
  await Excel.run(async c=>{
    const t=c.workbook.worksheets.getItem(SHEET).tables.getItem(TABLE);
    t.rows.add(null,[[new Date().toISOString(),person,activity,start.toLocaleDateString("en-CA"),start.toISOString(),end.toISOString(),Math.round(hours*100)/100,status,notes]]);
    await c.sync();
  });
}
function htmlTable(headers,rows){
  if(!rows.length)return "<p>No recorded sessions.</p>";
  return "<table><thead><tr>"+headers.map(h=>"<th>"+h+"</th>").join("")+"</tr></thead><tbody>"+rows.map(r=>"<tr>"+r.map(v=>"<td>"+String(v).replace(/[&<>"]/g,ch=>({"&":"&amp;","<":"&lt;",">":"&gt;",'"':"&quot;"}[ch]))+"</td>").join("")+"</tr>").join("")+"</tbody></table>";
}
function dateOf(v){if(v===""||v==null)return null;if(typeof v==="number")return new Date(Date.UTC(1899,11,30)+Math.floor(v)*864e5);const d=new Date(v);return Number.isNaN(d.getTime())?null:d;}
async function refresh(){
  const d=await Excel.run(async c=>{
    const log=c.workbook.worksheets.getItemOrNullObject(SHEET);
    const workflow=c.workbook.worksheets.getItemOrNullObject("Process Workflow");await c.sync();
    const l=log.isNullObject?null:log.getUsedRange(true),w=workflow.isNullObject?null:workflow.getUsedRange(true);
    if(l)l.load("values");if(w)w.load("values");await c.sync();
    return {sessions:l?l.values:[],workflow:w?w.values:[]};
  });
  const ss=d.sessions.slice(1).filter(r=>Number(r[6])>0);
  let total=0;const acts=new Map(),people=new Map();
  for(const r of ss){const n=Number(r[6]);total+=n;const a=String(r[2]),p=String(r[1]);if(!acts.has(a))acts.set(a,[]);acts.get(a).push(r);if(!people.has(p))people.set(p,{hours:0,count:0});people.get(p).hours+=n;people.get(p).count++;}
  $("pr-summary").textContent=ss.length+" sessions · "+total.toFixed(2)+" hours · "+people.size+" employees";
  $("pr-activity-summary").innerHTML=htmlTable(["Activity","Sessions","Avg hours"],[...acts].map(([a,r])=>[a,r.length,(r.reduce((sum,x)=>sum+Number(x[6]),0)/r.length).toFixed(2)]));
  $("pr-employee-summary").innerHTML=htmlTable(["Employee","Sessions","Total hours"],[...people].map(([p,v])=>[p,v.count,v.hours.toFixed(2)]));
  const flags=[];for(const [a,rs] of acts){const blocked=rs.filter(r=>String(r[7]).toLowerCase()==="blocked").length;if(blocked>=3)flags.push(a+" · "+blocked+" blocked sessions");}
  $("pr-bottleneck").textContent=flags.length?flags.join(" | "):"No recurring blocked activities based on available logs.";
  let found=false,overdue=[];
  const rows=d.workflow;
  for(let i=0;i<rows.length;i++){
    const h=rows[i].map(v=>String(v||"").trim().toLowerCase());
    const a=h.indexOf("activity"),di=h.indexOf("due date"),si=h.indexOf("status"),ci=h.indexOf("completed on");
    if(a<0||di<0||si<0||ci<0)continue;found=true;
    for(let j=i+1;j<rows.length;j++){const r=rows[j];const label=String(r[a]||"").trim();if(!label||label.toLowerCase()==="activity")break;
      const due=dateOf(r[di]),status=String(r[si]||"").trim().toLowerCase();if(due&&!["completed","closed","not required"].includes(status)&&!r[ci]&&due.getTime()<Date.now()-864e5)overdue.push(label);
    }
    break;
  }
  $("pr-overdue").textContent=!found?"Monthly workflow headers not detected.":overdue.length?overdue.join(" | "):"No overdue activities with recorded due dates.";
}
$("pr-start").onclick=()=>{
  if(active)return;
  const person=$("pr-employee").value.trim();if(!person){message("Enter employee name.",true);return;}
  active={start:new Date(),person,activity:ACTIVITIES[Number($("pr-activity").value)],status:$("pr-status").value,notes:$("pr-notes").value};
  $("pr-start").disabled=true;$("pr-stop").disabled=false;message("Started at "+active.start.toLocaleString()+". Keep this pane open.");
};
$("pr-stop").onclick=async()=>{
  if(!active)return;
  try{await save(active.start,new Date(),active.person,active.activity,active.status,active.notes);active=null;$("pr-start").disabled=false;$("pr-stop").disabled=true;message("Session saved in Productivity Log.");await refresh();}catch(e){message(e.message,true);}
};
$("pr-refresh").onclick=()=>refresh().catch(e=>message(e.message,true));
Office.onReady(info=>{if(info.host===Office.HostType.Excel)refresh().catch(e=>message(e.message,true));else message("Open within Excel.",true);});
})();