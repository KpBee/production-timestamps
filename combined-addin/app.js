"use strict";
(() => {
  const SHEET="Process Workflow", HISTORY="Productivity History", FIRST=11, LAST=30;
  const HEAD=["Session Key","Work Date","Activity","Employee","Start Time","End Time","Hours","Status","Recorded UTC"];
  const $=id=>document.getElementById(id);
  const empty=v=>v===null||v===undefined||v==="";
  const norm=v=>String(v??"").trim().toLowerCase().replace(/\s+/g," ");
  const excelIST=()=>25569+(Date.now()+330*60000)/86400000;
  const fmt=n=>Number(n).toFixed(2);
  const clean=v=>String(v??"").replace(/[&<>"']/g,c=>({"&":"&amp;","<":"&lt;",">":"&gt;",'"':"&quot;","'":"&#39;"}[c]));
  let queue=Promise.resolve(), enabled=false, sheetId=null;
  const message=(s,err=false)=>{$("state").textContent=s;$("state").className=err?"error":"ok";};
  function duration(start,end){
    if(typeof start!=="number"||typeof end!=="number"||end<start)return null;
    return Math.round((end-start)*24*100)/100;
  }
  function key(row,date,activity,start,end){
    return [row,Number(date).toFixed(6),String(activity),Number(start).toFixed(9),Number(end).toFixed(9)].join("|");
  }
  async function verify(){
    await Excel.run(async c=>{
      const s=c.workbook.worksheets.getItemOrNullObject(SHEET);await c.sync();
      if(s.isNullObject)throw Error("The worksheet 'Process Workflow' was not found.");
      const h=s.getRange("A9:G10");h.load("values");s.load("id");await c.sync();
      const r=h.values;
      if(!norm(r[0][0]).includes("daily workflow")||!norm(r[0][0]).includes("time log"))
        throw Error("The Daily Workflow section is not at row 9. No changes were made.");
      const expected=["Activity","Work date","Start Time","End Time","Hours Spent","Status"];
      for(let i=0;i<expected.length;i++)if(norm(r[1][i])!==norm(expected[i]))
        throw Error("Column "+String.fromCharCode(65+i)+" does not match the expected Daily Time Log. No changes were made.");
      sheetId=s.id;
      s.onChanged.add(ev=>{
        queue=queue.then(()=>handle(ev)).catch(e=>message("Timestamp error: "+e.message,true));
      });
      await c.sync();
    });
    enabled=true;
    message("LIVE: Monitoring only Process Workflow!F11:F30. Keep this pane open.");
  }
  async function handle(ev){
    if(!enabled|| (ev.source&&ev.source!=="Local"))return;
    if(ev.worksheetId&&ev.worksheetId!==sheetId)return;
    const a=String(ev.address||"").split("!").pop().replace(/\$/g,"");
    const m=/^F(\d{1,3})$/i.exec(a);
    if(!m)return;
    const r=Number(m[1]);
    if(r<FIRST||r>LAST)return;
    await stamp(r);
  }
  async function stamp(row){
    let completed=false,activity="";
    await Excel.run(async c=>{
      const s=c.workbook.worksheets.getItem(SHEET),range=s.getRange("A"+row+":G"+row);
      range.load("values");await c.sync();
      let [a,date,start,end,,status]=range.values[0];
      if(empty(a))return;
      activity=String(a);
      const mode=norm(status);
      if(mode!=="in progress"&&mode!=="completed")return;
      const serial=excelIST(),today=Math.floor(serial),time=serial-today;
      if(empty(date)){
        const target=s.getRange("B"+row);target.values=[[today]];
        target.numberFormat=[["dd-mmm-yyyy"]];
      }
      if(mode==="in progress"&&empty(start)){
        const target=s.getRange("C"+row);target.values=[[time]];
        target.numberFormat=[["hh:mm AM/PM"]];
      }
      if(mode==="completed"&&empty(end)){
        const target=s.getRange("D"+row);target.values=[[time]];
        target.numberFormat=[["hh:mm AM/PM"]];
      }
      completed=mode==="completed";
      await c.sync();
    });
    if(completed){
      try{await recordCompletedRows([row]);message("Completed timestamp saved: "+activity+" (row "+row+").");}
      catch(e){message("Timestamp saved, but productivity history could not update: "+e.message,true);}
    }else if(activity) message("Start timestamp captured: "+activity+" (row "+row+").");
    await refresh();
  }
  async function getOrCreateHistory(c){
    let s=c.workbook.worksheets.getItemOrNullObject(HISTORY);await c.sync();
    if(s.isNullObject){
      s=c.workbook.worksheets.add(HISTORY);
      s.getRange("A1:I1").values=[HEAD];
      s.getRange("A1:I1").format.fill.color="#157c60";
      s.getRange("A1:I1").format.font.color="#ffffff";
      s.getRange("A:I").format.columnWidth=110;
      s.getRange("C:C").format.columnWidth=210;
    }else{
      const h=s.getRange("A1:I1");h.load("values");await c.sync();
      if(h.values[0].join("|")!==HEAD.join("|"))throw Error("Productivity History has unexpected headings; history was not altered.");
    }
    await c.sync();return s;
  }
  async function recordCompletedRows(onlyRows){
    await Excel.run(async c=>{
      const source=c.workbook.worksheets.getItem(SHEET),data=source.getRange("A11:G30");
      data.load("values");await c.sync();
      const candidates=[];
      data.values.forEach((v,i)=>{
        const row=i+FIRST,[activity,date,start,end,,status]=v;
        if(onlyRows&&!onlyRows.includes(row))return;
        if(empty(activity)||norm(status)!=="completed"||empty(date))return;
        const hrs=duration(start,end);
        if(hrs===null)return;
        candidates.push({id:key(row,date,activity,start,end),activity:String(activity),date,start,end,hrs});
      });
      if(!candidates.length)return;
      const history=await getOrCreateHistory(c);
      const used=history.getUsedRange(true);used.load(["rowCount","rowIndex","values"]);await c.sync();
      const existing=new Set(used.values.slice(1).map(r=>String(r[0])));
      const employee=$("employee").value.trim();
      let next=used.rowIndex+used.rowCount+1;
      for(const s of candidates){
        if(existing.has(s.id))continue;
        history.getRange("A"+next+":I"+next).values=[[s.id,s.date,s.activity,employee,s.start,s.end,s.hrs,"Completed",new Date().toISOString()]];
        history.getRange("B"+next).numberFormat=[["dd-mmm-yyyy"]];
        history.getRange("E"+next+":F"+next).numberFormat=[["hh:mm AM/PM"]];
        existing.add(s.id);next++;
      }
      await c.sync();
    });
  }
  async function refresh(){
    const result=await Excel.run(async c=>{
      const source=c.workbook.worksheets.getItem(SHEET),rows=source.getRange("A11:G30");
      const history=c.workbook.worksheets.getItemOrNullObject(HISTORY);await c.sync();
      rows.load("values");let used=null;
      if(!history.isNullObject){used=history.getUsedRange(true);used.load("values");}
      await c.sync();return {daily:rows.values,history:used?used.values.slice(1):[]};
    });
    const valid=result.history.filter(r=>typeof r[6]==="number");
    const sum=valid.reduce((v,r)=>v+r[6],0);
    const counts={active:0,completed:0,blocked:0};
    for(const r of result.daily){
      const s=norm(r[5]);if(s==="in progress")counts.active++;if(s==="completed")counts.completed++;if(s==="blocked")counts.blocked++;
    }
    $("totals").textContent=valid.length+" logged sessions | "+fmt(sum)+" hours | "+counts.active+" in progress | "+counts.blocked+" blocked";
    const byTask=new Map(),byPerson=new Map();
    for(const r of valid){
      const name=String(r[2]),person=String(r[3]||"Unassigned"),hours=Number(r[6]);
      if(!byTask.has(name))byTask.set(name,[]);
      byTask.get(name).push(hours);
      byPerson.set(person,(byPerson.get(person)||0)+hours);
    }
    const taskRows=[...byTask].map(([name,a])=>[name,String(a.length),fmt(a.reduce((x,y)=>x+y,0)/a.length)+" h"]);
    const employeeRows=[...byPerson].map(([name,h])=>[name,fmt(h)+" h"]);
    drawTable("tasks",["Activity","Sessions","Average"],taskRows);
    drawTable("people",["Employee","Hours"],employeeRows);
    const avg=[...byTask].filter(([,a])=>a.length>=3).map(([name,a])=>({name,avg:a.reduce((x,y)=>x+y,0)/a.length}));
    const overall=valid.length?sum/valid.length:0;
    const slow=avg.filter(x=>x.avg>overall*1.5);
    $("flags").textContent=slow.length?slow.map(x=>x.name+" ("+fmt(x.avg)+" h avg)").join("; "):"No recurring duration bottleneck identified (requires at least 3 sessions per activity).";
  }
  function drawTable(id,head,rows){
    const h="<thead><tr>"+head.map(v=>"<th>"+clean(v)+"</th>").join("")+"</tr></thead>";
    const b=rows.map(r=>"<tr>"+r.map(v=>"<td>"+clean(v)+"</td>").join("")+"</tr>").join("");
    $(id).innerHTML="<table>"+h+"<tbody>"+b+"</tbody></table>"+(rows.length?"":"<p>No logged sessions yet.</p>");
  }
  $("sync").onclick=async()=>{
    try{await recordCompletedRows();await refresh();message("Completed daily rows synchronized. Existing history is preserved.");}
    catch(e){message("Synchronization error: "+e.message,true);}
  };
  $("refresh").onclick=()=>refresh().catch(e=>message("Refresh error: "+e.message,true));
  Office.onReady(async info=>{
    if(info.host!==Office.HostType.Excel){message("Open this add-in inside Microsoft Excel.",true);return;}
    try{await verify();await refresh();}
    catch(e){message("NOT TRACKING: "+e.message,true);}
  });
})();