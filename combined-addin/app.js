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
    // Excel can report a single cell, a pasted range, or a whole edited row.
    // Only react when the changed rectangle intersects the Status column (F).
    const address=String(ev.address||"").split("!").pop().replace(/\$/g,"").toUpperCase();
    const m=/^([A-Z]+)(\d+)(?::([A-Z]+)(\d+))?$/.exec(address);
    if(!m)return;
    const col=s=>[...s].reduce((n,c)=>n*26+c.charCodeAt(0)-64,0);
    const left=col(m[1]),right=col(m[3]||m[1]);
    if(left>6||right<6)return;
    const first=Math.max(FIRST,Number(m[2])),last=Math.min(LAST,Number(m[4]||m[2]));
    for(let r=first;r<=last;r++)await stamp(r);
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
      const monthly=source.getRange("A44:G50");monthly.load("values");
      const history=c.workbook.worksheets.getItemOrNullObject(HISTORY);await c.sync();
      rows.load("values");let used=null;
      if(!history.isNullObject){used=history.getUsedRange(true);used.load("values");}
      await c.sync();return {daily:rows.values,monthly:monthly.values,history:used?used.values.slice(1):[]};
    });
    const valid=result.history.filter(r=>typeof r[6]==="number");
    const sum=valid.reduce((v,r)=>v+r[6],0);
    const counts={active:0,completed:0,blocked:0};
    for(const r of result.daily){
      const s=norm(r[5]);if(s==="in progress")counts.active++;if(s==="completed")counts.completed++;if(s==="blocked")counts.blocked++;
    }
    $("totals").textContent=valid.length+" logged sessions | "+fmt(sum)+" hours | "+counts.active+" in progress | "+counts.blocked+" blocked";
    const today=Math.floor(excelIST());
    const overdue=result.monthly.filter(r=>{
      const due=r[2],status=norm(r[3]),done=r[4];
      return typeof due==="number"&&due<today&&!["completed","not required"].includes(status)&&empty(done);
    }).map(r=>String(r[0]));
    $("overdue").textContent=overdue.length?overdue.join("; "):"No overdue monthly activities with recorded due dates.";
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
  function formatClock(value){
    if(typeof value!=="number")return "—";
    const seconds=Math.round((((value%1)+1)%1)*86400)%86400;
    const hh=Math.floor(seconds/3600),mm=Math.floor((seconds%3600)/60),ss=seconds%60;
    return String(hh%12||12).padStart(2,"0")+":"+String(mm).padStart(2,"0")+":"+String(ss).padStart(2,"0")+(hh>=12?" PM":" AM");
  }
  async function showRow(){
    const n=Number($("row").value);
    if(!Number.isInteger(n)||n<FIRST||n>LAST){
      $("row-state").textContent="Select row 11–30.";
      return;
    }
    const values=await Excel.run(async c=>{
      const range=c.workbook.worksheets.getItem(SHEET).getRange("A"+n+":F"+n);
      range.load("values");await c.sync();return range.values[0];
    });
    const [activity,,start,end,hours,status]=values;
    const minutes=typeof start==="number"&&typeof end==="number"&&end>=start?(end-start)*1440:null;
    $("row-state").textContent="Row "+n+" · "+(activity||"No activity")+" · "+(status||"No status")+"\nStart: "+formatClock(start)+" · End: "+formatClock(end)+"\nElapsed: "+(minutes===null?"—":minutes.toFixed(2)+" minutes")+" · Excel hours: "+(typeof hours==="number"?hours.toFixed(2):"—")+"\nShort sessions can display 0.00 hours because the workbook rounds to 2 decimals.";
  }
  $("inspect").onclick=()=>showRow().catch(e=>message("Inspect error: "+e.message,true));
  $("row").onchange=()=>showRow().catch(e=>message("Inspect error: "+e.message,true));
  async function selectedDailyRow(){
    return Excel.run(async c=>{
      const active=c.workbook.worksheets.getActiveWorksheet();
      const selected=c.workbook.getSelectedRange();
      active.load("name");selected.load(["rowIndex","rowCount"]);
      await c.sync();
      const entered=$("row").value.trim();
      if(entered){
        const row=Number(entered);
        if(!Number.isInteger(row)||row<FIRST||row>LAST)throw Error("Row must be between 11 and 30.");
        return row;
      }
      const row=selected.rowIndex+1;
      if(active.name!==SHEET)throw Error("Select a cell in Process Workflow or enter a Daily row number.");
      if(row<FIRST||row>LAST||selected.rowCount!==1)throw Error("Select one Daily Time Log row (11–30), or enter its number.");
      return row;
    });
  }
  async function updateSelected(mode){
    try{
      if(!enabled)throw Error("The add-in is not connected to the verified Daily Workflow.");
      const row=await selectedDailyRow();
      await Excel.run(async c=>{
        const s=c.workbook.worksheets.getItem(SHEET);
        const activity=s.getRange("A"+row);
        activity.load("values");await c.sync();
        if(empty(activity.values[0][0]))throw Error("Choose an Activity on row "+row+" first.");
        s.getRange("F"+row).values=[[mode]];
        await c.sync();
      });
      await stamp(row);
      await showRow();
    }catch(e){message("Cannot "+mode.toLowerCase()+": "+e.message,true);}
  }
  $("begin").onclick=()=>updateSelected("In Progress");
  $("finish").onclick=()=>updateSelected("Completed");
  $("sync").onclick=async()=>{
    try{await recordCompletedRows();await refresh();message("Completed daily rows synchronized. Existing history is preserved.");}
    catch(e){message("Synchronization error: "+e.message,true);}
  };
  $("refresh").onclick=()=>refresh().catch(e=>message("Refresh error: "+e.message,true));
  Office.onReady(async info=>{
    if(info.host!==Office.HostType.Excel){message("Open this add-in inside Microsoft Excel.",true);return;}
    try{await verify();await refresh();await showRow();}
    catch(e){message("NOT TRACKING: "+e.message,true);}
  });
})();