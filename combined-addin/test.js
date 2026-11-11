const fs=require("fs"),vm=require("vm"),assert=require("assert/strict");
const code=fs.readFileSync(__dirname+"/app.js","utf8");
const HEAD=["Activity","Work date","Start Time","End Time","Hours Spent","Status","Notes"];
const cells=new Map(), sheetName="Process Workflow";
function key(sheet,row,col){return sheet+"!"+row+":"+col;}
function set(sheet,row,col,value){cells.set(key(sheet,row,col),value);}
function get(sheet,row,col){return cells.get(key(sheet,row,col))??"";}
function colNumber(s){return s.split("").reduce((n,ch)=>n*26+ch.charCodeAt(0)-64,0)-1;}
function rangeOf(addr,sheet){const m=/^([A-Z]+)(\d+)(?::([A-Z]+)(\d+))?$/i.exec(addr);
 if(!m)throw Error("Bad range "+addr);
 const c1=colNumber(m[1].toUpperCase()),r1=Number(m[2])-1;
 const c2=m[3]?colNumber(m[3].toUpperCase()):c1,r2=m[4]?Number(m[4])-1:r1;
 return {r1,c1,r2,c2};
}
const handlers=[],sheets={};
function makeRange(sheet,addr){
 const p=rangeOf(addr,sheet);const obj={
 load:()=>{}, numberFormat:[],
 get values(){return Array.from({length:p.r2-p.r1+1},(_,ri)=>Array.from({length:p.c2-p.c1+1},(_,ci)=>get(sheet,p.r1+ri,p.c1+ci)));},
 set values(vals){vals.forEach((r,i)=>r.forEach((v,j)=>set(sheet,p.r1+i,p.c1+j,v)));},
 get rowIndex(){return p.r1;},get rowCount(){return p.r2-p.r1+1;},
 format:{fill:{},font:{}}
 };
 return obj;
}
function worksheet(name){
 const obj={name,id:name,getRange:a=>{
   if(/^[A-Z]+:[A-Z]+$/.test(a))return {format:{columnWidth:100}};
   return makeRange(name,a);
 },load:()=>{},freezePanes:{freezeRows:()=>{}},onChanged:{add:fn=>handlers.push(fn)},
 getUsedRange:()=>{let maxr=0,maxc=0;for(const [k,v] of cells){if(!k.startsWith(name+"!")||v==null||v==="")continue;const z=k.split("!")[1].split(":").map(Number);maxr=Math.max(maxr,z[0]);maxc=Math.max(maxc,z[1]);}let x=makeRange(name,"A1:"+String.fromCharCode(65+maxc)+(maxr+1));x.rowIndex=0;x.rowCount=maxr+1;return x;}
 };sheets[name]=obj;return obj;
}
const s=worksheet(sheetName);
set(sheetName,8,0,"01  DAILY WORKFLOW  /  TIME LOG");
HEAD.forEach((h,i)=>set(sheetName,9,i,h));
set(sheetName,10,0,"AR Follow up");set(sheetName,10,5,"Not Started");
const els=new Map();
function element(id){if(!els.has(id))els.set(id,{id,value:id==="employee"?"Test Employee":"",textContent:"",className:"",innerHTML:"",onclick:null});return els.get(id);}
const excel={HostType:{},run:async cb=>cb({workbook:{
 worksheets:{getItemOrNullObject:n=>sheets[n]||{isNullObject:true},getItem:n=>sheets[n]||(()=>{throw Error(n)})(),getActiveWorksheet:()=>sheets[sheetName],add:n=>worksheet(n)},
 getSelectedRange:()=>({rowIndex:10,rowCount:1,load:()=>{}})
},sync:async()=>{}})};
const ready=[];
const env={Excel:excel,Office:{HostType:{Excel:"Excel"},onReady:fn=>ready.push(fn)},document:{getElementById:element},console,Date,Number,String,Math,Map,Set,Error,Promise};
vm.runInNewContext(code,env,{filename:"app.js"});
const settle=async()=>{for(let i=0;i<20;i++)await new Promise(resolve=>setImmediate(resolve));};
(async()=>{
 for(const fn of ready)await fn({host:"Excel"});await settle();
 assert.match(element("state").textContent,/LIVE/);
 set(sheetName,10,5,"In Progress");
 for(const fn of handlers.slice())fn({address:"F11",worksheetId:sheetName,source:"Local"});
 await settle();
 const start=get(sheetName,10,2),date=get(sheetName,10,1);
 assert.equal(typeof start,"number");assert.equal(typeof date,"number");
 assert.equal(get(sheetName,10,3),"");
 set(sheetName,10,5,"Completed");
 for(const fn of handlers.slice())fn({address:"F11",worksheetId:sheetName,source:"Local"});
 await settle();
 const end=get(sheetName,10,3);assert.equal(typeof end,"number");
 assert.equal(get("Productivity History",1,2),"AR Follow up");
 assert.equal(get("Productivity History",1,3),"Test Employee");
 assert.equal(get("Productivity History",1,7),"Completed");
 const savedKey=get("Productivity History",1,0);
 for(const fn of handlers.slice())fn({address:"F11",worksheetId:sheetName,source:"Local"});
 await settle();
 assert.equal(get("Productivity History",1,0),savedKey);
 assert.equal(get("Productivity History",2,0),"");
 // Pasting a two-row status range must capture timestamps on both rows.
 set(sheetName,11,0,"AP Invoice Processing");
 set(sheetName,11,5,"In Progress");
 for(const fn of handlers.slice())fn({address:"F11:F12",worksheetId:sheetName,source:"Local"});
 await settle();
 assert.equal(typeof get(sheetName,11,2),"number");
 assert.equal(typeof get(sheetName,11,1),"number");
 // A multi-column edit containing column F must also be handled.
 set(sheetName,11,5,"Completed");
 for(const fn of handlers.slice())fn({address:"E12:G12",worksheetId:sheetName,source:"Local"});
 await settle();
 assert.equal(typeof get(sheetName,11,3),"number");
 set(sheetName,43,3,"In Progress");
 for(const fn of handlers.slice())fn({address:"D44",worksheetId:sheetName,source:"Local"});
 await settle();
 assert.equal(get(sheetName,43,9),"");
 console.log("PASS: daily start/end timestamp; productivity history; idempotency; monthly isolation");
})().catch(e=>{console.error("FAIL",e);process.exitCode=1;});