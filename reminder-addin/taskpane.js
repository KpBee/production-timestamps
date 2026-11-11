/* Standalone Office add-in. Read-only task analysis and user-initiated Teams chat draft. */
const output=document.getElementById('result');
const scanButton=document.getElementById('scan');
const teamsButton=document.getElementById('teams');
let pending=[];
const normalize=value=>String(value??'').trim();
const isComplete=value=>normalize(value).toLowerCase()==='completed';
function excelDate(value){
 if(value instanceof Date && !isNaN(value)) return value.toISOString().slice(0,10);
 if(typeof value==='number'&&value>1) return new Date(Date.UTC(1899,11,30)+Math.floor(value)*86400000).toISOString().slice(0,10);
 if(typeof value==='string'&&/^\d{4}-\d{2}-\d{2}/.test(value)) return value.slice(0,10);
 return null;
}
function todayIST(){
 const p=new Intl.DateTimeFormat('en-US',{timeZone:'Asia/Kolkata',year:'numeric',month:'2-digit',day:'2-digit'}).formatToParts(new Date());
 const get=t=>p.find(x=>x.type===t).value;
 return [get('year'),get('month'),get('day')].join('-');
}
function inspectRows(values,today){
 const tasks=[];let undated=0;
 function add(row,rowNumber,section,dateColumn,statusColumn){
  const activity=normalize(row[0]);if(!activity||activity.toLowerCase().startsWith('tip:')) return;
  const date=excelDate(row[dateColumn]),status=normalize(row[statusColumn]);
  if(isComplete(status))return;
  if(!date){undated++;return;}
  if(date<=today) tasks.push({section,activity,date,status:status||'Not marked',row:rowNumber});
 }
 for(let i=10;i<29;i++) add(values[i],i+1,'Daily',1,5);
 for(let i=34;i<39;i++) add(values[i],i+1,'Weekly',1,5);
 for(let i=42;i<49;i++) add(values[i],i+1,'Monthly',2,3);
 return {tasks,undated};
}
async function scan(){
 scanButton.disabled=true;teamsButton.disabled=true;output.textContent='Reading current workbook...';
 try{
  const result=await Excel.run(async context=>{
   const sheet=context.workbook.worksheets.getItem('Process Workflow');
   const range=sheet.getRange('A1:G51');range.load('values');await context.sync();
   return inspectRows(range.values,todayIST());
  });
  pending=result.tasks;output.replaceChildren();
  const count=document.createElement('p');count.textContent=pending.length+' dated incomplete task(s) found';output.append(count);
  if(result.undated){const warn=document.createElement('p');warn.className='warning';warn.textContent=result.undated+' task(s) without work/due dates were excluded; enter dates to include them.';output.append(warn);}
  if(pending.length){
   const ul=document.createElement('ul');for(const task of pending){const li=document.createElement('li');li.textContent=task.section+' | '+task.activity+' | '+task.date+' | '+task.status;ul.append(li);}output.append(ul);teamsButton.disabled=false;
  }else{const p=document.createElement('p');p.textContent='No eligible incomplete dated tasks, so no reminder is prepared.';output.append(p);}
 }catch(e){output.textContent='Could not read Process Workflow: '+(e.message||e);pending=[];}
 finally{scanButton.disabled=false;}
}
function openTeams(){
 if(!pending.length)return;
 const lines=['Bee Team Production Tracker — incomplete tasks','Checked '+todayIST()+' (IST)','',
 ...pending.map(t=>'- '+t.section+': '+t.activity+' | '+t.date+' | '+t.status),'',
 'Please update the Production Tracker status and any blockers.'];
 const chat='https://teams.microsoft.com/l/chat/0/0?users='+encodeURIComponent('finny.jason@beeteam.in,jawahar.muthu@beeteam.in')+
 '&topicName='+encodeURIComponent('Production Reminders')+'&message='+encodeURIComponent(lines.join('\n'));
 try{Office.context.ui.openBrowserWindow(chat);}catch(e){window.open(chat,'_blank','noopener,noreferrer');}
 output.append(document.createTextNode(' Teams chat opened. The reminder is a draft; click Send in Teams.'));
}
Office.onReady(info=>{
 if(info.host===Office.HostType.Excel){
  scanButton.disabled=false;output.textContent='Ready to check this workbook.';
  scanButton.addEventListener('click',scan);teamsButton.addEventListener('click',openTeams);
 }
});
