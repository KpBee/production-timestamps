/* Bee Team reminder preview. No mail permissions, no sending, no workbook writes. */
const output=document.getElementById('result');
const button=document.getElementById('scan');
const normalize=value=>String(value??'').trim();
const completed=value=>normalize(value).toLowerCase()==='completed';
function excelDate(value){
  if(typeof value==='number' && value>1){const d=new Date(Date.UTC(1899,11,30)+value*86400000);return d.toISOString().slice(0,10);}
  if(typeof value==='string' && /^\\d{4}-\\d{2}-\\d{2}/.test(value))return value.slice(0,10);
  return null;
}
const todayIST=()=>new Intl.DateTimeFormat('en-CA',{timeZone:'Asia/Kolkata',year:'numeric',month:'2-digit',day:'2-digit'}).format(new Date());
function inspectRows(values,today){
 const tasks=[];
 for(const r of [...values.slice(10,30),...values.slice(35,40)]){
  const activity=normalize(r[0]),date=excelDate(r[1]),status=normalize(r[5]);
  if(!activity||!date||date>today||completed(status))continue;
  tasks.push({activity,date,status:status||'Not marked'});
 }
 return tasks;
}
async function scan(){
 button.disabled=true;output.textContent='Reading workbook...';
 try{
  const tasks=await Excel.run(async context=>{
   const sheet=context.workbook.worksheets.getItem('Process Workflow');
   const range=sheet.getRange('A1:G40');range.load('values');await context.sync();
   return inspectRows(range.values,todayIST());
  });
  output.replaceChildren();
  const title=document.createElement('p');title.textContent=tasks.length+' incomplete scheduled task(s) identified';output.append(title);
  const ul=document.createElement('ul');tasks.forEach(task=>{let li=document.createElement('li');li.textContent=task.date+' | '+task.activity+' | '+task.status;ul.append(li);});output.append(ul);
  const note=document.createElement('p');note.textContent='Sample recipients for future authorized testing: finny.jason@beeteam.in and jawahar.muthu@beeteam.in. No message was sent.';output.append(note);
 }catch(err){output.textContent='Unable to read Process Workflow: '+err.message;}finally{button.disabled=false;}
}
Office.onReady(info=>{if(info.host===Office.HostType.Excel){button.disabled=false;output.textContent='Ready for a read-only preview.';button.addEventListener('click',scan);}});
