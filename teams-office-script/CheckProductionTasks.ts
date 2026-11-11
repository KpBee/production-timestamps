function main(workbook: ExcelScript.Workbook, asOfDate: string) {
  if (!/^\d{4}-\d{2}-\d{2}$/.test(asOfDate)) throw new Error("asOfDate must be YYYY-MM-DD in IST");
  const sheet=workbook.getWorksheet("Process Workflow");
  if(!sheet) throw new Error("Process Workflow sheet missing");
  const rows=sheet.getRange("A1:G51").getValues(), tasks: string[]=[];
  let undatedCount=0;
  const val=(v:string|number|boolean)=>String(v??"").trim();
  const date=(v:string|number|boolean)=>{
    if(typeof v==="number"&&v>1) return new Date(Date.UTC(1899,11,30)+Math.floor(v)*86400000).toISOString().slice(0,10);
    const m=val(v).match(/^(\d{4}-\d{2}-\d{2})/);
    return m?m[1]:"";
  };
  const check=(r:number,kind:string,dc:number,sc:number)=>{
    const row=rows[r-1],task=val(row[0]),status=val(row[sc]),due=date(row[dc]);
    if(!task||task.toLowerCase().startsWith("tip:")||status.toLowerCase()==="completed")return;
    if(!due){undatedCount++;return;}
    if(due<=asOfDate)tasks.push(kind+": "+task+" | Due: "+due+" | "+(status||"Not marked")+" | Row: "+r);
  };
  for(let r=11;r<=29;r++)check(r,"Daily",1,5);
  for(let r=36;r<=40;r++)check(r,"Weekly",1,5);
  for(let r=44;r<=50;r++)check(r,"Monthly",2,3);
  return {count:tasks.length,undatedCount,asOfDate,message:tasks.length?
    "Bee Team Production Tracker — "+tasks.length+" incomplete task(s)\nChecked: "+asOfDate+" IST\n\n"+tasks.join("\n")+"\n\nPlease update the production tracker.":""};
}
