(() => {
"use strict";
const ORG={kpmg:"KPMG Assurance and Consulting Services LLP",default:"Edgeverve Systems Limited- NaBFID"};
const OUT=["Tickets#","Created on","Department","Prioritytitle","Type","subject","Classification","Organization","Wing","Closedon","resolution_steps","Status"];
const SUMMARY=["Client","Closed","Pending on COE","Pending on Customer","Grand Total"];
const EXCLUDED=[/not\s+reporting/i,/devices\s+not\s+working/i,/devices\s+not\s+reporting/i];
const $=id=>document.getElementById(id);
let createdRows=[],closedRows=[],processedRows=[],reviewRows=[],availableDates=[],selectedDate="";
const el={fileInput:$("fileInput"),dropzone:$("dropzone"),fileName:$("fileName"),dateSelection:$("dateSelection"),reportDate:$("reportDate"),datePickerButton:$("datePickerButton"),datePickerValue:$("datePickerValue"),datePickerPopover:$("datePickerPopover"),calendarHint:$("calendarHint"),calendarGrid:$("calendarGrid"),processBtn:$("processBtn"),dateSelectionStatus:$("dateSelectionStatus"),reviewPanel:$("reviewPanel"),reviewBody:$("reviewBody"),reviewCount:$("reviewCount"),continueBtn:$("continueBtn"),cancelReviewBtn:$("cancelReviewBtn"),resultSection:$("resultSection"),previewBtn:$("previewBtn"),editBtn:$("editBtn"),copyTableBtn:$("copyTableBtn"),downloadBtn:$("downloadBtn"),copyEmailBtn:$("copyEmailBtn"),copySubjectBtn:$("copySubjectBtn"),emailPreview:$("emailPreview"),dataPanel:$("dataPanel"),dataTitle:$("dataTitle"),dataSubtitle:$("dataSubtitle"),dataHead:$("dataHead"),dataBody:$("dataBody"),saveBtn:$("saveBtn"),closePanelBtn:$("closePanelBtn")};
function clean(v){return v==null?"":String(v).trim()}
function val(r,c){const k=Object.keys(r||{}).find(h=>clean(h).toLowerCase()===c.toLowerCase());return k?clean(r[k]):""}
function has(rows,c){return rows.length&&Object.keys(rows[0]).some(h=>clean(h).toLowerCase()===c.toLowerCase())}
function dateVal(v){if(v instanceof Date&&!isNaN(v.getTime()))return v;const p=parts(v);if(!p)return clean(v);const d=new Date(p.year,p.month-1,p.day,p.hour||0,p.minute||0,p.second||0);return isNaN(d.getTime())?clean(v):d}
function parts(v){
  if(v instanceof Date&&!isNaN(v.getTime()))return{day:v.getDate(),month:v.getMonth()+1,year:v.getFullYear(),hour:v.getHours(),minute:v.getMinutes(),second:v.getSeconds()};
  const s=clean(v);if(!s)return null;
  if(/^\d+(?:\.\d+)?$/.test(s)){const n=+s;if(n>0&&n<100000&&XLSX.SSF?.parse_date_code){const p=XLSX.SSF.parse_date_code(n);if(p?.y&&p?.m&&p?.d)return{day:p.d,month:p.m,year:p.y,hour:p.H||0,minute:p.M||0,second:Math.floor(p.S||0)}}return null}
  let m=s.match(/^(\d{4})[\/\-](\d{1,2})[\/\-](\d{1,2})(?:[ T,]+(\d{1,2}):(\d{2})(?::(\d{2}))?\s*(AM|PM))?$/i),year,month,day,hour=0,minute=0,second=0;
  if(m){year=+m[1];month=+m[2];day=+m[3]}else{m=s.match(/^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{4})(?:[ T,]+(\d{1,2}):(\d{2})(?::(\d{2}))?\s*(AM|PM))?$/i);if(m){const a=+m[1],b=+m[2];year=+m[3];if(a>12&&b<=12){day=a;month=b}else if(b>12&&a<=12){month=a;day=b}else{day=a;month=b}}}
  if(!m){const p=parseITSMDate(s);if(p)return{day:p.getDate(),month:p.getMonth()+1,year:p.getFullYear(),hour:p.getHours(),minute:p.getMinutes(),second:p.getSeconds()};return null}
  if(m[4]!=null){hour=+m[4];minute=+m[5];second=+(m[6]||0);const ap=(m[7]||"").toUpperCase();if(ap){if(hour<1||hour>12)return null;if(hour===12)hour=0;if(ap==="PM")hour+=12}else if(hour>23)return null;if(minute>59||second>59)return null}
  const d=new Date(year,month-1,day,hour,minute,second);return isNaN(d.getTime())||d.getFullYear()!==year||d.getMonth()!==month-1||d.getDate()!==day||d.getHours()!==hour||d.getMinutes()!==minute||d.getSeconds()!==second?null:{day,month,year,hour,minute,second}
}
function key(v){const p=parts(v);return p?[p.year,String(p.month).padStart(2,"0"),String(p.day).padStart(2,"0")].join("-"):""}
function detectDateFormat(v){
  const s=clean(v);const fmt={dateSep:"-",dayPad:true,monthPad:true,yearFirst:false,timeSep:" ",hourPad:false,seconds:false,ampm:false,ampmLower:false,hasTime:false,monthName:"",monthFirst:false};
  if(!s||/^\d+(?:\.\d+)?$/.test(s))return fmt;
  fmt.hasTime=/[ T,]+\d{1,2}:\d{2}/.test(s);fmt.seconds=/:\d{2}:\d{2}/.test(s);
  const ap=s.match(/\b(AM|PM)\b/i);fmt.ampm=!!ap;fmt.ampmLower=!!ap&&ap[1]===ap[1].toLowerCase();
  fmt.timeSep= /,\s*\d/.test(s)?", ":" ";
  const dp=s.split(/[ T,]+(?=\d)/)[0];
  const n=dp.match(/^(\d{1,4})([\/\-])(\d{1,2})\2(\d{1,4})$/);
  if(n){fmt.dateSep=n[2];fmt.yearFirst=n[1].length===4;if(fmt.yearFirst){fmt.monthPad=n[3].length===2;fmt.dayPad=n[4].length===2}else{fmt.dayPad=n[1].length===2;fmt.monthPad=n[3].length===2}}
  else if(/^[A-Za-z]{3,9}/.test(dp)){const m=dp.match(/^([A-Za-z]{3,9})\s+(\d{1,2})\s*,?\s*(\d{4})$/);if(m){fmt.monthName=m[1];fmt.dayPad=m[2].length===2;fmt.dateSep=" ";fmt.monthFirst=/^[A-Za-z]/.test(dp);}}
  const tm=s.match(/(?:[ T,]+)(\d{1,2}):(\d{2})(?::(\d{2}))?/);fmt.hourPad=!!tm&&tm[1].length===2;
  return fmt;
}
let dateOutputFormat=detectDateFormat("");
function applyCreatedFormat(v){dateOutputFormat=detectDateFormat(v);return dateOutputFormat}
function formatDateOutput(v,fmt=dateOutputFormat){
  const d=dateVal(v);if(!(d instanceof Date)||isNaN(d))return clean(v);
  const day=String(d.getDate()),month=String(d.getMonth()+1),year=String(d.getFullYear());
  let out;
  if(fmt.monthName){const names=["January","February","March","April","May","June","July","August","September","October","November","December"];const name=fmt.monthName.length<=3?names[d.getMonth()].slice(0,3):names[d.getMonth()];out=fmt.monthFirst?name+" "+(fmt.dayPad?day.padStart(2,"0"):day)+", "+year:(fmt.dayPad?day.padStart(2,"0"):day)+" "+name+" "+year}
  else if(fmt.yearFirst)out=year+fmt.dateSep+(fmt.monthPad?month.padStart(2,"0"):month)+fmt.dateSep+(fmt.dayPad?day.padStart(2,"0"):day);
  else out=(fmt.dayPad?day.padStart(2,"0"):day)+fmt.dateSep+(fmt.monthPad?month.padStart(2,"0"):month)+fmt.dateSep+year;
  if(fmt.hasTime){let h=fmt.ampm?(d.getHours()%12||12):d.getHours(),hs=String(h);if(fmt.hourPad)hs=hs.padStart(2,"0");out+=fmt.timeSep+hs+":"+String(d.getMinutes()).padStart(2,"0");if(fmt.seconds)out+=":"+String(d.getSeconds()).padStart(2,"0");if(fmt.ampm)out+=" "+(fmt.ampmLower?(d.getHours()>=12?"pm":"am"):(d.getHours()>=12?"PM":"AM"))}
  return out;
}
function displayDate(v){return formatDateOutput(v,dateOutputFormat)}
function excelDateFormat(fmt){
  let d;
  if(fmt.monthName)d=(fmt.dayPad?"dd":"d")+" "+(fmt.monthName.length<=3?"mmm":"mmmm")+" yyyy";
  else if(fmt.yearFirst)d="yyyy"+fmt.dateSep+(fmt.monthPad?"mm":"m")+fmt.dateSep+(fmt.dayPad?"dd":"d");
  else d=(fmt.dayPad?"dd":"d")+fmt.dateSep+(fmt.monthPad?"mm":"m")+fmt.dateSep+"yyyy";
  if(fmt.hasTime)d+=fmt.timeSep+(fmt.hourPad?"hh":"h")+":mm"+(fmt.seconds?":ss":"")+(fmt.ampm?(fmt.ampmLower?" am/pm":" AM/PM"):"");
  return d;
}

function ordinal(n){return n%100>=11&&n%100<=13?n+"th":n+({1:"st",2:"nd",3:"rd"}[n%10]||"th")}function detectDateFormat(v){
  const s=clean(v);if(!s||/^\d+(?:\.\d+)?$/.test(s))return"dd-mm-yyyy hh:mm:ss";
  const hasTime=/[ T]+\d{1,2}:\d{2}/.test(s),hasSec=/:[0-5]?\d:[0-5]?\d/.test(s),amp=/\b(?:AM|PM)\b/i.test(s);
  const comma=hasTime&&/,/.test(s)?", ":" ";
  const date=s.split(/[ T,]+/)[0];
  let dateFmt="dd-mm-yyyy";
  if(/^\d{4}[-\/]\d{1,2}[-\/]\d{1,2}$/.test(date)){const sep=date.includes("/")?"/":"-";const a=date.split(sep);dateFmt="yyyy"+sep+(a[1].length===2?"mm":"m")+sep+(a[2].length===2?"dd":"d")}
  else if(/^[A-Za-z]{3,9}/.test(date)){const m=date.match(/^([A-Za-z]{3,9})\s+(\d{1,2})/);dateFmt=m?("mmmm "+(m[2].length===2?"dd":"d")+", yyyy"):"dd-mmm-yyyy"}
  else{const sep=date.includes("/")?"/":"-";const a=date.split(sep);dateFmt=(a[0].length===2?"dd":"d")+sep+(a[1].length===2?"mm":"m")+sep+"yyyy"}
  if(!hasTime)return dateFmt;
  const tm=s.match(/(?:[ T]+|,\s*)(\d{1,2}):(\d{2})(?::(\d{2}))?/);if(!tm)return dateFmt;
  return dateFmt+comma+(tm[1].length===2?"hh":"h")+":mm"+(hasSec?":ss":"")+(amp?" AM/PM":"");
}
let dateOutputFormat="dd-mm-yyyy hh:mm:ss";
function displayDateWithFormat(v,fmt=dateOutputFormat){
  const d=dateVal(v);if(!(d instanceof Date)||isNaN(d))return clean(v);
  const p={d:d.getDate(),dd:String(d.getDate()).padStart(2,"0"),m:d.getMonth()+1,mm:String(d.getMonth()+1).padStart(2,"0"),yyyy:d.getFullYear(),h:d.getHours()%12||12,hh:String(d.getHours()%12||12),ss:String(d.getSeconds()).padStart(2,"0"),amp:d.getHours()>=12?"PM":"AM",mmmm:["January","February","March","April","May","June","July","August","September","October","November","December"][d.getMonth()]};
  return fmt.replace(/yyyy|mmmm|dd|mm|hh|ss|d|m|h|AM\/PM/g,t=>({yyyy:p.yyyy,mmmm:p.mmmm,dd:p.dd,mm:p.mm,hh:p.hh,ss:p.ss,d:p.d,m:p.m,h:p.h,"AM/PM":p.amp})[t]??t);
}

function emailDate(v){const p=parts(v);if(!p)return"";return ordinal(p.day)+" "+["January","February","March","April","May","June","July","August","September","October","November","December"][p.month-1]+" "+p.year}
function classify(s){const m=clean(s).match(/Domain:\s*([^|]*)/i);return m?clean(m[1])||"NABFID DC":"NABFID DC"}
function org(s){return /KPMG/i.test(clean(s))?ORG.kpmg:ORG.default}
function nab(r){return /nabfid/i.test(val(r,"subject"))||/nabfid/i.test(val(r,"Organization"))}
function soc(r){return val(r,"Type").toLowerCase()==="soc alert"}
function excluded(r){const s=val(r,"subject");return s&&EXCLUDED.some(p=>p.test(s))}
function status(r){return clean(val(r,"Status")).replace(/\s+/g," ").toLowerCase()}
function output(r){const s=val(r,"subject");return{"Tickets#":val(r,"Tickets#"),"Created on":dateVal(val(r,"Created on")),"Department":val(r,"Department"),"Prioritytitle":val(r,"Prioritytitle"),"Type":"SOC Alert",subject:s,"Classification":classify(s),"Organization":org(s),"Wing":val(r,"Wingname"),"Closedon":dateVal(val(r,"Closedon")),"resolution_steps":val(r,"resolution_steps"),"Status":val(r,"Status")}}
function valid(rows,name){if(!rows.length)throw Error(name+" file is empty.");for(const c of["Tickets#","Created on","subject","Type"])if(!has(rows,c))throw Error(name+' file is missing "'+c+'".')}
async function read(file){const wb=XLSX.read(await file.arrayBuffer(),{type:"array",cellDates:false,cellNF:true,cellText:true});const sh=wb.Sheets[wb.SheetNames[0]];return{file,rows:XLSX.utils.sheet_to_json(sh,{defval:"",raw:true})}}
function reset(){createdRows=[];closedRows=[];processedRows=[];reviewRows=[];availableDates=[];selectedDate="";el.reportDate.value="";el.datePickerValue.textContent="Select a date";el.processBtn.disabled=true;el.dateSelection.classList.add("hidden");el.reviewPanel.classList.add("hidden");el.resultSection.classList.add("hidden");el.dataPanel.classList.add("hidden")}
async function handleFiles(files){if(!files.length)return;if(files.length>2)return alert("Please upload at most 2 Excel files.");if(typeof XLSX==="undefined")return alert("Excel engine is not loaded. Please refresh.");reset();el.fileName.textContent="Reading workbook"+(files.length===1?"":"s")+"…";try{const a=await Promise.all([...files].map(read));let c=a.find(x=>!has(x.rows,"resolution_steps")),d=a.find(x=>has(x.rows,"resolution_steps"));if(files.length===2&&(!c||!d))throw Error('Upload one Created file without "resolution_steps" and one Closed file containing "resolution_steps".');if(files.length===1&&!c&&!d)throw Error('The workbook must contain a "resolution_steps" column to be used as the Closed dashboard, or omit it for the Created dashboard.');if(c)valid(c.rows,"Created");if(d)valid(d.rows,"Closed");createdRows=c?c.rows:[];closedRows=d?d.rows:[];applyCreatedFormat(c?.rows?.[0] ? val(c.rows[0],"Created on") : "");const dateRows=createdRows.length?createdRows:closedRows;availableDates=[...new Set(dateRows.map(r=>key(val(r,"Created on"))).filter(Boolean))].sort();if(!availableDates.length)throw Error('No valid "Created on" dates found in the uploaded file.');el.fileName.textContent=files.length===1?"Ready: "+a[0].file.name:"Ready: "+c.file.name+" + "+d.file.name;el.dateSelectionStatus.textContent=(createdRows.length?"Created: "+createdRows.length.toLocaleString()+" records":"")+" "+(createdRows.length&&closedRows.length?"• ":"")+(closedRows.length?"Closed: "+closedRows.length.toLocaleString()+" records":"");el.dateSelection.classList.remove("hidden");calendar();}catch(e){reset();el.fileName.textContent="Processing failed";alert(e.message||e)}}
function calendar(){el.calendarHint.textContent=availableDates.length+" date"+(availableDates.length===1?"":"s")+" available";el.calendarGrid.innerHTML="";[...availableDates].sort().reverse().forEach(k=>{const [y,m,d]=k.split("-").map(Number),b=document.createElement("button");b.type="button";b.className="date-option"+(k===el.reportDate.value?" is-selected":"");b.setAttribute("role","option");b.setAttribute("aria-selected",k===el.reportDate.value?"true":"false");const main=document.createElement("span");main.className="date-option-main";main.textContent=new Date(y,m-1,d).toLocaleDateString("en-GB",{day:"2-digit",month:"short",year:"numeric"});const sub=document.createElement("span");sub.className="date-option-sub";sub.textContent=new Date(y,m-1,d).toLocaleDateString("en-GB",{weekday:"long"});b.append(main,sub);b.onclick=()=>{el.reportDate.value=k;el.datePickerValue.textContent=new Date(y,m-1,d).toLocaleDateString("en-GB",{day:"numeric",month:"long",year:"numeric"});el.processBtn.disabled=false;closeCal();calendar()};el.calendarGrid.appendChild(b)})}
function closeCal(){el.datePickerPopover.classList.add("hidden");el.datePickerButton.setAttribute("aria-expanded","false")}
function eligibleCreated(d){return createdRows.filter(r=>key(val(r,"Created on"))===d).filter(nab).filter(soc).filter(r=>!excluded(r))}
function eligibleClosed(d){return closedRows.filter(r=>key(val(r,"Created on"))===d).filter(nab).filter(soc).filter(r=>!excluded(r)).filter(r=>!!val(r,"resolution_steps"))}
function processDate(){const d=clean(el.reportDate.value);if(!d)return alert("Please select a report date first.");selectedDate=d;const cr=eligibleCreated(d),cl=eligibleClosed(d),closedIds=new Set(cl.map(r=>val(r,"Tickets#")).filter(Boolean));reviewRows=cr.filter(r=>status(r)!=="closed"&&!closedIds.has(val(r,"Tickets#"))).map(r=>{const o=output(r);o.__original={Status:o.Status,Wing:o.Wing,Closedon:o.Closedon,resolution_steps:o.resolution_steps};o.__updates={};return o;});processedRows=cl.map(output);el.dateSelectionStatus.textContent=(createdRows.length?cr.length+" eligible Created records":"")+" "+(createdRows.length&&closedRows.length?"• ":"")+(closedRows.length?cl.length+" Closed records":"");if(reviewRows.length){renderReview();el.reviewPanel.classList.remove("hidden");el.resultSection.classList.add("hidden");el.reviewPanel.scrollIntoView({behavior:"smooth",block:"start"})}else finalize()}
function closeReviewDropdowns(){document.querySelectorAll(".review-dropdown-menu.is-open").forEach(m=>{m.classList.remove("is-open");m.remove()})}
function createReviewDropdown(field,current,options,row){
  const wrap=document.createElement("div");wrap.className="review-dropdown";
  const button=document.createElement("button");button.type="button";button.className="review-dropdown-trigger";button.setAttribute("aria-haspopup","listbox");button.setAttribute("aria-expanded","false");
  const label=document.createElement("span");label.className="review-dropdown-value";label.textContent=current||"Select";
  const icon=document.createElement("span");icon.className="review-dropdown-chevron";icon.textContent="⌄";button.append(label,icon);
  const open=()=>{
    closeReviewDropdowns();
    const menu=document.createElement("div");menu.className="review-dropdown-menu is-open";menu.setAttribute("role","listbox");
    const rect=button.getBoundingClientRect();const width=Math.max(rect.width,field==="Wing"?260:220);
    menu.style.width=width+"px";menu.style.left=Math.max(8,Math.min(rect.left,window.innerWidth-width-8))+"px";
    const below=window.innerHeight-rect.bottom, above=rect.top;
    menu.style.top=(below>=220||below>=above?rect.bottom+6:Math.max(8,rect.top-Math.min(220,options.length*42+16)-6))+"px";
    options.forEach(value=>{
      const item=document.createElement("button");item.type="button";item.className="review-dropdown-option"+(value.toLowerCase()===current.toLowerCase()?" is-selected":"");item.setAttribute("role","option");item.setAttribute("aria-selected",value.toLowerCase()===current.toLowerCase()?"true":"false");
      const text=document.createElement("span");text.textContent=value;item.appendChild(text);
      if(value.toLowerCase()===current.toLowerCase()){const check=document.createElement("span");check.className="review-dropdown-check";check.textContent="✓";item.appendChild(check)}
      item.onclick=e=>{e.stopPropagation();row.__updates[field]=value;label.textContent=value;closeReviewDropdowns();if(field==="Status"){const rowEl=button.closest("tr");const closedInput=rowEl?.querySelector(".review-closedon-input");const resolution=rowEl?.querySelector(".review-resolution-input");const isClosed=value.toLowerCase()==="closed";if(closedInput){closedInput.disabled=!isClosed;closedInput.classList.toggle("is-disabled",!isClosed);if(!isClosed){closedInput.value="";delete row.__updates.Closedon;}}if(resolution){resolution.classList.toggle("review-required",isClosed);resolution.setAttribute("aria-required",isClosed?"true":"false");}}};
      menu.appendChild(item)
    });
    document.body.appendChild(menu);button.setAttribute("aria-expanded","true");
  };
  button.onclick=e=>{e.stopPropagation();const openMenu=document.querySelector(".review-dropdown-menu.is-open");if(openMenu){closeReviewDropdowns();button.setAttribute("aria-expanded","false")}else open()};
  wrap.appendChild(button);return wrap
}
function renderReview(){
  el.reviewCount.textContent=reviewRows.length+" ticket"+(reviewRows.length===1?"":"s")+" require review";el.reviewBody.innerHTML="";
  reviewRows.forEach((r,index)=>{
    const tr=document.createElement("tr");
    [["Tickets#",r["Tickets#"]],["Created on",displayDate(r["Created on"])],["subject",r.subject],["Status",r.Status],["Wing",r.Wing]].forEach(([field,value],pos)=>{
      const td=document.createElement("td");td.textContent=value||"";
      if(pos===3||pos===4){td.className="review-current-cell";const valueEl=document.createElement("span");valueEl.className="review-current-value";valueEl.textContent=value||"—";td.replaceChildren(valueEl)}
      tr.appendChild(td)
    });
    const STATUS_OPTIONS=["Closed","Pending on customer","Pending on COE"];
    const WING_OPTIONS=["False Positive","True Positive - actioned by Teams","True Positive - NO Impact","True Positive Security Incident"];
    for(const [f,t] of[["Status","dropdown"],["Wing","dropdown"],["Closedon","text"],["resolution_steps","textarea"]]){
      const td=document.createElement("td");td.className=f==="Status"||f==="Wing"?"review-edit-cell":"";
      if(t==="dropdown"){
        const current=clean(r[f]);const opts=f==="Status"?STATUS_OPTIONS.slice():WING_OPTIONS.slice();
        if(current&&!opts.some(v=>v.toLowerCase()===current.toLowerCase()))opts.unshift(current);
        td.appendChild(createReviewDropdown(f,current,opts,r));
      }else{
        const i=document.createElement(t==="textarea"?"textarea":"input");
        if(f==="Closedon"){
          i.className="review-closedon-input";i.disabled=clean(r.Status).toLowerCase()!=="closed";
          i.placeholder="Available when status is Closed";i.value=r[f] instanceof Date&&!isNaN(r[f])?displayDate(r[f]):clean(r[f]);
          i.onblur=()=>{const raw=clean(i.value);if(!raw){delete r.__updates[f];i.classList.remove("input-invalid");return}const d=parseITSMDate(raw);if(d){r.__updates[f]=d;i.value=displayDate(d);i.classList.remove("input-invalid")}else i.classList.add("input-invalid")};
        }else{i.className="review-resolution-input";i.value=clean(r[f]);i.placeholder="Required when status is Closed";i.oninput=e=>{r.__updates[f]=e.target.value}}
        td.appendChild(i)
      }
      tr.appendChild(td)
    }
    const actionTd=document.createElement("td");const remove=document.createElement("button");remove.type="button";remove.className="review-remove-btn";remove.textContent="Remove";
    remove.onclick=()=>{closeReviewDropdowns();reviewRows.splice(index,1);renderReview();if(!reviewRows.length)el.reviewCount.textContent="No tickets require review"};
    actionTd.appendChild(remove);tr.appendChild(actionTd);el.reviewBody.appendChild(tr)
  })
}function validateReviewRows(){for(const r of reviewRows){const s=clean(Object.prototype.hasOwnProperty.call(r.__updates||{},"Status")?r.__updates.Status:r.Status).toLowerCase();if(s==="closed"){const resolution=clean(Object.prototype.hasOwnProperty.call(r.__updates||{},"resolution_steps")?r.__updates.resolution_steps:r.resolution_steps);if(!resolution)throw Error("Resolution Steps is required when status is Closed for ticket "+clean(r["Tickets#"]));const closedOn=Object.prototype.hasOwnProperty.call(r.__updates||{},"Closedon")?r.__updates.Closedon:r.Closedon;if(!clean(closedOn))throw Error("Closed on is required when status is Closed for ticket "+clean(r["Tickets#"]));}}return true}
function mergeReviewRow(r){const out={...r},original=r.__original||{},updates=r.__updates||{};for(const f of["Status","Wing","Closedon","resolution_steps"])out[f]=Object.prototype.hasOwnProperty.call(updates,f)&&clean(updates[f])!==""?updates[f]:original[f]??out[f];delete out.__original;delete out.__updates;return out}
function recalculateDerivedFields(r){r.Type="SOC Alert";r.Classification=classify(r.subject);r.Organization=org(r.subject);return r}
function validateProcessedRows(){for(const r of processedRows){if(r.Closedon!==""&&r.Closedon!=null&&!(r.Closedon instanceof Date)){const d=parseITSMDate(r.Closedon);if(!d)throw Error("Invalid Closedon date for ticket "+clean(r["Tickets#"]));r.Closedon=d}}return true}
function finalize(){validateReviewRows();const map=new Map();processedRows.forEach(r=>map.set(clean(r["Tickets#"]),r));reviewRows.forEach(r=>{const merged=recalculateDerivedFields(mergeReviewRow(r)),id=clean(merged["Tickets#"]);if(id&&!map.has(id))map.set(id,merged)});processedRows=[...map.values()].filter(r=>r["Tickets#"]);validateProcessedRows();el.reviewPanel.classList.add("hidden");renderSummary();el.resultSection.classList.remove("hidden");el.fileName.textContent="Processed: "+processedRows.length+" final records for "+emailDate(processedRows[0]?.["Created on"]);el.dateSelectionStatus.textContent="Final output: "+processedRows.length+" records • Created review records included: "+reviewRows.length;el.resultSection.scrollIntoView({behavior:"smooth",block:"start"})}
function summarize(rows){const m=new Map();rows.forEach(r=>{const c=clean(r.Classification)||"NABFID DC";if(!m.has(c))m.set(c,{Closed:0,"Pending on COE":0,"Pending on customer":0,total:0});const s=m.get(c),x=status(r);if(x==="closed")s.Closed++;else if(x==="pending on coe")s["Pending on COE"]++;else if(x==="pending on customer")s["Pending on customer"]++;s.total++});return m}
function cells(tr,a){a.forEach(v=>{const td=document.createElement("td");td.textContent=v;tr.appendChild(td)})}
function renderSummary(){const m=summarize(processedRows),h=$("summaryHead"),b=$("summaryBody");h.innerHTML="<tr>"+SUMMARY.map(x=>"<th>"+x+"</th>").join("")+"</tr>";b.innerHTML="";const t={Closed:0,"Pending on COE":0,"Pending on customer":0,total:0};[...m.keys()].sort().forEach(c=>{const s=m.get(c);t.Closed+=s.Closed;t["Pending on COE"]+=s["Pending on COE"];t["Pending on customer"]+=s["Pending on customer"];t.total+=s.total;const tr=document.createElement("tr");cells(tr,[c,s.Closed,s["Pending on COE"],s["Pending on customer"],s.total]);b.appendChild(tr)});const tr=document.createElement("tr");cells(tr,["Grand Total",t.Closed,t["Pending on COE"],t["Pending on customer"],t.total]);b.appendChild(tr);renderEmail()}
const DERIVED_FIELDS=new Set(["Type","Classification","Organization"]);const EDITABLE_FIELDS=new Set(["Tickets#","Created on","Department","Prioritytitle","subject","Wing","Closedon","resolution_steps","Status"]);function renderData(edit){el.dataPanel.classList.remove("hidden");el.saveBtn.classList.toggle("hidden",!edit);el.dataTitle.textContent=edit?"Edit processed data":"Preview processed data";el.dataHead.innerHTML="<tr>"+OUT.map(h=>"<th>"+h+"</th>").join("")+"</tr>";el.dataBody.innerHTML="";processedRows.forEach((r,i)=>{const tr=document.createElement("tr");OUT.forEach(h=>{const td=document.createElement("td");if(edit&&EDITABLE_FIELDS.has(h)){const x=document.createElement("input");x.value=r[h] instanceof Date?displayDate(r[h]):r[h]||"";x.oninput=e=>r[h]=e.target.value;if(h==="Closedon")x.onblur=()=>{const raw=clean(x.value);if(!raw){r[h]="";x.classList.remove("input-invalid");return}const d=parseITSMDate(raw);if(d){r[h]=d;x.value=displayDate(d);x.classList.remove("input-invalid")}else x.classList.add("input-invalid")};td.appendChild(x)}else{td.textContent=h==="Created on"||h==="Closedon"?displayDate(r[h]):r[h]||"";if(edit&&DERIVED_FIELDS.has(h))td.title="Derived from subject"}tr.appendChild(td)});el.dataBody.appendChild(tr)})}function table(){return document.querySelector("#summaryHead")?.closest("table")}
function styleCopy(t){const base='font-family:"Inter 18pt",Inter,Arial,sans-serif;font-size:10pt';t.setAttribute("border","1");t.setAttribute("cellspacing","0");t.setAttribute("cellpadding","0");t.setAttribute("style","border-collapse:collapse;border-spacing:0;width:720px;min-width:720px;table-layout:fixed;"+base);t.querySelectorAll("tr").forEach((r,i,a)=>r.querySelectorAll("th,td").forEach((c,j)=>{const widths=["26%","12%","20%","24%","18%"];c.setAttribute("style","border:1px solid #202020;border-style:solid;border-width:1px;padding:4px 8px;white-space:nowrap;text-align:center;vertical-align:middle;"+base+";width:"+widths[j]+";background:"+(i===0||i===a.length-1?"#0B2A5B":"#FFFFFF")+";color:"+(i===0||i===a.length-1?"#FFFFFF":"#111827")+";font-weight:"+(i===0||i===a.length-1?"700":"400"));}));return t}
function fallback(t){const a=document.createElement("textarea");a.value=t;a.style.cssText="position:fixed;left:-9999px";document.body.appendChild(a);a.select();document.execCommand("copy");a.remove()}
async function copy(t,html){if(navigator.clipboard&&window.ClipboardItem){try{await navigator.clipboard.write([new ClipboardItem({"text/html":new Blob([html],{type:"text/html"}),"text/plain":new Blob([t],{type:"text/plain"})})]);return}catch(e){}}fallback(t)}
function copied(b,s){const o=b.textContent;b.textContent=s;setTimeout(()=>b.textContent=o,1600)}
async function copySummary(){const t=table();if(!t)return;const text=[...t.rows].map(r=>[...r.cells].map(c=>c.textContent.trim()).join("\t")).join("\n");await copy(text,styleCopy(t.cloneNode(true)).outerHTML);copied(el.copyTableBtn,"Copied ✓")}
function renderEmail(){if(!processedRows.length)return;const d=emailDate(processedRows[0]["Created on"]),t=table()?.cloneNode(true);if(!d||!t)return;styleCopy(t);const base='font-family:"Inter 18pt",Inter,Arial,sans-serif;font-size:10pt;line-height:1.35';const c=[...summarize(processedRows).keys()].sort();const ct=c.length<=1?(c[0]||""):c.slice(0,-1).join(", ")+" and "+c.at(-1);const sub=document.createElement("div");sub.className="email-subject-preview";sub.textContent="Subject: NABFID Daily Offense Report || "+d;const body=document.createElement("div");body.className="email-preview-body";body.setAttribute("style",base+";color:#111827");const p1=document.createElement("p");p1.setAttribute("style",base+";margin:0 0 12px 0");p1.textContent="Hi Team,";const p2=document.createElement("p");p2.setAttribute("style",base+";margin:0 0 12px 0");p2.append(document.createTextNode("Please find the attached "+ct+" daily offense report for "));const strong=document.createElement("strong");strong.setAttribute("style",base+";font-weight:700");strong.textContent=d;p2.append(strong,document.createTextNode("."));body.append(p1,p2,t);el.emailPreview.replaceChildren(sub,body)}
async function copySubject(){if(!processedRows.length)return;await copy("NABFID Daily Offense Report || "+emailDate(processedRows[0]["Created on"]),"");copied(el.copySubjectBtn,"Subject Copied ✓")}
async function copyEmail(){renderEmail();const b=el.emailPreview.querySelector(".email-preview-body");if(!b)return;const t=[...b.querySelectorAll("p,tr")].map(x=>x.textContent.trim()).join("\n");await copy(t,b.outerHTML);copied(el.copyEmailBtn,"Email Copied ✓")}
function borders(){const s={style:"thin",color:{rgb:"202020"}};return{top:s,bottom:s,left:s,right:s}}
function styleSheet(sh,total=-1){const rg=XLSX.utils.decode_range(sh["!ref"]);for(let r=rg.s.r;r<=rg.e.r;r++)for(let c=rg.s.c;c<=rg.e.c;c++){const x=sh[XLSX.utils.encode_cell({r,c})];if(!x)continue;const isDateColumn=sh[XLSX.utils.encode_cell({r:0,c})]?.v==="Created on"||sh[XLSX.utils.encode_cell({r:0,c})]?.v==="Closedon";x.s=r===0?{font:{name:"Inter 18pt",sz:10,bold:true,color:{rgb:"FFFFFF"}},fill:{fgColor:{rgb:"0B2A5B"}},alignment:{horizontal:"center",vertical:"center",wrapText:false},border:borders()}:{font:{name:"Inter 18pt",sz:10,color:{rgb:r===total?"FFFFFF":"374151"}},fill:{fgColor:{rgb:r===total?"0B2A5B":"FFFFFF"}},alignment:{horizontal:"center",vertical:"center",wrapText:false},border:borders(),...(isDateColumn&&x.v instanceof Date?{numFmt:excelDateFormat(dateOutputFormat)}:{})}}}

function download(){if(!processedRows.length)return;const wb=XLSX.utils.book_new(),s=XLSX.utils.json_to_sheet(processedRows);for(const col of["B","J"])for(let r=2;r<=processedRows.length+1;r++){const c=s[col+r];if(c&&c.v instanceof Date)c.z="dd-mm-yyyy hh:mm:ss"}s["!freeze"]={xSplit:0,ySplit:1};s["!autofilter"]={ref:s["!ref"]};s["!cols"]=OUT.map(h=>({wch:({subject:42,Organization:38,resolution_steps:55,"Tickets#":16,"Created on":20,"Closedon":20,Wing:20}[h]||18)}));styleSheet(s);XLSX.utils.book_append_sheet(wb,s,"Offenses");const m=summarize(processedRows),rows=[SUMMARY],tot=[0,0,0,0];[...m.keys()].sort().forEach(c=>{const x=m.get(c);rows.push([c,x.Closed,x["Pending on COE"],x["Pending on customer"],x.total]);tot[0]+=x.Closed;tot[1]+=x["Pending on COE"];tot[2]+=x["Pending on customer"];tot[3]+=x.total});rows.push(["Grand Total",...tot]);const ss=XLSX.utils.aoa_to_sheet(rows);ss["!cols"]=[{wch:30},{wch:14},{wch:20},{wch:24},{wch:16}];styleSheet(ss,rows.length-1);XLSX.utils.book_append_sheet(wb,ss,"Classification");const p=parts(processedRows[0]["Created on"]),name=p?"NaBFID Offenses "+ordinal(p.day)+" "+["Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sept","Oct","Nov","Dec"][p.month-1]+".xlsx":"NaBFID Offenses.xlsx";const blob=new Blob([XLSX.write(wb,{bookType:"xlsx",type:"array"})],{type:"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"}),a=document.createElement("a"),u=URL.createObjectURL(blob);a.href=u;a.download=name;a.click();URL.revokeObjectURL(u)}
function bind(){document.addEventListener("click",e=>{if(!e.target.closest(".review-dropdown"))closeReviewDropdowns()});el.fileInput.onchange=e=>handleFiles(e.target.files);el.dropzone.ondragover=e=>{e.preventDefault();el.dropzone.classList.add("drop-active")};el.dropzone.ondragleave=()=>el.dropzone.classList.remove("drop-active");el.dropzone.ondrop=e=>{e.preventDefault();el.dropzone.classList.remove("drop-active");handleFiles(e.dataTransfer.files)};el.datePickerButton.onclick=()=>{if(el.datePickerPopover.classList.contains("hidden")){el.datePickerPopover.classList.remove("hidden");el.datePickerButton.setAttribute("aria-expanded","true");calendar()}else closeCal()};document.addEventListener("click",e=>{if(!el.datePickerPopover.contains(e.target)&&!el.datePickerButton.contains(e.target))closeCal()});el.processBtn.onclick=processDate;el.continueBtn.onclick=finalize;el.cancelReviewBtn.onclick=()=>el.reviewPanel.classList.add("hidden");el.previewBtn.onclick=()=>{renderData(false);el.dataPanel.scrollIntoView({behavior:"smooth"})};el.editBtn.onclick=()=>{renderData(true);el.dataPanel.scrollIntoView({behavior:"smooth"})};el.copyTableBtn.onclick=copySummary;el.downloadBtn.onclick=download;el.copySubjectBtn.onclick=copySubject;el.copyEmailBtn.onclick=copyEmail;el.saveBtn.onclick=()=>{try{processedRows=processedRows.map(r=>recalculateDerivedFields(r));validateProcessedRows();renderSummary();renderData(true)}catch(e){alert(e.message||e)}};el.closePanelBtn.onclick=()=>el.dataPanel.classList.add("hidden")}
bind();
})();