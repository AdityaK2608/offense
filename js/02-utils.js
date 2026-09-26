function clean(v){return v==null?"":String(v).trim()}

function val(r,c){const k=Object.keys(r||{}).find(h=>clean(h).toLowerCase()===c.toLowerCase());return k?clean(r[k]):""}

function rawVal(r,c){const k=Object.keys(r||{}).find(h=>clean(h).toLowerCase()===c.toLowerCase());return k?r[k]:""}

function has(rows,c){return rows.length&&Object.keys(rows[0]).some(h=>clean(h).toLowerCase()===c.toLowerCase())}

function ordinal(n){return n%100>=11&&n%100<=13?n+"th":n+({1:"st",2:"nd",3:"rd"}[n%10]||"th")}

function emailDate(v){const p=parts(v);if(!p)return"";return ordinal(p.day)+" "+["January","February","March","April","May","June","July","August","September","October","November","December"][p.month-1]+" "+p.year}

function classify(s){const m=clean(s).match(/Domain:\s*([^|]*)/i);return m?clean(m[1])||"NABFID DC":"NABFID DC"}

function org(s){return /KPMG/i.test(clean(s))?ORG.kpmg:ORG.default}

function nab(r){return /nabfid/i.test(val(r,"subject"))||/nabfid/i.test(val(r,"Organization"))}

function soc(r){return val(r,"Type").toLowerCase()==="soc alert"}

function excluded(r){const s=val(r,"subject");return s&&EXCLUDED.some(p=>p.test(s))}

function status(r){return clean(val(r,"Status")).replace(/\s+/g," ").toLowerCase()}

function output(r){const s=val(r,"subject");return{"Tickets#":val(r,"Tickets#"),"Created on":dateVal(val(r,"Created on")),"Department":val(r,"Department"),"Prioritytitle":val(r,"Prioritytitle"),"Type":"SOC Alert",subject:s,"Classification":classify(s),"Organization":org(s),"Wing":val(r,"Wing"),"Closedon":dateVal(val(r,"Closedon")),"resolution_steps":val(r,"resolution_steps"),"Status":val(r,"Status")}}

function valid(rows,name){if(!rows.length)throw Error(name+" file is empty.");const headers=Object.keys(rows[0]||{});const normalized=new Map();headers.forEach(h=>{const k=clean(h).toLowerCase();if(!k)return;const list=normalized.get(k)||[];list.push(h);normalized.set(k,list)});for(const [k,list] of normalized)if(list.length>1)throw Error(name+" file contains duplicate column headers: "+list.join(", ")+".");for(const c of REQUIRED_HEADERS)if(!has(rows,c))throw Error(name+' file is missing "'+c+'".')}

function cells(tr,a){a.forEach(v=>{const td=document.createElement("td");td.textContent=v;tr.appendChild(td)})}