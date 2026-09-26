function styleCopy(t){t.style.cssText="border-collapse:collapse;width:auto;font-family:Inter,Arial,sans-serif;font-size:10pt";[...t.querySelectorAll("tr")].forEach((r,i,a)=>r.querySelectorAll("th,td").forEach(c=>c.style.cssText="border:1px solid #202020;padding:4px 8px;white-space:nowrap;text-align:center;background:"+(i===0||i===a.length-1?"#0B2A5B":"#FFFFFF")+";color:"+(i===0||i===a.length-1?"#FFFFFF":"#111827")+";font-weight:"+(i===0||i===a.length-1?"700":"400")));return t}

function fallback(t){const a=document.createElement("textarea");a.value=t;a.style.cssText="position:fixed;left:-9999px";document.body.appendChild(a);a.select();document.execCommand("copy");a.remove()}

async function copy(t,html){if(navigator.clipboard&&window.ClipboardItem){try{await navigator.clipboard.write([new ClipboardItem({"text/html":new Blob([html],{type:"text/html"}),"text/plain":new Blob([t],{type:"text/plain"})})]);return}catch(e){}}fallback(t)}

function copied(b,s){const o=b.textContent;b.textContent=s;setTimeout(()=>b.textContent=o,1600)}

async function copySummary(){const t=table();if(!t)return;const text=[...t.rows].map(r=>[...r.cells].map(c=>c.textContent.trim()).join("\t")).join("\n");await copy(text,styleCopy(t.cloneNode(true)).outerHTML);copied(el.copyTableBtn,"Copied ✓")}

async function copySubject(){if(!processedRows.length)return;await copy("NABFID Daily Offense Report || "+emailDate(processedRows[0]["Created on"]),"");copied(el.copySubjectBtn,"Subject Copied ✓")}

async function copyEmail(){renderEmail();const b=el.emailPreview.querySelector(".email-preview-body");if(!b)return;const t=[...b.querySelectorAll("p,tr")].map(x=>x.textContent.trim()).join("\n");await copy(t,b.outerHTML);copied(el.copyEmailBtn,"Email Copied ✓")}