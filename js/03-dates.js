function excelSerialParts(n){if(!Number.isFinite(n)||n<=0||n>=100000)return null;const d=XLSX?.SSF?.parse_date_code?XLSX.SSF.parse_date_code(n):null;if(d?.y&&d?.m&&d?.d)return{day:d.d,month:d.m,year:d.y,hour:d.H||0,minute:d.M||0,second:Math.floor(d.S||0)};const ms=(n-25569)*86400000;const x=new Date(ms);if(isNaN(x.getTime()))return null;return{day:x.getUTCDate(),month:x.getUTCMonth()+1,year:x.getUTCFullYear(),hour:x.getUTCHours(),minute:x.getUTCMinutes(),second:x.getUTCSeconds()}}

function validDateParts(day,month,year){const d=new Date(year,month-1,day);return d.getFullYear()===year&&d.getMonth()===month-1&&d.getDate()===day}

function dateFromParts(day,month,year,hour=0,minute=0,second=0){if(!validDateParts(day,month,year)||hour<0||hour>23||minute<0||minute>59||second<0||second>59)return null;return new Date(year,month-1,day,hour,minute,second)}

function parseTextParts(v){const s=clean(v).replace(/\u00a0/g," ").replace(/\s+/g," ");if(!s)return null;let day,month,year;let m=s.match(/^(\d{4})[\/\-.](\d{1,2})[\/\-.](\d{1,2})/);if(m){year=+m[1];month=+m[2];day=+m[3]}else{m=s.match(/^(\d{1,2})[\/\-.](\d{1,2})[\/\-.](\d{4})/);if(!m)return null;const a=+m[1],b=+m[2];year=+m[3];if(a>12&&b<=12){day=a;month=b}else if(b>12&&a<=12){month=a;day=b}else{month=a;day=b}}const tm=s.match(/(?:T|\s+)(\d{1,2}):(\d{2})(?::(\d{2}))?\s*(AM|PM)?/i);let hour=0,minute=0,second=0;if(tm){hour=+tm[1];minute=+tm[2];second=+(tm[3]||0);const ap=(tm[4]||"").toUpperCase();if(ap){if(hour<1||hour>12)return null;if(hour===12)hour=0;if(ap==="PM")hour+=12}else if(hour>23)return null}return validDateParts(day,month,year)?{day,month,year,hour,minute,second}:null}

function parts(v){if(v instanceof Date&&!isNaN(v.getTime()))return{day:v.getDate(),month:v.getMonth()+1,year:v.getFullYear(),hour:v.getHours(),minute:v.getMinutes(),second:v.getSeconds()};if(typeof v==="number"){const p=excelSerialParts(v);if(p)return p}const s=clean(v);if(!s)return null;if(/^\d+(?:\.\d+)?$/.test(s)){const p=excelSerialParts(+s);if(p)return p}return parseTextParts(s)}

function dateVal(v){if(v instanceof Date&&!isNaN(v.getTime()))return v;const p=parts(v);if(p)return dateFromParts(p.day,p.month,p.year,p.hour,p.minute,p.second)||"";const d=new Date(clean(v));return isNaN(d.getTime())?clean(v):d}

function key(v){const p=parts(v);return p?[p.year,String(p.month).padStart(2,"0"),String(p.day).padStart(2,"0")].join("-"):""}

function displayDate(v){const d=dateVal(v);if(!(d instanceof Date)||isNaN(d.getTime()))return clean(v);let h=d.getHours(),a=h>=12?"PM":"AM";h=h%12||12;return d.getDate()+"-"+(d.getMonth()+1)+"-"+d.getFullYear()+" "+h+":"+String(d.getMinutes()).padStart(2,"0")+":"+String(d.getSeconds()).padStart(2,"0")+" "+a}

function emailDate(v){const p=parts(v);if(!p)return"";return ordinal(p.day)+" "+["January","February","March","April","May","June","July","August","September","October","November","December"][p.month-1]+" "+p.year}

function parseITSMDate(v){if(v instanceof Date&&!isNaN(v.getTime()))return v;const s=clean(v);if(!s)return null;const m=s.match(/^(\d{1,2})[-\/](\d{1,2})[-\/](\d{4})(?:\s+(\d{1,2}):(\d{2})(?::(\d{2}))?\s*(AM|PM)?)?$/i);if(!m)return null;let h=+(m[4]||0),mi=+(m[5]||0),sec=+(m[6]||0),ap=(m[7]||"").toUpperCase();if(ap){if(h<1||h>12)return null;if(h===12)h=0;if(ap==="PM")h+=12}return dateFromParts(+m[1],+m[2],+m[3],h,mi,sec)}