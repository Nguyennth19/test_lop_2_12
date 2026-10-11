function ss_(){const x=SpreadsheetApp.getActiveSpreadsheet();if(!x)throw Error('Hãy gắn Apps Script vào đúng Google Sheets Thu chi.');return x;}
function sh_(name){const x=ss_().getSheetByName(name);if(!x)throw Error('Thiếu tab '+name+'. Chạy migrateV6 trong Script Editor.');return x;}
function dateV6_(v){if(v instanceof Date&&!isNaN(v))return Utilities.formatDate(v,Session.getScriptTimeZone(),'yyyy-MM-dd');const s=String(v||'').trim();if(/^\d{4}-\d{2}-\d{2}$/.test(s))return s;const m=s.match(/^(\d{1,2})[\/.\-](\d{1,2})[\/.\-](\d{4})$/);return m?m[3]+'-'+m[2].padStart(2,'0')+'-'+m[1].padStart(2,'0'):'';}
function inputDate_(v){const s=dateV6_(v),d=new Date(s+'T12:00:00');if(!s||isNaN(d.getTime())||dateV6_(d)!==s)throw Error('Ngày không hợp lệ.');return d;}
function money_(x){let n=typeof x==='number'?x:Number(String(x||'').replace(/[^0-9-]/g,''));if(!Number.isSafeInteger(n)||n<0||n>1e12)throw Error('Số tiền không hợp lệ.');return n;}
function clean_(x,max){let t=String(x||'').trim();if(t.length>max)throw Error('Nội dung quá dài.');if(/^[=+@]/.test(t))t="'"+t;return t;}
function required_(x,max){let s=clean_(x,max);if(!s)throw Error('Vui lòng nhập đầy đủ thông tin bắt buộc.');return s;}
function rows_(name){const s=sh_(name),last=s.getLastRow();if(last<2)return [];return s.getRange(2,1,last-1,APP.headers[name]?.length||s.getLastColumn()).getValues().map((v,i)=>({row:i+2,v})).filter(x=>String(x.v[0]||'').trim());}
function active_(r,index){return String(r.v[index]||'ACTIVE').toUpperCase()!=='DELETED';}
function find_(name,id){return rows_(name).find(r=>String(r.v[0])===String(id));}
function append_(name,values){const lock=LockService.getScriptLock();lock.waitLock(30000);try{sh_(name).appendRow(values);}finally{lock.releaseLock();}}
function setRow_(name,row,values){const lock=LockService.getScriptLock();lock.waitLock(30000);try{sh_(name).getRange(row,1,1,values.length).setValues([values]);}finally{lock.releaseLock();}}
function newId_(prefix){return prefix+'-'+Utilities.formatDate(new Date(),Session.getScriptTimeZone(),'yyyyMMdd-HHmmss')+'-'+Utilities.getUuid().slice(0,7).toUpperCase();}
function roster_(){const s=sh_(APP.sheets.students),last=s.getLastRow();if(last<4)return [];return s.getRange(4,1,last-3,8).getValues().filter(r=>String(r[1]||'').trim()).map(r=>({id:String(r[0]||'').trim(),child:String(r[1]||'').trim(),parent:String(r[4]||'').trim(),phone:String(r[5]||'').trim(),status:String(r[7]||'').trim()}));}
function auditV6_(email,action,id,detail){sh_(APP.sheets.audit).appendRow([new Date(),email,action,id||'',String(detail||'').slice(0,600)]);}
function toDateText_(v){return dateV6_(v);}
