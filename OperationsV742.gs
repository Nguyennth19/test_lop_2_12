/** V7.4.2: quality checks and manager-only diagnostics. Never exposes private records publicly. */
const QUALITY_V742 = Object.freeze({
 sheet:'KIEM_TRA_DU_LIEU',
 headers:['Mã cảnh báo','Phát hiện lúc','Mức độ','Nguồn dữ liệu','Mã bản ghi','Nội dung','Trạng thái','Cập nhật lúc']
});
function prepareQualityV742(){
 const ss=ss_();let sh=ss.getSheetByName(QUALITY_V742.sheet),created=false;
 if(!sh){sh=ss.insertSheet(QUALITY_V742.sheet);created=true;}
 const h=QUALITY_V742.headers;
 if(sh.getMaxColumns()<h.length)sh.insertColumnsAfter(sh.getMaxColumns(),h.length-sh.getMaxColumns());
 const old=sh.getRange(1,1,1,h.length).getDisplayValues()[0];
 if(old.some((v,i)=>v&&v!==h[i]))throw Error('Sheet kiểm tra đã có nội dung khác chuẩn. Dừng để bảo vệ dữ liệu.');
 sh.getRange(1,1,1,h.length).setValues([h]);sh.setFrozenRows(1);
 return {ok:true,created,sheet:QUALITY_V742.sheet};
}
function qualityIssuesV742_(){
 const issues=[],add=(level,source,id,code,desc)=>issues.push({
  key:[source,id||'-',code].join('|'),level,source,id:String(id||''),description:desc
 });
 const students=roster_(),ids=new Set(),dup=new Set();
 students.forEach(x=>{if(!x.id)add('CAO',APP.sheets.students,x.sheetRow,'MISSING_STUDENT_ID','Thiếu mã học sinh');
 else if(ids.has(x.id))dup.add(x.id);else ids.add(x.id);
 if(x.status==='UNKNOWN')add('VUA',APP.sheets.students,x.id||x.sheetRow,'STATUS','Trạng thái học sinh không hợp lệ');
 if(!x.parent)add('THAP',APP.sheets.students,x.id||x.sheetRow,'PARENT','Chưa có tên phụ huynh');});
 dup.forEach(id=>add('CAO',APP.sheets.students,id,'DUP_STUDENT','Mã học sinh trùng'));
 [[APP.sheets.cash,9],[APP.sheets.expenses,9],[APP.sheets.goods,11]].forEach(([name,statusIx])=>{
  const all=rows_(name),seen=new Set();
  all.forEach(r=>{
   const v=r.v,id=String(v[0]||'').trim(),active=active_(r,statusIx);
   if(seen.has(id))add('CAO',name,id,'DUP_ID','Mã giao dịch bị trùng');seen.add(id);
   if(!active)return;
   if(name!==APP.sheets.expenses&&!ids.has(String(v[2]||'')))add('CAO',name,id,'NO_STUDENT','Không tìm thấy mã học sinh liên kết');
   if(name===APP.sheets.expenses){if(!APP.categories.includes(String(v[2]||'')))add('VUA',name,id,'CATEGORY','Nhóm khoản chi không nằm trong danh mục');
     if(!v[6])add('THAP',name,id,'RECEIPT','Khoản chi chưa có chứng từ');}
   if(name===APP.sheets.goods&&String(v[10]||'').toUpperCase()==='PUBLIC'&&!v[8])add('THAP',name,id,'GOODS_IMAGE','Hiện vật công khai chưa có ảnh');
  });
 });
 return issues;
}
function recordQualityV742_(issues){
 const sh=ss_().getSheetByName(QUALITY_V742.sheet);
 if(!sh)throw Error('Cần chạy prepareQualityV742() trước.');
 const rows=sh.getLastRow()>1?sh.getRange(2,1,sh.getLastRow()-1,8).getValues():[];
 const index=new Map(rows.map((r,i)=>[String(r[0]),{row:i+2,v:r}]));
 const now=new Date(),active=new Set();
 for(const item of issues){
  active.add(item.key);
  const previous=index.get(item.key);
  if(previous){
   // Retain first-seen date and do not overwrite manual notes in other columns.
   sh.getRange(previous.row,3,1,6).setValues([[item.level,item.source,item.id,item.description,'ĐANG MỞ',now]]);
  } else sh.appendRow([item.key,now,item.level,item.source,item.id,item.description,'ĐANG MỞ',now]);
 }
 for(const [key,p] of index)if(!active.has(key)&&String(p.v[6])==='ĐANG MỞ'){
  sh.getRange(p.row,7,1,2).setValues([['ĐÃ XỬ LÝ',now]]);
 }
}
function runQualityV742_(){
 const lock=LockService.getScriptLock();lock.waitLock(30000);
 try{prepareQualityV742();const issues=qualityIssuesV742_();recordQualityV742_(issues);
  return {ok:true,total:issues.length,high:issues.filter(i=>i.level==='CAO').length,
   medium:issues.filter(i=>i.level==='VUA').length,low:issues.filter(i=>i.level==='THAP').length};}
 finally{lock.releaseLock();}
}
/** Manual diagnostic for Admin; no private payer, bank notes or phone numbers returned. */
function adminQualityV742(token){
 authV6_(token,['ADMIN']);const issues=qualityIssuesV742_();
 const records=ss_().getSheetByName(QUALITY_V742.sheet);
 const lastRun=records&&records.getLastRow()>1?records.getRange(records.getLastRow(),8).getDisplayValue():'Chưa chạy';
 return {lastRun,total:issues.length,high:issues.filter(i=>i.level==='CAO').length,
  medium:issues.filter(i=>i.level==='VUA').length,low:issues.filter(i=>i.level==='THAP').length,
  recent:issues.slice(0,30).map(x=>({level:x.level,source:x.source,id:x.id,description:x.description}))};
}
function runQualityAdminV742(token){authV6_(token,['ADMIN']);return runQualityV742_();}
function weeklyQualityV742(){return runQualityV742_();}
function installWeeklyQualityV742(){
 // Run once as script owner; scheduled checks use script owner identity.
 const fn='weeklyQualityV742';
 ScriptApp.getProjectTriggers().filter(x=>x.getHandlerFunction()===fn).forEach(trigger=>ScriptApp.deleteTrigger(trigger));
 ScriptApp.newTrigger(fn).timeBased().onWeekDay(ScriptApp.WeekDay.SUNDAY).atHour(2).everyWeeks(1).inTimezone('Asia/Ho_Chi_Minh').create();
 return {ok:true,frequency:'Chủ nhật khoảng 02:00 (Việt Nam)'};
}
/** Manual QR metadata audit. Does not publish Drive sharing settings or QR contents. */
function checkMediaSetupV742(){
 const list=[
  {type:'QR',id:'1tHdkBZZKBfNkpfd4_VZtpHJaU8OO75Fm'},
  {type:'THU',id:'1mtiwgc690bRce_ADCQ7eeRdcxRKMGa_K'},
  {type:'HIEN_VAT',id:'1E-TVsFP6VsqHWXiudxycYvf6dl3g_U_8'},
  {type:'CHI',id:'1WeBf1AJr47Z2s7s3GLzHRM3nd6DtB02v'}
 ];
 return list.map(item=>{try{const f=item.type==='QR'?DriveApp.getFileById(item.id):DriveApp.getFolderById(item.id);
 return {kind:item.type,ok:!!f.getName(),mime:item.type==='QR'?f.getMimeType():'folder'};}
 catch(e){return {kind:item.type,ok:false,problem:'Không đọc được file/thư mục từ tài khoản đang chạy'};}});
}
function adminMediaStatusV742(token){authV6_(token,['ADMIN']);return checkMediaSetupV742();}

/**
 * V7.4.3: bounded, read-only inspection of referenced Drive files.
 * Does not download content, alter sharing, or expose names/URLs.
 * A cursor allows the owner to scan in small batches without timeouts.
 */
function inspectReceiptRefsV743_(offset,limit){
 const sources=[
  {sheet:APP.sheets.cash,idIndex:8,statusIndex:9,label:'Chứng từ thu'},
  {sheet:APP.sheets.expenses,idIndex:6,statusIndex:9,label:'Chứng từ chi'},
  {sheet:APP.sheets.expenses,idIndex:7,statusIndex:9,label:'Ảnh sản phẩm'},
  {sheet:APP.sheets.goods,idIndex:8,statusIndex:11,label:'Ảnh hiện vật'}
 ];
 const records=[];
 sources.forEach(s=>rows_(s.sheet).forEach(r=>{
  if(!active_(r,s.statusIndex))return;
  const fileId=String(r.v[s.idIndex]||'').trim();
  if(fileId)records.push({source:s.sheet,label:s.label,recordId:String(r.v[0]),fileId});
 }));
 const begin=Math.max(0,Math.floor(Number(offset)||0));
 const count=Math.min(30,Math.max(1,Math.floor(Number(limit)||15)));
 const batch=records.slice(begin,begin+count);
 const issues=[],statuses=[];
 for(const item of batch){
  try{
   const file=DriveApp.getFileById(item.fileId);
   const mime=file.getMimeType(),size=file.getSize();
   const allowed=/^image\/(jpeg|png|gif|webp)$/i.test(mime)||
    (item.label==='Chứng từ chi'&&mime==='application/pdf');
   if(!allowed)issues.push({code:'UNSUPPORTED_FILE',source:item.source,id:item.recordId,type:item.label});
   if(size>APP.maxFileBytes)issues.push({code:'LARGE_FILE',source:item.source,id:item.recordId,type:item.label});
   statuses.push({source:item.source,id:item.recordId,ok:allowed&&size<=APP.maxFileBytes});
  }catch(e){
   issues.push({code:'FILE_NOT_ACCESSIBLE',source:item.source,id:item.recordId,type:item.label});
   statuses.push({source:item.source,id:item.recordId,ok:false});
  }
 }
 return {ok:issues.length===0,checked:batch.length,total:records.length,
  nextOffset:begin+batch.length<records.length?begin+batch.length:null,issues,statuses};
}
function adminReceiptAuditV743(token,offset,limit){
 authV6_(token,['ADMIN']);
 return inspectReceiptRefsV743_(offset,limit);
}
/** Owner-only, read-only command, suitable for the Script Editor. */
function inspectReceiptRefsManualV743(){
 const output=inspectReceiptRefsV743_(0,30);
 console.log(JSON.stringify({checked:output.checked,total:output.total,
  nextOffset:output.nextOffset,issues:output.issues}));
 return output;
}
/** Owner-only summary; no finance writes or migration. */
function preDeployReviewV743(){
 const cash=rows_(APP.sheets.cash).filter(r=>active_(r,9));
 const expenses=rows_(APP.sheets.expenses).filter(r=>active_(r,9));
 const goods=rows_(APP.sheets.goods).filter(r=>active_(r,11));
 const byId=(rows)=>new Set(rows.map(r=>String(r.v[0]))).size===rows.length;
 const sum=(rows,index)=>rows.reduce((a,r)=>a+Number(r.v[index]||0),0);
 const totalCash=sum(cash,5),totalExpense=sum(expenses,4);
 const issues=qualityIssuesV742_();
 return {timestamp:new Date().toISOString(),finance:{
  cashCount:cash.length,cashTotal:totalCash,expenseCount:expenses.length,
  expenseTotal:totalExpense,balance:totalCash-totalExpense,goodsCount:goods.length
 },uniqueIds:{cash:byId(cash),expenses:byId(expenses),goods:byId(goods)},
 quality:{total:issues.length,high:issues.filter(x=>x.level==='CAO').length,
  medium:issues.filter(x=>x.level==='VUA').length,low:issues.filter(x=>x.level==='THAP').length},
 auditHeadersReady:sh_(APP.sheets.audit).getRange(1,6,1,6).getDisplayValues()[0]
  .join('|')===['Mã nhật ký','Bảng dữ liệu','Giá trị trước (JSON)','Giá trị sau (JSON)','Lý do điều chỉnh','Múi giờ'].join('|')
 };
}
