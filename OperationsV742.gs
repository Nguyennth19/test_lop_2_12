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
 ScriptApp.getProjectTriggers().filter(x=>x.getHandlerFunction()===fn).forEach(ScriptApp.deleteTrigger);
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
