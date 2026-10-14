/** V7.4.4: Phiếu thu display metadata. No new sheets or changes to finance values. */
const CASH_V744=Object.freeze({
 defaultBatch:'Ủng hộ quỹ lớp 2026–2027',
 nameHeader:'Tên học sinh',
 titleHeader:'Chức vụ khi tạo',
 accountTitleHeader:'Chức vụ'
});
function normalizeActorPhoneV744_(value){
 const s=String(value==null?'':value).trim();
 if(s==='NHAP_CU_2026')return s;
 try{return normalizePhoneV6_(s);}catch(e){}
 if(/^[35789]\d{8}$/.test(s)){
  const candidate='0'+s;
  const accounts=rows_(APP.sheets.users);
  if(accounts.some(r=>String(r.v[0]||'').replace(/\D/g,'')===candidate))return candidate;
 }
 return s;
}
function accountTitleV744_(actor){
 const key=normalizeActorPhoneV744_(actor);
 if(key==='NHAP_CU_2026')return 'Dữ liệu nhập lịch sử';
 if(!/^0\d{9}$/.test(key))return '';
 const sheet=sh_('TAI_KHOAN');
 if(sheet.getLastRow()<2)return '';
 const rows=sheet.getRange(2,1,sheet.getLastRow()-1,7).getDisplayValues();
 const match=rows.find(r=>normalizeActorPhoneV744_(r[0])===key);
 return match?String(match[6]||'').trim().slice(0,80):'';
}
function defaultCashBatchV744_(){
 const sh=sh_(APP.sheets.settings);
 const rows=sh.getLastRow()>1?sh.getRange(2,1,sh.getLastRow()-1,2).getDisplayValues():[];
 const found=rows.find(r=>String(r[0]||'').trim()==='DEFAULT_CASH_BATCH');
 return found&&String(found[1]||'').trim()||CASH_V744.defaultBatch;
}
/** Admin/editor login only, used when a new cash form opens. */
function cashFormDefaultsV744(token){authV6_(token,['ADMIN','EDITOR']);return {batch:defaultCashBatchV744_()};}
/** Semicolon separators are required by this spreadsheet's Vietnamese formula locale. */
function cashStudentFormulaV744_(row){
 return '=IF(C'+row+'="";"";IFERROR(VLOOKUP(VALUE(C'+row+');\'DANH SÁCH HỌC SINH\'!$A$4:$B;2;FALSE);"Chưa xác định"))';
}
function migrateCashDisplayV744(){
 const ss=ss_(),cash=sh_(APP.sheets.cash),accounts=sh_('TAI_KHOAN'),settings=sh_(APP.sheets.settings);
 if(cash.getRange(1,1,1,13).getDisplayValues()[0].join('|')!==APP.headers.THU_TIEN.slice(0,13).join('|'))
  throw Error('THU_TIEN A:M khác cấu trúc chuẩn. Dừng migration.');
 const cashHeader=cash.getRange(1,14,1,2).getDisplayValues()[0];
 if((cashHeader[0]&&cashHeader[0]!==CASH_V744.nameHeader)||
    (cashHeader[1]&&cashHeader[1]!==CASH_V744.titleHeader))
  throw Error('THU_TIEN N:O có dữ liệu không tương thích.');
 if(accounts.getRange(1,7).getDisplayValue()&&accounts.getRange(1,7).getDisplayValue()!==CASH_V744.accountTitleHeader)
  throw Error('TAI_KHOAN G đã được sử dụng. Không ghi đè.');
 const count=Math.max(0,cash.getLastRow()-1);
 // Confirm no pre-existing N/O data before writing.
 if(count){
  const existing=cash.getRange(2,14,count,2).getDisplayValues();
  if(existing.some(r=>r[0]&&!r[0].startsWith('=IF(')&&r[0]!=='Chưa xác định'))
   throw Error('THU_TIEN N có nội dung; cần kiểm tra thủ công trước khi chạy migration.');
 }
 cash.getRange(1,14,1,2).setValues([[CASH_V744.nameHeader,CASH_V744.titleHeader]]);
 accounts.getRange(1,7).setValue(CASH_V744.accountTitleHeader);
 accounts.getRange('A:A').setNumberFormat('@');cash.getRange('K:K').setNumberFormat('@');
 if(count){
  // Populate N by formulas and preserve O (position snapshot) when already populated.
  const formulas=Array.from({length:count},(_,i)=>[cashStudentFormulaV744_(i+2)]);
  cash.getRange(2,14,count,1).setFormulas(formulas);
  // Normalize legacy numeric phone only if it matches a real USERS account.
  const creatorValues=cash.getRange(2,11,count,1).getDisplayValues();
  creatorValues.forEach((r,i)=>{
   const raw=r[0],normalized=normalizeActorPhoneV744_(raw);
   if(normalized!==raw)cash.getRange(i+2,11).setValue(normalized);
  });
 }
 const keys=settings.getLastRow()>1?settings.getRange(2,1,settings.getLastRow()-1,1).getDisplayValues().flat():[];
 if(!keys.includes('DEFAULT_CASH_BATCH'))
  settings.getRange(settings.getLastRow()+1,1,1,3).setValues([['DEFAULT_CASH_BATCH',CASH_V744.defaultBatch,'Đợt thu mặc định, có thể sửa khi nhập']]);
 return {ok:true,rows:count,headers:['N: '+CASH_V744.nameHeader,'O: '+CASH_V744.titleHeader,'TAI_KHOAN!G: '+CASH_V744.accountTitleHeader]};
}
