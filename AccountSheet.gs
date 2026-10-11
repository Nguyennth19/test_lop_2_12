/** V6 Account Sheet Manager. Passwords must never be stored in cells.
 * Install in the original spreadsheet-bound Apps Script project.
 */
function onOpen(){
  SpreadsheetApp.getUi().createMenu('🔐 Quản lý tài khoản')
    .addItem('Tạo / đặt lại mật khẩu', 'openAccountSheetManager')
    .addItem('Khóa / mở tài khoản', 'openAccountSheetManager')
    .addToUi();
}
function openAccountSheetManager(){
  SpreadsheetApp.getUi().showSidebar(HtmlService.createHtmlOutputFromFile('AccountSidebar').setTitle('Tài khoản & phân quyền'));
}
function accountActor_(){
  const caller=String(Session.getActiveUser().getEmail()||'').trim().toLowerCase();
  if(!caller)throw Error('Không xác định được tài khoản Google. Vui lòng mở Google Sheets bằng tài khoản quản trị.');
  const owner=DriveApp.getFileById(ss_().getId()).getOwner();
  const ownerEmail=owner?String(owner.getEmail()||'').toLowerCase():'';
  if(caller!==ownerEmail) throw Error('Chỉ chủ sở hữu Google Sheets được mở trình quản lý tài khoản.');
  return {email:caller,isOwner:caller===ownerEmail};
}
function accountSheetState(){
  const actor=accountActor_();
  const list=rows_(APP.sheets.users).map(r=>({phone:String(r.v[0]||''),role:String(r.v[1]||''),status:String(r.v[3]||'')}));
  return {actor:actor.email,accounts:list};
}
function accountValidPassword_(p){
  if(typeof p!=='string'||p.length<6||p.length>128){
    throw Error('Mật khẩu phải có từ 6 đến 128 ký tự. Không bắt buộc chữ hoa, chữ số hoặc ký tự đặc biệt.');
  }
  return p;
}
function accountSheetSave(payload){
  const actor=accountActor_();
  const phone=normalizePhoneV6_(payload&&payload.phone);
  const role=String(payload&&payload.role||'EDITOR').toUpperCase();
  const password=accountValidPassword_(payload&&payload.password);
  if(!['ADMIN','EDITOR'].includes(role))throw Error('Vai trò không hợp lệ.');
  const props=PropertiesService.getScriptProperties();
  if(!props.getProperty('AUTH_PEPPER')){
    if(rows_(APP.sheets.users).length)throw Error('Thiếu AUTH_PEPPER. Không thay đổi khóa bí mật khi đã có tài khoản.');
    props.setProperty('AUTH_PEPPER',secretV6_());
  }
  const prior=userV6_(phone);
  if(!prior && rows_(APP.sheets.users).length===0 && !actor.isOwner)throw Error('Chỉ chủ file mới tạo được tài khoản đầu tiên.');
  if(prior&&prior.role==='ADMIN'&&role!=='ADMIN'){
    const activeAdmins=rows_(APP.sheets.users).filter(r=>String(r.v[1]).toUpperCase()==='ADMIN'&&String(r.v[3]).toUpperCase()==='ACTIVE');
    if(activeAdmins.length<=1)throw Error('Phải giữ ít nhất một Admin hoạt động.');
  }
  const lock=LockService.getScriptLock();lock.waitLock(30000);
  try{
    const latest=userV6_(phone);
    const sheet=sh_(APP.sheets.users);
    sheet.getRange('A:A').setNumberFormat('@'); // ALWAYS store phone as plain text, preserve leading zero
    const record=[phone,role,hashV6_(phone+'|'+password),'ACTIVE',latest?sheet.getRange(latest.row,5).getValue()||new Date():new Date()];
    if(latest){sheet.getRange(latest.row,1,1,record.length).setValues([record]);sheet.getRange(latest.row,1).setNumberFormat('@');}
    else { const next=Math.max(2,sheet.getLastRow()+1);sheet.getRange(next,1).setNumberFormat('@');sheet.getRange(next,1,1,record.length).setValues([record]); } 
    auditV6_(actor.email,latest?'SHEET_RESET_PASSWORD':'SHEET_CREATE_ACCOUNT','',phone+' '+role);
    return {ok:true,phone,role};
  }finally{lock.releaseLock();}
}
function accountSheetSetStatus(phone,status){
  const actor=accountActor_();
  phone=normalizePhoneV6_(phone); status=String(status||'').toUpperCase();
  if(!['ACTIVE','DISABLED'].includes(status))throw Error('Trạng thái không hợp lệ.');
  const u=userV6_(phone);if(!u)throw Error('Không tìm thấy tài khoản.');
  if(status==='DISABLED'&&u.role==='ADMIN'){}
  if(u.role==='ADMIN'&&status==='DISABLED'){
    const activeAdmins=rows_(APP.sheets.users).filter(r=>String(r.v[1]).toUpperCase()==='ADMIN'&&String(r.v[3]).toUpperCase()==='ACTIVE');
    if(activeAdmins.length<=1)throw Error('Không thể khóa Admin cuối cùng.');
  }
  sh_(APP.sheets.users).getRange(u.row,4).setValue(status);
  auditV6_(actor.email,'SHEET_SET_STATUS','',phone+' '+status);
  return {ok:true};
}
