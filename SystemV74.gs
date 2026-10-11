/** V7.4.1. Non-destructive integrity checks; no transaction ID is rewritten. */
function checkIntegrityV74(token){
 authV6_(token,['ADMIN']);
 const mapping=[
  [APP.sheets.cash,9],
  [APP.sheets.expenses,9],
  [APP.sheets.goods,11]
 ];
 const studentIds=roster_().map(x=>x.id),studentSet=new Set(studentIds);
 const problems=[];
 if(studentSet.size!==studentIds.length)problems.push({table:APP.sheets.students,type:'DUPLICATE_STUDENT_ID'});
 for(const [table,statusIdx] of mapping){
  const all=rows_(table),seen=new Set();
  for(const r of all){
   const id=String(r.v[0]||'').trim();
   if(!id)problems.push({table,row:r.row,type:'MISSING_ID'});
   if(seen.has(id))problems.push({table,row:r.row,id,type:'DUPLICATE_ID'});
   seen.add(id);
   const status=String(r.v[statusIdx]||'ACTIVE').toUpperCase();
   if(!['ACTIVE','DELETED'].includes(status))problems.push({table,row:r.row,id,type:'UNKNOWN_STATUS'});
   if(table!==APP.sheets.expenses&&!studentSet.has(String(r.v[2]||'')))
    problems.push({table,row:r.row,id,type:'UNKNOWN_STUDENT'});
  }
 }
 return {ok:problems.length===0,checkedAt:new Date().toISOString(),problems};
}
/** Explicit opt-in: adds configuration keys only when absent. */
function migrateSettingsV74(){
 const sh=sh_(APP.sheets.settings),last=sh.getLastRow();
 const current=last<2?[]:sh.getRange(2,1,last-1,2).getDisplayValues();
 const present=new Set(current.map(r=>String(r[0]||'').trim()));
 const values=[
  ['APP_SOURCE_VERSION','V7.4.1','Mã nguồn GitHub, không phải số bản Deploy'],
  ['FINANCE_COLOR_GREEN','#39A675','Màu xanh lá'],
  ['FINANCE_COLOR_ORANGE','#E58A49','Màu cam'],
  ['FINANCE_COLOR_BLUE','#5879D9','Màu xanh dương'],
  ['FINANCE_COLOR_PURPLE','#986BCF','Màu tím'],
  ['SOFT_DELETE_STATUS','DELETED','Chỉ cập nhật trạng thái, không xóa hàng'],
  ['AUDIT_SCHEMA_VERSION','V7.4.1','Nhật ký trước/sau'],
  ['AUDIT_REASON_REQUIRED','TRUE','Yêu cầu lý do khi sửa/xóa']
 ].filter(r=>!present.has(r[0]));
 if(values.length)sh.getRange(last+1,1,values.length,3).setValues(values);
 return {ok:true,added:values.map(v=>v[0])};
}
