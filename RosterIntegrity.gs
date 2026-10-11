/** V7.3 Bước 1: Audit KHÔNG sửa dữ liệu. Chạy auditRosterV73 trong Apps Script Editor. */
function auditRosterV73(){
 const roster=roster_(), contacts=sh_(APP.sheets.contacts);
 const all=contacts.getLastRow()>1?contacts.getRange(2,1,contacts.getLastRow()-1,11).getDisplayValues():[];
 const problems=[],warnings=[];const studentIds=new Map(),contactIds=new Map();
 roster.forEach(s=>{
  if(!s.id||!s.child){problems.push('Dòng '+s.sheetRow+': thiếu mã/tên học sinh');return;}
  if(studentIds.has(s.id))problems.push('Trùng mã học sinh '+s.id+' tại dòng '+studentIds.get(s.id)+' và '+s.sheetRow);
  studentIds.set(s.id,s.sheetRow);
  if(s.status==='UNKNOWN')problems.push('Trạng thái không nhận diện tại dòng '+s.sheetRow+': '+s.statusRaw);
  if(!s.parent)warnings.push('Chưa có tên phụ huynh: '+s.child+' (mã '+s.id+')');
 });
 all.forEach((r,i)=>{const id=String(r[0]||'').trim(),name=String(r[1]||'').trim();if(!id&&!name)return;
  if(!id){problems.push('LIEN_HE dòng '+(i+2)+': thiếu mã liên hệ');return;}
  if(contactIds.has(id))problems.push('Trùng mã liên hệ '+id+' tại '+contactIds.get(id)+' và '+(i+2));
  contactIds.set(id,i+2);
  const matching=roster.find(x=>x.id===id);
  if(matching&&matching.child!==name)problems.push('Lệch tên mã '+id+': DS='+matching.child+', LH='+name);
  if(!matching&&!/^(?:cô|co)\s/i.test(name))warnings.push('Liên hệ ngoài danh sách học sinh: '+id+' / '+name);
 });
 roster.forEach(s=>{if(s.id&&!contactIds.has(s.id))warnings.push('Thiếu liên hệ học sinh: '+s.id+' / '+s.child);});
 const report={ok:problems.length===0,students:roster.length,active:rosterActive_().length,contacts:all.filter(r=>r[0]||r[1]).length,errors:problems,warnings};
 Logger.log(JSON.stringify(report,null,2));return report;
}
/** Kiểm tra mã trước khi thêm học sinh mới. Không đổi mã hay thứ tự dòng. */
function assertNewStudentIdV73_(id){
 const key=String(id||'').trim();if(!/^\d+$/.test(key))throw Error('Mã mới phải là số nguyên dương.');
 const used=new Set(roster_().map(x=>x.id));
 const cs=sh_(APP.sheets.contacts),last=cs.getLastRow();
 if(last>1)cs.getRange(2,1,last-1,1).getDisplayValues().forEach(r=>used.add(String(r[0]).trim()));
 if(used.has(key))throw Error('Mã '+key+' đã được sử dụng; không cấp lại mã cũ.');
 return {ok:true,id:key};
}
