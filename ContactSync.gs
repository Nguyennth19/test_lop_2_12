/** V7.3.1 — sync registry A:H to LIEN_HE A:K. No destructive rewrites.
 * Source of truth for child/parent/phone: roster sheet. F,G,H,J remain manual.
 * A unique stable student ID is mandatory. New contacts are PRIVATE.
 */
function syncStudentContactsV731_(reason){
 const lock=LockService.getScriptLock();
 if(!lock.tryLock(25000))throw Error('Đồng bộ đang được thực hiện. Vui lòng thử lại sau.');
 try{
  const source=sh_(APP.sheets.students), target=sh_(APP.sheets.contacts);
  const students=roster_();
  const seen=new Set(),duplicates=[];
  students.forEach(x=>{if(!x.id||!x.child)duplicates.push(x.id||'(trống)');else if(seen.has(x.id))duplicates.push(x.id);seen.add(x.id);});
  if(duplicates.length)throw Error('Trùng hoặc thiếu mã/tên học sinh: '+duplicates.join(', '));
  if(students.some(x=>x.status==='UNKNOWN'))throw Error('Có trạng thái học sinh chưa hợp lệ, cần sửa trước khi đồng bộ.');
  const last=target.getLastRow();
  const existing=last<2?[]:target.getRange(2,1,last-1,11).getValues();
  const index=new Map();
  existing.forEach((r,i)=>{const id=String(r[0]||'').trim();if(!id)return;if(index.has(id))throw Error('Trùng mã liên hệ '+id);index.set(id,{row:i+2,values:r});});
  const updates=[],inserts=[],changes=[];
  let hidden=0,updated=0,added=0;
  students.forEach(s=>{
   const match=index.get(s.id),master=source.getRange(s.sheetRow,1,1,8).getDisplayValues()[0];
   if(!match){
    const label='Phụ huynh bé: '+String(s.child).trim().split(/\s+/).slice(-2).join(' ');
    inserts.push([s.id,s.child,master[3]||'',s.parent||'',s.phone||'','',label,'','PRIVATE','Tự thêm từ danh sách học sinh',new Date()]);
    changes.push(['THÊM',s.id,'Chờ xác nhận công khai']);added++;return;
   }
   const before=match.values, after=before.slice(), status=s.status;
   [s.id,s.child,master[3]||'',s.parent||'',s.phone||''].forEach((v,i)=>{after[i]=v;});
   if(status!=='ACTIVE'&&String(after[8]||'').toUpperCase()==='PUBLIC'){
     after[8]='PRIVATE';hidden++;changes.push(['ẨN',s.id,'Học sinh '+status]);
   }
   // No automatic re-publication on returning students.
   if(after.slice(0,9).some((v,i)=>String(v??'')!==String(before[i]??''))){
    after[10]=new Date();updates.push({row:match.row,after});updated++;
    changes.push(['SỬA',s.id,status]);
   }
  });
  // Existing student IDs that are absent in roster may be removed in error.
  // Fail safe: do not remove/privatize them automatically; warn for human review.
  const masterIds=new Set(students.map(s=>s.id));
  const orphans=[...index.keys()].filter(id=>!masterIds.has(id)&&!isContactTeacherV731_(index.get(id).values));
  if(updates.length){updates.forEach(u=>{
    // Only A:E, I, K updated. Never overwrite F/G/H/J.
    target.getRange(u.row,1,1,5).setValues([u.after.slice(0,5)]);
    target.getRange(u.row,9).setValue(u.after[8]);
    target.getRange(u.row,11).setValue(u.after[10]);
  });}
  if(inserts.length)target.getRange(target.getLastRow()+1,1,inserts.length,11).setValues(inserts);
  const result={ok:true,students:students.length,added,updated,hidden,orphans,changes:changes.slice(0,50)};
  console.log(JSON.stringify({reason:reason||'MANUAL',...result}));
  try{auditV6_('SYSTEM','CONTACT_SYNC_V731',reason||'',JSON.stringify({added,updated,hidden,orphans:orphans.length}));}catch(e){console.warn('Không thể ghi AUDIT_LOG: '+e.message);}
  return result;
 }finally{lock.releaseLock();}
}
function isContactTeacherV731_(r){
 const role=String(r[5]||'').toLowerCase();
 return /giáo viên|chủ nhiệm|bảo mẫu|bảo mãu|^cô/.test(role);
}
/** Manual preflight, safe to invoke in Apps Script editor. */
function previewContactSyncV731(){
 const students=roster_(), contacts=contactRowsV612_();
 const byId=new Map(contacts.map(x=>[String(x.v[0]).trim(),x.v]));
 const newIds=students.filter(s=>!byId.has(s.id)).map(s=>s.id);
 const changed=students.filter(s=>{const c=byId.get(s.id);return c&&[s.child,s.parent,s.phone].some((v,i)=>String(v||'')!==String(c[[1,3,4][i]]||''));}).map(s=>s.id);
 const wouldHide=students.filter(s=>s.status!=='ACTIVE'&&byId.has(s.id)&&String(byId.get(s.id)[8]).toUpperCase()==='PUBLIC').map(s=>s.id);
 const out={students:students.length,contacts:contacts.length,newIds,changed,wouldHide};
 console.log(JSON.stringify(out));return out;
}
/** Trigger is installed ONCE by project owner. Weekly Sunday ~00:00-01:00 Asia/Ho_Chi_Minh.
 * Apps Script time-driven triggers do not guarantee exact 00:00 execution.
 */
function installSundayContactSyncV731(){
 const fn='weeklyContactSyncV731';
 ScriptApp.getProjectTriggers().filter(t=>t.getHandlerFunction()===fn).forEach(ScriptApp.deleteTrigger);
 ScriptApp.newTrigger(fn).timeBased().onWeekDay(ScriptApp.WeekDay.SUNDAY).atHour(0).nearMinute(0).everyWeeks(1).inTimezone('Asia/Ho_Chi_Minh').create();
 console.log('Đã cài lịch Chủ nhật khoảng 00:00, giờ Việt Nam.');
}
function weeklyContactSyncV731(){return syncStudentContactsV731_('SCHEDULED_SUNDAY');}
