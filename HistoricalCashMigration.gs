/**
 * V7.3 – Bước 4 | Nhập các giao dịch đã được xác nhận từ sheet lịch sử.
 * Chạy previewHistoricalCashV734() trước, sau đó importHistoricalCashV734().
 * Đọc dữ liệu trực tiếp từ Google Sheets. Không nhúng thông tin người đóng vào mã.
 */
const CASH_MIG_V734 = Object.freeze({
  source: 'THU_TIEN_NHAP_CU_2026',
  target: 'THU_TIEN',
  roster: 'DANH SÁCH HỌC SINH',
  summary: 'TONG_HOP_THU_THEO_HOC_SINH',
  prefix: 'PT-CU26-',
  batch: 'Ủng hộ quỹ lớp 2026–2027',
  expectedRows: 32,
  expectedTotal: 15502000
});

function cashMigrationDataV734_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const c = CASH_MIG_V734;
  const source = ss.getSheetByName(c.source);
  const target = ss.getSheetByName(c.target);
  const roster = ss.getSheetByName(c.roster);
  if (!source || !target || !roster) throw Error('Thiếu sheet nguồn, đích hoặc danh sách học sinh.');
  const rosterRows = roster.getRange(4, 1, Math.max(1, roster.getLastRow()-3), 8).getValues();
  const studentMap = new Map();
  rosterRows.forEach(r => {
    const id=String(r[0]||'').trim(), name=String(r[1]||'').trim();
    if (!id || !name) return;
    const key=name.normalize('NFC').toLocaleLowerCase('vi-VN');
    if (studentMap.has(key)) throw Error('Trùng tên học sinh trong danh sách: '+name);
    studentMap.set(key, {id, name});
  });
  const sourceRows=source.getRange(2,1,Math.max(1,source.getLastRow()-1),13).getValues();
  const targetRows=target.getLastRow()>1 ? target.getRange(2,1,target.getLastRow()-1,13).getValues() : [];
  const existingIds = new Set(targetRows.map(r=>String(r[0]||'').trim()).filter(Boolean));
  const migrated = [];
  const problems = [];
  sourceRows.forEach((r,i)=>{
    const sheetRow=i+2;
    if (!r.some(v=>v!=='' && v!==null)) return;
    const id=c.prefix+String(i+1).padStart(3,'0');
    const note=String(r[7]||'');
    const match=note.match(/(?:^|\|)\s*Học sinh:\s*([^|]+)\s*(?:\||$)/i);
    const student=match&&studentMap.get(match[1].trim().normalize('NFC').toLocaleLowerCase('vi-VN'));
    const rawAmount=r[5];
    const amount=typeof rawAmount==='number'?rawAmount:Number(String(rawAmount).replace(/[.,\s]/g,''));
    const rawDate=r[1];
    let date=null;
    if (rawDate instanceof Date && !isNaN(rawDate)) date=rawDate;
    else {
      const m=String(rawDate).match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
      if(m){date=new Date(Number(m[3]),Number(m[2])-1,Number(m[1]),12);if(date.getFullYear()!==Number(m[3])||date.getMonth()!==Number(m[2])-1||date.getDate()!==Number(m[1]))date=null;}
    }
    if(!student) problems.push('Dòng '+sheetRow+': không tìm được mã học sinh từ ghi chú.');
    if(!Number.isSafeInteger(amount)||amount<=0)problems.push('Dòng '+sheetRow+': số tiền không hợp lệ.');
    if(!date)problems.push('Dòng '+sheetRow+': ngày thu không hợp lệ.');
    if(!String(r[3]||'').trim())problems.push('Dòng '+sheetRow+': thiếu người chuyển khoản.');
    const details=note+' | Nguồn: '+c.source+'!'+sheetRow+' | Khóa nhập: OLD2026-'+String(i+1).padStart(3,'0');
    const row=[id,date,student?student.id:'',String(r[3]||''),c.batch,amount,String(r[6]||'Chuyển khoản'),details,String(r[8]||''),'ACTIVE','NHAP_CU_2026',new Date(),new Date()];
    migrated.push({id,amount,student:student?student.name:'',row,already:existingIds.has(id)});
  });
  const total=migrated.reduce((s,x)=>s+x.amount,0);
  if(migrated.length!==c.expectedRows)problems.push('Số giao dịch nguồn khác '+c.expectedRows+': '+migrated.length);
  if(total!==c.expectedTotal)problems.push('Tổng tiền nguồn không khớp '+c.expectedTotal+': '+total);
  const counts=new Map();migrated.forEach(x=>counts.set(x.student,(counts.get(x.student)||0)+1));
  if(counts.size!==31)problems.push('Số học sinh trong nguồn khác 31: '+counts.size);
  // Nếu đã có giao dịch nhập thủ công với cùng khóa lịch sử nhưng mã thu khác, dừng để soát xét.
  targetRows.forEach((r,i)=>{
    const m=String(r[7]||'').match(/OLD2026-(\d{3})/);
    if(m){const expected=c.prefix+m[1];if(String(r[0]||'')!==expected)problems.push('Trùng khóa lịch sử ở THU_TIEN dòng '+(i+2)+': '+m[0]);}
  });
  // Nếu một giao dịch đã được nhập, dữ liệu chính phải khớp (tránh ghi đè thay đổi sau này).
  migrated.filter(x=>x.already).forEach(x=>{
    const old=targetRows.find(r=>String(r[0])===x.id);
    if(old && (String(old[2])!==String(x.row[2]) || Number(old[5])!==x.amount)) problems.push('Khác biệt giao dịch đã nhập: '+x.id);
  });
  return {ss,source,target,roster,migrated,total,problems,already:migrated.filter(x=>x.already).length,studentCount:counts.size};
}

function previewHistoricalCashV734(){
 const d=cashMigrationDataV734_();
 const summary={ok:d.problems.length===0,sourceTransactions:d.migrated.length,students:d.studentCount,total:d.total,alreadyImported:d.already,toImport:d.migrated.length-d.already,problems:d.problems};
 console.log(JSON.stringify(summary,null,2));return summary;
}

function importHistoricalCashV734(){
 const lock=LockService.getDocumentLock()||LockService.getScriptLock();
 if(!lock.tryLock(30000))throw Error('Đang có phiên nhập khác; vui lòng thử lại.');
 try{
  const d=cashMigrationDataV734_();
  if(d.problems.length)throw Error('Dừng để bảo toàn dữ liệu: '+d.problems.join('; '));
  const fresh=d.migrated.filter(x=>!x.already).map(x=>x.row);
  if(fresh.length){
    const start=d.target.getLastRow()+1;
    d.target.getRange(start,1,fresh.length,13).setValues(fresh);
    d.target.getRange(start,2,fresh.length,1).setNumberFormat('dd/MM/yyyy');
    d.target.getRange(start,6,fresh.length,1).setNumberFormat('#,##0');
  }
  buildCashStudentSummaryV734();
  const check=cashMigrationDataV734_();
  if(check.problems.length || check.already!==32)throw Error('Cần đối chiếu lại sau nhập: '+JSON.stringify({problems:check.problems,imported:check.already}));
  const outcome={ok:true,newlyImported:fresh.length,totalHistoryRows:check.already,historyTotal:check.total,studentCount:check.studentCount};
  console.log(JSON.stringify(outcome));return outcome;
 }finally{lock.releaseLock();}
}

function buildCashStudentSummaryV734(){
 const d=CASH_MIG_V734, ss=SpreadsheetApp.getActiveSpreadsheet();
 const sheet=ss.getSheetByName(d.summary)||ss.insertSheet(d.summary);
 const headings=['Mã học sinh','Họ tên học sinh','Số lần hỗ trợ','Tổng tiền hỗ trợ (VND)','Ngày hỗ trợ gần nhất','Tình trạng ghi nhận'];
 sheet.getRange(1,1,1,6).setValues([headings]);
 sheet.getRange(1,1,1,6).setFontWeight('bold').setBackground('#2f5597').setFontColor('#ffffff');
 sheet.setFrozenRows(1);
 // Dữ liệu chính đọc theo mã học sinh. Chỉ tính giao dịch còn hiệu lực.
 const roster=ss.getSheetByName(d.roster);
 const arr=roster.getRange(4,1,Math.max(1,roster.getLastRow()-3),2).getValues().filter(r=>r[0]&&r[1]);
 if(sheet.getLastRow()>1)sheet.getRange(2,1,sheet.getLastRow()-1,6).clearContent();
 const rows=arr.map((r,i)=>{
  const row=i+2,id=r[0],name=r[1];
  return [id,name,
   '=COUNTIFS(THU_TIEN!$C$2:$C,$A'+row+',THU_TIEN!$J$2:$J,"<>DELETED",THU_TIEN!$A$2:$A,"<>")',
   '=SUMIFS(THU_TIEN!$F$2:$F,THU_TIEN!$C$2:$C,$A'+row+',THU_TIEN!$J$2:$J,"<>DELETED",THU_TIEN!$A$2:$A,"<>")',
   '=IF(C'+row+'=0,"",MAXIFS(THU_TIEN!$B$2:$B,THU_TIEN!$C$2:$C,$A'+row+',THU_TIEN!$J$2:$J,"<>DELETED"))',
   '=IF(C'+row+'>0,"Đã ghi nhận hỗ trợ","Chưa ghi nhận hỗ trợ")'];
 });
 if(rows.length){sheet.getRange(2,1,rows.length,6).setValues(rows);sheet.getRange(2,4,rows.length,1).setNumberFormat('#,##0');sheet.getRange(2,5,rows.length,1).setNumberFormat('dd/MM/yyyy');}
 sheet.setColumnWidths(1,1,100);sheet.setColumnWidths(2,1,250);sheet.setColumnWidths(3,1,140);sheet.setColumnWidths(4,1,190);sheet.setColumnWidths(5,1,180);sheet.setColumnWidths(6,1,195);
 console.log('Đã tạo bảng tổng hợp '+arr.length+' học sinh. Khi thêm học sinh mới, chạy lại buildCashStudentSummaryV734().');
 return {ok:true,students:arr.length};
}
