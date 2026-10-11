/** V7.3 - Bước 4.1: xuất báo cáo hiện kim an toàn cho việc chia sẻ trong hội phụ huynh.
 * Quyền tạo: ADMIN / EDITOR. Nội dung PDF không chứa dữ liệu người chuyển khoản. */
function exportCashSupportPdfV7341(token){
  authV6_(token,['ADMIN','EDITOR']);
  const roster=roster_();
  const rosterMap={};roster.forEach(s=>rosterMap[String(s.id)]=s);
  const tx=rows_(APP.sheets.cash).filter(r=>active_(r,9)).map(r=>({
    id:String(r.v[0]||''),date:dateV6_(r.v[1]),studentId:String(r.v[2]||''),amount:Number(r.v[5]||0)
  })).filter(x=>x.id&&Number.isFinite(x.amount)&&x.amount>0);
  const sums={};
  tx.forEach(x=>{
    if(!sums[x.studentId])sums[x.studentId]={count:0,total:0,last:''};
    const s=sums[x.studentId];s.count++;s.total+=x.amount;
    if(x.date>s.last)s.last=x.date;
  });
  const money=x=>Math.round(Number(x)||0).toLocaleString('vi-VN')+' đ';
  const esc=x=>String(x===null||x===undefined?'':x).replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;').replace(/'/g,'&#39;');
  const date=x=>{const s=String(x||'');const m=s.match(/^(\d{4})-(\d{2})-(\d{2})/);return m?m[3]+'/'+m[2]+'/'+m[1]:s;};
  const rosterOrdered=roster.slice().sort((a,b)=>String(a.child).localeCompare(String(b.child),'vi'));
  const overall=tx.reduce((s,x)=>s+x.amount,0);
  const supported=rosterOrdered.filter(s=>sums[String(s.id)]&&sums[String(s.id)].count>0).length;
  const orphanIds=[...new Set(tx.map(x=>x.studentId))].filter(id=>!rosterMap[id]);
  const orderedTx=tx.slice().sort((a,b)=>a.date.localeCompare(b.date)||a.id.localeCompare(b.id));
  const now=Utilities.formatDate(new Date(),Session.getScriptTimeZone()||'Asia/Ho_Chi_Minh','dd/MM/yyyy HH:mm');
  const tableRows=rosterOrdered.map((s,i)=>{
    const v=sums[String(s.id)]||{count:0,total:0,last:''};
    return '<tr><td class="center">'+(i+1)+'</td><td>'+esc(s.child)+'</td><td class="center">'+v.count+'</td><td class="num">'+money(v.total)+'</td><td class="center">'+(v.last?esc(date(v.last)):'—')+'</td><td>'+(v.count?'Đã ghi nhận':'Chưa ghi nhận')+'</td></tr>';
  }).join('');
  const transactionRows=orderedTx.map((x,i)=>{
    const student=rosterMap[x.studentId];
    return '<tr><td class="center">'+(i+1)+'</td><td class="center">'+esc(date(x.date))+'</td><td>'+esc(student?student.child:'Mã học sinh '+x.studentId)+'</td><td class="num">'+money(x.amount)+'</td></tr>';
  }).join('');
  const html='<!DOCTYPE html><html><head><meta charset="UTF-8"><style>'+ 
    '@page{size:A4;margin:15mm 13mm}body{font-family:Arial,sans-serif;font-size:10px;color:#24354b}h1{text-align:center;font-size:19px;margin:0 0 5px}h2{font-size:13px;color:#20426a;margin:22px 0 7px}.sub{text-align:center;color:#65768b}.muted{color:#65768b}.stats{width:100%;margin:15px 0;border:1px solid #d8e3f1;border-collapse:collapse}.stats td{padding:9px;border:1px solid #d8e3f1;width:50%}table.data{width:100%;border-collapse:collapse;font-size:9px}table.data th{background:#eaf0f9;color:#213e67}table.data td,table.data th{border:1px solid #dde5f0;padding:5px 5px;vertical-align:middle}.data tr{page-break-inside:avoid}.num{text-align:right;white-space:nowrap}.center{text-align:center}.note{background:#f4f7fb;padding:9px;line-height:1.5}.footer{margin-top:22px;border-top:1px solid #ccc;padding-top:8px;font-size:9px;color:#576477}' +
    '</style></head><body>'+ 
    '<h1>BÁO CÁO HỖ TRỢ QUỸ LỚP 2/12</h1><div class="sub">Năm học 2026–2027 · Báo cáo hiện kim · Lập lúc '+esc(now)+'</div>'+ 
    '<table class="stats"><tr><td><b>Tổng tiền hỗ trợ</b><br>'+money(overall)+'</td><td><b>Số lần hỗ trợ</b><br>'+tx.length+' giao dịch</td></tr>'+ 
    '<tr><td><b>Học sinh đã ghi nhận</b><br>'+supported+' / '+roster.length+' học sinh</td><td><b>Học sinh chưa ghi nhận</b><br>'+(roster.length-supported)+' học sinh</td></tr></table>'+ 
    '<div class="note"><b>Hướng dẫn đọc:</b> Mỗi lần chuyển khoản/ủng hộ được tính là một giao dịch riêng; một học sinh có thể được hỗ trợ nhiều lần. Tổng tiền theo học sinh là tổng các giao dịch hiện còn hiệu lực. '+
    '“Chưa ghi nhận” nghĩa là chưa thấy khoản hỗ trợ trong dữ liệu, không hàm ý nghĩa vụ đóng góp. Báo cáo chỉ công khai họ tên học sinh, ngày và số tiền; không công khai tên người chuyển khoản, số điện thoại, chứng từ hay nội dung ngân hàng.</div>'+ 
    '<h2>1. Tổng hợp hỗ trợ theo học sinh</h2><table class="data"><thead><tr><th>STT</th><th>Họ tên học sinh</th><th>Số lần</th><th>Tổng tiền</th><th>Lần gần nhất</th><th>Ghi nhận</th></tr></thead><tbody>'+tableRows+'</tbody></table>'+ 
    '<h2>2. Nhật ký từng khoản hỗ trợ (công khai)</h2><table class="data"><thead><tr><th>STT</th><th>Ngày nhận</th><th>Họ tên học sinh</th><th>Số tiền</th></tr></thead><tbody>'+transactionRows+'</tbody></table>'+ 
    '<p><b>Đối chiếu:</b> '+tx.length+' giao dịch, tổng cộng '+money(overall)+'. '+(orphanIds.length?('Có '+orphanIds.length+' mã học sinh chưa khớp danh sách; cần quản trị đối chiếu. '):'Tất cả mã học sinh trong các giao dịch đều có hồ sơ trong danh sách hiện tại. ')+'</p>'+ 
    '<div class="footer">Nguồn dữ liệu: THU_TIEN và DANH SÁCH HỌC SINH tại thời điểm xuất báo cáo. Đây là báo cáo hỗ trợ hiện kim, không phải báo cáo số dư quỹ (chưa trừ khoản chi). Báo cáo phục vụ minh bạch nội bộ Hội Phụ huynh lớp 2/12.</div></body></html>';
  const blob=HtmlService.createHtmlOutput(html).getAs(MimeType.PDF).setName('Bao_cao_ho_tro_hien_kim_lop_2_12.pdf');
  const bytes=blob.getBytes();
  if(bytes.length>8*1024*1024)throw new Error('PDF vượt giới hạn dung lượng cho tải trực tiếp.');
  return {filename:'Bao_cao_ho_tro_hien_kim_2_12_'+Utilities.formatDate(new Date(),Session.getScriptTimeZone()||'Asia/Ho_Chi_Minh','yyyyMMdd')+'.pdf',base64:Utilities.base64Encode(bytes)};
}
