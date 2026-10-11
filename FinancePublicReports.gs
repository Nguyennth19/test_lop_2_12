/** V7.3 – Bước 5.3: Báo cáo PDF công khai, yêu cầu ADMIN hoặc EDITOR.
 * Dữ liệu lấy trực tiếp từ Sheets phía server, không xuất trường nội bộ.
 * Chỉ xuất các bản ghi ACTIVE; hiện vật phải đồng thời được duyệt PUBLIC.
 */
function financePdfEscapeV753_(v){return String(v==null?'':v).replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;').replace(/'/g,'&#39;');}
function financePdfMoneyV753_(v){return Math.round(Number(v)||0).toLocaleString('vi-VN')+' đ';}
function financePdfDateV753_(v){
  const s=String(v||''); const m=s.match(/^(\d{4})-(\d{2})-(\d{2})/);
  return m?m[3]+'/'+m[2]+'/'+m[1]:s;
}
function financePdfMakeV753_(title,description,statistics,body,note,slug){
  const e=financePdfEscapeV753_,tz=Session.getScriptTimeZone()||'Asia/Ho_Chi_Minh';
  const now=Utilities.formatDate(new Date(),tz,'dd/MM/yyyy HH:mm');
  const html='<!doctype html><html><head><meta charset="utf-8"><style>'+ 
  '@page{size:A4;margin:16mm 12mm}body{font-family:Arial,sans-serif;color:#26344d;font-size:10px;line-height:1.5}h1{font-size:19px;text-align:center;color:#234871;margin:0}h2{font-size:13px;color:#234871;margin:22px 0 8px}.subtitle{text-align:center;color:#596c83}.stat{width:100%;border-collapse:collapse;margin:16px 0}.stat td{border:1px solid #d5dfed;background:#f0f5fb;padding:10px;width:50%}.stat b{font-size:12px}.data{width:100%;border-collapse:collapse}.data th{background:#dfeaf7;color:#244a77;text-align:left}.data td,.data th{border:1px solid #dce4ef;padding:6px 5px;vertical-align:top}.data tr{page-break-inside:avoid}.num{text-align:right;white-space:nowrap}.center{text-align:center}.note{background:#f4f7fb;padding:10px;margin-top:14px}.foot{margin-top:22px;border-top:1px solid #ccd5e0;padding-top:8px;color:#68788a;font-size:9px}' +
  '</style></head><body><h1>'+e(title)+'</h1><div class="subtitle">Quỹ lớp 2/12 · Năm học 2026–2027 · Lập lúc '+e(now)+'</div>'+ 
  '<p class="subtitle">'+e(description)+'</p><table class="stat">'+statistics+'</table>'+body+
  '<div class="note"><b>Hướng dẫn đọc:</b> '+e(note)+'</div>'+ 
  '<div class="foot">Nguồn: Google Sheets Thu chi tại thời điểm xuất. Báo cáo phục vụ việc công khai cho Hội Phụ huynh lớp 2/12. Không bao gồm dữ liệu tài khoản nội bộ hoặc chứng từ riêng tư.</div></body></html>';
  const bytes=HtmlService.createHtmlOutput(html).getAs(MimeType.PDF).getBytes();
  if(bytes.length>8*1024*1024)throw new Error('Báo cáo PDF quá lớn để tải trực tiếp.');
  return {filename:slug+'_'+Utilities.formatDate(new Date(),tz,'yyyyMMdd_HHmm')+'.pdf',base64:Utilities.base64Encode(bytes)};
}
function exportExpensePdfV753(token){
  authV6_(token,['ADMIN','EDITOR']);
  const e=financePdfEscapeV753_,money=financePdfMoneyV753_,date=financePdfDateV753_;
  const records=rows_(APP.sheets.expenses).filter(r=>active_(r,9)).map(r=>({
    id:String(r.v[0]||''),day:dateV6_(r.v[1]),category:String(r.v[2]||''),
    description:String(r.v[3]||''),amount:Number(r.v[4]||0)
  })).filter(r=>r.id&&Number.isFinite(r.amount)&&r.amount>=0).sort((a,b)=>a.day.localeCompare(b.day)||a.id.localeCompare(b.id));
  const categories=(APP.categories||['Trang bị học tập','Hỗ trợ lớp học','Trang trí sự kiện','Khác']);
  const sums={};categories.forEach(x=>sums[x]={count:0,total:0});
  records.forEach(r=>{if(!sums[r.category])sums[r.category]={count:0,total:0};sums[r.category].count++;sums[r.category].total+=r.amount;});
  const grand=records.reduce((a,r)=>a+r.amount,0);
  const groups=Object.keys(sums);const active=groups.filter(x=>sums[x].count>0).length;
  const stats='<tr><td>Tổng số tiền chi<br><b>'+money(grand)+'</b></td><td>Số khoản chi<br><b>'+records.length+'</b></td></tr>'+ 
  '<tr><td>Số nhóm có phát sinh<br><b>'+active+' / '+groups.length+'</b></td><td>Phạm vi<br><b>Toàn bộ giao dịch còn hiệu lực</b></td></tr>';
  const summary='<h2>1. Tổng hợp theo nhóm khoản chi</h2><table class="data"><thead><tr><th>Nhóm khoản chi</th><th>Số khoản</th><th class="num">Số tiền</th><th class="num">Tỷ trọng</th></tr></thead><tbody>'+groups.map(g=>'<tr><td>'+e(g)+'</td><td class="center">'+sums[g].count+'</td><td class="num">'+money(sums[g].total)+'</td><td class="num">'+(grand?((100*sums[g].total/grand).toFixed(1)+'%'):'0%')+'</td></tr>').join('')+'</tbody></table>';
  const details='<h2>2. Danh sách các khoản chi</h2><table class="data"><thead><tr><th>STT</th><th>Ngày chi</th><th>Nhóm</th><th>Nội dung chi</th><th class="num">Số tiền</th></tr></thead><tbody>'+ (records.length?records.map((r,i)=>'<tr><td class="center">'+(i+1)+'</td><td>'+e(date(r.day))+'</td><td>'+e(r.category)+'</td><td>'+e(r.description)+'</td><td class="num">'+money(r.amount)+'</td></tr>').join(''):'<tr><td colspan="5">Chưa có khoản chi được ghi nhận.</td></tr>')+'</tbody><tfoot><tr><th colspan="4">TỔNG CỘNG</th><th class="num">'+money(grand)+'</th></tr></tfoot></table>';
  return financePdfMakeV753_('BÁO CÁO CÁC KHOẢN CHI','Báo cáo chi quỹ lớp – chi tiết và phân bổ theo nhóm',stats,summary+details,'Báo cáo chỉ tính các khoản chi chưa bị đánh dấu xóa. Tỷ trọng = tiền của nhóm / tổng tiền chi. Số liệu là khoản chi, không phải số dư quỹ. Không xuất ID chứng từ, người tạo hoặc dữ liệu quản trị.','Bao_cao_khoan_chi_lop_2_12');
}
function exportGoodsPdfV753(token){
  authV6_(token,['ADMIN','EDITOR']);
  const e=financePdfEscapeV753_,money=financePdfMoneyV753_,date=financePdfDateV753_;
  const records=rows_(APP.sheets.goods).filter(r=>active_(r,11)&&String(r.v[10]||'').trim().toUpperCase()==='PUBLIC').map(r=>({
    day:dateV6_(r.v[1]),donor:String(r.v[3]||''),product:String(r.v[4]||''),quantity:Number(r.v[5]||0),unit:String(r.v[6]||''),estimated:Number(r.v[7]||0),note:String(r.v[9]||'')
  })).filter(r=>r.product&&Number.isFinite(r.quantity)&&r.quantity>0).sort((a,b)=>a.day.localeCompare(b.day)||a.product.localeCompare(b.product,'vi'));
  const estimated=records.reduce((a,r)=>a+(Number.isFinite(r.estimated)&&r.estimated>0?r.estimated:0),0);
  const products=new Set(records.map(x=>x.product.toLocaleLowerCase('vi'))).size;
  const stats='<tr><td>Số lượt hỗ trợ hiện vật<br><b>'+records.length+'</b></td><td>Số loại sản phẩm<br><b>'+products+'</b></td></tr>'+ 
  '<tr><td>Giá trị ước tính được khai báo<br><b>'+money(estimated)+'</b></td><td>Phạm vi<br><b>Chỉ hiện vật đã duyệt công khai</b></td></tr>';
  const rows=records.length?records.map((r,i)=>'<tr><td class="center">'+(i+1)+'</td><td>'+e(date(r.day))+'</td><td>'+e(r.donor||'Không công bố')+'</td><td>'+e(r.product)+'</td><td class="num">'+e(r.quantity)+'</td><td>'+e(r.unit)+'</td><td class="num">'+(r.estimated>0?money(r.estimated):'—')+'</td><td>'+e(r.note)+'</td></tr>').join(''):'<tr><td colspan="8">Chưa có khoản hỗ trợ hiện vật được duyệt công khai trong kỳ.</td></tr>';
  const table='<h2>1. Chi tiết các khoản hỗ trợ hiện vật</h2><table class="data"><thead><tr><th>STT</th><th>Ngày nhận</th><th>Người hỗ trợ công khai</th><th>Hiện vật</th><th>Số lượng</th><th>Đơn vị</th><th>Ước tính</th><th>Ghi chú công khai</th></tr></thead><tbody>'+rows+'</tbody></table>';
  return financePdfMakeV753_('BÁO CÁO HỖ TRỢ HIỆN VẬT','Danh sách đóng góp sản phẩm và vật dụng cho lớp',stats,table,'Mỗi dòng thể hiện một lần tiếp nhận. Số lượng của các đơn vị khác nhau không cộng chung. Giá trị ước tính nếu có chỉ để tham khảo, KHÔNG cộng vào tiền mặt hoặc số dư quỹ. Chỉ công khai tên người hỗ trợ và ghi chú được phép hiển thị.','Bao_cao_hien_vat_lop_2_12');
}
