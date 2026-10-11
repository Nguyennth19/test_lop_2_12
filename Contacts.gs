/** V6.12 Contact Directory. Sheet LIEN_HE columns A:K:
 * A ID, B student/name, C gender, D private parent, E private phone,
 * F role, G public label, H Drive image ID, I PUBLIC/PRIVATE, J internal note, K updated date.
 * Never send phone fields or unpublished records to the public list.
 */
const CONTACT_PICTURES_V612 = Object.freeze({
 teacher_ngoc:'1Aj32ecqyO2r6AeE8epFwh8IjoUNkdnqu',
 teacher_nhuy:'13LrvyrPc1EPaGJ23tyOGnjOByQawe-xv',
 head:'1VHhhFgsXGpysialXkMo3_lh3iUoEBTsw',
 member_minhphuc:'1bLALCZtyhav-tFIFnHYP9tc6kAyc8bIY',
 deputy_lichkhe:'1ruti2jUDXHCQnqEvtyk64s1SEo5qexdh',
 secretary:'17zeloCbhC-54x5E4TTqvyb_eZZ0cOqin',
 boy:'1gq5MQnHj_yUSyogtn5tDNSf-jDnMcYPG',
 girl:'1nH5jlpfssU7Yt1ljjT0yFEs9NzcbOklg'
});
function contactRowsV612_(){
 const s=sh_(APP.sheets.contacts),last=s.getLastRow();
 return last<2?[]:s.getRange(2,1,last-1,11).getDisplayValues()
  .map((v,i)=>({row:i+2,v})).filter(r=>String(r.v[0]||'').trim());
}
function contactTypeV612_(r){
 const s=String(r.v[5]||'').toLowerCase();
 if(/cô chủ nhiệm|chủ nhiệm|giáo viên|bảo mẫu|bảo mãu/.test(s))return 'teacher';
 if(/trưởng ban|phó ban|thành viên|thư ký|thủ quỹ|thủ quỷ/.test(s))return 'board';
 return 'parent';
}
function contactPictureV612_(r,kind){
 const v=r.v,role=(v[5]||'').toLowerCase(),name=(v[1]+' '+v[6]).toLowerCase(),gender=(v[2]||'').toLowerCase();
 // A configured photo ID (H) is preferred; backend validates image metadata.
 const custom=String(v[7]||'').trim();
 if(/^[a-zA-Z0-9_-]{20,100}$/.test(custom))return custom;
 if(kind==='teacher')return /ngọc/.test(name)||/chủ nhiệm/.test(role)?CONTACT_PICTURES_V612.teacher_ngoc:CONTACT_PICTURES_V612.teacher_nhuy;
 if(kind==='board'){
  if(/trưởng ban/.test(role))return CONTACT_PICTURES_V612.head;
  if(/thư ký/.test(role))return CONTACT_PICTURES_V612.secretary;
  if(/thành viên/.test(role)&&/minh phúc/.test(name))return CONTACT_PICTURES_V612.member_minhphuc;
  if(/phó ban/.test(role))return CONTACT_PICTURES_V612.deputy_lichkhe;
 }
 return /nữ|nu|female/.test(gender)?CONTACT_PICTURES_V612.girl:CONTACT_PICTURES_V612.boy;
}
function contactDisplayV612_(r){
 let label=String(r.v[6]||'').trim();
 const kind=contactTypeV612_(r);
 if(kind==='parent'||kind==='board'){
   label=label.replace(/^phụ\s*huynh\s*bé\s*:?\s*/i,'').trim();
   if(!label)label=String(r.v[1]||'').trim();
 }
 return label;
}
function publicContactsV6(){
 const seen=new Set(),out=[];
 contactRowsV612_().forEach(r=>{
  if(String(r.v[8]||'').trim().toUpperCase()!=='PUBLIC')return;
  const id=String(r.v[0]||'').trim(),kind=contactTypeV612_(r),name=contactDisplayV612_(r);
  if(!id||!name||seen.has(id))return;
  seen.add(id);
  out.push({id,kind,name,role:String(r.v[5]||'').trim(),picture:contactPictureV612_(r,kind),canContact:/^(?:\+?84|0)\d{9}$/.test(String(r.v[4]||'').replace(/[\s.\-()]/g,''))});
 });
 const groupRank=r=>r.kind==='teacher'?0:r.kind==='board'?1:2;
 const boardRank=role=>{
   const s=String(role||'').toLowerCase().normalize('NFD').replace(/[\u0300-\u036f]/g,'').replace(/đ/g,'d');
   if(s.includes('truong ban'))return 0;
   if(s.includes('pho ban'))return 1;
   if(s.includes('thanh vien'))return 2;
   if(s.includes('thu ky'))return 3;
   if(s.includes('thu quy'))return 4;
   return 5;
 };
 return out.sort((a,b)=>groupRank(a)-groupRank(b)||
   (a.kind==='board'?boardRank(a.role)-boardRank(b.role):0));
}
/** Image is requested per approved public contact, never in a giant page payload.
 * We send a PNG/JPEG data URI directly, without CSS hiding or lazy-load on the client.
 */
function contactAvatarV612(id){
 const record=contactRowsV612_().find(r=>String(r.v[0])===String(id)&&String(r.v[8]).toUpperCase()==='PUBLIC');
 if(!record)throw Error('Liên hệ chưa được duyệt công khai.');
 const imageId=contactPictureV612_(record,contactTypeV612_(record));
 const cache=CacheService.getScriptCache(),key='contactPhoto_731_'+imageId;
 const cached=cache.get(key);if(cached)return {src:cached};
 const file=DriveApp.getFileById(imageId),blob=file.getBlob();
 const mime=blob.getContentType();
 if(!/^image\/(png|jpeg|webp)$/i.test(mime))throw Error('Ảnh không ở định dạng PNG/JPG/WEBP.');
 const bytes=blob.getBytes();if(bytes.length>4*1024*1024)throw Error('Ảnh gốc quá lớn, hãy tối ưu về dưới 4 MB.');
 const uri='data:'+mime+';base64,'+Utilities.base64Encode(bytes);
 if(uri.length<95000)cache.put(key,uri,21600);
 return {src:uri};
}
/** Return the Zalo URL only after the user clicks on an approved card.
 * The destination URL will necessarily reveal the phone number to a visitor.
 */
function contactZaloV612(id){
 const row=contactRowsV612_().find(r=>String(r.v[0])===String(id)&&String(r.v[8]).toUpperCase()==='PUBLIC');
 if(!row)throw Error('Liên hệ chưa công khai.');
 let phone=String(row.v[4]||'').replace(/[\s.\-()]/g,'');
 if(phone.startsWith('+84'))phone='0'+phone.slice(3);
 else if(phone.startsWith('84'))phone='0'+phone.slice(2);
 if(!/^0\d{9}$/.test(phone))throw Error('Liên hệ này chưa có số Zalo hợp lệ.');
 return {url:'https://zalo.me/'+phone};
}
function privateContactsV6(token){
 authV6_(token,['ADMIN','EDITOR']);
 return contactRowsV612_().map(r=>({row:r.row,studentId:String(r.v[0]),child:String(r.v[1]),parent:String(r.v[3]),phone:String(r.v[4]),role:String(r.v[5]),publicName:String(r.v[6]),publicContact:'',publish:String(r.v[8]||'PRIVATE'),note:String(r.v[9]||''),gender:String(r.v[2]||''),photoId:String(r.v[7]||'')}));
}
function saveContactV6(p){
 const actor=authV6_(p&&p.token,['ADMIN']),id=required_(p.studentId,100);
 const rec=contactRowsV612_().find(r=>String(r.v[0])===id);if(!rec)throw Error('Không tìm thấy liên hệ.');
 const role=clean_(p.role,100),name=clean_(p.publicName,150),publish=p.publish==='PUBLIC'?'PUBLIC':'PRIVATE';
 if(publish==='PUBLIC'&&!name)throw Error('Cần tên hiển thị trước khi công khai.');
 const s=sh_(APP.sheets.contacts);
 s.getRange(rec.row,7).setValue(name);s.getRange(rec.row,9).setValue(publish);s.getRange(rec.row,6).setValue(role);
 auditV6_(actor.phone||actor.email||'ADMIN','UPDATE_CONTACT',id,'Publish='+publish);
 return {ok:true};
}
/** New students are appended as PRIVATE and no existing data is overwritten. */
function syncContactsV6_OLD_UNUSED(token){
 const actor=authV6_(token,['ADMIN']),present=new Set(contactRowsV612_().map(r=>String(r.v[0]).trim()));
 const missing=roster_().filter(x=>x.id&&!present.has(x.id));
 if(missing.length){const s=sh_(APP.sheets.contacts);s.getRange(s.getLastRow()+1,1,missing.length,11).setValues(missing.map(x=>[x.id,x.child,'',x.parent,x.phone,'','Phụ huynh bé: '+x.child,'','PRIVATE','',new Date()]));}
 auditV6_(actor.phone||actor.email||'ADMIN','SYNC_CONTACTS','',missing.length+' new contacts');return {added:missing.length};
}

/** Admin-triggered safe sync, backed by the same engine used by scheduled trigger. */
function syncContactsV6(token){
 const actor=authV6_(token,['ADMIN']);
 return syncStudentContactsV731_('ADMIN:'+String(actor.phone||actor.email||''));
}
