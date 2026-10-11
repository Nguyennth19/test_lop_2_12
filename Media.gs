function uploadV6_(file,kind){if(!file||!file.base64)return {id:''};const mime=String(file.mime||''),isImage=/^image\/(jpeg|png|gif|webp)$/i.test(mime);if(!isImage&&!(kind==='receipt'&&mime==='application/pdf'))throw Error('Tệp không được hỗ trợ.');const enc=String(file.base64||'');if(!/^[A-Za-z0-9+/]*={0,2}$/.test(enc)||enc.length>Math.ceil(APP.maxFileBytes*4/3)+8)throw Error('Tệp không hợp lệ hoặc quá lớn.');const bytes=Utilities.base64Decode(enc);if(!bytes.length||bytes.length>APP.maxFileBytes)throw Error('Tệp vượt quá 5 MB.');const name=(kind+'_'+String(file.name||'image')).replace(/[^\p{L}\p{N}._-]/gu,'_').slice(0,100);const folderId=kind==='cash'?'1mtiwgc690bRce_ADCQ7eeRdcxRKMGa_K':kind==='goods'?'1E-TVsFP6VsqHWXiudxycYvf6dl3g_U_8':(kind==='receipt'||kind==='expense-product')?'1WeBf1AJr47Z2s7s3GLzHRM3nd6DtB02v':APP.driveFolderId;const folder=DriveApp.getFolderById(folderId);const created=folder.createFile(Utilities.newBlob(bytes,mime,name));return {id:created.getId()};}
function publicMediaV6(type,id,variant){const kind=String(type||''),key=String(id||'');let f='';if(kind==='goods'){const r=find_(APP.sheets.goods,key);if(!r||!active_(r,11)||String(r.v[10])!=='PUBLIC')throw Error('Không có quyền xem hình ảnh.');f=String(r.v[8]||'');}else if(kind==='expense'){const r=find_(APP.sheets.expenses,key);if(!r||!active_(r,9))throw Error('Không tìm thấy khoản chi.');f=String(r.v[variant==='product'?7:6]||'');}else throw Error('Không có quyền xem.');if(!f)return {type:'none'};const file=DriveApp.getFileById(f),mime=file.getMimeType(),blob=file.getBlob();if(blob.getBytes().length>APP.maxFileBytes)throw Error('Tệp quá lớn để xem.');if(!/^image\/(jpeg|png|gif|webp)$/i.test(mime)&&mime!=='application/pdf')throw Error('Loại tệp không hợp lệ.');return {type:mime==='application/pdf'?'pdf':'image',data:'data:'+mime+';base64,'+Utilities.base64Encode(blob.getBytes())};}
/** Strictly private cash-receipt preview. Never callable without an active session. */
function privateCashMediaV6(token,id){authV6_(token,['ADMIN','EDITOR']);const r=find_(APP.sheets.cash,id);if(!r||!active_(r,9))throw Error('Không tìm thấy giao dịch thu.');const f=String(r.v[8]||'');if(!f)return {type:'none'};const file=DriveApp.getFileById(f),mime=file.getMimeType();if(!/^image\/(jpeg|png|gif|webp)$/i.test(mime))throw Error('Ảnh không hợp lệ.');const bytes=file.getBlob().getBytes();if(bytes.length>APP.maxFileBytes)throw Error('Ảnh quá lớn.');return {type:'image',data:'data:'+mime+';base64,'+Utilities.base64Encode(bytes)};}

/** Read-only, public QR for the bank transfer guidance. File is an explicitly designated QR, not an upload. */
function publicTransferQrV68(){
  const id='1tHdkBZZKBfNkpfd4_VZtpHJaU8OO75Fm';
  const f=DriveApp.getFileById(id),mime=f.getMimeType();
  if(mime!=='image/png')throw Error('Ảnh QR phải là PNG.');
  const bytes=f.getBlob().getBytes();
  if(!bytes.length||bytes.length>1024*1024)throw Error('Không thể tải QR chuyển khoản.');
  return {type:'image',data:'data:image/png;base64,'+Utilities.base64Encode(bytes)};
}

/** V6.9: embed QR as safe JSON in the first HTML response; no client RPC required. */
function qrBootstrapJsonV69_(){
  try {
    const result=publicTransferQrV68();
    if (!result || result.type!=='image' || !/^data:image\/png;base64,[A-Za-z0-9+/=]+$/.test(String(result.data||''))) {
      throw new Error('Dữ liệu QR không hợp lệ.');
    }
    return JSON.stringify({ok:true,src:result.data});
  } catch(e) {
    // Avoid exposing detailed Drive errors in public HTML.
    console.error('QR preload failed: '+e);
    return JSON.stringify({ok:false,error:'Không thể đọc ảnh QR từ tài khoản triển khai Apps Script. Vui lòng kiểm tra quyền Drive và triển khai lại.'});
  }
}
