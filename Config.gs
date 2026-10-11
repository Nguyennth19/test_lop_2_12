/** V6 - one Google Spreadsheet, independent functional tabs. */
const APP = Object.freeze({
  version:'6.0', currency:'VND',
  driveFolderId:'1301rXcg99z9Ji1Pc64DZZRsJqfFnFrHM',
  maxFileBytes:5*1024*1024, sessionSeconds:10800,
  sheets:{students:'DANH SÁCH HỌC SINH',legacy:'SỔ THU CHI',cash:'THU_TIEN',goods:'THU_HIEN_VAT',expenses:'CHI_TIEU',contacts:'LIEN_HE',users:'USERS',audit:'AUDIT_LOG',loginAudit:'LOGIN_AUDIT',settings:'SETTINGS'},
  headers:{
    THU_TIEN:['Mã thu','Ngày thu','Mã học sinh/người góp','Người đóng (nội bộ)','Đợt thu','Số tiền (VND)','Hình thức','Ghi chú nội bộ','Ảnh chứng từ ID','Trạng thái','Người tạo','Ngày tạo','Ngày cập nhật'],
    THU_HIEN_VAT:['Mã hiện vật','Ngày nhận','Mã học sinh','Tên hiển thị người đóng góp','Tên sản phẩm','Số lượng','Đơn vị','Giá trị ước tính (VND)','Ảnh sản phẩm ID','Ghi chú công khai','Trạng thái công khai','Trạng thái','Người tạo','Ngày tạo','Ngày cập nhật'],
    CHI_TIEU:['Mã chi','Ngày chi','Nhóm khoản chi','Nội dung chi','Số tiền (VND)','Hình thức','Chứng từ ID','Ảnh sản phẩm ID','Ghi chú công khai','Trạng thái','Người tạo','Ngày tạo','Ngày cập nhật'],
    LIEN_HE:['Mã học sinh','Tên học sinh (nội bộ)','Tên phụ huynh (nội bộ)','Số điện thoại (nội bộ)','Vai trò liên hệ','Tên hiển thị công khai','Thông tin liên hệ công khai','Trạng thái công khai','Ghi chú nội bộ','Ngày cập nhật'],
    LOGIN_AUDIT:['Thời gian','Email','Sự kiện','Thành công','Chi tiết']
  },
  categories:['Trang bị học tập','Hỗ trợ lớp học','Trang trí sự kiện','Khác']
});
