# 🏠 Phần mềm Quản lý Phòng Trọ
Phần mềm desktop chạy trên Windows dùng để quản lý thu tiền phòng trọ, tính toán tiền điện nước và xuất báo cáo.

## 📋 Tình trạng dự án (04/04/2026)
✅ **HOÀN THÀNH**
- [x] Khởi tạo dự án Electron.js
- [x] Cài đặt đầy đủ dependencies
- [x] Cấu trúc thư mục chuẩn project
- [x] Cơ sở dữ liệu SQLite (Rooms, Settings, Billing)
- [x] Tự động tạo 12 phòng mặc định 1A-6A, 1B-6B
- [x] Giao diện chính Dashboard
- [x] Grid danh sách phòng với trạng thái
- [x] Modal chi tiết nhập liệu từng phòng
- [x] Bộ chọn tháng/năm lịch sử
- [x] Menu tab: Dashboard, Cài đặt, Người thuê, Báo cáo
- [x] Ứng dụng chạy thành công trên Electron

⏳ **ĐANG PHÁT TRIỂN**
- [ ] Logic tính toán tiền điện/nước realtime
- [ ] Hệ thống khóa/mở khóa dữ liệu
- [ ] Màn hình Cài đặt bảng giá
- [ ] Quản lý thông tin người thuê
- [ ] Upload ảnh CCCD & lưu vào resources
- [ ] Xuất báo cáo Excel
- [ ] Xuất phiếu thu Word
- [ ] Đóng gói file .exe cài đặt

## 🛠️ Công nghệ sử dụng
| Công nghệ | Mô tả |
|---|---|
| **Electron.js** | Framework tạo ứng dụng desktop từ Web |
| **SQLite3** | Cơ sở dữ liệu file đơn, không cần server |
| **Node.js** | Runtime backend |
| **ExcelJS** | Xuất báo cáo file Excel chuyên nghiệp |
| **docxtemplater** | Xuất phiếu thu file Word |
| **electron-builder** | Đóng gói thành file .exe cài đặt |

## 🚀 Chạy ứng dụng
```bash
# Cài đặt dependencies
npm install

# Chạy ở chế độ development
npm start

# Đóng gói thành file .exe
npm run dist
```

## 📁 Cấu trúc thư mục
```
Phần mềm Phòng trọ/
├── main.js              # Entry point Electron
├── package.json         # Config dự án
├── assets/              # Icon, hình ảnh
├── database/            # File data.db SQLite
├── resources/
│   └── uploads/         # Lưu ảnh CCCD người thuê
└── src/
    ├── main/            # Backend logic
    └── renderer/        # Giao diện frontend
        └── index.html   # Giao diện chính
```

## 📋 Danh sách tính năng
### ✅ Đã có
- Hiển thị danh sách 12 phòng
- Xem chi tiết từng phòng
- Chọn tháng/năm xem lịch sử
- Form nhập chỉ số điện nước
- Trạng thái phòng: Trống / Đang thuê

### ⏳ Sắp tới
- Tự động tính toán số sử dụng & tổng tiền
- Khóa dữ liệu sau khi chốt số liệu
- Cấu hình bảng giá điện, nước, rác, internet
- Quản lý thông tin người thuê & ảnh CCCD
- Xuất Excel báo cáo tổng hợp
- Xuất Phiếu thu tiền Word từng phòng
- Xuất hàng loạt nhiều phòng cùng lúc

## 🔒 Ràng buộc kỹ thuật đã implement
- [x] Database nằm cùng thư mục app, dễ dàng backup
- [x] Đường dẫn động, không dùng đường dẫn cứng
- [x] Tương thích khi đóng gói .exe
- [x] Tự động khởi tạo dữ liệu mặc định