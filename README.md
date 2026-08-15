# 🚀 ManagerTeams - Hệ Thống Quản Lý Thành Viên, Tổ (Nhóm Nhỏ) & Thăm Viếng

Hệ thống quản lý công việc Ban Điều Hành & Chăm sóc Thăm Viếng Hội Thánh, chạy trên **Google Apps Script Web App** kết hợp **Google Sheets làm CSDL** và tương thích 100% với **Google AppSheet**.

---

## 📌 1. Cấu Trúc Bảng CSDL (Google Sheets Database)

Hệ thống tự động khởi tạo 2 bảng dữ liệu chính trong Google Sheet khi chạy hàm `setupManagerTeamsDatabase()`:

### 1.1 Bảng `Members` (Quản lý Thành Viên)
| Cột (Header) | Mô tả | Kiểu dữ liệu / Ví dụ |
|---|---|---|
| `Member_ID` | Mã thành viên duy nhất | `MEM_1001` |
| `Full_Name` | Họ và tên đầy đủ | `Nguyễn Thị Hồng` |
| `Date_Of_Birth` | Ngày tháng năm sinh | `1995-08-15` (Tự động tính tuổi & báo sinh nhật tháng) |
| `Phone` | Số điện thoại liên hệ | `0988111222` |
| `Address` | Địa chỉ nhà | `123 Đường ABC, Q.1` |
| `Group_ID` | Mã Tổ / Nhóm nhỏ | `GRP_01` |
| `Group_Name` | Tên Tổ / Nhóm nhỏ | `Tổ 1 - Yêu Thương` |
| `Join_Date` | Ngày bắt đầu sinh hoạt | `2023-01-15` |
| `Last_Attendance_Date` | Ngày đi nhóm gần nhất | `2026-07-15` |
| `Status` | Trạng thái sinh hoạt | `Đang sinh hoạt` / `Cần thăm viếng` / `Tạm ngưng` |
| `Notes` | Ghi chú nhu cầu thăm viếng | `Vắng 3-4 tuần cần gọi thăm hỏi` |
| `Created_At` | Ngày tạo bản ghi | `2026-08-11` |
| `Updated_At` | Ngày cập nhật mới nhất | `2026-08-11` |

### 1.2 Bảng `Groups` (Quản lý Tổ / Nhóm Nhỏ)
| Cột (Header) | Mô tả | Kiểu dữ liệu / Ví dụ |
|---|---|---|
| `Group_ID` | Mã Tổ | `GRP_01` |
| `Group_Name` | Tên Tổ / Nhóm | `Tổ 1 - Yêu Thương` |
| `Leader_Name` | Họ tên Tổ trưởng | `Nguyễn Văn A` |
| `Leader_Phone` | SĐT Tổ trưởng | `0901234567` |
| `Notes` | Ghi chú sinh hoạt | `Sinh hoạt tối Thứ Bảy hàng tuần` |
| `Created_At` | Ngày tạo Tổ | `2026-08-11` |

---

## 🌟 2. Các Tính Năng Nổi Bật Dành Cho Vợ Của Bạn

1. 📊 **Tổng Quan & Cảnh Báo Thăm Viếng**:
   - **Tự động cảnh báo vắng (3-4 tuần)**: Phát hiện các thành viên có `Last_Attendance_Date` cách đây từ 21 - 28 ngày trở lên để đưa vào danh sách **Cần thăm viếng**.
   - **Cảnh báo Sinh nhật**: Tự động lọc danh sách thành viên có sinh nhật trong tháng hiện tại kèm số tuổi để tiện chúc mừng.
   - **Điểm danh 1-click**: Bấm nút "Đã thăm / Đi nhóm lại" để tự động cập nhật ngày đi nhóm về hôm nay.
2. 👥 **Quản Lý Thành Viên Toàn Diện**:
   - Lọc tìm kiếm tức thì theo Tên, Số điện thoại, Địa chỉ, Tổ hoặc Trạng thái.
   - Điểm danh hàng loạt (tích chọn danh sách và bấm "⚡ Điểm Danh Nhanh").
3. 🏡 **Quản Lý Tổ (Nhóm Nhỏ)**:
   - Tạo, sửa, xóa Tổ linh hoạt.
   - Tự động thống kê số lượng thành viên trong từng Tổ.
4. 🔒 **Bảo Mật & Tối Ưu Hiệu Năng (Theo [Luuy.md](file:///d:/HAINGUYEN/HOI%20THANH/HT%20TML/ManagerTeams/Luuy.md))**:
   - Bọc `LockService` chống đúp ghi dữ liệu.
   - Đọc/ghi dữ liệu theo mảng 2 chiều batching siêu nhanh.
   - Bọc `google.script.run` trong Promise `runGAS` giao diện phản hồi mượt mượt.

---

## 📱 3. Hướng Dẫn Triển Khai Web App & Kết Nối AppSheet

### 3.1 Phát Hành Web App (Apps Script)
1. Mở dự án trên Apps Script Editor.
2. Nhấn **Deploy** (Triển khai) $\rightarrow$ **New deployment** (Bản triển khai mới).
3. Select type: **Web app**.
   - Description: `v1.0 - Quản lý thành viên & Thăm viếng`.
   - Execute as: **Me**.
   - Who has access: **Anyone** (Hoặc Anyone with Google account).
4. Nhấn **Deploy** và sao chép URL Web App cấp cho chị nhà sử dụng trên điện thoại/máy tính.

### 3.2 Kết Nối Google AppSheet (Tùy chọn)
1. Truy cập: [https://www.appsheet.com](https://www.appsheet.com)
2. Chọn **Start with existing data** $\rightarrow$ Đặt tên App: `ManagerTeams`.
3. Chọn nguồn dữ liệu Google Sheet chứa 2 bảng `Members` và `Groups`. AppSheet sẽ tự động dựng App di động cực đẹp!
