# NEXUS_GRAPH.md - Memory Knowledge Graph & Architecture Map

Tài liệu bản đồ trí nhớ dự án (AI Nexus Memory Knowledge Graph) lưu trữ cấu trúc CSDL, sơ đồ thực thể, thành phần giao diện và các luồng dữ liệu của hệ thống **ManagerTeams**.

---

## 🕸️ 1. Entity Relationship Graph (ERD)

```
 [ Members (Thành Viên) ]
     │ (1:N)
     ├───► [ Visitations (Thăm Viếng) ] 
     │ (N:1)
     ├───► [ Groups (Tổ theo Năm) ]
     │ (1:N)
     └───► [ Attendance_Records (Điểm Danh Tuần) ]
               ▲
               │ (N:1)
   [ Weekly_Schedules (Tờ Chương Trình Quý) ] ◄───► [ Quarterly_Themes (Chủ Đề Quý & Năm) ]
```

---

## 📂 2. Data Schema Index (Google Sheet Tabs)

1. **`Members`**: `Mã Thành Viên`, `Họ và Tên`, `Chức Danh`, `Nhiệm Kỳ`, `Ngày Sinh`, `Số Điện Thoại`, `Địa Chỉ`, `Mã Tổ`, `Tên Tổ`, `Ngày Báp Têm`, `Tên Vợ Chồng`, `Ngày Sinh Hoạt`, `Lần Nhóm Cuối`, `Trạng Thái`, `Ghi Chú`
2. **`Groups`**: `Mã Tổ`, `Năm`, `Tên Tổ`, `Tổ Trưởng`, `Số ĐT Tổ Trưởng`, `Ghi Chú`
3. **`Visitations`**: `Mã Thăm Viếng`, `Mã Thành Viên`, `Họ và Tên`, `Người Thăm`, `Ngày Thăm`, `Hình Thức`, `Nhu Cầu Cầu Nguyện`, `Kết Quả & Ghi Chú`
4. **`Quarterly_Themes`**: `Mã Chủ Đề`, `Năm`, `Quý`, `Tên Chủ Đề`, `Câu Gốc`, `Nội Dung Câu Gốc`, `Bài Hát Chủ Đề`, `Ghi Chú`
5. **`Weekly_Schedules`**: `Mã Buổi Nhóm`, `Năm`, `Quý`, `Ngày Nhóm`, `Đề Tài`, `Câu Gốc`, `Đố Kinh Thánh`, `Tổ Phụ Trách`, `Phụ Trách (Diễn Giả)`, `Ghi Chú`
6. **`Attendance_Records`**: `Mã Điểm Danh`, `Ngày Nhóm`, `Mã Thành Viên`, `Họ và Tên`, `Trạng Thái`, `Thuộc Câu Gốc`, `Số Câu Đố KT`, `Ghi Chú`, `Thời Gian Ghi`

---

## ⚡ 3. Feature Registry & Active Component Map

| Tab ID | Feature Component | Client Handler | GAS Backend Endpoint |
|---|---|---|---|
| `#tab-grid-dashboard` | Grid Mini Apps & Quick Access | `openMiniAppGridDashboard()` | `getDashboardStats` |
| `#tab-checkin` | Điểm Danh Đa Tầng & Buổi Nhóm | `loadCheckinDataForDate()` | `getWeeklyAttendanceForDate`, `saveWeeklyAttendanceBatch` |
| `#tab-schedules` | Lịch Quý Tờ Chương Trình | `loadScheduleTab()` | `getWeeklySchedules`, `saveQuarterlySchedulesBatch` |
| `#tab-members` | Quản Lý Thành Viên & Hồ Sơ | `filterMembers()` | `getMembersData`, `saveMember`, `deleteMember` |
| `#tab-visitations` | Nhật Ký Thăm Viếng Đa Tầng | `filterVisitations()` | `getVisitationLogs`, `saveVisitationLog` |
| `#tab-groups` | Quản Lý Tổ Theo Năm | `renderGroupsTable()` | `getGroupsData`, `saveGroup`, `deleteGroup` |
| `#tab-themes` | Chủ Đề Quý & Cả Năm | `filterThemesByYear()` | `getQuarterlyThemes`, `saveTheme`, `deleteTheme` |
| `#tab-accounts` | Phân Quyền & Tài Khoản | `loadAccountsTab()` | `getAccountsData`, `saveAccount`, `deleteAccount` |

---

## 🧠 4. Superpowers Workflow Enforcement Rules

- **Red-Green TDD Verification**: Verification script `scratch/check_syntax.js` executed via AST parser before any deployment.
- **YAGNI & DRY**: Universal helper functions (`getSvgIcon`, `formatDisplayDate`, `escapeHtml`, `runGAS`).
- **Archive Strategy**: 1-click archiving moves older year attendance to `Attendance_Archive_YYYY` to maintain $<100\text{ms}$ query speed.
