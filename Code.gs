/**
 * ManagerTeams - Backend Google Apps Script (Version 3.2 - Complete 55 Members Dataset)
 */

const SHEET_NAMES = {
  MEMBERS: 'Members',
  GROUPS: 'Groups',
  VISITATIONS: 'Visitations',
  QUARTERLY_THEMES: 'Quarterly_Themes',
  WEEKLY_SCHEDULES: 'Weekly_Schedules',
  ATTENDANCE_RECORDS: 'Attendance_Records',
  ACCOUNTS: 'Accounts',
  CONFIG: 'Department_Config'
};

function doGet() {
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('ManagerTeams - Quản Lý Ban Ngành & Thành Viên')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1.0, maximum-scale=1.0, user-scalable=no')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

function getSpreadsheet() {
  const activeId = PropertiesService.getScriptProperties().getProperty('ACTIVE_SPREADSHEET_ID');
  if (activeId && activeId.trim() !== "") {
    try {
      return SpreadsheetApp.openById(activeId.trim());
    } catch (e) {
      console.warn("Failed to open spreadsheet by active ID:", e.message);
    }
  }
  return SpreadsheetApp.getActiveSpreadsheet();
}

function setupNewSpreadsheetDatabase(spreadsheetId, churchName, deptName, motto) {
  return SetupService.initializeNewDatabase(spreadsheetId, churchName, deptName, motto);
}

// BỘ DỮ LIỆU chuẩn 55 BAN VIÊN CHÍNH THỨC CỦA BAN NGÀNH
const INITIAL_55_MEMBERS = [
  ['MEM_001', "Ksor H' Miram", 'Ban viên', '2025-2027', '2000-04-12', '', '', 'GRP_01', 'Tổ 1', '', '', '2026-01-01', '2026-08-10', 'Đang sinh hoạt', ''],
  ['MEM_002', 'Rmah Toàn', 'Ban viên', '2025-2027', '1999-10-15', '', '', 'GRP_01', 'Tổ 1', '', '', '2026-01-01', '2026-08-10', 'Đang sinh hoạt', ''],
  ['MEM_003', 'Trần Thị ThanhThảo', 'Ban viên', '2025-2027', '1995-06-01', '0971520748', '', 'GRP_01', 'Tổ 1', '', '', '2026-01-01', '2026-08-10', 'Đang sinh hoạt', ''],
  ['MEM_004', 'Nguyễn Thế Hải', 'Ban viên', '2025-2027', '1993-03-05', '0934987202', '407/5, Nguyễn Thị Định, Cát Lái, Tp. Thủ Đức', 'GRP_01', 'Tổ 1', '', '', '2026-01-01', '2026-08-10', 'Đang sinh hoạt', ''],
  ['MEM_005', 'Lê Thị Hoàng Yến', 'Phó ban', '2025-2027', '1993-05-28', '0877115836', '407/5, Nguyễn Thị Định, Cát Lái, Tp. Thủ Đức', 'GRP_01', 'Tổ 1', '', '', '2026-01-01', '2026-08-10', 'Đang sinh hoạt', 'Phó ban | Tài khoản: yen'],
  ['MEM_006', 'Nguyễn Quốc Khánh', 'Ban viên', '2025-2027', '1993-01-26', '0773777910', '', 'GRP_01', 'Tổ 1', '', '', '2026-01-01', '2026-08-10', 'Đang sinh hoạt', ''],
  ['MEM_007', 'La Minh Hoàng', 'Ban viên', '2025-2027', '1992-11-29', '0934078190', '20, Đường 38, Bình Trưng Tây, Tp. Thủ Đức', 'GRP_01', 'Tổ 1', '', '', '2026-01-01', '2026-08-10', 'Đang sinh hoạt', ''],
  ['MEM_008', 'Nguyễn Đặng Thảo Nguyên', 'Ban viên', '2025-2027', '1992-04-17', '0935092920', '', 'GRP_01', 'Tổ 1', '', '', '2026-01-01', '2026-08-10', 'Đang sinh hoạt', ''],
  ['MEM_009', 'Phạm Kiều Trâm', 'Ban viên', '2025-2027', '1992-05-18', '0906860489', '', 'GRP_01', 'Tổ 1', '', '', '2026-01-01', '2026-08-10', 'Đang sinh hoạt', ''],
  ['MEM_010', 'Cao Thiên Ngọc', 'Ban viên', '2025-2027', '1992-04-04', '0932166032', '', 'GRP_01', 'Tổ 1', '', '', '2026-01-01', '2026-08-10', 'Đang sinh hoạt', ''],
  ['MEM_011', 'Y Đuen Êban', 'Thư ký', '2025-2027', '1991-10-13', '0903093041', '154, Đường 67, Cát Lái, Tp. Thủ Đức', 'GRP_02', 'Tổ 2', '', '', '2026-01-01', '2026-08-10', 'Đang sinh hoạt', 'Thư ký | Tài khoản: duen'],
  ['MEM_012', 'Nguyễn Ngọc Hà Thi', 'Ban viên', '2025-2027', '1991-05-19', '0902659747', '154, Đường 67, Cát Lái, Tp. Thủ Đức', 'GRP_02', 'Tổ 2', '', '', '2026-01-01', '2026-08-10', 'Đang sinh hoạt', ''],
  ['MEM_013', 'Thiên Nhân', 'Ban viên', '2025-2027', '1991-12-14', '0963577778', '', 'GRP_02', 'Tổ 2', '', '', '2026-01-01', '2026-08-10', 'Đang sinh hoạt', ''],
  ['MEM_014', 'Nhật Thuy', 'Ban viên', '2025-2027', '1991-07-21', '0939926929', '', 'GRP_02', 'Tổ 2', '', '', '2026-01-01', '2026-08-10', 'Đang sinh hoạt', ''],
  ['MEM_015', 'Nguyễn Thị Trà My', 'Ban viên', '2025-2027', '1990-03-05', '0978550426', '42, Thạnh Mỹ Lợi, Thạnh Mỹ Lợi, Tp. Thủ Đức', 'GRP_02', 'Tổ 2', '', '', '2026-01-01', '2026-08-10', 'Đang sinh hoạt', ''],
  ['MEM_016', 'Phạm Anh Tuấn', 'Thủ quỹ', '2025-2027', '1990-09-12', '0399182821', '42, Thạnh Mỹ Lợi, Thạnh Mỹ Lợi, Tp. Thủ Đức', 'GRP_02', 'Tổ 2', '', '', '2026-01-01', '2026-08-10', 'Đang sinh hoạt', 'Tổ trưởng Tổ 2 | Tài khoản: tuan'],
  ['MEM_017', 'Nguyễn NgọcTuệ', 'Ban viên', '2025-2027', '1990-11-07', '0905251525', '', 'GRP_02', 'Tổ 2', '', '', '2026-01-01', '2026-08-10', 'Đang sinh hoạt', ''],
  ['MEM_018', 'Huỳnh Giang Duy Vũ', 'Ban viên', '2025-2027', '1990-02-18', '', '', 'GRP_02', 'Tổ 2', '', '', '2026-01-01', '2026-08-10', 'Đang sinh hoạt', ''],
  ['MEM_019', 'Lê KimHuệ', 'Ban viên', '2025-2027', '1989-05-07', '0978436182', '20, Đường 38, Bình Trưng Tây, Tp. Thủ Đức', 'GRP_02', 'Tổ 2', '', '', '2026-01-01', '2026-08-10', 'Đang sinh hoạt', ''],
  ['MEM_020', 'Nguyễn Thị Nhung', 'Nghị viên', '2025-2027', '1988-02-29', '0928220493', '42, Đường 28, Cát Lái, Tp. Thủ Đức', 'GRP_02', 'Tổ 2', '', '', '2026-01-01', '2026-08-10', 'Đang sinh hoạt', 'Nghị viên | Tài khoản: nhung'],
  ['MEM_021', 'Trịnh Thế Hân', 'Ban viên', '2025-2027', '1988-12-18', '0919499857', '13, Vành đai tây, An Khánh, Tp. Thủ Đức', 'GRP_03', 'Tổ 3', '', '', '2026-01-01', '2026-08-10', 'Đang sinh hoạt', ''],
  ['MEM_022', 'Nguyễn Kim Long', 'Ban viên', '2025-2027', '1988-09-05', '0905673400', '', 'GRP_03', 'Tổ 3', '', '', '2026-01-01', '2026-08-10', 'Đang sinh hoạt', ''],
  ['MEM_023', 'Nguyễn Văn Hưng', 'Ban viên', '2025-2027', '1987-01-03', '0972400688', '194, Hồ Văn Huê, Phường 9, Quận Phú Nhuận', 'GRP_03', 'Tổ 3', '', '', '2026-01-01', '2026-08-10', 'Đang sinh hoạt', ''],
  ['MEM_024', 'Nguyễn Thanh Tuấn', 'Ban viên', '2025-2027', '1987-04-25', '0376258520', '42, Đường 28, Cát Lái, Tp. Thủ Đức', 'GRP_03', 'Tổ 3', '', '', '2026-01-01', '2026-08-10', 'Đang sinh hoạt', ''],
  ['MEM_025', 'Hoàng Thị Giang', 'Ban viên', '2025-2027', '1987-10-06', '0917555085', 'Hommyland 3', 'GRP_03', 'Tổ 3', '', '', '2026-01-01', '2026-08-10', 'Đang sinh hoạt', ''],
  ['MEM_026', 'Nguyễn Thị Nhật Thiên', 'Ban viên', '2025-2027', '1987-11-09', '0902680274', '', 'GRP_03', 'Tổ 3', '', '', '2026-01-01', '2026-08-10', 'Đang sinh hoạt', ''],
  ['MEM_027', 'Trương Thị Thanh Thảo', 'Ban viên', '2025-2027', '1986-04-26', '0901094521', '649/2, Nguyễn Thị Định, Cát Lái, Tp. Thủ Đức', 'GRP_03', 'Tổ 3', '', '', '2026-01-01', '2026-08-10', 'Đang sinh hoạt', ''],
  ['MEM_028', 'Nguyễn Thị Thu Hiền', 'Nghị viên', '2025-2027', '1986-08-10', '0933833967', '603, Nguyễn Thị Định, Cát Lái, Tp. Thủ Đức', 'GRP_03', 'Tổ 3', '', '', '2026-01-01', '2026-08-10', 'Đang sinh hoạt', 'Nghị viên | Tài khoản: hien'],
  ['MEM_029', 'Vũ Kiều Oanh', 'Ban viên', '2025-2027', '1986-07-17', '0904067880', '194, Hồ Văn Huê, Phường 9, Quận Phú Nhuận', 'GRP_03', 'Tổ 3', '', '', '2026-01-01', '2026-08-10', 'Đang sinh hoạt', ''],
  ['MEM_030', 'Huỳnh NhuậnTâm', 'Ban viên', '2025-2027', '1986-11-07', '0906488319', '301 Lô A1, Chung cư Thạnh Mỹ Lợi F, Thạnh Mỹ Lợi, Tp. Thủ Đức', 'GRP_03', 'Tổ 3', '', '', '2026-01-01', '2026-08-10', 'Đang sinh hoạt', ''],
  ['MEM_031', 'Nguyễn Ngọc Tuân', 'Ban viên', '2025-2027', '1986-06-13', '0905058053', '', 'GRP_04', 'Tổ 4', '', '', '2026-01-01', '2026-08-10', 'Đang sinh hoạt', ''],
  ['MEM_032', 'Lê Quang Trung Tín', 'Ban viên', '2025-2027', '1985-02-21', '0919762274', 'Hommyland 3', 'GRP_04', 'Tổ 4', '', '', '2026-01-01', '2026-08-10', 'Đang sinh hoạt', ''],
  ['MEM_033', 'Châu Thị Mai', 'Ban viên', '2025-2027', '1985-12-04', '', '', 'GRP_04', 'Tổ 4', '', '', '2026-01-01', '2026-08-10', 'Đang sinh hoạt', ''],
  ['MEM_034', 'Nguyễn Thị Mỹ Ngọc', 'Ban viên', '2025-2027', '1984-11-13', '0938603284', '45/9, Đường 32, Thạnh Mỹ Lợi, Tp. Thủ Đức', 'GRP_04', 'Tổ 4', '', '', '2026-01-01', '2026-08-10', 'Đang sinh hoạt', 'Tổ trưởng Tổ 4'],
  ['MEM_035', 'Lê Khắc Đại Lộc', 'Ban viên', '2025-2027', '1984-05-23', '0947998847', '30/5, Thạnh Mỹ Lợi, Thạnh Mỹ Lợi, Tp. Thủ Đức', 'GRP_04', 'Tổ 4', '', '', '2026-01-01', '2026-08-10', 'Đang sinh hoạt', ''],
  ['MEM_036', 'Tô Bích Trâm', 'Ban viên', '2025-2027', '1984-07-24', '', '', 'GRP_04', 'Tổ 4', '', '', '2026-01-01', '2026-08-10', 'Đang sinh hoạt', ''],
  ['MEM_037', 'Thái Nhựt Bình', 'Trưởng ban', '2025-2027', '1983-03-06', '0976181237', '31, Đường 18, Bình Trưng Tây, Tp. Thủ Đức', 'GRP_04', 'Tổ 4', '', '', '2026-01-01', '2026-08-10', 'Đang sinh hoạt', 'Trưởng ban | Tài khoản: binh'],
  ['MEM_038', 'Bùi Thị Thu Hiền', 'Ban viên', '2025-2027', '1983-02-07', '0362688466', '31, Đường 18, Bình Trưng Tây, Tp. Thủ Đức', 'GRP_04', 'Tổ 4', '', '', '2026-01-01', '2026-08-10', 'Đang sinh hoạt', ''],
  ['MEM_039', 'Huỳnh Nhật Kha My', 'Ban viên', '2025-2027', '1982-03-17', '0908988913', 'Ấp Cát, Xã Phú Hữu, Nhơn Trạch', 'GRP_04', 'Tổ 4', '', '', '2026-01-01', '2026-08-10', 'Đang sinh hoạt', ''],
  ['MEM_040', 'Hồ Thị Minh Nhựt', 'Ban viên', '2025-2027', '1982-10-01', '0705705474', '', 'GRP_04', 'Tổ 4', '', '', '2026-01-01', '2026-08-10', 'Đang sinh hoạt', ''],
  ['MEM_041', 'Hồ Thị DiễmNgọc', 'Ban viên', '2025-2027', '1981-12-19', '0972682035', '', 'GRP_05', 'Tổ 5', '', '', '2026-01-01', '2026-08-10', 'Đang sinh hoạt', ''],
  ['MEM_042', 'Nguyễn Nguyên Bá', 'Ban viên', '2025-2027', '1980-10-27', '0903101754', 'CC HomyLand 2, Bình Trưng Tây, Tp. Thủ Đức', 'GRP_05', 'Tổ 5', '', '', '2026-01-01', '2026-08-10', 'Đang sinh hoạt', ''],
  ['MEM_043', 'Phan Văn Hoàng', 'Ban viên', '2025-2027', '1978-01-01', '', '', 'GRP_05', 'Tổ 5', '', '', '2026-01-01', '2026-08-10', 'Đang sinh hoạt', ''],
  ['MEM_044', 'Nguyễn Hoa Thiên Lý', 'Ban viên', '2025-2027', '1978-03-17', '0902888001', '54A, Trần Văn Giáp, Hiệp Tân, Quận Tân Phú', 'GRP_05', 'Tổ 5', '', '', '2026-01-01', '2026-08-10', 'Đang sinh hoạt', ''],
  ['MEM_045', 'Lưu Văn Minh', 'Ban viên', '2025-2027', '1978-05-18', '0398687567', '40/2, Đường 836, Phú Hữu, Tp. Thủ Đức', 'GRP_05', 'Tổ 5', '', '', '2026-01-01', '2026-08-10', 'Đang sinh hoạt', ''],
  ['MEM_046', 'Nguyễn YếnThanh', 'Ban viên', '2025-2027', '1978-08-27', '0827708278', '', 'GRP_05', 'Tổ 5', '', '', '2026-01-01', '2026-08-10', 'Đang sinh hoạt', ''],
  ['MEM_047', 'Huỳnh Tấn Trực', 'Ban viên', '2025-2027', '1978-11-16', '', '', 'GRP_05', 'Tổ 5', '', '', '2026-01-01', '2026-08-10', 'Đang sinh hoạt', ''],
  ['MEM_048', 'Đặng Xuân Hương', 'Ban viên', '2025-2027', '1977-11-30', '0908475364', 'Căn 202 - Lô B2, CC Thạnh Mỹ Lợi, Tp. Thủ Đức', 'GRP_05', 'Tổ 5', '', '', '2026-01-01', '2026-08-10', 'Đang sinh hoạt', ''],
  ['MEM_049', 'Quách Thanh Dũng', 'Ban viên', '2025-2027', '1977-12-24', '0935913441', '126, Lê Văn Thịnh, Bình Trưng Tây, Tp. Thủ Đức', 'GRP_05', 'Tổ 5', '', '', '2026-01-01', '2026-08-10', 'Đang sinh hoạt', ''],
  ['MEM_050', 'Trần Thị Sen', 'Ban viên', '2025-2027', '1977-09-10', '0965430060', '126, Lê Văn Thịnh, Bình Trưng Tây, Tp. Thủ Đức', 'GRP_05', 'Tổ 5', '', '', '2026-01-01', '2026-08-10', 'Đang sinh hoạt', ''],
  ['MEM_051', 'Nguyễn Thị Hồng Việt', 'Ban viên', '2025-2027', '1976-01-17', '0908868909', 'CC HomyLand 2, Bình Trưng Tây, Tp. Thủ Đức', 'GRP_05', 'Tổ 5', '', '', '2026-01-01', '2026-08-10', 'Đang sinh hoạt', ''],
  ['MEM_052', 'Huỳnh Thị NgọcLâm', 'Ban viên', '2025-2027', '1975-10-02', '0898275113', '84/5, Bình Trưng, Bình Trưng Đông, Tp. Thủ Đức', 'GRP_05', 'Tổ 5', '', '', '2026-01-01', '2026-08-10', 'Đang sinh hoạt', ''],
  ['MEM_053', 'Nguyễn Sơn Đông', 'Ban viên', '2025-2027', '1971-10-23', '0903305029', '54A, Trần Văn Giáp, Hiệp Tân, Quận Tân Phú', 'GRP_05', 'Tổ 5', '', '', '2026-01-01', '2026-08-10', 'Đang sinh hoạt', ''],
  ['MEM_054', 'Nguyễn Thị Ngọc Mỹ', 'Ban viên', '2025-2027', '1971-02-20', '0905297004', '898/7, Nguyễn Thị Định, Thạnh Mỹ Lợi, Tp. Thủ Đức', 'GRP_05', 'Tổ 5', '', '', '2026-01-01', '2026-08-10', 'Đang sinh hoạt', ''],
  ['MEM_055', 'Nguyễn Thị Hạ Thương', 'Ban viên', '2025-2027', '', '', '', 'GRP_05', 'Tổ 5', '', '', '2026-01-01', '2026-08-10', 'Đang sinh hoạt', '']
];

function setupManagerTeamsDatabase() {
  const ss = getSpreadsheet();
  
  const defaultSheets = [
    {
      name: SHEET_NAMES.MEMBERS,
      headers: ['Mã Thành Viên', 'Họ và Tên', 'Chức Danh', 'Nhiệm Kỳ', 'Ngày Sinh', 'Số Điện Thoại', 'Địa Chỉ', 'Mã Tổ', 'Tên Tổ', 'Ngày Báp Têm', 'Tên Vợ Chồng', 'Ngày Sinh Hoạt', 'Lần Nhóm Cuối', 'Trạng Thái', 'Ghi Chú']
    },
    {
      name: SHEET_NAMES.GROUPS,
      headers: ['Mã Tổ', 'Tên Tổ', 'Tổ Trưởng', 'Số ĐT Tổ Trưởng', 'Ghi Chú']
    },
    {
      name: SHEET_NAMES.VISITATIONS,
      headers: ['Mã Thăm Viếng', 'Mã Thành Viên', 'Họ và Tên', 'Người Thăm', 'Ngày Thăm', 'Hình Thức', 'Nhu Cầu Cầu Nguyện', 'Kết Quả & Ghi Chú']
    },
    {
      name: SHEET_NAMES.QUARTERLY_THEMES,
      headers: ['Mã Chủ Đề', 'Năm', 'Quý', 'Tên Chủ Đề', 'Câu Gốc', 'Nội Dung Câu Gốc', 'Bài Hát Chủ Đề', 'Ghi Chú']
    },
    {
      name: SHEET_NAMES.WEEKLY_SCHEDULES,
      headers: ['Mã Buổi Nhóm', 'Năm', 'Quý', 'Ngày Nhóm', 'Đề Tài', 'Câu Gốc', 'Đố Kinh Thánh', 'Tổ Phụ Trách', 'Phụ Trách (Diễn Giả)', 'Ghi Chú']
    },
    {
      name: SHEET_NAMES.ATTENDANCE_RECORDS,
      headers: ['Mã Điểm Danh', 'Ngày Nhóm', 'Mã Thành Viên', 'Họ và Tên', 'Trạng Thái', 'Thuộc Câu Gốc', 'Số Câu Đố KT', 'Ghi Chú', 'Thời Gian Ghi']
    },
    {
      name: SHEET_NAMES.ACCOUNTS,
      headers: ['Username', 'Password', 'FullName', 'Position', 'Status', 'CreatedAt']
    },
    {
      name: SHEET_NAMES.CONFIG,
      headers: ['Tên Ban Ngành', 'Thứ Nhóm', 'Giờ Nhóm', 'Thông Báo Chân Trang']
    }
  ];

  defaultSheets.forEach(item => {
    let sheet = ss.getSheetByName(item.name);
    if (!sheet) {
      sheet = ss.insertSheet(item.name);
      sheet.appendRow(item.headers);
      sheet.getRange(1, 1, 1, item.headers.length).setFontWeight('bold').setBackground('#1e3a8a').setFontColor('#ffffff');
    }
  });

  // Tự động kiểm tra và làm sạch dữ liệu 55 thành viên chính thức
  const memSheet = ss.getSheetByName(SHEET_NAMES.MEMBERS);
  if (memSheet) {
    const headers = defaultSheets[0].headers;
    memSheet.clear();
    memSheet.appendRow(headers);
    memSheet.getRange(1, 1, 1, headers.length).setFontWeight('bold').setBackground('#1e3a8a').setFontColor('#ffffff');
    INITIAL_55_MEMBERS.forEach(row => memSheet.appendRow(row));
  }

  // Khởi tạo các Tổ
  const groupSheet = ss.getSheetByName(SHEET_NAMES.GROUPS);
  if (groupSheet && groupSheet.getLastRow() <= 1) {
    const defaultGroups = [
      ['GRP_01', 'Tổ 1', 'Lê Thị Hoàng Yến', '0877115836', 'Sinh hoạt Tổ 1'],
      ['GRP_02', 'Tổ 2', 'Phạm Anh Tuấn', '0399182821', 'Sinh hoạt Tổ 2'],
      ['GRP_03', 'Tổ 3', 'Nguyễn Thị Thu Hiền', '0933833967', 'Sinh hoạt Tổ 3'],
      ['GRP_04', 'Tổ 4', 'Nguyễn Thị Mỹ Ngọc', '0938603284', 'Sinh hoạt Tổ 4'],
      ['GRP_05', 'Tổ 5', 'Thái Nhựt Bình', '0976181237', 'Sinh hoạt Tổ 5']
    ];
    defaultGroups.forEach(g => groupSheet.appendRow(g));
  }

  // Tự động khởi tạo 6 tài khoản mặc định
  const accountSheet = ss.getSheetByName(SHEET_NAMES.ACCOUNTS);
  if (accountSheet && accountSheet.getLastRow() <= 1) {
    const defaultAccounts = [
      ['binh', '123456', 'Thái Nhựt Bình', 'Trưởng ban', 'Hoạt động', new Date().toISOString()],
      ['yen', '123456', 'Lê Thị Hoàng Yến', 'Phó ban', 'Hoạt động', new Date().toISOString()],
      ['duen', '123456', 'Y Đuen Êban', 'Thư ký', 'Hoạt động', new Date().toISOString()],
      ['tuan', '123456', 'Phạm Anh Tuấn', 'Thủ quỹ', 'Hoạt động', new Date().toISOString()],
      ['nhung', '123456', 'Nguyễn Thị Nhung', 'Nghị viên', 'Hoạt động', new Date().toISOString()],
      ['hien', '123456', 'Nguyễn Thị Thu Hiền', 'Ban viên', 'Hoạt động', new Date().toISOString()]
    ];
    defaultAccounts.forEach(acc => accountSheet.appendRow(acc));
  }

  return { success: true, data: "🎉 Đã cập nhật nạp đủ 55 Ban viên chính thức và làm sạch CSDL thành công!" };
}

// =========================================================================
// QUẢN LÝ DỮ LIỆU THÀNH VIÊN VỚI DYNAMIC HEADER PARSER (CHỐNG LỆCH CỘT 100%)
// =========================================================================
function getMembersData() {
  try {
    const ss = getSpreadsheet();
    const sheet = ss.getSheetByName(SHEET_NAMES.MEMBERS);
    if (!sheet) return { success: true, data: [] };

    const rows = sheet.getDataRange().getValues();
    if (rows.length <= 1) return { success: true, data: [] };

    // Đọc động hàng Tiêu đề 1 để lấy vị trí chỉ số cột chính xác
    const headerRow = rows[0].map(h => String(h || '').toLowerCase().trim());
    
    const getCol = (possibleNames, fallbackIndex) => {
      for (let name of possibleNames) {
        const idx = headerRow.indexOf(name.toLowerCase());
        if (idx !== -1) return idx;
      }
      return fallbackIndex;
    };

    const colId = getCol(['mã thành viên', 'mã tv', 'id'], 0);
    const colName = getCol(['họ và tên', 'tên', 'họ tên', 'fullname'], 1);
    const colPos = getCol(['chức danh', 'chức vụ', 'position'], 2);
    const colTenure = getCol(['nhiệm kỳ', 'khóa'], 3);
    const colDob = getCol(['ngày sinh', 'dob', 'sinh nhật'], 4);
    const colPhone = getCol(['số điện thoại', 'sđt', 'phone'], 5);
    const colAddress = getCol(['địa chỉ', 'nơi ở', 'address'], 6);
    const colGroupId = getCol(['mã tổ', 'mã nhóm'], 7);
    const colGroupName = getCol(['tên tổ', 'tổ', 'nhóm'], 8);
    const colBaptism = getCol(['ngày bápêm', 'ngày báp tập', 'bápêm', 'bápêm', 'ngày báp têm'], 9);
    const colSpouse = getCol(['tên vợ chồng', 'vợ/chồng', 'vợ chồng'], 10);
    const colJoin = getCol(['ngày sinh hoạt', 'ngày tham gia'], 11);
    const colLastAtt = getCol(['lần nhóm cuối', 'ngày nhóm cuối', 'đi nhóm cuối'], 12);
    const colStatus = getCol(['trạng thái', 'tình trạng'], 13);
    const colNotes = getCol(['ghi chú', 'lưu ý'], 14);

    const today = new Date();
    const currentMonth = today.getMonth() + 1;

    const members = [];
    for (let i = 1; i < rows.length; i++) {
      const row = rows[i];
      const fullNameVal = String(row[colName] || '').trim();
      const memberId = String(row[colId] || '');
      
      if (!memberId && !fullNameVal) continue;

      const dobStr = formatDateStandard(row[colDob]);
      const lastAttStr = formatDateStandard(row[colLastAtt]);
      
      let isBirthdayThisMonth = false;
      let age = null;
      if (dobStr) {
        const parts = dobStr.split('-');
        if (parts.length === 3) {
          const birthMonth = parseInt(parts[1], 10);
          const birthYear = parseInt(parts[0], 10);
          if (birthMonth === currentMonth) isBirthdayThisMonth = true;
          if (!isNaN(birthYear) && birthYear > 1900) {
            age = today.getFullYear() - birthYear;
          }
        }
      }

      let daysAbsent = 0;
      let isAbsentAlert = false;
      if (lastAttStr) {
        const lastDate = new Date(lastAttStr);
        if (!isNaN(lastDate.getTime())) {
          const diffTime = Math.abs(today - lastDate);
          daysAbsent = Math.ceil(diffTime / (1000 * 60 * 60 * 24));
          if (daysAbsent >= 21) {
            isAbsentAlert = true;
          }
        }
      } else {
        isAbsentAlert = true;
        daysAbsent = 30;
      }

      members.push({
        memberId: memberId || ('MEM_' + String(i).padStart(3, '0')),
        fullName: fullNameVal,
        position: String(row[colPos] || 'Ban viên'),
        tenure: String(row[colTenure] || '2025-2027'),
        dateOfBirth: dobStr,
        phone: String(row[colPhone] || ''),
        address: String(row[colAddress] || ''),
        groupId: String(row[colGroupId] || 'GRP_01'),
        groupName: String(row[colGroupName] || 'Tổ 1'),
        baptismDate: formatDateStandard(row[colBaptism]),
        spouseName: String(row[colSpouse] || ''),
        joinDate: formatDateStandard(row[colJoin]),
        lastAttendanceDate: lastAttStr,
        status: String(row[colStatus] || 'Đang sinh hoạt'),
        notes: String(row[colNotes] || ''),
        isBirthdayThisMonth: isBirthdayThisMonth,
        age: age,
        daysAbsent: daysAbsent,
        isAbsentAlert: isAbsentAlert
      });
    }

    return { success: true, data: members };
  } catch (err) {
    return { success: false, error: err.message };
  }
}

function saveMember(memberData) {
  try {
    const ss = getSpreadsheet();
    let sheet = ss.getSheetByName(SHEET_NAMES.MEMBERS);
    if (!sheet) {
      setupManagerTeamsDatabase();
      sheet = ss.getSheetByName(SHEET_NAMES.MEMBERS);
    }

    const rows = sheet.getDataRange().getValues();
    let memberId = memberData.memberId;
    let foundIndex = -1;

    if (memberId) {
      for (let i = 1; i < rows.length; i++) {
        if (String(rows[i][0]) === memberId) {
          foundIndex = i + 1;
          break;
        }
      }
    }

    if (!memberId) {
      memberId = 'MEM_' + Date.now().toString().slice(-6);
    }

    const rowValue = [
      memberId,
      memberData.fullName,
      memberData.position || 'Ban viên',
      memberData.tenure || '2025-2027',
      memberData.dateOfBirth || '',
      memberData.phone || '',
      memberData.address || '',
      memberData.groupId || '',
      memberData.groupName || '',
      memberData.baptismDate || '',
      memberData.spouseName || '',
      memberData.joinDate || '',
      memberData.lastAttendanceDate || '',
      memberData.status || 'Đang sinh hoạt',
      memberData.notes || ''
    ];

    if (foundIndex > 0) {
      sheet.getRange(foundIndex, 1, 1, rowValue.length).setValues([rowValue]);
    } else {
      sheet.appendRow(rowValue);
    }

    return { success: true, data: { memberId }, error: "Lưu thành viên thành công!" };
  } catch (err) {
    return { success: false, error: err.message };
  }
}

function deleteMember(memberId) {
  try {
    const ss = getSpreadsheet();
    const sheet = ss.getSheetByName(SHEET_NAMES.MEMBERS);
    if (!sheet) return { success: false, error: "Bảng dữ liệu không tồn tại!" };

    const rows = sheet.getDataRange().getValues();
    for (let i = 1; i < rows.length; i++) {
      if (String(rows[i][0]) === memberId) {
        sheet.deleteRow(i + 1);
        return { success: true, error: "Đã xóa thành viên thành công!" };
      }
    }
    return { success: false, error: "Không tìm thấy thành viên để xóa!" };
  } catch (err) {
    return { success: false, error: err.message };
  }
}

function getMemberProfileDetail(memberId) {
  try {
    const ss = getSpreadsheet();
    const memRes = getMembersData();
    const members = memRes.success ? memRes.data : [];
    const member = members.find(m => m.memberId === memberId);

    if (!member) {
      return { success: false, error: "Không tìm thấy ban viên!" };
    }

    // Lấy 12 tuần điểm danh gần nhất của ban viên này
    let attendanceHistory = [];
    const attSheet = ss.getSheetByName(SHEET_NAMES.ATTENDANCE_RECORDS);
    if (attSheet && attSheet.getLastRow() > 1) {
      const attRows = attSheet.getDataRange().getValues();
      
      // Tìm các ngày nhóm duy nhất trong CSDL
      const dateSet = new Set();
      for (let i = 1; i < attRows.length; i++) {
        const d = formatDateStandard(attRows[i][1]);
        if (d) dateSet.add(d);
      }
      const sortedDates = Array.from(dateSet).sort().reverse().slice(0, 12);

      // Tra cứu kết quả từng ngày cho ban viên này
      sortedDates.forEach(dateStr => {
        let status = 'Vắng';
        let isVerse = false;
        let score = 0;
        let found = false;

        for (let i = 1; i < attRows.length; i++) {
          const rDate = formatDateStandard(attRows[i][1]);
          const mId = String(attRows[i][2]);
          if (rDate === dateStr && mId === memberId) {
            status = String(attRows[i][4] || 'Có mặt');
            const vVal = attRows[i][5];
            isVerse = String(vVal).toLowerCase() === 'có' || String(vVal).toLowerCase() === 'true';
            score = Number(attRows[i][6]) || (status === 'Có mặt' ? 10 : 0);
            found = true;
            break;
          }
        }

        attendanceHistory.push({
          date: dateStr,
          status: status,
          isVerse: isVerse,
          score: score,
          recorded: found
        });
      });
    }

    // Lấy nhật ký thăm viếng của ban viên này
    let visitationLogs = [];
    const visSheet = ss.getSheetByName(SHEET_NAMES.VISITATIONS);
    if (visSheet && visSheet.getLastRow() > 1) {
      const visRows = visSheet.getDataRange().getValues();
      for (let i = 1; i < visRows.length; i++) {
        if (String(visRows[i][1]) === memberId) {
          visitationLogs.push({
            visitationId: String(visRows[i][0]),
            visitorName: String(visRows[i][3]),
            visitDate: formatDateStandard(visRows[i][4]),
            visitType: String(visRows[i][5]),
            prayerRequests: String(visRows[i][6] || ''),
            resultNotes: String(visRows[i][7] || '')
          });
        }
      }
    }

    return {
      success: true,
      data: {
        member: member,
        attendanceHistory: attendanceHistory,
        visitationLogs: visitationLogs
      }
    };
  } catch (err) {
    return { success: false, error: err.message };
  }
}

function batchUpdateMemberGroup(memberIds, newGroupId, newGroupName) {
  const lock = LockService.getScriptLock();
  const success = lock.tryLock(10000);
  if (!success) {
    return { success: false, error: "Hệ thống đang bận ghi dữ liệu. Vui lòng thử lại!" };
  }

  try {
    const ss = getSpreadsheet();
    const sheet = ss.getSheetByName(SHEET_NAMES.MEMBERS);
    if (!sheet || sheet.getLastRow() <= 1) {
      return { success: false, error: "Bảng Members rỗng!" };
    }

    const rows = sheet.getDataRange().getValues();
    let updatedCount = 0;

    for (let i = 1; i < rows.length; i++) {
      const mId = String(rows[i][0]);
      if (memberIds.includes(mId)) {
        sheet.getRange(i + 1, 8).setValue(newGroupId); // Mã Tổ (Cột H)
        sheet.getRange(i + 1, 9).setValue(newGroupName); // Tên Tổ (Cột I)
        updatedCount++;
      }
    }

    return { success: true, error: `Đã chuyển ${updatedCount} thành viên sang ${newGroupName} thành công!` };
  } catch (err) {
    return { success: false, error: err.message };
  } finally {
    lock.releaseLock();
  }
}

function batchUpdateMemberStatus(memberIds, newStatus) {
  const lock = LockService.getScriptLock();
  const success = lock.tryLock(10000);
  if (!success) {
    return { success: false, error: "Hệ thống đang bận ghi dữ liệu. Vui lòng thử lại!" };
  }

  try {
    const ss = getSpreadsheet();
    const sheet = ss.getSheetByName(SHEET_NAMES.MEMBERS);
    if (!sheet || sheet.getLastRow() <= 1) {
      return { success: false, error: "Bảng Members rỗng!" };
    }

    const rows = sheet.getDataRange().getValues();
    let updatedCount = 0;

    for (let i = 1; i < rows.length; i++) {
      const mId = String(rows[i][0]);
      if (memberIds.includes(mId)) {
        sheet.getRange(i + 1, 14).setValue(newStatus); // Trạng thái (Cột N)
        updatedCount++;
      }
    }

    return { success: true, error: `Đã cập nhật trạng thái thành "${newStatus}" cho ${updatedCount} thành viên!` };
  } catch (err) {
    return { success: false, error: err.message };
  } finally {
    lock.releaseLock();
  }
}

// =========================================================================
// HỆ THỐNG XÁC THỰC ĐĂNG NHẬP & QUẢN LÝ TÀI KHOẢN
// =========================================================================
function verifyAccountLogin(username, password) {
  try {
    const ss = getSpreadsheet();
    const sheet = ss.getSheetByName(SHEET_NAMES.ACCOUNTS);

    username = (username || '').toLowerCase().trim();
    password = (password || '').trim();

    if (!username || !password) {
      return { success: false, error: 'Vui lòng nhập đầy đủ tên đăng nhập và mật khẩu!' };
    }

    if (!sheet || sheet.getLastRow() <= 1) {
      setupManagerTeamsDatabase();
    }

    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      const u = String(data[i][0] || '').toLowerCase().trim();
      const p = String(data[i][1] || '').trim();
      const fullName = String(data[i][2] || '');
      const position = String(data[i][3] || 'Ban viên');
      const status = String(data[i][4] || 'Hoạt động');

      if (u === username) {
        if (p === password) {
          if (status === 'Bị khóa') {
            return { success: false, error: 'Tài khoản này hiện đang bị khóa!' };
          }
          return {
            success: true,
            user: {
              username: u,
              fullName: fullName || u,
              position: position
            }
          };
        } else {
          return { success: false, error: 'Mật khẩu nhập không đúng!' };
        }
      }
    }

    return { success: false, error: 'Tài khoản không tồn tại trên hệ thống!' };
  } catch (err) {
    return { success: false, error: err.message };
  }
}

function getAccountsData() {
  try {
    const ss = getSpreadsheet();
    const sheet = ss.getSheetByName(SHEET_NAMES.ACCOUNTS);
    if (!sheet || sheet.getLastRow() <= 1) return { success: true, data: [] };

    const rows = sheet.getDataRange().getValues();
    const accounts = [];
    for (let i = 1; i < rows.length; i++) {
      accounts.push({
        username: String(rows[i][0]),
        password: String(rows[i][1]),
        fullName: String(rows[i][2]),
        position: String(rows[i][3]),
        status: String(rows[i][4] || 'Hoạt động'),
        createdAt: String(rows[i][5] || '')
      });
    }
    return { success: true, data: accounts };
  } catch (err) {
    return { success: false, error: err.message };
  }
}

function saveAccount(accData) {
  try {
    const ss = getSpreadsheet();
    let sheet = ss.getSheetByName(SHEET_NAMES.ACCOUNTS);
    if (!sheet) {
      setupManagerTeamsDatabase();
      sheet = ss.getSheetByName(SHEET_NAMES.ACCOUNTS);
    }

    const username = String(accData.username || '').toLowerCase().trim();
    const password = String(accData.password || '123456').trim();
    const fullName = String(accData.fullName || '').trim();
    const position = String(accData.position || 'Ban viên');
    const status = String(accData.status || 'Hoạt động');

    if (!username) return { success: false, error: 'Tên đăng nhập không được để trống!' };

    const rows = sheet.getDataRange().getValues();
    let foundIndex = -1;

    for (let i = 1; i < rows.length; i++) {
      if (String(rows[i][0]).toLowerCase().trim() === username) {
        foundIndex = i + 1;
        break;
      }
    }

    if (foundIndex > 0) {
      sheet.getRange(foundIndex, 2, 1, 4).setValues([[password, fullName, position, status]]);
    } else {
      sheet.appendRow([username, password, fullName, position, status, new Date().toISOString()]);
    }

    return { success: true, error: 'Đã lưu thông tin tài khoản thành công!' };
  } catch (err) {
    return { success: false, error: err.message };
  }
}

function deleteAccount(username) {
  try {
    const ss = getSpreadsheet();
    const sheet = ss.getSheetByName(SHEET_NAMES.ACCOUNTS);
    if (!sheet) return { success: false, error: 'Chưa có bảng tài khoản!' };

    username = String(username).toLowerCase().trim();
    const rows = sheet.getDataRange().getValues();

    for (let i = 1; i < rows.length; i++) {
      if (String(rows[i][0]).toLowerCase().trim() === username) {
        sheet.deleteRow(i + 1);
        return { success: true, error: `Đã xóa tài khoản "${username}" thành công!` };
      }
    }

    return { success: false, error: 'Không tìm thấy tài khoản để xóa!' };
  } catch (err) {
    return { success: false, error: err.message };
  }
}

function changeAccountPassword(username, oldPassword, newPassword) {
  try {
    const ss = getSpreadsheet();
    const sheet = ss.getSheetByName(SHEET_NAMES.ACCOUNTS);
    if (!sheet) return { success: false, error: 'Chưa có dữ liệu!' };

    username = String(username).toLowerCase().trim();
    const rows = sheet.getDataRange().getValues();

    for (let i = 1; i < rows.length; i++) {
      if (String(rows[i][0]).toLowerCase().trim() === username) {
        const currentPass = String(rows[i][1]).trim();
        if (currentPass !== String(oldPassword).trim()) {
          return { success: false, error: 'Mật khẩu cũ nhập không đúng!' };
        }
        sheet.getRange(i + 1, 2).setValue(String(newPassword).trim());
        return { success: true, error: 'Đã đổi mật khẩu thành công!' };
      }
    }

    return { success: false, error: 'Không tìm thấy tài khoản!' };
  } catch (err) {
    return { success: false, error: err.message };
  }
}

function getDepartmentConfig() {
  try {
    const ss = getSpreadsheet();
    const sheet = ss.getSheetByName(SHEET_NAMES.CONFIG);
    if (!sheet || sheet.getLastRow() <= 1) {
      return {
        success: true,
        data: {
          departmentName: 'BAN THANH TRÁNG',
          meetingDay: 'Chúa Nhật',
          meetingTime: '16h00 chiều',
          footerNotice: 'Chương trình này thay lời mời, xin các Anh/Chị thêm lời cầu nguyện và tham dự đầy đủ.'
        }
      };
    }
    const row = sheet.getRange(2, 1, 1, 4).getValues()[0];
    return {
      success: true,
      data: {
        departmentName: row[0] || 'BAN THANH TRÁNG',
        meetingDay: row[1] || 'Chúa Nhật',
        meetingTime: row[2] || '16h00 chiều',
        footerNotice: row[3] || 'Chương trình này thay lời mời, xin các Anh/Chị thêm lời cầu nguyện và tham dự đầy đủ.'
      }
    };
  } catch (err) {
    return { success: false, error: err.message };
  }
}

function saveDepartmentConfig(config) {
  try {
    const ss = getSpreadsheet();
    let sheet = ss.getSheetByName(SHEET_NAMES.CONFIG);
    if (!sheet) {
      setupManagerTeamsDatabase();
      sheet = ss.getSheetByName(SHEET_NAMES.CONFIG);
    }
    if (sheet.getLastRow() <= 1) {
      sheet.appendRow([config.departmentName, config.meetingDay, config.meetingTime, config.footerNotice]);
    } else {
      sheet.getRange(2, 1, 1, 4).setValues([[config.departmentName, config.meetingDay, config.meetingTime, config.footerNotice]]);
    }
    return { success: true };
  } catch (err) {
    return { success: false, error: err.message };
  }
}

function verifyUserPin(pin) {
  pin = String(pin).trim();
  if (pin === '1234') {
    return { success: true, data: { role: 'Admin', title: 'Ban Điều Hành (Admin)' } };
  } else if (pin === '5678') {
    return { success: true, data: { role: 'Care_Leader', title: 'Tổ Trưởng / Nghị Viên' } };
  } else {
    return { success: false, error: 'Mã PIN không đúng!' };
  }
}

function sendExecutiveSummaryEmail(recipientEmail) {
  try {
    if (!recipientEmail || !recipientEmail.includes('@')) {
      return { success: false, error: 'Địa chỉ Email không hợp lệ!' };
    }

    const membersRes = getMembersData();
    const members = membersRes.success ? membersRes.data : [];
    const absentAlerts = members.filter(m => m.isAbsentAlert || m.status === 'Cần thăm viếng');
    const birthdays = members.filter(m => m.isBirthdayThisMonth);

    let htmlBody = `
      <div style="font-family: Arial, sans-serif; padding: 20px; color: #1e293b;">
        <h2 style="color: #1e3a8a;">⛪ BÁO CÁO TỔNG HỢP BAN THẮNG TRÁNG</h2>
        <p>Kính gửi Ban Điều Hành,</p>
        <p>Hệ thống ManagerTeams xin gửi báo cáo tình hình nhân sự mới nhất:</p>
        
        <h3 style="color: #dc2626;">🚑 Ban Viên Vắng 3-4 Tuần Cần Thăm Viếng (${absentAlerts.length})</h3>
        <table border="1" cellpadding="8" cellspacing="0" style="border-collapse: collapse; width: 100%;">
          <tr style="background: #fee2e2; color: #991b1b;">
            <th>Họ và Tên</th><th>Chức Danh</th><th>Tổ</th><th>SĐT</th><th>Số Ngày Vắng</th>
          </tr>
          ${absentAlerts.map(m => `
            <tr>
              <td><strong>${m.fullName}</strong></td>
              <td>${m.position}</td>
              <td>${m.groupName}</td>
              <td>${m.phone}</td>
              <td><span style="color:red; font-weight:bold;">${m.daysAbsent} ngày</span></td>
            </tr>
          `).join('')}
        </table>

        <h3 style="color: #d97706; margin-top: 20px;">🎂 Sinh Nhật Trong Tháng (${birthdays.length})</h3>
        <ul>
          ${birthdays.map(m => `<li><strong>${m.fullName}</strong> - Ngày sinh: ${m.dateOfBirth} (${m.groupName})</li>`).join('')}
        </ul>

        <p style="margin-top: 30px; font-style: italic; color: #64748b;">Trân trọng,<br>Hệ thống Quản lý ManagerTeams</p>
      </div>
    `;

    MailApp.sendEmail({
      to: recipientEmail,
      subject: `[ManagerTeams] Báo Cáo Chăm Sóc & Thăm Viếng Ban Ngành - ${new Date().toLocaleDateString('vi-VN')}`,
      htmlBody: htmlBody
    });

    return { success: true, error: `Đã gửi báo cáo thành công tới ${recipientEmail}!` };
  } catch (err) {
    return { success: false, error: err.message };
  }
}

function getWeeklyAttendanceForDate(dateStr) {
  try {
    const ss = getSpreadsheet();
    const membersRes = getMembersData();
    const members = membersRes.success ? membersRes.data : [];

    const attSheet = ss.getSheetByName(SHEET_NAMES.ATTENDANCE_RECORDS);
    let existingRecords = {};
    let recentDates = [];

    if (attSheet && attSheet.getLastRow() > 1) {
      const rows = attSheet.getDataRange().getValues();
      
      // Lấy danh sách các ngày nhóm gần đây để tính tần suất tham gia
      const dateSet = new Set();
      for (let i = 1; i < rows.length; i++) {
        const d = formatDateStandard(rows[i][1]);
        if (d) dateSet.add(d);
      }
      recentDates = Array.from(dateSet).sort().reverse().slice(0, 4);

      // Đọc bản ghi của ngày dateStr
      for (let i = 1; i < rows.length; i++) {
        const recordDate = formatDateStandard(rows[i][1]);
        if (recordDate === dateStr) {
          const mId = String(rows[i][2]);
          const statusVal = String(rows[i][4] || 'Có mặt');
          let verseVal = rows[i][5];
          let scoreVal = rows[i][6];
          let notesVal = String(rows[i][7] || rows[i][5] || ''); // Hỗ trợ cả schema cũ & mới

          // Xử lý kiểu dữ liệu câu gốc & điểm KT
          let isVerse = true;
          if (verseVal !== undefined && verseVal !== null && verseVal !== '') {
            isVerse = String(verseVal).toLowerCase() === 'có' || String(verseVal).toLowerCase() === 'true';
          } else {
            isVerse = statusVal === 'Có mặt';
          }

          let quizCount = Number(scoreVal);
          if (isNaN(quizCount)) {
            quizCount = 0;
          } else if (quizCount > 5) {
            quizCount = 5; // Chuẩn hóa từ dữ liệu 10 điểm cũ xuống tối đa 5 câu
          }

          existingRecords[mId] = {
            status: statusVal,
            isVerseMemorized: isVerse,
            bibleQuizCount: quizCount,
            notes: notesVal
          };
        }
      }
    }

    // Tính toán tần suất tham gia của từng thành viên
    let memberAttendanceCounts = {};
    if (attSheet && attSheet.getLastRow() > 1 && recentDates.length > 0) {
      const rows = attSheet.getDataRange().getValues();
      for (let i = 1; i < rows.length; i++) {
        const rDate = formatDateStandard(rows[i][1]);
        const mId = String(rows[i][2]);
        const status = String(rows[i][4]);
        if (recentDates.includes(rDate) && status === 'Có mặt') {
          memberAttendanceCounts[mId] = (memberAttendanceCounts[mId] || 0) + 1;
        }
      }
    }

    // Lấy thông tin bài học & câu gốc tuần đó từ WEEKLY_SCHEDULES
    let weeklyTopic = { title: '', verse: '', quiz: '' };
    const schedSheet = ss.getSheetByName(SHEET_NAMES.WEEKLY_SCHEDULES);
    if (schedSheet && schedSheet.getLastRow() > 1) {
      const schedRows = schedSheet.getDataRange().getValues();
      for (let i = 1; i < schedRows.length; i++) {
        const d = formatDateStandard(schedRows[i][3]);
        if (d === dateStr) {
          weeklyTopic = {
            title: String(schedRows[i][4] || ''),
            verse: String(schedRows[i][5] || ''),
            quiz: String(schedRows[i][6] || '')
          };
          break;
        }
      }
    }

    const checkinList = members.map(m => {
      const rec = existingRecords[m.memberId];
      const count = memberAttendanceCounts[m.memberId] || 0;
      const totalRecent = Math.max(recentDates.length, 1);
      const rate = Math.round((count / totalRecent) * 100);

      // Ban viên đi thường xuyên nếu tỉ lệ >= 75% hoặc là Nhân sự/Ban điều hành
      const isLeader = ['Trưởng ban', 'Phó ban', 'Thư ký', 'Thủ quỹ', 'Nghị viên'].includes(m.position);
      const isFrequent = rate >= 75 || isLeader || recentDates.length === 0;

      const isPresent = rec ? rec.status === 'Có mặt' : true;
      const verseMemo = rec ? rec.isVerseMemorized : isPresent;
      const qCount = rec ? rec.bibleQuizCount : 0;

      return {
        memberId: m.memberId,
        fullName: m.fullName,
        groupName: m.groupName || 'Chưa xếp tổ',
        position: m.position || 'Ban viên',
        status: isPresent ? 'Có mặt' : 'Vắng',
        isVerseMemorized: verseMemo,
        bibleQuizCount: Math.min(Math.max(qCount, 0), 5),
        notes: rec ? rec.notes : '',
        isFrequent: isFrequent,
        attendanceRate: rate
      };
    });

    return {
      success: true,
      data: {
        date: dateStr,
        records: checkinList,
        weeklyTopic: weeklyTopic
      }
    };
  } catch (err) {
    return { success: false, error: err.message };
  }
}

function saveWeeklyAttendanceBatch(attendanceList, dateStr) {
  const lock = LockService.getScriptLock();
  const success = lock.tryLock(10000);
  if (!success) {
    return { success: false, error: 'Hệ thống đang bận do có người khác đang ghi dữ liệu. Vui lòng thử lại sau!' };
  }

  try {
    const ss = getSpreadsheet();
    let sheet = ss.getSheetByName(SHEET_NAMES.ATTENDANCE_RECORDS);
    if (!sheet) {
      setupManagerTeamsDatabase();
      sheet = ss.getSheetByName(SHEET_NAMES.ATTENDANCE_RECORDS);
    }

    const memSheet = ss.getSheetByName(SHEET_NAMES.MEMBERS);
    const timestamp = new Date().toISOString();

    // Xóa các dòng điểm danh cũ của ngày dateStr để ghi mới sạch sẽ (tránh nhân đôi dòng)
    if (sheet.getLastRow() > 1) {
      const rows = sheet.getDataRange().getValues();
      for (let i = rows.length - 1; i >= 1; i--) {
        const rDate = formatDateStandard(rows[i][1]);
        if (rDate === dateStr) {
          sheet.deleteRow(i + 1);
        }
      }
    }

    const batchRows = [];
    const presentMemberIds = [];

    attendanceList.forEach(item => {
      const attId = `ATT_${dateStr.replace(/-/g, '')}_${item.memberId}`;
      const isPresent = item.status === 'Có mặt';
      const verseMemo = item.isVerseMemorized ? 'Có' : 'Không';
      const quizCount = Math.min(Math.max(Number(item.bibleQuizCount || item.bibleScore || 0), 0), 5);
      const notes = item.notes || '';

      batchRows.push([
        attId,
        dateStr,
        item.memberId,
        item.fullName,
        item.status,
        verseMemo,
        quizCount,
        notes,
        timestamp
      ]);

      if (isPresent) {
        presentMemberIds.push(item.memberId);
      }
    });

    if (batchRows.length > 0) {
      const startRow = sheet.getLastRow() + 1;
      sheet.getRange(startRow, 1, batchRows.length, batchRows[0].length).setValues(batchRows);
    }

    // Cập nhật ngày nhóm gần nhất cho các thành viên có mặt vào bảng Members
    if (presentMemberIds.length > 0 && memSheet && memSheet.getLastRow() > 1) {
      const memValues = memSheet.getDataRange().getValues();
      for (let i = 1; i < memValues.length; i++) {
        const mId = String(memValues[i][0]);
        if (presentMemberIds.includes(mId)) {
          memSheet.getRange(i + 1, 13).setValue(dateStr);
          memSheet.getRange(i + 1, 14).setValue('Đang sinh hoạt');
        }
      }
    }

    return { success: true, error: `Đã lưu điểm danh ngày ${dateStr} thành công cho ${attendanceList.length} ban viên!` };
  } catch (err) {
    return { success: false, error: err.message };
  } finally {
    lock.releaseLock();
  }
}

function updateQuickAttendance(memberIds, dateStr) {
  try {
    if (!memberIds || !Array.isArray(memberIds) || memberIds.length === 0) {
      return { success: false, error: 'Chưa chọn thành viên nào!' };
    }
    const targetDate = dateStr || formatDateStandard(new Date());
    const memRes = getMembersData();
    const members = memRes.success ? memRes.data : [];

    const attendanceList = memberIds.map(id => {
      const m = members.find(x => x.memberId === id);
      return {
        memberId: id,
        fullName: m ? m.fullName : '',
        status: 'Có mặt',
        isVerseMemorized: true,
        bibleScore: 10,
        notes: 'Điểm danh nhanh 1-Click'
      };
    });

    return saveWeeklyAttendanceBatch(attendanceList, targetDate);
  } catch (err) {
    return { success: false, error: err.message };
  }
}

function getAttendanceAnalytics(year, quarter) {
  try {
    const ss = getSpreadsheet();
    const groupsRes = getGroupsData();
    const groups = groupsRes.success ? groupsRes.data : [];

    const sheet = ss.getSheetByName(SHEET_NAMES.ATTENDANCE_RECORDS);
    if (!sheet || sheet.getLastRow() <= 1) {
      return { success: true, data: { groupStats: [] } };
    }

    const memRes = getMembersData();
    const members = memRes.success ? memRes.data : [];

    const rows = sheet.getDataRange().getValues();
    let groupCounts = {};

    groups.forEach(g => {
      groupCounts[g.groupName] = { total: 0, present: 0 };
    });

    for (let i = 1; i < rows.length; i++) {
      const mId = String(rows[i][2]);
      const status = String(rows[i][4]);
      const mem = members.find(m => m.memberId === mId);

      if (mem && mem.groupName && groupCounts[mem.groupName]) {
        groupCounts[mem.groupName].total++;
        if (status === 'Có mặt') {
          groupCounts[mem.groupName].present++;
        }
      }
    }

    const groupStats = Object.keys(groupCounts).map(gName => {
      const item = groupCounts[gName];
      const rate = item.total > 0 ? Math.round((item.present / item.total) * 100) : 0;
      return {
        groupName: gName,
        total: item.total,
        present: item.present,
        rate: rate
      };
    });

    return { success: true, data: { groupStats } };
  } catch (err) {
    return { success: false, error: err.message };
  }
}

function getDashboardStats() {
  try {
    const membersRes = getMembersData();
    const groupsRes = getGroupsData();
    
    if (!membersRes.success) throw new Error(membersRes.error);
    
    const members = membersRes.data || [];
    const groups = groupsRes.success ? groupsRes.data : [];

    const totalMembers = members.length;
    const absentAlertMembers = members.filter(m => m.isAbsentAlert || m.status === 'Cần thăm viếng').length;
    const birthdayMembers = members.filter(m => m.isBirthdayThisMonth).length;
    const totalGroups = groups.length;

    return {
      success: true,
      data: {
        totalMembers,
        absentAlertMembers,
        birthdayMembers,
        totalGroups
      }
    };
  } catch (err) {
    return { success: false, error: err.message };
  }
}

function getGroupsData() {
  try {
    const ss = getSpreadsheet();
    const sheet = ss.getSheetByName(SHEET_NAMES.GROUPS);
    if (!sheet) return { success: true, data: [] };

    const rows = sheet.getDataRange().getValues();
    if (rows.length <= 1) return { success: true, data: [] };

    const membersRes = getMembersData();
    const members = membersRes.success ? membersRes.data : [];

    const groups = [];
    for (let i = 1; i < rows.length; i++) {
      const groupId = String(rows[i][0] || '');
      if (!groupId) continue;

      const memberCount = members.filter(m => m.groupId === groupId).length;

      groups.push({
        groupId: groupId,
        groupName: String(rows[i][1] || ''),
        leaderName: String(rows[i][2] || ''),
        leaderPhone: String(rows[i][3] || ''),
        notes: String(rows[i][4] || ''),
        memberCount: memberCount
      });
    }

    return { success: true, data: groups };
  } catch (err) {
    return { success: false, error: err.message };
  }
}

function saveGroup(groupData) {
  try {
    const ss = getSpreadsheet();
    let sheet = ss.getSheetByName(SHEET_NAMES.GROUPS);
    if (!sheet) {
      setupManagerTeamsDatabase();
      sheet = ss.getSheetByName(SHEET_NAMES.GROUPS);
    }

    const rows = sheet.getDataRange().getValues();
    let groupId = groupData.groupId;
    let foundIndex = -1;

    if (groupId) {
      for (let i = 1; i < rows.length; i++) {
        if (String(rows[i][0]) === groupId) {
          foundIndex = i + 1;
          break;
        }
      }
    }

    if (!groupId) {
      groupId = 'GRP_' + Date.now().toString().slice(-4);
    }

    const rowValue = [
      groupId,
      groupData.groupName,
      groupData.leaderName || '',
      groupData.leaderPhone || '',
      groupData.notes || ''
    ];

    if (foundIndex > 0) {
      sheet.getRange(foundIndex, 1, 1, rowValue.length).setValues([rowValue]);
    } else {
      sheet.appendRow(rowValue);
    }

    return { success: true, data: { groupId }, error: "Lưu Tổ thành công!" };
  } catch (err) {
    return { success: false, error: err.message };
  }
}

function deleteGroup(groupId) {
  try {
    const ss = getSpreadsheet();
    const sheet = ss.getSheetByName(SHEET_NAMES.GROUPS);
    if (!sheet) return { success: false, error: "Không tìm thấy bảng CSDL Tổ!" };

    const rows = sheet.getDataRange().getValues();
    for (let i = 1; i < rows.length; i++) {
      if (String(rows[i][0]) === groupId) {
        sheet.deleteRow(i + 1);
        return { success: true, error: "Đã xóa Tổ thành công!" };
      }
    }
    return { success: false, error: "Không tìm thấy Tổ để xóa!" };
  } catch (err) {
    return { success: false, error: err.message };
  }
}

function getVisitationLogs() {
  try {
    const ss = getSpreadsheet();
    const sheet = ss.getSheetByName(SHEET_NAMES.VISITATIONS);
    if (!sheet) return { success: true, data: [] };

    const rows = sheet.getDataRange().getValues();
    if (rows.length <= 1) return { success: true, data: [] };

    const visitations = [];
    for (let i = rows.length - 1; i >= 1; i--) {
      const visId = String(rows[i][0] || '');
      if (!visId) continue;

      visitations.push({
        visitationId: visId,
        memberId: String(rows[i][1] || ''),
        memberName: String(rows[i][2] || ''),
        visitorName: String(rows[i][3] || ''),
        visitDate: formatDateStandard(rows[i][4]),
        visitType: String(rows[i][5] || ''),
        prayerRequests: String(rows[i][6] || ''),
        resultNotes: String(rows[i][7] || '')
      });
    }

    return { success: true, data: visitations };
  } catch (err) {
    return { success: false, error: err.message };
  }
}

function saveVisitationLog(logData) {
  try {
    const repo = new VisitationRepository();
    const result = repo.saveVisitationLog(logData);
    return { success: true, logId: result.logId, error: "Đã lưu nhật ký thăm viếng thành công!" };
  } catch (err) {
    return { success: false, error: err.message };
  }
}

function deleteVisitationLog(logId) {
  try {
    const repo = new VisitationRepository();
    const result = repo.deleteVisitationLog(logId);
    return result;
  } catch (err) {
    return { success: false, error: err.message };
  }
}

function getQuarterlyThemes() {
  try {
    const ss = getSpreadsheet();
    const sheet = ss.getSheetByName(SHEET_NAMES.QUARTERLY_THEMES);
    if (!sheet) return { success: true, data: [] };

    const rows = sheet.getDataRange().getValues();
    if (rows.length <= 1) return { success: true, data: [] };

    const themes = [];
    for (let i = 1; i < rows.length; i++) {
      const themeId = String(rows[i][0] || '');
      if (!themeId) continue;

      themes.push({
        themeId: themeId,
        year: String(rows[i][1] || ''),
        quarter: String(rows[i][2] || ''),
        themeName: String(rows[i][3] || ''),
        keyVerse: String(rows[i][4] || ''),
        keyVerseContent: String(rows[i][5] || ''),
        themeSong: String(rows[i][6] || ''),
        notes: String(rows[i][7] || '')
      });
    }

    return { success: true, data: themes };
  } catch (err) {
    return { success: false, error: err.message };
  }
}

function saveQuarterlyTheme(themeData) {
  try {
    const ss = getSpreadsheet();
    let sheet = ss.getSheetByName(SHEET_NAMES.QUARTERLY_THEMES);
    if (!sheet) {
      setupManagerTeamsDatabase();
      sheet = ss.getSheetByName(SHEET_NAMES.QUARTERLY_THEMES);
    }

    const rows = sheet.getDataRange().getValues();
    let themeId = themeData.themeId;
    let foundIndex = -1;

    if (themeId) {
      for (let i = 1; i < rows.length; i++) {
        if (String(rows[i][0]) === themeId) {
          foundIndex = i + 1;
          break;
        }
      }
    }

    if (!themeId) {
      themeId = 'THM_' + Date.now().toString().slice(-6);
    }

    const rowValue = [
      themeId,
      themeData.year,
      themeData.quarter,
      themeData.themeName,
      themeData.keyVerse || '',
      themeData.keyVerseContent || '',
      themeData.themeSong || '',
      themeData.notes || ''
    ];

    if (foundIndex > 0) {
      sheet.getRange(foundIndex, 1, 1, rowValue.length).setValues([rowValue]);
    } else {
      sheet.appendRow(rowValue);
    }

    return { success: true, error: "Lưu chủ đề thành công!" };
  } catch (err) {
    return { success: false, error: err.message };
  }
}

function deleteQuarterlyTheme(themeId) {
  try {
    const ss = getSpreadsheet();
    const sheet = ss.getSheetByName(SHEET_NAMES.QUARTERLY_THEMES);
    if (!sheet) return { success: false, error: "Không tìm thấy CSDL chủ đề!" };

    const rows = sheet.getDataRange().getValues();
    for (let i = 1; i < rows.length; i++) {
      if (String(rows[i][0]) === themeId) {
        sheet.deleteRow(i + 1);
        return { success: true, error: "Đã xóa chủ đề thành công!" };
      }
    }
    return { success: false, error: "Không tìm thấy chủ đề để xóa!" };
  } catch (err) {
    return { success: false, error: err.message };
  }
}

function getWeeklySchedules(year, quarter) {
  try {
    const ss = getSpreadsheet();
    const sheet = ss.getSheetByName(SHEET_NAMES.WEEKLY_SCHEDULES);
    if (!sheet) return { success: true, data: [] };

    const rows = sheet.getDataRange().getValues();
    if (rows.length <= 1) return { success: true, data: [] };

    const schedules = [];
    for (let i = 1; i < rows.length; i++) {
      const rowYear = String(rows[i][1] || '');
      const rowQuarter = String(rows[i][2] || '');

      if ((!year || rowYear === year) && (!quarter || rowQuarter === quarter)) {
        schedules.push({
          scheduleId: String(rows[i][0] || ''),
          year: rowYear,
          quarter: rowQuarter,
          meetingDate: formatDateStandard(rows[i][3]),
          topic: String(rows[i][4] || ''),
          keyVerse: String(rows[i][5] || ''),
          bibleQuiz: String(rows[i][6] || ''),
          assignedGroup: String(rows[i][7] || ''),
          speakerLeader: String(rows[i][8] || ''),
          notes: String(rows[i][9] || '')
        });
      }
    }

    return { success: true, data: schedules };
  } catch (err) {
    return { success: false, error: err.message };
  }
}

function saveAllWeeklySchedulesBatch(schedulesList) {
  try {
    const ss = getSpreadsheet();
    let sheet = ss.getSheetByName(SHEET_NAMES.WEEKLY_SCHEDULES);
    if (!sheet) {
      setupManagerTeamsDatabase();
      sheet = ss.getSheetByName(SHEET_NAMES.WEEKLY_SCHEDULES);
    }

    if (schedulesList.length > 0) {
      const targetYear = schedulesList[0].year;
      const targetQuarter = schedulesList[0].quarter;

      const rows = sheet.getDataRange().getValues();
      for (let i = rows.length - 1; i >= 1; i--) {
        if (String(rows[i][1]) === targetYear && String(rows[i][2]) === targetQuarter) {
          sheet.deleteRow(i + 1);
        }
      }

      schedulesList.forEach((s, idx) => {
        const schId = s.scheduleId || `SCH_${targetYear}_${targetQuarter.replace(/\s+/g, '')}_${idx}`;
        sheet.appendRow([
          schId,
          s.year,
          s.quarter,
          s.meetingDate,
          s.topic,
          s.keyVerse || '',
          s.bibleQuiz || '',
          s.assignedGroup || '',
          s.speakerLeader || '',
          s.isSpecial ? 'Đặc biệt' : ''
        ]);
      });
    }

    return { success: true, error: "Đã lưu lịch chương trình cả Quý thành công!" };
  } catch (err) {
    return { success: false, error: err.message };
  }
}

function formatDateStandard(dateVal) {
  if (!dateVal) return '';
  let d = null;

  if (dateVal instanceof Date) {
    d = dateVal;
  } else {
    const str = String(dateVal).trim();
    if (str.match(/^\d{4}-\d{2}-\d{2}/)) {
      return str.substring(0, 10);
    }
    const dmYMatch = str.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})/);
    if (dmYMatch) {
      const day = dmYMatch[1].padStart(2, '0');
      const month = dmYMatch[2].padStart(2, '0');
      const year = dmYMatch[3];
      return `${year}-${month}-${day}`;
    }
    const parsed = Date.parse(str);
    if (!isNaN(parsed)) {
      d = new Date(parsed);
    }
  }

  if (d && !isNaN(d.getTime())) {
    const y = d.getFullYear();
    const m = String(d.getMonth() + 1).padStart(2, '0');
    const day = String(d.getDate()).padStart(2, '0');
    return `${y}-${m}-${day}`;
  }

  return String(dateVal).trim();
}

/**
 * Kết chuyển dữ liệu điểm danh năm cũ sang Sheet lưu trữ riêng (ví dụ: Attendance_Archive_2025)
 */
function archiveYearlyAttendance(targetYear) {
  try {
    if (!targetYear) return { success: false, error: 'Chưa chọn năm kết chuyển!' };
    targetYear = String(targetYear).trim();

    const ss = getSpreadsheet();
    const sheet = ss.getSheetByName(SHEET_NAMES.ATTENDANCE_RECORDS);
    if (!sheet || sheet.getLastRow() <= 1) {
      return { success: false, error: 'Bảng điểm danh chính hiện đang trống, không có dữ liệu để kết chuyển.' };
    }

    const rows = sheet.getDataRange().getValues();
    const headers = rows[0];

    const rowsToArchive = [];
    const rowsToKeepIndexes = [];

    for (let i = 1; i < rows.length; i++) {
      const dateVal = formatDateStandard(rows[i][1]);
      if (dateVal && dateVal.startsWith(targetYear)) {
        rowsToArchive.push(rows[i]);
      } else {
        rowsToKeepIndexes.push(i);
      }
    }

    if (rowsToArchive.length === 0) {
      return { success: false, error: `Không tìm thấy bản ghi điểm danh nào thuộc năm ${targetYear} trong bảng chính.` };
    }

    // Tạo hoặc mở Sheet lưu trữ riêng theo năm
    const archiveSheetName = `Attendance_Archive_${targetYear}`;
    let archiveSheet = ss.getSheetByName(archiveSheetName);
    if (!archiveSheet) {
      archiveSheet = ss.insertSheet(archiveSheetName);
      archiveSheet.appendRow(headers);
      archiveSheet.getRange(1, 1, 1, headers.length).setFontWeight('bold').setBackground('#0f172a').setFontColor('#ffffff');
    }

    // Ghi dữ liệu kết chuyển sang sheet lưu trữ
    const startRow = archiveSheet.getLastRow() + 1;
    archiveSheet.getRange(startRow, 1, rowsToArchive.length, headers.length).setValues(rowsToArchive);

    // Xóa các dòng đã kết chuyển khỏi sheet chính (xóa từ dưới lên để giữ nguyên chỉ số dòng)
    for (let i = rows.length - 1; i >= 1; i--) {
      const dateVal = formatDateStandard(rows[i][1]);
      if (dateVal && dateVal.startsWith(targetYear)) {
        sheet.deleteRow(i + 1);
      }
    }

    return {
      success: true,
      error: `🎉 Đã kết chuyển thành công ${rowsToArchive.length} bản ghi điểm danh năm ${targetYear} sang sheet "${archiveSheetName}"! Sheet chính hiện tại rất nhẹ và mượt.`
    };
  } catch (err) {
    return { success: false, error: err.message };
  }
}

/**
 * Xóa hàng loạt các lượt điểm danh đã chọn của 1 ngày nhóm
 */
function bulkDeleteWeeklyAttendanceRecords(memberIds, dateStr) {
  try {
    if (!memberIds || !Array.isArray(memberIds) || memberIds.length === 0) {
      return { success: false, error: 'Chưa chọn ban viên nào để xóa điểm danh!' };
    }
    if (!dateStr) return { success: false, error: 'Chưa chọn ngày điểm danh!' };

    const ss = getSpreadsheet();
    const sheet = ss.getSheetByName(SHEET_NAMES.ATTENDANCE_RECORDS);
    if (!sheet || sheet.getLastRow() <= 1) return { success: true, error: 'Không có dữ liệu điểm danh.' };

    const rows = sheet.getDataRange().getValues();
    let count = 0;

    for (let i = rows.length - 1; i >= 1; i--) {
      const recordDate = formatDateStandard(rows[i][1]);
      const mId = String(rows[i][2]);
      if (recordDate === dateStr && memberIds.includes(mId)) {
        sheet.deleteRow(i + 1);
        count++;
      }
    }

    return { success: true, error: `Đã xóa ${count} lượt điểm danh ngày ${dateStr} thành công!` };
  } catch (err) {
    return { success: false, error: err.message };
  }
}
