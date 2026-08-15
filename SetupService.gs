/**
 * SetupService.gs - 1-Click Automated Spreadsheet Database Setup & Deployment Tool
 * Enterprise Service for Reusing ManagerTeams Across Multiple Churches & Departments
 */
class SetupService {
  static initializeNewDatabase(spreadsheetId, churchName, deptName, motto = "") {
    try {
      let ss = null;

      if (spreadsheetId && spreadsheetId.trim() !== "") {
        ss = SpreadsheetApp.openById(spreadsheetId.trim());
      } else {
        ss = SpreadsheetApp.getActiveSpreadsheet();
      }

      if (!ss) {
        return { success: false, error: "Không tìm thấy Google Sheet với ID đã cung cấp." };
      }

      churchName = churchName || "HỘI THÁNH TIN LÀNH THỦ ĐỨC";
      deptName = deptName || "BAN THẮNG TRÁNG";
      motto = motto || "SỐNG VỚI CÂU GỐC & CẦU NGUYỆN LÊN CHÚA";

      // 1. Setup Members Sheet
      this.ensureSheetWithHeaders(ss, "Members", [
        "Mã Thành Viên", "Họ và Tên", "Chức Danh", "Nhiệm Kỳ", "Ngày Sinh",
        "Số Điện Thoại", "Địa Chỉ", "Mã Tổ", "Tên Tổ", "Ngày Báp Têm",
        "Tên Vợ Chồng", "Ngày Sinh Hoạt", "Lần Nhóm Cuối", "Trạng Thái", "Ghi Chú"
      ]);

      // 2. Setup Groups Sheet
      const groupSheet = this.ensureSheetWithHeaders(ss, "Groups", [
        "Mã Tổ", "Năm", "Tên Tổ", "Tổ Trưởng", "Số ĐT Tổ Trưởng", "Ghi Chú"
      ]);
      if (groupSheet.getLastRow() <= 1) {
        groupSheet.appendRow(["G01", "2026", "Tổ 1 - Giô-suê", "Nguyễn Văn Bình", "0901234567", "Tổ nhóm tối Thứ 5"]);
        groupSheet.appendRow(["G02", "2026", "Tổ 2 - Đa-ni-ên", "Trần Thị Yến", "0909876543", "Tổ nhóm tối Thứ 6"]);
      }

      // 3. Setup Weekly_Schedules Sheet
      this.ensureSheetWithHeaders(ss, "Weekly_Schedules", [
        "Mã Buổi Nhóm", "Năm", "Quý", "Ngày Nhóm", "Đề Tài", "Câu Gốc",
        "Đố Kinh Thánh", "Tổ Phụ Trách", "Phụ Trách (Diễn Giả)", "Ghi Chú"
      ]);

      // 4. Setup Attendance_Records Sheet
      this.ensureSheetWithHeaders(ss, "Attendance_Records", [
        "Mã Điểm Danh", "Ngày Nhóm", "Mã Thành Viên", "Họ và Tên", "Trạng Thái",
        "Thuộc Câu Gốc", "Số Câu Đố KT", "Ghi Chú", "Thời Gian Ghi"
      ]);

      // 5. Setup Visitations Sheet
      this.ensureSheetWithHeaders(ss, "Visitations", [
        "Mã Thăm Viếng", "Mã Thành Viên", "Họ và Tên", "Người Thăm", "Ngày Thăm",
        "Hình Thức", "Nhu Cầu Cầu Nguyện", "Kết Quả & Ghi Chú"
      ]);

      // 6. Setup Quarterly_Themes Sheet
      const themeSheet = this.ensureSheetWithHeaders(ss, "Quarterly_Themes", [
        "Mã Chủ Đề", "Năm", "Quý", "Tên Chủ Đề", "Câu Gốc", "Nội Dung Câu Gốc", "Bài Hát Chủ Đề", "Ghi Chú"
      ]);
      if (themeSheet.getLastRow() <= 1) {
        themeSheet.appendRow([
          "TH_2026_Q1", "2026", "Quý I", "TƯỚI NƯỚC CÂY SỰ SỐNG",
          "Thi-thiên 1:3", "Người sẽ như cây trồng gần dòng nước, sinh bông trái theo thì tiết.",
          "Bài Hát Ban Ngành 2026", "Chủ đề sinh hoạt Quý I"
        ]);
      }

      // 7. Setup Accounts Sheet
      const accSheet = this.ensureSheetWithHeaders(ss, "Accounts", [
        "Tên Đăng Nhập", "Mật Khẩu", "Họ và Tên", "Vai Trò", "Trạng Thái"
      ]);
      if (accSheet.getLastRow() <= 1) {
        accSheet.appendRow(["binh", "123456", "Nguyễn Văn Bình", "Trưởng ban", "Hoạt động"]);
        accSheet.appendRow(["yen", "123456", "Trần Thị Yến", "Phó ban", "Hoạt động"]);
        accSheet.appendRow(["duen", "123456", "Lê Thị Duyên", "Thư ký", "Hoạt động"]);
        accSheet.appendRow(["tuan", "123456", "Phạm Văn Tuấn", "Thủ quỹ", "Hoạt động"]);
        accSheet.appendRow(["nhung", "123456", "Hoàng Thị Nhung", "Nghị viên", "Hoạt động"]);
        accSheet.appendRow(["hien", "123456", "Vũ Thị Hiền", "Ban viên", "Hoạt động"]);
      }

      // 8. Setup Config Sheet
      const cfgSheet = this.ensureSheetWithHeaders(ss, "Config", ["Khóa Cấu Hình", "Giá Trị"]);
      const cfgMap = {
        "churchName": churchName,
        "departmentName": deptName,
        "departmentMotto": motto,
        "meetingDay": "Chúa Nhật",
        "meetingTime": "16h00 chiều",
        "footerNotice": "Xin quý Anh/Chị thêm lời cầu nguyện và tham dự đầy đủ."
      };
      
      cfgSheet.clearContents();
      cfgSheet.appendRow(["Khóa Cấu Hình", "Giá Trị"]);
      Object.keys(cfgMap).forEach(key => {
        cfgSheet.appendRow([key, cfgMap[key]]);
      });

      // Save Active Spreadsheet ID into ScriptProperties if specified
      if (spreadsheetId && spreadsheetId.trim() !== "") {
        PropertiesService.getScriptProperties().setProperty('ACTIVE_SPREADSHEET_ID', spreadsheetId.trim());
      }

      return {
        success: true,
        spreadsheetId: ss.getId(),
        spreadsheetUrl: ss.getUrl(),
        error: `Đã khởi tạo thành công CSDL cho [${churchName} - ${deptName}]!`
      };

    } catch (err) {
      return { success: false, error: err.message };
    }
  }

  static ensureSheetWithHeaders(ss, sheetName, headers) {
    let sheet = ss.getSheetByName(sheetName);
    if (!sheet) {
      sheet = ss.insertSheet(sheetName);
    }
    if (sheet.getLastRow() === 0) {
      sheet.appendRow(headers);
      const headerRange = sheet.getRange(1, 1, 1, headers.length);
      headerRange.setBackground("#1e293b").setFontColor("#ffffff").setFontWeight("bold");
    }
    return sheet;
  }
}
