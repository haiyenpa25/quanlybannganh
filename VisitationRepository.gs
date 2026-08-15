/**
 * VisitationRepository.gs - OOP Repository for Pastoral Care & Visitation Domain Entity
 * Extends BaseRepository (0_BaseRepository.gs)
 */
class VisitationRepository extends BaseRepository {
  constructor() {
    super("Visitations");
  }

  getVisitationLogs() {
    const rawLogs = this.getAllAsObjects();

    return rawLogs.map(item => ({
      logId: String(item["Mã Thăm Viếng"] || ""),
      memberId: String(item["Mã Thành Viên"] || ""),
      memberName: String(item["Họ và Tên"] || ""),
      visitorName: String(item["Người Thăm"] || ""),
      visitDate: String(item["Ngày Thăm"] || ""),
      visitType: String(item["Hình Thức"] || "Tại nhà / Trực tiếp"),
      prayerRequests: String(item["Nhu Cầu Cầu Nguyện"] || ""),
      resultNotes: String(item["Kết Quả & Ghi Chú"] || ""),
      _rowIndex: item._rowIndex
    }));
  }

  saveVisitationLog(logData) {
    const sheet = this.getSheet();
    const data = sheet.getDataRange().getValues();
    const headers = data[0] || ["Mã Thăm Viếng", "Mã Thành Viên", "Họ và Tên", "Người Thăm", "Ngày Thăm", "Hình Thức", "Nhu Cầu Cầu Nguyện", "Kết Quả & Ghi Chú"];

    let targetRowIndex = -1;
    let logId = logData.logId;

    if (logId) {
      // Find existing log row to update
      for (let i = 1; i < data.length; i++) {
        if (String(data[i][0]).trim() === String(logId).trim()) {
          targetRowIndex = i + 1;
          break;
        }
      }
    }

    if (!logId) {
      // Auto generate new log ID: VIS-YYYYMMDD-XXX
      const todayStr = new Date().toISOString().split('T')[0].replace(/-/g, '');
      const randNum = Math.floor(100 + Math.random() * 900);
      logId = `VIS-${todayStr}-${randNum}`;
    }

    const rowValues = [
      logId,
      logData.memberId || "",
      logData.memberName || "",
      logData.visitorName || "",
      logData.visitDate || "",
      logData.visitType || "Tại nhà / Trực tiếp",
      logData.prayerRequests || "",
      logData.resultNotes || ""
    ];

    if (targetRowIndex > 1) {
      this.updateRow(targetRowIndex, rowValues);
    } else {
      this.appendRow(rowValues);
    }

    // Auto update member status if requested
    if (logData.updateAttendance && logData.memberId) {
      this.updateMemberStatusAfterVisit(logData.memberId);
    }

    return { success: true, logId: logId };
  }

  deleteVisitationLog(logId) {
    if (!logId) return { success: false, error: "Mã lượt thăm không hợp lệ." };

    const sheet = this.getSheet();
    const data = sheet.getDataRange().getValues();

    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]).trim() === String(logId).trim()) {
        this.deleteRow(i + 1);
        return { success: true, logId: logId };
      }
    }

    return { success: false, error: "Không tìm thấy lượt thăm cần xóa." };
  }

  updateMemberStatusAfterVisit(memberId) {
    try {
      const ss = this.getSpreadsheet();
      const memSheet = ss.getSheetByName("Members");
      if (!memSheet) return;

      const data = memSheet.getDataRange().getValues();
      for (let i = 1; i < data.length; i++) {
        if (String(data[i][0]).trim() === String(memberId).trim()) {
          // Col index 14 is "Trạng Thái"
          memSheet.getRange(i + 1, 14).setValue("Đang thăm viếng");
          break;
        }
      }
    } catch (e) {
      console.warn("Failed to auto-update member status:", e.message);
    }
  }
}
