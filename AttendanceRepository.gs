/**
 * AttendanceRepository.gs - OOP Repository for Weekly Attendance Domain Entity
 */
class AttendanceRepository extends BaseRepository {
  constructor() {
    super("Attendance_Records");
  }

  getWeeklyRecordsForDate(dateStr) {
    const allRecords = this.getAllAsObjects();
    return allRecords.filter(r => String(r["Ngày Nhóm"]).trim() === String(dateStr).trim());
  }

  archiveYearRecords(targetYear) {
    const ss = this.getSpreadsheet();
    const mainSheet = this.getSheet();
    const data = mainSheet.getDataRange().getValues();
    if (data.length <= 1) return 0;

    const headers = data[0];
    const archiveSheetName = `Attendance_Archive_${targetYear}`;
    let archiveSheet = ss.getSheetByName(archiveSheetName);
    if (!archiveSheet) {
      archiveSheet = ss.insertSheet(archiveSheetName);
      archiveSheet.appendRow(headers);
    }

    const rowsToKeep = [headers];
    const rowsToMove = [];

    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const rowDate = String(row[1] || "");
      if (rowDate.startsWith(String(targetYear))) {
        rowsToMove.push(row);
      } else {
        rowsToKeep.push(row);
      }
    }

    if (rowsToMove.length > 0) {
      archiveSheet.getRange(archiveSheet.getLastRow() + 1, 1, rowsToMove.length, headers.length).setValues(rowsToMove);
      mainSheet.clearContents();
      mainSheet.getRange(1, 1, rowsToKeep.length, headers.length).setValues(rowsToKeep);
    }

    return rowsToMove.length;
  }
}
