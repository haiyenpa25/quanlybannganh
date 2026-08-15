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

  saveSingleSchedule(schData) {
    const ss = this.getSpreadsheet();
    let sheet = ss.getSheetByName("Weekly_Schedules");
    if (!sheet) sheet = ss.insertSheet("Weekly_Schedules");

    const rows = sheet.getDataRange().getValues();
    let schId = schData.scheduleId;
    if (!schId) {
      schId = 'SCH_' + (schData.date ? schData.date.replace(/-/g, '') : Date.now());
    }

    const year = schData.date ? schData.date.substring(0, 4) : '2026';
    const month = schData.date ? parseInt(schData.date.substring(5, 7)) : 1;
    let quarter = 'Quý I';
    if (month >= 4 && month <= 6) quarter = 'Quý II';
    else if (month >= 7 && month <= 9) quarter = 'Quý III';
    else if (month >= 10) quarter = 'Quý IV';

    const rowData = [
      schId, year, quarter, schData.date, schData.topic,
      schData.keyVerse, schData.quizGroup, schData.speaker, schData.notes || ''
    ];

    let foundIndex = -1;
    for (let i = 1; i < rows.length; i++) {
      if (rows[i][0] === schId || (rows[i][3] && String(rows[i][3]).startsWith(schData.date))) {
        foundIndex = i + 1;
        break;
      }
    }

    if (foundIndex > 0) {
      sheet.getRange(foundIndex, 1, 1, rowData.length).setValues([rowData]);
    } else {
      sheet.appendRow(rowData);
    }

    return { success: true, scheduleId: schId };
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
