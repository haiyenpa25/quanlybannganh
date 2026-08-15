/**
 * 0_BaseRepository.gs - Abstract OOP Base Repository for Google Sheets Data Access
 * Prefix '0_' ensures Google Apps Script loads this file BEFORE any subclasses.
 */
class BaseRepository {
  constructor(sheetName) {
    this.sheetName = sheetName;
  }

  getSpreadsheet() {
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

  getSheet() {
    const ss = this.getSpreadsheet();
    let sheet = ss.getSheetByName(this.sheetName);
    if (!sheet) {
      sheet = ss.insertSheet(this.sheetName);
    }
    return sheet;
  }

  getAllRows() {
    const sheet = this.getSheet();
    const data = sheet.getDataRange().getValues();
    if (data.length <= 1) return { headers: data[0] || [], rows: [] };
    
    const headers = data[0].map(h => String(h).trim());
    const rows = data.slice(1);
    return { headers, rows };
  }

  getAllAsObjects() {
    const { headers, rows } = this.getAllRows();
    if (headers.length === 0) return [];

    return rows.map((row, index) => {
      const obj = { _rowIndex: index + 2 };
      headers.forEach((h, colIdx) => {
        obj[h] = row[colIdx] !== undefined ? row[colIdx] : '';
      });
      return obj;
    });
  }

  appendRow(rowArray) {
    const sheet = this.getSheet();
    sheet.appendRow(rowArray);
  }

  updateRow(rowIndex, rowArray) {
    const sheet = this.getSheet();
    const range = sheet.getRange(rowIndex, 1, 1, rowArray.length);
    range.setValues([rowArray]);
  }

  deleteRow(rowIndex) {
    const sheet = this.getSheet();
    sheet.deleteRow(rowIndex);
  }
}
