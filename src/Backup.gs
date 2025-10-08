function createSpreadsheetBackup() {
  const ss = getSpreadsheet();
  const name = ss.getName() + ' Backup ' + Utilities.formatDate(new Date(), CONFIG.timezone, 'yyyyMMdd_HHmm');
  const file = DriveApp.getFileById(ss.getId());
  file.makeCopy(name);
  recordAudit('createSpreadsheetBackup', 'バックアップ作成: ' + name);
}