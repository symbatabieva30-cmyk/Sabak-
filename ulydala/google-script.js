function doPost(e) {
  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    var sh = SpreadsheetApp.getActiveSpreadsheet().getSheets()[0];
    var d = JSON.parse(e.postData.contents);
    if (sh.getLastRow() === 0) {
      sh.appendRow(['Аты', 'Жеткен қаласы', 'Дұрыс жауап', 'Барлық сұрақ', 'Тиын', 'Бітірді ме', 'Соңғы уақыты', 'ID']);
      sh.setFrozenRows(1);
      sh.getRange(1, 1, 1, 8).setFontWeight('bold');
    }
    var row = [d.name, d.city, d.correct, d.total, d.coins, d.finished ? 'Иә ✅' : 'Жоқ', new Date(), d.id];
    var last = sh.getLastRow();
    var ids = last > 1 ? sh.getRange(2, 8, last - 1, 1).getValues().map(function (r) { return r[0]; }) : [];
    var i = ids.indexOf(d.id);
    if (i >= 0) sh.getRange(i + 2, 1, 1, 8).setValues([row]);
    else sh.appendRow(row);
    return ContentService.createTextOutput('ok');
  } finally {
    lock.releaseLock();
  }
}
