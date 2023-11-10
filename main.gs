function main() {
  const threads = GmailApp.search('label:ここにラベル名  is:Unread', 0, 100);

  threads.forEach((thread) => {
    thread.getMessages().forEach((message) => {
      if (!message.isUnread()) { return; }
       var date = `${message.getDate()}`;
       var froms =`${message.getFrom()}`;
       var subject=` ${message.getSubject()}`;
       var body =`${message.getPlainBody()}`;
       message.markRead();
       addToSpreadsheet(date, froms,  subject,body)
    });
  });
}

function  addToSpreadsheet(date, froms,subject,body) {
  var sheet = SpreadsheetApp.openById("ファイルのID").getSheetByName("件名シート名");
  var lastRow = sheet.getLastRow();
  var formattedDate = Utilities.formatDate(new Date(date), "JST", "yyyy/MM/dd HH:mm:ss");
  sheet.getRange(lastRow + 1, 1).setValue(formattedDate);
  sheet.getRange(lastRow + 1, 2).setValue(froms);
  sheet.getRange(lastRow + 1, 3).setValue(subject);
  sheet.getRange(lastRow + 1, 4).setValue(body);
}
