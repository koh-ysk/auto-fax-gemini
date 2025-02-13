/**
 * スプレッドシートが開かれたときに実行される関数
 * カスタムメニューから関数を実行できるようにする
 * => 権限で関係でエラーが発生しまくるという事象が発生
 */
function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('カスタムメニュー')
      .addItem('FAX取込実行', 'handleMails')
      .addItem('AI分析実行', 'handleRows')
      .addToUi(); // スプレッドシートのUIにメニューを追加
}
