// GoogleDriveのPDF保存先フォルダ
const FOLDER_ID = PropertiesService.getScriptProperties().getProperty("FOLDER_ID")
// 監視対象のメールアドレス
const EMAIL_ADDRESS = PropertiesService.getScriptProperties().getProperty("EMAIL_ADDRESS")

// FAXの分析結果を出力するスプレッドシート
const faxMainSheet = SpreadsheetApp.getActive().getSheetByName('main');
// FAXのメール実行ログを管理するシート
const faxMailLogSheet = SpreadsheetApp.getActive().getSheetByName('mailExecuteLog');
// FAXのファイルを管理するシート(ファイル名で取得済みのものを管理)
const faxFileNamesLogSheet = SpreadsheetApp.getActive().getSheetByName('faxFileNamesLog');

/**
 * メールを処理していく関数
 */
function handleMails() {
  // FAXを格納する共有フォルダのIDを指定
  const folder = DriveApp.getFolderById(FOLDER_ID);
  const now = new Date();
  let errorLog = "";

  try {
    // 重複がないように、すでに取得済みのFAXのファイル名を取得
    const existingFaxFiles = getExistingFaxFiles();

    // メール実行ログの最後の行を実行時間のログとして使用
    const lastRow = faxMailLogSheet.getLastRow();
    const lastRunRange = faxMailLogSheet.getRange("A" + lastRow);
    const lastRunDate = new Date(lastRunRange.getValue());

    // 前回の実行時間から現在までの間に受信したメールを検索
    const searchQuery = `to:${EMAIL_ADDRESS} after:` + Math.floor(lastRunDate.getTime() / 1000);
    const threads = GmailApp.search(searchQuery);

    console.log(searchQuery);

    // 取得したメールを送信日時で昇順ソート
    threads.sort(function(a, b) {
      const lastMessageDateA = a.getMessages()[a.getMessageCount() - 1].getDate();
      const lastMessageDateB = b.getMessages()[b.getMessageCount() - 1].getDate();
      return lastMessageDateA.getTime() - lastMessageDateB.getTime();
    });

    // メールを1つずつ分析
    for (let i = 0; i < threads.length; i++) {
      const messages = threads[i].getMessages();
      for (let j = 0; j < messages.length; j++) {
        try {
          // mainのシートに書き込む
          handleAttachments(messages[j], folder, faxMainSheet, now, existingFaxFiles);
        } catch (e) {
          console.error(e.toString());
        }
      }
    }
  } catch (e) {
    errorLog = e.toString();
    console.error(errorLog);
  }

  // mailログに書き込む
  faxMailLogSheet.appendRow([now, errorLog]);
}

/**
 * メールの添付ファイルを処理する関数
 * @param メールオブジェクト
 * @param Google Drive Folder
 * @param スプレッドシートのシートオブジェクト
 * @param Dateオブジェクト
 * @param array
 */
function handleAttachments(message, folder, logSheet, now, existingAttachments) {
  const attachments = message.getAttachments();

  for (let k = 0; k < attachments.length; k++) {
    try {
      if (attachments[k].getContentType() === "application/pdf"
        && !existingAttachments.includes(attachments[k].getName()))
      {
        const file = folder.createFile(attachments[k]);
        writeInfo(logSheet, message, file, now);

        // faxログに書き込む
        faxFileNamesLogSheet.appendRow([file.getName()]);
      }
    } catch (e) {
      console.error(e.toString())
    }
  }
}

/**
 * スプレッドシートに情報を書き込む
 * @param スプレッドシートのシートオブジェクト
 * @param メールオブジェクト
 * @param 添付ファイル
 * @param Dateオブジェクト
 */
function writeInfo(logSheet, message, file, now) {
  // ログをスプレッドシートに書き込む
  logSheet.appendRow([
    message.getDate(), // 受信日時
    "", // AI分析内容
    "", //転送記録
    file.getUrl(), // Google Driveのパス
    "", //AI分析送信者
    "", //AI分析受信者
    file.getName(), // 添付ファイル名
    now, // 実行時間
    message.getFrom(), // 送信者
  ]);
}

/**
 * すでに分析済みのFaxファイル名を取得する関数
 * @return { array }: ファイル名一覧
 */
function getExistingFaxFiles() {

  //最後の行を取得
  const lastRow = faxFileNamesLogSheet.getLastRow();

  //1列目の値を配列に格納
  let values = [];
  for (let i = 1; i <= lastRow; i++) {
    values.push(faxFileNamesLogSheet.getRange(i, 1).getValue());
  }
  return values;
}
