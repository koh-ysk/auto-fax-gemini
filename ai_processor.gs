// WEBHOOK_URL(Google Chat通知用)
const WEBHOOK_URL = PropertiesService.getScriptProperties().getProperty("WEBHOOK_URL")

/**
 * スプレッドシートの情報を読み込んで、Geminiに投げていく関数
 */
function handleRows() {
  const lastRow = faxMainSheet.getLastRow();

  // 1行ずつ分析
  for (let i = 1; i <= lastRow; i++) {
    const context = faxMainSheet.getRange("B" + i).getValue();
    // 内容がすでに書き込まれていたらスキップ
    if (context) {
      console.info(`${i}行目はすでに書き込まれています`);
    } else {
      try {
        callGemini(i);
      } catch (e) {
        console.error(e.toString())
      }
    }
  }
}

/**
 * スプレッドシートの特定行を読み込んで、Geminiに分析依頼
 * @params {integer} row - 行番号
 */
function callGemini(row) {
  const prompt = getPrompt("EmailAnalysis")
  const fileUrl = faxMainSheet.getRange("D" + row).getValue();
  const fileId = extractFileID(fileUrl);
  const result = Gemini_common_functions.gemini(prompt, fileId);

  insertAnalyzedData(row, result, fileUrl);
}

/**
 * Gemini分析結果をシートに書き込む
 * @param {integer} row - 行番号
 * @param {string} result - Geminiライブラリからの返り値
 * @param {string} fileUrl - Google Drive上のファイル
 */
function insertAnalyzedData(row, result, fileUrl) {
  const object = extractAndParseJsonFromMarkdown(result);
  console.info(object);

  if (object == null) {
    console.error(`${row}行目のGemini分析が失敗しました`);
  } else {
    faxMainSheet.getRange("B" + row).setFormula('=HYPERLINK("' + fileUrl + '", "' + object.content + '")');
    faxMainSheet.getRange("E" + row).setValue(object.sender);
    faxMainSheet.getRange("F" + row).setValue(object.receiver);
    faxMainSheet.getRange("J" + row).setValue(object.is_important);

    // 重要なものが来た際は、通知を飛ばす
    if (object.is_important === true) {
      sendChatMessageIfTrue(object, row, fileUrl);
    }
  }

}

/**
 * Google Drive URLからfileIDを取得
 * @param {string} url - Google Drive上のファイル
 */
function extractFileID(url) {
  // 正規表現パターン
  const pattern = /\/d\/([a-zA-Z0-9_-]+)\//;
  const matches = url.match(pattern);

  // マッチした場合、fileIDを取り出す
  if (matches && matches.length > 1) {
    return matches[1];
  } else {
    return null;
  }
}

/**
 * マークダウン表記のテキストからJSON文字列を抽出し、パースする
 * @param {string} markdownText - マークダウン表記を含むテキスト
 * @return {Object|null} パースされたJSONオブジェクト、または抽出・パースに失敗した場合はnull
 */
function extractAndParseJsonFromMarkdown(markdownText) {
  // マークダウンコードブロック内のJSON文字列を抽出するための正規表現
  const regex = /```(?:json)?\s*([\s\S]*?)\s*```/;
  const match = markdownText.match(regex);

  if (match && match[1]) {
    // 抽出された文字列をJSONオブジェクトにパースする
    try {
      const jsonObject = JSON.parse(match[1]);
      return jsonObject;
    } catch (e) {
      console.error('JSONのパースに失敗しました: ' + e.toString());
    }
  } else {
    console.error('マークダウンからJSON文字列を抽出できませんでした。');
  }
  return null;
}

/**
 * スプレッドシートからプロンプトを取得する関数
 */
function getPrompt(promptName) {
  const faxFileNamesLogSheet = SpreadsheetApp.getActive().getSheetByName('geminiPrompt');
  const data = faxFileNamesLogSheet.getDataRange().getValues();
  let promptTemplate = '';

  for (let i = 1; i < data.length; i++) { // ヘッダーをスキップするために1から開始
    if (data[i][0] === promptName) {
      promptTemplate = data[i][1];
      break;
    }
  }

  if (!promptTemplate) {
    throw new Error('プロンプトが見つかりません');
  }

  return promptTemplate;
}


/**
 * WebHookに通知を飛ばす関数
 */
function sendChatMessageIfTrue(object, row, fileUrl) {
  // 曜日の日本語表記の配列
  const daysOfWeek = ["日", "月", "火", "水", "木", "金", "土"];

  try {
    const faxMainSheet = SpreadsheetApp.getActive().getSheetByName('main');


    // 受信日時を取得・フォーマット (時刻情報も含める)
    const receivedDateTimeCell = faxMainSheet.getRange("A" + row);
    const receivedDateTimeValue = receivedDateTimeCell.getValue();
    const receivedDateTime = new Date(receivedDateTimeValue);
    const formattedDateTime = Utilities.formatDate(receivedDateTime, "JST", "yyyy年M月d日 H時mm分");
    const dayOfWeekIndex = receivedDateTime.getDay();
    const dayOfWeekJp = daysOfWeek[dayOfWeekIndex];

    // メッセージのフォーマットを作成 (object.content全体をハイパーリンクに変換)
    const contentWithHyperlink = `<${fileUrl}|${object.content}>`;

    // メッセージのフォーマットを作成
    const message =
      `受信日時：${formattedDateTime} (${dayOfWeekJp})\n` +
      `内容：${contentWithHyperlink}\n` +
      `送信元：${object.sender}`;

    const payload = JSON.stringify({ text: message });
    const options = { method: "POST", contentType: "application/json", payload: payload };
    UrlFetchApp.fetch(WEBHOOK_URL, options);
    console.log("メッセージが送信されました");

  } catch (error) {
    console.error("sendChatMessageIfTrue 関数でエラーが発生しました: " + error);
    console.error("スタックトレース: " + error.stack);
  }
}
