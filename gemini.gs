const PROPERTIES = PropertiesService.getScriptProperties();
const API_ENDPOINT = PROPERTIES.getProperty("API_ENDPOINT");
const PROJECT_ID = PROPERTIES.getProperty("PROJECT_ID");
const MODEL_ID = PROPERTIES.getProperty("MODEL_ID");
const LOCATION_ID = PROPERTIES.getProperty("LOCATION_ID");

const DEFAULT_GEN_CONFIG = {
    maxOutputTokens: 8192,
    temperature: 1,
    topP: 0.95,
    topK: 32
};

/**
 * Gemini APIを利用してテキストやマルチモーダルデータを処理する関数。
 *
 * @param {string} prompt - Gemini APIに送信するプロンプト。
 * @param {string|Object} [file] - ファイルID（Google Drive）またはオブジェクト（GCSのファイル情報）。
 * @param {Object} [config={}] - 生成設定（maxOutputTokens, temperature, topP, topK）。
 * @param {string} [model=MODEL_ID] - 使用するモデルのID。
 * @param {string} [projectId] - デフォルト以外のプロジェクトを指定する場合。
 * @returns {string} - Gemini APIのレスポンス（テキスト形式）。
 */
function gemini(prompt, file, config = {}, model = MODEL_ID, projectId = PROJECT_ID) {
    const genConfig = { ...DEFAULT_GEN_CONFIG, ...config }; // デフォルト値をスプレッド構文で設定
    const geminiEndpoint = `https://${API_ENDPOINT}/v1/projects/${projectId}/locations/${LOCATION_ID}/publishers/google/models/${model}:streamGenerateContent`;

    console.log("Gemini API Endpoint:", geminiEndpoint);

    const parts = [{ text: prompt }];
    if (file) {
        console.log("マルチモーダル入力を処理します");
        if (typeof file === "string") { // Drive の場合は fileId（文字列）、GCS はオブジェクト
            const driveFile = DriveApp.getFileById(file);
            const blob = driveFile.getBlob();
            const base64Data = Utilities.base64Encode(blob.getBytes());
            const mimeType = driveFile.getMimeType() || "image/jpeg";

            console.log(`DriveファイルのMIMEタイプ: ${mimeType}`);

            parts.unshift({ // 配列の先頭に追加
                inlineData: {
                    mimeType: mimeType,
                    data: base64Data
                }
            });
        }
    } else {
        console.log("テキスト入力のみを処理します");
    }

    const request = {
        contents: [{ role: "user", parts }],
        generation_config: genConfig
    };

    const headers = {
        "Authorization": "Bearer " + ScriptApp.getOAuthToken(),
        "Content-Type": "application/json"
    };

    const options = {
        method: "POST",
        headers: headers,
        payload: JSON.stringify(request),
        muteHttpExceptions: true,
    };

    try {
        const response = UrlFetchApp.fetch(geminiEndpoint, options);
        const responseData = JSON.parse(response.getContentText());

        // レスポンスが期待通りでない場合のエラーハンドリング
        if (!responseData || !responseData.length || !responseData[0].candidates || !responseData[0].candidates[0].content || !responseData[0].candidates[0].content.parts || !responseData[0].candidates[0].content.parts.length) {
            console.error("Unexpected API Response:", responseData);
            return "Error: Unexpected API Response";
        }

        return responseData.map(data => data.candidates[0].content.parts.map(part => part.text).join("")).join(""); // 複数の候補とパートに対応
    } catch (error) {
        console.error("Failed to fetch from Gemini API", error);
        return "Error fetching response from Gemini API";
    }
}

/**
 * Gemini APIのテスト関数（テキストのみ）。
 */
function testGemini() {
    const prompt = "日本の首都はどこですか？";
    const response = gemini(prompt);
    console.log("Test Response:", response);
}

/**
 * Google Drive のファイルを含めたテスト
 */
function testGeminiWithDrive() {
    const fileId = "1sam6ple123SampleSampleSa4mp4leSAMPLE"; // 実際のGoogle DriveのファイルIDを使用
    const prompt = "この写真に写っている動物は何ですか？";

    const response = gemini(prompt, fileId);
    console.log("Test Response with Drive File:", response);
}
