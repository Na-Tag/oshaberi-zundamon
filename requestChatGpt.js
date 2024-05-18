const SPREAD_SHEET_ID = '1CURaAYrxoW-HV8pvaSNfIX30DqTE4u2_v7rwQYYlY44';

function getRows(){
  const rows = SpreadsheetApp.openById(SPREAD_SHEET_ID).getSheetByName("プロンプト").getDataRange().getValues().slice(2);
  return rows;
} 

const systemContent = getRows()[1][1];

// JSDocの記法の一部がスプレッドシートのオートコンプリートに反映される
// 参考：https://qiita.com/suzuki_sh/items/e44c2526d63fe9fa89ab
/**
 * GPTを呼び出すカスタム関数
 *
 * @param {string} input GPTへの入力文字列
 * @param {model} model モデル名。デフォルトは"gpt-3.5-turbo"
 * @param {boolean} useCache キャッシュを使用するかどうか。長い入力を使う時はキャッシュを無効化する必要がある。デフォルトは"true"
 * @return GPTで生成した結果
 * @customfunction
 */
function GPT(input, model, useCache) {
  const scriptProperties = PropertiesService.getScriptProperties();
  if (model === void 0) { model = "gpt-3.5-turbo"; }
  if (useCache === void 0) { useCache = true; }
  // キャッシュから取得
    const cache = CacheService.getScriptCache();
  if (useCache) {
    try {
      const cachedResult = cache.get(input);
      if (cachedResult) {
        SpreadsheetApp.getActiveSpreadsheet().toast("cahched");
        return cachedResult;
      }
    }
    catch (e) {
      if (e.message.includes("Argument too large: key")) {
          throw new Error("キャッシュの取得に失敗しました。入力を短くするか、キャッシュを無効にしてください。");
      }
      else {
          throw e;
      }
    }
  }
  // OpenAI Chat Completion APIで生成
  const URL = "https://api.openai.com/v1/chat/completions";
  const headers = {
      "Content-Type": "application/json",
      "Authorization": "Bearer ".concat(scriptProperties.getProperty("OPENAI_API_KEY"))
  };
  const body = {
      "model": model,
      "messages": [
        { "role": "system", "content": systemContent},
        { "role": "user", "content": input}
      ]
  };
  try{
    const response = UrlFetchApp.fetch(URL, {
      "method": "post",
      "headers": headers,
      "payload": JSON.stringify(body),
      "muteHttpExceptions" : true
    });
    console.log(response.getContentText());
    const result = JSON.parse(response.getContentText()).choices[0].message.content;
    // キャッシュに保存
    if (useCache) {
        cache.put(input, result, 21600);
    }
    return result;
  }catch(e){
    console.error("Error : " + e);
  }
    
}

function requestChatGptMain(){
  const sheet = SpreadsheetApp.openById(SPREAD_SHEET_ID).getSheetByName("プロンプト");
  const range = sheet.getDataRange();
  const rows = range.getValues();
  if(deployManager(rows)){
    const input = rows[3][0];
    const answer = GPT(input,void 0, true);
    rows[3][2] = answer;
    range.setValues(rows);
    setLog([input,answer,systemContent])
  }else{
    SpreadsheetApp.getActiveSpreadsheet().toast("実行不可状態。");
  }
}

function deployManager(rows){
  const value = rows[0][0];
  if(value == "スクリプトの実行可能状況：可能")
    return 1;
  else  
    return 0;
}

function deploySwitch(){
  const range = SpreadsheetApp.openById(SPREAD_SHEET_ID).getSheetByName("プロンプト").getDataRange();
  const rows = range.getValues();
  if(rows[0][0] == "スクリプトの実行可能状況：不可")
    rows[0][0] = "スクリプトの実行可能状況：可能";
  else
    rows[0][0] = "スクリプトの実行可能状況：不可";
  range.setValues(rows);
}

