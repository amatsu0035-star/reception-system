// ====== 設定 ======
const SHEET_NAME = '来場予定表';
const VISION_API_KEY = 'AIzaSyChJZYyzo6pLRjCcnE-w4IUkaWDarVnEu0'; // ここにGoogle Cloud Vision APIのAPIキーを貼り付けてください

// スプレッドシートの列インデックス設定（A列=0, B列=1, C列=2...）
// ※実際のシートの列配置に合わせて適宜変更してください。
const COL_KEY = 0;          // キー番号
const COL_INVITER = 1;      // 招待元企業名
const COL_CUSTOMER = 2;     // お客様名
const COL_COMPANY = 3;      // お客様会社名
const COL_NOTE = 4;         // 特記事項
const COL_STATUS = 5;       // 来場状況（「来場済み」等を記録する列）
const COL_ARRIVED_AT = 6;   // 来場日時（「yyyy/mm/dd(aaa)_hh:mm」形式を記録する列）
const COL_GUEST_COUNT = 7;  // 来場人数を記録する列（H列）
const COL_CHILD_COUNT = 8;  // お子様の人数を記録する列（I列）
// ==================

function doGet() {
  return HtmlService.createHtmlOutputFromFile('index')
      .setTitle('来賓受付システム')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
      // iPadでのピンチズーム・ダブルタップズームを無効化して誤操作を防ぐ
      .addMetaTag('viewport', 'width=device-width, initial-scale=1.0, maximum-scale=1.0, user-scalable=no');
}

/**
 * 現在日時を「yyyy/mm/dd(aaa)_hh:mm」の形式で取得する関数
 */
function getFormattedDate() {
  const date = new Date();
  const yyyy = date.getFullYear();
  const mm = String(date.getMonth() + 1).padStart(2, '0');
  const dd = String(date.getDate()).padStart(2, '0');
  const days = ['日', '月', '火', '水', '木', '金', '土'];
  const aaa = days[date.getDay()];
  const hh = String(date.getHours()).padStart(2, '0');
  const min = String(date.getMinutes()).padStart(2, '0');
  
  return `${yyyy}/${mm}/${dd}(${aaa})_${hh}:${min}`;
}

/**
 * フロントエンドからの受付処理を受け取る関数（排他制御付き）
 * @param {string} keyNumber 検索キー
 * @param {number} guestCount 来場人数
 * @param {number} childCount お子様人数
 * @return {object} 照合結果とゲスト情報を含むオブジェクト
 */
function processReception(keyNumber, guestCount, childCount) {
  // 複数端末からの同時実行を防ぐためロックを取得
  const lock = LockService.getScriptLock();
  // 15秒間ロック取得を試みる。取得できない場合はエラーを返す
  if (!lock.tryLock(15000)) { 
    throw new Error('現在、他の端末が処理中です。少し待ってから再度お試しください。');
  }

  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(SHEET_NAME);
    
    if (!sheet) {
      throw new Error('「' + SHEET_NAME + '」シートが見つかりません。');
    }
    
    const data = sheet.getDataRange().getValues();
    
    // 1行目は見出しと想定し、2行目(インデックス1)から検索
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      
      // キー番号を文字列化して完全一致で判定
      if (String(row[COL_KEY]) === String(keyNumber)) {
        const isArrived = (String(row[COL_STATUS]) === '来場済み');
        
        const currentCount = isArrived ? row[COL_GUEST_COUNT] : (guestCount || 1);
        
        const guestInfo = {
          inviter: row[COL_INVITER],
          customer: row[COL_CUSTOMER],
          company: row[COL_COMPANY],
          note: row[COL_NOTE],
          guestCount: currentCount
        };
        
        if (isArrived) {
          return { status: 'already_arrived', guest: guestInfo };
        } else {
          // 未来場の場合 -> スプレッドシートを更新
          const nowFormatted = getFormattedDate();
          
          sheet.getRange(i + 1, COL_STATUS + 1).setValue('来場済み');
          sheet.getRange(i + 1, COL_ARRIVED_AT + 1).setValue(nowFormatted);
          sheet.getRange(i + 1, COL_GUEST_COUNT + 1).setValue(guestCount || 1);
          sheet.getRange(i + 1, COL_CHILD_COUNT + 1).setValue(childCount || 0);
          
          // データ更新を即時に反映させる（排他制御時に確実に行うため）
          SpreadsheetApp.flush();
          
          return { status: 'success', guest: guestInfo };
        }
      }
    }
    
    return { status: 'not_found' };
    
  } finally {
    // 処理が完了した（またはエラーになった）後、必ずロックを解除する
    lock.releaseLock();
  }
}

/**
 * 受付を取り消す関数（排他制御付き）
 * @param {string} keyNumber 検索キー
 * @return {object} 処理結果
 */
function resetReception(keyNumber) {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(15000)) {
    throw new Error('現在、他の端末が処理中です。少し待ってから再度お試しください。');
  }

  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(SHEET_NAME);
    
    if (!sheet) {
      throw new Error('「' + SHEET_NAME + '」シートが見つかりません。');
    }
    
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][COL_KEY]) === String(keyNumber)) {
        // 来場状況、来場日時、人数のセルを空にする（取り消し）
        sheet.getRange(i + 1, COL_STATUS + 1).clearContent();
        sheet.getRange(i + 1, COL_ARRIVED_AT + 1).clearContent();
        sheet.getRange(i + 1, COL_GUEST_COUNT + 1).clearContent();
        sheet.getRange(i + 1, COL_CHILD_COUNT + 1).clearContent();
        
        SpreadsheetApp.flush();
        return { status: 'reset_success' };
      }
    }
    
    return { status: 'not_found' };
    
  } finally {
    lock.releaseLock();
  }
}

/**
 * 外部からのPOST通信（GitHub Pages等のフロントエンドから）を受け取る関数
 */
function doPost(e) {
  try {
    // 送信されてきたデータをJSONとして解析（Preflight回避のためContent-Typeはtext/plainを想定）
    const params = JSON.parse(e.postData.contents);
    let result = {};
    
    // リクエストのactionに応じて処理を振り分け
    if (params.action === 'processReception') {
      result = processReception(params.keyNumber, params.guestCount, params.childCount);
    } else if (params.action === 'resetReception') {
      result = resetReception(params.keyNumber);
    } else if (params.action === 'processVision') {
      result = processVision(params.base64Image);
    } else {
      result = { status: 'error', message: '不明なアクションです' };
    }
    
    // 処理結果をJSON形式の文字列にして返却
    return ContentService.createTextOutput(JSON.stringify(result))
      .setMimeType(ContentService.MimeType.JSON);
      
  } catch (error) {
    // エラー発生時はエラー内容を返す
    const errorResult = { status: 'error', message: error.message };
    return ContentService.createTextOutput(JSON.stringify(errorResult))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

/**
 * 画像からテキストを抽出する関数 (Cloud Vision APIを使用)
 */
function processVision(base64Image) {
  if (!VISION_API_KEY) {
    throw new Error('VISION_API_KEYが設定されていません。GASのコードを編集してAPIキーを設定してください。');
  }

  // Base64文字列から"data:image/jpeg;base64,"の部分を取り除く
  const content = base64Image.replace(/^data:image\/[a-z]+;base64,/, "");

  const url = 'https://vision.googleapis.com/v1/images:annotate?key=' + VISION_API_KEY;
  const payload = {
    "requests": [
      {
        "image": {
          "content": content
        },
        "features": [
          {
            "type": "DOCUMENT_TEXT_DETECTION"
          }
        ]
      }
    ]
  };

  const options = {
    "method": "post",
    "contentType": "application/json",
    "payload": JSON.stringify(payload)
  };

  try {
    const response = UrlFetchApp.fetch(url, options);
    const result = JSON.parse(response.getContentText());

    if (result.responses && result.responses[0] && result.responses[0].fullTextAnnotation) {
      const detectedText = result.responses[0].fullTextAnnotation.text;
      
      // 改行や空白を除去して返す
      const cleanText = detectedText.replace(/[\r\n\s]+/g, '');
      
      // 抽出したテキストから記号以外の英数字を探す
      const match = cleanText.match(/[a-zA-Z0-9-]{3,}/);
      
      if (match) {
        return { status: 'success', text: match[0] };
      } else if (cleanText.length > 0) {
        return { status: 'success', text: cleanText };
      } else {
        return { status: 'not_found', message: '読み取れる文字が見つかりませんでした。' };
      }
    } else {
      return { status: 'not_found', message: '文字を検出できませんでした。' };
    }
  } catch (error) {
    throw new Error('Vision APIの呼び出しに失敗しました: ' + error.message);
  }
}


