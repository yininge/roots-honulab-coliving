/**
 * 池上共居祭 — 報名表收件端（Google Apps Script Web App）
 *
 * 用途：活動頁（index.html / y2k.html）的報名表送出時，POST 到這支 Web App，
 *       驗證後把每一筆報名 append 進一張 Google Sheet，Yi 跟 Cindy 共管審核。
 *
 * 安全層：
 *   1. honeypot 隱藏欄位 — 機器人若填了 `website` 欄，直接靜默丟棄
 *   2. reCAPTCHA v3 — 驗證 token；無 token / 驗證失敗 → 擋下不寫入（擋直接打 API 的灌水）
 *   3. 公式注入防護 — 欄位開頭是 = + - @ 會被前綴單引號，避免被 Sheet 當公式執行
 *
 * ── 一次性設定步驟 ──────────────────────────────────────────────
 * A. 建立 Sheet
 *    1. 新建 Google Sheet「池上共居祭報名」，把第一個分頁改名為「報名」。
 *    2. 第一列填好表頭（順序對齊下方 appendRow）：
 *       時間 | 姓名 | Email | 聯絡方式 | 方案 | 同住夥伴 | 背景 | 動機 | 期待偏好 | 簡聊時段 | 197晚餐(9/10) | 泥火山豆腐(9/13) | 拍攝同意 | 來源頁 | 驗證分數
 *
 * B. 申請 reCAPTCHA v3 金鑰
 *    1. 到 https://www.google.com/recaptcha/admin/create
 *    2. 類型選「reCAPTCHA v3」，網域填 yininge.github.io（本機測試可再加 localhost）。
 *    3. 取得兩把金鑰：
 *       - 「網站金鑰 Site Key」（公開）→ 等下貼進 index.html / y2k.html 的 RECAPTCHA_SITE_KEY
 *       - 「祕密金鑰 Secret Key」（不可外流）→ 下一步存進 Script Property，不要貼進任何網頁或 repo
 *
 * C. 貼上腳本 + 設定祕密金鑰 + 設定通知收件人
 *    1. 在「池上共居祭報名」Sheet：擴充功能 → Apps Script，把本檔整段貼進去，儲存。
 *    2. 左側齒輪「專案設定」→ 指令碼屬性 → 新增屬性：
 *       名稱 RECAPTCHA_SECRET ，值 = 上面的祕密金鑰。存檔。
 *       （祕密金鑰只存在這裡，永遠不進前端、不進 GitHub repo。）
 *    3. 再新增一個屬性：名稱 NOTIFY_EMAIL ，值 = 報名通知信的收件人。
 *       可逗號分隔多位（例：cindy@x.com,yi@honulab.com）。沒設定就不寄通知信。
 *
 * D. 部署
 *    1. 部署 → 新增部署 → 類型「網頁應用程式」：執行身分=我、誰可存取=任何人。
 *    2. 複製「網頁應用程式」網址。
 *
 * E. 接回前端
 *    1. 把 D 的網址 + B 的 Site Key 回貼給 Yi（或自己填進 index.html / y2k.html 的
 *       FORM_ENDPOINT 與 RECAPTCHA_SITE_KEY）。
 *    2. 把 Sheet 分享給 Cindy（編輯權限），右側自行加狀態欄（待審/錄取/已通知/已付款）。
 *
 * 之後若修改本檔，需「管理部署 → 編輯 → 版本：新版本」重新部署才生效。
 * 註：RECAPTCHA_SECRET 尚未設定前，腳本會略過 reCAPTCHA 驗證（仍可收件，但少一層防護）。
 */
function doPost(e) {
  try {
    var d = JSON.parse(e.postData.contents);

    // 1. honeypot：真人不會看到也不會填這個欄位，有值 = 機器人 → 靜默丟棄
    if (d.website) {
      return json({ ok: true });
    }

    // 2. reCAPTCHA v3 驗證（設了祕密金鑰才驗）
    var score = '';
    var secret = PropertiesService.getScriptProperties().getProperty('RECAPTCHA_SECRET');
    if (secret) {
      if (!d.recaptcha) {
        return json({ ok: false, error: 'missing token' });
      }
      var resp = UrlFetchApp.fetch('https://www.google.com/recaptcha/api/siteverify', {
        method: 'post',
        payload: { secret: secret, response: d.recaptcha },
        muteHttpExceptions: true
      });
      var r = JSON.parse(resp.getContentText());
      if (!r.success) {
        return json({ ok: false, error: 'recaptcha failed' });
      }
      score = r.score; // 0.0（像機器人）~ 1.0（像真人）。低分仍寫入但留分數讓你判斷
    }

    // 3. 寫入（所有字串欄位過公式注入防護）
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('報名');
    sheet.appendRow([
      new Date(),
      c(d.name), c(d.email), c(d.contact), c(d.plan),
      c(d.partner), c(d.background), c(d.motivation), c(d.expectation),
      c(d.calltime), c(d.dinner197), c(d.tofu),
      d.photo ? '是' : '否', c(d.source), score
    ]);

    // 4. 寄送報名通知信給主辦（收件人存在指令碼屬性 NOTIFY_EMAIL）
    notifyOrganizers(d, score);

    return json({ ok: true });
  } catch (err) {
    return json({ ok: false, error: String(err) });
  }
}

/**
 * 報名成功後寄通知信給主辦（Cindy / Yi）。
 * 收件人存在指令碼屬性 NOTIFY_EMAIL，可逗號分隔多位（例：cindy@x.com,yi@honulab.com）。
 * 沒設定就略過；寄信失敗也不影響報名寫入。
 */
function notifyOrganizers(d, score) {
  try {
    var to = PropertiesService.getScriptProperties().getProperty('NOTIFY_EMAIL');
    if (!to) return; // 沒設定收件人就不寄
    var name = d.name || '（未填姓名）';
    var body = [
      '有人填好報名表了 🎉',
      '',
      '姓名：' + (d.name || ''),
      'Email：' + (d.email || ''),
      '聯絡方式：' + (d.contact || ''),
      '方案：' + (d.plan || ''),
      '同住夥伴：' + (d.partner || ''),
      '9/10 197 風味餐晚餐：' + (d.dinner197 || ''),
      '9/13 泥火山豆腐體驗：' + (d.tofu || ''),
      '',
      '背景：' + (d.background || ''),
      '報名動機：' + (d.motivation || ''),
      '期待 / 偏好：' + (d.expectation || ''),
      '方便簡聊時段：' + (d.calltime || ''),
      '拍攝同意：' + (d.photo ? '是' : '否'),
      '來源頁：' + (d.source || ''),
      'reCAPTCHA 分數：' + (score === '' ? '（未啟用）' : score),
      '',
      '— 完整名單在報名試算表，記得去審核 ✦'
    ].join('\n');
    var options = { name: '池上共居祭報名通知' };
    if (d.email) options.replyTo = d.email; // 直接回信即回給報名者
    MailApp.sendEmail(to, '新報名：' + name, body, options);
  } catch (err) {
    console.error('notify failed: ' + err); // 寄信失敗只記錄，不擋報名
  }
}

/**
 * 一鍵測試：在編輯器上方函式選單選 sendTestNotice → 執行。
 * 第一次會跳授權視窗（要按同意，授予寄信權限）；之後會寄一封測試通知信到 NOTIFY_EMAIL。
 * 收到 = 收件人設定與寄信權限都 OK，問題就只剩「有沒有重新部署新版本」。
 */
function sendTestNotice() {
  var to = PropertiesService.getScriptProperties().getProperty('NOTIFY_EMAIL');
  if (!to) {
    throw new Error('NOTIFY_EMAIL 還沒設定：專案設定 → 指令碼屬性 → 新增 NOTIFY_EMAIL');
  }
  notifyOrganizers({
    name: '測試報名',
    email: 'test@example.com',
    contact: 'LINE: test',
    plan: '單人 NT$7,777',
    dinner197: '參加',
    tofu: '不參加',
    background: '（這是 sendTestNotice 寄出的測試信）',
    motivation: '測試通知信',
    photo: true,
    source: 'test'
  }, '');
  Logger.log('已嘗試寄測試信到：' + to);
}

/** 公式注入防護：開頭是 = + - @ 的值前綴單引號，讓 Sheet 當純文字 */
function c(v) {
  v = (v == null) ? '' : String(v);
  return /^[=+\-@]/.test(v) ? ("'" + v) : v;
}

function json(o) {
  return ContentService
    .createTextOutput(JSON.stringify(o))
    .setMimeType(ContentService.MimeType.JSON);
}
