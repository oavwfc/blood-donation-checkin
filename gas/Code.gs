// =============================================
// 捐血報到系統 - Google Apps Script 後端
// =============================================
// 部署說明:
// 1. 開啟 Google Sheet，點擊「擴充功能」>「Apps Script」
// 2. 將此檔案內容貼入，存檔
// 3. 點擊「部署」>「新增部署作業」
// 4. 類型選「網頁應用程式」
// 5. 執行身分: 我 (你的 Google 帳號)
// 6. 存取權: 所有人 (含匿名使用者)
// 7. 部署後複製網址，貼到 js/config.js 的 GAS_URL
// =============================================

const SHEET_NAME = 'blood_donation';

// ── 主要進入點 ──────────────────────────────

function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    const action = data.action;

    let result;
    switch (action) {
      case 'register':        result = register(data);        break;
      case 'checkin':         result = checkin(data);         break;
      case 'checkin_new':     result = checkinNew(data);      break;
      case 'lookup_checkin':  result = lookupCheckin(data);   break;
      case 'lookup_phone':    result = lookupPhone(data);     break;
      case 'update_donation': result = updateDonation(data);  break;
      case 'update_profile':  result = updateProfile(data);   break;
      case 'admin_login':     result = adminLogin(data);      break;
      case 'admin_lookup':    result = adminLookup(data);     break;
      case 'update_gift':     result = updateGift(data);      break;
      default:
        result = { status: 'error', message: '未知的操作類型' };
    }

    return buildResponse(result);
  } catch (err) {
    return buildResponse({ status: 'error', message: err.toString() });
  }
}

function doGet(e) {
  return buildResponse({ status: 'ok', message: '捐血報到系統運作正常' });
}

function buildResponse(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

// ── 工具函式 ────────────────────────────────

function getSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);

  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    const headers = [
      '報名時間', '公司行號', '姓名', '電話',
      '捐血CC數', '報到時間', '報到編號', '報到', '捐血成功', '禮品領取'
    ];
    sheet.appendRow(headers);
    sheet.setFrozenRows(1);

    const headerRange = sheet.getRange(1, 1, 1, headers.length);
    headerRange.setBackground('#C62828');
    headerRange.setFontColor('#FFFFFF');
    headerRange.setFontWeight('bold');
  }

  return sheet;
}

function formatDate(date) {
  const d = new Date(date);
  const yyyy = d.getFullYear();
  const mm   = String(d.getMonth() + 1).padStart(2, '0');
  const dd   = String(d.getDate()).padStart(2, '0');
  return `${yyyy}/${mm}/${dd}`;
}

function generateCheckinNumber(sheet) {
  const data = sheet.getDataRange().getValues();
  let maxNum = 0;

  for (let i = 1; i < data.length; i++) {
    const val = data[i][6]; // 報到編號 (G欄, index 6)
    if (val) {
      const num = parseInt(val);
      if (!isNaN(num) && num > maxNum) maxNum = num;
    }
  }

  return String(maxNum + 1).padStart(4, '0');
}

// ── 報名 ────────────────────────────────────
// 欄位對應: A報名時間 B公司行號 C姓名 D電話 E捐血CC數
//           F報到時間 G報到編號 H報到 I捐血成功 J禮品領取

function register(data) {
  const sheet = getSheet();
  const rows  = sheet.getDataRange().getValues();

  // 重複報名檢查 (姓名 + 電話)
  for (let i = 1; i < rows.length; i++) {
    if (rows[i][2] === data.姓名 && String(rows[i][3]) === String(data.電話)) {
      return { status: 'duplicate', message: '此姓名與手機號碼已完成報名，無需重複報名。' };
    }
  }

  const now    = formatDate(new Date());
  const newRow = sheet.getLastRow() + 1;

  sheet.getRange(newRow, 1, 1, 10).setValues([[
    now, data.公司行號, data.姓名, '',
    data.捐血CC數, '', '', '', '', 'Y'
  ]]);

  var phoneCell = sheet.getRange(newRow, 4); // D 電話
  phoneCell.setNumberFormat('@');
  phoneCell.setValue(String(data.電話));

  return { status: 'success', message: '報名成功！' };
}

// ── 報到（已報名） ───────────────────────────

function checkin(data) {
  const sheet = getSheet();
  const rows  = sheet.getDataRange().getValues();

  for (let i = 1; i < rows.length; i++) {
    if (rows[i][2] === data.姓名 &&
        String(rows[i][3]) === String(data.電話)) {

      // 已報到
      if (rows[i][7] === 'Y') {
        return {
          status: 'already_checkin',
          message: `${rows[i][2]} 您已完成報到`,
          報到編號: rows[i][6]
        };
      }

      const checkinNum = generateCheckinNumber(sheet);
      const now        = formatDate(new Date());
      const rowNum     = i + 1;

      sheet.getRange(rowNum, 6).setValue(now);                              // F 報到時間
      sheet.getRange(rowNum, 7).setNumberFormat('@').setValue(checkinNum); // G 報到編號
      sheet.getRange(rowNum, 8).setValue('Y');                             // H 報到

      return {
        status: 'success',
        message: '報到成功！',
        報到編號: checkinNum,
        姓名: rows[i][2],
        捐血CC數: rows[i][4]
      };
    }
  }

  return {
    status: 'not_found',
    message: '查無報名資料，請確認輸入是否正確，或改用「未報名」方式報到。'
  };
}

// ── 報到（未報名，現場直接新增） ─────────────

function checkinNew(data) {
  const sheet      = getSheet();
  const now        = formatDate(new Date());
  const checkinNum = generateCheckinNumber(sheet);

  const newRow = sheet.getLastRow() + 1;

  sheet.getRange(newRow, 1, 1, 10).setValues([[
    now, data.公司行號, data.姓名, '',
    data.捐血CC數, now, '', 'Y', '', 'Y'
  ]]);

  var phoneCell = sheet.getRange(newRow, 4); // D 電話
  phoneCell.setNumberFormat('@');
  phoneCell.setValue(String(data.電話));

  sheet.getRange(newRow, 7).setNumberFormat('@').setValue(checkinNum); // G 報到編號

  return {
    status: 'success',
    message: '報到成功！',
    報到編號: checkinNum,
    姓名: data.姓名,
    捐血CC數: data.捐血CC數
  };
}

// ── 查詢報到編號（只查不改） ──────────────────

function lookupCheckin(data) {
  const sheet  = getSheet();
  const rows   = sheet.getDataRange().getValues();
  const target = String(data.報到編號).padStart(4, '0');

  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][6]).padStart(4, '0') === target) {
      return {
        status: 'success',
        公司行號: rows[i][1],
        姓名:     rows[i][2],
        電話:     rows[i][3],
        報到編號: rows[i][6],
        捐血CC數: rows[i][4],
        捐血成功: rows[i][8]
      };
    }
  }

  return {
    status: 'not_found',
    message: `查無報到編號 ${data.報到編號}，請確認後再試。`
  };
}

// ── 捐血後更新 ───────────────────────────────

function updateDonation(data) {
  const sheet  = getSheet();
  const rows   = sheet.getDataRange().getValues();
  const target = String(data.報到編號).padStart(4, '0');

  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][6]).padStart(4, '0') === target) {
      const rowNum = i + 1;
      sheet.getRange(rowNum, 9).setValue(data.捐血成功);  // I 捐血成功
      if (data.捐血成功 === 'N') {
        sheet.getRange(rowNum, 10).setValue('無符合資格'); // J 禮品領取
      }

      return {
        status: 'success',
        message: '資料更新成功！',
        公司行號: rows[i][1],
        姓名:     rows[i][2],
        電話:     rows[i][3],
        報到編號: rows[i][6],
        捐血CC數: rows[i][4]
      };
    }
  }

  return {
    status: 'not_found',
    message: `查無報到編號 ${data.報到編號}，請確認後再試。`
  };
}

// ── 修改報名資料 ─────────────────────────────

function updateProfile(data) {
  const sheet = getSheet();
  const rows  = sheet.getDataRange().getValues();
  const phone = String(data.電話);

  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][3]) === phone) {
      const rowNum = i + 1;
      if (data.公司行號) sheet.getRange(rowNum, 2).setValue(data.公司行號); // B
      if (data.姓名)     sheet.getRange(rowNum, 3).setValue(data.姓名);     // C
      if (data.捐血CC數) sheet.getRange(rowNum, 5).setValue(data.捐血CC數); // E

      return {
        status:   'success',
        message:  '資料更新成功！',
        公司行號: data.公司行號 || rows[i][1],
        姓名:     data.姓名     || rows[i][2],
        電話:     rows[i][3],
        捐血CC數: data.捐血CC數 || rows[i][4],
        報到:     rows[i][7],
        報到編號: rows[i][6],
        捐血成功: rows[i][8]
      };
    }
  }

  return { status: 'not_found', message: '查無此資料，無法更新。' };
}

// ── 管理者登入 ───────────────────────────────


function adminLogin(data) {
  const storedHash = PropertiesService.getScriptProperties().getProperty('ADMIN_PASSWORD');
  if (!storedHash) {
    return { status: 'error', message: '管理者密碼尚未設定，請聯絡系統管理員。' };
  }

  // 將輸入密碼 SHA-256 後與儲存的 hash 比對
  const inputDigest = Utilities.computeDigest(
    Utilities.DigestAlgorithm.SHA_256,
    data.password || ''
  );
  const inputHash = inputDigest.map(function(b) {
    return ('0' + (b & 0xff).toString(16)).slice(-2);
  }).join('');

  if (inputHash !== storedHash) {
    return { status: 'error', message: '密碼錯誤，請重新輸入。' };
  }

  // 產生小時級 token（SHA-256(hash:小時數)），每小時自動更新
  const hourly = Math.floor(Date.now() / 3600000);
  const tokenDigest = Utilities.computeDigest(
    Utilities.DigestAlgorithm.SHA_256,
    storedHash + ':' + hourly
  );
  const token = tokenDigest.map(function(b) {
    return ('0' + (b & 0xff).toString(16)).slice(-2);
  }).join('');

  return { status: 'success', token: token };
}

function validateAdminToken(token) {
  const stored = PropertiesService.getScriptProperties().getProperty('ADMIN_PASSWORD');
  if (!stored || !token) return false;

  const now = Math.floor(Date.now() / 3600000);
  // 允許目前小時與前一小時（跨整點緩衝）
  for (var i = 0; i <= 1; i++) {
    const digest = Utilities.computeDigest(
      Utilities.DigestAlgorithm.SHA_256,
      stored + ':' + (now - i)
    );
    const expected = digest.map(function(b) {
      return ('0' + (b & 0xff).toString(16)).slice(-2);
    }).join('');
    if (token === expected) return true;
  }
  return false;
}

// ── 管理者查詢（手機電話 或 報到編號）──────────

function adminLookup(data) {
  if (!validateAdminToken(data.token)) {
    return { status: 'unauthorized', message: '登入逾時，請重新登入。' };
  }

  const sheet = getSheet();
  const rows  = sheet.getDataRange().getValues();
  var row = null;

  if (data.電話) {
    const phone = String(data.電話);
    for (let i = 1; i < rows.length; i++) {
      if (String(rows[i][3]) === phone) { row = rows[i]; break; }
    }
  } else if (data.報到編號) {
    const target = String(data.報到編號).padStart(4, '0');
    for (let i = 1; i < rows.length; i++) {
      if (String(rows[i][6]).padStart(4, '0') === target) { row = rows[i]; break; }
    }
  }

  if (!row) {
    return { status: 'not_found', message: '查無資料，請確認輸入是否正確。' };
  }

  return {
    status:   'success',
    報名時間: row[0],
    公司行號: row[1],
    姓名:     row[2],
    電話:     String(row[3]),
    捐血CC數: row[4],
    報到時間: row[5],
    報到編號: String(row[6]),
    報到:     row[7],
    捐血成功: row[8]
  };
}

// ── 禮品領取更新 ─────────────────────────────

function updateGift(data) {
  if (!validateAdminToken(data.token)) {
    return { status: 'unauthorized', message: '登入逾時，請重新登入。' };
  }

  const sheet  = getSheet();
  const rows   = sheet.getDataRange().getValues();
  const target = String(data.報到編號).padStart(4, '0');

  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][6]).padStart(4, '0') === target) {
      sheet.getRange(i + 1, 10).setValue(data.禮品領取); // J 禮品領取
      return { status: 'success', message: '禮品領取狀態已更新。' };
    }
  }

  return { status: 'not_found', message: '查無此報到編號，無法更新。' };
}

// ── 手機號碼查詢（捐血前入口） ─────────────────

function lookupPhone(data) {
  const sheet = getSheet();
  const rows  = sheet.getDataRange().getValues();
  const phone = String(data.電話);

  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][3]) === phone) {
      return {
        status: 'found',
        公司行號: rows[i][1],
        姓名:     rows[i][2],
        電話:     rows[i][3],
        捐血CC數: rows[i][4],
        報到:     rows[i][7],
        報到編號: rows[i][6],
        捐血成功: rows[i][8]
      };
    }
  }

  return { status: 'not_found', message: '查無此手機號碼的報名資料。' };
}
