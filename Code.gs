// 連携先スプレッドシートのID
const SPREADSHEET_ID = "1SjRkvg9kk1YFjJKROBlxXFHp10xO3_mC-JSq9yAM9lk";

// マスターシート名
const MASTER_SHEET_NAME = "勤怠マスタ";
const MEMBER_SHEET_SUFFIX = "_勤務表";
const CATEGORY_USERS = ["田中", "牛嶋", "長谷川", "住吉"];
const WORK_CATEGORIES = ["アカデミー", "ホームワイン", "その他（WT業務）"];
// 管理者認証情報はスクリプトプロパティで管理する。

// マスターシートのヘッダー
const MASTER_HEADERS = [
  "タイムスタンプ",
  "名前",
  "区分",
  "打刻日時",
  "対象月",
  "システムID",
  "交通機関",
  "備考",
  "業務区分",
  "実働時間"
];

/**
 * GETリクエスト受信用
 * HTML側からの action=read / action=add / action=delete /
 * action=addUser / action=deleteUser をここで処理します。
 */
function doGet(e) {
  return handleRequest_(e, false);
}

/**
 * POSTリクエスト受信用
 * 念のためPOSTでも同じ処理を通します。
 */
function doPost(e) {
  return handleRequest_(e, true);
}

/**
 * メイン処理
 */
function handleRequest_(e, isPost) {
  e = e || {};
  const params = e.parameter || {};
  const action = params.action || "read";

  const adminActions = ["adminLogin", "addUser", "deleteUser", "delete"];
  if (action === "adminResult") {
    const key = String(params.receipt || "");
    const value = /^[a-f0-9-]{36}$/.test(key) ? CacheService.getScriptCache().get("result:" + key) : null;
    return createResponse_(e, value ? JSON.parse(value) : {ok:false, pending:true});
  }
  if (adminActions.indexOf(action) !== -1) {
    if (!isPost || params.callback) return createResponse_(e, {ok:false, error:"post_required"});
    if (!/^[a-f0-9-]{36}$/.test(String(params.receipt || ""))) return createResponse_(e, {ok:false, error:"invalid_receipt"});
    let result;
    try {
      if (action === "adminLogin") result = loginAdmin_(params);
      else {
        result = validateAdmin_(params, action);
        if (!result) {
          const adminSS = SpreadsheetApp.openById(SPREADSHEET_ID);
          if (action === "addUser") result = addUser_(adminSS, params);
          else if (action === "deleteUser") result = deleteUser_(adminSS, params);
          else result = deleteLog_(getOrCreateMasterSheet_(adminSS), params);
        }
      }
    } catch (err) { result = {ok:false, error:"admin_action_failed", message:"管理者操作に失敗しました。再同期して状態を確認してください。"}; }
    CacheService.getScriptCache().put("result:" + params.receipt, JSON.stringify(result), 60);
    return createResponse_(e, result);
  }
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = getOrCreateMasterSheet_(ss);

    if (action === "read") {
      return createResponse_(e, readData_(ss, sheet));
    }

    if (action === "add") {
      return createResponse_(e, addLog_(sheet, params));
    }

    if (action === "delete") {
      return createResponse_(e, deleteLog_(sheet, params));
    }

    if (action === "addUser") {
      return createResponse_(e, addUser_(ss, params));
    }

    if (action === "deleteUser") {
      return createResponse_(e, deleteUser_(ss, params));
    }

    return createResponse_(e, {
      ok: false,
      error: "unknown_action",
      message: "不明なactionです: " + action
    });

  } catch (err) {
    return createResponse_(e, {
      ok: false,
      error: String(err && err.message ? err.message : err),
      stack: String(err && err.stack ? err.stack : "")
    });
  }
}

/**
 * 勤怠マスタシートを取得。なければ作成。
 */
function getOrCreateMasterSheet_(ss) {
  let sheet = ss.getSheetByName(MASTER_SHEET_NAME);

  if (!sheet) {
    sheet = ss.insertSheet(MASTER_SHEET_NAME);
    sheet.appendRow(MASTER_HEADERS);
    sheet.getRange(1, 1, 1, MASTER_HEADERS.length).setFontWeight("bold");
    sheet.setFrozenRows(1);
    return sheet;
  }

  if (sheet.getLastRow() === 0) {
    sheet.appendRow(MASTER_HEADERS);
    sheet.getRange(1, 1, 1, MASTER_HEADERS.length).setFontWeight("bold");
    sheet.setFrozenRows(1);
    return sheet;
  }

  const headerRange = sheet.getRange(1, 1, 1, MASTER_HEADERS.length);
  const currentHeaders = headerRange.getDisplayValues()[0];
  let needsHeaderFix = false;

  for (let i = 0; i < MASTER_HEADERS.length; i++) {
    if (!currentHeaders[i]) {
      needsHeaderFix = true;
      break;
    }
  }

  if (needsHeaderFix) {
    headerRange.setValues([MASTER_HEADERS]);
    headerRange.setFontWeight("bold");
    sheet.setFrozenRows(1);
  }

  return sheet;
}

/**
 * 打刻データ追加
 */
function addLog_(sheet, params) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);

  try {
    const name = params.name || "";
    const type = params.type || "";
    const time = params.time || "";
    const month = params.month || deriveMonth_(time);
    const transport = params.transport || "";
    const memo = params.memo || "";
    const category = String(params.category || "");
    if (category && (!CATEGORY_USERS.includes(name) || !WORK_CATEGORIES.includes(category))) {
      return {ok:false, action:"add", error:"invalid_category", message:"業務区分を確認してください。"};
    }

    if (!name || !type || !time) {
      return {
        ok: false,
        action: "add",
        error: "missing_required_params",
        message: "name / type / time のいずれかが不足しています。"
      };
    }

    // HTML側が発行したIDをそのまま保存する。同一リクエストが再送されても、
    // 同じIDまたは同じ打刻内容なら二重登録しない。
    const logId = String(params.requestId || params.id || new Date().getTime());
    const existingLogId = findExistingLogId_(sheet, logId, name, type, time, category);
    if (existingLogId) {
      return {
        ok: true,
        action: "add",
        id: existingLogId,
        duplicate: true,
        message: "すでに登録済みのため、重複登録を防止しました。"
      };
    }

    sheet.appendRow([
      new Date(),
      name,
      type,
      time,
      month,
      logId,
      transport,
      memo,
      category,
      ""
    ]);
    const rowNumber = sheet.getLastRow();
    // 同一人物・業務区分の直前の打刻が勤務中なら、その区間の実働を加算。
    // セル参照なので、管理者が打刻時刻を修正した場合にも再計算される。
    if (category && rowNumber > 2) {
      const r = rowNumber, p = r - 1;
      const match = '($B$2:$B$'+p+'=B'+r+')*($I$2:$I$'+p+'=I'+r+')';
      const prevType = 'LOOKUP(2,ARRAYFORMULA(1/('+match+')),$C$2:$C$'+p+')';
      const prevTime = 'LOOKUP(2,ARRAYFORMULA(1/('+match+')),$D$2:$D$'+p+')';
      sheet.getRange(r,10).setFormula('=IFERROR(IF(AND(OR(C'+r+'="退勤",C'+r+'="休憩開始"),OR('+prevType+'="出勤",'+prevType+'="休憩終了")),MAX(0,VALUE(D'+r+')-VALUE('+prevTime+')),0),0)').setNumberFormat('[h]:mm:ss');
    }
    SpreadsheetApp.flush();

    return {
      ok: true,
      action: "add",
      id: logId,
      message: "SUCCESS"
    };

  } finally {
    lock.releaseLock();
  }
}

/**
 * 同じ送信ID、または名前・区分・打刻日時が完全一致する行を探す。
 * 通信遅延による同一URLの再送を安全に1件へまとめる。
 */
function findExistingLogId_(sheet, logId, name, type, time, category) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return "";

  const rows = sheet.getRange(2, 1, lastRow - 1, MASTER_HEADERS.length).getDisplayValues();
  for (let i = rows.length - 1; i >= 0; i--) {
    const row = rows[i];
    const rowId = row[5] ? String(row[5]) : "";
    const sameId = rowId && rowId === logId;
    const samePunch = row[1] === name && row[2] === type && row[3] === time && (row[8] || "") === (category || "");

    if (sameId || samePunch) {
      return rowId || logId;
    }
  }

  return "";
}

/**
 * データ読み取り
 */
function readData_(ss, sheet) {
  const data = sheet.getDataRange().getDisplayValues();
  const logs = [];

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const id = row[5] ? String(row[5]) : String(i);
    const name = row[1] || "";
    const type = row[2] || "";
    const time = row[3] || "";
    const month = row[4] || deriveMonth_(time);
    const transport = row[6] || "";
    const memo = row[7] || "";

    if (!name && !type && !time) {
      continue;
    }

    logs.push({
      id: id,
      name: name,
      type: type,
      time: time,
      month: month,
      transport: transport,
      memo: memo,
      category: row[8] || ""
    });
  }

  logs.reverse();
  const roster = readRoster_(ss);

  return {
    ok: true,
    action: "read",
    schemaVersion: 2,
    logs: logs,
    users: roster.users,
    transportationCosts: roster.transportationCosts,
    serverTime: Utilities.formatDate(
      new Date(),
      "Asia/Tokyo",
      "yyyy-MM-dd HH:mm:ss"
    )
  };
}

/**
 * 「〇〇_勤務表」シートを全端末共通の名簿として読み取る。
 */
function readRoster_(ss) {
  const users = [];
  const transportationCosts = {};

  ss.getSheets().forEach(function(sheet) {
    const sheetName = sheet.getName();

    if (!sheetName.endsWith(MEMBER_SHEET_SUFFIX)) {
      return;
    }

    const name = sheetName.slice(0, -MEMBER_SHEET_SUFFIX.length);
    if (!name) {
      return;
    }

    users.push(name);

    try {
      const costValue = sheet.getRange("I1").getValue();
      const cost = parseInt(costValue, 10);
      transportationCosts[name] = isNaN(cost) ? 0 : cost;
    } catch (err) {
      transportationCosts[name] = 0;
    }
  });

  return {
    users: users,
    transportationCosts: transportationCosts
  };
}

/**
 * 管理者画面からメンバーを追加する。
 * 既存の勤務表をテンプレートとして複製し、C1へ名前、I1へ交通費0円を設定する。
 */
function addUser_(ss, params) {
  const authError = validateAdmin_(params, "addUser");
  if (authError) return authError;

  const lock = LockService.getScriptLock();
  lock.waitLock(10000);

  try {
    const nameResult = validateUserName_(params.name);
    if (!nameResult.ok) return nameResult;

    const name = nameResult.name;
    const newSheetName = name + MEMBER_SHEET_SUFFIX;

    if (ss.getSheetByName(newSheetName)) {
      return {
        ok: false,
        action: "addUser",
        error: "already_exists",
        message: name + "さんはすでに名簿に登録されています。"
      };
    }

    const memberSheets = getMemberSheets_(ss);
    if (!memberSheets.length) {
      return {
        ok: false,
        action: "addUser",
        error: "template_not_found",
        message: "複製元になる勤務表が見つかりません。"
      };
    }

    const newSheet = memberSheets[0].copyTo(ss).setName(newSheetName);
    newSheet.getRange("C1").setValue(name);
    newSheet.getRange("I1").setValue(0);
    ss.setActiveSheet(newSheet);
    ss.moveActiveSheet(ss.getNumSheets());
    SpreadsheetApp.flush();

    const roster = readRoster_(ss);
    return {
      ok: true,
      action: "addUser",
      users: roster.users,
      transportationCosts: roster.transportationCosts,
      message: name + "さんを名簿に追加しました。"
    };

  } finally {
    lock.releaseLock();
  }
}

/**
 * 管理者画面からメンバーを削除する。
 * 個別勤務表は非表示の削除済みシートとして退避し、勤怠マスタの過去ログも残す。
 */
function deleteUser_(ss, params) {
  const authError = validateAdmin_(params, "deleteUser");
  if (authError) return authError;

  const lock = LockService.getScriptLock();
  lock.waitLock(10000);

  try {
    const nameResult = validateUserName_(params.name);
    if (!nameResult.ok) return nameResult;

    const name = nameResult.name;
    const targetSheet = ss.getSheetByName(name + MEMBER_SHEET_SUFFIX);

    if (!targetSheet) {
      return {
        ok: false,
        action: "deleteUser",
        error: "not_found",
        message: name + "さんは名簿に登録されていません。"
      };
    }

    const memberSheets = getMemberSheets_(ss);
    if (memberSheets.length <= 1) {
      return {
        ok: false,
        action: "deleteUser",
        error: "last_user",
        message: "名簿には1人以上必要です。"
      };
    }

    preserveMonthSourceBeforeDelete_(memberSheets, targetSheet);
    const archiveName = createArchivedSheetName_(ss, name);
    targetSheet.setName(archiveName);
    targetSheet.hideSheet();
    SpreadsheetApp.flush();

    const roster = readRoster_(ss);
    return {
      ok: true,
      action: "deleteUser",
      users: roster.users,
      transportationCosts: roster.transportationCosts,
      message: name + "さんを名簿から削除しました。個別勤務表と過去ログは退避・保持されています。"
    };

  } finally {
    lock.releaseLock();
  }
}

function createArchivedSheetName_(ss, name) {
  const timestamp = Utilities.formatDate(
    new Date(),
    "Asia/Tokyo",
    "yyyyMMdd_HHmmss"
  );
  const baseName = name + MEMBER_SHEET_SUFFIX + "_削除済_" + timestamp;
  let candidate = baseName;
  let suffix = 2;

  while (ss.getSheetByName(candidate)) {
    candidate = baseName + "_" + suffix;
    suffix++;
  }

  return candidate.slice(0, 100);
}

function getMemberSheets_(ss) {
  return ss.getSheets().filter(function(sheet) {
    return sheet.getName().endsWith(MEMBER_SHEET_SUFFIX);
  });
}

/**
 * 各勤務表のB1が削除対象シートを参照している場合、別の勤務表へ参照を付け替える。
 */
function preserveMonthSourceBeforeDelete_(memberSheets, targetSheet) {
  const targetSheetName = targetSheet.getName();
  const replacementSheet = memberSheets.find(function(sheet) {
    return sheet.getSheetId() !== targetSheet.getSheetId();
  });

  if (!replacementSheet) return;

  const targetMonth = targetSheet.getRange("B1").getValue();
  const replacementMonthCell = replacementSheet.getRange("B1");
  const replacementFormula = replacementMonthCell.getFormula();

  if (replacementFormula.indexOf("'" + targetSheetName.replace(/'/g, "''") + "'") !== -1) {
    replacementMonthCell.setValue(targetMonth);
  }

  const replacementNameEscaped = replacementSheet.getName().replace(/'/g, "''");
  memberSheets.forEach(function(sheet) {
    if (sheet.getSheetId() === targetSheet.getSheetId() ||
        sheet.getSheetId() === replacementSheet.getSheetId()) {
      return;
    }

    const monthCell = sheet.getRange("B1");
    const formula = monthCell.getFormula();
    if (formula.indexOf("'" + targetSheetName.replace(/'/g, "''") + "'") !== -1) {
      monthCell.setFormula("='" + replacementNameEscaped + "'!B1");
    }
  });
}

function loginAdmin_(params) {
  const props = PropertiesService.getScriptProperties();
  const expected = props.getProperty("ADMIN_PASSWORD_SHA256");
  if (!expected) return {ok:false, error:"admin_not_configured", message:"管理者認証が未設定です。"};
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const cache = CacheService.getScriptCache();
    const failures = Number(cache.get("adminFailures") || 0);
    if (failures >= 20) return {ok:false, error:"rate_limited", message:"しばらく待ってから認証してください。"};
    const actual = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, String(params.adminPassword || ""), Utilities.Charset.UTF_8)
      .map(function(b){ return (b & 255).toString(16).padStart(2, "0"); }).join("");
    if (actual !== expected) {
      cache.put("adminFailures", String(failures + 1), 300);
      return {ok:false, error:"unauthorized", message:"パスワードが正しくありません。"};
    }
    cache.remove("adminFailures");
    const token = Utilities.getUuid() + Utilities.getUuid();
    cache.put("admin:" + token, "valid", 1800);
    return {ok:true, action:"adminLogin", token:token};
  } finally { lock.releaseLock(); }
}

function validateAdmin_(params, action) {
  const token = String(params.adminToken || "");
  if (/^[a-f0-9-]{72}$/.test(token) && CacheService.getScriptCache().get("admin:" + token) === "valid") return null;
  return {ok:false, action:action, error:"unauthorized", message:"管理者認証が必要です。もう一度ログインしてください。"};
}

function validateUserName_(value) {
  const name = String(value || "").trim().replace(/\s+/g, " ");

  if (!name) {
    return {
      ok: false,
      error: "missing_name",
      message: "追加・削除する方のお名前を入力してください。"
    };
  }

  if (name.length > 20) {
    return {
      ok: false,
      error: "name_too_long",
      message: "お名前は20文字以内で入力してください。"
    };
  }

  if (/[\\\/?*\[\]:]/.test(name)) {
    return {
      ok: false,
      error: "invalid_name",
      message: "お名前に使用できない記号が含まれています。"
    };
  }

  return { ok: true, name: name };
}

/**
 * 打刻データ削除
 * idがあればid優先。idがない場合は name + time で削除。
 */
function deleteLog_(sheet, params) {
  const authError = validateAdmin_(params, "delete");
  if (authError) return authError;
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);

  try {
    const targetId = params.id || "";
    const targetName = params.name || "";
    const targetTime = params.time || "";

    if (!targetId && (!targetName || !targetTime)) {
      return {
        ok: false,
        action: "delete",
        error: "missing_delete_params",
        message: "削除対象の id または name/time が不足しています。"
      };
    }

    const data = sheet.getDataRange().getDisplayValues();
    let isDeleted = false;

    for (let i = data.length - 1; i >= 1; i--) {
      const row = data[i];
      const rowName = row[1] || "";
      const rowTime = row[3] || "";
      const rowId = row[5] ? String(row[5]) : "";
      const matchedById = targetId && rowId === String(targetId);
      const matchedByNameTime = !targetId && rowName === targetName && rowTime === targetTime;

      if (matchedById || matchedByNameTime) {
        sheet.deleteRow(i + 1);
        isDeleted = true;
        break;
      }
    }

    if (isDeleted) {
      return { ok: true, action: "delete", message: "SUCCESS" };
    }

    return {
      ok: false,
      action: "delete",
      error: "not_found",
      message: "ERROR: 対象データなし"
    };

  } finally {
    lock.releaseLock();
  }
}

/**
 * JSON / JSONP 両対応のレスポンス作成
 */
function createResponse_(e, result) {
  const params = e && e.parameter ? e.parameter : {};
  const output = JSON.stringify(result);

  if (params.callback) {
    const callback = String(params.callback);
    if (/^[A-Za-z_$][0-9A-Za-z_$]*$/.test(callback)) {
      return ContentService
        .createTextOutput(callback + "(" + output + ");")
        .setMimeType(ContentService.MimeType.JAVASCRIPT);
    }
  }

  return ContentService
    .createTextOutput(output)
    .setMimeType(ContentService.MimeType.JSON);
}

/**
 * timeから yyyy-MM を推定
 */
function deriveMonth_(time) {
  if (!time) {
    return Utilities.formatDate(new Date(), "Asia/Tokyo", "yyyy-MM");
  }

  const s = String(time);
  if (/^\d{4}-\d{2}/.test(s)) {
    return s.slice(0, 7);
  }

  if (/^\d{4}\/\d{2}/.test(s)) {
    return s.slice(0, 7).replace("/", "-");
  }

  const d = new Date(s);
  if (!isNaN(d.getTime())) {
    return Utilities.formatDate(d, "Asia/Tokyo", "yyyy-MM");
  }

  return Utilities.formatDate(new Date(), "Asia/Tokyo", "yyyy-MM");
}

function testRead() {
  const response = handleRequest_({ parameter: { action: "read" } });
  Logger.log(response.getContent());
}

function testReadJsonp() {
  const response = handleRequest_({
    parameter: { action: "read", callback: "test" }
  });
  Logger.log(response.getContent());
}
