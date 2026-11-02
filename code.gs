
/**
 * 數位科技概論｜版本三
 * GAS 雲端存檔後端
 */

// =================================
// 1. 初始化 Google 試算表
// =================================

function setup() {
  const properties = PropertiesService.getScriptProperties();
  let spreadsheetId = properties.getProperty("SPREADSHEET_ID");

  let spreadsheet;

  if (spreadsheetId) {
    spreadsheet = SpreadsheetApp.openById(spreadsheetId);
  } else {
    spreadsheet = SpreadsheetApp.create("幾A幾B_版本三_玩家資料");
    spreadsheetId = spreadsheet.getId();

    properties.setProperty("SPREADSHEET_ID", spreadsheetId);
  }

  let sheet = spreadsheet.getSheetByName("玩家存檔");

  if (!sheet) {
    sheet = spreadsheet.insertSheet("玩家存檔");
  }

  if (sheet.getLastRow() === 0) {
    sheet.appendRow([
      "玩家代號",
      "遊戲資料(JSON)",
      "最後更新時間"
    ]);
  }

  sheet.setFrozenRows(1);

  Logger.log("初始化完成！");
  Logger.log("試算表網址：" + spreadsheet.getUrl());

  return spreadsheet.getUrl();
}


// =================================
// 2. 網頁入口
// =================================

function doGet() {
  return HtmlService
    .createHtmlOutputFromFile("Index")
    .setTitle("四位數幾 A 幾 B｜版本三")
    .addMetaTag("viewport", "width=device-width, initial-scale=1");
}


// =================================
// 3. 取得試算表
// =================================

function getDatabase() {
  const id = PropertiesService
    .getScriptProperties()
    .getProperty("SPREADSHEET_ID");

  if (!id) {
    throw new Error("尚未初始化資料庫，請先執行 setup。");
  }

  const spreadsheet = SpreadsheetApp.openById(id);
  const sheet = spreadsheet.getSheetByName("玩家存檔");

  if (!sheet) {
    throw new Error("找不到玩家存檔工作表。");
  }

  return sheet;
}


// =================================
// 4. 檢查玩家代號
// =================================

function validateIdentifier(identifier) {
  if (typeof identifier !== "string") {
    throw new Error("代號格式錯誤。");
  }

  identifier = identifier.trim();

  if (identifier.length < 1 || identifier.length > 40) {
    throw new Error("代號長度必須為 1～40 個字元。");
  }

  if (/[\u0000-\u001f]/.test(identifier)) {
    throw new Error("代號包含不允許的字元。");
  }

  return identifier;
}


// =================================
// 5. 查詢玩家資料
// =================================

function getPlayerData(identifier) {
  identifier = validateIdentifier(identifier);

  const sheet = getDatabase();
  const lastRow = sheet.getLastRow();

  if (lastRow < 2) {
    return {
      success: true,
      identifier: identifier,
      games: []
    };
  }

  const data = sheet
    .getRange(2, 1, lastRow - 1, 2)
    .getValues();

  for (let i = 0; i < data.length; i++) {
    if (String(data[i][0]) === identifier) {

      let games;

      try {
        games = JSON.parse(data[i][1] || "[]");
      } catch (error) {
        throw new Error("玩家資料格式錯誤，無法讀取。");
      }

      if (!Array.isArray(games)) {
        throw new Error("玩家遊戲資料格式錯誤。");
      }

      return {
        success: true,
        identifier: identifier,
        games: games
      };
    }
  }

  // 新代號沒有舊資料
  return {
    success: true,
    identifier: identifier,
    games: []
  };
}


// =================================
// 6. 儲存玩家資料
// =================================

function savePlayerData(identifier, games) {
  identifier = validateIdentifier(identifier);

  if (!Array.isArray(games)) {
    throw new Error("遊戲資料格式錯誤。");
  }

  if (games.length > 1000) {
    throw new Error("遊戲紀錄數量超出限制。");
  }

  const json = JSON.stringify(games);

  if (json.length > 40000) {
    throw new Error("遊戲資料太大，無法儲存。");
  }

  const lock = LockService.getScriptLock();

  lock.waitLock(10000);

  try {
    const sheet = getDatabase();
    const lastRow = sheet.getLastRow();

    let targetRow = -1;

    if (lastRow >= 2) {
      const identifiers = sheet
        .getRange(2, 1, lastRow - 1, 1)
        .getDisplayValues();

      for (let i = 0; i < identifiers.length; i++) {
        if (identifiers[i][0] === identifier) {
          targetRow = i + 2;
          break;
        }
      }
    }

    const updatedAt = new Date();

    if (targetRow === -1) {
      targetRow = sheet.getLastRow() + 1;
    }

    sheet.getRange(targetRow, 1, 1, 3).setValues([[
      identifier,
      json,
      updatedAt
    ]]);

    return {
      success: true,
      message: "雲端存檔成功！",
      updatedAt: updatedAt.toISOString()
    };

  } finally {
    lock.releaseLock();
  }
}
