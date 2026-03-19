/**
 * ==========================================
 * SERVER SIDE: Google Apps Script (รหัส.gs)
 * v4 — Smart Merge Engine
 * - autoDetectHeaderRow: keyword ตรงกับไฟล์จริง
 * - normalizeHeader_: ลบวงเล็บ/บาท/newline ก่อน compare
 * - getSheetMeta: คืน hidden status + detected header row
 * - mergeSelectedFiles: รองรับ defaultValues per file (True5G_Max เป็นต้น)
 * - getPreviewData: ใช้ defaultValues ด้วย
 * ==========================================
 */
const MAIN_FOLDER_ID   = "1-yY9mu_wRwCPCd2P0MYGROUVKUKGbbMW";
const TARGET_FOLDER_ID = "1rQ5Eo_oxY8yL9C7V72_Lr9CIKVWOjaaV";
const UPDATE_FOLDER_ID = "1GGcZoKsDROkFp808WcRhjNuVBB9CWy0R";

/* Template columns — ชื่อจริงตาม template */
var TEMPLATE_HEADERS = [
  "ประเภท",
  "ประเภทลูกค้าที่สามารถซื้อโปรโมชั่นนี้ได้",
  "แบรนด์และรุ่น",
  "ชื่อโปรโมชั่น",
  "รายละเอียด",
  "ช่องทางขาย",
  "Shop type",
  "Shop code",
  "ราคาปกติ (บาท)",
  "ส่วนลดค่าเครื่อง (บาท)",
  "ส่วนลดค่าเครื่องเพิ่มเติม (บาท)",
  "ส่วนลดค่าเครื่องเพิ่มเติม MNP (บาท)",
  "ราคาหลังหักส่วนลด (บาท)",
  "ค่าบริการเหมาจ่ายที่ชำระไว้ก่อน",
  "โปรโมชั่นเริ่มต้น",
  "สิทธิพิเศษเพิ่มเติม",
  "สัญญาการใช้งาน (เดือน)",
  "Start Sale date",
  "Sale to date"
];

/* คอลัมน์ส่วนลด: ถ้า cell ว่างให้เติม 0 อัตโนมัติ */
var ZERO_FILL_HEADERS = [
  "ส่วนลดค่าเครื่อง (บาท)",
  "ส่วนลดค่าเครื่องเพิ่มเติม (บาท)",
  "ส่วนลดค่าเครื่องเพิ่มเติม MNP (บาท)"
];

/* คอลัมน์ที่ใส่วันที่ปัจจุบันอัตโนมัติ (ถ้าไม่มีใน source) */
var AUTO_DATE_HEADERS = ["Start Sale date", "Sale to date"];

/* ==========================================
   AUTH & SECURITY
   ========================================== */
function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index').setTitle("AI CHAT PROMOTION");
}

function validateUser(username, password) {
  var cache      = CacheService.getScriptCache();
  var attemptKey = "login_attempts_" + username.toLowerCase();
  var lockKey    = "login_lock_" + username.toLowerCase();
  if (cache.get(lockKey)) return { success: false, error: "พยายามเข้าสู่ระบบมากเกินไป กรุณารอ 5 นาที" };

  var attempts    = parseInt(cache.get(attemptKey) || "0");
  var props       = PropertiesService.getScriptProperties();
  var correctUser = props.getProperty('APP_USERNAME') || "admin";
  var correctPass = props.getProperty('APP_PASSWORD') || "123456";

  if (username === correctUser && password === correctPass) {
    cache.remove(attemptKey);
    var token = Utilities.getUuid();
    cache.put("session_" + token, username, 28800);
    return { success: true, token: token };
  }
  attempts++;
  cache.put(attemptKey, String(attempts), 300);
  if (attempts >= 5) cache.put(lockKey, "1", 300);
  return { success: false, error: "ชื่อผู้ใช้หรือรหัสผ่านไม่ถูกต้อง" };
}

function requireAuth_(token) {
  if (!token) throw new Error("SESSION_EXPIRED");
  if (!CacheService.getScriptCache().get("session_" + token)) throw new Error("SESSION_EXPIRED");
}

function isFileInAllowedFolder_(fileId) {
  try {
    var parents = DriveApp.getFileById(fileId).getParents();
    while (parents.hasNext()) {
      var pid = parents.next().getId();
      if (pid === TARGET_FOLDER_ID || pid === MAIN_FOLDER_ID || pid === UPDATE_FOLDER_ID) return true;
    }
    return false;
  } catch (e) { return false; }
}

/* ==========================================
   FILE EXPLORATION
   ========================================== */
function getFileList(token) {
  requireAuth_(token);
  var files = DriveApp.getFolderById(TARGET_FOLDER_ID).getFilesByType(MimeType.GOOGLE_SHEETS);
  var list  = [];
  while (files.hasNext()) {
    var f = files.next();
    list.push({ name: f.getName(), id: f.getId() });
  }
  return list;
}

function getFolderList(token) {
  requireAuth_(token);
  var folders = DriveApp.getFolderById(MAIN_FOLDER_ID).getFolders();
  var list    = [{ name: "\uD83D\uDCC1 โฟลเดอร์หลัก (Root)", id: MAIN_FOLDER_ID }];
  while (folders.hasNext()) {
    var f = folders.next();
    list.push({ name: f.getName(), id: f.getId() });
  }
  return list;
}

/**
 * getSheetNamesMulti — คืน sheet list พร้อม isHidden
 */
function getSheetNamesMulti(token, fileIds) {
  requireAuth_(token);
  return fileIds.map(function(id) {
    if (!isFileInAllowedFolder_(id)) return { fileId: id, fileName: "ไม่มีสิทธิ์", sheets: [] };
    try {
      var ss = SpreadsheetApp.openById(id);
      return {
        fileId:   id,
        fileName: ss.getName(),
        sheets:   ss.getSheets().map(function(s) {
          return { name: s.getName(), rows: s.getLastRow(), isHidden: s.isSheetHidden() };
        })
      };
    } catch (e) { return { fileId: id, fileName: "เปิดไม่ได้", sheets: [] }; }
  });
}

/**
 * getSheetMeta — เรียกเมื่อ user เลือก sheet
 * คืน: detectedHeaderRow, headerScore, rawHeaders, missingTemplateKeys
 */
function getSheetMeta(token, fileId, sheetName) {
  requireAuth_(token);
  if (!isFileInAllowedFolder_(fileId)) throw new Error("ไม่มีสิทธิ์เข้าถึงไฟล์นี้");
  var ss    = SpreadsheetApp.openById(fileId);
  var sheet = ss.getSheetByName(sheetName);
  if (!sheet) throw new Error("ไม่พบ Sheet: " + sheetName);

  var detection  = autoDetectHeaderRow_(sheet);
  var hRow       = detection.row;
  var lastCol    = sheet.getLastColumn();
  var rawHeaders = lastCol > 0
    ? sheet.getRange(hRow, 1, 1, lastCol).getDisplayValues()[0].map(function(h) { return h.toString().trim(); })
    : [];

  // template col ที่ยังไม่มี mapping ใน rawHeaders (ไม่นับ date col)
  var missing = TEMPLATE_HEADERS.filter(function(th) {
    if (AUTO_DATE_HEADERS.indexOf(th) !== -1) return false;
    var normTh = normalizeHeader_(th);
    for (var i = 0; i < rawHeaders.length; i++) {
      if (normalizeHeader_(rawHeaders[i]) === normTh) return false;
    }
    return true;
  });

  return {
    detectedHeaderRow:   hRow,
    headerScore:         detection.score,
    headerDetected:      detection.detected,
    rawHeaders:          rawHeaders,
    missingTemplateKeys: missing
  };
}

/* ==========================================
   SYNC EXCEL
   ========================================== */
function syncExcelFiles(token) {
  requireAuth_(token);
  var source = DriveApp.getFolderById(MAIN_FOLDER_ID);
  var files  = source.getFilesByType(MimeType.MICROSOFT_EXCEL);
  var count  = 0, log = [];
  while (files.hasNext()) {
    var f    = files.next();
    var name = f.getName().replace(/\.xlsx$|\.xls$/i, "").replace(/[^\wก-๙ ]/g, "_");
    try {
      var existing = DriveApp.getFolderById(TARGET_FOLDER_ID).getFilesByName(name);
      while (existing.hasNext()) existing.next().setTrashed(true);
      Drive.Files.copy(
        { title: name, mimeType: MimeType.GOOGLE_SHEETS, parents: [{ id: TARGET_FOLDER_ID }] },
        f.getId(), { convert: true }
      );
      count++;
      log.push("OK: " + name);
    } catch (e) { log.push("ERR: " + name + " (" + e.message + ")"); }
  }
  return "Sync สำเร็จ " + count + " ไฟล์\n" + log.join("\n");
}

/* ==========================================
   AI LOGIC
   ========================================== */
function askAI(token, userQuestion, fileConfigs) {
  requireAuth_(token);
  var props  = PropertiesService.getScriptProperties();
  var apiKey = props.getProperty('TYPHOON_API_KEY');
  if (!apiKey || !apiKey.trim()) return "ไม่พบ TYPHOON_API_KEY กรุณาตั้งค่าใน Script Properties";

  var apiUrl = "https://api.opentyphoon.ai/v1/chat/completions";
  var context = "";
  fileConfigs.forEach(function(cfg) {
    if (!isFileInAllowedFolder_(cfg.id)) return;
    var ss = SpreadsheetApp.openById(cfg.id);
    context += "\n\nFILE: " + ss.getName();
    ss.getSheets().forEach(function(sh) {
      if (cfg.sheets && cfg.sheets.indexOf(sh.getName()) === -1) return;
      var lastRow = Math.min(sh.getLastRow(), 1500);
      var lastCol = sh.getLastColumn();
      if (!lastRow || !lastCol) return;
      context += "\nSHEET: " + sh.getName() + "\nDATA: " + JSON.stringify(sh.getRange(1,1,lastRow,lastCol).getValues());
    });
  });

  var prompt = "คุณคือ AI วิเคราะห์โปรโมชั่น\n\nข้อมูล:\n" + context +
    "\n\nคำถาม: " + userQuestion + "\n\nตอบเป็นภาษาไทย ใช้ตาราง HTML ในกรณีจำเป็น ห้ามใช้ Markdown";

  var models = ["typhoon-v2.5-30b-a3b-instruct", "typhoon-v2.1-12b-instruct", "typhoon-v2-8b-instruct"];
  var lastError = "";
  for (var i = 0; i < models.length; i++) {
    try {
      var res = UrlFetchApp.fetch(apiUrl, {
        method: "post",
        contentType: "application/json",
        headers: { "Authorization": "Bearer " + apiKey },
        payload: JSON.stringify({
          model: models[i],
          messages: [{ role: "user", content: prompt }],
          max_tokens: 8192
        }),
        muteHttpExceptions: true
      });
      var json = JSON.parse(res.getContentText());
      if (res.getResponseCode() === 200) return json.choices[0].message.content;
      lastError = json.error ? json.error.message : res.getContentText();
    } catch(e) {
      lastError = e.message;
    }
  }
  return "Error: Typhoon API ล้มเหลวทุก model — " + lastError;
}

/* ==========================================
   FOLDER MANAGEMENT
   ========================================== */
function createNewFolder(token, name) {
  requireAuth_(token);
  var f = DriveApp.getFolderById(MAIN_FOLDER_ID).createFolder(name);
  return { id: f.getId(), name: f.getName() };
}

/* ==========================================
   HEADER DETECTION (internal)
   keywords ครอบคลุมทุก col จริงในไฟล์ 12 ไฟล์
   ========================================== */
var DETECT_KEYWORDS_ = [
  "ประเภท","ลูกค้า","แบรนด์","รุ่น","รุ่นที่ร่วมรายการ",
  "ชื่อโปรโมชั่น","ชื่อโปรโมชัน","โปรโมชั่น","โปรโมชัน",
  "รายละเอียด","ราคาปกติ","ราคา","ส่วนลด","ส่วนลดค่าเครื่อง",
  "ราคาหลัง","ราคาสุทธิ","mnp","สัญญา","เดือน",
  "sale","date","ค่าบริการ","ค่าบริการเหมาจ่าย",
  "โปรโมชั่นเริ่มต้น","สิ่งที่เปลี่ยนแปลง"
];

function autoDetectHeaderRow_(sheet) {
  var lastCol  = sheet.getLastColumn();
  var scanRows = Math.min(sheet.getLastRow(), 10);
  if (!scanRows || !lastCol) return { detected: false, row: 1, score: 0 };

  var data = sheet.getRange(1, 1, scanRows, lastCol).getDisplayValues();
  var bestRow = 1, bestScore = 0;
  for (var r = 0; r < scanRows; r++) {
    var score = 0;
    for (var c = 0; c < data[r].length; c++) {
      var norm = normalizeHeader_(data[r][c]);
      if (!norm) continue;
      for (var k = 0; k < DETECT_KEYWORDS_.length; k++) {
        if (norm.indexOf(normalizeHeader_(DETECT_KEYWORDS_[k])) !== -1) { score++; break; }
      }
    }
    if (score > bestScore) { bestScore = score; bestRow = r + 1; }
  }
  return bestScore === 0
    ? { detected: false, row: 1, score: 0 }
    : { detected: true, row: bestRow, score: bestScore };
}

function autoDetectHeaderRow(token, fileId, sheetName) {
  requireAuth_(token);
  if (!isFileInAllowedFolder_(fileId)) throw new Error("ไม่มีสิทธิ์เข้าถึงไฟล์นี้");
  var ss    = SpreadsheetApp.openById(fileId);
  var sheet = ss.getSheetByName(sheetName);
  if (!sheet) return { detected: false, row: 1, score: 0 };
  return autoDetectHeaderRow_(sheet);
}

/* ==========================================
   NORMALIZE HEADER (internal)
   ลบ: newline, วงเล็บพร้อมเนื้อหา, คำว่า "บาท", /, -, .
   จากนั้นลบ space ทั้งหมด lowercase
   ========================================== */
function normalizeHeader_(str) {
  if (!str) return "";
  return str.toString()
    .toLowerCase()
    .replace(/\n/g, " ")
    .replace(/\(.*?\)/g, "")
    .replace(/\bบาท\b/g, "")
    .replace(/[\/\-\.]/g, "")
    .replace(/\s+/g, "")
    .trim();
}

/* ==========================================
   MAPPING PROFILES
   ========================================== */
var PROFILE_FILE_NAME_ = "_Mapping_Profiles";

var PROFILE_FOLDER_ID_ = "1WBdOd-eozC1mweNOGSYE-FjX9lzqRKCt";

function getOrCreateProfileSheet_() {
  var folder = DriveApp.getFolderById(PROFILE_FOLDER_ID_);
  var files  = folder.getFilesByName(PROFILE_FILE_NAME_);
  var ss;
  if (files.hasNext()) {
    ss = SpreadsheetApp.openById(files.next().getId());
  } else {
    ss = SpreadsheetApp.create(PROFILE_FILE_NAME_);
    DriveApp.getFileById(ss.getId()).moveTo(folder);
    var sh = ss.getSheets()[0];
    sh.setName("Profiles");
    sh.getRange(1,1,1,3).setValues([["ProfileName","MappingJSON","UpdatedAt"]])
      .setFontWeight("bold").setBackground("#334155").setFontColor("#ffffff");
    sh.setFrozenRows(1);
    sh.setColumnWidths(1, 3, [200,600,160]);
    return sh;
  }
  var sheet = ss.getSheetByName("Profiles") || ss.insertSheet("Profiles");
  if (sheet.getLastRow() < 1) {
    sheet.getRange(1,1,1,3).setValues([["ProfileName","MappingJSON","UpdatedAt"]])
      .setFontWeight("bold").setBackground("#334155").setFontColor("#ffffff");
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function loadMappingProfiles(token) {
  requireAuth_(token);
  try {
    var sheet   = getOrCreateProfileSheet_();
    var lastRow = sheet.getLastRow();
    if (lastRow < 2) return [];
    return sheet.getRange(2,1,lastRow-1,3).getValues().reduce(function(acc, row) {
      var name = (row[0]||"").toString().trim();
      if (!name) return acc;
      var mapping = {}; try { mapping = JSON.parse(row[1]||"{}"); } catch(e) {}
      acc.push({ name: name, mapping: mapping, updatedAt: (row[2]||"").toString() });
      return acc;
    }, []);
  } catch(e) { return []; }
}

function saveMappingProfile(token, profileName, mappingObj) {
  requireAuth_(token);
  if (!profileName || !profileName.trim()) throw new Error("กรุณาตั้งชื่อ Profile");
  var sheet    = getOrCreateProfileSheet_();
  var lastRow  = sheet.getLastRow();
  var nowStr   = Utilities.formatDate(new Date(), "GMT+7", "yyyy-MM-dd HH:mm");
  var jsonStr  = JSON.stringify(mappingObj);
  var trimName = profileName.trim();
  if (lastRow >= 2) {
    var names = sheet.getRange(2,1,lastRow-1,1).getValues();
    for (var i = 0; i < names.length; i++) {
      if (names[i][0].toString().trim() === trimName) {
        sheet.getRange(i+2,1,1,3).setValues([[trimName, jsonStr, nowStr]]);
        return { success: true, count: sheet.getLastRow() - 1 };
      }
    }
  }
  sheet.appendRow([trimName, jsonStr, nowStr]);
  return { success: true, count: sheet.getLastRow() - 1 };
}

function deleteMappingProfile(token, profileName) {
  requireAuth_(token);
  var sheet   = getOrCreateProfileSheet_();
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return { success: false };
  var names = sheet.getRange(2,1,lastRow-1,1).getValues();
  for (var i = names.length-1; i >= 0; i--) {
    if (names[i][0].toString().trim() === profileName) {
      sheet.deleteRow(i+2);
      return { success: true };
    }
  }
  return { success: false };
}

/* ==========================================
   PARALLEL DATA READER — UrlFetchApp.fetchAll()
   อ่านข้อมูลทุกไฟล์พร้อมกัน เร็วกว่า sequential 5-10x
   ========================================== */

function colToA1_(col) {
  var s = "";
  while (col > 0) { col--; s = String.fromCharCode(65 + (col % 26)) + s; col = Math.floor(col / 26); }
  return s;
}

/**
 * buildSheetsApiUrl_ — สร้าง URL สำหรับ Sheets API v4
 */
function buildSheetsApiUrl_(spreadsheetId, range, fields) {
  var url = "https://sheets.googleapis.com/v4/spreadsheets/" + spreadsheetId
          + "?ranges=" + encodeURIComponent(range)
          + "&fields=" + encodeURIComponent(fields);
  return url;
}

/**
 * parallelFetchSheetData_ — อ่านข้อมูล + formatting จากหลายไฟล์พร้อมกัน
 * คืน array ของ {values, merges, strikeRows, error} ตาม index ของ input
 */
function parallelFetchSheetData_(fileInfos) {
  var oauthToken = ScriptApp.getOAuthToken();
  var authHeaders = { "Authorization": "Bearer " + oauthToken };
  var BATCH_SIZE = 3; // 3 ไฟล์ต่อ batch (ไม่เกิน 60 req/min)
  var BATCH_DELAY = 4000; // 4 วินาทีระหว่าง batch
  var MAX_RETRIES = 3; // retry สูงสุด 3 รอบ
  var RETRY_DELAY = 5000; // 5 วินาทีต่อ retry
  var fields = "sheets.data.rowData.values(effectiveValue,formattedValue,effectiveFormat(backgroundColor,textFormat(foregroundColor,bold,strikethrough),horizontalAlignment)),sheets.merges";

  // สร้าง result placeholders
  var results = fileInfos.map(function(fi) {
    return {
      headers: [], disp: [], vals: [],
      bgs: [], fcs: [], fws: [], als: [],
      strikeSet: {}, merges: [],
      rowCount: fi.dataEndRow - fi.dataStartRow + 1,
      colCount: fi.colCount,
      error: null
    };
  });

  // สร้าง request สำหรับทุกไฟล์
  var allRequests = fileInfos.map(function(fi) {
    var safeSheet = "'" + fi.sheetName.replace(/'/g, "''") + "'";
    var lastCol = colToA1_(fi.colCount);
    var fullRange = safeSheet + "!A" + fi.headerRow + ":" + lastCol + fi.dataEndRow;
    return {
      url: buildSheetsApiUrl_(fi.spreadsheetId, fullRange, fields),
      method: "get",
      headers: authHeaders,
      muteHttpExceptions: true
    };
  });

  // pending = indices ที่ยังไม่สำเร็จ
  var pending = [];
  for (var i = 0; i < fileInfos.length; i++) pending.push(i);

  for (var attempt = 0; attempt < MAX_RETRIES && pending.length > 0; attempt++) {
    // แบ่ง pending เป็น batch
    for (var batchStart = 0; batchStart < pending.length; batchStart += BATCH_SIZE) {
      if (attempt > 0 || batchStart > 0) Utilities.sleep(attempt === 0 ? BATCH_DELAY : RETRY_DELAY);

      var batchEnd = Math.min(batchStart + BATCH_SIZE, pending.length);
      var batchRequests = [];
      var batchPendingIdx = [];
      for (var b = batchStart; b < batchEnd; b++) {
        batchRequests.push(allRequests[pending[b]]);
        batchPendingIdx.push(pending[b]);
      }

      var responses = UrlFetchApp.fetchAll(batchRequests);

      responses.forEach(function(resp, ri) {
        var fileIdx = batchPendingIdx[ri];
        var result = results[fileIdx];
        var fi = fileInfos[fileIdx];
        try {
          var code = resp.getResponseCode();
          if (code === 429) return; // ยังอยู่ใน pending ลอง retry รอบถัดไป
          if (code !== 200) { result.error = "HTTP " + code; return; }
          var json = JSON.parse(resp.getContentText());
          parseSheetResponse_(json, result, fi);
          result.error = null; // สำเร็จ
        } catch(e) {
          result.error = e.message;
        }
      });
    }

    // อัปเดต pending: เก็บเฉพาะที่ยังไม่มี data + ไม่มี error ถาวร (ยกเว้น 429)
    var newPending = [];
    for (var p = 0; p < pending.length; p++) {
      var idx = pending[p];
      if (results[idx].disp.length === 0 && results[idx].headers.length === 0 && !results[idx].error) {
        newPending.push(idx); // 429 → retry
      }
    }
    pending = newPending;
  }

  // ถ้ายัง pending อยู่ ให้ mark error
  pending.forEach(function(idx) {
    if (!results[idx].error) results[idx].error = "Rate limit (retry หมดแล้ว)";
  });

  return results;
}

/**
 * parseSheetResponse_ — parse JSON response จาก Sheets API v4
 */
function parseSheetResponse_(json, result, fi) {
  if (json.sheets && json.sheets[0] && json.sheets[0].merges) {
    result.merges = json.sheets[0].merges;
  }
  var sheetData = json.sheets && json.sheets[0] && json.sheets[0].data && json.sheets[0].data[0];
  if (!sheetData || !sheetData.rowData) return;
  var rowData = sheetData.rowData;

  if (rowData[0] && rowData[0].values) {
    result.headers = rowData[0].values.map(function(c) {
      return (c && c.formattedValue) ? c.formattedValue.toString().trim() : "";
    });
  }

  for (var r = 1; r < rowData.length; r++) {
    var dRow=[], vRow=[], bRow=[], cRow=[], wRow=[], aRow=[];
    var hasStrike = false;

    if (!rowData[r] || !rowData[r].values) {
      for (var c = 0; c < fi.colCount; c++) {
        dRow.push(""); vRow.push(""); bRow.push("#ffffff");
        cRow.push("#000000"); wRow.push("normal"); aRow.push("left");
      }
    } else {
      var cells = rowData[r].values;
      for (var c = 0; c < fi.colCount; c++) {
        var cell = (c < cells.length) ? cells[c] : null;
        if (!cell) {
          dRow.push(""); vRow.push(""); bRow.push("#ffffff");
          cRow.push("#000000"); wRow.push("normal"); aRow.push("left");
          continue;
        }
        dRow.push(cell.formattedValue || "");
        var ev = cell.effectiveValue;
        if (ev) {
          vRow.push(ev.numberValue !== undefined ? ev.numberValue :
                    ev.stringValue !== undefined ? ev.stringValue :
                    ev.boolValue !== undefined ? ev.boolValue :
                    cell.formattedValue || "");
        } else { vRow.push(cell.formattedValue || ""); }

        var ef = cell.effectiveFormat || {};
        var tf = ef.textFormat || {};
        var bg = ef.backgroundColor;
        if (bg) {
          var rr = Math.round((bg.red||0)*255), gg = Math.round((bg.green||0)*255), bb = Math.round((bg.blue||0)*255);
          bRow.push("#" + ((1<<24)+(rr<<16)+(gg<<8)+bb).toString(16).slice(1));
        } else { bRow.push("#ffffff"); }
        var fc = tf.foregroundColor;
        if (fc) {
          var rr = Math.round((fc.red||0)*255), gg = Math.round((fc.green||0)*255), bb2 = Math.round((fc.blue||0)*255);
          cRow.push("#" + ((1<<24)+(rr<<16)+(gg<<8)+bb2).toString(16).slice(1));
        } else { cRow.push("#000000"); }
        wRow.push(tf.bold ? "bold" : "normal");
        aRow.push((ef.horizontalAlignment || "LEFT").toLowerCase());
        if (tf.strikethrough) hasStrike = true;
      }
    }
    result.disp.push(dRow);
    result.vals.push(vRow);
    result.bgs.push(bRow);
    result.fcs.push(cRow);
    result.fws.push(wRow);
    result.als.push(aRow);
    if (hasStrike) result.strikeSet[r - 1] = true;
  }
}

/**
 * applyMerges_ — fill merged cells จาก Sheets API merge info
 * merges format: [{startRowIndex, endRowIndex, startColumnIndex, endColumnIndex}]
 */
function applyMerges_(result, dataStartRow) {
  if (!result.merges || !result.merges.length) return;
  result.merges.forEach(function(mg) {
    // convert sheet-absolute index → data-relative index
    var sr = mg.startRowIndex - (dataStartRow - 1);
    var er = mg.endRowIndex - (dataStartRow - 1) - 1;
    var sc = mg.startColumnIndex;
    var ec = mg.endColumnIndex - 1;
    if (er < 0 || sr >= result.rowCount) return;
    if (sr < 0) sr = 0;

    var bv = result.vals[sr] ? result.vals[sr][sc] : "";
    var bd = result.disp[sr] ? result.disp[sr][sc] : "";
    var bb = result.bgs[sr]  ? result.bgs[sr][sc]  : "#ffffff";
    var bf = result.fcs[sr]  ? result.fcs[sr][sc]  : "#000000";
    var bw = result.fws[sr]  ? result.fws[sr][sc]  : "normal";
    var ba = result.als[sr]  ? result.als[sr][sc]  : "left";

    for (var r = sr; r <= er && r < result.rowCount; r++) {
      for (var c = sc; c <= ec && c < result.colCount; c++) {
        if (result.vals[r]) result.vals[r][c] = bv;
        if (result.disp[r]) result.disp[r][c] = bd;
        if (result.bgs[r])  result.bgs[r][c] = bb;
        if (result.fcs[r])  result.fcs[r][c] = bf;
        if (result.fws[r])  result.fws[r][c] = bw;
        if (result.als[r])  result.als[r][c] = ba;
      }
    }
  });
}

/* ==========================================
   MERGE ENGINE HELPERS
   ========================================== */

/**
 * resolveValue_
 * ลำดับ: mapping[th] → direct match by name → defaultValues[th]
 */
function resolveValue_(templateCol, mapping, defaultValues, sourceHeaders, sourceRow) {
  var candidates = mapping[templateCol];
  if (!candidates || (Array.isArray(candidates) && !candidates.length)) candidates = [templateCol];
  if (!Array.isArray(candidates)) candidates = [candidates];

  for (var ci = 0; ci < candidates.length; ci++) {
    if (!candidates[ci]) continue;
    var normC = normalizeHeader_(candidates[ci]);
    for (var hi = 0; hi < sourceHeaders.length; hi++) {
      if (normalizeHeader_(sourceHeaders[hi]) === normC) {
        var v = sourceRow[hi];
        if (v !== undefined && v !== null && v !== "") return { value: v, found: true };
      }
    }
  }
  // fallback default
  var def = defaultValues && defaultValues[templateCol];
  if (def !== undefined && def !== null && def !== "") return { value: def, found: false };
  // คอลัมน์ส่วนลด: ว่าง → 0
  if (ZERO_FILL_HEADERS.indexOf(templateCol) !== -1) return { value: 0, found: false };
  return { value: "", found: false };
}

function buildOutputRow_(masterHeaders, special, autoDateHeaders, mapping, defaultValues,
                          block, r, disp, bgs, fcs, fws, als, nowStr, sourceName) {
  var vRow=[], bRow=[], cRow=[], wRow=[], aRow=[];
  masterHeaders.forEach(function(h) {

    // ---- คอลัมน์ [ที่มา] ----
    if (h === "[ที่มา]") {
      vRow.push(sourceName || "");
      bRow.push("#e0f2fe"); cRow.push("#0369a1"); wRow.push("normal"); aRow.push("left");
      return;
    }

    // ---- คอลัมน์ AUTO_DATE ----
    if (special.indexOf(h) !== -1) {
      vRow.push(autoDateHeaders.indexOf(h) !== -1 ? nowStr : "");
      bRow.push("#ffffff"); cRow.push("#000000"); wRow.push("normal"); aRow.push("left");
      return;
    }

    var isTemplateCol = TEMPLATE_HEADERS.indexOf(h) !== -1;
    var sIdx = -1;
    var val  = "";

    if (isTemplateCol) {
      // ---- คอลัมน์ template: ใช้ mapping ----
      var res = resolveValue_(h, mapping, defaultValues, block.headers, disp[r]);
      val = res.value !== undefined ? res.value : "";
      var cands = mapping[h];
      if (!cands || (Array.isArray(cands) && !cands.length)) cands = [h];
      if (!Array.isArray(cands)) cands = [cands];
      for (var ci = 0; ci < cands.length && sIdx === -1; ci++) {
        var normC = normalizeHeader_(cands[ci]);
        for (var hi = 0; hi < block.headers.length; hi++) {
          if (normalizeHeader_(block.headers[hi]) === normC) { sIdx = hi; break; }
        }
      }
    } else {
      // ---- คอลัมน์ extra (ไม่ได้ map): match ตรงๆ ด้วยชื่อ ----
      var normH = normalizeHeader_(h);
      for (var hi = 0; hi < block.headers.length; hi++) {
        if (normalizeHeader_(block.headers[hi]) === normH) {
          sIdx = hi;
          val  = (disp[r][hi] !== undefined && disp[r][hi] !== null) ? disp[r][hi] : "";
          break;
        }
      }
    }

    vRow.push(val);
    bRow.push(sIdx !== -1 ? bgs[r][sIdx] : "#ffffff");
    cRow.push(sIdx !== -1 ? fcs[r][sIdx] : "#000000");
    wRow.push(sIdx !== -1 ? fws[r][sIdx] : "normal");
    aRow.push(sIdx !== -1 ? als[r][sIdx] : "left");
  });
  return { v: vRow, b: bRow, c: cRow, w: wRow, a: aRow };
}

function applyStyles_(destSheet, startRow, colCount, vals, bgs, fcs, fws, als) {
  var rng = destSheet.getRange(startRow, 1, vals.length, colCount);
  rng.setValues(vals).setBackgrounds(bgs).setFontColors(fcs)
     .setFontWeights(fws).setHorizontalAlignments(als)
     .setBorder(true,true,true,true,true,true,"#e2e8f0",SpreadsheetApp.BorderStyle.SOLID);
}

/* ==========================================
   PREVIEW DATA
   ========================================== */
function getPreviewData(token, configs, manualMapping) {
  requireAuth_(token);
  var mapping      = manualMapping || {};
  var nowStr       = Utilities.formatDate(new Date(), "GMT+7", "yyyy-MM-dd");
  var previewRows  = [];
  var totalSkipped = 0;
  var totalEmpty   = 0;
  var totalNoBrand = 0;
  var totalSource  = 0;
  var MAX_TOTAL    = 50, MAX_PER_FILE = 15;

  // PHASE 1: Collect file info
  var fileInfos = [];
  configs.forEach(function(cfg) {
    if (!isFileInAllowedFolder_(cfg.fileId)) return;
    try {
      var ss    = SpreadsheetApp.openById(cfg.fileId);
      var sheet = ss.getSheetByName(cfg.sheetName);
      if (!sheet || sheet.isSheetHidden()) return;
      var lastRow = sheet.getLastRow();
      var lastCol = sheet.getLastColumn();
      var hRow    = cfg.headerRow || 1;
      if (lastRow <= hRow || !lastCol) return;

      var scanRows = Math.min(lastRow - hRow, MAX_PER_FILE + 10);
      totalSource += (lastRow - hRow);

      fileInfos.push({
        cfg: cfg,
        spreadsheetId: cfg.fileId,
        sheetName: cfg.sheetName,
        headerRow: hRow,
        dataStartRow: hRow + 1,
        dataEndRow: hRow + scanRows,
        colCount: lastCol
      });
    } catch(e) { console.error("Preview discovery [" + cfg.fileName + "]: " + e.message); }
  });

  if (!fileInfos.length) return { headers: TEMPLATE_HEADERS.concat(["[ที่มา]"]), rows: [], skipped: 0, emptyRows: 0, noBrand: 0, totalSource: 0, total: 0 };

  // PHASE 2: Parallel fetch
  var fetched = parallelFetchSheetData_(fileInfos);

  // PHASE 3: Process rows
  var brandIdx = TEMPLATE_HEADERS.indexOf("แบรนด์และรุ่น");

  fetched.forEach(function(result, idx) {
    if (previewRows.length >= MAX_TOTAL) return;
    if (result.error) return;
    var fi  = fileInfos[idx];
    var cfg = fi.cfg;
    var cfgMapping    = cfg.mapping || mapping;
    var defaultValues = cfg.defaultValues || {};
    var fileAdded     = 0;

    // Apply merged cells
    applyMerges_(result, fi.dataStartRow);

    for (var r = 0; r < result.disp.length && fileAdded < MAX_PER_FILE && previewRows.length < MAX_TOTAL; r++) {
      if (result.strikeSet[r]) { totalSkipped++; continue; }
      var allEmpty = result.disp[r].every(function(v){ return !v || !v.toString().trim(); });
      if (allEmpty) { totalEmpty++; continue; }

      var outRow = TEMPLATE_HEADERS.map(function(th) {
        if (AUTO_DATE_HEADERS.indexOf(th) !== -1) return nowStr;
        return resolveValue_(th, cfgMapping, defaultValues, result.headers, result.disp[r]).value;
      });

      // เช็ค "แบรนด์และรุ่น" ว่าง → ลบ row ทิ้ง
      if (brandIdx !== -1) {
        var brandVal = outRow[brandIdx];
        if (!brandVal || !brandVal.toString().trim()) { totalNoBrand++; continue; }
      }

      outRow.push(cfg.fileName + " / " + cfg.sheetName);
      previewRows.push(outRow);
      fileAdded++;
    }
  });

  return {
    headers:     TEMPLATE_HEADERS.concat(["[ที่มา]"]),
    rows:        previewRows,
    skipped:     totalSkipped,
    emptyRows:   totalEmpty,
    noBrand:     totalNoBrand,
    totalSource: totalSource,
    total:       previewRows.length
  };
}

/* ==========================================
   MERGE SELECTED FILES
   cfg per sheet: { fileId, fileName, sheetName, headerRow,
                    mapping: {th: [srcCol,...]},
                    defaultValues: {th: "value"} }
   ========================================== */
function mergeSelectedFiles(token, configs, newFileName, targetFolderId) {
  requireAuth_(token);
  if (!newFileName || !newFileName.trim()) return JSON.stringify({ error: "กรุณาระบุชื่อไฟล์ใหม่" });

  var nowStr        = Utilities.formatDate(new Date(), "GMT+7", "yyyy-MM-dd");
  var masterHeaders = TEMPLATE_HEADERS.slice();
  var special       = AUTO_DATE_HEADERS.slice();
  var errorFiles    = [];

  try {
    // PHASE 1: Discovery — หา sheet metadata (ใช้ SpreadsheetApp เบาๆ แค่ header + row/col count)
    var fileInfos = [];
    configs.forEach(function(cfg) {
      if (!isFileInAllowedFolder_(cfg.fileId)) return;
      try {
        var ss    = SpreadsheetApp.openById(cfg.fileId);
        var sheet = ss.getSheetByName(cfg.sheetName);
        if (!sheet || sheet.isSheetHidden()) return;

        var lastRow = sheet.getLastRow();
        var lastCol = sheet.getLastColumn();
        var hRow    = cfg.headerRow || 1;
        if (!lastRow || !lastCol || lastRow <= hRow) return;

        fileInfos.push({
          cfg: cfg,
          spreadsheetId: cfg.fileId,
          sheetName: cfg.sheetName,
          headerRow: hRow,
          dataStartRow: hRow + 1,
          dataEndRow: lastRow,
          colCount: lastCol
        });
      } catch(e) {
        errorFiles.push({ name: cfg.fileName + " / " + cfg.sheetName, reason: e.message });
      }
    });

    if (!fileInfos.length) return JSON.stringify({ error: "ไม่พบข้อมูลในไฟล์ที่เลือก" });

    // PHASE 2: Parallel fetch — อ่านทุกไฟล์พร้อมกัน
    var fetched = parallelFetchSheetData_(fileInfos);

    // Build masterHeaders from fetched headers + apply merges
    fetched.forEach(function(result, idx) {
      if (result.error) {
        errorFiles.push({ name: fileInfos[idx].cfg.fileName + " / " + fileInfos[idx].cfg.sheetName, reason: result.error });
        return;
      }

      // Apply merged cells
      applyMerges_(result, fileInfos[idx].dataStartRow);

      // Extra cols discovery
      var cfgMap = fileInfos[idx].cfg.mapping || {};
      result.headers.forEach(function(h) {
        if (!h) return;
        var normH = normalizeHeader_(h);
        var isMapped = Object.keys(cfgMap).some(function(th) {
          var vals = cfgMap[th];
          if (!Array.isArray(vals)) vals = [vals];
          return vals.some(function(v) { return v && normalizeHeader_(v) === normH; });
        });
        if (!isMapped && !masterHeaders.some(function(mh){ return normalizeHeader_(mh) === normH; })) {
          masterHeaders.push(h);
        }
      });
    });

    // แทรก [ที่มา] ต่อจาก TEMPLATE_HEADERS (ก่อน extra cols)
    masterHeaders.splice(TEMPLATE_HEADERS.length, 0, "[ที่มา]");

    // PHASE 3: สร้างไฟล์
    var newSS  = SpreadsheetApp.create(newFileName);
    var dest   = targetFolderId ? DriveApp.getFolderById(targetFolderId) : DriveApp.getFolderById(MAIN_FOLDER_ID);
    DriveApp.getFileById(newSS.getId()).moveTo(dest);

    var destSheet = newSS.getSheets()[0];
    destSheet.setName("Combined_Data");
    destSheet.getRange(1,1,1,masterHeaders.length)
      .setValues([masterHeaders]).setFontWeight("bold")
      .setBackground("#334155").setFontColor("#ffffff")
      .setHorizontalAlignment("center").setVerticalAlignment("middle");
    destSheet.setFrozenRows(1);

    // PHASE 4: Process + batch write
    var allV=[], allB=[], allC=[], allW=[], allA=[];
    var totalRows = 0, totalSourceRows = 0;
    var strikeRows = 0, emptyRows = 0, noBrandRows = 0;

    fetched.forEach(function(result, idx) {
      if (result.error) return;
      var fi  = fileInfos[idx];
      var cfg = fi.cfg;
      var mapping       = cfg.mapping || {};
      var defaultValues = cfg.defaultValues || {};
      var block = { headers: result.headers, colCount: result.colCount };

      totalSourceRows += result.rowCount;

      for (var r = 0; r < result.disp.length; r++) {
        // skip strikethrough
        if (result.strikeSet[r]) { strikeRows++; continue; }
        // skip empty
        var allEmpty = result.disp[r].every(function(v){ return !v || !v.toString().trim(); });
        if (allEmpty) { emptyRows++; continue; }

        var sourceName = cfg.fileName + " / " + cfg.sheetName;
        var row = buildOutputRow_(masterHeaders, special, AUTO_DATE_HEADERS,
                                   mapping, defaultValues,
                                   block, r, result.disp, result.bgs, result.fcs,
                                   result.fws, result.als, nowStr, sourceName);

        // เช็ค "แบรนด์และรุ่น" ว่าง → ลบ row ทิ้ง
        var brandMIdx = masterHeaders.indexOf("แบรนด์และรุ่น");
        if (brandMIdx !== -1) {
          var bv = row.v[brandMIdx];
          if (!bv || !bv.toString().trim()) { noBrandRows++; continue; }
        }

        allV.push(row.v); allB.push(row.b); allC.push(row.c);
        allW.push(row.w); allA.push(row.a);
      }
    });

    // Batch write ทีเดียว
    totalRows = allV.length;
    if (totalRows > 0) {
      applyStyles_(destSheet, 2, masterHeaders.length, allV, allB, allC, allW, allA);
      destSheet.getRange(1,1,totalRows+1,masterHeaders.length).setVerticalAlignment("middle");
    }
    destSheet.autoResizeColumns(1, masterHeaders.length);

    var msg = "รวมสำเร็จ: " + totalRows + " แถว  |  ไฟล์: " + newFileName;
    if (strikeRows > 0 || emptyRows > 0 || noBrandRows > 0) {
      msg += "\n📊 สรุป: จากทั้งหมด " + totalSourceRows + " แถว";
      if (strikeRows  > 0) msg += " | ลบขีดฆ่า " + strikeRows + " แถว";
      if (noBrandRows > 0) msg += " | ไม่มีแบรนด์/รุ่น " + noBrandRows + " แถว";
      if (emptyRows   > 0) msg += " | ข้ามแถวว่าง " + emptyRows + " แถว";
    }
    if (errorFiles.length > 0) {
      msg += "\n⚠️ ข้ามไฟล์ที่มีปัญหา " + errorFiles.length + " ไฟล์:";
      errorFiles.forEach(function(ef) { msg += "\n   - " + ef.name + " (" + ef.reason + ")"; });
    }

    return JSON.stringify({
      message:     msg,
      url:         newSS.getUrl(),
      fileName:    newFileName,
      rows:        totalRows,
      totalSource: totalSourceRows,
      strikeRows:  strikeRows,
      noBrandRows: noBrandRows,
      emptyRows:   emptyRows,
      errorFiles:  errorFiles
    });
  } catch(e) { return JSON.stringify({ error: "ผิดพลาด: " + e.message }); }
}

/* ==========================================
   getDiscoveryHeaders — ใช้ใน mapping page
   ========================================== */
function getDiscoveryHeaders(token, configs) {
  requireAuth_(token);
  var result = [];
  configs.forEach(function(cfg) {
    try {
      var ss    = SpreadsheetApp.openById(cfg.fileId);
      var sheet = ss.getSheetByName(cfg.sheetName);
      if (!sheet) return;
      var lastCol = sheet.getLastColumn();
      if (!lastCol) return;
      var hRow    = cfg.headerRow || 1;
      var headers = sheet.getRange(hRow,1,1,lastCol).getDisplayValues()[0]
                         .map(function(h){ return h.toString().trim(); })
                         .filter(function(h){ return !!h; });
      result.push({ fileId: cfg.fileId, fileName: cfg.fileName, sheetName: cfg.sheetName, headerRow: hRow, headers: headers });
    } catch(e) { console.error("Discovery: " + e.message); }
  });
  return result;
}

/* ==========================================
   PDF EXTRACT ENGINE
   ==========================================
   Step 1: OCR.space  → แปลง PDF เป็น text (ฟรี 500 req/วัน)
   Step 2: Gemini     → แปลง text เป็น structured JSON
   Fallback: ส่ง PDF ตรงไป Gemini ถ้า OCR ล้มเหลว
   - savePdfExtractedData : บันทึกข้อมูลที่ user ยืนยันแล้วลง Google Sheet
   ========================================== */

var PDF_OUTPUT_FOLDER_ID = "1vq3IEaIwJXwUGeQoDEOuVrl0OU64orJb";

/* Mapping จาก short key (Gemini JSON) → Template column */
var PDF_FIELD_KEYS_ = [
  "type","customer","brand","promo","detail",
  "normalPrice","discount","extraDiscount","netPrice",
  "mnpDiscount","advancePayment","campaign","contract",
  "startDate","endDate"
];

/* ── OCR.space helper (single JPEG image) ── */
function ocrImagePage_(base64Data, ocrKey) {
  var url = "https://api.ocr.space/parse/image";
  var res = UrlFetchApp.fetch(url, {
    method: "post",
    payload: {
      base64Image:       "data:image/jpeg;base64," + base64Data,
      apikey:            ocrKey,
      language:          "tha",
      isOverlayRequired: "false",
      filetype:          "JPG",
      OCREngine:         "2",
      scale:             "true",
      isTable:           "true"
    },
    muteHttpExceptions: true
  });

  var code = res.getResponseCode();
  var json = JSON.parse(res.getContentText());

  if (code !== 200) return { error: "OCR.space HTTP " + code };
  if (json.IsErroredOnProcessing) return { error: (json.ErrorMessage || JSON.stringify(json.ErrorDetails)) };
  if (!json.ParsedResults || json.ParsedResults.length === 0) return { error: "ไม่พบข้อความ" };

  return { text: json.ParsedResults[0].ParsedText || "" };
}

/* ── OCR.space helper (full PDF — legacy) ── */
function ocrPdfPages_(base64Data, ocrKey) {
  var url = "https://api.ocr.space/parse/image";
  var res = UrlFetchApp.fetch(url, {
    method: "post",
    payload: {
      base64Image:       "data:application/pdf;base64," + base64Data,
      apikey:            ocrKey,
      language:          "tha",          // ภาษาไทย
      isOverlayRequired: "false",
      filetype:          "PDF",
      OCREngine:         "2",            // Engine 2 รองรับภาษาไทยดีกว่า
      scale:             "true",
      isTable:           "true"          // รักษาโครงสร้างตาราง
    },
    muteHttpExceptions: true
  });

  var code = res.getResponseCode();
  var json = JSON.parse(res.getContentText());

  if (code !== 200) return { error: "OCR.space HTTP " + code };
  if (json.IsErroredOnProcessing) return { error: "OCR error: " + (json.ErrorMessage || JSON.stringify(json.ErrorDetails)) };
  if (!json.ParsedResults || json.ParsedResults.length === 0) return { error: "OCR ไม่พบข้อความใน PDF" };

  // รวม text ทุกหน้า พร้อมระบุเลขหน้า
  var allText = "";
  for (var i = 0; i < json.ParsedResults.length; i++) {
    allText += "\n=== หน้า " + (i + 1) + " ===\n" + (json.ParsedResults[i].ParsedText || "");
  }
  return { text: allText.trim(), pages: json.ParsedResults.length };
}

/* ── Gemini: text → structured JSON ── */
function typhoonTextToPromo_(apiKey, ocrText, pageRange) {
  var pageInst = (!pageRange || pageRange === "all")
    ? "ทุกหน้า"
    : "เฉพาะหน้า " + pageRange + " เท่านั้น (ข้ามหน้าอื่น)";

  var systemMsg =
    'คุณคือ AI ที่เชี่ยวชาญการสกัดข้อมูลโปรโมชั่นจากเอกสาร PDF ของ True/dtac\n' +
    'คุณจะได้รับข้อความที่ OCR มาจาก PDF แล้วต้องสกัดข้อมูลโปรโมชั่นเป็น JSON array\n' +
    'ตอบเป็น JSON array เท่านั้น ห้ามมี markdown, ห้ามมีข้อความอื่นนอก JSON';

  var userMsg =
    'ข้อความด้านล่างนี้ถูก OCR มาจาก PDF (' + pageInst + ')\n' +
    'สกัดข้อมูลโปรโมชั่น/แพ็กเกจทั้งหมดที่พบในตาราง\n' +
    'คืนผลเป็น JSON array โดยแต่ละ object มี key ดังนี้:\n\n' +
    '- "type"           : ประเภท เช่น 2P, 4POTT, STL, FTTH\n' +
    '- "customer"       : ประเภทลูกค้า เช่น ลูกค้าใหม่, ย้ายค่าย, ซิมรายเดือน\n' +
    '- "brand"          : ชื่อแคมเปญหลัก เช่น True Fiber My Plan, Combo Max, Security\n' +
    '- "promo"          : MKT Code หรือชื่อโปรโมชั่นเฉพาะ เช่น FTTS203-1000\n' +
    '- "detail"         : รายละเอียด เช่น ความเร็ว, ซิมเน็ต, CCTV, ความบันเทิง, อุปกรณ์\n' +
    '- "normalPrice"    : ราคาปกติก่อนลด (บาท/เดือน) ถ้ามี\n' +
    '- "discount"       : ส่วนลดค่าเครื่องหรือส่วนลดค่าบริการ (บาท)\n' +
    '- "extraDiscount"  : ส่วนลดเพิ่มเติมอื่นๆ\n' +
    '- "netPrice"       : ราคาที่ลูกค้าจ่ายจริง (บาท/เดือน)\n' +
    '- "mnpDiscount"    : ส่วนลดย้ายค่าย (บาท)\n' +
    '- "advancePayment" : ค่าบริการเหมาจ่ายที่ชำระไว้ก่อน\n' +
    '- "campaign"       : Campaign Name / Profile ที่ใช้ออกออเดอร์ เช่น Join us get more\n' +
    '- "contract"       : ระยะสัญญา (เดือน) เช่น 12, 24\n' +
    '- "startDate"      : วันเริ่มขาย (ถ้ามี)\n' +
    '- "endDate"        : วันสิ้นสุด (ถ้ามี)\n' +
    '- "page"           : หมายเลขหน้าใน PDF ที่ดึงข้อมูลมา\n\n' +
    'กฎสำคัญ:\n' +
    '- ถ้าข้อมูลไม่มีหรือไม่เกี่ยวข้อง ให้ใส่ "" (string ว่าง)\n' +
    '- แต่ละ object = 1 รายการโปรโมชั่นที่แตกต่างกัน (ต่าง MKT Code / ต่างราคา / ต่างความเร็ว)\n' +
    '- ถ้า 1 หน้ามีหลายแพ็กเกจ/หลายราคา ให้แยกเป็นหลาย object\n' +
    '- ข้ามหน้าที่ไม่มีตารางโปรโมชั่น (หน้าปก, สารบัญ, ขั้นตอนการขาย, รูปภาพ, flow chart)\n' +
    '- ตอบเป็น JSON array เท่านั้น ห้ามมี markdown, ห้ามมีข้อความอื่นนอก JSON\n\n' +
    '=== OCR TEXT ===\n' + ocrText;

  var apiUrl = "https://api.opentyphoon.ai/v1/chat/completions";

  // ลอง model ทีละตัว (ใหม่สุดก่อน)
  var models = [
    "typhoon-v2.5-30b-a3b-instruct",
    "typhoon-v2.1-12b-instruct",
    "typhoon-v2-8b-instruct"
  ];

  var lastError = "";
  for (var m = 0; m < models.length; m++) {
    var modelName = models[m];
    var payload = {
      model: modelName,
      messages: [
        { role: "system", content: systemMsg },
        { role: "user",   content: userMsg }
      ],
      temperature: 0.1,
      max_tokens: 8192
    };

    for (var attempt = 1; attempt <= 2; attempt++) {
      try {
        var res = UrlFetchApp.fetch(apiUrl, {
          method: "post",
          contentType: "application/json",
          headers: { "Authorization": "Bearer " + apiKey },
          payload: JSON.stringify(payload),
          muteHttpExceptions: true
        });
        var code = res.getResponseCode();
        var body = res.getContentText();

        if (code === 200) {
          var json = JSON.parse(body);
          var text = json.choices[0].message.content;
          text = text.replace(/^```(?:json)?\s*/i, "").replace(/\s*```\s*$/, "").trim();
          var rows = JSON.parse(text);
          if (!Array.isArray(rows)) rows = [rows];
          return { rows: rows, total: rows.length, model: modelName, method: "OCR+Typhoon" };
        }
        if (code === 400) {
          // Model not found → ลอง model ถัดไป
          lastError = "HTTP 400 [" + modelName + "]: " + body.substring(0, 200);
          break;
        }
        if (code === 429 && attempt < 2) { Utilities.sleep(5000); continue; }
        lastError = "HTTP " + code + " [" + modelName + "]: " + body.substring(0, 200);
        break;
      } catch(e) {
        lastError = modelName + ": " + e.message;
        if (attempt < 2) { Utilities.sleep(3000); continue; }
        break;
      }
    }
  }
  return { error: "Typhoon API ล้มเหลวทุก model — " + lastError };
}


/**
 * extractPdfPageImages  (New — receives page images from client pdf.js)
 * images = [{ pageNum: 3, base64: "..." }, { pageNum: 4, base64: "..." }]
 */
function extractPdfPageImages(token, images, pageRange) {
  requireAuth_(token);
  var props  = PropertiesService.getScriptProperties();
  var apiKey = props.getProperty('TYPHOON_API_KEY');
  if (!apiKey || !apiKey.trim()) return { error: "ไม่พบ TYPHOON_API_KEY กรุณาตั้งค่าใน Script Properties" };

  var ocrKey = props.getProperty('OCR_API_KEY') || 'helloworld';

  if (!images || images.length === 0) return { error: "ไม่ได้รับรูปภาพจาก PDF" };

  // OCR each page image individually via OCR.space
  var allText   = "";
  var ocrErrors = [];

  for (var i = 0; i < images.length; i++) {
    var img = images[i];
    try {
      var ocrResult = ocrImagePage_(img.base64, ocrKey);
      if (ocrResult.error) {
        ocrErrors.push("หน้า " + img.pageNum + ": " + ocrResult.error);
      } else if (ocrResult.text) {
        allText += "\n=== หน้า " + img.pageNum + " ===\n" + ocrResult.text;
      }
    } catch(e) {
      ocrErrors.push("หน้า " + img.pageNum + ": " + e.message);
    }
    // หน่วงเล็กน้อยระหว่าง request เพื่อไม่ให้โดน rate-limit
    if (i < images.length - 1) Utilities.sleep(500);
  }

  if (!allText || allText.trim().length < 30) {
    return { error: "OCR ไม่สามารถอ่านข้อความจากรูปภาพได้\n" + ocrErrors.join("\n") };
  }

  // ส่ง OCR text ไป Typhoon เพื่อแปลงเป็น structured JSON
  var result = typhoonTextToPromo_(apiKey, allText.trim(), pageRange);
  if (result.error) return result;

  result.ocrPages  = images.length;
  result.ocrErrors = ocrErrors;
  result.method    = "OCR+Typhoon (Images)";
  return result;
}

/**
 * savePdfExtractedData
 * รับ JSON array (หลังจาก user แก้ไขใน preview) → สร้าง Google Sheet ใน PDF_OUTPUT_FOLDER_ID
 */
function savePdfExtractedData(token, rowsJson, newFileName) {
  requireAuth_(token);
  try {
    var rows = JSON.parse(rowsJson);
    if (!rows || rows.length === 0) return { error: "ไม่มีข้อมูลที่จะบันทึก" };
    if (!newFileName || !newFileName.trim()) return { error: "กรุณาระบุชื่อไฟล์" };

    var headers = TEMPLATE_HEADERS.slice();

    var newSS = SpreadsheetApp.create(newFileName.trim());
    var dest  = DriveApp.getFolderById(PDF_OUTPUT_FOLDER_ID);
    DriveApp.getFileById(newSS.getId()).moveTo(dest);

    var sheet = newSS.getSheets()[0];
    sheet.setName("Extracted_Data");

    /* Header row */
    sheet.getRange(1, 1, 1, headers.length)
      .setValues([headers]).setFontWeight("bold")
      .setBackground("#334155").setFontColor("#ffffff")
      .setHorizontalAlignment("center").setVerticalAlignment("middle");
    sheet.setFrozenRows(1);

    /* Map short keys → template columns */
    var dataRows = rows.map(function(row) {
      return headers.map(function(h, idx) {
        var key = PDF_FIELD_KEYS_[idx];
        if (!key) return "";
        var val = row[key];
        return (val !== undefined && val !== null) ? val.toString() : "";
      });
    });

    if (dataRows.length > 0) {
      sheet.getRange(2, 1, dataRows.length, headers.length).setValues(dataRows);
    }

    sheet.autoResizeColumns(1, headers.length);

    return {
      message:  "บันทึกสำเร็จ: " + dataRows.length + " แถว | ไฟล์: " + newFileName,
      url:      newSS.getUrl(),
      fileName: newFileName,
      rows:     dataRows.length
    };
  } catch(e) { return { error: "บันทึกล้มเหลว: " + e.message }; }
}

/**
 * saveTyphoonApiDoc_  — สร้าง Google Doc เก็บข้อมูล API config ไว้ใน PDF output folder
 */
function saveTyphoonApiDoc_() {
  var doc = DocumentApp.create("Typhoon API Config");
  var body = doc.getBody();

  body.appendParagraph("Typhoon API Configuration").setHeading(DocumentApp.ParagraphHeading.HEADING1);
  body.appendParagraph("อัปเดตล่าสุด: " + new Date().toLocaleString("th-TH"));
  body.appendParagraph("");

  body.appendParagraph("API Endpoint").setHeading(DocumentApp.ParagraphHeading.HEADING2);
  body.appendParagraph("URL: https://api.opentyphoon.ai/v1/chat/completions");
  body.appendParagraph("Method: POST");
  body.appendParagraph("Auth: Bearer token (ใน header Authorization)");
  body.appendParagraph("");

  body.appendParagraph("Model ที่ใช้").setHeading(DocumentApp.ParagraphHeading.HEADING2);
  body.appendParagraph("typhoon-v2.1-12b-instruct (56K context, Thai optimized)");
  body.appendParagraph("temperature: 0.1 | max_tokens: 8192");
  body.appendParagraph("");

  body.appendParagraph("API Key Location").setHeading(DocumentApp.ParagraphHeading.HEADING2);
  body.appendParagraph("เก็บไว้ใน Script Properties ชื่อ: TYPHOON_API_KEY");
  body.appendParagraph("ห้าม hardcode ใน source code!");
  body.appendParagraph("");

  body.appendParagraph("Rate Limits (Free Tier)").setHeading(DocumentApp.ParagraphHeading.HEADING2);
  body.appendParagraph("5 requests/sec | 200 requests/min");
  body.appendParagraph("");

  body.appendParagraph("Flow การทำงาน").setHeading(DocumentApp.ParagraphHeading.HEADING2);
  body.appendParagraph("1. Client: pdf.js แปลง PDF → JPEG images (ทีละหน้า)");
  body.appendParagraph("2. Server: OCR.space อ่านรูปภาพ → text");
  body.appendParagraph("3. Server: Typhoon API แปลง text → structured JSON (โปรโมชั่น)");
  body.appendParagraph("4. Client: แสดง editable preview → user แก้ไข → บันทึกเป็น Google Sheet");

  doc.saveAndClose();

  var file = DriveApp.getFileById(doc.getId());
  var dest = DriveApp.getFolderById(PDF_OUTPUT_FOLDER_ID);
  file.moveTo(dest);

  return { url: doc.getUrl(), name: "Typhoon API Config" };
}