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
  "ประเภทลูกค้าที่สามารถซื้อโปโมชั่นนี้ได้",
  "แบรนด์และรุ่น",
  "ชื่อโปรโมชั่น",
  "รายละเอียด",
  "ราคาปกติ (บาท)",
  "ส่วนลดค่าเครื่อง (บาท)",
  "ส่วนลดค่าเครื่องเพิ่มเติม",
  "ราคาหลังหักส่วนลด (บาท)",
  "ส่วนลดค่าเครื่องเพิ่มเติม MNP (บาท)",
  "ค่าบริการเหมาจ่ายที่ชำระไว้ก่อน",
  "โปรโมชั่นเริ่มต้น",
  "สัญญาการใช้งาน (เดือน)",
  "Start Sale date",
  "Sale to date"
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
  var apiKey = props.getProperty('GEMINI_API_KEY');
  if (!apiKey || !apiKey.trim()) return "ไม่พบ GEMINI_API_KEY กรุณาตั้งค่าใน Script Properties";

  var apiUrl  = "https://generativelanguage.googleapis.com/v1beta/models/gemini-2.0-flash:generateContent?key=" + apiKey;
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
  var res  = UrlFetchApp.fetch(apiUrl, {
    method: "post", contentType: "application/json",
    payload: JSON.stringify({ contents: [{ parts: [{ text: prompt }] }] }),
    muteHttpExceptions: true
  });
  var json = JSON.parse(res.getContentText());
  if (res.getResponseCode() === 200) return json.candidates[0].content.parts[0].text;
  return "Error: " + (json.error ? json.error.message : "AI ไม่ตอบสนอง");
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
   STRIKETHROUGH DETECTION — ใช้ Sheets API v4
   ต้องเปิด Advanced Service: Google Sheets API (ชื่อ "Sheets")
   ========================================== */

/**
 * colToA1_ — แปลง column index (1-based) เป็น A1 notation letter
 */
function colToA1_(col) {
  var s = "";
  while (col > 0) {
    col--;
    s = String.fromCharCode(65 + (col % 26)) + s;
    col = Math.floor(col / 26);
  }
  return s;
}

/**
 * getStrikeRowSet_ — ใช้ Sheets API v4 อ่าน effectiveFormat.textFormat.strikethrough
 * คืน object { 0: true, 3: true, ... } สำหรับ row ที่มี strikethrough (0-based relative to data start)
 */
function getStrikeRowSet_(spreadsheetId, sheetName, dataStartRow, dataEndRow, colCount) {
  var strikeSet = {};
  try {
    var safeSheet = "'" + sheetName.replace(/'/g, "''") + "'";
    var lastColLetter = colToA1_(colCount);
    var a1 = safeSheet + "!A" + dataStartRow + ":" + lastColLetter + dataEndRow;

    var resp = Sheets.Spreadsheets.get(spreadsheetId, {
      ranges: [a1],
      fields: "sheets.data.rowData.values.effectiveFormat.textFormat.strikethrough"
    });

    var sheetData = resp.sheets && resp.sheets[0] && resp.sheets[0].data && resp.sheets[0].data[0];
    if (!sheetData || !sheetData.rowData) return strikeSet;
    var rowData = sheetData.rowData;

    for (var r = 0; r < rowData.length; r++) {
      if (!rowData[r] || !rowData[r].values) continue;
      var cells = rowData[r].values;
      for (var c = 0; c < cells.length; c++) {
        if (cells[c] &&
            cells[c].effectiveFormat &&
            cells[c].effectiveFormat.textFormat &&
            cells[c].effectiveFormat.textFormat.strikethrough === true) {
          strikeSet[r] = true;
          break;  // พบ strikethrough ใน row นี้แล้ว ไม่ต้องเช็ค col ถัดไป
        }
      }
    }
  } catch (e) {
    console.error("getStrikeRowSet_ error: " + e.message);
    // fallback: ลองใช้ getFontLines แบบเดิม (ถ้า Sheets API ยังไม่เปิด)
  }
  return strikeSet;
}

/**
 * flattenMergedDisp_ — เวอร์ชันเบาสำหรับ preview (flatten เฉพาะ display values)
 */
function flattenMergedDisp_(srcRange, disp, rowCount, colCount) {
  srcRange.getMergedRanges().forEach(function(mr) {
    var sr = mr.getRow() - srcRange.getRow();
    var er = sr + mr.getNumRows() - 1;
    var sc = mr.getColumn() - 1;
    var ec = sc + mr.getNumColumns() - 1;
    var bd = disp[sr][sc];
    for (var r = sr; r <= er; r++) {
      for (var c = sc; c <= ec; c++) {
        if (r < rowCount && c < colCount) {
          disp[r][c] = bd;
        }
      }
    }
  });
}

/* ==========================================
   MERGE ENGINE HELPERS
   ========================================== */
function flattenMergedCells_(srcRange, vals, disp, bgs, fcs, fws, als, rowCount, colCount) {
  srcRange.getMergedRanges().forEach(function(mr) {
    var sr = mr.getRow() - srcRange.getRow();
    var er = sr + mr.getNumRows() - 1;
    var sc = mr.getColumn() - 1;
    var ec = sc + mr.getNumColumns() - 1;
    var bv = vals[sr][sc], bd = disp[sr][sc], bb = bgs[sr][sc];
    var bf = fcs[sr][sc],  bw = fws[sr][sc],  ba = als[sr][sc];
    for (var r = sr; r <= er; r++) {
      for (var c = sc; c <= ec; c++) {
        if (r < rowCount && c < colCount) {
          vals[r][c]=bv; disp[r][c]=bd; bgs[r][c]=bb;
          fcs[r][c]=bf;  fws[r][c]=bw;  als[r][c]=ba;
        }
      }
    }
  });
}

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
  var totalNoBrand = 0; // แบรนด์และรุ่น ว่าง
  var totalSource  = 0;  // จำนวน row ทั้งหมดก่อนกรอง
  var MAX_TOTAL    = 50, MAX_PER_FILE = 15;

  configs.forEach(function(cfg) {
    if (previewRows.length >= MAX_TOTAL) return;
    if (!isFileInAllowedFolder_(cfg.fileId)) return;
    var defaultValues = cfg.defaultValues || {};
    var cfgMapping    = cfg.mapping || mapping;

    try {
      var ss    = SpreadsheetApp.openById(cfg.fileId);
      var sheet = ss.getSheetByName(cfg.sheetName);
      if (!sheet || sheet.isSheetHidden()) return;

      var lastRow = sheet.getLastRow();
      var lastCol = sheet.getLastColumn();
      var hRow    = cfg.headerRow || 1;
      if (lastRow <= hRow || !lastCol) return;

      totalSource += (lastRow - hRow);

      var headers   = sheet.getRange(hRow,1,1,lastCol).getDisplayValues()[0].map(function(h){return h.toString().trim();});
      var scanRows  = Math.min(lastRow - hRow, MAX_PER_FILE + 10);
      var dataRange = sheet.getRange(hRow+1,1,scanRows,lastCol);
      var dispVals  = dataRange.getDisplayValues();

      // flatten merged cells → เติมค่าให้ cell ที่ merge ไว้
      flattenMergedDisp_(dataRange, dispVals, scanRows, lastCol);

      // ใช้ Sheets API v4 ตรวจ strikethrough
      var strikeSet = getStrikeRowSet_(cfg.fileId, cfg.sheetName, hRow+1, hRow+scanRows, lastCol);
      var fileAdded = 0;

      for (var r = 0; r < scanRows && fileAdded < MAX_PER_FILE && previewRows.length < MAX_TOTAL; r++) {
        if (strikeSet[r]) { totalSkipped++; continue; }
        var allEmpty = dispVals[r].every(function(v){ return !v || !v.toString().trim(); });
        if (allEmpty) { totalEmpty++; continue; }

        var outRow = TEMPLATE_HEADERS.map(function(th) {
          if (AUTO_DATE_HEADERS.indexOf(th) !== -1) return nowStr;
          return resolveValue_(th, cfgMapping, defaultValues, headers, dispVals[r]).value;
        });

        // เช็ค "แบรนด์และรุ่น" ว่าง → ลบ row ทิ้ง
        var brandIdx = TEMPLATE_HEADERS.indexOf("แบรนด์และรุ่น");
        if (brandIdx !== -1) {
          var brandVal = outRow[brandIdx];
          if (!brandVal || !brandVal.toString().trim()) { totalNoBrand++; continue; }
        }

        outRow.push(cfg.fileName + " / " + cfg.sheetName);
        previewRows.push(outRow);
        fileAdded++;
      }
    } catch(e) { console.error("Preview [" + cfg.fileName + "]: " + e.message); }
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

  var nowStr       = Utilities.formatDate(new Date(), "GMT+7", "yyyy-MM-dd");
  var masterHeaders = TEMPLATE_HEADERS.slice();
  var special       = AUTO_DATE_HEADERS.slice();
  var sourceBlocks  = [];

  try {
    // PHASE 1: Discovery
    configs.forEach(function(cfg) {
      if (!isFileInAllowedFolder_(cfg.fileId)) return;
      var ss    = SpreadsheetApp.openById(cfg.fileId);
      var sheet = ss.getSheetByName(cfg.sheetName);
      if (!sheet || sheet.isSheetHidden()) return;

      var lastRow = sheet.getLastRow();
      var lastCol = sheet.getLastColumn();
      var hRow    = cfg.headerRow || 1;
      if (!lastRow || !lastCol) return;

      var rawH = sheet.getRange(hRow,1,1,lastCol).getDisplayValues()[0].map(function(h){return h.toString().trim();});

      // เพิ่ม extra col ต่อท้าย master เฉพาะที่ยังไม่ได้ map ไป template ใดๆ
      var cfgMap = cfg.mapping || {};
      rawH.forEach(function(h) {
        if (!h) return;
        var normH = normalizeHeader_(h);
        // ถ้าถูก map เป็น value ของ template col ใดๆ ให้ข้าม (จะถูกใส่ใน template col นั้นแล้ว)
        var isMapped = Object.keys(cfgMap).some(function(th) {
          var vals = cfgMap[th];
          if (!Array.isArray(vals)) vals = [vals];
          return vals.some(function(v) { return v && normalizeHeader_(v) === normH; });
        });
        if (!isMapped && !masterHeaders.some(function(mh){ return normalizeHeader_(mh) === normH; })) {
          masterHeaders.push(h);
        }
      });

      if (lastRow > hRow) {
        sourceBlocks.push({
          fileConfig:    cfg,
          headers:       rawH,
          mapping:       cfg.mapping || {},
          defaultValues: cfg.defaultValues || {},
          range:         sheet.getRange(hRow+1, 1, lastRow-hRow, lastCol),
          rowCount:      lastRow - hRow,
          colCount:      lastCol,
          spreadsheetId: cfg.fileId,
          sheetName:     cfg.sheetName,
          dataStartRow:  hRow + 1,
          dataEndRow:    lastRow
        });
      }
    });

    if (!sourceBlocks.length) return JSON.stringify({ error: "ไม่พบข้อมูลในไฟล์ที่เลือก" });

    // แทรก [ที่มา] ต่อจาก TEMPLATE_HEADERS (ก่อน extra cols)
    masterHeaders.splice(TEMPLATE_HEADERS.length, 0, "[ที่มา]");

    // PHASE 2: สร้างไฟล์
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

    // PHASE 3: Copy data
    var currentRow = 2, totalRows = 0, totalSourceRows = 0;
    var strikeRows = 0, emptyRows = 0, noBrandRows = 0;
    sourceBlocks.forEach(function(block) {
      totalSourceRows += block.rowCount;
      var srcRange  = block.range;
      var vals      = srcRange.getValues();
      var disp      = srcRange.getDisplayValues();
      var bgs       = srcRange.getBackgrounds();
      var fcs       = srcRange.getFontColors();
      var fws       = srcRange.getFontWeights();
      var als       = srcRange.getHorizontalAlignments();

      flattenMergedCells_(srcRange, vals, disp, bgs, fcs, fws, als, block.rowCount, block.colCount);

      // ใช้ Sheets API v4 ตรวจ strikethrough
      var strikeSet = getStrikeRowSet_(block.spreadsheetId, block.sheetName, block.dataStartRow, block.dataEndRow, block.colCount);

      var outV=[], outB=[], outC=[], outW=[], outA=[];
      for (var r = 0; r < block.rowCount; r++) {
        if (strikeSet[r]) { strikeRows++; continue; }
        var allEmpty = disp[r].every(function(v){ return !v || !v.toString().trim(); });
        if (allEmpty) { emptyRows++; continue; }

        var sourceName = block.fileConfig.fileName + " / " + block.fileConfig.sheetName;
        var row = buildOutputRow_(masterHeaders, special, AUTO_DATE_HEADERS,
                                   block.mapping, block.defaultValues,
                                   block, r, disp, bgs, fcs, fws, als, nowStr, sourceName);

        // เช็ค "แบรนด์และรุ่น" ว่าง → ลบ row ทิ้ง
        var brandMIdx = masterHeaders.indexOf("แบรนด์และรุ่น");
        if (brandMIdx !== -1) {
          var bv = row.v[brandMIdx];
          if (!bv || !bv.toString().trim()) { noBrandRows++; continue; }
        }

        outV.push(row.v); outB.push(row.b); outC.push(row.c);
        outW.push(row.w); outA.push(row.a);
      }
      if (outV.length) {
        applyStyles_(destSheet, currentRow, masterHeaders.length, outV, outB, outC, outW, outA);
        currentRow += outV.length;
        totalRows  += outV.length;
      }
    });

    if (currentRow > 2) destSheet.getRange(1,1,currentRow-1,masterHeaders.length).setVerticalAlignment("middle");
    destSheet.autoResizeColumns(1, masterHeaders.length);

    var msg = "รวมสำเร็จ: " + totalRows + " แถว  |  ไฟล์: " + newFileName;
    if (strikeRows > 0 || emptyRows > 0 || noBrandRows > 0) {
      msg += "\n📊 สรุป: จากทั้งหมด " + totalSourceRows + " แถว";
      if (strikeRows  > 0) msg += " | ลบขีดฆ่า " + strikeRows + " แถว";
      if (noBrandRows > 0) msg += " | ไม่มีแบรนด์/รุ่น " + noBrandRows + " แถว";
      if (emptyRows   > 0) msg += " | ข้ามแถวว่าง " + emptyRows + " แถว";
    }

    return JSON.stringify({
      message:     msg,
      url:         newSS.getUrl(),
      fileName:    newFileName,
      rows:        totalRows,
      totalSource: totalSourceRows,
      strikeRows:  strikeRows,
      noBrandRows: noBrandRows,
      emptyRows:   emptyRows
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
   UPDATE ENGINE
   ==========================================
   - getUpdateFileList    : ดึงไฟล์จาก UPDATE_FOLDER_ID
   - getCombinedFileList  : ดึงไฟล์จาก MAIN + TARGET (สำหรับ target picker)
   - analyzeUpdateRows    : เปรียบเทียบ source vs target
                            → duplicate / newRows / conflicts / invalids
   - applyNewRows         : เพิ่มแถวใหม่เข้า target
   - saveFailedRows       : บันทึก failed rows ไปยัง sheet ที่กำหนด
   ========================================== */

function getUpdateFileList(token) {
  requireAuth_(token);
  var files = DriveApp.getFolderById(UPDATE_FOLDER_ID).getFilesByType(MimeType.GOOGLE_SHEETS);
  var list  = [];
  while (files.hasNext()) {
    var f = files.next();
    list.push({ name: f.getName(), id: f.getId() });
  }
  return list;
}

function getCombinedFileList(token) {
  requireAuth_(token);
  var list = [], seen = {};
  var addFromFolder = function(folderId) {
    try {
      var files = DriveApp.getFolderById(folderId).getFilesByType(MimeType.GOOGLE_SHEETS);
      while (files.hasNext()) {
        var f = files.next();
        if (!seen[f.getId()]) { seen[f.getId()] = true; list.push({ name: f.getName(), id: f.getId() }); }
      }
    } catch(e) {}
  };
  addFromFolder(MAIN_FOLDER_ID);
  addFromFolder(TARGET_FOLDER_ID);
  return list;
}

/**
 * analyzeUpdateRows
 * เปรียบเทียบ source sheet กับ target sheet แล้วแบ่งแถวออกเป็น 4 กลุ่ม:
 *   duplicate  — แถวที่เหมือนกันทุก col (ข้าม)
 *   newRows    — แถวใหม่ที่ยังไม่มีใน target (พร้อมเพิ่ม)
 *   conflicts  — มี key เดียวกัน (ชื่อโปรโมชั่น + แบรนด์) แต่ข้อมูลต่างกัน
 *   invalids   — ไม่มีข้อมูลใน key column
 */
function analyzeUpdateRows(token, srcFileId, srcSheetName, tgtFileId, tgtSheetName) {
  requireAuth_(token);
  try {
    var srcSS    = SpreadsheetApp.openById(srcFileId);
    var srcSheet = srcSS.getSheetByName(srcSheetName);
    if (!srcSheet) return { error: 'ไม่พบ Sheet ต้นฉบับ: ' + srcSheetName };

    var srcLastRow = srcSheet.getLastRow();
    var srcLastCol = srcSheet.getLastColumn();
    if (srcLastRow < 2 || !srcLastCol)
      return { duplicates: 0, newRowsCount: 0, conflicts: [], invalids: [], srcHeaders: [] };

    var srcHeaders = srcSheet.getRange(1, 1, 1, srcLastCol).getDisplayValues()[0]
                             .map(function(h){ return h.toString().trim(); });
    var srcData    = srcSheet.getRange(2, 1, srcLastRow - 1, srcLastCol).getDisplayValues();

    var tgtSS    = SpreadsheetApp.openById(tgtFileId);
    var tgtSheet = tgtSS.getSheetByName(tgtSheetName);
    if (!tgtSheet) return { error: 'ไม่พบ Sheet ปลายทาง: ' + tgtSheetName };

    var tgtLastRow = tgtSheet.getLastRow();
    var tgtLastCol = tgtSheet.getLastColumn();
    var tgtData    = (tgtLastRow > 1 && tgtLastCol > 0)
      ? tgtSheet.getRange(2, 1, tgtLastRow - 1, tgtLastCol).getDisplayValues() : [];

    /* Build exact-match set (ใช้ \x00 เป็น delimiter) */
    var tgtRowSet = {};
    tgtData.forEach(function(row) {
      tgtRowSet[ row.map(function(v){ return (v||'').toString().trim(); }).join('\x00') ] = true;
    });

    /* หา index ของ key columns ใน source */
    var KEY_COLS  = ['ชื่อโปรโมชั่น', 'แบรนด์และรุ่น'];
    var srcKeyIdx = KEY_COLS.map(function(k) {
      for (var i = 0; i < srcHeaders.length; i++) {
        if (normalizeHeader_(srcHeaders[i]) === normalizeHeader_(k)) return i;
      }
      return -1;
    });

    /* Build partial-key set จาก target (เพื่อตรวจ conflict) */
    var tgtKeySet = {};
    tgtData.forEach(function(row) {
      var pk = srcKeyIdx.map(function(ki){
        return ki >= 0 && ki < row.length ? (row[ki]||'').toString().trim() : '';
      }).join('\x00');
      if (pk.replace(/\x00/g,'').trim()) tgtKeySet[pk] = true;
    });

    var duplicates = 0, newRowsCount = 0, conflicts = [], invalids = [];
    var MAX_REPORT = 200;

    srcData.forEach(function(row, idx) {
      var allEmpty = row.every(function(v){ return !v || !v.toString().trim(); });
      if (allEmpty) return;

      var fullKey = row.map(function(v){ return (v||'').toString().trim(); }).join('\x00');
      var pk      = srcKeyIdx.map(function(ki){
        return ki >= 0 && ki < row.length ? (row[ki]||'').toString().trim() : '';
      }).join('\x00');
      var hasKey  = srcKeyIdx.some(function(ki){ return ki >= 0 && row[ki] && row[ki].toString().trim(); });

      if (tgtRowSet[fullKey]) {
        duplicates++;
      } else if (!hasKey) {
        if (invalids.length < MAX_REPORT)
          invalids.push({ rowNum: idx + 2, data: row.map(function(v){ return (v||'').toString(); }),
                          reason: 'ไม่มีข้อมูลหลัก (ชื่อโปรโมชั่น / แบรนด์และรุ่น)' });
      } else if (tgtKeySet[pk]) {
        if (conflicts.length < MAX_REPORT)
          conflicts.push({ rowNum: idx + 2, data: row.map(function(v){ return (v||'').toString(); }),
                           reason: 'มีข้อมูลหลักซ้ำแต่รายละเอียดต่างกัน' });
      } else {
        newRowsCount++;
      }
    });

    return {
      duplicates:   duplicates,
      newRowsCount: newRowsCount,
      conflicts:    conflicts,
      invalids:     invalids,
      srcHeaders:   srcHeaders
    };
  } catch(e) { return { error: e.message }; }
}

/**
 * applyNewRows — เพิ่มเฉพาะแถวที่เป็น "new" เข้า target sheet
 */
function applyNewRows(token, srcFileId, srcSheetName, tgtFileId, tgtSheetName) {
  requireAuth_(token);
  try {
    var srcSS    = SpreadsheetApp.openById(srcFileId);
    var srcSheet = srcSS.getSheetByName(srcSheetName);
    var tgtSS    = SpreadsheetApp.openById(tgtFileId);
    var tgtSheet = tgtSS.getSheetByName(tgtSheetName);

    var srcLastRow = srcSheet.getLastRow();
    var srcLastCol = srcSheet.getLastColumn();
    if (srcLastRow < 2 || !srcLastCol) return { added: 0 };

    var srcHeaders = srcSheet.getRange(1, 1, 1, srcLastCol).getDisplayValues()[0];
    var srcData    = srcSheet.getRange(2, 1, srcLastRow - 1, srcLastCol).getDisplayValues();

    var tgtLastRow = tgtSheet.getLastRow();
    var tgtLastCol = tgtSheet.getLastColumn();
    var tgtData    = (tgtLastRow > 1 && tgtLastCol > 0)
      ? tgtSheet.getRange(2, 1, tgtLastRow - 1, tgtLastCol).getDisplayValues() : [];

    var tgtRowSet = {};
    tgtData.forEach(function(row) {
      tgtRowSet[ row.map(function(v){ return (v||'').toString().trim(); }).join('\x00') ] = true;
    });

    var KEY_COLS  = ['ชื่อโปรโมชั่น', 'แบรนด์และรุ่น'];
    var srcKeyIdx = KEY_COLS.map(function(k) {
      for (var i = 0; i < srcHeaders.length; i++) {
        if (normalizeHeader_(srcHeaders[i]) === normalizeHeader_(k)) return i;
      }
      return -1;
    });

    var tgtKeySet = {};
    tgtData.forEach(function(row) {
      var pk = srcKeyIdx.map(function(ki){
        return ki >= 0 && ki < row.length ? (row[ki]||'').toString().trim() : '';
      }).join('\x00');
      if (pk.replace(/\x00/g,'').trim()) tgtKeySet[pk] = true;
    });

    var toAdd = [];
    srcData.forEach(function(row) {
      var allEmpty = row.every(function(v){ return !v || !v.toString().trim(); });
      if (allEmpty) return;

      var fullKey = row.map(function(v){ return (v||'').toString().trim(); }).join('\x00');
      if (tgtRowSet[fullKey]) return;

      var pk     = srcKeyIdx.map(function(ki){
        return ki >= 0 && ki < row.length ? (row[ki]||'').toString().trim() : '';
      }).join('\x00');
      var hasKey = srcKeyIdx.some(function(ki){ return ki >= 0 && row[ki] && row[ki].toString().trim(); });

      if (tgtKeySet[pk] || !hasKey) return;
      toAdd.push(row);
    });

    if (toAdd.length === 0) return { added: 0 };

    /* เพิ่ม header row ถ้า target ว่างเปล่า */
    if (tgtLastRow === 0) {
      tgtSheet.appendRow(srcHeaders);
      tgtLastRow = 1;
    }

    /* Batch setValues แทน appendRow ทีละแถว */
    tgtSheet.getRange(tgtSheet.getLastRow() + 1, 1, toAdd.length, srcLastCol).setValues(toAdd);

    return { added: toAdd.length };
  } catch(e) { return { error: e.message }; }
}

/**
 * saveFailedRows — บันทึก conflict/invalid rows ไปยัง sheet ที่กำหนด
 * สร้าง sheet ใหม่ถ้ายังไม่มี, เพิ่ม header row อัตโนมัติ
 */
function saveFailedRows(token, failedRowsJson, destFileId, destSheetName, headersJson) {
  requireAuth_(token);
  try {
    var rows    = JSON.parse(failedRowsJson);
    var headers = JSON.parse(headersJson || '[]');
    if (!rows || rows.length === 0) return { saved: 0 };

    var ss    = SpreadsheetApp.openById(destFileId);
    var sheet = ss.getSheetByName(destSheetName);
    if (!sheet) sheet = ss.insertSheet(destSheetName);

    if (sheet.getLastRow() === 0 && headers.length > 0) {
      var hdrs = headers.concat(['เหตุผลที่ไม่ได้อัปเดท', 'แถวต้นฉบับ (Row#)']);
      sheet.appendRow(hdrs);
      sheet.getRange(1, 1, 1, hdrs.length)
           .setFontWeight('bold').setBackground('#334155').setFontColor('#ffffff');
      sheet.setFrozenRows(1);
    }

    var rowsToWrite = rows.map(function(r) {
      return r.data.concat([ r.reason || '', 'Row ' + (r.rowNum || '') ]);
    });

    var startRow = sheet.getLastRow() + 1;
    var colCount = rowsToWrite[0].length;
    sheet.getRange(startRow, 1, rowsToWrite.length, colCount).setValues(rowsToWrite);

    return { saved: rows.length, url: ss.getUrl() };
  } catch(e) { return { error: e.message }; }
}