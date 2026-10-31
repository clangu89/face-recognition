/**
 * ==============================================================================
 * Web Application: Face Recognition & Voice Greeting (ระบบจดจำใบหน้าและทักทายด้วยเสียง)
 * Platform: Google Apps Script + Google Sheets + Google Drive
 * ==============================================================================
 */

// รหัสเชื่อมต่อฐานข้อมูล Google Sheets โดยตรงเพื่อความเร็วสูงสุด (ไม่ค้นหา Drive)
const DB_SPREADSHEET_ID = "15iYZLwqkoWcw4yPBOy9TysfaKd0MfrmZR_JgDKl_VEE";
const SHEET_PROFILES = "FaceProfiles";
const SHEET_LOGS = "ScanLogs";
const SHEET_USERS = "SystemUsers";
const SHEET_MESSAGES = "ProfileMessages";
const SHEET_SETTINGS = "SystemSettings";
const SHEET_ATTENDANCE = "AttendanceLogs";
const DRIVE_FOLDER_NAME = "FaceRecognition_Photos";
const DB_SPREADSHEET_NAME = "AI_Face_Recognition_Database";

/**
 * ให้บริการหน้าเว็บ Web App
 */
function doGet(e) {
  // รองรับการเรียกแบบ REST API สำหรับ Frontend ภายนอก (เช่น GitHub Pages)
  if (e && e.parameter && e.parameter.action) {
    var action = e.parameter.action;
    var result = { success: false, message: "Invalid action" };
    if (action === "getProfiles") {
      result = getFaceProfiles();
    } else if (action === "init") {
      result = initSheetAndFolder();
    } else if (action === "getUsers") {
      result = getSystemUsers();
    } else if (action === "getMessages") {
      result = getProfileMessages();
    } else if (action === "saveMessage") {
      result = saveProfileMessage(e.parameter);
    } else if (action === "deleteMessage") {
      result = deleteProfileMessage(e.parameter.id || e.parameter.messageId);
    } else if (action === "clearMessages") {
      result = clearProfileMessages(e.parameter.profileId, e.parameter.personName);
    } else if (action === "getSettings") {
      result = getSystemSettings();
    } else if (action === "saveSettings") {
      result = saveSystemSettings(e.parameter);
    } else if (action === "getTodayAttendance") {
      result = getTodayAttendance(e.parameter.date || e.parameter.targetDate);
    } else if (action === "recordAttendance") {
      result = recordAttendanceLog(e.parameter);
    } else if (action === "updateProfileType") {
      var gId = e.parameter.id || e.parameter.profileId;
      var gType = e.parameter.personType;
      var gName = e.parameter.name || e.parameter.personName;
      result = updateProfileType(gId, gType, gName);
    } else if (action === "getRealtimeLogs") {
      result = getRealtimeDetectionLogs(e.parameter);
    } else if (action === "clearRealtimeLogs") {
      result = clearRealtimeDetectionLogs();
    } else if (action === "triggerGlobalReload") {
      result = triggerGlobalForceReload(e.parameter);
    } else if (action === "clientPulse") {
      result = syncClientPresenceAndCommands(e.parameter);
    }
    return ContentService.createTextOutput(JSON.stringify(result))
      .setMimeType(ContentService.MimeType.JSON);
  }

  // ส่งมอบหน้าจอ Web App HTML
  var htmlOutput;
  try {
    htmlOutput = HtmlService.createHtmlOutputFromFile("index");
  } catch (e) {
    htmlOutput = HtmlService.createHtmlOutputFromFile("Index");
  }
  return htmlOutput
    .setTitle("ระบบจดจำใบหน้าและทักทายด้วยเสียง (Face Recognition & Greeting)")
    .addMetaTag("viewport", "width=device-width, initial-scale=1, maximum-scale=1, user-scalable=no")
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/**
 * ฟังก์ชัน Include สำหรับนำเข้าไฟล์ HTML ย่อย
 */
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

/**
 * ดึง Google Spreadsheet ความเร็วสูง (เปิดตรงผ่าน ID ไม่ต้อง Scan Drive)
 */
function getDatabaseSpreadsheet() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    if (ss) return ss;
  } catch (e) {}

  if (DB_SPREADSHEET_ID) {
    try {
      return SpreadsheetApp.openById(DB_SPREADSHEET_ID);
    } catch (e) {
      console.warn("Open by ID failed: " + e.message);
    }
  }

  return SpreadsheetApp.create(DB_SPREADSHEET_NAME);
}

/**
 * ตรวจสอบและสร้างโครงสร้าง Google Sheet และ Google Drive Folder ครั้งแรก
 */
function initSheetAndFolder() {
  try {
    var ss = getDatabaseSpreadsheet();

    // 1. ตรวจสอบ Sheet ข้อมูลโปรไฟล์ใบหน้า (FaceProfiles)
    var profileSheet = ss.getSheetByName(SHEET_PROFILES);
    if (!profileSheet) {
      profileSheet = ss.insertSheet(SHEET_PROFILES);
      profileSheet.appendRow([
        "Profile_ID", 
        "Full_Name", 
        "Descriptor_Data", 
        "Drive_Image_URL", 
        "Drive_File_ID", 
        "Created_At", 
        "Last_Seen_At",
        "Person_Type"
      ]);
      profileSheet.getRange(1, 1, 1, 8)
        .setBackground("#1E293B")
        .setFontColor("#FFFFFF")
        .setFontWeight("bold");
      profileSheet.setFrozenRows(1);
    } else {
      // ตรวจสอบว่ามีคอลัมน์ Person_Type หรือยัง
      if (profileSheet.getLastColumn() < 8) {
        profileSheet.getRange(1, 8).setValue("Person_Type")
          .setBackground("#1E293B")
          .setFontColor("#FFFFFF")
          .setFontWeight("bold");
      }
    }

    // 2. ตรวจสอบ Sheet ประวัติการสแกน (ScanLogs)
    var logSheet = ss.getSheetByName(SHEET_LOGS);
    var logHeaders = ["Timestamp", "Person_Name", "Match_Status", "Confidence", "Person_Type", "Detected_By_User", "Detected_By_Name", "Detail"];
    if (!logSheet) {
      logSheet = ss.insertSheet(SHEET_LOGS);
      logSheet.appendRow(logHeaders);
      logSheet.getRange(1, 1, 1, logHeaders.length)
        .setBackground("#1E293B")
        .setFontColor("#FFFFFF")
        .setFontWeight("bold");
      logSheet.setFrozenRows(1);
    } else if (logSheet.getLastColumn() < logHeaders.length) {
      logSheet.getRange(1, 1, 1, logHeaders.length).setValues([logHeaders])
        .setBackground("#1E293B")
        .setFontColor("#FFFFFF")
        .setFontWeight("bold");
    }

    // 3. ตรวจสอบ Sheet ข้อมูลผู้ใช้งานระบบ (SystemUsers)
    var userSheet = ss.getSheetByName(SHEET_USERS);
    if (!userSheet) {
      userSheet = ss.insertSheet(SHEET_USERS);
      userSheet.appendRow([
        "User_ID", 
        "Username", 
        "Password_Hash", 
        "Full_Name", 
        "Role", 
        "Status", 
        "Created_At", 
        "Last_Login_At"
      ]);
      userSheet.getRange(1, 1, 1, 8)
        .setBackground("#1E293B")
        .setFontColor("#FFFFFF")
        .setFontWeight("bold");
      userSheet.setFrozenRows(1);

      // สร้างบัญชี Admin เริ่มต้น (admin / admin1234)
      var nowStr = Utilities.formatDate(new Date(), "GMT+7", "dd/MM/yyyy HH:mm:ss");
      userSheet.appendRow([
        "USR_1",
        "admin",
        hashPassword("admin1234"),
        "ผู้ดูแลระบบ (Admin)",
        "admin",
        "ACTIVE",
        nowStr,
        "-"
      ]);
    }

    // 4. ตรวจสอบ Sheet กล่องฝากข้อความประจำบุคคล (ProfileMessages)
    var msgSheet = ss.getSheetByName(SHEET_MESSAGES);
    if (!msgSheet) {
      msgSheet = ss.insertSheet(SHEET_MESSAGES);
      msgSheet.appendRow([
        "Message_ID", 
        "Profile_ID", 
        "Person_Name", 
        "Message_Text", 
        "Created_At"
      ]);
      msgSheet.getRange(1, 1, 1, 5)
        .setBackground("#1E293B")
        .setFontColor("#FFFFFF")
        .setFontWeight("bold");
      msgSheet.setFrozenRows(1);
    }

    // 5. ตรวจสอบ Sheet การตั้งค่าระบบและค่าคงที่ (SystemSettings)
    var settingsSheet = ss.getSheetByName(SHEET_SETTINGS);
    if (!settingsSheet) {
      settingsSheet = ss.insertSheet(SHEET_SETTINGS);
      settingsSheet.appendRow([
        "Setting_Key",
        "Setting_Value",
        "Description",
        "Updated_At"
      ]);
      settingsSheet.getRange(1, 1, 1, 4)
        .setBackground("#1E293B")
        .setFontColor("#FFFFFF")
        .setFontWeight("bold");
      settingsSheet.setFrozenRows(1);

      var nowStr = Utilities.formatDate(new Date(), "GMT+7", "dd/MM/yyyy HH:mm:ss");
      settingsSheet.appendRow(["GREETING_KNOWN_TEMPLATE", "สวัสดี{particle}คุณ {name}", "ข้อความทักทายสมาชิกหรือบุคลากรในระบบ", nowStr]);
      settingsSheet.appendRow(["GREETING_UNKNOWN_TEXT", "สวัสดี{particle}คุณลูกค้า วันนี้รับอะไรดี{particle}", "ข้อความทักทายลูกค้าหรือผู้มาติดต่อทั่วไป", nowStr]);
      settingsSheet.appendRow(["VOICE_PARTICLE", "ค่ะ", "คำลงท้ายเสียงพูดทักทาย (ค่ะ/ครับ)", nowStr]);
      settingsSheet.appendRow(["BELL_CHIME_ENABLED", "true", "เปิด/ปิดเสียงกระดิ่งเตือนก่อนทักทาย (true/false)", nowStr]);
      settingsSheet.appendRow(["BELL_CHIME_SOUND", "dingdong", "รูปแบบเสียงกระดิ่งเตือน (dingdong/melody3/single_ding/marimba/crystal/digital)", nowStr]);
      settingsSheet.appendRow(["WORK_START_TIME", "08:30", "เวลาเริ่มเข้างานพนักงาน (HH:mm)", nowStr]);
      settingsSheet.appendRow(["WORK_END_TIME", "17:30", "เวลาเลิกงานพนักงาน (HH:mm)", nowStr]);
      settingsSheet.appendRow(["DETECTION_SPEED_MODE", "balanced", "โหมดความเร็วตรวจจับใบหน้า AI (turbo/balanced/precision)", nowStr]);
      settingsSheet.appendRow(["DETECTION_ENGINE", "tiny", "โมเดล AI ตรวจจับใบหน้า (tiny: TinyFaceDetector / ssd: SSD MobileNet V1)", nowStr]);
      settingsSheet.appendRow(["SSD_MIN_CONFIDENCE", "0.50", "เกณฑ์ความมั่นใจขั้นต่ำโมเดล SSD MobileNet V1 (0.30 - 0.70)", nowStr]);
    }

    // 6. ตรวจสอบ Sheet บันทึกเวลาเข้า-ออกงาน (AttendanceLogs)
    var attSheet = ss.getSheetByName(SHEET_ATTENDANCE);
    if (!attSheet) {
      attSheet = ss.insertSheet(SHEET_ATTENDANCE);
      attSheet.appendRow([
        "Log_ID",
        "Date",
        "Profile_ID",
        "Person_Name",
        "Person_Type",
        "Action_Type",
        "Action_Time",
        "Work_Status",
        "Timestamp",
        "Work_Schedule"
      ]);
      attSheet.getRange(1, 1, 1, 10)
        .setBackground("#1E293B")
        .setFontColor("#FFFFFF")
        .setFontWeight("bold");
      attSheet.setFrozenRows(1);
    } else if (attSheet.getLastColumn() < 10) {
      attSheet.getRange(1, 10).setValue("Work_Schedule")
        .setBackground("#1E293B")
        .setFontColor("#FFFFFF")
        .setFontWeight("bold");
    }

    // 7. ตรวจสอบหรือสร้างโฟลเดอร์สำหรับเก็บภาพใบหน้าใน Google Drive
    getOrCreateDriveFolder();

    return {
      success: true,
      message: "ตั้งค่าแผ่นงานและโฟลเดอร์จัดเก็บข้อมูลเรียบร้อยแล้ว",
      spreadsheetUrl: ss.getUrl()
    };
  } catch (error) {
    return {
      success: false,
      message: error.toString()
    };
  }
}

/**
 * ดึงหรือสร้างโฟลเดอร์ใน Google Drive โดยแคช Folder ID ไว้ใน Script Properties
 */
function getOrCreateDriveFolder() {
  var props = PropertiesService.getScriptProperties();
  var savedFolderId = props.getProperty("DRIVE_FOLDER_ID");
  if (savedFolderId) {
    try {
      return DriveApp.getFolderById(savedFolderId);
    } catch (e) {}
  }

  var folders = DriveApp.getFoldersByName(DRIVE_FOLDER_NAME);
  var folder = folders.hasNext() ? folders.next() : DriveApp.createFolder(DRIVE_FOLDER_NAME);
  try {
    folder.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  } catch (e) {}

  props.setProperty("DRIVE_FOLDER_ID", folder.getId());
  return folder;
}

/**
 * ดึงรายการโปรไฟล์ใบหน้าแบบเร่งด่วน (ไม่เรียก initSheet และไม่แตะ DriveApp)
 */
function getFaceProfiles() {
  try {
    var ss = getDatabaseSpreadsheet();
    var sheet = ss.getSheetByName(SHEET_PROFILES);
    if (!sheet) {
      initSheetAndFolder();
      sheet = ss.getSheetByName(SHEET_PROFILES);
    }

    var data = sheet.getDataRange().getValues();
    if (data.length <= 1) {
      return { success: true, profiles: [] };
    }

    var profiles = [];
    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      var id = row[0];
      var name = row[1];
      var rawDescriptors = row[2];
      var imageUrl = row[3];
      var createdAt = row[5];
      var rawType = row[7] ? String(row[7]).trim().toLowerCase() : '';
      var personType = (rawType === 'vip') ? 'vip' : 'employee';

      var descriptorArray = [];
      try {
        descriptorArray = JSON.parse(rawDescriptors);
      } catch (e) {}

      if (id && name && descriptorArray && descriptorArray.length > 0) {
        profiles.push({
          id: id,
          name: name,
          descriptors: descriptorArray,
          imageUrl: imageUrl || "",
          personType: personType,
          createdAt: createdAt ? Utilities.formatDate(new Date(createdAt), "GMT+7", "yyyy-MM-dd HH:mm:ss") : ""
        });
      }
    }

    return {
      success: true,
      profiles: profiles
    };
  } catch (error) {
    return {
      success: false,
      message: error.toString(),
      profiles: []
    };
  }
}

/**
 * บันทึกโปรไฟล์ใบหน้าใหม่ลง Google Sheets และอัปโหลดรูปลง Google Drive
 */
function saveFaceProfile(payload) {
  try {
    if (!payload || !payload.name || !payload.descriptor) {
      throw new Error("ข้อมูลไม่ครบถ้วน (ต้องระบุชื่อและข้อมูลใบหน้า)");
    }

    var ss = getDatabaseSpreadsheet();
    var sheet = ss.getSheetByName(SHEET_PROFILES);
    if (!sheet) {
      initSheetAndFolder();
      sheet = ss.getSheetByName(SHEET_PROFILES);
    }

    var profileId = "FACE_" + Utilities.getUuid().substring(0, 8).toUpperCase();
    var now = new Date();
    var imageUrl = "";
    var fileId = "";
    var rawType = payload.personType ? String(payload.personType).trim().toLowerCase() : 'employee';
    var personType = (rawType === 'vip') ? 'vip' : 'employee';

    // บันทึกรูปภาพลง Google Drive ถ้ามี base64 ส่งมา
    if (payload.imageBase64 && payload.imageBase64.indexOf("data:image") !== -1) {
      try {
        var folder = getOrCreateDriveFolder();
        var base64Data = payload.imageBase64.split(",")[1];
        var decoded = Utilities.base64Decode(base64Data);
        var blob = Utilities.newBlob(decoded, "image/jpeg", profileId + "_" + payload.name + ".jpg");
        var file = folder.createFile(blob);
        
        try {
          file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
        } catch (e) {}

        fileId = file.getId();
        imageUrl = "https://lh3.googleusercontent.com/d/" + fileId;
      } catch (errDrive) {
        console.error("Save image to Drive failed: " + errDrive.toString());
      }
    }

    // บันทึกลง Sheet (รวมคอลัมน์ Person_Type)
    var descriptorJson = JSON.stringify(payload.descriptor);
    sheet.appendRow([
      profileId,
      payload.name.trim(),
      descriptorJson,
      imageUrl,
      fileId,
      now,
      now,
      personType
    ]);

    return {
      success: true,
      message: "บันทึกโปรไฟล์ " + payload.name + " สำเร็จ (" + (personType === 'vip' ? 'VIP' : 'พนักงาน') + ")",
      profile: {
        id: profileId,
        name: payload.name.trim(),
        descriptors: payload.descriptor,
        imageUrl: imageUrl,
        personType: personType,
        createdAt: Utilities.formatDate(now, "GMT+7", "yyyy-MM-dd HH:mm:ss")
      }
    };
  } catch (error) {
    return {
      success: false,
      message: error.toString()
    };
  }
}

/**
 * อัปเดตประเภทโปรไฟล์ใบหน้า (VIP หรือ พนักงาน)
 */
function updateProfileType(profileId, personType, personName) {
  try {
    var pId = profileId ? String(profileId).trim() : "";
    var pName = personName ? String(personName).trim() : "";
    if (!pId && !pName) {
      return { success: false, message: "ไม่พบ Profile ID หรือ ชื่อบุคคล" };
    }
    var cleanType = (String(personType).toLowerCase() === 'vip') ? 'vip' : 'employee';
    var ss = getDatabaseSpreadsheet();
    var sheet = ss.getSheetByName(SHEET_PROFILES);
    if (!sheet) {
      initSheetAndFolder();
      sheet = ss.getSheetByName(SHEET_PROFILES);
    }

    // ตรวจสอบว่ามีคอลัมน์ Person_Type (คอลัมน์ 8) หรือไม่
    if (sheet.getLastColumn() < 8) {
      sheet.getRange(1, 8).setValue("Person_Type")
        .setBackground("#1E293B")
        .setFontColor("#FFFFFF")
        .setFontWeight("bold");
    }

    var data = sheet.getDataRange().getValues();
    var matchedRowIndex = -1;

    for (var i = 1; i < data.length; i++) {
      var rowId = String(data[i][0] || "").trim();
      var rowName = String(data[i][1] || "").trim();

      // แมตช์ด้วย Profile_ID ก่อน หากไม่มีให้แมตช์ด้วยชื่อ
      if (pId && rowId === pId) {
        matchedRowIndex = i + 1;
        break;
      } else if (!pId && pName && rowName === pName) {
        matchedRowIndex = i + 1;
        break;
      }
    }

    // หากยังไม่เจอ แต่มีทั้งสองค่า ลองหาด้วยชื่อสำรอง
    if (matchedRowIndex === -1 && pName) {
      for (var j = 1; j < data.length; j++) {
        if (String(data[j][1] || "").trim() === pName) {
          matchedRowIndex = j + 1;
          break;
        }
      }
    }

    if (matchedRowIndex !== -1) {
      sheet.getRange(matchedRowIndex, 8).setValue(cleanType);
      SpreadsheetApp.flush(); // บันทึกลง Google Sheets ทันที
      return {
        success: true,
        message: "อัปเดตประเภทเป็น " + (cleanType === 'vip' ? 'VIP' : 'พนักงาน') + " สำเร็จ",
        profileId: pId || String(data[matchedRowIndex - 1][0]),
        personType: cleanType
      };
    }

    return { success: false, message: "ไม่พบรหัสโปรไฟล์ในระบบ: " + (pId || pName) };
  } catch (e) {
    return { success: false, message: e.toString() };
  }
}

/**
 * ลบโปรไฟล์ใบหน้าออกจาก Sheet และ Drive
 */
function deleteFaceProfile(profileId) {
  try {
    var ss = getDatabaseSpreadsheet();
    var sheet = ss.getSheetByName(SHEET_PROFILES);
    var data = sheet.getDataRange().getValues();

    for (var i = 1; i < data.length; i++) {
      if (data[i][0] === profileId) {
        var fileId = data[i][4];
        if (fileId) {
          try {
            DriveApp.getFileById(fileId).setTrashed(true);
          } catch (e) {}
        }
        var profileName = data[i][1];
        sheet.deleteRow(i + 1);
        try {
          clearProfileMessages(profileId, profileName);
        } catch (eMsg) {}
        return {
          success: true,
          message: "ลบโปรไฟล์รหัส " + profileId + " เรียบร้อยแล้ว"
        };
      }
    }

    return {
      success: false,
      message: "ไม่พบรหัสโปรไฟล์ที่ต้องการลบ"
    };
  } catch (error) {
    return {
      success: false,
      message: error.toString()
    };
  }
}

/**
 * บันทึกประวัติการตรวจจับใบหน้า (Log) พร้อมอัปเดต Real-Time Buffer ใน CacheService ทันที
 * รองรับทั้งแบบส่ง Object รายละเอียดครบถ้วน หรือแบบเดิม (name, status, confidence)
 */
function logScanEvent(payloadOrName, status, confidence) {
  try {
    var now = new Date();
    var nowMs = now.getTime();
    var isObj = (payloadOrName && typeof payloadOrName === "object");
    var p = isObj ? payloadOrName : {};

    var id = String(p.id || p.eventId || ("LOG_" + nowMs + "_" + Math.random().toString(36).substr(2, 4).toUpperCase()));
    var timestampMs = Number(p.timestamp || p.timestampMs) || nowMs;
    var dateObj = new Date(timestampMs);
    var timeStr = Utilities.formatDate(dateObj, "GMT+7", "HH:mm:ss");
    var dateStr = Utilities.formatDate(dateObj, "GMT+7", "dd/MM/yyyy");

    var personName = String(isObj ? (p.personName || p.name || "") : (payloadOrName || "")).trim() || "ลูกค้าทั่วไป";
    var matchStatus = String(isObj ? (p.status || p.matchStatus || "") : (status || "")).trim().toUpperCase() || (personName === "ลูกค้าทั่วไป" ? "UNKNOWN" : "KNOWN");
    
    var confVal = isObj ? (p.confidence !== undefined ? p.confidence : p.confidencePct) : confidence;
    var confPct = 0;
    if (typeof confVal === "number") {
      confPct = confVal <= 1 ? Math.round(confVal * 100) : Math.round(confVal);
    } else if (typeof confVal === "string" && confVal.indexOf("%") !== -1) {
      confPct = Math.round(parseFloat(confVal)) || 0;
    }
    var confStr = confPct > 0 ? (confPct + "%") : "-";

    var personType = String(p.personType || (matchStatus === "KNOWN" ? "employee" : "unknown")).toLowerCase();
    var detectedByUsername = String(p.detectedByUsername || p.username || "unknown");
    var detectedByFullName = String(p.detectedByFullName || p.fullName || detectedByUsername);
    var detectedByRole = String(p.detectedByRole || p.role || "operator");
    var sessionId = String(p.sessionId || "");
    var detail = String(p.detail || (matchStatus === "KNOWN" ? "ยืนยันตัวตนสำเร็จ" : "ตรวจพบผู้มาติดต่อ"));

    var logItem = {
      id: id,
      timestamp: timestampMs,
      timeStr: timeStr,
      dateStr: dateStr,
      personName: personName,
      profileId: String(p.profileId || ""),
      personType: personType,
      status: matchStatus,
      confidencePct: confPct,
      detectedByUsername: detectedByUsername,
      detectedByFullName: detectedByFullName,
      detectedByRole: detectedByRole,
      sessionId: sessionId,
      detail: detail
    };

    // 1. บันทึกลง Real-Time Ring Buffer ใน CacheService (ดึงเร็วระดับ ms สำหรับ Admin)
    try {
      var cache = CacheService.getScriptCache();
      var rawLogs = cache.get(REALTIME_SCAN_LOGS_KEY);
      var logList = [];
      if (rawLogs) {
        try { logList = JSON.parse(rawLogs) || []; } catch (eJson) { logList = []; }
      }
      var existIdx = -1;
      for (var li = 0; li < logList.length; li++) {
        if (logList[li].id === id) { existIdx = li; break; }
      }
      if (existIdx >= 0) {
        logList[existIdx] = logItem;
      } else {
        logList.unshift(logItem);
      }
      if (logList.length > MAX_REALTIME_SCAN_LOGS) {
        logList = logList.slice(0, MAX_REALTIME_SCAN_LOGS);
      }
      cache.put(REALTIME_SCAN_LOGS_KEY, JSON.stringify(logList), 21600); // แคช 6 ชั่วโมง
    } catch (eCache) {
      console.warn("CacheService put error:", eCache);
    }

    // 2. บันทึกประวัติถาวรลงชีต ScanLogs
    try {
      var ss = getDatabaseSpreadsheet();
      var sheet = ss.getSheetByName(SHEET_LOGS);
      if (sheet) {
        var logHeaders = ["Timestamp", "Person_Name", "Match_Status", "Confidence", "Person_Type", "Detected_By_User", "Detected_By_Name", "Detail"];
        if (sheet.getLastColumn() < logHeaders.length) {
          sheet.getRange(1, 1, 1, logHeaders.length).setValues([logHeaders])
            .setBackground("#1E293B").setFontColor("#FFFFFF").setFontWeight("bold");
        }
        sheet.appendRow([
          dateObj,
          personName,
          matchStatus,
          confStr,
          personType,
          detectedByUsername,
          detectedByFullName,
          detail
        ]);
      }
    } catch (eSheet) {
      console.warn("Sheet append error:", eSheet);
    }

    return { success: true, log: logItem };
  } catch (e) {
    return { success: false, error: e.toString() };
  }
}

/**
 * ฟังก์ชันเข้ารหัสรหัสผ่านด้วย SHA-256
 */
function hashPassword(password) {
  var rawHash = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, String(password), Utilities.Charset.UTF_8);
  var hex = "";
  for (var i = 0; i < rawHash.length; i++) {
    var byteVal = rawHash[i];
    if (byteVal < 0) byteVal += 256;
    var byteHex = byteVal.toString(16);
    if (byteHex.length === 1) byteHex = "0" + byteHex;
    hex += byteHex;
  }
  return hex;
}

/**
 * เข้าสู่ระบบด้วย Username และ Password
 */
function loginUser(username, password) {
  try {
    if (!username || !password) {
      return { success: false, message: "กรุณากรอกชื่อผู้ใช้และรหัสผ่าน" };
    }
    initSheetAndFolder();
    var ss = getDatabaseSpreadsheet();
    var sheet = ss.getSheetByName(SHEET_USERS);
    if (!sheet) return { success: false, message: "ไม่พบตารางข้อมูลผู้ใช้งาน" };

    var data = sheet.getDataRange().getValues();
    if (data.length <= 1) return { success: false, message: "ไม่มีข้อมูลผู้ใช้งานในระบบ" };

    var inputHash = hashPassword(password);
    var cleanUser = String(username).trim().toLowerCase();

    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      var rowUser = String(row[1]).trim().toLowerCase();
      var rowHash = String(row[2]).trim();
      var rowStatus = String(row[5]).trim().toUpperCase();

      if (rowUser === cleanUser) {
        if (rowStatus !== "ACTIVE") {
          return { success: false, message: "บัญชีผู้ใช้นี้ถูกระงับการใช้งาน กรุณาติดต่อผู้ดูแลระบบ" };
        }
        if (rowHash === inputHash) {
          var nowStr = Utilities.formatDate(new Date(), "GMT+7", "dd/MM/yyyy HH:mm:ss");
          sheet.getRange(i + 1, 8).setValue(nowStr);

          return {
            success: true,
            message: "เข้าสู่ระบบสำเร็จ",
            user: {
              id: String(row[0]),
              username: String(row[1]),
              fullName: String(row[3]),
              role: String(row[4]).toLowerCase(),
              status: rowStatus
            }
          };
        } else {
          return { success: false, message: "รหัสผ่านไม่ถูกต้อง กรุณาลองใหม่อีกครั้ง" };
        }
      }
    }

    return { success: false, message: "ไม่พบบัญชีผู้ใช้นี้ในระบบ" };
  } catch (error) {
    return { success: false, message: error.toString() };
  }
}

/**
 * ดึงรายชื่อผู้ใช้งานระบบทั้งหมด (ไม่ส่ง Password Hash เพื่อความปลอดภัย)
 */
function getSystemUsers() {
  try {
    initSheetAndFolder();
    var ss = getDatabaseSpreadsheet();
    var sheet = ss.getSheetByName(SHEET_USERS);
    if (!sheet) return { success: true, users: [] };

    var data = sheet.getDataRange().getValues();
    var users = [];
    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      if (!row[0] && !row[1]) continue;
      users.push({
        id: String(row[0]),
        username: String(row[1]),
        fullName: String(row[3]),
        role: String(row[4]).toLowerCase(),
        status: String(row[5]).toUpperCase(),
        createdAt: String(row[6] || "-"),
        lastLogin: String(row[7] || "-")
      });
    }
    return { success: true, users: users };
  } catch (error) {
    return { success: false, message: error.toString(), users: [] };
  }
}

/**
 * สร้างผู้ใช้งานระบบใหม่
 */
function createSystemUser(userData) {
  try {
    if (!userData || !userData.username || !userData.password) {
      return { success: false, message: "กรุณาระบุชื่อผู้ใช้และรหัสผ่าน" };
    }
    var cleanUser = String(userData.username).trim().toLowerCase();
    if (cleanUser.length < 3) {
      return { success: false, message: "ชื่อผู้ใช้ต้องมีความยาวอย่างน้อย 3 ตัวอักษร" };
    }
    if (String(userData.password).length < 4) {
      return { success: false, message: "รหัสผ่านต้องมีความยาวอย่างน้อย 4 ตัวอักษร" };
    }

    initSheetAndFolder();
    var ss = getDatabaseSpreadsheet();
    var sheet = ss.getSheetByName(SHEET_USERS);
    var data = sheet.getDataRange().getValues();

    for (var i = 1; i < data.length; i++) {
      if (String(data[i][1]).trim().toLowerCase() === cleanUser) {
        return { success: false, message: "ชื่อผู้ใช้นี้มีในระบบแล้ว กรุณาใช้ชื่ออื่น" };
      }
    }

    var newId = "USR_" + (data.length + Date.now().toString().slice(-4));
    var nowStr = Utilities.formatDate(new Date(), "GMT+7", "dd/MM/yyyy HH:mm:ss");
    var fullName = String(userData.fullName || cleanUser).trim();
    var role = String(userData.role || "operator").trim().toLowerCase();
    var status = String(userData.status || "ACTIVE").trim().toUpperCase();

    sheet.appendRow([
      newId,
      cleanUser,
      hashPassword(userData.password),
      fullName,
      role,
      status,
      nowStr,
      "-"
    ]);

    return {
      success: true,
      message: "สร้างผู้ใช้งาน " + cleanUser + " เรียบร้อยแล้ว",
      user: {
        id: newId,
        username: cleanUser,
        fullName: fullName,
        role: role,
        status: status
      }
    };
  } catch (error) {
    return { success: false, message: error.toString() };
  }
}

/**
 * แก้ไขข้อมูลผู้ใช้งาน หรือเปลี่ยนรหัสผ่าน
 */
function updateSystemUser(userData) {
  try {
    if (!userData || !userData.id) {
      return { success: false, message: "ไม่พบรหัสผู้ใช้ที่ต้องการแก้ไข" };
    }
    var ss = getDatabaseSpreadsheet();
    var sheet = ss.getSheetByName(SHEET_USERS);
    var data = sheet.getDataRange().getValues();

    for (var i = 1; i < data.length; i++) {
      if (String(data[i][0]) === String(userData.id)) {
        var rowNum = i + 1;
        if (userData.fullName) sheet.getRange(rowNum, 4).setValue(String(userData.fullName).trim());
        if (userData.role) sheet.getRange(rowNum, 5).setValue(String(userData.role).trim().toLowerCase());
        if (userData.status) sheet.getRange(rowNum, 6).setValue(String(userData.status).trim().toUpperCase());
        if (userData.password && String(userData.password).trim().length > 0) {
          if (String(userData.password).length < 4) {
            return { success: false, message: "รหัสผ่านใหม่ต้องมีความยาวอย่างน้อย 4 ตัวอักษร" };
          }
          sheet.getRange(rowNum, 3).setValue(hashPassword(userData.password));
        }

        return { success: true, message: "อัปเดตข้อมูลผู้ใช้เรียบร้อยแล้ว" };
      }
    }

    return { success: false, message: "ไม่พบผู้ใช้ในระบบ" };
  } catch (error) {
    return { success: false, message: error.toString() };
  }
}

/**
 * ลบผู้ใช้งานระบบ (ป้องกันการลบ Admin คนสุดท้าย)
 */
function deleteSystemUser(userId) {
  try {
    if (!userId) return { success: false, message: "ไม่พบรหัสผู้ใช้" };
    var ss = getDatabaseSpreadsheet();
    var sheet = ss.getSheetByName(SHEET_USERS);
    var data = sheet.getDataRange().getValues();

    var adminCount = 0;
    var targetIndex = -1;
    var targetRole = "";

    for (var i = 1; i < data.length; i++) {
      var role = String(data[i][4]).trim().toLowerCase();
      if (role === "admin") adminCount++;
      if (String(data[i][0]) === String(userId)) {
        targetIndex = i;
        targetRole = role;
      }
    }

    if (targetIndex === -1) {
      return { success: false, message: "ไม่พบผู้ใช้ที่ต้องการลบ" };
    }

    if (targetRole === "admin" && adminCount <= 1) {
      return { success: false, message: "ไม่อนุญาตให้ลบผู้ดูแลระบบคนสุดท้ายของระบบ" };
    }

    sheet.deleteRow(targetIndex + 1);
    return { success: true, message: "ลบผู้ใช้งานเรียบร้อยแล้ว" };
  } catch (error) {
    return { success: false, message: error.toString() };
  }
}

/**
 * รองรับ POST Request จาก Frontend ภายนอก (GitHub Pages)
 */
function doPost(e) {
  try {
    var data = {};
    if (e.postData && e.postData.contents) {
      data = JSON.parse(e.postData.contents);
    }
    var action = (e.parameter && e.parameter.action) || data.action;
    var result = { success: false, message: "Invalid action" };

    if (action === "login") {
      result = loginUser(data.username, data.password);
    } else if (action === "getUsers") {
      result = getSystemUsers();
    } else if (action === "createUser") {
      result = createSystemUser(data.userData || data);
    } else if (action === "updateUser") {
      result = updateSystemUser(data.userData || data);
    } else if (action === "deleteUser") {
      result = deleteSystemUser(data.id || (e.parameter && e.parameter.id));
    } else if (action === "saveProfile") {
      result = saveFaceProfile(data.payload || data);
    } else if (action === "deleteProfile") {
      result = deleteFaceProfile(data.id || (e.parameter && e.parameter.id));
    } else if (action === "logScan") {
      result = logScanEvent(data.payload || data, data.status, data.confidence);
    } else if (action === "getProfiles") {
      result = getFaceProfiles();
    } else if (action === "getMessages") {
      result = getProfileMessages();
    } else if (action === "saveMessage") {
      result = saveProfileMessage(data.messageData || data.payload || data);
    } else if (action === "deleteMessage") {
      result = deleteProfileMessage(data.id || (e.parameter && e.parameter.id));
    } else if (action === "clearMessages") {
      result = clearProfileMessages(data.profileId || (e.parameter && e.parameter.profileId), data.personName);
    } else if (action === "getSettings") {
      result = getSystemSettings();
    } else if (action === "saveSettings") {
      result = saveSystemSettings(data.settings || data);
    } else if (action === "getTodayAttendance") {
      result = getTodayAttendance(data.date || data.targetDate || (e.parameter && (e.parameter.date || e.parameter.targetDate)));
    } else if (action === "recordAttendance") {
      result = recordAttendanceLog(data.payload || data);
    } else if (action === "updateProfileType") {
      var pId = data.id || data.profileId || (e.parameter && (e.parameter.id || e.parameter.profileId));
      var pType = data.personType || (e.parameter && e.parameter.personType);
      var pName = data.name || data.personName || (e.parameter && (e.parameter.name || e.parameter.personName));
      result = updateProfileType(pId, pType, pName);
    } else if (action === "getRealtimeLogs") {
      result = getRealtimeDetectionLogs(data.payload || data || (e.parameter && e.parameter));
    } else if (action === "clearRealtimeLogs") {
      result = clearRealtimeDetectionLogs();
    } else if (action === "triggerGlobalReload") {
      result = triggerGlobalForceReload(data.payload || data || (e.parameter && e.parameter));
    } else if (action === "clientPulse") {
      result = syncClientPresenceAndCommands(data.payload || data || (e.parameter && e.parameter));
    }

    return ContentService.createTextOutput(JSON.stringify(result))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (error) {
    return ContentService.createTextOutput(JSON.stringify({ success: false, message: error.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

/**
 * ดึงหรือสร้าง Sheet ProfileMessages โดยอัตโนมัติพร้อมหัวตาราง
 */
function getOrCreateProfileMessagesSheet(ss) {
  var sheet = ss.getSheetByName(SHEET_MESSAGES);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_MESSAGES);
    sheet.appendRow([
      "Message_ID", 
      "Profile_ID", 
      "Person_Name", 
      "Message_Text", 
      "Created_At"
    ]);
    sheet.getRange(1, 1, 1, 5)
      .setBackground("#1E293B")
      .setFontColor("#FFFFFF")
      .setFontWeight("bold");
    sheet.setFrozenRows(1);
  }
  return sheet;
}

/**
 * ดึงรายการข้อความฝากทั้งหมดจากตาราง ProfileMessages
 */
function getProfileMessages() {
  try {
    var ss = getDatabaseSpreadsheet();
    var sheet = getOrCreateProfileMessagesSheet(ss);
    var data = sheet.getDataRange().getValues();
    var messages = {};

    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      var msgId = String(row[0] || '').trim();
      var profileId = String(row[1] || '').trim();
      var personName = String(row[2] || '').trim();
      var text = String(row[3] || '').trim();
      var createdAt = String(row[4] || '').trim();

      if (msgId && text) {
        var item = {
          id: msgId,
          profileId: profileId,
          personName: personName,
          text: text,
          createdAt: createdAt
        };
        // ทำดัชนีทั้งสองคีย์เพื่อความรวดเร็วและแม่นยำในการค้นหา
        if (profileId) {
          if (!messages[profileId]) messages[profileId] = [];
          messages[profileId].push(item);
        }
        if (personName && personName !== profileId) {
          if (!messages[personName]) messages[personName] = [];
          messages[personName].push(item);
        }
      }
    }
    return { success: true, messages: messages };
  } catch (error) {
    return { success: false, message: error.toString(), messages: {} };
  }
}

/**
 * บันทึกข้อความฝากใหม่ลงตาราง ProfileMessages
 */
function saveProfileMessage(payload) {
  try {
    if (!payload) return { success: false, message: "ไม่มีข้อมูลข้อความ" };
    var text = String(payload.text || payload.messageText || "").trim();
    if (!text) {
      return { success: false, message: "กรุณาระบุข้อความ" };
    }
    var ss = getDatabaseSpreadsheet();
    var sheet = getOrCreateProfileMessagesSheet(ss);

    var msgId = payload.id || ("MSG_" + Date.now().toString());
    var profileId = String(payload.profileId || "").trim();
    var personName = String(payload.personName || payload.name || "").trim();
    var createdAt = payload.createdAt || Utilities.formatDate(new Date(), "GMT+7", "dd/MM/yyyy HH:mm");

    sheet.appendRow([msgId, profileId, personName, text, createdAt]);
    return { success: true, message: "บันทึกข้อความเรียบร้อยแล้ว", id: msgId };
  } catch (error) {
    return { success: false, message: error.toString() };
  }
}

/**
 * ลบข้อความฝากตาม Message ID (ลบทุกแถวที่ตรงกันแบบย้อนหลังจากล่างขึ้นบน)
 */
function deleteProfileMessage(messageId) {
  try {
    if (!messageId) return { success: false, message: "ไม่พบรหัสข้อความ" };
    var ss = getDatabaseSpreadsheet();
    var sheet = getOrCreateProfileMessagesSheet(ss);
    var data = sheet.getDataRange().getValues();
    var targetId = String(messageId).trim().toLowerCase();
    var deletedCount = 0;

    for (var i = data.length - 1; i >= 1; i--) {
      var rowId = String(data[i][0] || '').trim().toLowerCase();
      if (rowId === targetId) {
        sheet.deleteRow(i + 1);
        deletedCount++;
      }
    }

    if (deletedCount > 0) {
      return { success: true, count: deletedCount, message: "ลบข้อความเรียบร้อยแล้ว (" + deletedCount + " แถว)" };
    }
    return { success: false, message: "ไม่พบข้อความที่ต้องการลบ" };
  } catch (error) {
    return { success: false, message: error.toString() };
  }
}

/**
 * ล้างข้อความทั้งหมดของ Profile ID หรือ Person Name
 */
function clearProfileMessages(profileId, personName) {
  try {
    var ss = getDatabaseSpreadsheet();
    var sheet = getOrCreateProfileMessagesSheet(ss);
    var data = sheet.getDataRange().getValues();
    var cleanId = String(profileId || "").trim().toLowerCase();
    var cleanName = String(personName || "").trim().toLowerCase();
    var deletedCount = 0;

    for (var i = data.length - 1; i >= 1; i--) {
      var rowProfId = String(data[i][1] || '').trim().toLowerCase();
      var rowName = String(data[i][2] || '').trim().toLowerCase();
      if ((cleanId && rowProfId === cleanId) || (cleanName && rowName === cleanName)) {
        sheet.deleteRow(i + 1);
        deletedCount++;
      }
    }
    return { success: true, count: deletedCount, message: "ล้างข้อความทั้งหมดเรียบร้อยแล้ว (" + deletedCount + " แถว)" };
  } catch (error) {
    return { success: false, message: error.toString() };
  }
}

/**
 * ดึงหรือสร้าง Sheet SystemSettings โดยอัตโนมัติพร้อมหัวตารางและค่าคงที่เริ่มต้น
 */
function getOrCreateSystemSettingsSheet(ss) {
  var sheet = ss.getSheetByName(SHEET_SETTINGS);
  if (!sheet) {
    initSheetAndFolder();
    sheet = ss.getSheetByName(SHEET_SETTINGS);
  }
  return sheet;
}

/**
 * แปลงค่าเวลาทุกรูปแบบ (Date, String 1899, 'HH:mm) ให้เป็น "HH:mm" อย่างแม่นยำ
 * โดยใช้ getHours/getMinutes เพื่อเลี่ยงปัญหา Local Mean Time +06:42 ของปี 1899 ที่ทำให้เวลาเลื่อน 17 นาที
 */
function normalizeTimeHHMM_(val) {
  if (val === null || val === undefined || val === "") return "";
  if (val instanceof Date) {
    var h = val.getHours();
    var m = val.getMinutes();
    return ("0" + h).slice(-2) + ":" + ("0" + m).slice(-2);
  }
  var s = String(val).trim();
  if (s.charAt(0) === "'") s = s.substring(1);
  if (s.includes("1899") || s.includes("GMT")) {
    var d = new Date(s);
    if (!isNaN(d.getTime())) {
      return ("0" + d.getHours()).slice(-2) + ":" + ("0" + d.getMinutes()).slice(-2);
    }
  }
  var match = s.match(/(\d{1,2}):(\d{2})/);
  return match ? ("0" + match[1]).slice(-2) + ":" + match[2] : s;
}

/**
 * ดึงการตั้งค่าระบบและข้อความทักทายที่เป็นค่าคงที่ถาวร
 */
function getSystemSettings() {
  try {
    var ss = getDatabaseSpreadsheet();
    var sheet = getOrCreateSystemSettingsSheet(ss);
    var data = sheet.getDataRange().getValues();

    var settings = {
      employeeGreetingTemplate: "สวัสดี{particle}คุณ {name} วันนี้{datetime}",
      vipGreetingTemplate: "สวัสดี{particle}คุณ {name}",
      knownGreetingTemplate: "สวัสดี{particle}คุณ {name}",
      unknownGreetingText: "สวัสดี{particle}คุณลูกค้า วันนี้รับอะไรดี{particle}",
      voiceParticle: "ค่ะ",
      bellChimeEnabled: true,
      bellChimeSound: "dingdong",
      voiceCategoryConfigs: null,
      workStartTime: "08:30",
      workEndTime: "17:30",
      faceMatchThreshold: 0.41,
      detectionSpeedMode: "balanced",
      detectionEngine: "tiny",
      ssdMinConfidence: 0.50
    };

    if (data && data.length > 1) {
      for (var i = 1; i < data.length; i++) {
        var key = String(data[i][0] || "").trim().toUpperCase();
        var val = data[i][1];
        if (key === "GREETING_EMPLOYEE_TEMPLATE" && val !== undefined && val !== null && String(val).trim() !== "") {
          settings.employeeGreetingTemplate = String(val).trim();
        } else if (key === "GREETING_VIP_TEMPLATE" && val !== undefined && val !== null && String(val).trim() !== "") {
          settings.vipGreetingTemplate = String(val).trim();
          settings.knownGreetingTemplate = String(val).trim();
        } else if (key === "GREETING_KNOWN_TEMPLATE" && val !== undefined && val !== null && String(val).trim() !== "") {
          settings.knownGreetingTemplate = String(val).trim();
          if (!settings._hasExplicitVip) {
            settings.vipGreetingTemplate = String(val).trim();
          }
        } else if (key === "GREETING_UNKNOWN_TEXT" && val !== undefined && val !== null && String(val).trim() !== "") {
          settings.unknownGreetingText = String(val).trim();
        } else if (key === "VOICE_PARTICLE" && val !== undefined && val !== null && String(val).trim() !== "") {
          settings.voiceParticle = String(val).trim();
        } else if (key === "BELL_CHIME_ENABLED" && val !== undefined && val !== null) {
          settings.bellChimeEnabled = (String(val).trim().toLowerCase() === "true");
        } else if (key === "BELL_CHIME_SOUND" && val !== undefined && val !== null && String(val).trim() !== "") {
          settings.bellChimeSound = String(val).trim().toLowerCase();
        } else if (key === "VOICE_CATEGORY_CONFIGS" && val !== undefined && val !== null && String(val).trim() !== "") {
          try {
            settings.voiceCategoryConfigs = JSON.parse(String(val).trim());
          } catch (eJson) {}
        } else if (key === "WORK_START_TIME" && val !== undefined && val !== null && String(val).trim() !== "") {
          settings.workStartTime = normalizeTimeHHMM_(val) || "08:30";
        } else if (key === "WORK_END_TIME" && val !== undefined && val !== null && String(val).trim() !== "") {
          settings.workEndTime = normalizeTimeHHMM_(val) || "17:30";
        } else if (key === "FACE_MATCH_THRESHOLD" && val !== undefined && val !== null && String(val).trim() !== "") {
          var parsedThr = parseFloat(String(val).trim());
          if (!isNaN(parsedThr) && parsedThr >= 0.30 && parsedThr <= 0.55) {
            settings.faceMatchThreshold = Math.round(parsedThr * 1000) / 1000;
          }
        } else if (key === "DETECTION_SPEED_MODE" && val !== undefined && val !== null && String(val).trim() !== "") {
          var smVal = String(val).trim().toLowerCase();
          if (smVal === "turbo" || smVal === "precision" || smVal === "balanced") {
            settings.detectionSpeedMode = smVal;
          }
        } else if (key === "DETECTION_ENGINE" && val !== undefined && val !== null && String(val).trim() !== "") {
          var engVal = String(val).trim().toLowerCase();
          if (engVal === "tiny" || engVal === "ssd") {
            settings.detectionEngine = engVal;
          }
        } else if (key === "SSD_MIN_CONFIDENCE" && val !== undefined && val !== null && String(val).trim() !== "") {
          var confVal = parseFloat(String(val).trim());
          if (!isNaN(confVal) && confVal >= 0.20 && confVal <= 0.90) {
            settings.ssdMinConfidence = Math.round(confVal * 100) / 100;
          }
        }
        if (key === "GREETING_VIP_TEMPLATE") {
          settings._hasExplicitVip = true;
        }
      }
      delete settings._hasExplicitVip;
    }

    return {
      success: true,
      settings: settings
    };
  } catch (error) {
    return {
      success: false,
      message: error.toString(),
      settings: {
        employeeGreetingTemplate: "สวัสดี{particle}คุณ {name} วันนี้{datetime}",
        vipGreetingTemplate: "สวัสดี{particle}คุณ {name}",
        knownGreetingTemplate: "สวัสดี{particle}คุณ {name}",
        unknownGreetingText: "สวัสดี{particle}คุณลูกค้า วันนี้รับอะไรดี{particle}",
        voiceParticle: "ค่ะ",
        bellChimeEnabled: true,
        bellChimeSound: "dingdong",
        voiceCategoryConfigs: null,
        workStartTime: "08:30",
        workEndTime: "17:30",
        faceMatchThreshold: 0.41,
        detectionSpeedMode: "balanced",
        detectionEngine: "tiny",
        ssdMinConfidence: 0.50
      }
    };
  }
}

/**
 * บันทึกการตั้งค่าระบบและข้อความทักทายเป็นค่าคงที่ถาวรลงใน Google Sheets
 */
function saveSystemSettings(payload) {
  try {
    if (!payload) return { success: false, message: "ไม่มีข้อมูลการตั้งค่า" };
    var ss = getDatabaseSpreadsheet();
    var sheet = getOrCreateSystemSettingsSheet(ss);
    var data = sheet.getDataRange().getValues();

    var nowStr = Utilities.formatDate(new Date(), "GMT+7", "dd/MM/yyyy HH:mm:ss");

    var updateMap = {};
    if (payload.employeeGreetingTemplate !== undefined) {
      updateMap["GREETING_EMPLOYEE_TEMPLATE"] = {
        value: String(payload.employeeGreetingTemplate).trim(),
        desc: "ข้อความทักทายพนักงาน (Employee)"
      };
    }
    if (payload.vipGreetingTemplate !== undefined) {
      updateMap["GREETING_VIP_TEMPLATE"] = {
        value: String(payload.vipGreetingTemplate).trim(),
        desc: "ข้อความทักทายบุคคลสำคัญ (VIP)"
      };
      updateMap["GREETING_KNOWN_TEMPLATE"] = {
        value: String(payload.vipGreetingTemplate).trim(),
        desc: "ข้อความทักทายสมาชิกหรือบุคลากรในระบบ"
      };
    } else if (payload.knownGreetingTemplate !== undefined) {
      updateMap["GREETING_KNOWN_TEMPLATE"] = {
        value: String(payload.knownGreetingTemplate).trim(),
        desc: "ข้อความทักทายสมาชิกหรือบุคลากรในระบบ"
      };
    }
    if (payload.unknownGreetingText !== undefined) {
      updateMap["GREETING_UNKNOWN_TEXT"] = {
        value: String(payload.unknownGreetingText).trim(),
        desc: "ข้อความทักทายลูกค้าหรือผู้มาติดต่อทั่วไป"
      };
    }
    if (payload.voiceParticle !== undefined) {
      updateMap["VOICE_PARTICLE"] = {
        value: String(payload.voiceParticle).trim(),
        desc: "คำลงท้ายเสียงพูดทักทาย (ค่ะ/ครับ)"
      };
    }
    if (payload.bellChimeEnabled !== undefined) {
      updateMap["BELL_CHIME_ENABLED"] = {
        value: String(payload.bellChimeEnabled).trim().toLowerCase() === "true" ? "true" : "false",
        desc: "เปิด/ปิดเสียงกระดิ่งเตือนก่อนทักทาย (true/false)"
      };
    }
    if (payload.bellChimeSound !== undefined) {
      updateMap["BELL_CHIME_SOUND"] = {
        value: String(payload.bellChimeSound).trim().toLowerCase() || "dingdong",
        desc: "รูปแบบเสียงกระดิ่งเตือน (dingdong/melody3/single_ding/marimba/crystal/digital)"
      };
    }
    if (payload.voiceCategoryConfigs !== undefined && payload.voiceCategoryConfigs !== null) {
      var jsonStr = (typeof payload.voiceCategoryConfigs === "string")
        ? payload.voiceCategoryConfigs
        : JSON.stringify(payload.voiceCategoryConfigs);
      updateMap["VOICE_CATEGORY_CONFIGS"] = {
        value: jsonStr,
        desc: "การตั้งค่าสุ้มเสียงแยกตามประเภท (พนักงาน / VIP / ผู้มาติดต่อ)"
      };
    }
    if (payload.workStartTime !== undefined) {
      updateMap["WORK_START_TIME"] = {
        value: normalizeTimeHHMM_(payload.workStartTime) || "08:30",
        desc: "เวลาเริ่มเข้างานพนักงาน (HH:mm)",
        isText: true
      };
    }
    if (payload.workEndTime !== undefined) {
      updateMap["WORK_END_TIME"] = {
        value: normalizeTimeHHMM_(payload.workEndTime) || "17:30",
        desc: "เวลาเลิกงานพนักงาน (HH:mm)",
        isText: true
      };
    }
    if (payload.faceMatchThreshold !== undefined) {
      var thrNum = parseFloat(String(payload.faceMatchThreshold));
      if (!isNaN(thrNum) && thrNum >= 0.30 && thrNum <= 0.55) {
        updateMap["FACE_MATCH_THRESHOLD"] = {
          value: String(Math.round(thrNum * 1000) / 1000),
          desc: "เกณฑ์ระยะห่างความแม่นยำใบหน้า (ยิ่งต่ำยิ่งเข้มงวด แนะนำ 0.41 สำหรับแยกพี่น้อง/แม่ลูก)",
          isText: true
        };
      }
    }
    if (payload.detectionSpeedMode !== undefined) {
      var smMode = String(payload.detectionSpeedMode).trim().toLowerCase();
      if (smMode === "turbo" || smMode === "precision" || smMode === "balanced") {
        updateMap["DETECTION_SPEED_MODE"] = {
          value: smMode,
          desc: "โหมดความเร็วตรวจจับใบหน้า AI (turbo/balanced/precision)"
        };
      }
    }
    if (payload.detectionEngine !== undefined) {
      var engMode = String(payload.detectionEngine).trim().toLowerCase();
      if (engMode === "tiny" || engMode === "ssd") {
        updateMap["DETECTION_ENGINE"] = {
          value: engMode,
          desc: "โมเดล AI ตรวจจับใบหน้า (tiny: TinyFaceDetector / ssd: SSD MobileNet V1)"
        };
      }
    }
    if (payload.ssdMinConfidence !== undefined) {
      var sConf = parseFloat(payload.ssdMinConfidence);
      if (!isNaN(sConf) && sConf >= 0.20 && sConf <= 0.90) {
        updateMap["SSD_MIN_CONFIDENCE"] = {
          value: String(Math.round(sConf * 100) / 100),
          desc: "เกณฑ์ความมั่นใจขั้นต่ำโมเดล SSD MobileNet V1 (0.30 - 0.70)"
        };
      }
    }

    // อัปเดตแถวที่มีอยู่เดิม
    var foundKeys = {};
    for (var i = 1; i < data.length; i++) {
      var rowKey = String(data[i][0] || "").trim().toUpperCase();
      if (updateMap[rowKey]) {
        var cell = sheet.getRange(i + 1, 2);
        if (updateMap[rowKey].isText) {
          cell.setNumberFormat("@").setValue("'" + updateMap[rowKey].value);
        } else {
          cell.setValue(updateMap[rowKey].value);
        }
        sheet.getRange(i + 1, 4).setValue(nowStr);
        foundKeys[rowKey] = true;
      }
    }

    // เพิ่มคีย์ใหม่ที่ยังไม่มีในตาราง
    for (var key in updateMap) {
      if (!foundKeys[key]) {
        var valToStore = updateMap[key].isText ? ("'" + updateMap[key].value) : updateMap[key].value;
        sheet.appendRow([key, valToStore, updateMap[key].desc, nowStr]);
        if (updateMap[key].isText) {
          sheet.getRange(sheet.getLastRow(), 2).setNumberFormat("@");
        }
      }
    }

    // สั่ง flush ลงแผ่นงานทันที
    SpreadsheetApp.flush();

    // บันทึกลง Script Properties เพื่อการอ่านที่รวดเร็ว
    try {
      var props = PropertiesService.getScriptProperties();
      for (var k in updateMap) {
        props.setProperty(k, String(updateMap[k].value));
      }
    } catch (eProp) {}

    return {
      success: true,
      message: "บันทึกค่าคงที่ระบบถาวรเรียบร้อยแล้ว",
      updatedAt: nowStr
    };
  } catch (error) {
    return { success: false, message: error.toString() };
  }
}

/**
 * บันทึกประวัติเวลาเข้า-ออกงานพนักงานลงตาราง AttendanceLogs
 */
function recordAttendanceLog(payload) {
  try {
    if (!payload || !payload.name) {
      return { success: false, message: "กรุณาระบุข้อมูลพนักงาน" };
    }
    var ss = getDatabaseSpreadsheet();
    var sheet = ss.getSheetByName(SHEET_ATTENDANCE);
    if (!sheet) {
      initSheetAndFolder();
      sheet = ss.getSheetByName(SHEET_ATTENDANCE);
    }

    var now = new Date();
    var nowStr = Utilities.formatDate(now, "GMT+7", "dd/MM/yyyy HH:mm:ss");
    var dateStr = payload.date || Utilities.formatDate(now, "GMT+7", "yyyy-MM-dd");

    // Format clean HH:mm string
    var rawTime = payload.time || Utilities.formatDate(now, "GMT+7", "HH:mm");
    var timeStr = String(rawTime).trim();
    if (timeStr.length > 5 && timeStr.indexOf(':') === 2) {
      timeStr = timeStr.substring(0, 5);
    }

    var actionType = String(payload.actionType || payload.logType || "CHECK_IN").toUpperCase();
    if (actionType !== 'CHECK_OUT') actionType = 'CHECK_IN';

    var workStatus = String(payload.workStatus || payload.status || (actionType === 'CHECK_IN' ? 'ตรงเวลา' : 'เลิกงานตามเวลา')).trim();
    var workSchedule = String(payload.workSchedule || "").trim();
    var logId = "ATT_" + Date.now() + "_" + Math.random().toString(36).substr(2, 4).toUpperCase();
    var personName = String(payload.name).trim();
    var profileId = String(payload.profileId || "").trim();
    var personType = String(payload.personType || "พนักงาน").trim();

    // Store with prepended single quote so Google Sheets doesn't convert to 1899 Date
    sheet.appendRow([
      logId,
      dateStr,
      profileId,
      personName,
      personType,
      actionType,
      "'" + timeStr,
      workStatus,
      nowStr,
      workSchedule
    ]);

    var logData = {
      id: logId,
      date: dateStr,
      profileId: profileId,
      name: personName,
      personType: personType,
      actionType: actionType,
      logType: actionType,
      time: timeStr,
      workStatus: workStatus,
      status: workStatus,
      workSchedule: workSchedule,
      timestamp: nowStr
    };

    return {
      success: true,
      message: "บันทึก " + (actionType === 'CHECK_IN' ? "เข้างาน" : "ออกงาน") + " สำเร็จ (" + timeStr + ")",
      log: logData,
      attendance: logData,
      logs: [logData]
    };
  } catch (e) {
    return { success: false, message: e.toString() };
  }
}

/**
 * ดึงรายการลงเวลาเข้า-ออกงานของวันที่ระบุ (ค่าเริ่มต้น: วันนี้)
 */
function getTodayAttendance(targetDate) {
  try {
    var ss = getDatabaseSpreadsheet();
    var sheet = ss.getSheetByName(SHEET_ATTENDANCE);
    if (!sheet) return { success: true, attendance: [], logs: [], count: 0 };

    var now = new Date();
    var dateQuery = targetDate || Utilities.formatDate(now, "GMT+7", "yyyy-MM-dd");
    // Normalize dateQuery if passed as dd/MM/yyyy
    if (dateQuery && /^\d{1,2}\/\d{1,2}\/\d{4}$/.test(dateQuery)) {
      var qp = dateQuery.split('/');
      dateQuery = qp[2] + '-' + qp[1].padStart(2, '0') + '-' + qp[0].padStart(2, '0');
    }

    var data = sheet.getDataRange().getValues();
    var list = [];

    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      if (!row[0] && !row[3]) continue;

      // Normalize Date
      var rDate = "";
      if (row[1] instanceof Date) {
        rDate = Utilities.formatDate(row[1], "GMT+7", "yyyy-MM-dd");
      } else {
        rDate = String(row[1] || "").trim();
        if (/^\d{1,2}\/\d{1,2}\/\d{4}$/.test(rDate)) {
          var dp = rDate.split('/');
          rDate = dp[2] + '-' + dp[1].padStart(2, '0') + '-' + dp[0].padStart(2, '0');
        }
      }

      if (rDate === dateQuery) {
        // Normalize Time cleanly to HH:mm (avoids 1899 LMT +06:42 17-minute shift)
        var rTime = normalizeTimeHHMM_(row[6]);

        var actType = String(row[5] || "CHECK_IN").toUpperCase();
        var wStatus = String(row[7] || "ตรงเวลา").trim();
        var sched = String(row[9] || "").trim();

        list.push({
          id: String(row[0] || ""),
          date: rDate,
          profileId: String(row[2] || ""),
          name: String(row[3] || ""),
          personType: String(row[4] || "พนักงาน"),
          actionType: actType,
          logType: actType,
          time: rTime,
          workStatus: wStatus,
          status: wStatus,
          workSchedule: sched,
          timestamp: String(row[8] || "")
        });
      }
    }

    return {
      success: true,
      date: dateQuery,
      attendance: list,
      logs: list,
      count: list.length
    };
  } catch (e) {
    return { success: false, message: e.toString(), attendance: [], logs: [], count: 0 };
  }
}

/**
 * ==============================================================================
 * ระบบติดตามสถานะเครื่องออนไลน์และรายการตรวจจับใบหน้าแบบ Real-Time (CacheService)
 * ==============================================================================
 */
var CLIENT_PRESENCE_INDEX_KEY = "ACTIVE_CLIENT_SESSIONS_V1";
var CLIENT_SESSION_TTL_MS = 28000;
var GLOBAL_RELOAD_CMD_KEY = "GLOBAL_FORCE_RELOAD_CMD_V1";
var REALTIME_SCAN_LOGS_KEY = "REALTIME_SCAN_LOGS_V1";
var REALTIME_SCAN_LOGS_CLEARED_KEY = "REALTIME_SCAN_LOGS_CLEARED_V1";
var MAX_REALTIME_SCAN_LOGS = 80;

/**
 * ดึงคำสั่ง Force Reload ล่าสุดจาก CacheService (หรือสำรองจาก ScriptProperties)
 */
function getLatestReloadCommand_() {
  try {
    var cache = CacheService.getScriptCache();
    var raw = cache.get(GLOBAL_RELOAD_CMD_KEY);
    if (raw) {
      return JSON.parse(raw);
    }
    var props = PropertiesService.getScriptProperties();
    var pRaw = props.getProperty(GLOBAL_RELOAD_CMD_KEY);
    if (pRaw) {
      var cmd = JSON.parse(pRaw);
      if (cmd && (Date.now() - (cmd.timestamp || 0) < 21600000)) {
        cache.put(GLOBAL_RELOAD_CMD_KEY, pRaw, 21600);
        return cmd;
      }
    }
  } catch (e) {}
  return null;
}

/**
 * สั่งให้ทุกเครื่องที่ออนไลน์อยู่ทำการ Hard Reload (Ctrl+F5)
 */
function triggerGlobalForceReload(payload) {
  try {
    var p = payload || {};
    var cache = CacheService.getScriptCache();
    var nowMs = Date.now();
    var timeStr = Utilities.formatDate(new Date(nowMs), "GMT+7", "dd/MM/yyyy HH:mm:ss");
    var reloadId = "RLD_" + nowMs + "_" + Math.random().toString(36).substr(2, 4).toUpperCase();

    var reloadCmd = {
      reloadId: reloadId,
      timestamp: nowMs,
      timeStr: timeStr,
      triggeredBy: String(p.username || p.triggeredBy || "admin"),
      includeSelf: p.includeSelf !== false,
      reason: String(p.reason || "ผู้ดูแลระบบสั่งอัปเดตและรีโหลดหน้าจอใหม่ (Ctrl+F5)")
    };

    var cmdJson = JSON.stringify(reloadCmd);
    cache.put(GLOBAL_RELOAD_CMD_KEY, cmdJson, 21600); // แคช 6 ชั่วโมง

    try {
      var props = PropertiesService.getScriptProperties();
      props.setProperty(GLOBAL_RELOAD_CMD_KEY, cmdJson);
    } catch (eProp) {}

    // คำนวณจำนวนเครื่องที่ออนไลน์ขณะนี้
    var presenceRes = getOnlineClientsInfo_();

    return {
      success: true,
      message: "ส่งคำสั่ง Hard Reload ไปยังทุกเครื่องสำเร็จ",
      reloadCommand: reloadCmd,
      onlineCount: presenceRes.onlineCount,
      timestamp: nowMs
    };
  } catch (e) {
    return { success: false, message: e.toString() };
  }
}

/**
 * Helper ดึงข้อมูลและจำนวนเครื่องที่ออนไลน์ (นับทั้งจำนวนผู้ใช้ออนไลน์ และจำนวนที่เปิดกล้องตรวจจับ)
 */
function getOnlineClientsInfo_() {
  var cache = CacheService.getScriptCache();
  var nowMs = Date.now();
  var rawPresence = cache.get(CLIENT_PRESENCE_INDEX_KEY);
  var presenceMap = {};
  if (rawPresence) {
    try { presenceMap = JSON.parse(rawPresence) || {}; } catch (e) {}
  }

  var activeUsers = [];
  var changed = false;
  var keys = Object.keys(presenceMap);
  var activeCamerasCount = 0;

  for (var i = 0; i < keys.length; i++) {
    var k = keys[i];
    var item = presenceMap[k];
    if (!item || (nowMs - (item.lastSeen || 0) > CLIENT_SESSION_TTL_MS)) {
      delete presenceMap[k];
      changed = true;
    } else {
      if (item.cameraStarted) {
        activeCamerasCount++;
      }
      activeUsers.push(item);
    }
  }

  if (changed) {
    cache.put(CLIENT_PRESENCE_INDEX_KEY, JSON.stringify(presenceMap), 300);
  }

  // เรียงลำดับ: เครื่องที่เปิดกล้องขึ้นก่อน ตามด้วยเวลาล่าสุด
  activeUsers.sort(function(a, b) {
    if (!!b.cameraStarted !== !!a.cameraStarted) {
      return (b.cameraStarted ? 1 : 0) - (a.cameraStarted ? 1 : 0);
    }
    return (b.lastSeen || 0) - (a.lastSeen || 0);
  });

  return {
    onlineCount: activeUsers.length,
    activeCameraCount: activeCamerasCount,
    users: activeUsers,
    clients: activeUsers
  };
}

/**
 * รับ Heartbeat จากเครื่องผู้ใช้งานทั่วไป (Client Presence Pulse)
 */
function syncClientPresenceAndCommands(payload) {
  try {
    var p = payload || {};
    var cache = CacheService.getScriptCache();
    var nowMs = Date.now();
    var sessionId = String(p.sessionId || "").trim();
    if (!sessionId) {
      return { success: false, message: "Missing sessionId" };
    }

    var rawPresence = cache.get(CLIENT_PRESENCE_INDEX_KEY);
    var presenceMap = {};
    if (rawPresence) {
      try { presenceMap = JSON.parse(rawPresence) || {}; } catch (e) {}
    }

    // ล้างตัวหมดอายุ
    var keys = Object.keys(presenceMap);
    for (var i = 0; i < keys.length; i++) {
      var k = keys[i];
      if (!presenceMap[k] || (nowMs - (presenceMap[k].lastSeen || 0) > CLIENT_SESSION_TTL_MS)) {
        delete presenceMap[k];
      }
    }

    presenceMap[sessionId] = {
      sessionId: sessionId,
      username: String(p.username || "user"),
      fullName: String(p.fullName || p.username || "ผู้ใช้งาน"),
      role: String(p.role || "operator"),
      cameraStarted: !!p.cameraStarted,
      currentTab: String(p.currentTab || "scanner"),
      lastSeen: nowMs
    };

    cache.put(CLIENT_PRESENCE_INDEX_KEY, JSON.stringify(presenceMap), 300);

    var latestReload = getLatestReloadCommand_();
    var onlineInfo = getOnlineClientsInfo_();

    return {
      success: true,
      onlineCount: onlineInfo.onlineCount,
      activeCameraCount: onlineInfo.activeCameraCount,
      clients: onlineInfo.clients,
      reloadCommand: latestReload,
      timestamp: nowMs
    };
  } catch (e) {
    return { success: false, message: e.toString() };
  }
}

/**
 * ดึงรายการตรวจจับใบหน้าแบบ Real-Time จาก CacheService Ring Buffer
 * หาก Cache ว่าง (Cold start) จะอ่านจากชีต ScanLogs ล่าสุด 40 รายการมาเติมให้อัตโนมัติ
 */
function getRealtimeDetectionLogs(params) {
  try {
    var p = params || {};
    var cache = CacheService.getScriptCache();
    var rawLogs = cache.get(REALTIME_SCAN_LOGS_KEY);
    var rawCleared = cache.get(REALTIME_SCAN_LOGS_CLEARED_KEY);
    var clearedAt = rawCleared ? Number(rawCleared) : 0;
    if (isNaN(clearedAt)) clearedAt = 0;
    var logList = [];
    var isCold = false;

    if (rawLogs !== null) {
      try {
        logList = JSON.parse(rawLogs) || [];
      } catch (eJson) {
        logList = [];
      }
    } else {
      // เฉพาะกรณีที่ Cache ยังไม่เคยถูกสร้าง (rawLogs === null เช่น เริ่มระบบใหม่) จึงจะอ่านจากชีต ScanLogs ล่าสุด 40 แถว
      // หากมีการสั่งล้างจอไปแล้ว (rawLogs !== null แม้จะเป็น "[]") จะไม่ดึงข้อมูลเก่าจากชีตมาเติมซ้ำเด็ดขาด
      isCold = true;
      try {
        var ss = getDatabaseSpreadsheet();
        var sheet = ss.getSheetByName(SHEET_LOGS);
        if (sheet && sheet.getLastRow() > 1) {
          var lastRow = sheet.getLastRow();
          var startRow = Math.max(2, lastRow - 39);
          var numRows = (lastRow - startRow) + 1;
          var numCols = Math.max(4, sheet.getLastColumn());
          var rangeVals = sheet.getRange(startRow, 1, numRows, numCols).getValues();

          for (var r = rangeVals.length - 1; r >= 0; r--) {
            var row = rangeVals[r];
            var dateVal = row[0];
            var dateObj = (dateVal instanceof Date && !isNaN(dateVal.getTime())) ? dateVal : new Date(dateVal);
            var timestampMs = (!isNaN(dateObj.getTime())) ? dateObj.getTime() : Date.now();
            var timeStr = Utilities.formatDate(dateObj, "GMT+7", "HH:mm:ss");
            var dateStr = Utilities.formatDate(dateObj, "GMT+7", "dd/MM/yyyy");

            var pName = String(row[1] || "").trim() || "ลูกค้าทั่วไป";
            var mStatus = String(row[2] || "").trim().toUpperCase() || (pName === "ลูกค้าทั่วไป" ? "UNKNOWN" : "KNOWN");
            var cStr = String(row[3] || "").trim();
            var cPct = 0;
            if (cStr && cStr.indexOf("%") !== -1) {
              cPct = Math.round(parseFloat(cStr)) || 0;
            } else if (!isNaN(Number(cStr)) && Number(cStr) > 0) {
              var n = Number(cStr);
              cPct = n <= 1 ? Math.round(n * 100) : Math.round(n);
            }

            var pType = row.length >= 5 ? String(row[4] || "").trim().toLowerCase() : (mStatus === "KNOWN" ? "employee" : "unknown");
            var detUser = row.length >= 6 ? String(row[5] || "").trim() : "unknown";
            var detName = row.length >= 7 ? String(row[6] || "").trim() : detUser;
            var detDetail = row.length >= 8 ? String(row[7] || "").trim() : (mStatus === "KNOWN" ? "ยืนยันตัวตนสำเร็จ" : "ตรวจพบผู้มาติดต่อ");

            logList.push({
              id: "SEED_" + timestampMs + "_" + r,
              timestamp: timestampMs,
              timeStr: timeStr,
              dateStr: dateStr,
              personName: pName,
              profileId: "",
              personType: pType || (mStatus === "KNOWN" ? "employee" : "unknown"),
              status: mStatus,
              confidencePct: cPct,
              detectedByUsername: detUser || "unknown",
              detectedByFullName: detName || detUser || "ไม่ระบุเครื่อง",
              detectedByRole: "operator",
              sessionId: "",
              detail: detDetail
            });
          }

          if (logList.length > 0) {
            cache.put(REALTIME_SCAN_LOGS_KEY, JSON.stringify(logList), 21600);
          }
        }
      } catch (eSeed) {
        console.warn("ScanLogs seed error:", eSeed);
      }
    }

    // กรองรายการตามเวลาสั่งล้างจอ (clearedAt)
    if (clearedAt > 0 && logList.length > 0) {
      logList = logList.filter(function(item) {
        return (item.timestamp || 0) > clearedAt;
      });
    }

    var sinceTs = Number(p.since || p.sinceTimestamp) || 0;
    var limit = Math.min(100, Math.max(1, Number(p.limit) || 50));
    var filtered = logList;

    if (sinceTs > 0) {
      filtered = logList.filter(function(item) {
        return (item.timestamp || 0) > sinceTs;
      });
    }

    if (filtered.length > limit) {
      filtered = filtered.slice(0, limit);
    }

    var onlineInfo = getOnlineClientsInfo_();
    var latestReload = getLatestReloadCommand_();

    return {
      success: true,
      logs: filtered,
      count: filtered.length,
      totalCached: logList.length,
      isColdStartSeeded: isCold,
      clearedAt: clearedAt,
      onlineCount: onlineInfo.onlineCount,
      activeCameraCount: onlineInfo.activeCameraCount,
      onlineClients: onlineInfo.clients,
      reloadCommand: latestReload,
      serverTime: Date.now()
    };
  } catch (e) {
    return {
      success: false,
      message: e.toString(),
      logs: [],
      count: 0
    };
  }
}

/**
 * ล้างข้อมูลรายการตรวจจับบนหน้าจอแสดงผลสด (ล้างเฉพาะ Cache ไม่ลบชีตประวัติถาวร)
 */
function clearRealtimeDetectionLogs() {
  try {
    var cache = CacheService.getScriptCache();
    var now = Date.now();
    cache.put(REALTIME_SCAN_LOGS_KEY, JSON.stringify([]), 21600);
    cache.put(REALTIME_SCAN_LOGS_CLEARED_KEY, String(now), 21600);
    return {
      success: true,
      message: "ล้างรายการตรวจจับบนหน้าจอแสดงผลสดเรียบร้อยแล้ว (ประวัติในชีตยังคงอยู่ครบถ้วน)",
      clearedAt: now,
      timestamp: now
    };
  } catch (e) {
    return { success: false, message: e.toString() };
  }
}



