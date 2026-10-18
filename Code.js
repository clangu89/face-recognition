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
      result = getTodayAttendance(e.parameter.date);
    } else if (action === "recordAttendance") {
      result = recordAttendanceLog(e.parameter);
    } else if (action === "updateProfileType") {
      result = updateProfileType(e.parameter.id, e.parameter.personType);
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
    if (!logSheet) {
      logSheet = ss.insertSheet(SHEET_LOGS);
      logSheet.appendRow(["Timestamp", "Person_Name", "Match_Status", "Confidence"]);
      logSheet.getRange(1, 1, 1, 4)
        .setBackground("#1E293B")
        .setFontColor("#FFFFFF")
        .setFontWeight("bold");
      logSheet.setFrozenRows(1);
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
      settingsSheet.appendRow(["WORK_START_TIME", "08:30", "เวลาเริ่มเข้างานพนักงาน (HH:mm)", nowStr]);
      settingsSheet.appendRow(["WORK_END_TIME", "17:30", "เวลาเลิกงานพนักงาน (HH:mm)", nowStr]);
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
        "Timestamp"
      ]);
      attSheet.getRange(1, 1, 1, 9)
        .setBackground("#1E293B")
        .setFontColor("#FFFFFF")
        .setFontWeight("bold");
      attSheet.setFrozenRows(1);
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
function updateProfileType(profileId, personType) {
  try {
    if (!profileId) return { success: false, message: "ไม่พบ Profile ID" };
    var cleanType = (String(personType).toLowerCase() === 'vip') ? 'vip' : 'employee';
    var ss = getDatabaseSpreadsheet();
    var sheet = ss.getSheetByName(SHEET_PROFILES);
    var data = sheet.getDataRange().getValues();

    for (var i = 1; i < data.length; i++) {
      if (String(data[i][0]).trim() === String(profileId).trim()) {
        sheet.getRange(i + 1, 8).setValue(cleanType);
        return {
          success: true,
          message: "อัปเดตประเภทเป็น " + (cleanType === 'vip' ? 'VIP' : 'พนักงาน') + " สำเร็จ",
          profileId: profileId,
          personType: cleanType
        };
      }
    }
    return { success: false, message: "ไม่พบรหัสโปรไฟล์" };
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
 * บันทึกประวัติการตรวจจับใบหน้า (Log)
 */
function logScanEvent(name, status, confidence) {
  try {
    var ss = getDatabaseSpreadsheet();
    var sheet = ss.getSheetByName(SHEET_LOGS);
    if (!sheet) return { success: false };
    
    sheet.appendRow([
      new Date(),
      name || "ลูกค้าทั่วไป",
      status || "UNKNOWN",
      confidence ? (confidence * 100).toFixed(1) + "%" : "-"
    ]);

    return { success: true };
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
      result = logScanEvent(data.name, data.status, data.confidence);
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
      result = getTodayAttendance(data.date || (e.parameter && e.parameter.date));
    } else if (action === "recordAttendance") {
      result = recordAttendanceLog(data.payload || data);
    } else if (action === "updateProfileType") {
      result = updateProfileType(data.id || (e.parameter && e.parameter.id), data.personType || (e.parameter && e.parameter.personType));
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
 * ดึงการตั้งค่าระบบและข้อความทักทายที่เป็นค่าคงที่ถาวร
 */
function getSystemSettings() {
  try {
    var ss = getDatabaseSpreadsheet();
    var sheet = getOrCreateSystemSettingsSheet(ss);
    var data = sheet.getDataRange().getValues();

    var settings = {
      knownGreetingTemplate: "สวัสดี{particle}คุณ {name}",
      unknownGreetingText: "สวัสดี{particle}คุณลูกค้า วันนี้รับอะไรดี{particle}",
      voiceParticle: "ค่ะ",
      bellChimeEnabled: true,
      workStartTime: "08:30",
      workEndTime: "17:30"
    };

    if (data && data.length > 1) {
      for (var i = 1; i < data.length; i++) {
        var key = String(data[i][0] || "").trim().toUpperCase();
        var val = data[i][1];
        if (key === "GREETING_KNOWN_TEMPLATE" && val !== undefined && val !== null && String(val).trim() !== "") {
          settings.knownGreetingTemplate = String(val).trim();
        } else if (key === "GREETING_UNKNOWN_TEXT" && val !== undefined && val !== null && String(val).trim() !== "") {
          settings.unknownGreetingText = String(val).trim();
        } else if (key === "VOICE_PARTICLE" && val !== undefined && val !== null && String(val).trim() !== "") {
          settings.voiceParticle = String(val).trim();
        } else if (key === "BELL_CHIME_ENABLED" && val !== undefined && val !== null) {
          settings.bellChimeEnabled = (String(val).trim().toLowerCase() === "true");
        } else if (key === "WORK_START_TIME" && val !== undefined && val !== null && String(val).trim() !== "") {
          settings.workStartTime = String(val).trim();
        } else if (key === "WORK_END_TIME" && val !== undefined && val !== null && String(val).trim() !== "") {
          settings.workEndTime = String(val).trim();
        }
      }
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
        knownGreetingTemplate: "สวัสดี{particle}คุณ {name}",
        unknownGreetingText: "สวัสดี{particle}คุณลูกค้า วันนี้รับอะไรดี{particle}",
        voiceParticle: "ค่ะ",
        bellChimeEnabled: true,
        workStartTime: "08:30",
        workEndTime: "17:30"
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
    if (payload.knownGreetingTemplate !== undefined) {
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
    if (payload.workStartTime !== undefined) {
      updateMap["WORK_START_TIME"] = {
        value: String(payload.workStartTime).trim(),
        desc: "เวลาเริ่มเข้างานพนักงาน (HH:mm)"
      };
    }
    if (payload.workEndTime !== undefined) {
      updateMap["WORK_END_TIME"] = {
        value: String(payload.workEndTime).trim(),
        desc: "เวลาเลิกงานพนักงาน (HH:mm)"
      };
    }

    // อัปเดตแถวที่มีอยู่เดิม
    var foundKeys = {};
    for (var i = 1; i < data.length; i++) {
      var rowKey = String(data[i][0] || "").trim().toUpperCase();
      if (updateMap[rowKey]) {
        sheet.getRange(i + 1, 2).setValue(updateMap[rowKey].value);
        sheet.getRange(i + 1, 4).setValue(nowStr);
        foundKeys[rowKey] = true;
      }
    }

    // เพิ่มคีย์ใหม่ที่ยังไม่มีในตาราง
    for (var key in updateMap) {
      if (!foundKeys[key]) {
        sheet.appendRow([key, updateMap[key].value, updateMap[key].desc, nowStr]);
      }
    }

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
    var timeStr = payload.time || Utilities.formatDate(now, "GMT+7", "HH:mm:ss");
    var actionType = (String(payload.actionType).toUpperCase() === 'CHECK_OUT') ? 'CHECK_OUT' : 'CHECK_IN';
    var workStatus = payload.workStatus || (actionType === 'CHECK_IN' ? 'ตรงเวลา' : 'เลิกงานตามเวลา');
    var logId = "ATT_" + Date.now() + "_" + Math.random().toString(36).substr(2, 4).toUpperCase();
    var personName = String(payload.name).trim();
    var profileId = String(payload.profileId || "").trim();
    var personType = String(payload.personType || "พนักงาน").trim();

    sheet.appendRow([
      logId,
      dateStr,
      profileId,
      personName,
      personType,
      actionType,
      timeStr,
      workStatus,
      nowStr
    ]);

    return {
      success: true,
      message: "บันทึก " + (actionType === 'CHECK_IN' ? "เข้างาน" : "ออกงาน") + " สำเร็จ (" + timeStr + ")",
      log: {
        id: logId,
        date: dateStr,
        profileId: profileId,
        name: personName,
        actionType: actionType,
        time: timeStr,
        workStatus: workStatus
      }
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
    if (!sheet) return { success: true, attendance: [] };

    var now = new Date();
    var dateQuery = targetDate || Utilities.formatDate(now, "GMT+7", "yyyy-MM-dd");
    var data = sheet.getDataRange().getValues();
    var list = [];

    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      var rDate = String(row[1] || "").trim();
      if (row[1] instanceof Date) {
        rDate = Utilities.formatDate(row[1], "GMT+7", "yyyy-MM-dd");
      }
      if (rDate === dateQuery) {
        list.push({
          id: String(row[0] || ""),
          date: rDate,
          profileId: String(row[2] || ""),
          name: String(row[3] || ""),
          personType: String(row[4] || ""),
          actionType: String(row[5] || ""),
          time: String(row[6] || ""),
          workStatus: String(row[7] || ""),
          timestamp: String(row[8] || "")
        });
      }
    }

    return { success: true, date: dateQuery, attendance: list };
  } catch (e) {
    return { success: false, message: e.toString(), attendance: [] };
  }
}


