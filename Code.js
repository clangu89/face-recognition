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
        "Last_Seen_At"
      ]);
      profileSheet.getRange(1, 1, 1, 7)
        .setBackground("#1E293B")
        .setFontColor("#FFFFFF")
        .setFontWeight("bold");
      profileSheet.setFrozenRows(1);
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

    // 3. ตรวจสอบหรือสร้างโฟลเดอร์สำหรับเก็บภาพใบหน้าใน Google Drive
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

    // บันทึกลง Sheet
    var descriptorJson = JSON.stringify(payload.descriptor);
    sheet.appendRow([
      profileId,
      payload.name.trim(),
      descriptorJson,
      imageUrl,
      fileId,
      now,
      now
    ]);

    return {
      success: true,
      message: "บันทึกโปรไฟล์ " + payload.name + " สำเร็จ",
      profile: {
        id: profileId,
        name: payload.name.trim(),
        descriptors: payload.descriptor,
        imageUrl: imageUrl,
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
        sheet.deleteRow(i + 1);
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
 * รองรับ POST Request จาก Frontend ภายนอก
 */
function doPost(e) {
  try {
    var data = {};
    if (e.postData && e.postData.contents) {
      data = JSON.parse(e.postData.contents);
    }
    var action = (e.parameter && e.parameter.action) || data.action;
    var result = { success: false, message: "Invalid action" };

    if (action === "saveProfile") {
      result = saveFaceProfile(data.payload || data);
    } else if (action === "deleteProfile") {
      result = deleteFaceProfile(data.id || (e.parameter && e.parameter.id));
    } else if (action === "logScan") {
      result = logScanEvent(data.name, data.status, data.confidence);
    } else if (action === "getProfiles") {
      result = getFaceProfiles();
    }

    return ContentService.createTextOutput(JSON.stringify(result))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (error) {
    return ContentService.createTextOutput(JSON.stringify({ success: false, message: error.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

