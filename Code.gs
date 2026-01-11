/**
 * ========================================
 * Booking Control System - Main Backend
 * ========================================
 * Google Apps Script Backend for Adventure Tour Booking Management System
 *
 * @author Antigravity AI
 * @version 1.0
 * @date 2025-12-29
 */

// ========================================
// WEB APP ENTRY POINT
// ========================================

/**
 * doGet - Handle HTTP GET requests
 * รองรับ URL parameters เช่น ?page=dashboard
 */
function doGet(e) {
  // Get page parameter from URL (e.g., ?page=dashboard)
  const page = e.parameter.page || null;

  // Create template
  const template = HtmlService.createTemplateFromFile("index");
  template.initialPage = page; // Pass page to template

  return template
    .evaluate()
    .setTitle("Booking Control")
    .setFaviconUrl("https://img.icons8.com/fluency/48/calendar.png")
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag("viewport", "width=device-width, initial-scale=1");
}

/**
 * Include HTML files (for server-side includes)
 * ใช้สำหรับ <?!= include('filename'); ?>
 */
function include(filename) {
  try {
    return HtmlService.createTemplateFromFile(filename).evaluate().getContent();
  } catch (e) {
    Logger.log("Error including file " + filename + ": " + e.toString());
    return "<!-- Error including " + filename + " -->";
  }
}

// ========================================
// CONFIGURATION - Fixed Sheet ID
// ========================================

const CONFIG = {
  SPREADSHEET_ID: "1VHWoJ3UyBTUWLXVRBZFu4iURzRZY_H3HDQRRJfqfq8k", // ⚠️ ใส่ Spreadsheet ID ของคุณที่นี่

  // Drive Folder IDs
  BOOKING_SLIPS_FOLDER_ID: "1lxu8dGE5wBkphXRJtYGU84w41pIxRaby", // Folder สำหรับสลิปการจอง
  REFUND_SLIPS_FOLDER_ID: "12VmXp8TrEWFS2WiOW1S0ng4_qFSQv9PA", // Folder สำหรับสลิปการคืนเงิน

  // Sheet Names
  SHEETS: {
    USERS: "Users",
    BOOKING_RAW: "Booking_Raw",
    LOCATIONS: "Locations",
    PROGRAMS: "Programs",
    BOOKING_STATUS_HISTORY: "Booking_Status_History",
    CUSTOMER: "Customer",
    REFUND: "Refund",
  },

  // User Roles
  ROLES: {
    SALE: "Sale",
    OP: "OP",
    ADMIN: "Admin",
    AR_AP: "AR_AP",
    COST: "Cost",
    OWNER: "Owner",
  },

  // Booking Status
  STATUS: {
    CONFIRM: "Confirm",
    COMPLETE: "Completed",
    CANCEL: "Cancel",
  },

  // Default Password
  DEFAULT_PASSWORD: "password123",
};

// ========================================
// UTILITY FUNCTIONS
// ========================================

/**
 * Get Spreadsheet by ID
 */
function getSpreadsheet() {
  return SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
}

/**
 * Get Sheet by Name
 */
function getSheet(sheetName) {
  return getSpreadsheet().getSheetByName(sheetName);
}

/**
 * Hash Password using SHA-256
 */
function hashPassword(password) {
  const rawHash = Utilities.computeDigest(
    Utilities.DigestAlgorithm.SHA_256,
    password,
    Utilities.Charset.UTF_8
  );
  return rawHash
    .map((byte) => {
      const v = byte < 0 ? 256 + byte : byte;
      return ("0" + v.toString(16)).slice(-2);
    })
    .join("");
}

/**
 * Generate Unique ID
 */
function generateUniqueId(prefix = "") {
  const timestamp = new Date().getTime();
  const random = Math.floor(Math.random() * 10000);
  return `${prefix}${timestamp}${random}`;
}

/**
 * Get Current Timestamp
 */
function getCurrentTimestamp() {
  return Utilities.formatDate(
    new Date(),
    Session.getScriptTimeZone(),
    "yyyy-MM-dd HH:mm:ss"
  );
}

/**
 * Format Date to String
 */
function formatDate(date, format = "dd/MM/yyyy") {
  if (!date) return "";
  return Utilities.formatDate(
    new Date(date),
    Session.getScriptTimeZone(),
    format
  );
}

// ========================================
// SESSION MANAGEMENT (Client-Side)
// ========================================
// หมายเหตุ: Session จะถูกเก็บใน Browser (localStorage) แทน Server
// เพื่อให้สามารถ Login หลาย Browser ด้วย User คนละคนได้

/**
 * Validate Session Token
 * ตรวจสอบ Session Token ที่ส่งมาจาก Client
 */
function validateSession(sessionToken) {
  if (!sessionToken) return null;

  try {
    // Decode session token (Base64)
    const decoded = Utilities.newBlob(
      Utilities.base64Decode(sessionToken)
    ).getDataAsString();

    const sessionData = JSON.parse(decoded);

    // ตรวจสอบว่า Session หมดอายุหรือไม่ (24 ชั่วโมง)
    const now = new Date().getTime();
    const sessionAge = now - sessionData.loginTime;
    const maxAge = 24 * 60 * 60 * 1000; // 24 hours

    if (sessionAge > maxAge) {
      return null; // Session หมดอายุ
    }

    // ตรวจสอบว่า User ยังมีอยู่และ Active
    const sheet = getSheet(CONFIG.SHEETS.USERS);
    const data = sheet.getDataRange().getValues();

    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      if (row[0] === sessionData.userId && row[7] === "เปิดใช้งาน") {
        return {
          userId: sessionData.userId,
          username: sessionData.username,
          role: sessionData.role,
          fullName: row[4],
          phone: row[5],
          loginTime: sessionData.loginTime,
        };
      }
    }

    return null; // User ไม่พบหรือไม่ Active
  } catch (error) {
    Logger.log("Session validation error: " + error.message);
    return null;
  }
}

/**
 * Create Session Token
 * สร้าง Session Token สำหรับส่งกลับไปยัง Client
 */
function createSessionToken(userId, username, role) {
  const sessionData = {
    userId: userId,
    username: username,
    role: role,
    loginTime: new Date().getTime(),
  };

  // Encode to Base64
  const jsonString = JSON.stringify(sessionData);
  const encoded = Utilities.base64Encode(jsonString);

  return encoded;
}

/**
 * Check User Role (รับ sessionToken จาก Client)
 */
function hasRoleWithToken(sessionToken, requiredRole) {
  const session = validateSession(sessionToken);
  if (!session) return false;

  // Owner has access to everything
  if (session.role === CONFIG.ROLES.OWNER) return true;

  // Check specific role
  if (Array.isArray(requiredRole)) {
    return requiredRole.includes(session.role);
  }
  return session.role === requiredRole;
}

/**
 * Get Current User (Optimized)
 * ดึงข้อมูล User ปัจจุบันจาก Session Token
 */
function getCurrentUser(sessionToken) {
  try {
    if (!sessionToken) {
      return {
        success: false,
        message: "ไม่พบ Session Token",
      };
    }

    // Validate session (already checks user exists and is active)
    const session = validateSession(sessionToken);

    if (!session) {
      return {
        success: false,
        message: "Session หมดอายุ กรุณาเข้าสู่ระบบใหม่",
      };
    }

    // Return user data from validated session
    return {
      success: true,
      data: {
        userId: session.userId,
        username: session.username,
        fullName: session.fullName,
        phone: session.phone,
        role: session.role,
        loginTime: session.loginTime,
      },
    };
  } catch (error) {
    return {
      success: false,
      message: "เกิดข้อผิดพลาด: " + error.message,
    };
  }
}

// ========================================
// LEGACY SESSION FUNCTIONS (สำหรับ Backward Compatibility)
// ========================================
// ฟังก์ชันเหล่านี้ยังคงทำงานได้ชั่วคราวโดยใช้ ScriptProperties
// เพื่อไม่ให้โค้ดเก่าเสีย แต่แนะนำให้ใช้ Client-Side Session

/**
 * @deprecated ใช้ Client-Side Session แทน
 * ฟังก์ชันนี้ยังทำงานได้ชั่วคราวเพื่อ Backward Compatibility
 */
function setSession(userId, username, role) {
  try {
    const scriptProperties = PropertiesService.getScriptProperties();
    const sessionData = {
      userId: userId,
      username: username,
      role: role,
      loginTime: new Date().getTime(),
    };
    // ใช้ username เป็น key เพื่อรองรับหลาย session
    scriptProperties.setProperty(
      "session_" + username,
      JSON.stringify(sessionData)
    );
  } catch (error) {
    Logger.log("setSession error: " + error.message);
  }
}

/**
 * @deprecated ใช้ validateSession(sessionToken) แทน
 * ฟังก์ชันนี้ยังทำงานได้ชั่วคราวเพื่อ Backward Compatibility
 */
function getSession() {
  try {
    const scriptProperties = PropertiesService.getScriptProperties();
    const allProperties = scriptProperties.getProperties();

    // หา session ที่ใหม่ที่สุด
    let latestSession = null;
    let latestTime = 0;

    for (let key in allProperties) {
      if (key.startsWith("session_")) {
        try {
          const sessionData = JSON.parse(allProperties[key]);
          if (sessionData.loginTime > latestTime) {
            latestTime = sessionData.loginTime;
            latestSession = sessionData;
          }
        } catch (e) {
          // Skip invalid session
        }
      }
    }

    return latestSession;
  } catch (error) {
    Logger.log("getSession error: " + error.message);
    return null;
  }
}

/**
 * @deprecated ใช้ Client-Side Session แทน
 * ฟังก์ชันนี้ยังทำงานได้ชั่วคราวเพื่อ Backward Compatibility
 */
function clearSession() {
  try {
    const scriptProperties = PropertiesService.getScriptProperties();
    const allProperties = scriptProperties.getProperties();

    // ลบ session ทั้งหมด
    for (let key in allProperties) {
      if (key.startsWith("session_")) {
        scriptProperties.deleteProperty(key);
      }
    }
  } catch (error) {
    Logger.log("clearSession error: " + error.message);
  }
}

/**
 * @deprecated ใช้ validateSession(sessionToken) แทน
 * ฟังก์ชันนี้ยังทำงานได้ชั่วคราวเพื่อ Backward Compatibility
 */
function isLoggedIn() {
  return getSession() !== null;
}

/**
 * @deprecated ใช้ hasRoleWithToken(sessionToken, requiredRole) แทน
 * ฟังก์ชันนี้ยังทำงานได้ชั่วคราวเพื่อ Backward Compatibility
 */
function hasRole(requiredRole) {
  const session = getSession();
  if (!session) return false;

  // Owner has access to everything
  if (session.role === CONFIG.ROLES.OWNER) return true;

  // Check specific role
  if (Array.isArray(requiredRole)) {
    return requiredRole.includes(session.role);
  }
  return session.role === requiredRole;
}

// ========================================
// AUTHENTICATION FUNCTIONS
// ========================================

/**
 * Login User
 */
function loginUser(username, password) {
  try {
    const sheet = getSheet(CONFIG.SHEETS.USERS);
    const data = sheet.getDataRange().getValues();

    // Skip header row
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const dbUsername = row[2]; // Column C: ชื่อผู้ใช้
      const dbPassword = row[3]; // Column D: รหัสผ่าน
      const fullName = row[4]; // Column E: ชื่อ-นามสกุล
      const phone = row[5]; // Column F: เบอร์โทร
      const role = row[6]; // Column G: บทบาท
      const status = row[7]; // Column H: สถานะการใช้งาน

      if (dbUsername === username && status === "เปิดใช้งาน") {
        const hashedPassword = hashPassword(password);
        if (dbPassword === hashedPassword) {
          const userId = row[0]; // Column A: รหัสผู้ใช้

          // สร้าง Session Token สำหรับ Client-Side
          const sessionToken = createSessionToken(userId, username, role);

          // ตั้งค่า Session แบบเก่าด้วย (สำหรับ Backward Compatibility)
          setSession(userId, username, role);

          return {
            success: true,
            message: "เข้าสู่ระบบสำเร็จ",
            sessionToken: sessionToken, // ส่ง Token ที่ root level
            data: {
              userId: userId,
              username: username,
              fullName: fullName,
              phone: phone,
              role: role,
            },
          };
        }
      }
    }

    return {
      success: false,
      message: "ชื่อผู้ใช้หรือรหัสผ่านไม่ถูกต้อง",
    };
  } catch (error) {
    return {
      success: false,
      message: "เกิดข้อผิดพลาด: " + error.message,
    };
  }
}

/**
 * Logout User
 * (Client จะลบ Session Token เอง)
 */
function logoutUser() {
  return {
    success: true,
    message: "ออกจากระบบสำเร็จ",
  };
}

/**
 * Get Current User Info (รับ sessionToken จาก Client)
 */
function getCurrentUser(sessionToken) {
  const session = validateSession(sessionToken);
  if (!session) {
    return {
      success: false,
      message: "ไม่พบข้อมูลผู้ใช้หรือ Session หมดอายุ",
    };
  }

  return {
    success: true,
    data: session,
  };
}

/**
 * Setup Initial Owner User (Run this once)
 * ฟังก์ชันนี้ใช้สำหรับสร้าง User Owner คนแรกในระบบ
 * รันครั้งเดียวเมื่อเริ่มต้นใช้งานระบบ
 */
function setupInitialOwner() {
  try {
    const sheet = getSheet(CONFIG.SHEETS.USERS);

    // Check if sheet is empty (only header or no data)
    const data = sheet.getDataRange().getValues();
    if (data.length > 1) {
      Logger.log("มี User อยู่ในระบบแล้ว");
      return {
        success: false,
        message: "มี User อยู่ในระบบแล้ว ไม่สามารถสร้าง Owner ใหม่ได้",
      };
    }

    // Create header if not exists
    if (data.length === 0) {
      const headers = [
        "รหัสผู้ใช้",
        "อีเมล",
        "ชื่อผู้ใช้",
        "รหัสผ่าน",
        "ชื่อ-นามสกุล",
        "เบอร์โทร",
        "บทบาท",
        "สถานะการใช้งาน",
        "วันที่สร้าง",
        "วันที่แก้ไขล่าสุด",
      ];
      sheet.appendRow(headers);
    }

    // Create Owner user
    const userId = "USR001";
    const hashedPassword = hashPassword("password123");
    const timestamp = getCurrentTimestamp();

    const ownerRow = [
      userId,
      "owner@example.com",
      "owner",
      hashedPassword,
      "เจ้าของ",
      "", // เบอร์โทร (เว้นว่าง)
      "Owner",
      "เปิดใช้งาน",
      timestamp,
      timestamp,
    ];

    sheet.appendRow(ownerRow);

    Logger.log("สร้าง Owner สำเร็จ!");
    Logger.log("Username: owner");
    Logger.log("Password: password123");

    return {
      success: true,
      message: "สร้าง Owner สำเร็จ!\nUsername: owner\nPassword: password123",
      data: {
        username: "admin",
        password: "password123",
      },
    };
  } catch (error) {
    Logger.log("Error: " + error.message);
    return {
      success: false,
      message: "เกิดข้อผิดพลาด: " + error.message,
    };
  }
}

/**
 * Setup All Sheets with Headers
 * สร้าง Headers สำหรับทุก Sheet
 */
function setupAllSheets() {
  try {
    const ss = getSpreadsheet();

    // Setup Users Sheet
    let usersSheet = ss.getSheetByName(CONFIG.SHEETS.USERS);
    if (!usersSheet) {
      usersSheet = ss.insertSheet(CONFIG.SHEETS.USERS);
    }
    if (usersSheet.getLastRow() === 0) {
      usersSheet.appendRow([
        "รหัสผู้ใช้",
        "อีเมล",
        "ชื่อผู้ใช้",
        "รหัสผ่าน",
        "ชื่อ-นามสกุล",
        "เบอร์โทร",
        "บทบาท",
        "สถานะการใช้งาน",
        "วันที่สร้าง",
        "วันที่แก้ไขล่าสุด",
      ]);
      // Set Column F (Phone) as Plain Text to keep leading zero
      usersSheet.getRange("F:F").setNumberFormat("@");
    }

    // Setup Booking_Raw Sheet
    let bookingSheet = ss.getSheetByName(CONFIG.SHEETS.BOOKING_RAW);
    if (!bookingSheet) {
      bookingSheet = ss.insertSheet(CONFIG.SHEETS.BOOKING_RAW);
    }
    if (bookingSheet.getLastRow() === 0) {
      bookingSheet.appendRow([
        "รหัสการจอง",
        "วันที่จอง",
        "วันที่เดินทาง",
        "ชื่อสถานที่",
        "โปรแกรม",
        "ผู้ใหญ่ (คน)",
        "เด็ก (คน)",
        "ราคาผู้ใหญ่",
        "ราคาเด็ก",
        "ค่าใช้จ่ายเพิ่มเติม",
        "ส่วนลด (บาท)",
        "รวม VAT 7%",
        "สถานะ",
        "เงินโอน",
        "เงินสด",
        "Cash on tour",
        "URL สลิปการชำระเงิน",
        "Agent",
        "หมายเหตุ",
        "ยอดขายต่อรายการ",
        "ผู้สร้าง",
        "วันที่สร้าง",
        "ผู้แก้ไขล่าสุด",
        "วันที่แก้ไขล่าสุด",
      ]);
    }

    // Setup Locations Sheet
    let locationsSheet = ss.getSheetByName(CONFIG.SHEETS.LOCATIONS);
    if (!locationsSheet) {
      locationsSheet = ss.insertSheet(CONFIG.SHEETS.LOCATIONS);
    }
    if (locationsSheet.getLastRow() === 0) {
      locationsSheet.appendRow([
        "รหัสสถานที่",
        "ชื่อสถานที่",
        "ชื่อเซลล์",
        "วันที่สร้าง",
        "วันที่แก้ไขล่าสุด",
      ]);
    }

    // Setup Programs Sheet
    let programsSheet = ss.getSheetByName(CONFIG.SHEETS.PROGRAMS);
    if (!programsSheet) {
      programsSheet = ss.insertSheet(CONFIG.SHEETS.PROGRAMS);
    }
    if (programsSheet.getLastRow() === 0) {
      programsSheet.appendRow([
        "รหัสโปรแกรม",
        "รายละเอียด",
        "ราคาผู้ใหญ่",
        "ราคาเด็ก",
        "สถานะการใช้งาน",
        "วันที่สร้าง",
        "วันที่แก้ไขล่าสุด",
      ]);
    }

    // Setup Booking_Status_History Sheet
    let historySheet = ss.getSheetByName(CONFIG.SHEETS.BOOKING_STATUS_HISTORY);
    if (!historySheet) {
      historySheet = ss.insertSheet(CONFIG.SHEETS.BOOKING_STATUS_HISTORY);
    }
    if (historySheet.getLastRow() === 0) {
      historySheet.appendRow([
        "รหัสประวัติ",
        "รหัสการจอง",
        "สถานะเดิม",
        "สถานะใหม่",
        "ผู้เปลี่ยนสถานะ",
        "วันที่เปลี่ยนสถานะ",
        "เหตุผล",
      ]);
    }

    Logger.log("Setup all sheets สำเร็จ!");
    return {
      success: true,
      message: "สร้าง Sheets และ Headers สำเร็จทั้งหมด",
    };
  } catch (error) {
    Logger.log("Error: " + error.message);
    return {
      success: false,
      message: "เกิดข้อผิดพลาด: " + error.message,
    };
  }
}

/**
 * Debug Login - ตรวจสอบปัญหาการ Login
 * รันฟังก์ชันนี้เพื่อดูข้อมูล User ในระบบ
 */
function debugLogin() {
  try {
    Logger.log("=== เริ่มตรวจสอบระบบ ===");

    // 1. ตรวจสอบ Spreadsheet ID
    Logger.log("1. Spreadsheet ID: " + CONFIG.SPREADSHEET_ID);
    if (
      CONFIG.SPREADSHEET_ID === "1VHWoJ3UyBTUWLXVRBZFu4iURzRZY_H3HDQRRJfqfq8k"
    ) {
      Logger.log("❌ ERROR: คุณยังไม่ได้ใส่ SPREADSHEET_ID ใน Code.gs");
      return {
        success: false,
        message: "กรุณาใส่ SPREADSHEET_ID ใน Code.gs",
      };
    }

    // 2. ตรวจสอบ Sheet Users
    const sheet = getSheet(CONFIG.SHEETS.USERS);
    Logger.log("2. Sheet Users: พบแล้ว");

    // 3. ตรวจสอบข้อมูล User
    const data = sheet.getDataRange().getValues();
    Logger.log("3. จำนวนแถวทั้งหมด: " + data.length);

    if (data.length === 0) {
      Logger.log("❌ ERROR: Sheet Users ว่างเปล่า");
      Logger.log("แก้ไข: รันฟังก์ชัน setupAllSheets() และ setupInitialOwner()");
      return {
        success: false,
        message: "Sheet Users ว่างเปล่า กรุณารัน setupInitialOwner()",
      };
    }

    if (data.length === 1) {
      Logger.log("❌ ERROR: มีแค่ Header ไม่มี User");
      Logger.log("แก้ไข: รันฟังก์ชัน setupInitialOwner()");
      return {
        success: false,
        message: "ไม่มี User ในระบบ กรุณารัน setupInitialOwner()",
      };
    }

    // 4. แสดงข้อมูล User ทั้งหมด
    Logger.log("4. User ในระบบ:");
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      Logger.log(`   - User ${i}:`);
      Logger.log(`     รหัส: ${row[0]}`);
      Logger.log(`     Username: ${row[2]}`);
      Logger.log(`     Role: ${row[5]}`);
      Logger.log(`     Status: ${row[6]}`);
      Logger.log(`     Password Hash: ${row[3].substring(0, 20)}...`);
    }

    // 5. ทดสอบ Login
    Logger.log("5. ทดสอบ Login...");
    const testResult = loginUser("admin", "password123");
    Logger.log("   ผลลัพธ์: " + JSON.stringify(testResult));

    if (testResult.success) {
      Logger.log("✅ SUCCESS: Login สำเร็จ!");
    } else {
      Logger.log("❌ ERROR: Login ไม่สำเร็จ - " + testResult.message);
    }

    Logger.log("=== สิ้นสุดการตรวจสอบ ===");

    return {
      success: true,
      message: "ตรวจสอบเสร็จสิ้น ดู Execution log สำหรับรายละเอียด",
      data: {
        totalUsers: data.length - 1,
        loginTest: testResult,
      },
    };
  } catch (error) {
    Logger.log("❌ CRITICAL ERROR: " + error.message);
    Logger.log("Stack: " + error.stack);
    return {
      success: false,
      message: "เกิดข้อผิดพลาดร้ายแรง: " + error.message,
    };
  }
}

/**
 * ทดสอบ Login โดยตรง (สำหรับ Debug)
 * รันฟังก์ชันนี้เพื่อทดสอบการ Login ด้วย username และ password
 */
function testLogin() {
  Logger.log("=== ทดสอบการ Login ===");

  const username = "admin";
  const password = "password123";

  Logger.log("Username: " + username);
  Logger.log("Password: " + password);

  const result = loginUser(username, password);

  Logger.log("=== ผลการทดสอบ Login ===");
  Logger.log(JSON.stringify(result, null, 2));

  if (result.success) {
    Logger.log("✅ Login สำเร็จ!");
  } else {
    Logger.log("❌ Login ไม่สำเร็จ: " + result.message);
  }

  return result;
}

/**
 * ตรวจสอบข้อมูล User ทั้งหมดในระบบ
 * รันฟังก์ชันนี้เพื่อดูข้อมูล User ทั้งหมดที่มีใน Sheet
 */
function checkAllUsers() {
  try {
    Logger.log("=== ตรวจสอบข้อมูล User ทั้งหมด ===");

    const sheet = getSheet(CONFIG.SHEETS.USERS);
    const data = sheet.getDataRange().getValues();

    Logger.log("จำนวนแถวทั้งหมด: " + data.length);
    Logger.log("");

    // แสดง Header
    Logger.log("Header (แถวที่ 1):");
    Logger.log(JSON.stringify(data[0]));
    Logger.log("");

    // แสดงข้อมูล User
    for (let i = 1; i < data.length; i++) {
      Logger.log(`แถวที่ ${i + 1}:`);
      Logger.log(`  A (รหัสผู้ใช้): ${data[i][0]}`);
      Logger.log(`  B (อีเมล): ${data[i][1]}`);
      Logger.log(`  C (ชื่อผู้ใช้): ${data[i][2]}`);
      Logger.log(`  D (รหัสผ่าน): ${data[i][3]}`);
      Logger.log(`  E (ชื่อ-นามสกุล): ${data[i][4]}`);
      Logger.log(`  F (เบอร์โทร): ${data[i][5]}`);
      Logger.log(`  G (บทบาท): ${data[i][6]}`);
      Logger.log(`  H (สถานะ): ${data[i][7]}`);
      Logger.log(`  I (วันที่สร้าง): ${data[i][8]}`);
      Logger.log(`  J (วันที่แก้ไข): ${data[i][9]}`);
      Logger.log("");
    }

    return {
      success: true,
      totalRows: data.length,
      totalUsers: data.length - 1,
      data: data,
    };
  } catch (error) {
    Logger.log("❌ ERROR: " + error.message);
    return {
      success: false,
      message: error.message,
    };
  }
}

/**
 * ทดสอบ Session Management
 * รันฟังก์ชันนี้เพื่อทดสอบการทำงานของ Session
 */
function testSession() {
  Logger.log("=== ทดสอบ Session ===");

  // 1. ลบ Session เดิม (ถ้ามี)
  clearSession();
  Logger.log("1. ลบ Session เดิม");

  // 2. ตรวจสอบว่าไม่มี Session
  let session = getSession();
  Logger.log("2. Session หลังลบ: " + JSON.stringify(session));
  Logger.log("   Is Logged In: " + isLoggedIn());

  // 3. ตั้งค่า Session ใหม่
  setSession("USR001", "admin", "Owner");
  Logger.log("3. ตั้งค่า Session ใหม่");

  // 4. อ่าน Session
  session = getSession();
  Logger.log("4. Session หลังตั้งค่า: " + JSON.stringify(session));
  Logger.log("   Is Logged In: " + isLoggedIn());

  // 5. ทดสอบ hasRole
  Logger.log("5. ทดสอบ hasRole:");
  Logger.log("   hasRole('Owner'): " + hasRole(CONFIG.ROLES.OWNER));
  Logger.log("   hasRole('Sales'): " + hasRole(CONFIG.ROLES.SALES));
  Logger.log(
    "   hasRole(['OP', 'Owner']): " +
      hasRole([CONFIG.ROLES.OP, CONFIG.ROLES.OWNER])
  );

  // 6. ลบ Session
  clearSession();
  Logger.log("6. ลบ Session");

  session = getSession();
  Logger.log("7. Session หลังลบ: " + JSON.stringify(session));
  Logger.log("   Is Logged In: " + isLoggedIn());

  return {
    success: true,
    message: "ทดสอบ Session เสร็จสิ้น",
  };
}

/**
 * ตรวจสอบ Password Hash
 * รันฟังก์ชันนี้เพื่อดู Hash ของรหัสผ่านที่ถูกต้อง
 */
function verifyPasswordHash() {
  Logger.log("=== ตรวจสอบ Password Hash ===");

  const password = "password123";
  const correctHash =
    "5e884898da28047151d0e56f8dc6292773603d0d6aabbdd62a11ef721d1542d8";

  const generatedHash = hashPassword(password);

  Logger.log("Password: " + password);
  Logger.log("Hash ที่ถูกต้อง: " + correctHash);
  Logger.log("Hash ที่สร้างขึ้น: " + generatedHash);
  Logger.log("ตรงกันหรือไม่: " + (correctHash === generatedHash));

  // ทดสอบกับข้อมูลใน Sheet
  try {
    const sheet = getSheet(CONFIG.SHEETS.USERS);
    const data = sheet.getDataRange().getValues();

    if (data.length > 1) {
      const dbHash = data[1][3]; // Password hash ของ User แรก
      Logger.log("");
      Logger.log("Hash ใน Database: " + dbHash);
      Logger.log("ตรงกับ Hash ที่ถูกต้องหรือไม่: " + (dbHash === correctHash));
      Logger.log(
        "ตรงกับ Hash ที่สร้างขึ้นหรือไม่: " + (dbHash === generatedHash)
      );
    }
  } catch (error) {
    Logger.log("ไม่สามารถอ่านข้อมูลจาก Sheet ได้: " + error.message);
  }

  return {
    success: true,
    password: password,
    correctHash: correctHash,
    generatedHash: generatedHash,
    isMatch: correctHash === generatedHash,
  };
}

/**
 * Create New User
 */
function createUser(userData) {
  if (!hasRole(CONFIG.ROLES.OWNER)) {
    return { success: false, message: "ไม่มีสิทธิ์เข้าถึง" };
  }

  try {
    const sheet = getSheet(CONFIG.SHEETS.USERS);

    // Check if username or email already exists
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (data[i][1] === userData.email || data[i][2] === userData.username) {
        return { success: false, message: "อีเมลหรือชื่อผู้ใช้นี้มีอยู่แล้ว" };
      }
    }

    const userId = generateUniqueId("USR");
    const defaultPassword = userData.phone || CONFIG.DEFAULT_PASSWORD;
    const hashedPassword = hashPassword(defaultPassword);
    const timestamp = getCurrentTimestamp();

    const newRow = [
      userId, // A: รหัสผู้ใช้
      userData.email, // B: อีเมล
      userData.username, // C: ชื่อผู้ใช้
      hashedPassword, // D: รหัสผ่าน
      userData.fullName, // E: ชื่อ-นามสกุล
      userData.phone ? "'" + userData.phone : "", // F: เบอร์โทร (Force text to keep leading zero)
      userData.role, // G: บทบาท
      "เปิดใช้งาน", // H: สถานะการใช้งาน
      timestamp, // I: วันที่สร้าง
      timestamp, // J: วันที่แก้ไขล่าสุด
    ];

    sheet.appendRow(newRow);

    return {
      success: true,
      message: "เพิ่มพนักงานสำเร็จ",
      data: { userId: userId },
    };
  } catch (error) {
    return { success: false, message: error.message };
  }
}

/**
 * Update User
 */
function updateUser(userId, userData) {
  if (!hasRole(CONFIG.ROLES.OWNER)) {
    return { success: false, message: "ไม่มีสิทธิ์เข้าถึง" };
  }

  try {
    const sheet = getSheet(CONFIG.SHEETS.USERS);
    const data = sheet.getDataRange().getValues();

    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === userId) {
        const timestamp = getCurrentTimestamp();

        // Update only provided fields
        if (userData.email) sheet.getRange(i + 1, 2).setValue(userData.email);
        if (userData.username)
          sheet.getRange(i + 1, 3).setValue(userData.username);
        if (userData.fullName)
          sheet.getRange(i + 1, 5).setValue(userData.fullName);
        if (userData.phone)
          sheet.getRange(i + 1, 6).setValue("'" + userData.phone);
        if (userData.role) sheet.getRange(i + 1, 7).setValue(userData.role);
        if (userData.status) sheet.getRange(i + 1, 8).setValue(userData.status);

        // Update timestamp
        sheet.getRange(i + 1, 10).setValue(timestamp);

        return { success: true, message: "อัพเดทข้อมูลสำเร็จ" };
      }
    }

    return { success: false, message: "ไม่พบผู้ใช้" };
  } catch (error) {
    return { success: false, message: error.message };
  }
}

/**
 * Reset User Password
 */
function resetUserPassword(userId) {
  if (!hasRole(CONFIG.ROLES.OWNER)) {
    return { success: false, message: "ไม่มีสิทธิ์เข้าถึง" };
  }

  try {
    const sheet = getSheet(CONFIG.SHEETS.USERS);
    const data = sheet.getDataRange().getValues();

    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === userId) {
        const phone = data[i][5]; // Column F: เบอร์โทร
        const defaultPassword = phone || CONFIG.DEFAULT_PASSWORD;
        const hashedPassword = hashPassword(defaultPassword);
        const timestamp = getCurrentTimestamp();

        sheet.getRange(i + 1, 4).setValue(hashedPassword);
        sheet.getRange(i + 1, 10).setValue(timestamp);

        return {
          success: true,
          message:
            "รีเซ็ตรหัสผ่านสำเร็จ รหัสผ่านใหม่: " +
            (phone || CONFIG.DEFAULT_PASSWORD),
        };
      }
    }

    return { success: false, message: "ไม่พบผู้ใช้" };
  } catch (error) {
    return { success: false, message: error.message };
  }
}

/**
 * Delete User (Soft Delete)
 */
function deleteUser(userId) {
  if (!hasRole(CONFIG.ROLES.OWNER)) {
    return { success: false, message: "ไม่มีสิทธิ์เข้าถึง" };
  }

  try {
    const sheet = getSheet(CONFIG.SHEETS.USERS);
    const data = sheet.getDataRange().getValues();

    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === userId) {
        // Check if user is Owner
        if (data[i][6] === CONFIG.ROLES.OWNER) {
          return { success: false, message: "ไม่สามารถลบ Owner ได้" };
        }

        const timestamp = getCurrentTimestamp();
        sheet.getRange(i + 1, 8).setValue("ปิดใช้งาน");
        sheet.getRange(i + 1, 10).setValue(timestamp);

        return { success: true, message: "ลบพนักงานสำเร็จ" };
      }
    }

    return { success: false, message: "ไม่พบผู้ใช้" };
  } catch (error) {
    return { success: false, message: error.message };
  }
}

// ========================================
// BOOKING MANAGEMENT
// ========================================

/**
 * Upload Slip to Google Drive
 */

// ========================================
// LOCATIONS MANAGEMENT
// ========================================

/**
 * Get All Locations
 */
function getAllLocations(sessionToken) {
  try {
    // Validate session
    const session = validateSession(sessionToken);
    if (!session) {
      return {
        success: false,
        message: "Session หมดอายุ กรุณาเข้าสู่ระบบใหม่",
      };
    }

    const sheet = getSheet(CONFIG.SHEETS.LOCATIONS);
    const data = sheet.getDataRange().getValues();

    const locations = data
      .slice(1)
      .filter((row) => row[0]) // Ensure locationId exists
      .map((row) => ({
        locationId: String(row[0] || ""),
        locationName: String(row[1] || ""),
        cellName: String(row[2] || ""),
        isActive:
          row[3] === "เปิดใช้งาน" || row[3] === true || row[3] === "Active",
        createdAt:
          row[4] instanceof Date ? row[4].toISOString() : String(row[4] || ""),
        updatedAt:
          row[5] instanceof Date ? row[5].toISOString() : String(row[5] || ""),
      }));

    return { success: true, data: locations };
  } catch (error) {
    return { success: false, message: error.message };
  }
}

/**
 * Create Location (Admin/Owner)
 */
function createLocation(sessionToken, locationData) {
  try {
    const session = validateSession(sessionToken);
    if (
      !session ||
      !hasRoleWithToken(sessionToken, [CONFIG.ROLES.ADMIN, CONFIG.ROLES.OWNER])
    ) {
      return {
        success: false,
        message: "ไม่มีสิทธิ์เข้าถึง (Admin/Owner Only)",
      };
    }

    const sheet = getSheet(CONFIG.SHEETS.LOCATIONS);
    const locationId = generateUniqueId("LOC");
    const timestamp = getCurrentTimestamp();

    const newRow = [
      locationId,
      locationData.locationName,
      locationData.cellName || "",
      "เปิดใช้งาน", // Default status
      timestamp,
      timestamp,
    ];

    sheet.appendRow(newRow);

    return {
      success: true,
      message: "เพิ่มสถานที่สำเร็จ",
      data: { locationId: locationId },
    };
  } catch (error) {
    return { success: false, message: error.message };
  }
}

/**
 * Update Location (Admin/Owner)
 */
function updateLocation(sessionToken, locationId, locationData) {
  try {
    const session = validateSession(sessionToken);
    if (
      !session ||
      !hasRoleWithToken(sessionToken, [CONFIG.ROLES.ADMIN, CONFIG.ROLES.OWNER])
    ) {
      return {
        success: false,
        message: "ไม่มีสิทธิ์เข้าถึง (Admin/Owner Only)",
      };
    }

    const sheet = getSheet(CONFIG.SHEETS.LOCATIONS);
    const data = sheet.getDataRange().getValues();
    let rowIndex = -1;

    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]).trim() === String(locationId).trim()) {
        rowIndex = i + 1;
        break;
      }
    }

    if (rowIndex === -1) {
      return { success: false, message: "ไม่พบข้อมูลสถานที่" };
    }

    const timestamp = getCurrentTimestamp();
    sheet.getRange(rowIndex, 2).setValue(locationData.locationName);
    sheet.getRange(rowIndex, 3).setValue(locationData.cellName || "");
    sheet.getRange(rowIndex, 4).setValue(locationData.status || "เปิดใช้งาน");
    sheet.getRange(rowIndex, 6).setValue(timestamp);

    return { success: true, message: "แก้ไขสถานที่สำเร็จ" };
  } catch (error) {
    return { success: false, message: error.message };
  }
}

/**
 * Toggle Location Status (Admin/Owner)
 */
function toggleLocationStatus(sessionToken, locationId) {
  try {
    const session = validateSession(sessionToken);
    if (
      !session ||
      !hasRoleWithToken(sessionToken, [CONFIG.ROLES.ADMIN, CONFIG.ROLES.OWNER])
    ) {
      return {
        success: false,
        message: "ไม่มีสิทธิ์เข้าถึง (Admin/Owner Only)",
      };
    }

    const sheet = getSheet(CONFIG.SHEETS.LOCATIONS);
    const data = sheet.getDataRange().getValues();
    let rowIndex = -1;
    let currentIsActive = false;

    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]).trim() === String(locationId).trim()) {
        rowIndex = i + 1;
        currentIsActive =
          data[i][3] === "เปิดใช้งาน" ||
          data[i][3] === true ||
          data[i][3] === "Active";
        break;
      }
    }

    if (rowIndex === -1) {
      return { success: false, message: "ไม่พบข้อมูลสถานที่" };
    }

    const newStatus = currentIsActive ? "ปิดใช้งาน" : "เปิดใช้งาน";
    const timestamp = getCurrentTimestamp();

    sheet.getRange(rowIndex, 4).setValue(newStatus);
    sheet.getRange(rowIndex, 6).setValue(timestamp);

    return {
      success: true,
      message: `เปลี่ยนสถานะเป็น ${newStatus} สำเร็จ`,
    };
  } catch (error) {
    return { success: false, message: error.message };
  }
}

/**
 * Delete Location (Admin/Owner) - Actually deactivates the location
 */
function deleteLocation(sessionToken, locationId) {
  try {
    const session = validateSession(sessionToken);
    if (
      !session ||
      !hasRoleWithToken(sessionToken, [CONFIG.ROLES.ADMIN, CONFIG.ROLES.OWNER])
    ) {
      return {
        success: false,
        message: "ไม่มีสิทธิ์เข้าถึง (Admin/Owner Only)",
      };
    }

    const sheet = getSheet(CONFIG.SHEETS.LOCATIONS);
    const data = sheet.getDataRange().getValues();
    let rowIndex = -1;

    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]).trim() === String(locationId).trim()) {
        rowIndex = i + 1;
        break;
      }
    }

    if (rowIndex === -1) {
      return { success: false, message: "ไม่พบข้อมูลสถานที่" };
    }

    // Instead of deleting, update status to "ปิดใช้งาน"
    const timestamp = getCurrentTimestamp();
    sheet.getRange(rowIndex, 4).setValue("ปิดใช้งาน");
    sheet.getRange(rowIndex, 6).setValue(timestamp);

    return { success: true, message: "ปิดการใช้งานสถานที่สำเร็จ" };
  } catch (error) {
    return { success: false, message: error.message };
  }
}

// ========================================
// PROGRAMS MANAGEMENT
// ========================================

/**
 * Get All Programs
 */
function getAllPrograms(sessionToken) {
  try {
    // Validate session
    const session = validateSession(sessionToken);
    if (!session) {
      return {
        success: false,
        message: "Session หมดอายุ กรุณาเข้าสู่ระบบใหม่",
      };
    }

    const sheet = getSheet(CONFIG.SHEETS.PROGRAMS);
    if (!sheet) {
      return {
        success: false,
        message: "ไม่พบ Sheet Programs กรุณารัน setupAllSheets หนึ่งครั้ง",
      };
    }
    const data = sheet.getDataRange().getValues();
    const programs = [];

    // Header: รหัสโปรแกรม0, รายละเอียด1, ราคาผู้ใหญ่2, ราคาเด็ก3, สถานะการใช้งาน4, วันที่สร้าง5, วันที่แก้ไขล่าสุด6
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      if (!row[0]) continue;

      programs.push({
        programId: String(row[0] || ""),
        description: String(row[1] || ""),
        adultPrice: parseFloat(row[2]) || 0,
        childPrice: parseFloat(row[3]) || 0,
        isActive:
          row[4] === "เปิดใช้งาน" || row[4] === true || row[4] === "Active",
        createdAt: formatDate(row[5], "dd/MM/yyyy HH:mm:ss"),
        updatedAt: formatDate(row[6], "dd/MM/yyyy HH:mm:ss"),
      });
    }

    return { success: true, data: programs };
  } catch (error) {
    return { success: false, message: error.message };
  }
}

/**
 * Create Program (Owner Only)
 */
function createProgram(sessionToken, programData) {
  try {
    const session = validateSession(sessionToken);
    if (!session || session.role !== CONFIG.ROLES.OWNER) {
      return { success: false, message: "ไม่มีสิทธิ์เข้าถึง (Owner Only)" };
    }

    const sheet = getSheet(CONFIG.SHEETS.PROGRAMS);
    if (!sheet) {
      return { success: false, message: "ไม่พบ Sheet Programs" };
    }

    const data = sheet.getDataRange().getValues();
    const programId = String(programData.programId).trim();

    // Check for duplicate ID
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]) === programId) {
        return { success: false, message: "รหัสโปรแกรมนี้มีอยู่แล้วในระบบ" };
      }
    }

    const timestamp = getCurrentTimestamp();
    const newRow = [
      programId,
      programData.description || "",
      programData.adultPrice,
      programData.childPrice,
      programData.isActive ? "เปิดใช้งาน" : "ปิดใช้งาน",
      timestamp,
      timestamp,
    ];

    sheet.appendRow(newRow);

    return {
      success: true,
      message: "เพิ่มโปรแกรมสำเร็จ",
      data: { programId: programId },
    };
  } catch (error) {
    return { success: false, message: error.message };
  }
}

/**
 * Update Program (Owner Only)
 */
function updateProgram(sessionToken, oldId, programData) {
  try {
    const session = validateSession(sessionToken);
    if (
      !session ||
      (session.role !== CONFIG.ROLES.OWNER &&
        session.role !== CONFIG.ROLES.AR_AP)
    ) {
      return {
        success: false,
        message: "ไม่มีสิทธิ์เข้าถึง (Owner or AR_AP Only)",
      };
    }
    const isOwner = session.role === CONFIG.ROLES.OWNER;
    const isArAp = session.role === CONFIG.ROLES.AR_AP;

    const sheet = getSheet(CONFIG.SHEETS.PROGRAMS);
    if (!sheet) {
      return { success: false, message: "ไม่พบ Sheet Programs" };
    }
    const data = sheet.getDataRange().getValues();
    const newId = String(programData.programId).trim();
    let rowIndex = -1;

    // Find current row and check for duplicate of NEW ID (if changed)
    for (let i = 1; i < data.length; i++) {
      const currentIdInSheet = String(data[i][0]).trim();

      // Find the record we want to update
      if (currentIdInSheet === oldId) {
        rowIndex = i + 1;
      }

      // Check if the NEW ID already exists (and is not the one we are currently editing)
      // Only check this if Owner (since AR_AP cannot change ID)
      if (isOwner && newId !== oldId && currentIdInSheet === newId) {
        return {
          success: false,
          message: "รหัสโปรแกรมใหม่ขัดแย้งกับโปรแกรมอื่นที่มีอยู่แล้ว",
        };
      }
    }

    if (rowIndex === -1) {
      return { success: false, message: "ไม่พบข้อมูลโปรแกรมเดิม" };
    }

    const timestamp = getCurrentTimestamp();

    if (isOwner) {
      // Owner updates everything
      sheet.getRange(rowIndex, 1).setValue(newId);
      sheet.getRange(rowIndex, 2).setValue(programData.description || "");
      sheet.getRange(rowIndex, 3).setValue(programData.adultPrice);
      sheet.getRange(rowIndex, 4).setValue(programData.childPrice);
      sheet
        .getRange(rowIndex, 5)
        .setValue(programData.isActive ? "เปิดใช้งาน" : "ปิดใช้งาน");
    } else if (isArAp) {
      // AR_AP updates ONLY prices and timestamp
      sheet.getRange(rowIndex, 3).setValue(programData.adultPrice);
      sheet.getRange(rowIndex, 4).setValue(programData.childPrice);
    }

    sheet.getRange(rowIndex, 7).setValue(timestamp);

    return { success: true, message: "แก้ไขโปรแกรมสำเร็จ" };
  } catch (error) {
    return { success: false, message: error.message };
  }
}

/**
 * Delete Program (Soft Delete - Owner Only)
 */
function deleteProgram(sessionToken, programId) {
  try {
    const session = validateSession(sessionToken);
    if (!session || session.role !== CONFIG.ROLES.OWNER) {
      return { success: false, message: "ไม่มีสิทธิ์เข้าถึง (Owner Only)" };
    }

    const sheet = getSheet(CONFIG.SHEETS.PROGRAMS);
    if (!sheet) {
      return { success: false, message: "ไม่พบ Sheet Programs" };
    }
    const data = sheet.getDataRange().getValues();
    let rowIndex = -1;

    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === programId) {
        rowIndex = i + 1;
        break;
      }
    }

    if (rowIndex === -1) {
      return { success: false, message: "ไม่พบข้อมูลโปรแกรม" };
    }

    const timestamp = getCurrentTimestamp();
    sheet.getRange(rowIndex, 5).setValue("ปิดใช้งาน");
    sheet.getRange(rowIndex, 7).setValue(timestamp);

    return {
      success: true,
      message: "ลบโปรแกรมสำเร็จ (เปลี่ยนสถานะเป็นปิดใช้งาน)",
    };
  } catch (error) {
    return { success: false, message: error.message };
  }
}

/**
 * Toggle Program Status (Owner Only)
 */
function toggleProgramStatus(sessionToken, programId) {
  try {
    const session = validateSession(sessionToken);
    if (!session || session.role !== CONFIG.ROLES.OWNER) {
      return { success: false, message: "ไม่มีสิทธิ์เข้าถึง (Owner Only)" };
    }

    const sheet = getSheet(CONFIG.SHEETS.PROGRAMS);
    if (!sheet) return { success: false, message: "ไม่พบ Sheet Programs" };

    const data = sheet.getDataRange().getValues();
    let rowIndex = -1;
    let currentIsActive = false;

    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]).trim() === String(programId).trim()) {
        rowIndex = i + 1;
        currentIsActive = data[i][4] === "เปิดใช้งาน";
        break;
      }
    }

    if (rowIndex === -1)
      return { success: false, message: "ไม่พบข้อมูลโปรแกรม" };

    const newStatus = currentIsActive ? "ปิดใช้งาน" : "เปิดใช้งาน";
    const timestamp = getCurrentTimestamp();

    sheet.getRange(rowIndex, 5).setValue(newStatus);
    sheet.getRange(rowIndex, 7).setValue(timestamp);

    return {
      success: true,
      message: `เปลี่ยนสถานะเป็น ${newStatus} สำเร็จ`,
      data: { isActive: !currentIsActive },
    };
  } catch (error) {
    return { success: false, message: error.message };
  }
}

// ========================================
// DASHBOARD & REPORTS
// ========================================

/**
 * Get Dashboard Data
 */
function getDashboardData(sessionToken, startDate, endDate) {
  try {
    // Validate session
    const session = validateSession(sessionToken);
    if (!session) {
      Logger.log("Session validation failed");
      return {
        success: false,
        message: "Session หมดอายุ กรุณาเข้าสู่ระบบใหม่",
      };
    }

    // Check role permissions (Owner, Cost, and Admin can access)
    const allowedRoles = [
      CONFIG.ROLES.OWNER,
      CONFIG.ROLES.COST,
      CONFIG.ROLES.ADMIN,
      CONFIG.ROLES.OP,
    ];
    if (!allowedRoles.includes(session.role)) {
      return { success: false, message: "ไม่มีสิทธิ์เข้าถึง Dashboard" };
    }

    // Parse date range
    // Helper to parse YYYY-MM-DD to Local Script Timezone (not UTC)
    const parseLocalYMD = (dateStr) => {
      if (typeof dateStr === "string" && /^\d{4}-\d{2}-\d{2}$/.test(dateStr)) {
        const [y, m, d] = dateStr.split("-").map(Number);
        return new Date(y, m - 1, d);
      }
      return dateStr ? new Date(dateStr) : null;
    };

    const filterStart = parseLocalYMD(startDate);
    const filterEnd = parseLocalYMD(endDate);

    if (filterStart) {
      filterStart.setHours(0, 0, 0, 0);
    }
    if (filterEnd) {
      filterEnd.setHours(23, 59, 59, 999);
    }

    // Calculate previous period for comparison
    let prevFilterStart = null;
    let prevFilterEnd = null;

    if (filterStart && filterEnd) {
      const daysDiff =
        Math.ceil((filterEnd - filterStart) / (1000 * 60 * 60 * 24)) + 1;
      prevFilterEnd = new Date(filterStart);
      prevFilterEnd.setDate(prevFilterEnd.getDate() - 1);
      prevFilterEnd.setHours(23, 59, 59, 999);

      prevFilterStart = new Date(prevFilterEnd);
      prevFilterStart.setDate(prevFilterStart.getDate() - daysDiff + 1);
      prevFilterStart.setHours(0, 0, 0, 0);
    }

    const sheet = getSheet(CONFIG.SHEETS.BOOKING_RAW);
    const data = sheet.getDataRange().getValues();

    // Initialize counters for current period
    let totalSales = 0;
    let completedBookings = 0;
    let confirmedBookings = 0;
    let pendingAmount = 0;
    let totalBookings = 0;
    let cancelledBookings = 0;

    // Initialize counters for previous period
    let prevTotalSales = 0;
    let prevCompletedBookings = 0;
    let prevTotalBookings = 0;

    const programStats = {};
    const agentStats = {};
    const locationStats = {};

    // Process filtered bookings
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const bookingDate = new Date(row[1]); // Column B: Booking Date
      const status = row[12]; // Column M: Status
      const totalAmount = Number(row[19]) || 0; // Column T: Total Amount
      const program = row[4]; // Column E: Program
      const agent = row[17]; // Column R: Agent

      // Check if in current period
      const isInCurrentPeriod =
        (!filterStart || bookingDate >= filterStart) &&
        (!filterEnd || bookingDate <= filterEnd);

      // Check if in previous period
      const isInPreviousPeriod =
        prevFilterStart &&
        prevFilterEnd &&
        bookingDate >= prevFilterStart &&
        bookingDate <= prevFilterEnd;

      // Process current period data
      if (isInCurrentPeriod) {
        totalBookings++;

        if (status === CONFIG.STATUS.CANCEL) {
          cancelledBookings++;
        }

        if (status === CONFIG.STATUS.COMPLETE) {
          completedBookings++;
          totalSales += totalAmount;
        }

        if (status === CONFIG.STATUS.CONFIRM) {
          confirmedBookings++;
          // คำนวณยอดค้างรับ = ยอดรวม - (เงินโอน + เงินสด + Cash on tour)
          const transferAmount = Number(row[13]) || 0; // Column N: เงินโอน
          const cashAmount = Number(row[14]) || 0; // Column O: เงินสด
          const cashOnTour = Number(row[15]) || 0; // Column P: Cash on tour
          const totalPaid = transferAmount + cashAmount + cashOnTour;
          const remaining = totalAmount - totalPaid;
          pendingAmount += remaining > 0 ? remaining : 0; // เอาเฉพาะยอดที่ยังค้างรับ (ไม่ติดลบ)
        }

        // Program stats
        if (status === CONFIG.STATUS.COMPLETE) {
          if (!programStats[program]) {
            programStats[program] = { count: 0, amount: 0 };
          }
          programStats[program].count++;
          programStats[program].amount += totalAmount;
        }

        // Agent stats
        if (status === CONFIG.STATUS.COMPLETE && agent) {
          if (!agentStats[agent]) {
            agentStats[agent] = 0;
          }
          agentStats[agent] += totalAmount;
        }

        // Location stats
        if (status === CONFIG.STATUS.COMPLETE && row[3]) {
          const locationName = row[3];
          if (!locationStats[locationName]) {
            locationStats[locationName] = 0;
          }
          locationStats[locationName] += totalAmount;
        }
      }

      // Process previous period data (for comparison)
      if (isInPreviousPeriod) {
        prevTotalBookings++;

        if (status === CONFIG.STATUS.COMPLETE) {
          prevCompletedBookings++;
          prevTotalSales += totalAmount;
        }
      }
    }

    // Calculate cancel rate
    const cancelRate =
      totalBookings > 0
        ? ((cancelledBookings / totalBookings) * 100).toFixed(2)
        : 0;

    // Calculate growth rates (% change from previous period)
    const calculateGrowth = (current, previous) => {
      if (!previous || previous === 0) return null;
      return (((current - previous) / previous) * 100).toFixed(2);
    };

    // Process Refund Data
    let totalRefundAmount = 0;
    let refundCount = 0;
    let prevTotalRefundAmount = 0;
    let prevRefundCount = 0;

    try {
      const refundSheet = getSheet(CONFIG.SHEETS.REFUND);
      if (refundSheet) {
        const refundData = refundSheet.getDataRange().getValues();
        for (let j = 1; j < refundData.length; j++) {
          let rDateValue = refundData[j][9]; // Column J: Created At
          let rDate;

          if (rDateValue instanceof Date) {
            rDate = rDateValue;
          } else if (
            typeof rDateValue === "string" &&
            rDateValue.includes("/")
          ) {
            // Parse dd/MM/yyyy
            const parts = rDateValue.split(" ");
            const dateParts = parts[0].split("/");
            if (dateParts.length === 3) {
              const day = parseInt(dateParts[0], 10);
              const month = parseInt(dateParts[1], 10) - 1;
              const year = parseInt(dateParts[2], 10);
              rDate = new Date(year, month, day);
            }
          }

          if (!rDate || isNaN(rDate.getTime())) continue;

          const rAmount = Number(refundData[j][5]) || 0; // Column F: Refund Amount

          const isInCurrent =
            (!filterStart || rDate >= filterStart) &&
            (!filterEnd || rDate <= filterEnd);
          const isInPrev =
            prevFilterStart &&
            prevFilterEnd &&
            rDate >= prevFilterStart &&
            rDate <= prevFilterEnd;

          if (isInCurrent) {
            totalRefundAmount += rAmount;
            refundCount++;
          }
          if (isInPrev) {
            prevTotalRefundAmount += rAmount;
            prevRefundCount++;
          }
        }
      }
    } catch (e) {
      Logger.log("Dashboard Refund Process Error: " + e.message);
    }

    // Get top 5 programs
    const topPrograms = Object.entries(programStats)
      .sort((a, b) => b[1].amount - a[1].amount)
      .slice(0, 5)
      .map(([name, stats]) => ({ name, ...stats }));

    // Growth rates
    const salesGrowth = calculateGrowth(totalSales, prevTotalSales);
    const bookingsGrowth = calculateGrowth(totalBookings, prevTotalBookings);
    const refundGrowth = calculateGrowth(
      totalRefundAmount,
      prevTotalRefundAmount
    );

    return {
      success: true,
      data: {
        totalSales: totalSales,
        completedBookings: completedBookings,
        confirmedBookings: confirmedBookings,
        pendingAmount: pendingAmount,
        totalBookings: totalBookings,
        cancelledBookings: cancelledBookings,
        cancelRate: cancelRate,
        // Refund Data
        totalRefundAmount: totalRefundAmount,
        refundCount: refundCount,
        refundGrowth: refundGrowth,
        // Comparison data
        salesGrowth: salesGrowth,
        bookingsGrowth: bookingsGrowth,
        prevTotalSales: prevTotalSales,
        prevTotalBookings: prevTotalBookings,
        topPrograms: topPrograms,
        salesByAgent: agentStats,
        salesByLocation: locationStats,
      },
    };
  } catch (error) {
    Logger.log("Get dashboard data error: " + error.message);
    return { success: false, message: error.message };
  }
}

// ========================================
// USER MANAGEMENT (Owner Only)
// ========================================

/**
 * Get All Users (Owner only)
 */
function getAllUsers(sessionToken) {
  try {
    const session = validateSession(sessionToken);

    if (!session) {
      return {
        success: false,
        message: "Session หมดอายุ กรุณาเข้าสู่ระบบใหม่",
      };
    }

    const sheet = getSheet(CONFIG.SHEETS.USERS);

    if (!sheet) {
      return {
        success: false,
        message: "ไม่พบ Sheet Users",
      };
    }

    const data = sheet.getDataRange().getValues();

    const users = [];
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      if (row[0]) {
        // Convert to plain strings for JSON serialization
        users.push({
          userId: String(row[0] || ""),
          email: String(row[1] || ""),
          username: String(row[2] || ""),
          fullName: String(row[4] || ""),
          phone: String(row[5] || ""),
          role: String(row[6] || ""),
          status: String(row[7] || ""),
          createdAt: row[8] ? String(row[8]) : "",
          updatedAt: row[9] ? String(row[9]) : "",
        });
      }
    }

    return {
      success: true,
      data: users,
    };
  } catch (error) {
    return {
      success: false,
      message: "เกิดข้อผิดพลาด: " + error.message,
    };
  }
}

/**
 * Create New User (Owner only)
 */
function createUser(sessionToken, userData) {
  try {
    // Validate session
    const session = validateSession(sessionToken);
    if (!session || session.role !== CONFIG.ROLES.OWNER) {
      return { success: false, message: "คุณไม่มีสิทธิ์ดำเนินการนี้" };
    }

    // Validate required fields
    if (
      !userData.fullName ||
      !userData.email ||
      !userData.username ||
      !userData.role
    ) {
      return { success: false, message: "กรุณากรอกข้อมูลให้ครบถ้วน" };
    }

    // Validate email format
    const emailRegex = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
    if (!emailRegex.test(userData.email)) {
      return { success: false, message: "รูปแบบอีเมลไม่ถูกต้อง" };
    }

    // Validate role
    const validRoles = ["Sale", "OP", "Admin", "AR_AP", "Cost", "Owner"];
    if (!validRoles.includes(userData.role)) {
      return { success: false, message: "ตำแหน่งไม่ถูกต้อง" };
    }

    // Check duplicate email and username
    const sheet = getSheet(CONFIG.SHEETS.USERS);
    const data = sheet.getDataRange().getValues();

    for (let i = 1; i < data.length; i++) {
      if (data[i][1] === userData.email) {
        return { success: false, message: "อีเมลนี้ถูกใช้งานแล้ว" };
      }
      if (data[i][2] === userData.username) {
        return { success: false, message: "ชื่อผู้ใช้นี้ถูกใช้งานแล้ว" };
      }
    }

    // Generate user ID
    const userId = generateUserId();

    // Default password (Updated: use phone number if available)
    const defaultPassword = userData.phone || "password123";
    const hashedPassword = hashPassword(defaultPassword);

    // Prepare data
    const now = getCurrentTimestamp();
    const newRow = [
      userId,
      userData.email,
      userData.username,
      hashedPassword,
      userData.fullName,
      userData.phone ? "'" + userData.phone : "", // เบอร์โทร (Force text to keep leading zero)
      userData.role,
      "เปิดใช้งาน",
      now,
      now,
    ];

    // Append to sheet
    sheet.appendRow(newRow);

    return {
      success: true,
      message: "เพิ่มพนักงานสำเร็จ",
      data: { userId: userId },
    };
  } catch (error) {
    return {
      success: false,
      message: "เกิดข้อผิดพลาด: " + error.message,
    };
  }
}

/**
 * Update User (Owner only)
 */
function updateUser(sessionToken, userId, userData) {
  try {
    // Validate session
    const session = validateSession(sessionToken);
    if (!session || session.role !== CONFIG.ROLES.OWNER) {
      return { success: false, message: "คุณไม่มีสิทธิ์ดำเนินการนี้" };
    }

    // Validate required fields
    if (
      !userData.fullName ||
      !userData.email ||
      !userData.username ||
      !userData.role
    ) {
      return { success: false, message: "กรุณากรอกข้อมูลให้ครบถ้วน" };
    }

    // Validate email format
    const emailRegex = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
    if (!emailRegex.test(userData.email)) {
      return { success: false, message: "รูปแบบอีเมลไม่ถูกต้อง" };
    }

    const sheet = getSheet(CONFIG.SHEETS.USERS);
    const data = sheet.getDataRange().getValues();

    // Find user
    let rowIndex = -1;
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === userId) {
        rowIndex = i + 1; // Sheet row is 1-indexed
        break;
      }
    }

    if (rowIndex === -1) {
      return { success: false, message: "ไม่พบข้อมูลพนักงาน" };
    }

    // Check duplicate email and username (exclude current user)
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] !== userId) {
        if (data[i][1] === userData.email) {
          return { success: false, message: "อีเมลนี้ถูกใช้งานแล้ว" };
        }
        if (data[i][2] === userData.username) {
          return { success: false, message: "ชื่อผู้ใช้นี้ถูกใช้งานแล้ว" };
        }
      }
    }

    // Update data
    sheet.getRange(rowIndex, 2).setValue(userData.email);
    sheet.getRange(rowIndex, 3).setValue(userData.username);
    sheet.getRange(rowIndex, 5).setValue(userData.fullName);
    sheet
      .getRange(rowIndex, 6)
      .setValue(userData.phone ? "'" + userData.phone : "");
    sheet.getRange(rowIndex, 7).setValue(userData.role);
    sheet.getRange(rowIndex, 8).setValue(userData.status);
    sheet.getRange(rowIndex, 10).setValue(getCurrentTimestamp());

    return {
      success: true,
      message: "แก้ไขข้อมูลสำเร็จ",
    };
  } catch (error) {
    return {
      success: false,
      message: "เกิดข้อผิดพลาด: " + error.message,
    };
  }
}

/**
 * Delete User (Soft delete - Owner only)
 */
function deleteUser(sessionToken, userId) {
  try {
    // Validate session
    const session = validateSession(sessionToken);
    if (!session || session.role !== CONFIG.ROLES.OWNER) {
      return { success: false, message: "คุณไม่มีสิทธิ์ดำเนินการนี้" };
    }

    const sheet = getSheet(CONFIG.SHEETS.USERS);
    const data = sheet.getDataRange().getValues();

    // Find user
    let rowIndex = -1;
    let userRole = null;
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === userId) {
        rowIndex = i + 1;
        userRole = data[i][6]; // Column G: บทบาท
        break;
      }
    }

    if (rowIndex === -1) {
      return { success: false, message: "ไม่พบข้อมูลพนักงาน" };
    }

    // Prevent deleting Owner
    if (userRole === CONFIG.ROLES.OWNER) {
      return { success: false, message: "ไม่สามารถลบ Owner ได้" };
    }

    // Soft delete - change status to Inactive
    sheet.getRange(rowIndex, 8).setValue("ปิดใช้งาน");
    sheet.getRange(rowIndex, 10).setValue(getCurrentTimestamp());

    return {
      success: true,
      message: "ลบพนักงานสำเร็จ",
    };
  } catch (error) {
    return {
      success: false,
      message: "เกิดข้อผิดพลาด: " + error.message,
    };
  }
}

/**
 * Reset User Password (Owner only)
 */
function resetUserPassword(sessionToken, userId) {
  try {
    // Validate session
    const session = validateSession(sessionToken);
    if (!session || session.role !== CONFIG.ROLES.OWNER) {
      return { success: false, message: "คุณไม่มีสิทธิ์ดำเนินการนี้" };
    }

    const sheet = getSheet(CONFIG.SHEETS.USERS);
    const data = sheet.getDataRange().getValues();

    // Find user
    let rowIndex = -1;
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === userId) {
        rowIndex = i + 1;
        break;
      }
    }

    if (rowIndex === -1) {
      return { success: false, message: "ไม่พบข้อมูลพนักงาน" };
    }

    // Reset password to default (Updated: use phone number from Column F)
    const phone = data[rowIndex - 1][5]; // Column F
    const defaultPassword = phone || "password123";
    const hashedPassword = hashPassword(defaultPassword);

    sheet.getRange(rowIndex, 4).setValue(hashedPassword);
    sheet.getRange(rowIndex, 10).setValue(getCurrentTimestamp());

    return {
      success: true,
      message: "รีเซ็ตรหัสผ่านสำเร็จ รหัสผ่านใหม่: " + (phone || "password123"),
    };
  } catch (error) {
    return {
      success: false,
      message: "เกิดข้อผิดพลาด: " + error.message,
    };
  }
}

/**
 * Toggle User Status (Owner only)
 */
function toggleUserStatus(sessionToken, userId) {
  if (!hasRoleWithToken(sessionToken, CONFIG.ROLES.OWNER)) {
    return { success: false, message: "ไม่มีสิทธิ์เข้าถึง" };
  }

  try {
    const sheet = getSheet(CONFIG.SHEETS.USERS);
    if (!sheet) return { success: false, message: "ไม่พบ Sheet Users" };

    const data = sheet.getDataRange().getValues();
    let rowIndex = -1;
    let currentIsActive = false;

    // Skip header (row 1)
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]).trim() === String(userId).trim()) {
        rowIndex = i + 1;
        currentIsActive = data[i][7] === "เปิดใช้งาน"; // Column H: สถานะ
        break;
      }
    }

    if (rowIndex === -1) {
      return { success: false, message: "ไม่พบข้อมูลผู้ใช้" };
    }

    const newStatus = currentIsActive ? "ปิดใช้งาน" : "เปิดใช้งาน";
    const timestamp = getCurrentTimestamp();

    // Column 8 is Status, Column 10 is Updated At
    sheet.getRange(rowIndex, 8).setValue(newStatus);
    sheet.getRange(rowIndex, 10).setValue(timestamp);

    return {
      success: true,
      message: `เปลี่ยนสถานะเป็น ${newStatus} สำเร็จ`,
      data: { isActive: !currentIsActive },
    };
  } catch (error) {
    return { success: false, message: error.message };
  }
}

/**
 * Generate User ID
 */
function generateUserId() {
  const sheet = getSheet(CONFIG.SHEETS.USERS);
  const lastRow = sheet.getLastRow();

  if (lastRow === 1) {
    return "USR001";
  }

  const lastId = sheet.getRange(lastRow, 1).getValue();
  const number = parseInt(lastId.replace("USR", "")) + 1;
  return "USR" + number.toString().padStart(3, "0");
}

// ========================================
// BOOKING MANAGEMENT FUNCTIONS
// ========================================

/**
 * Generate Booking ID
 */
function generateBookingId() {
  const sheet = getSheet(CONFIG.SHEETS.BOOKING_RAW);
  const lastRow = sheet.getLastRow();

  // If only header exists, start with BK001
  if (lastRow <= 1) {
    return "BK001";
  }

  // Get all booking IDs and find the highest number
  const data = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
  let maxNumber = 0;

  for (let i = 0; i < data.length; i++) {
    const id = data[i][0];
    if (id && typeof id === "string" && id.startsWith("BK")) {
      const numStr = id.replace("BK", "");
      const num = parseInt(numStr);
      if (!isNaN(num) && num > maxNumber) {
        maxNumber = num;
      }
    }
  }

  // Generate next ID
  const nextNumber = maxNumber + 1;
  return "BK" + nextNumber.toString().padStart(3, "0");
}

/**
 * Get All Bookings
 * ดึงข้อมูลการจองทั้งหมด
 * OP: ดูได้ทั้งหมด แต่แก้ไขได้เฉพาะที่ตนเองสร้าง
 * Owner: ดูและแก้ไขได้ทั้งหมด
 */
function getAllBookings(sessionToken, startDate, endDate) {
  try {
    // Validate session
    const session = validateSession(sessionToken);
    if (!session) {
      return {
        success: false,
        message: "Session หมดอายุ กรุณาเข้าสู่ระบบใหม่",
      };
    }

    // Check permissions
    const allowedRoles = [
      CONFIG.ROLES.OP,
      CONFIG.ROLES.OWNER,
      CONFIG.ROLES.ADMIN,
      CONFIG.ROLES.AR_AP,
      CONFIG.ROLES.SALE,
      CONFIG.ROLES.COST,
    ];

    if (!allowedRoles.includes(session.role)) {
      return {
        success: false,
        message: "คุณไม่มีสิทธิ์เข้าถึงข้อมูลนี้",
      };
    }

    // Parse date range
    const filterStart = startDate ? new Date(startDate) : null;
    const filterEnd = endDate ? new Date(endDate) : null;
    if (filterStart) filterStart.setHours(0, 0, 0, 0);
    if (filterEnd) filterEnd.setHours(23, 59, 59, 999);

    const viewingAllRoles = [
      CONFIG.ROLES.OWNER,
      CONFIG.ROLES.ADMIN,
      CONFIG.ROLES.AR_AP,
      CONFIG.ROLES.COST,
      CONFIG.ROLES.OP,
    ];
    const canViewAll = viewingAllRoles.includes(session.role);
    const currentUser = session.username;

    const sheet = getSheet(CONFIG.SHEETS.BOOKING_RAW);
    const data = sheet.getDataRange().getValues();

    // Fetch History Reasons
    const historySheet = getSheet(CONFIG.SHEETS.BOOKING_STATUS_HISTORY);
    const historyData = historySheet
      ? historySheet.getDataRange().getValues()
      : [];
    const historyMap = {};

    // Map latest reason to booking ID
    for (let i = 1; i < historyData.length; i++) {
      const hRow = historyData[i];
      const hBookingId = hRow[1]; // Column B: Booking ID
      const hReason = hRow[6]; // Column G: Reason

      if (hReason && String(hReason).trim() !== "") {
        historyMap[hBookingId] = hReason;
      }
    }

    const bookings = [];
    for (let i = 1; i < data.length; i++) {
      const row = data[i];

      // Date Filter
      const bookingDateRaw = new Date(row[1]); // Column B: Booking Date
      if (filterStart && bookingDateRaw < filterStart) continue;
      if (filterEnd && bookingDateRaw > filterEnd) continue;

      const createdBy = row[20]; // Column U: ผู้สร้าง

      // Filter data: If not in privileged roles, only show own bookings
      if (!canViewAll && createdBy !== currentUser) {
        continue;
      }

      bookings.push({
        bookingId: row[0],
        bookingDate: formatDate(row[1]),
        travelDate: formatDate(row[2]),
        location: row[3],
        program: row[4],
        adults: row[5],
        children: row[6],
        adultPrice: row[7],
        childPrice: row[8],
        additionalCost: row[9],
        discount: row[10],
        includeVat:
          row[11] === true || row[11] === "TRUE" || row[11] === "true",
        status: row[12],
        transferAmount: row[13],
        cashAmount: row[14],
        cashOnTour: row[15],
        slipUrl: row[16],
        agent: row[17],
        notes: row[18],
        totalAmount: row[19],
        createdBy: row[20],
        createdAt: formatDate(row[21], "dd/MM/yyyy HH:mm:ss"),
        updatedBy: row[22],
        updatedAt: formatDate(row[23], "dd/MM/yyyy HH:mm:ss"),
        reason: historyMap[row[0]] || "", // Add reason from history
      });
    }

    return {
      success: true,
      data: bookings,
    };
  } catch (error) {
    return {
      success: false,
      message: "เกิดข้อผิดพลาด: " + error.message,
    };
  }
}

/**
 * Get Booking by ID
 */
function getBookingById(sessionToken, bookingId) {
  try {
    // Validate session
    const session = validateSession(sessionToken);
    if (!session) {
      return {
        success: false,
        message: "Session หมดอายุ กรุณาเข้าสู่ระบบใหม่",
      };
    }

    const sheet = getSheet(CONFIG.SHEETS.BOOKING_RAW);
    const data = sheet.getDataRange().getValues();

    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      if (row[0] === bookingId) {
        return {
          success: true,
          data: {
            bookingId: row[0],
            bookingDate: formatDate(row[1]),
            travelDate: formatDate(row[2]),
            location: row[3],
            program: row[4],
            adults: row[5],
            children: row[6],
            adultPrice: row[7],
            childPrice: row[8],
            additionalCost: row[9], // Column J
            discount: row[10], // Column K
            includeVat:
              row[11] === true || row[11] === "TRUE" || row[11] === "true", // Column L: รวม VAT 7% (boolean)
            status: row[12], // Column M
            transferAmount: row[13], // Column N: เงินโอน
            cashAmount: row[14], // Column O: เงินสด
            cashOnTour: row[15], // Column P: Cash on tour
            slipUrl: row[16], // Column Q
            agent: row[17], // Column R
            notes: row[18], // Column S
            totalAmount: row[19], // Column T
            createdBy: row[20], // Column U
            createdAt: formatDate(row[21], "dd/MM/yyyy HH:mm:ss"), // Column V
            updatedBy: row[22], // Column W
            updatedAt: formatDate(row[23], "dd/MM/yyyy HH:mm:ss"), // Column X
          },
        };
      }
    }

    return {
      success: false,
      message: "ไม่พบข้อมูลการจอง",
    };
  } catch (error) {
    return {
      success: false,
      message: "เกิดข้อผิดพลาด: " + error.message,
    };
  }
}

/**
 * Create Booking
 * สร้างการจองใหม่ (OP, Owner)
 */
function createBooking(sessionToken, bookingData) {
  try {
    // Validate session
    const session = validateSession(sessionToken);
    if (!session) {
      return {
        success: false,
        message: "Session หมดอายุ กรุณาเข้าสู่ระบบใหม่",
      };
    }

    // Check permissions (OP and Owner can create)
    if (
      session.role !== CONFIG.ROLES.OP &&
      session.role !== CONFIG.ROLES.OWNER
    ) {
      return {
        success: false,
        message: "คุณไม่มีสิทธิ์สร้างการจอง",
      };
    }

    // Validate required fields
    if (
      !bookingData.bookingDate ||
      !bookingData.travelDate ||
      !bookingData.location ||
      !bookingData.program
    ) {
      return {
        success: false,
        message: "กรุณากรอกข้อมูลให้ครบถ้วน",
      };
    }

    // Calculate total amount (including additional cost)
    const subtotal =
      (bookingData.adults || 0) * (bookingData.adultPrice || 0) +
      (bookingData.children || 0) * (bookingData.childPrice || 0) +
      (bookingData.additionalCost || 0) -
      (bookingData.discount || 0);

    // Calculate total with or without VAT based on includeVat flag
    const includeVat = bookingData.includeVat === true;
    const totalAmount = includeVat ? Math.round(subtotal * 1.07) : subtotal;

    // Generate booking ID
    const bookingId = generateBookingId();
    const now = getCurrentTimestamp();

    // Prepare data
    const newRow = [
      bookingId,
      bookingData.bookingDate,
      bookingData.travelDate,
      bookingData.location,
      bookingData.program,
      bookingData.adults || 0,
      bookingData.children || 0,
      bookingData.adultPrice || 0,
      bookingData.childPrice || 0,
      bookingData.additionalCost || 0, // Column J: ค่าใช้จ่ายเพิ่มเติม
      bookingData.discount || 0, // Column K: ส่วนลด
      includeVat, // Column L: รวม VAT 7% (boolean)
      CONFIG.STATUS.CONFIRM, // Column M: สถานะ
      0, // Column N: เงินโอน (ค่าเริ่มต้น)
      0, // Column O: เงินสด (ค่าเริ่มต้น)
      0, // Column P: Cash on tour (ค่าเริ่มต้น)
      "", // Column Q: Slip URL
      bookingData.agent || "", // Column R: Agent
      bookingData.notes || "", // Column S: หมายเหตุ
      totalAmount, // Column T: ยอดขายต่อรายการ
      session.username, // Column U: Created by
      now, // Column V: Created at
      session.username, // Column W: Updated by
      now, // Column X: Updated at
    ];

    // Append to sheet
    const sheet = getSheet(CONFIG.SHEETS.BOOKING_RAW);
    sheet.appendRow(newRow);

    return {
      success: true,
      message: "สร้างการจองสำเร็จ",
      data: { bookingId: bookingId },
    };
  } catch (error) {
    return {
      success: false,
      message: "เกิดข้อผิดพลาด: " + error.message,
    };
  }
}

/**
 * Update Booking
 * แก้ไขการจอง
 * OP: แก้ไขได้เฉพาะที่ตนเองสร้าง
 * Owner: แก้ไขได้ทั้งหมด
 */
function updateBooking(sessionToken, bookingId, bookingData) {
  try {
    // Validate session
    const session = validateSession(sessionToken);
    if (!session) {
      return {
        success: false,
        message: "Session หมดอายุ กรุณาเข้าสู่ระบบใหม่",
      };
    }

    const sheet = getSheet(CONFIG.SHEETS.BOOKING_RAW);
    const data = sheet.getDataRange().getValues();

    // Find booking
    let rowIndex = -1;
    let createdBy = null;
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === bookingId) {
        rowIndex = i + 1; // Sheet row is 1-indexed
        createdBy = data[i][20]; // Column U: ผู้สร้าง
        break;
      }
    }

    if (rowIndex === -1) {
      return {
        success: false,
        message: "ไม่พบข้อมูลการจอง",
      };
    }

    // Check permissions
    // OP can only edit their own bookings, Owner can edit all
    if (session.role === CONFIG.ROLES.OP && createdBy !== session.username) {
      return {
        success: false,
        message:
          "คุณไม่มีสิทธิ์แก้ไขการจองนี้ (สามารถแก้ไขได้เฉพาะที่ตนเองสร้าง)",
      };
    }

    if (
      session.role !== CONFIG.ROLES.OP &&
      session.role !== CONFIG.ROLES.OWNER
    ) {
      return {
        success: false,
        message: "คุณไม่มีสิทธิ์แก้ไขการจอง",
      };
    }

    // Calculate total amount (including additional cost)
    const subtotal =
      (bookingData.adults || 0) * (bookingData.adultPrice || 0) +
      (bookingData.children || 0) * (bookingData.childPrice || 0) +
      (bookingData.additionalCost || 0) -
      (bookingData.discount || 0);

    // Calculate total with or without VAT based on includeVat flag
    const includeVat = bookingData.includeVat === true;
    const totalAmount = includeVat ? Math.round(subtotal * 1.07) : subtotal;

    const now = getCurrentTimestamp();

    // Update data (keep existing status and slip URL)
    sheet.getRange(rowIndex, 2).setValue(bookingData.bookingDate);
    sheet.getRange(rowIndex, 3).setValue(bookingData.travelDate);
    sheet.getRange(rowIndex, 4).setValue(bookingData.location);
    sheet.getRange(rowIndex, 5).setValue(bookingData.program);
    sheet.getRange(rowIndex, 6).setValue(bookingData.adults || 0);
    sheet.getRange(rowIndex, 7).setValue(bookingData.children || 0);
    sheet.getRange(rowIndex, 8).setValue(bookingData.adultPrice || 0);
    sheet.getRange(rowIndex, 9).setValue(bookingData.childPrice || 0);
    sheet.getRange(rowIndex, 10).setValue(bookingData.additionalCost || 0); // Column J: ค่าใช้จ่ายเพิ่มเติม
    sheet.getRange(rowIndex, 11).setValue(bookingData.discount || 0); // Column K: ส่วนลด
    sheet.getRange(rowIndex, 12).setValue(includeVat); // Column L: รวม VAT 7% (boolean)
    sheet.getRange(rowIndex, 18).setValue(bookingData.agent || ""); // Column R: Agent
    sheet.getRange(rowIndex, 19).setValue(bookingData.notes || ""); // Column S: หมายเหตุ
    sheet.getRange(rowIndex, 20).setValue(totalAmount); // Column T: ยอดขายต่อรายการ
    sheet.getRange(rowIndex, 23).setValue(session.username); // Column W: Updated by
    sheet.getRange(rowIndex, 24).setValue(now); // Column X: Updated at

    return {
      success: true,
      message: "แก้ไขการจองสำเร็จ",
    };
  } catch (error) {
    return {
      success: false,
      message: "เกิดข้อผิดพลาด: " + error.message,
    };
  }
}

/**
 * Delete Booking
 * ลบการจอง (Soft delete - เปลี่ยนสถานะเป็น Cancel)
 * OP: ลบได้เฉพาะที่ตนเองสร้าง
 * Owner: ลบได้ทั้งหมด
 */
function deleteBooking(sessionToken, bookingId, reason) {
  try {
    // Validate session
    const session = validateSession(sessionToken);
    if (!session) {
      return {
        success: false,
        message: "Session หมดอายุ กรุณาเข้าสู่ระบบใหม่",
      };
    }

    const sheet = getSheet(CONFIG.SHEETS.BOOKING_RAW);
    const data = sheet.getDataRange().getValues();

    // Find booking
    let rowIndex = -1;
    let createdBy = null;
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === bookingId) {
        rowIndex = i + 1;
        createdBy = data[i][20]; // Column U: ผู้สร้าง (Correct Index: 20)
        break;
      }
    }

    if (rowIndex === -1) {
      return {
        success: false,
        message: "ไม่พบข้อมูลการจอง",
      };
    }

    // Check permissions
    if (session.role === CONFIG.ROLES.OP && createdBy !== session.username) {
      return {
        success: false,
        message: "คุณไม่มีสิทธิ์ลบการจองนี้ (สามารถลบได้เฉพาะที่ตนเองสร้าง)",
      };
    }

    if (
      session.role !== CONFIG.ROLES.OP &&
      session.role !== CONFIG.ROLES.OWNER
    ) {
      return {
        success: false,
        message: "คุณไม่มีสิทธิ์ลบการจอง",
      };
    }

    // Soft delete - change status to Cancel
    const now = getCurrentTimestamp();
    const oldStatus = data[rowIndex - 1][12]; // Column M: สถานะ (Index 12)

    sheet.getRange(rowIndex, 13).setValue(CONFIG.STATUS.CANCEL); // Column M: Status -> Cancel (Col 13)

    sheet.getRange(rowIndex, 23).setValue(session.username); // Column W: Updated By (Col 23)
    sheet.getRange(rowIndex, 24).setValue(now); // Column X: Updated At (Col 24)

    // Log to History
    try {
      const historySheet = getSheet(CONFIG.SHEETS.BOOKING_STATUS_HISTORY);
      const historyId =
        "HIS" + Utilities.formatDate(new Date(), "GMT+7", "yyyyMMddHHmmssSSS");

      historySheet.appendRow([
        historyId,
        bookingId,
        oldStatus,
        CONFIG.STATUS.CANCEL,
        session.username,
        now,
        reason,
      ]);
    } catch (e) {
      Logger.log("Error logging history: " + e.toString());
    }

    return {
      success: true,
      message: "ยกเลิกการจองสำเร็จ",
    };
  } catch (error) {
    return {
      success: false,
      message: "เกิดข้อผิดพลาด: " + error.message,
    };
  }
}

/**
 * Update Booking Status
 * เปลี่ยนสถานะการจอง (สำหรับ AR/AP, Owner)
 */

// ========================================
// SLIP UPLOAD FUNCTIONS
// ========================================

/**
 * Upload Slip to Google Drive
 * อัพโหลดสลิปการชำระเงินไปยัง Google Drive
 */

/**
 * Update Booking Slip URL
 * อัพเดท URL สลิปในการจอง
 */
function updateBookingSlip(sessionToken, bookingId, slipUrl) {
  try {
    // Validate session
    const session = validateSession(sessionToken);
    if (!session) {
      return {
        success: false,
        message: "Session หมดอายุ กรุณาเข้าสู่ระบบใหม่",
      };
    }

    const sheet = getSheet(CONFIG.SHEETS.BOOKING_RAW);
    const data = sheet.getDataRange().getValues();

    // Find booking
    let rowIndex = -1;
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === bookingId) {
        rowIndex = i + 1;
        break;
      }
    }

    if (rowIndex === -1) {
      return {
        success: false,
        message: "ไม่พบข้อมูลการจอง",
      };
    }

    // Update slip URL (Column Q) and ensure status is CONFIRM (Column M)
    sheet.getRange(rowIndex, 17).setValue(slipUrl); // Column Q: Slip URL
    sheet.getRange(rowIndex, 13).setValue(CONFIG.STATUS.CONFIRM); // Column M: Status → CONFIRM
    sheet.getRange(rowIndex, 23).setValue(session.username); // Column W: Updated by
    sheet.getRange(rowIndex, 24).setValue(getCurrentTimestamp()); // Column X: Updated at

    return {
      success: true,
      message: "อัพเดทสลิปสำเร็จ",
    };
  } catch (error) {
    return {
      success: false,
      message: "เกิดข้อผิดพลาด: " + error.message,
    };
  }
}

// ========================================
// PAYMENT APPROVAL SYSTEM (AR/AP, Owner)
// ========================================

/**
 * Get Bookings for Approval
 * ดึงรายการจองที่รอการอนุมัติ (สำหรับ AR/AP และ Owner)
 */
function getBookingsForApproval(
  sessionToken,
  filterStatus = null,
  startDate = null,
  endDate = null
) {
  try {
    // Validate session
    const session = validateSession(sessionToken);
    if (!session) {
      return {
        success: false,
        message: "Session หมดอายุ กรุณาเข้าสู่ระบบใหม่",
      };
    }

    // Check role (AR/AP, Owner, or Cost only)
    if (
      session.role !== CONFIG.ROLES.AR_AP &&
      session.role !== CONFIG.ROLES.OWNER &&
      session.role !== CONFIG.ROLES.COST
    ) {
      return {
        success: false,
        message: "คุณไม่มีสิทธิ์เข้าถึงหน้านี้",
      };
    }

    // Parse date range
    const filterStart = startDate ? new Date(startDate) : null;
    const filterEnd = endDate ? new Date(endDate) : null;
    if (filterStart) filterStart.setHours(0, 0, 0, 0);
    if (filterEnd) filterEnd.setHours(23, 59, 59, 999);

    const sheet = getSheet(CONFIG.SHEETS.BOOKING_RAW);
    const data = sheet.getDataRange().getValues();

    const bookings = [];

    // Skip header row
    for (let i = 1; i < data.length; i++) {
      const row = data[i];

      // Skip empty rows
      if (!row[0]) continue;

      const status = row[12]; // Column M: สถานะ

      // Filter by status if specified
      if (filterStatus && status !== filterStatus) {
        continue;
      }

      // Filter by Date (Booking Date - Column B)
      if (filterStart || filterEnd) {
        const bookingDate = row[1];
        let dateObj;

        if (bookingDate instanceof Date) {
          dateObj = bookingDate;
        } else {
          dateObj = new Date(bookingDate);
        }

        if (dateObj && !isNaN(dateObj.getTime())) {
          if (filterStart && dateObj < filterStart) continue;
          if (filterEnd && dateObj > filterEnd) continue;
        }
      }

      // Include all bookings except Cancelled (or filter by status)
      if (!filterStatus || status === filterStatus) {
        bookings.push({
          bookingId: row[0],
          bookingDate: formatDate(row[1], "dd/MM/yyyy"),
          travelDate: formatDate(row[2], "dd/MM/yyyy"),
          location: row[3],
          program: row[4],
          adults: row[5],
          children: row[6],
          adultPrice: row[7],
          childPrice: row[8],
          additionalCost: row[9], // Column J: ค่าใช้จ่ายเพิ่มเติม
          discount: row[10], // Column K: ส่วนลด
          includeVat:
            row[11] === true || row[11] === "TRUE" || row[11] === "true", // Column L: รวม VAT 7%
          status: status, // Column M: สถานะ
          transferAmount: row[13] || 0, // Column N: เงินโอน
          cashAmount: row[14] || 0, // Column O: เงินสด
          cashOnTour: row[15] || 0, // Column P: Cash on tour
          slipUrl: row[16], // Column Q: Slip URL
          agent: row[17], // Column R: Agent
          notes: row[18], // Column S: หมายเหตุ
          totalAmount: row[19], // Column T: ยอดขายต่อรายการ
          createdBy: row[20], // Column U: ผู้สร้าง
          createdAt: formatDate(row[21], "dd/MM/yyyy HH:mm:ss"), // Column V: วันที่สร้าง
          updatedBy: row[22], // Column W: ผู้แก้ไขล่าสุด
          updatedAt: formatDate(row[23], "dd/MM/yyyy HH:mm:ss"), // Column X: วันที่แก้ไขล่าสุด
        });
      }
    }

    // Sort by booking date (newest first)
    bookings.sort((a, b) => {
      const dateA = new Date(a.bookingDate.split("/").reverse().join("-"));
      const dateB = new Date(b.bookingDate.split("/").reverse().join("-"));
      return dateB - dateA;
    });

    return {
      success: true,
      data: bookings,
    };
  } catch (error) {
    Logger.log("Get bookings for approval error: " + error.message);
    return {
      success: false,
      message: "เกิดข้อผิดพลาด: " + error.message,
    };
  }
}

/**
 * Update Booking Status
 * อัพเดทสถานะการจอง และบันทึกประวัติ
 */
/**
 * Update Booking Status
 * อัพเดทสถานะการจอง และบันทึกประวัติการเปลี่ยนสถานะ
 */
function updateBookingStatus(sessionToken, bookingId, newStatus, reason = "") {
  try {
    // 1. Validate session
    const session = validateSession(sessionToken);
    if (!session) {
      return {
        success: false,
        message: "Session หมดอายุ กรุณาเข้าสู่ระบบใหม่",
      };
    }

    // 2. Check role (Admin, AR_AP, or Owner only)
    if (
      session.role !== CONFIG.ROLES.ADMIN &&
      session.role !== CONFIG.ROLES.AR_AP &&
      session.role !== CONFIG.ROLES.OWNER
    ) {
      return {
        success: false,
        message: "คุณไม่มีสิทธิ์เปลี่ยนสถานะการจอง",
      };
    }

    // 3. Validate status
    const validStatuses = Object.values(CONFIG.STATUS);
    if (!validStatuses.includes(newStatus)) {
      return {
        success: false,
        message: "สถานะไม่ถูกต้อง",
      };
    }

    const sheet = getSheet(CONFIG.SHEETS.BOOKING_RAW);
    const data = sheet.getDataRange().getValues();

    // 4. Find booking row
    let rowIndex = -1;
    let oldStatus = "";
    let bookingData = null;

    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === bookingId) {
        rowIndex = i + 1;
        oldStatus = data[i][12] || ""; // Column M: สถานะ
        bookingData = data[i];
        break;
      }
    }

    if (rowIndex === -1) {
      return {
        success: false,
        message: "ไม่พบข้อมูลการจองรหัส: " + bookingId,
      };
    }

    // 5. Update Status in Booking_Raw
    const now = getCurrentTimestamp();
    sheet.getRange(rowIndex, 13).setValue(newStatus); // Column M: สถานะ
    sheet.getRange(rowIndex, 23).setValue(session.username); // Column W: ผู้แก้ไขล่าสุด
    sheet.getRange(rowIndex, 24).setValue(now); // Column X: วันที่แก้ไขล่าสุด

    // 6. Save to Status History
    const historyResult = saveStatusHistory(
      bookingId,
      oldStatus,
      newStatus,
      session.username,
      reason
    );

    if (!historyResult.success) {
      Logger.log("Failed to save history: " + historyResult.message);
    }

    return {
      success: true,
      message: `เปลี่ยนสถานะเป็น ${newStatus} เรียบร้อยแล้ว`,
      data: {
        bookingId: bookingId,
        oldStatus: oldStatus,
        newStatus: newStatus,
      },
    };
  } catch (error) {
    Logger.log("Update booking status error: " + error.message);
    return {
      success: false,
      message: "เกิดข้อผิดพลาด: " + error.message,
    };
  }
}

/**
 * Update Payment Amounts
 * อัพเดทข้อมูลการชำระเงิน (เงินโอน, เงินสด, Cash on tour)
 */
function updatePaymentAmounts(
  sessionToken,
  bookingId,
  transferAmount,
  cashAmount,
  cashOnTour
) {
  try {
    // 1. Validate session
    const session = validateSession(sessionToken);
    if (!session) {
      return {
        success: false,
        message: "Session หมดอายุ กรุณาเข้าสู่ระบบใหม่",
      };
    }

    // 2. Check role (AR_AP, Owner, Admin only)
    const allowedRoles = [
      CONFIG.ROLES.OWNER,
      CONFIG.ROLES.ADMIN,
      CONFIG.ROLES.AR_AP,
    ];

    if (!allowedRoles.includes(session.role)) {
      return {
        success: false,
        message: "คุณไม่มีสิทธิ์บันทึกข้อมูลการชำระเงิน",
      };
    }

    // 3. Validate amounts
    if (transferAmount < 0 || cashAmount < 0 || cashOnTour < 0) {
      return {
        success: false,
        message: "จำนวนเงินต้องไม่ติดลบ",
      };
    }

    const sheet = getSheet(CONFIG.SHEETS.BOOKING_RAW);
    const data = sheet.getDataRange().getValues();

    // 4. Find booking row
    let rowIndex = -1;
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === bookingId) {
        rowIndex = i + 1;
        break;
      }
    }

    if (rowIndex === -1) {
      return {
        success: false,
        message: "ไม่พบข้อมูลการจองรหัส: " + bookingId,
      };
    }

    // 5. Update payment amounts
    const now = getCurrentTimestamp();
    sheet.getRange(rowIndex, 14).setValue(transferAmount); // Column N: เงินโอน
    sheet.getRange(rowIndex, 15).setValue(cashAmount); // Column O: เงินสด
    sheet.getRange(rowIndex, 16).setValue(cashOnTour); // Column P: Cash on tour
    sheet.getRange(rowIndex, 23).setValue(session.username); // Column W: ผู้แก้ไขล่าสุด
    sheet.getRange(rowIndex, 24).setValue(now); // Column X: วันที่แก้ไขล่าสุด

    return {
      success: true,
      message: "บันทึกข้อมูลการชำระเงินสำเร็จ",
      data: {
        bookingId: bookingId,
        transferAmount: transferAmount,
        cashAmount: cashAmount,
        cashOnTour: cashOnTour,
        totalPaid: transferAmount + cashAmount + cashOnTour,
      },
    };
  } catch (error) {
    Logger.log("Update payment amounts error: " + error.message);
    return {
      success: false,
      message: "เกิดข้อผิดพลาด: " + error.message,
    };
  }
}

/**
 * Save Status History
 * บันทึกประวัติการเปลี่ยนสถานะ
 */
function saveStatusHistory(
  bookingId,
  oldStatus,
  newStatus,
  changedBy,
  reason = ""
) {
  try {
    const sheet = getSheet(CONFIG.SHEETS.BOOKING_STATUS_HISTORY);

    // Generate history ID
    const historyId = generateUniqueId("HIST");
    const timestamp = getCurrentTimestamp();

    const historyRow = [
      historyId, // Column A: รหัสประวัติ
      bookingId, // Column B: รหัสการจอง
      oldStatus, // Column C: สถานะเดิม
      newStatus, // Column D: สถานะใหม่
      changedBy, // Column E: ผู้เปลี่ยนสถานะ
      timestamp, // Column F: วันที่เปลี่ยนสถานะ
      reason, // Column G: เหตุผล
    ];

    sheet.appendRow(historyRow);

    return {
      success: true,
      message: "บันทึกประวัติสำเร็จ",
    };
  } catch (error) {
    Logger.log("Save status history error: " + error.message);
    return {
      success: false,
      message: "เกิดข้อผิดพลาด: " + error.message,
    };
  }
}

/**
 * Get Booking Status History
 * ดึงประวัติการเปลี่ยนสถานะของการจอง
 */
function getBookingStatusHistory(sessionToken, bookingId) {
  try {
    // Validate session
    const session = validateSession(sessionToken);
    if (!session) {
      return {
        success: false,
        message: "Session หมดอายุ กรุณาเข้าสู่ระบบใหม่",
      };
    }

    const sheet = getSheet(CONFIG.SHEETS.BOOKING_STATUS_HISTORY);
    const data = sheet.getDataRange().getValues();

    const history = [];

    // Skip header row
    for (let i = 1; i < data.length; i++) {
      const row = data[i];

      if (row[1] === bookingId) {
        // Column B: รหัสการจอง
        history.push({
          historyId: row[0],
          bookingId: row[1],
          oldStatus: row[2],
          newStatus: row[3],
          changedBy: row[4],
          changedAt: formatDate(row[5], "dd/MM/yyyy HH:mm:ss"),
          reason: row[6],
        });
      }
    }

    // Sort by date (newest first)
    history.sort((a, b) => {
      const dateA = new Date(
        a.changedAt.split(" ")[0].split("/").reverse().join("-")
      );
      const dateB = new Date(
        b.changedAt.split(" ")[0].split("/").reverse().join("-")
      );
      return dateB - dateA;
    });

    return {
      success: true,
      data: history,
    };
  } catch (error) {
    Logger.log("Get booking status history error: " + error.message);
    return {
      success: false,
      message: "เกิดข้อผิดพลาด: " + error.message,
    };
  }
}

// ========================================
// REPORTS FUNCTIONS
// ========================================

/**
 * Get Daily Sales Report by Location
 * รายงานยอดขายรายวันแยกตามสถานที่ (เฉพาะ Completed)
 */
function getDailySalesReport(sessionToken, startDate, endDate) {
  try {
    // Validate session
    const session = validateSession(sessionToken);
    if (!session) {
      return {
        success: false,
        message: "Session หมดอายุ กรุณาเข้าสู่ระบบใหม่",
      };
    }

    const sheet = getSheet(CONFIG.SHEETS.BOOKING_RAW);
    const data = sheet.getDataRange().getValues();

    // Parse dates
    // Fix: Parse YYYY-MM-DD string manually to ensure it uses Script Timezone (not UTC)
    let filterStartDate, filterEndDate;

    if (
      typeof startDate === "string" &&
      /^\d{4}-\d{2}-\d{2}$/.test(startDate)
    ) {
      const [sy, sm, sd] = startDate.split("-").map(Number);
      filterStartDate = new Date(sy, sm - 1, sd);
    } else {
      filterStartDate = new Date(startDate);
    }

    if (typeof endDate === "string" && /^\d{4}-\d{2}-\d{2}$/.test(endDate)) {
      const [ey, em, ed] = endDate.split("-").map(Number);
      filterEndDate = new Date(ey, em - 1, ed);
    } else {
      filterEndDate = new Date(endDate);
    }

    filterStartDate.setHours(0, 0, 0, 0);
    filterEndDate.setHours(23, 59, 59, 999);

    // Aggregate by location
    const locationSales = {};

    for (let i = 1; i < data.length; i++) {
      const row = data[i];

      // Skip empty rows
      if (!row[0]) continue;

      const status = row[12]; // Column M: Status

      // Robust Date Parsing for Booking Date (Column B / Index 1)
      let bookingDate;
      const dateVal = row[1];
      if (dateVal instanceof Date) {
        bookingDate = dateVal;
      } else {
        // Try parsing string manually if standard parsing fails or for specific formats
        // Assuming DD/MM/YYYY or similar if string
        bookingDate = new Date(dateVal);
      }

      const location = row[3] || "ไม่ระบุ"; // Column D: Location
      const totalAmount = Number(row[19]) || 0; // Column T: Total Amount

      // Filter: Only Completed status
      if (status !== CONFIG.STATUS.COMPLETE) continue;

      // Filter: Date range
      if (bookingDate < filterStartDate || bookingDate > filterEndDate)
        continue;

      // Aggregate
      if (!locationSales[location]) {
        locationSales[location] = {
          location: location,
          bookingCount: 0,
          totalSales: 0,
        };
      }

      locationSales[location].bookingCount++;
      locationSales[location].totalSales += totalAmount;
    }

    // Convert to array and sort by total sales (descending)
    const result = Object.values(locationSales).sort(
      (a, b) => b.totalSales - a.totalSales
    );

    return {
      success: true,
      data: {
        dailySales: result,
      },
    };
  } catch (error) {
    Logger.log("Get daily sales report error: " + error.message);
    return {
      success: false,
      message: "เกิดข้อผิดพลาด: " + error.message,
    };
  }
}

/**
 * Get Program Summary Report
 * รายงานสรุปยอดแต่ละโปรแกรม
 */
function getProgramSummaryReport(sessionToken, startDate, endDate) {
  try {
    // Validate session
    const session = validateSession(sessionToken);
    if (!session) {
      return {
        success: false,
        message: "Session หมดอายุ กรุณาเข้าสู่ระบบใหม่",
      };
    }

    const sheet = getSheet(CONFIG.SHEETS.BOOKING_RAW);
    const data = sheet.getDataRange().getValues();

    // Parse dates
    // Fix: Parse YYYY-MM-DD string manually to ensure it uses Script Timezone (not UTC)
    let filterStartDate, filterEndDate;

    if (
      typeof startDate === "string" &&
      /^\d{4}-\d{2}-\d{2}$/.test(startDate)
    ) {
      const [sy, sm, sd] = startDate.split("-").map(Number);
      filterStartDate = new Date(sy, sm - 1, sd);
    } else {
      filterStartDate = new Date(startDate);
    }

    if (typeof endDate === "string" && /^\d{4}-\d{2}-\d{2}$/.test(endDate)) {
      const [ey, em, ed] = endDate.split("-").map(Number);
      filterEndDate = new Date(ey, em - 1, ed);
    } else {
      filterEndDate = new Date(endDate);
    }

    filterStartDate.setHours(0, 0, 0, 0);
    filterEndDate.setHours(23, 59, 59, 999);

    // Aggregate by program
    const programSummary = {};

    for (let i = 1; i < data.length; i++) {
      const row = data[i];

      // Skip empty rows
      if (!row[0]) continue;

      const status = row[12]; // Column M: Status

      // Robust Date Parsing for Booking Date (Column B / Index 1)
      let bookingDate;
      const dateVal = row[1];
      if (dateVal instanceof Date) {
        bookingDate = dateVal;
      } else {
        bookingDate = new Date(dateVal);
      }

      const program = row[4] || "ไม่ระบุ"; // Column E: Program
      const adults = Number(row[5]) || 0; // Column F: Adults
      const children = Number(row[6]) || 0; // Column G: Children
      const totalAmount = Number(row[19]) || 0; // Column T: Total Amount

      // Filter: Only Completed status
      if (status !== CONFIG.STATUS.COMPLETE) continue;

      // Filter: Date range
      if (bookingDate < filterStartDate || bookingDate > filterEndDate)
        continue;

      // Aggregate
      if (!programSummary[program]) {
        programSummary[program] = {
          program: program,
          bookingCount: 0,
          totalAdults: 0,
          totalChildren: 0,
          totalPeople: 0,
          totalRevenue: 0,
        };
      }

      programSummary[program].bookingCount++;
      programSummary[program].totalAdults += adults;
      programSummary[program].totalChildren += children;
      programSummary[program].totalPeople += adults + children;
      programSummary[program].totalRevenue += totalAmount;
    }

    // Convert to array and sort by total revenue (descending)
    const result = Object.values(programSummary).sort(
      (a, b) => b.totalRevenue - a.totalRevenue
    );

    return {
      success: true,
      data: {
        programSummary: result,
      },
    };
  } catch (error) {
    Logger.log("Get program summary report error: " + error.message);
    return {
      success: false,
      message: "เกิดข้อผิดพลาด: " + error.message,
    };
  }
}

// ========================================
// REFUND SYSTEM - BACKEND FUNCTIONS
// ========================================

/**
 * ========================================
 * UNIFIED ID GENERATION SYSTEM
 * ========================================
 * สร้างรหัสอัตโนมัติในรูปแบบเดียวกัน: PREFIX-YYYYMMDD-XXX
 * - BK: Booking (การจอง)
 * - HIS: History (ประวัติการเปลี่ยนสถานะ)
 * - CUS: Customer (ลูกค้า)
 * - REF: Refund (การคืนเงิน)
 */

/**
 * Generate ID with Unified Pattern
 * สร้างรหัสในรูปแบบ: PREFIX-YYYYMMDD-XXX
 * @param {string} prefix - คำนำหน้า (BK, HIS, CUS, REF)
 * @param {string} sheetName - ชื่อ Sheet ที่เก็บข้อมูล
 * @param {number} columnIndex - Index ของคอลัมน์ที่เก็บรหัส (0-based)
 * @returns {string} - รหัสที่สร้างขึ้น
 */
function generateId(prefix, sheetName, columnIndex = 0) {
  try {
    const ss = getSpreadsheet();
    const sheet = ss.getSheetByName(sheetName);

    if (!sheet) {
      throw new Error(`ไม่พบ Sheet: ${sheetName}`);
    }

    const data = sheet.getDataRange().getValues();
    const today = Utilities.formatDate(new Date(), "Asia/Bangkok", "yyyyMMdd");
    const fullPrefix = `${prefix}-${today}-`;

    let maxNumber = 0;
    for (let i = 1; i < data.length; i++) {
      const id = data[i][columnIndex];
      if (id && id.toString().startsWith(fullPrefix)) {
        const parts = id.split("-");
        if (parts.length >= 3) {
          const num = parseInt(parts[2]);
          if (num > maxNumber) maxNumber = num;
        }
      }
    }

    const newNumber = String(maxNumber + 1).padStart(3, "0");
    return fullPrefix + newNumber;
  } catch (error) {
    Logger.log(`Generate ID error (${prefix}): ` + error.message);
    throw error;
  }
}

/**
 * Generate Booking ID
 * สร้างรหัสการจองอัตโนมัติ รูปแบบ: BK-YYYYMMDD-XXX
 */
function generateBookingId() {
  return generateId("BK", CONFIG.SHEETS.BOOKING_RAW, 0);
}

/**
 * Generate History ID
 * สร้างรหัสประวัติอัตโนมัติ รูปแบบ: HIS-YYYYMMDD-XXX
 */
function generateHistoryId() {
  return generateId("HIS", CONFIG.SHEETS.BOOKING_STATUS_HISTORY, 0);
}

/**
 * Generate Customer ID
 * สร้างรหัสลูกค้าอัตโนมัติ รูปแบบ: CUS-YYYYMMDD-XXX
 */
function generateCustomerId() {
  return generateId("CUS", "Customer", 0);
}

/**
 * Generate Refund ID
 * สร้างรหัสการคืนเงินอัตโนมัติ รูปแบบ: REF-YYYYMMDD-XXX
 */
function generateRefundId() {
  return generateId("REF", "Refund", 0);
}

/**
 * Upload Refund Slip to Google Drive
 * อัปโหลดสลิปการโอนเงินคืนไป Google Drive
 * @param {Object} slipFile - ข้อมูลไฟล์ {data, mimeType, filename}
 * @param {string} refundId - รหัสการคืนเงิน
 * @returns {string} - URL ของไฟล์ที่อัปโหลด
 */
function uploadRefundSlip(slipFile, refundId) {
  try {
    // Validate folder ID
    if (
      !CONFIG.REFUND_SLIPS_FOLDER_ID ||
      CONFIG.REFUND_SLIPS_FOLDER_ID.trim() === ""
    ) {
      throw new Error(
        "ไม่พบ Folder ID สำหรับสลิปการคืนเงิน กรุณาตรวจสอบ CONFIG.REFUND_SLIPS_FOLDER_ID"
      );
    }

    // Get Refund Slips folder directly by ID
    let folder;
    try {
      folder = DriveApp.getFolderById(CONFIG.REFUND_SLIPS_FOLDER_ID);
    } catch (folderError) {
      Logger.log(
        "Cannot access folder with ID: " + CONFIG.REFUND_SLIPS_FOLDER_ID
      );
      Logger.log("Error: " + folderError.message);
      throw new Error(
        "ไม่สามารถเข้าถึง Folder สลิปการคืนเงินได้ กรุณาตรวจสอบ Folder ID: " +
          CONFIG.REFUND_SLIPS_FOLDER_ID
      );
    }

    const blob = Utilities.newBlob(
      Utilities.base64Decode(slipFile.data),
      slipFile.mimeType,
      refundId + "_" + slipFile.filename
    );
    const file = folder.createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    return file.getUrl();
  } catch (error) {
    Logger.log("Upload refund slip error: " + error.message);
    throw new Error("ไม่สามารถอัปโหลดสลิปได้: " + error.message);
  }
}

/**
 * Upload Booking Slip to Google Drive
 * อัปโหลดสลิปการชำระเงินการจองไป Google Drive
 * @param {Object} slipFile - ข้อมูลไฟล์ {data, mimeType, filename}
 * @param {string} bookingId - รหัสการจอง
 * @returns {string} - URL ของไฟล์ที่อัปโหลด
 */
/**
 * Upload Booking Slip to Google Drive
 * อัปโหลดสลิปการชำระเงินการจองไป Google Drive
 * @param {string} sessionToken - Session Token ของผู้ใช้
 * @param {Object} slipFile - ข้อมูลไฟล์ {data, mimeType, filename}
 * @param {string} bookingId - รหัสการจอง
 * @returns {string} - URL ของไฟล์ที่อัปโหลด
 */
function uploadBookingSlip(sessionToken, slipFile, bookingId) {
  try {
    // Validate session
    const session = validateSession(sessionToken);
    if (!session) {
      throw new Error("Session หมดอายุ กรุณาเข้าสู่ระบบใหม่");
    }

    // Validate folder ID
    if (
      !CONFIG.BOOKING_SLIPS_FOLDER_ID ||
      CONFIG.BOOKING_SLIPS_FOLDER_ID.trim() === ""
    ) {
      throw new Error(
        "ไม่พบ Folder ID สำหรับสลิปการจอง กรุณาตรวจสอบ CONFIG.BOOKING_SLIPS_FOLDER_ID"
      );
    }

    // Get Booking Slips folder directly by ID
    let folder;
    try {
      folder = DriveApp.getFolderById(CONFIG.BOOKING_SLIPS_FOLDER_ID);
    } catch (folderError) {
      Logger.log(
        "Cannot access folder with ID: " + CONFIG.BOOKING_SLIPS_FOLDER_ID
      );
      Logger.log("Error: " + folderError.message);
      throw new Error(
        "ไม่สามารถเข้าถึง Folder สลิปการจองได้ กรุณาตรวจสอบ Folder ID: " +
          CONFIG.BOOKING_SLIPS_FOLDER_ID
      );
    }

    const blob = Utilities.newBlob(
      Utilities.base64Decode(slipFile.data),
      slipFile.mimeType,
      bookingId + "_" + slipFile.filename
    );
    const file = folder.createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    return file.getUrl();
  } catch (error) {
    Logger.log("Upload booking slip error: " + error.message);
    throw new Error("ไม่สามารถอัปโหลดสลิปได้: " + error.message);
  }
}

/**
 * Create Customer
 * สร้างข้อมูลลูกค้าใหม่
 */
function createCustomer(sessionToken, customerData) {
  try {
    // Validate session
    const session = validateSession(sessionToken);
    if (!session) {
      return {
        success: false,
        message: "Session หมดอายุ กรุณาเข้าสู่ระบบใหม่",
      };
    }

    // Validate input
    if (
      !customerData.customerName ||
      !customerData.bankName ||
      !customerData.accountNumber
    ) {
      return { success: false, message: "กรุณากรอกข้อมูลให้ครบถ้วน" };
    }

    const ss = getSpreadsheet();
    const sheet = ss.getSheetByName("Customer");

    if (!sheet) {
      return {
        success: false,
        message: "ไม่พบ Sheet: Customer กรุณาสร้าง Sheet ก่อน",
      };
    }

    const data = sheet.getDataRange().getValues();
    let existingCustomerId = null;
    let foundIndex = -1;

    // Check if customer name already exists
    for (let i = 1; i < data.length; i++) {
      if (data[i][1] === customerData.customerName) {
        existingCustomerId = data[i][0];
        foundIndex = i;
        break;
      }
    }

    const now = Utilities.formatDate(
      new Date(),
      "Asia/Bangkok",
      "dd/MM/yyyy HH:mm:ss"
    );

    if (foundIndex !== -1) {
      // Name matches, check if bank or account is different
      const existingBank = data[foundIndex][2];
      const existingAcc = String(data[foundIndex][3]).replace(/^'/, "");
      const newAcc = String(customerData.accountNumber).replace(/^'/, "");

      if (existingBank !== customerData.bankName || existingAcc !== newAcc) {
        // Update Bank and Account (Columns 3, 4)
        sheet
          .getRange(foundIndex + 1, 3, 1, 2)
          .setValues([
            [customerData.bankName, "'" + customerData.accountNumber],
          ]);
        // Update Last Edited info (Columns 7, 8)
        sheet
          .getRange(foundIndex + 1, 7, 1, 2)
          .setValues([[session.username, now]]);
        return {
          success: true,
          message: "อัปเดตข้อมูลบัญชีลูกค้าเดิมเรียบร้อย",
          data: {
            customerId: existingCustomerId,
            customerName: customerData.customerName,
            bankName: customerData.bankName,
            accountNumber: customerData.accountNumber,
            updatedBy: session.username,
            updatedAt: now,
          },
        };
      } else {
        // Details are the same, just return existing data
        return {
          success: true,
          message: "ใช้ข้อมูลลูกค้าเดิมที่มีอยู่ในระบบ",
          data: {
            customerId: existingCustomerId,
            customerName: customerData.customerName,
            bankName: customerData.bankName,
            accountNumber: customerData.accountNumber,
          },
        };
      }
    }

    // Case: New Customer
    const customerId = generateCustomerId();

    // Prepare data
    const rowData = [
      customerId,
      customerData.customerName,
      customerData.bankName,
      "'" + customerData.accountNumber,
      session.username,
      now,
      session.username,
      now,
    ];

    // Append to sheet
    sheet.appendRow(rowData);

    return {
      success: true,
      message: "สร้างข้อมูลลูกค้าใหม่สำเร็จ",
      data: {
        customerId: customerId,
        customerName: customerData.customerName,
        bankName: customerData.bankName,
        accountNumber: customerData.accountNumber,
        createdBy: session.username,
        createdAt: now,
      },
    };
  } catch (error) {
    Logger.log("Create/Update customer error: " + error.message);
    return { success: false, message: "เกิดข้อผิดพลาด: " + error.message };
  }
}

/**
 * Get All Customers
 * ดึงข้อมูลลูกค้าทั้งหมด
 */
function getAllCustomers(sessionToken) {
  try {
    // Validate session
    const session = validateSession(sessionToken);
    if (!session) {
      return {
        success: false,
        message: "Session หมดอายุ กรุณาเข้าสู่ระบบใหม่",
      };
    }

    const ss = getSpreadsheet();
    const sheet = ss.getSheetByName("Customer");

    if (!sheet) {
      return { success: false, message: "ไม่พบ Sheet: Customer" };
    }

    const data = sheet.getDataRange().getValues();
    const customers = [];

    for (let i = 1; i < data.length; i++) {
      if (data[i][0]) {
        customers.push({
          customerId: data[i][0],
          customerName: data[i][1],
          bankName: data[i][2],
          accountNumber: data[i][3],
          createdBy: data[i][4],
          createdAt: data[i][5],
          updatedBy: data[i][6],
          updatedAt: data[i][7],
        });
      }
    }

    return {
      success: true,
      message: "ดึงข้อมูลลูกค้าสำเร็จ",
      data: customers,
    };
  } catch (error) {
    Logger.log("Get all customers error: " + error.message);
    return { success: false, message: "เกิดข้อผิดพลาด: " + error.message };
  }
}

/**
 * Create Refund
 * สร้างรายการคืนเงินใหม่
 */
function createRefund(sessionToken, refundData, slipFile) {
  try {
    // Validate session
    const session = validateSession(sessionToken);
    if (!session) {
      return {
        success: false,
        message: "Session หมดอายุ กรุณาเข้าสู่ระบบใหม่",
      };
    }

    // Check permission (only AR/AP and Owner)
    if (!["AR_AP", "Owner"].includes(session.role)) {
      return {
        success: false,
        message: "คุณไม่มีสิทธิ์ในการสร้างรายการคืนเงิน",
      };
    }

    // Validate input
    if (
      !refundData.bookingId ||
      !refundData.customerId ||
      !refundData.refundAmount
    ) {
      return { success: false, message: "กรุณากรอกข้อมูลให้ครบถ้วน" };
    }

    if (refundData.refundAmount <= 0) {
      return { success: false, message: "ยอดเงินคืนต้องมากกว่า 0" };
    }

    const ss = getSpreadsheet();
    const refundSheet = ss.getSheetByName("Refund");
    const customerSheet = ss.getSheetByName("Customer");

    if (!refundSheet) {
      return {
        success: false,
        message: "ไม่พบ Sheet: Refund กรุณาสร้าง Sheet ก่อน",
      };
    }
    if (!customerSheet) {
      return { success: false, message: "ไม่พบ Sheet: Customer" };
    }

    // Verify customer exists
    const customerData = customerSheet.getDataRange().getValues();
    let customerInfo = null;
    for (let i = 1; i < customerData.length; i++) {
      if (customerData[i][0] === refundData.customerId) {
        customerInfo = {
          customerId: customerData[i][0],
          customerName: customerData[i][1],
          bankName: customerData[i][2],
          accountNumber: customerData[i][3],
        };
        break;
      }
    }

    if (!customerInfo) {
      return { success: false, message: "ไม่พบข้อมูลลูกค้า" };
    }

    // Generate Refund ID
    const refundId = generateRefundId();

    // Upload slip
    let slipUrl = "";
    if (slipFile && slipFile.data) {
      slipUrl = uploadRefundSlip(slipFile, refundId);
    }

    const now = Utilities.formatDate(
      new Date(),
      "Asia/Bangkok",
      "dd/MM/yyyy HH:mm:ss"
    );

    // Prepare data
    const rowData = [
      refundId,
      refundData.bookingId,
      customerInfo.customerName, // Column C (New)
      customerInfo.bankName, // Column D (New)
      "'" + customerInfo.accountNumber, // Column E (New)
      refundData.refundAmount, // Column G (Shifted from D)
      slipUrl, // Column H (Shifted from E)
      refundData.note || "", // Column I (Shifted from F)
      session.username, // Column J (Shifted from G)
      now, // Column K (Shifted from H)
    ];

    // Append to sheet
    refundSheet.appendRow(rowData);

    return {
      success: true,
      message: "สร้างรายการคืนเงินสำเร็จ",
      data: {
        refundId: refundId,
        bookingId: refundData.bookingId,
        customerId: customerInfo.customerId,
        customerName: customerInfo.customerName,
        bankName: customerInfo.bankName,
        accountNumber: customerInfo.accountNumber,
        refundAmount: refundData.refundAmount,
        slipUrl: slipUrl,
        note: refundData.note || "",
        createdBy: session.username,
        createdAt: now,
      },
    };
  } catch (error) {
    Logger.log("Create refund error: " + error.message);
    return { success: false, message: "เกิดข้อผิดพลาด: " + error.message };
  }
}

/**
 * Get All Refunds
 * ดึงข้อมูลการคืนเงินทั้งหมด
 */
function getAllRefunds(sessionToken, dateFrom = "", dateTo = "") {
  try {
    // Validate session
    const session = validateSession(sessionToken);
    if (!session) {
      return {
        success: false,
        message: "Session หมดอายุ กรุณาเข้าสู่ระบบใหม่",
      };
    }

    // Check permission (only AR/AP, Owner and Cost)
    if (!["AR_AP", "Owner", "Cost"].includes(session.role)) {
      return { success: false, message: "คุณไม่มีสิทธิ์ในการดูรายการคืนเงิน" };
    }

    const ss = getSpreadsheet();
    const refundSheet = ss.getSheetByName("Refund");
    const customerSheet = ss.getSheetByName("Customer");

    if (!refundSheet) {
      return { success: false, message: "ไม่พบ Sheet: Refund" };
    }
    if (!customerSheet) {
      return { success: false, message: "ไม่พบ Sheet: Customer" };
    }

    const refundData = refundSheet.getDataRange().getValues();
    const customerData = customerSheet.getDataRange().getValues();

    // Create customer lookup map
    const customerMap = {};
    for (let i = 1; i < customerData.length; i++) {
      customerMap[customerData[i][0]] = {
        customerName: customerData[i][1],
        bankName: customerData[i][2],
        accountNumber: customerData[i][3],
      };
    }

    const refunds = [];
    for (let i = 1; i < refundData.length; i++) {
      if (refundData[i][0]) {
        const createdAtDate = refundData[i][9]; // Column J (index 9)

        // Apply date filter if provided
        if (dateFrom || dateTo) {
          const refundDate = new Date(createdAtDate);
          const refundDateStr = Utilities.formatDate(
            refundDate,
            Session.getScriptTimeZone(),
            "yyyy-MM-dd"
          );

          // Check if date is within range
          if (dateFrom && refundDateStr < dateFrom) continue;
          if (dateTo && refundDateStr > dateTo) continue;
        }

        refunds.push({
          refundId: refundData[i][0],
          bookingId: refundData[i][1],
          customerName: refundData[i][2] || "N/A",
          bankName: refundData[i][3] || "N/A",
          accountNumber: refundData[i][4] || "N/A",
          refundAmount: refundData[i][5],
          slipUrl: refundData[i][6],
          note: refundData[i][7],
          createdBy: refundData[i][8],
          createdAt: formatDate(createdAtDate, "dd/MM/yyyy HH:mm:ss"),
        });
      }
    }

    // Sort by created date (newest first)
    refunds.sort((a, b) => {
      const dateA = new Date(a.createdAt);
      const dateB = new Date(b.createdAt);
      return dateB - dateA;
    });

    return {
      success: true,
      message: "ดึงข้อมูลการคืนเงินสำเร็จ",
      data: refunds,
    };
  } catch (error) {
    Logger.log("Get all refunds error: " + error.message);
    return { success: false, message: "เกิดข้อผิดพลาด: " + error.message };
  }
}

/**
 * Get Refund by Booking ID
 * ดึงข้อมูลการคืนเงินตามรหัสการจอง
 */
function getRefundByBookingId(sessionToken, bookingId) {
  try {
    // Validate session
    const session = validateSession(sessionToken);
    if (!session) {
      return {
        success: false,
        message: "Session หมดอายุ กรุณาเข้าสู่ระบบใหม่",
      };
    }

    const ss = getSpreadsheet();
    const refundSheet = ss.getSheetByName("Refund");
    const customerSheet = ss.getSheetByName("Customer");

    if (!refundSheet || !customerSheet) {
      return { success: false, message: "ไม่พบ Sheet ที่จำเป็น" };
    }

    const refundData = refundSheet.getDataRange().getValues();

    // Find refund by booking ID
    for (let i = 1; i < refundData.length; i++) {
      if (refundData[i][1] === bookingId) {
        return {
          success: true,
          message: "พบข้อมูลการคืนเงิน",
          data: {
            refundId: refundData[i][0],
            bookingId: refundData[i][1],
            customerName: refundData[i][2] || "N/A",
            bankName: refundData[i][3] || "N/A",
            accountNumber: refundData[i][4] || "N/A",
            refundAmount: refundData[i][5],
            slipUrl: refundData[i][6],
            note: refundData[i][7],
            createdBy: refundData[i][8],
            createdAt: formatDate(refundData[i][9], "dd/MM/yyyy HH:mm:ss"),
          },
        };
      }
    }

    return {
      success: false,
      message: "ไม่พบข้อมูลการคืนเงินสำหรับรหัสการจองนี้",
    };
  } catch (error) {
    Logger.log("Get refund by booking ID error: " + error.message);
    return { success: false, message: "เกิดข้อผิดพลาด: " + error.message };
  }
}

/**
 * Update Current User Profile
 */
function updateCurrentUserProfile(sessionToken, profileData) {
  try {
    const session = validateSession(sessionToken);
    if (!session)
      return {
        success: false,
        message: "Session หมดอายุ กรุณาเข้าสู่ระบบใหม่",
      };

    const sheet = getSheet(CONFIG.SHEETS.USERS);
    const data = sheet.getDataRange().getValues();
    let rowIndex = -1;

    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === session.userId) {
        rowIndex = i + 1;
        break;
      }
    }

    if (rowIndex === -1)
      return { success: false, message: "ไม่พบข้อมูลผู้ใช้" };

    // Update Full Name (Column E: index 4) and Phone (Column F: index 5)
    if (profileData.fullName)
      sheet.getRange(rowIndex, 5).setValue(profileData.fullName);
    if (profileData.phone)
      sheet.getRange(rowIndex, 6).setValue("'" + profileData.phone);

    // Update timestamp (Column J: index 9)
    sheet.getRange(rowIndex, 10).setValue(getCurrentTimestamp());

    return {
      success: true,
      message: "อัพเดทข้อมูลส่วนตัวสำเร็จ",
      data: {
        fullName: profileData.fullName || session.fullName,
        phone: profileData.phone || session.phone,
      },
    };
  } catch (error) {
    return { success: false, message: "เกิดข้อผิดพลาด: " + error.message };
  }
}

/**
 * Change Current User Password
 */
function changeCurrentUserPassword(sessionToken, passwordData) {
  try {
    const session = validateSession(sessionToken);
    if (!session)
      return {
        success: false,
        message: "Session หมดอายุ กรุณาเข้าสู่ระบบใหม่",
      };

    const sheet = getSheet(CONFIG.SHEETS.USERS);
    const data = sheet.getDataRange().getValues();
    let rowIndex = -1;
    let currentHashedPassword = "";

    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === session.userId) {
        rowIndex = i + 1;
        currentHashedPassword = data[i][3]; // Column D: index 3
        break;
      }
    }

    if (rowIndex === -1)
      return { success: false, message: "ไม่พบข้อมูลผู้ใช้" };

    // Verify Old Password
    if (hashPassword(passwordData.oldPassword) !== currentHashedPassword) {
      return { success: false, message: "รหัสผ่านเดิมไม่ถูกต้อง" };
    }

    // Hash New Password
    const newHashedPassword = hashPassword(passwordData.newPassword);
    sheet.getRange(rowIndex, 4).setValue(newHashedPassword);

    // Update timestamp
    sheet.getRange(rowIndex, 10).setValue(getCurrentTimestamp());

    return { success: true, message: "เปลี่ยนรหัสผ่านสำเร็จ" };
  } catch (error) {
    return { success: false, message: "เกิดข้อผิดพลาด: " + error.message };
  }
}
