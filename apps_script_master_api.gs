/**
 * InstaLab Master Sheet Database API (v7 - Row Master & Live Device Control)
 * Sheets: Authorized_Users | Topics_Row_Master | Special_Days | LoginLogs
 */

const SHEET_NAMES = {
  USERS: "Authorized_Users",
  TOPICS: "Topics_Row_Master",
  SPECIAL: "Special_Days",
  LOGS: "LoginLogs"
};

function doGet(e) {
  return routeRequest(e);
}

function doPost(e) {
  return routeRequest(e);
}

function routeRequest(e) {
  const p = (e && e.parameter) ? e.parameter : {};
  let body = {};

  if (e && e.postData && e.postData.contents) {
    try {
      body = JSON.parse(e.postData.contents);
    } catch (err) {
      body = {};
    }
  }

  const req = Object.assign({}, p, body);
  const action = String(req.action || "getActivities").trim();

  try {
    if (action === "verifyUser" || action === "verifyLogin") {
      return jsonOut(handleVerifyUser(req));
    }
    if (action === "checkSession") {
      return jsonOut(handleCheckSession(req));
    }
    if (action === "getActivities" || action === "bootstrap" || action === "getMasterData") {
      return jsonOut(handleGetMasterData());
    }
    if (action === "updateProfile") {
      return jsonOut(handleUpdateProfile(req));
    }
    return jsonOut({ ok: false, success: false, message: "Unknown action: " + action });
  } catch (err) {
    return jsonOut({ ok: false, success: false, allowed: false, message: String(err) });
  }
}

function jsonOut(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

function normHeader(h) {
  return String(h || "").trim().toLowerCase().replace(/\s+/g, "_");
}

function formatCell(val, tz) {
  if (val === null || val === undefined) return "";
  if (val instanceof Date && !isNaN(val.getTime())) {
    return Utilities.formatDate(val, tz, "yyyy-MM-dd");
  }
  return String(val).trim();
}

function readSheet(name) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(name);
  if (!sheet) return { sheet: null, headers: [], rows: [] };

  const values = sheet.getDataRange().getValues();
  if (!values || values.length <= 1) return { sheet, headers: [], rows: [] };

  const tz = Session.getScriptTimeZone() || "Asia/Kolkata";
  const headers = values[0].map(normHeader);
  const rows = [];

  for (let i = 1; i < values.length; i++) {
    const r = values[i];
    if (r.join("").trim() === "") continue;
    const obj = { _row: i + 1 };
    headers.forEach((col, idx) => {
      obj[col] = formatCell(r[idx], tz);
    });
    rows.push(obj);
  }

  return { sheet, headers, rows };
}

function isDateActive(todayStr, openStr, closeStr, exactDateStr) {
  if (openStr && closeStr) return todayStr >= openStr && todayStr <= closeStr;
  if (openStr) return todayStr >= openStr;
  if (closeStr) return todayStr <= closeStr;
  if (exactDateStr) return todayStr === exactDateStr;
  return true;
}

/**
 * Verify ONLY Email_ID + Status ('Active') + Expiry_Date + Device_ID against Max_Devices
 */
function handleVerifyUser(req) {
  const email = String(req.email || "").trim().toLowerCase();
  const deviceId = String(req.device_id || req.deviceId || "").trim();
  const deviceInfo = String(req.device_info || req.deviceLabel || "").trim();
  const reqMedium = String(req.medium || "").trim();

  if (!email || email.indexOf("@") === -1) {
    return {
      ok: false,
      success: false,
      allowed: false,
      message: "ದಯವಿಟ್ಟು ಸರಿಯಾದ Email ID ನಮೂದಿಸಿ • Please enter a valid Email ID."
    };
  }

  const { sheet, headers, rows } = readSheet(SHEET_NAMES.USERS);
  if (!sheet || !rows.length) {
    writeLog(email, "", "", deviceId, "DENIED_NO_USERS");
    return {
      ok: false,
      success: false,
      allowed: false,
      message: "Authorized_Users ಶೀಟ್‌ನಲ್ಲಿ ಬಳಕೆದಾರರಿಲ್ಲ • No users in Authorized_Users."
    };
  }

  const tz = Session.getScriptTimeZone() || "Asia/Kolkata";
  const todayStr = Utilities.formatDate(new Date(), tz, "yyyy-MM-dd");

  const lock = LockService.getScriptLock();
  lock.waitLock(10000);

  try {
    for (let i = 0; i < rows.length; i++) {
      const u = rows[i];
      const rowEmail = String(u.email_id || u.email || "").trim().toLowerCase();
      if (rowEmail !== email) continue;

      const status = String(u.status || "").trim().toLowerCase();
      if (status !== "active") {
        writeLog(email, u.teacher_name, u.school_name_en || u.school_name_kn, deviceId, "DENIED_INACTIVE");
        return {
          ok: false,
          success: false,
          allowed: false,
          message: "ಈ Email ID Active ಆಗಿಲ್ಲ • Account is Inactive."
        };
      }

      const expiry = String(u.expiry_date || "").trim().slice(0, 10);
      if (expiry && expiry < todayStr) {
        writeLog(email, u.teacher_name, u.school_name_en || u.school_name_kn, deviceId, "DENIED_EXPIRED");
        return {
          ok: false,
          success: false,
          allowed: false,
          message: "ಪ್ರವೇಶ ಅವಧಿ ಮುಕ್ತಾಯವಾಗಿದೆ (" + expiry + ") • Access expired."
        };
      }

      const maxDevices = Number(u.max_devices) > 0 ? Math.floor(Number(u.max_devices)) : 2;
      const rawConn = String(u.connected_devices || "").trim();
      const connList = rawConn ? rawConn.split(",").map(s => s.trim()).filter(Boolean) : [];
      const existingIds = connList.map(entry => entry.split("|")[0].trim());

      const alreadyPresent = deviceId && existingIds.indexOf(deviceId) !== -1;
      if (!alreadyPresent) {
        if (connList.length >= maxDevices) {
          writeLog(email, u.teacher_name, u.school_name_en || u.school_name_kn, deviceId, "DENIED_MAX_DEVICES");
          return {
            ok: false,
            success: false,
            allowed: false,
            message: "ಗರಿಷ್ಠ ಡಿವೈಸ್ ಮಿತಿ (" + maxDevices + ") ಪೂರ್ಣವಾಗಿದೆ • Max devices reached."
          };
        }
        if (deviceId) {
          const safeInfo = deviceInfo.replace(/[,|]/g, " ").slice(0, 28);
          connList.push(safeInfo ? (deviceId + "|" + safeInfo) : deviceId);
          const connCol = headers.indexOf("connected_devices");
          if (connCol >= 0) {
            sheet.getRange(u._row, connCol + 1).setValue(connList.join(", "));
          }
        }
      }

      const chosenMedium = reqMedium || u.medium || "Both";
      const medCol = headers.indexOf("medium");
      if (medCol >= 0 && reqMedium) {
        sheet.getRange(u._row, medCol + 1).setValue(reqMedium);
      }

      const schoolKn = u.school_name_kn || u.school_name_en || "";
      const schoolEn = u.school_name_en || u.school_name_kn || "";

      writeLog(email, u.teacher_name, schoolEn || schoolKn, deviceId, "ALLOWED");

      return {
        ok: true,
        success: true,
        allowed: true,
        teacher: {
          email: rowEmail,
          name: u.teacher_name || "",
          school_kn: schoolKn,
          school_en: schoolEn,
          school: schoolKn || schoolEn,
          medium: chosenMedium,
          role: u.role || "Teacher",
          device_id: deviceId
        }
      };
    }

    writeLog(email, "", "", deviceId, "DENIED_NOT_FOUND");
    return {
      ok: false,
      success: false,
      allowed: false,
      message: "ಈ Email ID Authorized_Users ಪಟ್ಟಿಯಲ್ಲಿಲ್ಲ • Email ID not authorized."
    };
  } finally {
    lock.releaseLock();
  }
}

/**
 * Live check: If Device_ID was cleared from Connected_Devices in sheet, return valid: false so app auto-logouts
 */
function handleCheckSession(req) {
  const email = String(req.email || "").trim().toLowerCase();
  const deviceId = String(req.device_id || req.deviceId || "").trim();
  if (!email || !deviceId) return { valid: false };

  const { rows } = readSheet(SHEET_NAMES.USERS);
  const tz = Session.getScriptTimeZone() || "Asia/Kolkata";
  const todayStr = Utilities.formatDate(new Date(), tz, "yyyy-MM-dd");

  for (let i = 0; i < rows.length; i++) {
    const u = rows[i];
    if (String(u.email_id || "").trim().toLowerCase() !== email) continue;

    if (String(u.status || "").trim().toLowerCase() !== "active") {
      return { valid: false };
    }
    const expiry = String(u.expiry_date || "").trim().slice(0, 10);
    if (expiry && expiry < todayStr) {
      return { valid: false };
    }

    const rawConn = String(u.connected_devices || "").trim();
    const connList = rawConn ? rawConn.split(",").map(s => s.trim()).filter(Boolean) : [];
    const existingIds = connList.map(entry => entry.split("|")[0].trim());

    if (existingIds.indexOf(deviceId) === -1) {
      writeLog(email, u.teacher_name, u.school_name_en || u.school_name_kn, deviceId, "DEVICE_CLEARED_LOGOUT");
      return { valid: false };
    }

    return {
      valid: true,
      teacher: {
        email: email,
        name: u.teacher_name || "",
        school_kn: u.school_name_kn || u.school_name_en || "",
        school_en: u.school_name_en || u.school_name_kn || "",
        medium: u.medium || "Both"
      }
    };
  }

  return { valid: false };
}

/**
 * Fetch Topics_Row_Master and Special_Days
 */
function handleGetMasterData() {
  const tz = Session.getScriptTimeZone() || "Asia/Kolkata";
  const todayStr = Utilities.formatDate(new Date(), tz, "yyyy-MM-dd");

  const topicRows = readSheet(SHEET_NAMES.TOPICS).rows
    .filter(r => String(r.status || "").trim().toLowerCase() === "active")
    .filter(r => isDateActive(todayStr, r.open_date, r.close_date, ""))
    .map(r => ({
      topic_id: r.topic_id || "",
      grade_kn: r.grade_kn || "",
      grade_en: r.grade_en || "",
      topic_name_kn: r.topic_name_kn || r.topic_name_en || "",
      topic_name_en: r.topic_name_en || r.topic_name_kn || "",
      ppt_link: r.ppt_link || "",
      ppt_access: r.ppt_access || "No Access",
      sim_link: r.sim_link || "",
      sim_access: r.sim_access || "No Access",
      worksheet_link: r.worksheet_link || "",
      worksheet_access: r.worksheet_access || "No Access",
      extra_type: r.extra_type || "Extra",
      extra_link: r.extra_link || "",
      extra_access: r.extra_access || "No Access",
      medium: r.medium || "Both",
      search_keywords: r.search_keywords || ""
    }));

  const specialRows = readSheet(SHEET_NAMES.SPECIAL).rows
    .filter(r => String(r.status || "").trim().toLowerCase() === "active")
    .map(r => ({
      date: r.date || "",
      event_title_kn: r.event_title_kn || r.event_title_en || "",
      event_title_en: r.event_title_en || r.event_title_kn || "",
      wish_kn: r.wish_kn || r.wish_en || "",
      wish_en: r.wish_en || r.wish_kn || "",
      sim_link: r.sim_link || "",
      sim_access: r.sim_access || "No Access",
      pdf_link: r.pdf_link || "",
      pdf_access: r.pdf_access || "No Access",
      medium: r.medium || "Both",
      search_keywords: r.search_keywords || "",
      open_date: r.open_date || "",
      close_date: r.close_date || ""
    }));

  const todaySpecial = specialRows.filter(s =>
    isDateActive(todayStr, s.open_date, s.close_date, s.date)
  );

  return {
    ok: true,
    success: true,
    today_date: todayStr,
    topics: topicRows,
    special_days: specialRows,
    today_special: todaySpecial
  };
}

function handleUpdateProfile(req) {
  const email = String(req.email || "").trim().toLowerCase();
  const name = String(req.name || "").trim();
  const school = String(req.school || "").trim();
  const medium = String(req.medium || "").trim();
  if (!email) return { ok: false, success: false };

  const { sheet, headers, rows } = readSheet(SHEET_NAMES.USERS);
  if (!sheet) return { ok: false, success: false };

  const nameCol = headers.indexOf("teacher_name");
  const schKnCol = headers.indexOf("school_name_kn");
  const schEnCol = headers.indexOf("school_name_en");
  const medCol = headers.indexOf("medium");

  for (let i = 0; i < rows.length; i++) {
    const r = rows[i];
    if (String(r.email_id || "").trim().toLowerCase() === email) {
      if (nameCol >= 0 && name) sheet.getRange(r._row, nameCol + 1).setValue(name);
      if (schKnCol >= 0 && school) sheet.getRange(r._row, schKnCol + 1).setValue(school);
      if (schEnCol >= 0 && school && schKnCol < 0) sheet.getRange(r._row, schEnCol + 1).setValue(school);
      if (medCol >= 0 && medium) sheet.getRange(r._row, medCol + 1).setValue(medium);
      return { ok: true, success: true };
    }
  }
  return { ok: false, success: false };
}

function writeLog(email, name, school, deviceId, result) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName(SHEET_NAMES.LOGS);
    if (!sheet) {
      sheet = ss.insertSheet(SHEET_NAMES.LOGS);
      sheet.appendRow(["Timestamp", "Email_ID", "Teacher_Name", "School_Name", "Device_ID", "Result"]);
    }
    sheet.appendRow([new Date(), email, name || "", school || "", deviceId || "", result]);
  } catch (e) {}
}