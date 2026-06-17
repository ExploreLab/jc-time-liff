/**
 * VERSION: 31
 * UPDATED: 2026-06-17
 * CHANGES: Serve GAS as backend API; host LIFF frontend on a stable HTTPS URL.
 * ══════════════════════════════════════════════════════════════
 *  JC-Time Attendance — Google Apps Script (Backend API)
 *  วางโค้ดทั้งหมดนี้ใน Google Apps Script แล้ว Deploy เป็น Web App
 *
 *  ขั้นตอน Deploy:
 *  1. เปิด script.google.com → สร้างโปรเจกต์ใหม่
 *  2. วางโค้ดนี้ทั้งหมด
 *  3. แก้ค่า SPREADSHEET_ID และค่า Config อื่นๆ ด้านล่าง
 *  4. Deploy → New Deployment → Web App
 *     - Execute as: Me
 *     - Who has access: Anyone
 *  5. คัดลอก Web App URL ไปใส่ใน index.html (ค่า API_URL)
 * ══════════════════════════════════════════════════════════════
 */

/* ═══════════════════════════════════════════════════════
   CONFIG — แก้ค่าตรงนี้เท่านั้น
═══════════════════════════════════════════════════════ */
const CONFIG = {
  SPREADSHEET_ID: '19gpAezpGs2EZgLvChBJ-DtTf8_yBhbKrNCNT8_dzc4o',  // ← ID ของ Google Sheet

  // ชื่อ Sheet tabs (ต้องตรงกับ Google Sheet)
  SHEET_USERS:      'Users',
  SHEET_ATTENDANCE: 'Attendance',
  SHEET_LEAVE:      'LeaveRequests',
  SHEET_OT:         'OTRequests',

  // เวลางานเริ่ม/สิ้นสุด (ชั่วโมง 24 ชม.)
  WORK_START_HOUR:  8,
  WORK_START_MIN:   0,
  LATE_THRESHOLD_MIN: 10,  // สายหากเช็คอินหลัง 08:10 น.

  // รัศมี GPS ที่อนุญาต (เมตร)
  GPS_RADIUS_METERS: 100,

  // วันลาพักร้อนต่อปี (ค่า default)
  ANNUAL_LEAVE_DAYS: 10,

  // LINE Notify Token (สำหรับแจ้ง HR เมื่อมีคำขอ)
  LINE_NOTIFY_TOKEN: 'YOUR_LINE_NOTIFY_TOKEN',  // ← Optional
};

/* ═══════════════════════════════════════════════════════
   SPREADSHEET HELPER
═══════════════════════════════════════════════════════ */
function getSheet(name) {
  return SpreadsheetApp
    .openById(CONFIG.SPREADSHEET_ID)
    .getSheetByName(name);
}

function sheetToObjects(sheet) {
  const [headers, ...rows] = sheet.getDataRange().getValues();
  return rows
    .filter(r => r[0] !== '')
    .map(r => {
      const obj = {};
      headers.forEach((h, i) => { obj[h] = r[i]; });
      return obj;
    });
}

function appendRow(sheet, rowObj, headers) {
  sheet.appendRow(headers.map(h => rowObj[h] ?? ''));
}

function updateRow(sheet, rowIndex, rowObj, headers) {
  // rowIndex คือแถวใน sheet (1-based, +1 เพราะ header อยู่แถว 1)
  headers.forEach((h, colIndex) => {
    if (rowObj[h] !== undefined)
      sheet.getRange(rowIndex + 2, colIndex + 1).setValue(rowObj[h]);
  });
}

/* ═══════════════════════════════════════════════════════
   GPS UTILITY
═══════════════════════════════════════════════════════ */
function calcDistance(lat1, lng1, lat2, lng2) {
  const R = 6371000; // เมตร
  const dLat = (lat2 - lat1) * Math.PI / 180;
  const dLng = (lng2 - lng1) * Math.PI / 180;
  const a =
    Math.sin(dLat / 2) ** 2 +
    Math.cos(lat1 * Math.PI / 180) *
    Math.cos(lat2 * Math.PI / 180) *
    Math.sin(dLng / 2) ** 2;
  return R * 2 * Math.atan2(Math.sqrt(a), Math.sqrt(1 - a));
}

/* ═══════════════════════════════════════════════════════
   DATE / TIME UTILITY
═══════════════════════════════════════════════════════ */
function toDateStr(d) {
  // "YYYY-MM-DD"
  const dt = new Date(d);
  const y  = dt.getFullYear();
  const m  = String(dt.getMonth() + 1).padStart(2, '0');
  const day = String(dt.getDate()).padStart(2, '0');
  return `${y}-${m}-${day}`;
}

function toTimeStr(d) {
  // "HH:MM:SS"
  const dt = new Date(d);
  return dt.toTimeString().slice(0, 8);
}

function formatSheetTime(value) {
  if (!value) return null;
  if (value instanceof Date) {
    return Utilities.formatDate(value, 'Asia/Bangkok', 'HH:mm');
  }
  const str = String(value).trim();
  const match = str.match(/(\d{1,2}):(\d{2})(?::\d{2})?/);
  if (match) {
    return `${match[1].padStart(2, '0')}:${match[2]}`;
  }
  const dt = new Date(str);
  if (!isNaN(dt.getTime())) {
    return Utilities.formatDate(dt, 'Asia/Bangkok', 'HH:mm');
  }
  return str.slice(0, 5);
}

function getThailandNow() {
  // Google Apps Script รันใน UTC; แปลงเป็น Asia/Bangkok
  return new Date(new Date().toLocaleString('en-US', { timeZone: 'Asia/Bangkok' }));
}

function isLate(timeStr) {
  const [h, m] = timeStr.split(':').map(Number);
  const totalMin = h * 60 + m;
  const threshold = CONFIG.WORK_START_HOUR * 60 + CONFIG.WORK_START_MIN + CONFIG.LATE_THRESHOLD_MIN;
  return totalMin > threshold;
}

/* ═══════════════════════════════════════════════════════
   LINE NOTIFY (Optional)
═══════════════════════════════════════════════════════ */
function lineNotify(message) {
  if (!CONFIG.LINE_NOTIFY_TOKEN || CONFIG.LINE_NOTIFY_TOKEN === 'YOUR_LINE_NOTIFY_TOKEN') return;
  try {
    UrlFetchApp.fetch('https://notify-api.line.me/api/notify', {
      method: 'post',
      headers: { Authorization: 'Bearer ' + CONFIG.LINE_NOTIFY_TOKEN },
      payload: { message },
    });
  } catch (e) {
    Logger.log('LINE Notify error: ' + e);
  }
}

/* ═══════════════════════════════════════════════════════
   CORS WRAPPER
═══════════════════════════════════════════════════════ */
function jsonResponse(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

/* ═══════════════════════════════════════════════════════
   HTTP HANDLERS
═══════════════════════════════════════════════════════ */
function doGet(e) {
  const action = e.parameter.action;

  // Preview only. LIFF should use a stable static host and call this GAS URL as API.
  if (!action) {
    const tmpl = HtmlService.createTemplateFromFile('index');
    tmpl.API_URL = ScriptApp.getService().getUrl();
    return tmpl.evaluate()
      .setTitle('JC-Time Attendance')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }

  // Parse data parameter for write actions
  let data = {};
  try { data = e.parameter.data ? JSON.parse(e.parameter.data) : {}; } catch (_) {}

  try {
    switch (action) {
      // Read actions
      case 'getUserInfo':         return jsonResponse(getUserInfo(e.parameter));
      case 'getGpsStatus':        return jsonResponse(getGpsStatus(e.parameter));
      case 'getTodayStatus':      return jsonResponse(getTodayStatus(e.parameter));
      case 'getMonthlySummary':   return jsonResponse(getMonthlySummary(e.parameter));
      case 'getRecentAttendance': return jsonResponse(getRecentAttendance(e.parameter));
      case 'getMonthlyHistory':   return jsonResponse(getMonthlyHistory(e.parameter));
      case 'getMyRequests':       return jsonResponse(getMyRequests(e.parameter));
      case 'getPendingLeave':     return jsonResponse(getPendingLeave(e.parameter));
      case 'getPendingOT':        return jsonResponse(getPendingOT(e.parameter));
      // Write actions (via GET + data param)
      case 'checkIn':      return jsonResponse(checkIn(data));
      case 'checkOut':     return jsonResponse(checkOut(data));
      case 'submitLeave':  return jsonResponse(submitLeave(data));
      case 'submitOT':     return jsonResponse(submitOT(data));
      case 'approveLeave': return jsonResponse(approveLeave(data));
      case 'rejectLeave':  return jsonResponse(rejectLeave(data));
      case 'approveOT':    return jsonResponse(approveOT(data));
      case 'rejectOT':     return jsonResponse(rejectOT(data));
      case 'registerUser': return jsonResponse(registerUser(data));
      default:
        return jsonResponse({ error: 'Unknown action: ' + action });
    }
  } catch (err) {
    Logger.log('doGet error: ' + err);
    return jsonResponse({ error: err.toString() });
  }
}

function doPost(e) {
  let body;
  try {
    body = JSON.parse(e.postData.contents);
  } catch {
    return jsonResponse({ success: false, message: 'Invalid JSON' });
  }
  try {
    switch (body.action) {
      case 'checkIn':      return jsonResponse(checkIn(body));
      case 'checkOut':     return jsonResponse(checkOut(body));
      case 'submitLeave':  return jsonResponse(submitLeave(body));
      case 'submitOT':     return jsonResponse(submitOT(body));
      // Admin
      case 'approveLeave': return jsonResponse(approveLeave(body));
      case 'rejectLeave':  return jsonResponse(rejectLeave(body));
      case 'approveOT':    return jsonResponse(approveOT(body));
      case 'rejectOT':     return jsonResponse(rejectOT(body));
      // Setup
      case 'registerUser': return jsonResponse(registerUser(body));
      default:
        return jsonResponse({ success: false, message: 'Unknown action: ' + body.action });
    }
  } catch (err) {
    Logger.log('doPost error: ' + err);
    return jsonResponse({ success: false, message: err.toString() });
  }
}

/* ═══════════════════════════════════════════════════════
   ACTION: getUserInfo
   GET ?action=getUserInfo&userId=U123
═══════════════════════════════════════════════════════ */
function getUserInfo({ userId }) {
  Logger.log('getUserInfo called with userId: ' + userId);
  const users = sheetToObjects(getSheet(CONFIG.SHEET_USERS));
  const user  = users.find(u => u.LineUserID === userId);
  if (!user) return { found: false };
  return {
    found:    true,
    name:     user.Name,
    dept:     user.Department,
    position: user.Position,
    workplaceLat: user.WorkplaceLat,
    workplaceLng: user.WorkplaceLng,
    workplaceName: user.WorkplaceName,
    leaveBalance: Number(user.LeaveBalance) || CONFIG.ANNUAL_LEAVE_DAYS,
  };
}

/* ═══════════════════════════════════════════════════════
   ACTION: getGpsStatus
   GET ?action=getGpsStatus&userId=U123&lat=13.75&lng=100.5
═══════════════════════════════════════════════════════ */
function getGpsStatus({ userId, lat, lng }) {
  const users = sheetToObjects(getSheet(CONFIG.SHEET_USERS));
  const user  = users.find(u => u.LineUserID === userId);
  if (!user) return { inRange: false, message: 'ไม่พบข้อมูลผู้ใช้' };

  const dist = calcDistance(
    Number(lat), Number(lng),
    Number(user.WorkplaceLat), Number(user.WorkplaceLng)
  );
  return {
    inRange:       dist <= CONFIG.GPS_RADIUS_METERS,
    distance:      Math.round(dist),
    workplaceName: user.WorkplaceName || 'สำนักงาน',
  };
}

/* ═══════════════════════════════════════════════════════
   ACTION: getTodayStatus
   GET ?action=getTodayStatus&userId=U123
═══════════════════════════════════════════════════════ */
function getTodayStatus({ userId }) {
  const todayStr   = toDateStr(getThailandNow());
  const attendance = sheetToObjects(getSheet(CONFIG.SHEET_ATTENDANCE));
  const rec        = attendance.find(
    r => r.UserID === userId && toDateStr(r.Date) === todayStr
  );
  if (!rec) return { checkedIn: false, checkedOut: false };
  return {
    checkedIn:  !!rec.TimeIn,
    checkedOut: !!rec.TimeOut,
    timeIn:     formatSheetTime(rec.TimeIn),
    timeOut:    formatSheetTime(rec.TimeOut),
    status:     rec.Status,
  };
}

/* ═══════════════════════════════════════════════════════
   ACTION: getMonthlySummary
   GET ?action=getMonthlySummary&userId=U123
═══════════════════════════════════════════════════════ */
function getMonthlySummary({ userId }) {
  const now        = getThailandNow();
  const month      = now.getMonth() + 1;
  const year       = now.getFullYear();
  const attendance = sheetToObjects(getSheet(CONFIG.SHEET_ATTENDANCE));
  const monthly    = attendance.filter(r => {
    const d = new Date(r.Date);
    return r.UserID === userId &&
           d.getMonth() + 1 === month &&
           d.getFullYear() === year;
  });

  const ot = sheetToObjects(getSheet(CONFIG.SHEET_OT));
  const otPending = ot.filter(
    r => r.UserID === userId && r.Status === 'รออนุมัติ'
  ).length;

  const users = sheetToObjects(getSheet(CONFIG.SHEET_USERS));
  const user  = users.find(u => u.LineUserID === userId);

  return {
    present:      monthly.filter(r => r.Status === 'ปกติ' || r.Status === 'สาย').length,
    late:         monthly.filter(r => r.Status === 'สาย').length,
    leave:        monthly.filter(r => r.Status === 'ลา').length,
    absent:       monthly.filter(r => r.Status === 'ขาดงาน').length,
    leaveBalance: user ? Number(user.LeaveBalance) : CONFIG.ANNUAL_LEAVE_DAYS,
    otPending,
  };
}

/* ═══════════════════════════════════════════════════════
   ACTION: getRecentAttendance
   GET ?action=getRecentAttendance&userId=U123&limit=5
═══════════════════════════════════════════════════════ */
function getRecentAttendance({ userId, limit = 5 }) {
  const attendance = sheetToObjects(getSheet(CONFIG.SHEET_ATTENDANCE));
  const userRecs   = attendance
    .filter(r => r.UserID === userId)
    .sort((a, b) => new Date(b.Date) - new Date(a.Date))
    .slice(0, Number(limit));
  return {
    records: userRecs.map(r => ({
      date:      toDateStr(r.Date),
      timeIn:    formatSheetTime(r.TimeIn),
      timeOut:   formatSheetTime(r.TimeOut),
      status:    r.Status,
      leaveType: r.LeaveType || null,
      leaveApproved: r.LeaveApproved === true || r.LeaveApproved === 'TRUE',
      workplace: r.Workplace || '',
    })),
  };
}

/* ═══════════════════════════════════════════════════════
   ACTION: getMonthlyHistory
   GET ?action=getMonthlyHistory&userId=U123&month=6&year=2025
═══════════════════════════════════════════════════════ */
function getMonthlyHistory({ userId, month, year }) {
  const m = Number(month);
  const y = Number(year);
  const attendance = sheetToObjects(getSheet(CONFIG.SHEET_ATTENDANCE));
  const records    = attendance
    .filter(r => {
      const d = new Date(r.Date);
      return r.UserID === userId &&
             d.getMonth() + 1 === m &&
             d.getFullYear() === y;
    })
    .sort((a, b) => new Date(b.Date) - new Date(a.Date));

  return {
    records: records.map(r => ({
      date:    toDateStr(r.Date),
      timeIn:  formatSheetTime(r.TimeIn),
      timeOut: formatSheetTime(r.TimeOut),
      status:  r.Status,
      leaveType: r.LeaveType || null,
    })),
  };
}

/* ═══════════════════════════════════════════════════════
   ACTION: checkIn
   POST { action, userId, lat, lng }
═══════════════════════════════════════════════════════ */
function checkIn({ userId, lat, lng }) {
  // 1. ตรวจสอบ GPS
  const gps = getGpsStatus({ userId, lat, lng });
  if (!gps.inRange) {
    return { success: false, message: `GPS อยู่ห่าง ${gps.distance} ม. (เกินรัศมีที่กำหนด ${CONFIG.GPS_RADIUS_METERS} ม.)` };
  }

  const now     = getThailandNow();
  const todayStr = toDateStr(now);
  const timeStr  = toTimeStr(now);

  // 2. เช็คว่าเช็คอินซ้ำไหม
  const sheet  = getSheet(CONFIG.SHEET_ATTENDANCE);
  const records = sheetToObjects(sheet);
  const existing = records.find(
    r => r.UserID === userId && toDateStr(r.Date) === todayStr
  );
  if (existing && existing.TimeIn) {
    return { success: false, message: `เช็คอินไปแล้ว เวลา ${formatSheetTime(existing.TimeIn)}` };
  }

  // 3. ดึงข้อมูล user
  const users = sheetToObjects(getSheet(CONFIG.SHEET_USERS));
  const user  = users.find(u => u.LineUserID === userId);
  if (!user) return { success: false, message: 'ไม่พบข้อมูลพนักงาน' };

  // 4. คำนวณสถานะ
  const status = isLate(timeStr) ? 'สาย' : 'ปกติ';

  // 5. สร้าง RecordID
  const recordId = 'ATT' + now.getTime();

  // 6. บันทึกลง Sheet
  const headers = ['RecordID','UserID','Date','TimeIn','TimeOut','LatIn','LngIn','LatOut','LngOut','Status','LeaveType','LeaveApproved','Workplace','Note'];
  const newRow = {
    RecordID:  recordId,
    UserID:    userId,
    Date:      todayStr,
    TimeIn:    timeStr,
    TimeOut:   '',
    LatIn:     lat,
    LngIn:     lng,
    LatOut:    '',
    LngOut:    '',
    Status:    status,
    LeaveType: '',
    LeaveApproved: '',
    Workplace: gps.workplaceName,
    Note:      '',
  };
  appendRow(sheet, newRow, headers);

  Logger.log(`CheckIn: ${userId} at ${timeStr} (${status})`);
  return {
    success: true,
    timeIn:  timeStr.slice(0, 5),
    status,
    workplace: gps.workplaceName,
  };
}

/* ═══════════════════════════════════════════════════════
   ACTION: checkOut
   POST { action, userId, lat, lng }
═══════════════════════════════════════════════════════ */
function checkOut({ userId, lat, lng }) {
  const now      = getThailandNow();
  const todayStr = toDateStr(now);
  const timeStr  = toTimeStr(now);

  const sheet   = getSheet(CONFIG.SHEET_ATTENDANCE);
  const data    = sheet.getDataRange().getValues();
  const headers = data[0];
  const rows    = data.slice(1);

  const rowIdx = rows.findIndex(
    r => r[headers.indexOf('UserID')] === userId &&
         toDateStr(r[headers.indexOf('Date')]) === todayStr &&
         r[headers.indexOf('TimeIn')] !== ''
  );

  if (rowIdx === -1) {
    return { success: false, message: 'ไม่พบข้อมูลการเช็คอินวันนี้' };
  }
  if (rows[rowIdx][headers.indexOf('TimeOut')] !== '') {
    return { success: false, message: `เช็คเอาท์ไปแล้ว เวลา ${formatSheetTime(rows[rowIdx][headers.indexOf('TimeOut')])}` };
  }

  // อัปเดต TimeOut, LatOut, LngOut
  sheet.getRange(rowIdx + 2, headers.indexOf('TimeOut') + 1).setValue(timeStr);
  sheet.getRange(rowIdx + 2, headers.indexOf('LatOut')  + 1).setValue(lat);
  sheet.getRange(rowIdx + 2, headers.indexOf('LngOut')  + 1).setValue(lng);

  Logger.log(`CheckOut: ${userId} at ${timeStr}`);
  return {
    success: true,
    timeOut: timeStr.slice(0, 5),
  };
}

/* ═══════════════════════════════════════════════════════
   ACTION: submitLeave
   POST { action, userId, leaveType, startDate, endDate, reason }
═══════════════════════════════════════════════════════ */
function submitLeave({ userId, leaveType, startDate, endDate, reason }) {
  // ตรวจสอบวันลาคงเหลือ
  const users     = sheetToObjects(getSheet(CONFIG.SHEET_USERS));
  const userIdx   = users.findIndex(u => u.LineUserID === userId);
  if (userIdx === -1) return { success: false, message: 'ไม่พบข้อมูลพนักงาน' };
  const user = users[userIdx];

  // คำนวณจำนวนวัน
  const start = new Date(startDate);
  const end   = new Date(endDate);
  const days  = Math.round((end - start) / 86400000) + 1;

  if (leaveType === 'ลาพักร้อน') {
    const balance = Number(user.LeaveBalance) || 0;
    if (days > balance) {
      return { success: false, message: `วันลาพักร้อนคงเหลือไม่พอ (คงเหลือ ${balance} วัน, ขอ ${days} วัน)` };
    }
  }

  const leaveId = 'LV' + new Date().getTime();
  const headers = ['LeaveID','UserID','LeaveType','StartDate','EndDate','Days','Reason','Status','SubmittedAt','ApprovedBy','ApprovedAt','Note'];
  const newRow  = {
    LeaveID:     leaveId,
    UserID:      userId,
    LeaveType:   leaveType,
    StartDate:   startDate,
    EndDate:     endDate,
    Days:        days,
    Reason:      reason || '',
    Status:      'รออนุมัติ',
    SubmittedAt: toDateStr(getThailandNow()),
    ApprovedBy:  '',
    ApprovedAt:  '',
    Note:        '',
  };
  appendRow(getSheet(CONFIG.SHEET_LEAVE), newRow, headers);

  // แจ้ง HR ผ่าน LINE Notify
  lineNotify(`\n[JC-Time] คำขอลางาน\nพนักงาน: ${user.Name}\nประเภท: ${leaveType}\nวันที่: ${startDate} ถึง ${endDate} (${days} วัน)\nเหตุผล: ${reason || '—'}`);

  Logger.log(`Leave submitted: ${userId} - ${leaveType} ${startDate}~${endDate}`);
  return { success: true, leaveId, days };
}

/* ═══════════════════════════════════════════════════════
   ACTION: submitOT
   POST { action, userId, date, startTime, endTime, purpose }
═══════════════════════════════════════════════════════ */
function submitOT({ userId, date, startTime, endTime, purpose }) {
  const users = sheetToObjects(getSheet(CONFIG.SHEET_USERS));
  const user  = users.find(u => u.LineUserID === userId);
  if (!user) return { success: false, message: 'ไม่พบข้อมูลพนักงาน' };

  const otId   = 'OT' + new Date().getTime();
  const headers = ['OTID','UserID','Date','StartTime','EndTime','Purpose','Status','SubmittedAt','ApprovedBy','ApprovedAt','Note'];
  const newRow  = {
    OTID:        otId,
    UserID:      userId,
    Date:        date,
    StartTime:   startTime,
    EndTime:     endTime,
    Purpose:     purpose,
    Status:      'รออนุมัติ',
    SubmittedAt: toDateStr(getThailandNow()),
    ApprovedBy:  '',
    ApprovedAt:  '',
    Note:        '',
  };
  appendRow(getSheet(CONFIG.SHEET_OT), newRow, headers);

  lineNotify(`\n[JC-Time] คำขอ OT\nพนักงาน: ${user.Name}\nวันที่: ${date}\nเวลา: ${startTime}–${endTime}\nงาน: ${purpose}`);

  Logger.log(`OT submitted: ${userId} - ${date} ${startTime}~${endTime}`);
  return { success: true, otId };
}

/* ═══════════════════════════════════════════════════════
   ACTION: getMyRequests
   GET ?action=getMyRequests&userId=U123&limit=8
═══════════════════════════════════════════════════════ */
function getMyRequests({ userId, limit = 8 }) {
  const leaves = sheetToObjects(getSheet(CONFIG.SHEET_LEAVE))
    .filter(r => r.UserID === userId)
    .map(r => ({
      type:       'leave',
      id:         r.LeaveID,
      leaveType:  r.LeaveType,
      startDate:  toDateStr(r.StartDate),
      endDate:    toDateStr(r.EndDate),
      days:       r.Days,
      status:     r.Status,
      approvedBy: r.ApprovedBy || null,
      submittedAt: toDateStr(r.SubmittedAt),
      _sort: new Date(r.SubmittedAt),
    }));

  const ots = sheetToObjects(getSheet(CONFIG.SHEET_OT))
    .filter(r => r.UserID === userId)
    .map(r => ({
      type:       'ot',
      id:         r.OTID,
      date:       toDateStr(r.Date),
      startTime:  r.StartTime,
      endTime:    r.EndTime,
      purpose:    r.Purpose,
      status:     r.Status,
      approvedBy: r.ApprovedBy || null,
      submittedAt: toDateStr(r.SubmittedAt),
      _sort: new Date(r.SubmittedAt),
    }));

  const all = [...leaves, ...ots]
    .sort((a, b) => b._sort - a._sort)
    .slice(0, Number(limit))
    .map(({ _sort, ...rest }) => rest);

  return { requests: all };
}

/* ═══════════════════════════════════════════════════════
   ADMIN: getPendingLeave
   GET ?action=getPendingLeave
═══════════════════════════════════════════════════════ */
function getPendingLeave() {
  const leaves = sheetToObjects(getSheet(CONFIG.SHEET_LEAVE))
    .filter(r => r.Status === 'รออนุมัติ');
  const users  = sheetToObjects(getSheet(CONFIG.SHEET_USERS));
  return {
    requests: leaves.map(r => {
      const user = users.find(u => u.LineUserID === r.UserID);
      return {
        leaveId:   r.LeaveID,
        userId:    r.UserID,
        userName:  user ? user.Name : '—',
        dept:      user ? user.Department : '—',
        leaveType: r.LeaveType,
        startDate: toDateStr(r.StartDate),
        endDate:   toDateStr(r.EndDate),
        days:      r.Days,
        reason:    r.Reason,
        submitted: toDateStr(r.SubmittedAt),
      };
    }),
  };
}

/* ═══════════════════════════════════════════════════════
   ADMIN: getPendingOT
   GET ?action=getPendingOT
═══════════════════════════════════════════════════════ */
function getPendingOT() {
  const ots   = sheetToObjects(getSheet(CONFIG.SHEET_OT))
    .filter(r => r.Status === 'รออนุมัติ');
  const users = sheetToObjects(getSheet(CONFIG.SHEET_USERS));
  return {
    requests: ots.map(r => {
      const user = users.find(u => u.LineUserID === r.UserID);
      return {
        otId:      r.OTID,
        userId:    r.UserID,
        userName:  user ? user.Name : '—',
        dept:      user ? user.Department : '—',
        date:      toDateStr(r.Date),
        startTime: r.StartTime,
        endTime:   r.EndTime,
        purpose:   r.Purpose,
        submitted: toDateStr(r.SubmittedAt),
      };
    }),
  };
}

/* ═══════════════════════════════════════════════════════
   ADMIN: approveLeave / rejectLeave
   POST { action:'approveLeave', leaveId, approvedBy }
═══════════════════════════════════════════════════════ */
function approveLeave({ leaveId, approvedBy }) {
  return _updateLeaveStatus(leaveId, 'อนุมัติ', approvedBy);
}
function rejectLeave({ leaveId, approvedBy, note }) {
  return _updateLeaveStatus(leaveId, 'ปฏิเสธ', approvedBy, note);
}

function _updateLeaveStatus(leaveId, status, approvedBy, note) {
  const sheet   = getSheet(CONFIG.SHEET_LEAVE);
  const data    = sheet.getDataRange().getValues();
  const headers = data[0];
  const rows    = data.slice(1);

  const rowIdx = rows.findIndex(r => r[headers.indexOf('LeaveID')] === leaveId);
  if (rowIdx === -1) return { success: false, message: 'ไม่พบคำขอลา' };

  const now = toDateStr(getThailandNow());
  sheet.getRange(rowIdx + 2, headers.indexOf('Status')     + 1).setValue(status);
  sheet.getRange(rowIdx + 2, headers.indexOf('ApprovedBy') + 1).setValue(approvedBy || '');
  sheet.getRange(rowIdx + 2, headers.indexOf('ApprovedAt') + 1).setValue(now);
  if (note) sheet.getRange(rowIdx + 2, headers.indexOf('Note') + 1).setValue(note);

  // ถ้าอนุมัติ: ตัดวันลาพักร้อน + บันทึก Attendance
  if (status === 'อนุมัติ') {
    const leave = {
      UserID:    rows[rowIdx][headers.indexOf('UserID')],
      LeaveType: rows[rowIdx][headers.indexOf('LeaveType')],
      StartDate: rows[rowIdx][headers.indexOf('StartDate')],
      EndDate:   rows[rowIdx][headers.indexOf('EndDate')],
      Days:      Number(rows[rowIdx][headers.indexOf('Days')]) || 1,
    };

    // หักวันลาพักร้อน
    if (leave.LeaveType === 'ลาพักร้อน') {
      const usersSheet = getSheet(CONFIG.SHEET_USERS);
      const uData      = usersSheet.getDataRange().getValues();
      const uHeaders   = uData[0];
      const uRows      = uData.slice(1);
      const uIdx       = uRows.findIndex(r => r[uHeaders.indexOf('LineUserID')] === leave.UserID);
      if (uIdx !== -1) {
        const balCol = uHeaders.indexOf('LeaveBalance');
        const cur    = Number(uRows[uIdx][balCol]) || 0;
        usersSheet.getRange(uIdx + 2, balCol + 1).setValue(Math.max(0, cur - leave.Days));
      }
    }

    // บันทึก Attendance แต่ละวันที่ลา
    const attSheet   = getSheet(CONFIG.SHEET_ATTENDANCE);
    const attHeaders = ['RecordID','UserID','Date','TimeIn','TimeOut','LatIn','LngIn','LatOut','LngOut','Status','LeaveType','LeaveApproved','Workplace','Note'];
    const start = new Date(leave.StartDate);
    const end   = new Date(leave.EndDate);
    for (let d = new Date(start); d <= end; d.setDate(d.getDate() + 1)) {
      const dow = d.getDay();
      if (dow === 0 || dow === 6) continue; // ข้ามวันหยุด
      appendRow(attSheet, {
        RecordID:     'LV_ATT_' + d.getTime(),
        UserID:       leave.UserID,
        Date:         toDateStr(d),
        TimeIn:       '', TimeOut: '',
        LatIn: '', LngIn: '', LatOut: '', LngOut: '',
        Status:       'ลา',
        LeaveType:    leave.LeaveType,
        LeaveApproved: 'TRUE',
        Workplace:    '',
        Note:         'อนุมัติโดย ' + (approvedBy || ''),
      }, attHeaders);
    }

    // แจ้งพนักงาน
    lineNotify(`\n[JC-Time] อนุมัติใบลา\nประเภท: ${leave.LeaveType}\nวันที่: ${toDateStr(leave.StartDate)} – ${toDateStr(leave.EndDate)}\nอนุมัติโดย: ${approvedBy || '—'}`);
  }

  Logger.log(`Leave ${leaveId} → ${status} by ${approvedBy}`);
  return { success: true, status };
}

/* ═══════════════════════════════════════════════════════
   ADMIN: approveOT / rejectOT
   POST { action:'approveOT', otId, approvedBy }
═══════════════════════════════════════════════════════ */
function approveOT({ otId, approvedBy }) {
  return _updateOTStatus(otId, 'อนุมัติ', approvedBy);
}
function rejectOT({ otId, approvedBy, note }) {
  return _updateOTStatus(otId, 'ปฏิเสธ', approvedBy, note);
}

function _updateOTStatus(otId, status, approvedBy, note) {
  const sheet   = getSheet(CONFIG.SHEET_OT);
  const data    = sheet.getDataRange().getValues();
  const headers = data[0];
  const rows    = data.slice(1);

  const rowIdx = rows.findIndex(r => r[headers.indexOf('OTID')] === otId);
  if (rowIdx === -1) return { success: false, message: 'ไม่พบคำขอ OT' };

  const now = toDateStr(getThailandNow());
  sheet.getRange(rowIdx + 2, headers.indexOf('Status')     + 1).setValue(status);
  sheet.getRange(rowIdx + 2, headers.indexOf('ApprovedBy') + 1).setValue(approvedBy || '');
  sheet.getRange(rowIdx + 2, headers.indexOf('ApprovedAt') + 1).setValue(now);
  if (note) sheet.getRange(rowIdx + 2, headers.indexOf('Note') + 1).setValue(note);

  Logger.log(`OT ${otId} → ${status} by ${approvedBy}`);
  return { success: true, status };
}

/* ═══════════════════════════════════════════════════════
   SETUP: registerUser
   POST { action:'registerUser', userId, name, dept, position,
          workplaceLat, workplaceLng, workplaceName }
═══════════════════════════════════════════════════════ */
function registerUser({ userId, name, dept, position, workplaceLat, workplaceLng, workplaceName }) {
  const sheet   = getSheet(CONFIG.SHEET_USERS);
  const users   = sheetToObjects(sheet);
  const exists  = users.find(u => u.LineUserID === userId);
  if (exists) return { success: false, message: 'พนักงานลงทะเบียนแล้ว' };

  const headers = ['LineUserID','Name','Department','Position','WorkplaceLat','WorkplaceLng','WorkplaceName','LeaveBalance','RegisteredAt'];
  appendRow(sheet, {
    LineUserID:    userId,
    Name:          name,
    Department:    dept || '',
    Position:      position || '',
    WorkplaceLat:  workplaceLat,
    WorkplaceLng:  workplaceLng,
    WorkplaceName: workplaceName || 'สำนักงาน',
    LeaveBalance:  CONFIG.ANNUAL_LEAVE_DAYS,
    RegisteredAt:  toDateStr(getThailandNow()),
  }, headers);

  Logger.log(`New user registered: ${userId} - ${name}`);
  return { success: true };
}

/* ═══════════════════════════════════════════════════════
   ONE-TIME SETUP: ลงทะเบียนผู้ใช้ (รันมือครั้งเดียว)
═══════════════════════════════════════════════════════ */
function registerMyUser() {
  return registerUser({
    userId:        'U351141375047c6a0e1020e1640049e30',
    name:          'ผู้ดูแลระบบ',
    dept:          'ฝ่าย IT',
    position:      'Admin',
    workplaceLat:  13.9377499,
    workplaceLng:  100.4737766,
    workplaceName: 'สำนักงาน JC',
  });
}

/* ═══════════════════════════════════════════════════════
   SETUP HELPER: สร้าง Sheet ครั้งแรก (รันมือครั้งเดียว)
   ไปที่ Apps Script → Run → setupSheets
═══════════════════════════════════════════════════════ */
function setupSheets() {
  const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);

  const sheets = {
    [CONFIG.SHEET_USERS]: [
      'LineUserID','Name','Department','Position',
      'WorkplaceLat','WorkplaceLng','WorkplaceName',
      'LeaveBalance','RegisteredAt'
    ],
    [CONFIG.SHEET_ATTENDANCE]: [
      'RecordID','UserID','Date','TimeIn','TimeOut',
      'LatIn','LngIn','LatOut','LngOut',
      'Status','LeaveType','LeaveApproved','Workplace','Note'
    ],
    [CONFIG.SHEET_LEAVE]: [
      'LeaveID','UserID','LeaveType','StartDate','EndDate',
      'Days','Reason','Status','SubmittedAt','ApprovedBy','ApprovedAt','Note'
    ],
    [CONFIG.SHEET_OT]: [
      'OTID','UserID','Date','StartTime','EndTime',
      'Purpose','Status','SubmittedAt','ApprovedBy','ApprovedAt','Note'
    ],
  };

  Object.entries(sheets).forEach(([name, headers]) => {
    let sheet = ss.getSheetByName(name);
    if (!sheet) {
      sheet = ss.insertSheet(name);
      Logger.log(`Created sheet: ${name}`);
    }
    // เขียน header ถ้าว่างอยู่
    if (sheet.getLastRow() === 0) {
      sheet.appendRow(headers);
      sheet.getRange(1, 1, 1, headers.length)
        .setFontWeight('bold')
        .setBackground('#1a3a5c')
        .setFontColor('#ffffff');
      sheet.setFrozenRows(1);
      Logger.log(`Headers set for: ${name}`);
    }
  });

  Logger.log('✅ Setup complete!');
  return 'Setup complete!';
}
