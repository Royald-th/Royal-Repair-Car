/**
 * ============================================================
 *  ระบบแจ้งซ่อมรถยนต์ — Google Apps Script Backend
 *  วาง Code นี้ใน Google Apps Script แล้ว Deploy เป็น Web App
 * ============================================================
 * 
 *  SETUP:
 *  1. สร้าง Google Spreadsheet ใหม่
 *  2. ไปที่ Extensions > Apps Script
 *  3. วาง Code นี้ทั้งหมด
 *  4. กรอกค่าใน CONFIG section ด้านล่าง
 *  5. Deploy > New Deployment > Web App
 *     - Execute as: Me
 *     - Who has access: Anyone
 *  6. คัดลอก URL มาใส่ใน CONFIG.GAS_URL ของ index.html
 * ============================================================
 */

/* ============================================================
   CONFIG — แก้ค่าเหล่านี้
   ============================================================ */
const CONFIG = {
  SPREADSHEET_ID: 'YOUR_SPREADSHEET_ID',
  LINE_CHANNEL_ACCESS_TOKEN: 'YOUR_LINE_CHANNEL_ACCESS_TOKEN',

  // ===== SECURITY =====
  // Secret token สำหรับกัน request นอกระบบ
  // ต้องตรงกับ API_SECRET ใน index.html (CONFIG.API_SECRET)
  // เปลี่ยนเป็นค่าสุ่มที่ยาวและซับซ้อน เช่น: 'xK9#mP2$qR7@nL4'
  API_SECRET: 'YOUR_API_SECRET_KEY',

  // LINE IDs
  LINE_SUPERVISOR_IDS:   ['YOUR_SUPERVISOR_LINE_ID'],
  LINE_MANAGER_IDS:      ['YOUR_MANAGER_LINE_ID'],
  LINE_ADMIN_USER_IDS:   ['YOUR_LINE_ADMIN_USER_ID'],

  DRIVE_FOLDER_REPAIR: 'YOUR_REPAIR_IMAGES_FOLDER_ID',
  DRIVE_FOLDER_BILL:   'YOUR_BILL_IMAGES_FOLDER_ID',
};

/* ============================================================
   SHEET NAMES
   ============================================================ */
const SHEETS = {
  JOBS: 'Jobs',
  USERS: 'Users',
  VEHICLES: 'Vehicles',
  STATUS_LOG: 'StatusLog',   // ประวัติการเปลี่ยนสถานะ
};

/* ============================================================
   HEADERS
   ============================================================ */
const JOB_HEADERS     = ['jobId','lineUid','userName','plate','mileage','detail','estimate','location','imageUrl','viewUrl','status','note','managerNote','actualCost','billUrl','billViewUrl','createdAt','updatedAt'];
// col map: A=jobId B=lineUid C=userName D=plate E=mileage F=detail G=estimate H=location
//          I=imageUrl J=viewUrl K=status L=note M=managerNote N=actualCost
//          O=billUrl P=billViewUrl Q=createdAt R=updatedAt
const USER_HEADERS    = ['lineUid','name','dept','role','avatar','createdAt'];
const VEHICLE_HEADERS  = ['plate','model','createdAt'];
const STATUS_LOG_HEADERS = ['logId','jobId','plate','userName','oldStatus','newStatus','note','actualCost','changedBy','changedAt'];

/* ============================================================
   MAIN ENTRY POINT
   ============================================================ */

/* ============================================================
   SECURITY
   ============================================================ */

// ตรวจ API Secret — return error object ถ้าไม่ผ่าน, null ถ้าผ่าน
function verifySecret(data) {
  if (!CONFIG.API_SECRET || CONFIG.API_SECRET === 'YOUR_API_SECRET_KEY') return null; // ยังไม่ได้ตั้งค่า skip
  if (data._secret !== CONFIG.API_SECRET) {
    Logger.log('[SECURITY] Invalid secret from action=' + data.action + ' lineUid=' + data.lineUid);
    return { status: 'error', message: 'Unauthorized' };
  }
  return null;
}

// ตรวจ lineUid ว่าอยู่ใน Sheet Users — ป้องกันคนนอก call API
function verifyUser(lineUid, requireRole) {
  if (!lineUid) return { status: 'error', message: 'Unauthorized: no lineUid' };
  const sheet = getSheet(SHEETS.USERS);
  const users = sheetToObjects(sheet, USER_HEADERS);
  const user  = users.find(u => u.lineUid === lineUid);
  if (!user) return { status: 'error', message: 'Unauthorized: user not found' };
  if (requireRole) {
    const hierarchy = ['user','supervisor','manager','admin'];
    const userLevel = hierarchy.indexOf(user.role);
    const reqLevel  = hierarchy.indexOf(requireRole);
    if (userLevel < reqLevel) return { status: 'error', message: 'Forbidden: insufficient role' };
  }
  return null; // ผ่าน
}

// ฟังก์ชันกลางสำหรับ route action
function routeAction(data) {
  const action = data.action;
  switch (action) {
    case 'getUser':           return getUser(data);
    case 'registerUser':      return registerUser(data);
    case 'updateUserAvatar':  return updateUserAvatar(data);
    case 'getJobs':           return getJobs(data);
    case 'createJob':         return createJob(data);
    case 'updateJob':         return updateJob(data);
    case 'updateStatus':      return updateStatus(data);
    case 'getUsers':          return getUsers();
    case 'addUser':           return addUser(data);
    case 'updateUser':        return updateUser(data);
    case 'deleteUser':        return deleteUser(data);
    case 'getVehicles':       return getVehicles();
    case 'addVehicle':        return addVehicle(data);
    case 'deleteVehicle':     return deleteVehicle(data);
    case 'deleteJob':         return deleteJob(data);
    case 'uploadImage':       return uploadImage(data);
    case 'ping':          return { status: 'ok', pong: true };
    case 'notifyNewJob':      notifySupervisors(data, data.jobId, data.now || new Date().toISOString()); return { status: 'ok' };
    case 'getStatusLog':      return getStatusLog(data);
    case 'getVehicleHistory': return getVehicleHistory(data);
    case 'getYearlyReport':   return getYearlyReport(data);
    case 'exportJobs':        return exportJobs(data);
    case 'ping':            return { status: 'ok', message: 'pong' };
    default: return { status: 'error', message: 'Unknown action: ' + action };
  }
}

function outputJson(result) {
  return ContentService
    .createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}

// GET — รับ payload เป็น query string (?payload=...)
// วิธีนี้หลีกเลี่ยง CORS preflight ได้ทั้งหมด
function doGet(e) {
  try {
    const raw  = e.parameter.payload;
    if (!raw) return outputJson({ status: 'ok', message: 'Car Repair API Running' });
    const data = JSON.parse(decodeURIComponent(raw));
    // ตรวจ Secret Token
    const check = verifySecret(data);
    if (check) return outputJson(check);
    return outputJson(routeAction(data));
  } catch (err) {
    Logger.log('doGet error: ' + err);
    return outputJson({ status: 'error', message: err.message });
  }
}

// POST — รองรับทุกรูปแบบที่ browser ส่งมา
function doPost(e) {
  try {
    let data;
    let source = '';

    // 1) e.parameter.payload (form-urlencoded ที่ GAS parse ให้อัตโนมัติ)
    if (e.parameter && e.parameter.payload) {
      source = 'e.parameter.payload';
      data = JSON.parse(e.parameter.payload);
    }
    // 2) e.postData.contents (raw body)
    else if (e.postData && e.postData.contents) {
      const raw = e.postData.contents;
      source = 'e.postData.contents (type=' + e.postData.type + ')';
      if (raw.indexOf('payload=') === 0) {
        // urlencoded: payload=...
        data = JSON.parse(decodeURIComponent(raw.slice(8)));
      } else if (raw.charAt(0) === '{') {
        // raw JSON
        data = JSON.parse(raw);
      } else {
        // พยายาม decode ทั้ง string
        data = JSON.parse(decodeURIComponent(raw));
      }
    } else {
      Logger.log('doPost: No payload found | parameters=' + JSON.stringify(e.parameter));
      return outputJson({ status: 'error', message: 'No payload received' });
    }

    Logger.log('doPost OK | source=' + source + ' | action=' + data.action
      + ' | keys=' + Object.keys(data).join(','));
    // ตรวจ Secret Token
    const check = verifySecret(data);
    if (check) return outputJson(check);
    return outputJson(routeAction(data));

  } catch (err) {
    Logger.log('doPost ERROR: ' + err
      + ' | postData type=' + (e.postData ? e.postData.type : 'none')
      + ' | raw=' + (e.postData ? e.postData.contents.substring(0, 300) : 'none'));
    return outputJson({ status: 'error', message: 'doPost parse error: ' + err.message });
  }
}


/* ============================================================
   SPREADSHEET HELPERS
   ============================================================ */
function getSheet(name) {
  const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  let sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    // Add headers
    const headers = name === SHEETS.JOBS ? JOB_HEADERS : name === SHEETS.USERS ? USER_HEADERS : VEHICLE_HEADERS;
    sheet.appendRow(headers);
    sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold').setBackground('#2E7D32').setFontColor('#ffffff');
  }
  return sheet;
}

function sheetToObjects(sheet, headers) {
  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return [];
  return data.slice(1)
    .filter(row => row.some(cell => cell !== null && cell !== undefined && String(cell).trim() !== ''))
    .map(row => {
      const obj = {};
      headers.forEach((h, i) => { obj[h] = row[i] !== undefined ? String(row[i]) : ''; });
      return obj;
    });
}

function findRowIndex(sheet, colIndex, value) {
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][colIndex]) === String(value)) return i + 1; // 1-based
  }
  return -1;
}

/* ============================================================
   JOB ID GENERATOR
   JOB-YYYYMMDD-XXX (reset monthly)
   ============================================================ */
function generateJobId() {
  const now     = new Date();
  const yyyy    = now.getFullYear();
  const mm      = String(now.getMonth() + 1).padStart(2, '0');
  const dd      = String(now.getDate()).padStart(2, '0');
  const prefix  = `JOB-${yyyy}${mm}${dd}-`;

  const sheet   = getSheet(SHEETS.JOBS);
  const data    = sheet.getDataRange().getValues();
  
  // filter jobs this month
  const thisMonth = `JOB-${yyyy}${mm}`;
  const monthJobs = data.slice(1).filter(row => String(row[0]).startsWith(thisMonth));
  const nextNum   = monthJobs.length + 1;
  return `${prefix}${String(nextNum).padStart(3, '0')}`;
}

/* ============================================================
   USER FUNCTIONS
   ============================================================ */
function getUser(data) {
  if (!data.lineUid) return { status: 'error', message: 'lineUid required' };
  const sheet = getSheet(SHEETS.USERS);
  const users = sheetToObjects(sheet, USER_HEADERS);
  const user  = users.find(u => u.lineUid === data.lineUid);
  if (user) return { status: 'ok', user };
  return { status: 'error', message: 'User not found' };
}

function getUsers() {
  const sheet = getSheet(SHEETS.USERS);
  const users = sheetToObjects(sheet, USER_HEADERS);
  return { status: 'ok', users };
}

/**
 * registerUser — เรียกจาก LIFF เมื่อ Login ครั้งแรก
 * ถ้ายังไม่มีใน Sheet → เพิ่มใหม่ role = user
 * ถ้ามีแล้ว → คืนข้อมูลเดิม (ป้องกัน race condition)
 */
function registerUser(data) {
  const sheet = getSheet(SHEETS.USERS);
  const users = sheetToObjects(sheet, USER_HEADERS);
  const existing = users.find(u => u.lineUid === data.lineUid);
  if (existing) {
    // มีแล้ว ส่งคืน user เดิม
    return { status: 'ok', user: existing, isNew: false };
  }
  // เพิ่มใหม่ โดยกำหนด role = user เสมอ
  sheet.appendRow([
    data.lineUid,
    data.name   || 'ผู้ใช้งาน',
    data.dept   || '',
    'user',                        // role ตายตัว
    data.avatar || '',
    new Date().toISOString(),
  ]);
  Logger.log('New user registered: ' + data.lineUid + ' / ' + data.name);
  return { status: 'ok', isNew: true };
}

/**
 * updateUserAvatar — อัปเดตรูปโปรไฟล์และชื่อ LINE ล่าสุด
 * (เรียกแบบ fire-and-forget ทุกครั้งที่ login)
 */
function updateUserAvatar(data) {
  const sheet = getSheet(SHEETS.USERS);
  const idx   = findRowIndex(sheet, 0, data.lineUid);
  if (idx < 0) return { status: 'error', message: 'ไม่พบผู้ใช้' };
  // อัปเดตเฉพาะ avatar (col 5) — ไม่แก้ชื่อที่ Admin อาจแก้เอง
  sheet.getRange(idx, 5).setValue(data.avatar || '');
  return { status: 'ok' };
}

function addUser(data) {
  const authErr = verifyUser(data.adminUid || data.lineUid, 'supervisor');
  if (authErr) return authErr;
  // supervisor ไม่สามารถสร้าง user ที่มี role = admin
  const isCallerAdmin = !verifyUser(data.adminUid || data.lineUid, 'admin');
  if (!isCallerAdmin && data.role === 'admin') {
    return { status: 'error', message: 'Forbidden: ไม่มีสิทธิ์กำหนด role เป็น admin' };
  }
  const sheet = getSheet(SHEETS.USERS);
  const users = sheetToObjects(sheet, USER_HEADERS);
  if (users.find(u => u.lineUid === data.lineUid)) {
    return { status: 'error', message: 'LINE UID นี้มีอยู่แล้ว' };
  }
  sheet.appendRow([data.lineUid, data.name, data.dept || '', data.role || 'user', '', new Date().toISOString()]);
  return { status: 'ok' };
}

function updateUser(data) {
  const authErr = verifyUser(data.adminUid || data.lineUid, 'supervisor');
  if (authErr) return authErr;
  const sheet = getSheet(SHEETS.USERS);
  const users = sheetToObjects(sheet, USER_HEADERS);
  const target = users.find(u => u.lineUid === data.lineUid);
  if (!target) return { status: 'error', message: 'ไม่พบผู้ใช้' };
  // supervisor ไม่สามารถแก้ไข admin หรือเปลี่ยน role เป็น admin
  const isCallerAdmin = !verifyUser(data.adminUid || data.lineUid, 'admin');
  if (!isCallerAdmin) {
    if (target.role === 'admin') return { status: 'error', message: 'Forbidden: ไม่มีสิทธิ์แก้ไข admin' };
    if (data.role === 'admin') return { status: 'error', message: 'Forbidden: ไม่มีสิทธิ์กำหนด role เป็น admin' };
  }
  const idx = findRowIndex(sheet, 0, data.lineUid);
  sheet.getRange(idx, 2).setValue(data.name);
  sheet.getRange(idx, 3).setValue(data.dept || '');
  sheet.getRange(idx, 4).setValue(data.role || 'user');
  return { status: 'ok' };
}

function deleteUser(data) {
  const authErr = verifyUser(data.adminUid || data.lineUid, 'supervisor');
  if (authErr) return authErr;
  const sheet = getSheet(SHEETS.USERS);
  const users = sheetToObjects(sheet, USER_HEADERS);
  const target = users.find(u => u.lineUid === data.lineUid);
  if (!target) return { status: 'error', message: 'ไม่พบผู้ใช้' };
  // supervisor ไม่สามารถลบ admin
  const isCallerAdmin = !verifyUser(data.adminUid || data.lineUid, 'admin');
  if (!isCallerAdmin && target.role === 'admin') {
    return { status: 'error', message: 'Forbidden: ไม่มีสิทธิ์ลบ admin' };
  }
  const idx = findRowIndex(sheet, 0, data.lineUid);
  sheet.deleteRow(idx);
  return { status: 'ok' };
}

/* ============================================================
   VEHICLE FUNCTIONS
   ============================================================ */
function getVehicles() {
  const sheet    = getSheet(SHEETS.VEHICLES);
  const vehicles = sheetToObjects(sheet, VEHICLE_HEADERS);
  return { status: 'ok', vehicles };
}

function addVehicle(data) {
  const authErr = verifyUser(data.adminUid || data.lineUid, 'admin');
  if (authErr) return authErr;
  const sheet    = getSheet(SHEETS.VEHICLES);
  const vehicles = sheetToObjects(sheet, VEHICLE_HEADERS);
  if (vehicles.find(v => v.plate === data.plate)) {
    return { status: 'error', message: 'ทะเบียนนี้มีอยู่แล้ว' };
  }
  sheet.appendRow([data.plate, data.model || '', new Date().toISOString()]);
  return { status: 'ok' };
}

function deleteVehicle(data) {
  const authErr = verifyUser(data.adminUid || data.lineUid, 'admin');
  if (authErr) return authErr;
  const sheet = getSheet(SHEETS.VEHICLES);
  const idx   = findRowIndex(sheet, 0, data.plate);
  if (idx < 0) return { status: 'error', message: 'ไม่พบทะเบียน' };
  sheet.deleteRow(idx);
  return { status: 'ok' };
}

/* ============================================================
   JOB FUNCTIONS
   ============================================================ */
function deleteJob(data) {
  const authErr = verifyUser(data.adminUid, 'supervisor');
  if (authErr) return authErr;
  // เฉพาะ Admin เท่านั้น
  if (!data.adminUid) return { status: 'error', message: 'ไม่มีสิทธิ์' };
  const sheet = getSheet(SHEETS.JOBS);
  const idx   = findRowIndex(sheet, 0, data.jobId);
  if (idx < 0) return { status: 'error', message: 'ไม่พบงานซ่อม: ' + data.jobId };
  sheet.deleteRow(idx);
  Logger.log('Job deleted: ' + data.jobId + ' by ' + data.adminUid);
  return { status: 'ok' };
}

function getJobs(data) {
  const sheet = getSheet(SHEETS.JOBS);
  let jobs    = sheetToObjects(sheet, JOB_HEADERS);

  // Non-admin: return only own jobs
  if (!data.isAdmin) {
    jobs = jobs.filter(j => j.lineUid === data.lineUid);
  }

  // Sort by createdAt DESC
  jobs.sort((a, b) => new Date(b.createdAt) - new Date(a.createdAt));
  return { status: 'ok', jobs };
}

function createJob(data) {
  if (!data.plate || !data.mileage || !data.detail || !data.lineUid) {
    return { status: 'error', message: 'ข้อมูลไม่ครบ: plate, mileage, detail, lineUid จำเป็น' };
  }
  const sheet  = getSheet(SHEETS.JOBS);
  const jobId  = generateJobId();
  const now    = new Date().toISOString();

  sheet.appendRow([
    jobId,
    data.lineUid,
    data.userName,
    data.plate,
    data.mileage,
    data.detail,
    data.estimate || '',
    data.location,
    data.imageUrl  || '',  // imageUrl (thumbnail)
    data.viewUrl   || '',  // viewUrl (Drive link)
    'รอดำเนินการ',
    data.note      || '',
    '',   // actualCost
    '',   // billUrl
    '',   // billViewUrl
    now,
    now,
  ]);

  // Rename รูปแจ้งซ่อมให้ตรงกับ jobId (กรณี upload ก่อนที่จะมี jobId)
  if (data.imageFileId) {
    try {
      const imgFile = DriveApp.getFileById(data.imageFileId);
      const ext = imgFile.getMimeType().split('/')[1].replace('jpeg','jpg');
      imgFile.setName(jobId + '_repair.' + ext);
    } catch(e) { Logger.log('rename repair image error: ' + e); }
  }

  return { status: 'ok', jobId };
}

function updateJob(data) {
  const sheet = getSheet(SHEETS.JOBS);
  const idx   = findRowIndex(sheet, 0, data.jobId);
  if (idx < 0) return { status: 'error', message: 'ไม่พบงานซ่อม' };

  const cols = { plate:4, mileage:5, detail:6, estimate:7, location:8, imageUrl:9, note:11 };
  // JOB_HEADERS index (1-based col):
  // 1=jobId,2=lineUid,3=userName,4=plate,5=mileage,6=detail,7=estimate,8=location,9=imageUrl,10=status,11=note,...,15=updatedAt
  // JOB_HEADERS cols (1-based):
  // 1=jobId,2=lineUid,3=userName,4=plate,5=mileage,6=detail,7=estimate,
  // 8=location,9=imageUrl,10=viewUrl,11=status,12=note,13=actualCost,
  // 14=billUrl,15=billViewUrl,16=createdAt,17=updatedAt
  sheet.getRange(idx, 4).setValue(data.plate);
  sheet.getRange(idx, 5).setValue(data.mileage);
  sheet.getRange(idx, 6).setValue(data.detail);
  sheet.getRange(idx, 7).setValue(data.estimate || '');
  sheet.getRange(idx, 8).setValue(data.location);
  if (data.imageUrl)  sheet.getRange(idx, 9).setValue(data.imageUrl);
  if (data.viewUrl)   sheet.getRange(idx, 10).setValue(data.viewUrl);
  sheet.getRange(idx, 12).setValue(data.note || '');
  sheet.getRange(idx, 17).setValue(new Date().toISOString());

  // If was "ส่งกลับแก้ไข", set back to "รอดำเนินการ"
  const currentStatus = sheet.getRange(idx, 10).getValue();
  if (currentStatus === 'ส่งกลับแก้ไข') {
    sheet.getRange(idx, 10).setValue('รอดำเนินการ');
  }

  return { status: 'ok' };
}

function updateStatus(data) {
  // supervisor และ manager ต่างก็อัปเดตสถานะได้
  const authErr = verifyUser(data.adminUid, 'supervisor');
  // manager มี role ต่ำกว่า supervisor ใน hierarchy — เช็คแยก
  let finalAuthErr = authErr;
  if (authErr) {
    const managerErr = verifyUser(data.adminUid, 'manager');
    if (!managerErr) finalAuthErr = null; // manager ผ่าน
  }
  if (finalAuthErr) return finalAuthErr;

  const sheet = getSheet(SHEETS.JOBS);
  const idx   = findRowIndex(sheet, 0, data.jobId);
  if (idx < 0) return { status: 'error', message: 'ไม่พบงานซ่อม' };

  // col map: K=11=status  L=12=note  M=13=managerNote  N=14=actualCost
  //          O=15=billUrl  P=16=billViewUrl  R=18=updatedAt
  const currentStatusOld = sheet.getRange(idx, 11).getValue();
  sheet.getRange(idx, 11).setValue(data.status);
  if (data.note)        sheet.getRange(idx, 12).setValue(data.note);
  if (data.managerNote) sheet.getRange(idx, 13).setValue(data.managerNote); // col M
  if (data.actualCost)  sheet.getRange(idx, 14).setValue(data.actualCost);
  if (data.billUrl)     sheet.getRange(idx, 15).setValue(data.billUrl);
  if (data.billViewUrl) sheet.getRange(idx, 16).setValue(data.billViewUrl);
  sheet.getRange(idx, 18).setValue(new Date().toISOString());

  const userLineUid = sheet.getRange(idx, 2).getValue();
  const userName    = sheet.getRange(idx, 3).getValue();
  const plate       = sheet.getRange(idx, 4).getValue();

  // ===== Workflow แจ้งเตือนตาม role =====
  if (data.status === 'รอการอนุมัติ') {
    // หัวหน้าส่งต่อ → แจ้ง ผู้บริหาร
    notifyManagers(data.jobId, plate, userName, data.note || '');
  } else if (['อนุมัติ','ไม่อนุมัติ','ส่งกลับแก้ไข'].includes(data.status)) {
    // ผู้บริหารตัดสินใจ → แจ้ง หัวหน้า
    notifySupervisorsDecision(data.jobId, plate, data.status, data.note || '');
  }
  // หมายเหตุ: ไม่แจ้งเตือนผู้ใช้ทุกกรณี

  // ===== บันทึก StatusLog =====
  try {
    const logSheet = getSheet(SHEETS.STATUS_LOG);
    // สร้าง header ถ้ายังไม่มี
    if (logSheet.getLastRow() === 0) {
      logSheet.appendRow(STATUS_LOG_HEADERS);
    }
    const logId = 'LOG-' + new Date().getTime();
    logSheet.appendRow([
      logId,
      data.jobId,
      plate,
      sheet.getRange(idx, 3).getValue(),  // userName
      currentStatusOld,
      data.status,
      data.note       || '',
      data.actualCost || '',
      data.adminUid   || '',
      new Date().toISOString(),
    ]);
    Logger.log('StatusLog บันทึกสำเร็จ: ' + logId);
  } catch(e) {
    Logger.log('StatusLog ERROR: ' + e);
  }

  return { status: 'ok' };
}

/* ============================================================
   IMAGE UPLOAD TO GOOGLE DRIVE
   folderType: 'repair' = รูปแจ้งซ่อม, 'bill' = บิล/ใบเสร็จ
   ============================================================ */
function uploadImage(data) {
  try {
    // Log ทุกอย่างเพื่อ debug
    Logger.log('uploadImage called | folderType=' + data.folderType
      + ' | filename=' + data.filename
      + ' | mimeType=' + data.mimeType
      + ' | base64 length=' + (data.base64 ? data.base64.length : 0)
      + ' | REPAIR_ID=' + CONFIG.DRIVE_FOLDER_REPAIR
      + ' | BILL_ID='   + CONFIG.DRIVE_FOLDER_BILL);

    // เลือกโฟลเดอร์ตามประเภท — ถ้า folderType ไม่มาให้ใช้ REPAIR เป็น default
    const isBill   = (String(data.folderType) === 'bill');
    const folderId = isBill ? CONFIG.DRIVE_FOLDER_BILL : CONFIG.DRIVE_FOLDER_REPAIR;

    Logger.log('Using folderId: ' + folderId + ' (isBill=' + isBill + ')');

    if (!folderId || folderId === '' || folderId.indexOf('YOUR_') === 0) {
      return { status: 'error', message: 'Folder ID ไม่ถูกต้อง: "' + folderId + '" กรุณาตรวจสอบ CONFIG' };
    }

    if (!data.base64 || data.base64.length === 0) {
      return { status: 'error', message: 'ไม่ได้รับข้อมูลรูปภาพ (base64 ว่างเปล่า)' };
    }

    const folder = DriveApp.getFolderById(folderId);

    // ตั้งชื่อไฟล์ตามเลขงาน ถ้ามี jobId
    const ext      = (data.mimeType || 'image/jpeg').split('/')[1].replace('jpeg','jpg');
    const fileType = isBill ? 'bill' : 'repair';
    const stamp    = Utilities.formatDate(new Date(), 'Asia/Bangkok', 'yyyyMMdd_HHmmss');
    const fname    = data.jobId
      ? (data.jobId + '_' + fileType + '.' + ext)
      : (fileType + '_' + stamp + '.' + ext);

    const blob = Utilities.newBlob(
      Utilities.base64Decode(data.base64.replace(/\s/g, '')),
      data.mimeType || 'image/jpeg',
      fname
    );
    const file = folder.createFile(blob);
    // หมายเหตุ: ไม่ต้อง setSharing ที่ไฟล์ทุกครั้ง
    // ให้ตั้ง "Anyone with the link can view" ที่ folder ใน Drive ครั้งเดียว
    // ไฟล์ใหม่จะ inherit permission จาก folder อัตโนมัติ → เร็วขึ้น ~1-3 วิ

    const fileId  = file.getId();
    const url     = 'https://drive.google.com/thumbnail?id=' + fileId + '&sz=w400';
    const viewUrl = 'https://drive.google.com/file/d/' + fileId + '/view';

    Logger.log('Upload SUCCESS: fileId=' + fileId);
    return { status: 'ok', url: url, viewUrl: viewUrl, fileId: fileId };

  } catch (e) {
    Logger.log('uploadImage ERROR: ' + e.toString());
    return { status: 'error', message: e.toString() };
  }
}

/* ============================================================
   EXPORT CSV DATA
   ============================================================ */
function exportJobs(data) {
  const sheet = getSheet(SHEETS.JOBS);
  let jobs    = sheetToObjects(sheet, JOB_HEADERS);

  // Admin เห็นทั้งหมด, User เห็นแค่ของตัวเอง
  if (!data.isAdmin) {
    jobs = jobs.filter(j => j.lineUid === data.lineUid);
  }

  // filter สถานะ (ถ้ามี)
  if (data.filterStatus) {
    jobs = jobs.filter(j => j.status === data.filterStatus);
  }

  // Sort newest first
  jobs.sort((a, b) => new Date(b.createdAt) - new Date(a.createdAt));

  // คืนเฉพาะ fields ที่ต้องการ (ไม่รวม lineUid, imageUrl, viewUrl ฯลฯ)
  const exportFields = ['jobId','userName','plate','mileage','detail','estimate','location','status','note','actualCost','createdAt','updatedAt'];
  const exportHeaders = ['เลขที่งาน','ผู้แจ้ง','ทะเบียน','ไมล์','อาการ','ประเมินราคา','สถานที่ซ่อม','สถานะ','หมายเหตุ','ค่าใช้จ่ายจริง','วันที่แจ้ง','อัปเดตล่าสุด'];

  const rows = jobs.map(j => exportFields.map(f => j[f] || ''));
  return { status: 'ok', headers: exportHeaders, rows: rows, total: rows.length };
}

/* ============================================================
   MIGRATE — รันครั้งเดียวเพื่ออัปเดต Sheet Jobs ให้มี header ใหม่
   ใช้เมื่อ Sheet เดิมสร้างไว้แล้วด้วย headers ชุดเก่า
   ============================================================ */
function migrateJobsSheet() {
  const ss    = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  const sheet = ss.getSheetByName(SHEETS.JOBS);
  if (!sheet) { Logger.log('ไม่พบ Sheet Jobs'); return; }

  const currentHeaders = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map(String);
  Logger.log('Headers ปัจจุบัน: ' + JSON.stringify(currentHeaders));
  Logger.log('Headers เป้าหมาย: ' + JSON.stringify(JOB_HEADERS));

  // ถ้า header ตรงกันแล้ว ไม่ต้องทำอะไร
  if (JSON.stringify(currentHeaders) === JSON.stringify(JOB_HEADERS)) {
    Logger.log('Headers ถูกต้องแล้ว ไม่ต้องแก้ไข');
    return;
  }

  // สร้าง map ของ header เก่า → index
  const oldMap = {};
  currentHeaders.forEach((h, i) => { oldMap[h] = i; });

  // อ่านข้อมูลทั้งหมด
  const allData = sheet.getDataRange().getValues();
  const dataRows = allData.slice(1); // ไม่รวม header row

  // สร้างข้อมูลใหม่ตาม JOB_HEADERS
  const newRows = dataRows.map(row => {
    return JOB_HEADERS.map(h => {
      const oldIdx = oldMap[h];
      return oldIdx !== undefined ? row[oldIdx] : '';
    });
  });

  // Clear sheet และเขียนใหม่
  sheet.clearContents();
  sheet.getRange(1, 1, 1, JOB_HEADERS.length).setValues([JOB_HEADERS]);
  sheet.getRange(1, 1, 1, JOB_HEADERS.length)
       .setFontWeight('bold').setBackground('#2E7D32').setFontColor('#ffffff');

  if (newRows.length > 0) {
    sheet.getRange(2, 1, newRows.length, JOB_HEADERS.length).setValues(newRows);
  }

  Logger.log('Migration สำเร็จ! ' + newRows.length + ' rows, ' + JOB_HEADERS.length + ' columns');
  Logger.log('Headers ใหม่: ' + JSON.stringify(JOB_HEADERS));
}

/* สร้างโฟลเดอร์ Drive อัตโนมัติ (รันครั้งเดียว) */
function createDriveFolders() {
  const root        = DriveApp.getRootFolder();
  const mainFolder  = root.createFolder('ระบบแจ้งซ่อมรถยนต์');
  const repairFolder = mainFolder.createFolder('รูปภาพแจ้งซ่อม');
  const billFolder   = mainFolder.createFolder('บิล-ใบเสร็จ');
  Logger.log('REPAIR folder ID: ' + repairFolder.getId());
  Logger.log('BILL   folder ID: ' + billFolder.getId());
  Logger.log('คัดลอก ID เหล่านี้ไปใส่ใน CONFIG');
}

/* ============================================================
   LINE NOTIFICATION
   ============================================================ */
// ส่ง message ไปหากลุ่ม LINE IDs
function notifyGroup(ids, msg) {
  if (!Array.isArray(ids)) ids = [ids];
  ids.forEach(uid => {
    if (uid && !uid.startsWith('YOUR_')) sendLineMessage(uid, msg);
  });
}

// admin — ใช้สำหรับแจ้งเตือนระบบ (backup ฯลฯ)
function notifyAllAdmins(msg) {
  notifyGroup(CONFIG.LINE_ADMIN_USER_IDS, msg);
}

// หัวหน้า — แจ้งเมื่อมีงานแจ้งซ่อมใหม่
function notifySupervisors(data, jobId, now) {
  const msg = [
    '🚗 มีการแจ้งซ่อมรถยนต์ใหม่!',
    '─────────────────────',
    `📋 เลขที่งาน : ${jobId}`,
    `👤 ผู้แจ้ง   : ${data.userName}`,
    `🚙 ทะเบียน  : ${data.plate}`,
    `📏 ไมล์     : ${Number(data.mileage).toLocaleString()} กม.`,
    `🔧 อาการ    : ${data.detail}`,
    `💰 ประเมิน  : ${data.estimate ? Number(data.estimate).toLocaleString() + ' บาท' : '-'}`,
    `📍 สถานที่  : ${data.location || '-'}`,
    `🕐 เวลา     : ${new Date(now).toLocaleString('th-TH')}`,
    '─────────────────────',
    'กรุณาเข้าระบบเพื่อตรวจสอบและส่งต่อผู้บริหาร',
  ].join('\n');
  notifyGroup(CONFIG.LINE_SUPERVISOR_IDS, msg);
}

// ผู้บริหาร — แจ้งเมื่อหัวหน้าส่งต่อให้อนุมัติ
function notifySupervisorsDecision(jobId, plate, status, note) {
  const emoji = { 'อนุมัติ':'✅', 'ไม่อนุมัติ':'❌', 'ส่งกลับแก้ไข':'⚠️' }[status] || '📋';
  const msg = [
    `${emoji} ผู้บริหารได้ตัดสินใจ: ${status}`,
    '─────────────────────',
    `📋 เลขที่งาน : ${jobId}`,
    `🚙 ทะเบียน  : ${plate}`,
    `📌 ผลการอนุมัติ : ${status}`,
    note ? `📝 หมายเหตุ  : ${note}` : '',
    '─────────────────────',
    'กรุณาเข้าระบบเพื่อดำเนินการต่อ',
  ].filter(Boolean).join('\n');
  notifyGroup(CONFIG.LINE_SUPERVISOR_IDS, msg);
}

function notifyManagers(jobId, plate, userName, note) {
  const msg = [
    '📋 รออนุมัติ: ใบแจ้งซ่อมรถยนต์',
    '─────────────────────',
    `📋 เลขที่งาน : ${jobId}`,
    `🚙 ทะเบียน  : ${plate}`,
    `👤 ผู้แจ้ง   : ${userName}`,
    note ? `📝 หมายเหตุ  : ${note}` : '',
    '─────────────────────',
    'กรุณาเข้าระบบเพื่ออนุมัติหรือไม่อนุมัติ',
  ].filter(Boolean).join('\n');
  notifyGroup(CONFIG.LINE_MANAGER_IDS, msg);
}

function notifyUserRevise(lineUid, jobId, note) {
  const msg = [
    '⚠️ แจ้งเตือน: ใบแจ้งซ่อมถูกส่งกลับแก้ไข',
    '─────────────────────',
    `📋 เลขที่งาน : ${jobId}`,
    `📝 หมายเหตุ  : ${note}`,
    '─────────────────────',
    'กรุณาเข้าระบบเพื่อแก้ไขและส่งใหม่',
  ].join('\n');
  sendLineMessage(lineUid, msg);
}

function notifyUserStatusChange(lineUid, jobId, plate, status, note, actualCost) {
  const emoji = {
    'อนุมัติ':        '✅',
    'ไม่อนุมัติ':    '❌',
    'กำลังซ่อม':     '🔧',
    'เสร็จสิ้น':     '🎉',
    'ส่งกลับแก้ไข': '⚠️',
    'รอดำเนินการ':   '⏳',
  }[status] || '📋';

  const lines = [
    `${emoji} อัปเดตสถานะงานซ่อมของคุณ`,
    '─────────────────────',
    `📋 เลขที่งาน : ${jobId}`,
    `🚙 ทะเบียน   : ${plate}`,
    `📌 สถานะใหม่ : ${status}`,
  ];
  if (note)       lines.push(`📝 หมายเหตุ  : ${note}`);
  if (actualCost) lines.push(`💰 ค่าใช้จ่ายจริง : ${Number(actualCost).toLocaleString()} บาท`);
  lines.push('─────────────────────');
  if (status === 'เสร็จสิ้น')     lines.push('งานซ่อมเสร็จสมบูรณ์แล้ว ขอบคุณครับ/ค่ะ');
  else if (status === 'ไม่อนุมัติ') lines.push('กรุณาติดต่อผู้ดูแลระบบเพื่อสอบถามเพิ่มเติม');
  else                              lines.push('กรุณาเข้าระบบเพื่อติดตามสถานะ');

  sendLineMessage(lineUid, lines.join('\n'));
}

function sendLineMessage(userId, message) {
  if (!CONFIG.LINE_CHANNEL_ACCESS_TOKEN || CONFIG.LINE_CHANNEL_ACCESS_TOKEN === 'YOUR_LINE_CHANNEL_ACCESS_TOKEN') {
    Logger.log('[LINE] Token not set, skip sending.');
    return;
  }
  try {
    const res = UrlFetchApp.fetch('https://api.line.me/v2/bot/message/push', {
      method: 'post',
      headers: {
        'Content-Type': 'application/json',
        'Authorization': `Bearer ${CONFIG.LINE_CHANNEL_ACCESS_TOKEN}`,
      },
      payload: JSON.stringify({
        to: userId,
        messages: [{ type: 'text', text: message }],
      }),
      muteHttpExceptions: true,
    });
    const code = res.getResponseCode();
    if (code !== 200) {
      Logger.log('[LINE] ERROR to=' + userId + ' status=' + code + ' body=' + res.getContentText());
    } else {
      Logger.log('[LINE] OK to=' + userId);
    }
  } catch (e) {
    Logger.log('[LINE] exception: ' + e.message);
  }
}


/* ============================================================
   ข้อ 5 — ประวัติการเปลี่ยนสถานะ
   ============================================================ */
function getStatusLog(data) {
  const sheet = getSheet(SHEETS.STATUS_LOG);
  let logs = sheetToObjects(sheet, STATUS_LOG_HEADERS);

  if (data.jobId) logs = logs.filter(l => l.jobId === data.jobId);
  logs.sort((a, b) => new Date(b.changedAt) - new Date(a.changedAt));
  return { status: 'ok', logs };
}

/* ============================================================
   ข้อ 6 — ประวัติซ่อมตามทะเบียน
   ============================================================ */
function getVehicleHistory(data) {
  if (!data.plate) return { status: 'error', message: 'กรุณาระบุทะเบียน' };

  const jobSheet = getSheet(SHEETS.JOBS);
  const jobs = sheetToObjects(jobSheet, JOB_HEADERS)
    .filter(j => j.plate === data.plate)
    .sort((a, b) => new Date(b.createdAt) - new Date(a.createdAt));

  const totalCost = jobs.reduce((s, j) => s + (parseFloat(j.actualCost) || 0), 0);
  const doneCount = jobs.filter(j => j.status === 'เสร็จสิ้น').length;

  return { status: 'ok', plate: data.plate, jobs, totalCost, doneCount };
}

/* ============================================================
   ข้อ 7 — รายงานรายปี
   ============================================================ */
function getYearlyReport(data) {
  const year = parseInt(data.year) || new Date().getFullYear();
  const jobSheet = getSheet(SHEETS.JOBS);
  const allJobs = sheetToObjects(jobSheet, JOB_HEADERS);

  // กรองเฉพาะปีที่เลือก
  const jobs = allJobs.filter(j => {
    if (!j.createdAt) return false;
    return new Date(j.createdAt).getFullYear() === year;
  });

  // สรุปรายเดือน
  const monthly = Array.from({length: 12}, (_, i) => ({
    month: i + 1,
    total: 0, done: 0, cost: 0, approved: 0, rejected: 0
  }));
  jobs.forEach(j => {
    const m = new Date(j.createdAt).getMonth(); // 0-based
    monthly[m].total++;
    if (j.status === 'เสร็จสิ้น') {
      monthly[m].done++;
      monthly[m].cost += parseFloat(j.actualCost) || 0;
    }
    if (j.status === 'อนุมัติ')    monthly[m].approved++;
    if (j.status === 'ไม่อนุมัติ') monthly[m].rejected++;
  });

  // รถที่ซ่อมบ่อย
  const plateCounts = {};
  const plateCosts  = {};
  jobs.forEach(j => {
    plateCounts[j.plate] = (plateCounts[j.plate] || 0) + 1;
    plateCosts[j.plate]  = (plateCosts[j.plate]  || 0) + (parseFloat(j.actualCost) || 0);
  });
  const topPlates = Object.entries(plateCounts)
    .map(([plate, count]) => ({ plate, count, cost: plateCosts[plate] || 0 }))
    .sort((a, b) => b.count - a.count)
    .slice(0, 10);

  // สรุปภาพรวม
  const totalJobs      = jobs.length;
  const totalDone      = jobs.filter(j => j.status === 'เสร็จสิ้น').length;
  const totalCost      = jobs.reduce((s, j) => s + (parseFloat(j.actualCost) || 0), 0);
  const totalApproved  = jobs.filter(j => j.status === 'อนุมัติ').length;
  const totalRejected  = jobs.filter(j => j.status === 'ไม่อนุมัติ').length;
  const approvedBudget = jobs.filter(j => j.status === 'อนุมัติ').reduce((s,j) => s+(parseFloat(j.estimate)||0), 0);

  return { status: 'ok', year, totalJobs, totalDone, totalCost, totalApproved, totalRejected, approvedBudget, monthly, topPlates };
}

/* ============================================================
   ข้อ 1 — Backup อัตโนมัติ (ตั้ง Time Trigger รันฟังก์ชันนี้ทุกวัน)
   วิธีตั้ง: Apps Script → Triggers → Add Trigger
     - Function: dailyBackup
     - Event: Time-driven → Day timer → เลือกเวลา
   ============================================================ */
function dailyBackup() {
  try {
    const ss        = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    const backupName = 'Backup_' + Utilities.formatDate(new Date(), 'Asia/Bangkok', 'yyyy-MM-dd');
    const backupFolder = getOrCreateBackupFolder();

    // Export เป็น Excel (.xlsx)
    const url  = 'https://docs.google.com/spreadsheets/d/' + CONFIG.SPREADSHEET_ID + '/export?format=xlsx';
    const blob = UrlFetchApp.fetch(url, {
      headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() },
      muteHttpExceptions: true,
    }).getBlob().setName(backupName + '.xlsx');

    backupFolder.createFile(blob);

    // ลบ backup เก่าเกิน 30 วัน
    const files = backupFolder.getFiles();
    const cutoff = new Date();
    cutoff.setDate(cutoff.getDate() - 30);
    while (files.hasNext()) {
      const f = files.next();
      if (f.getDateCreated() < cutoff) f.setTrashed(true);
    }

    Logger.log('Backup สำเร็จ: ' + backupName);
    notifyAllAdmins('✅ Backup ข้อมูลสำเร็จ\n📁 ' + backupName + '.xlsx\n🗑 ลบ backup เก่ากว่า 30 วันแล้ว');
  } catch(e) {
    Logger.log('Backup ERROR: ' + e);
    notifyAllAdmins('❌ Backup ล้มเหลว: ' + e.message);
  }
}

function getOrCreateBackupFolder() {
  const root    = DriveApp.getRootFolder();
  const folders = root.getFoldersByName('ระบบแจ้งซ่อม_Backup');
  if (folders.hasNext()) return folders.next();
  return root.createFolder('ระบบแจ้งซ่อม_Backup');
}

/* ============================================================
   SETUP FUNCTION — รันครั้งเดียวเพื่อสร้าง Sheets
   ============================================================ */
function setupSheets() {
  getSheet(SHEETS.JOBS);
  getSheet(SHEETS.USERS);
  getSheet(SHEETS.VEHICLES);
  getSheet(SHEETS.STATUS_LOG);

  // เพิ่ม Admin ตัวอย่าง
  const userSheet = getSheet(SHEETS.USERS);
  const users = sheetToObjects(userSheet, USER_HEADERS);
  if (!users.length) {
    userSheet.appendRow(['U_ADMIN_REPLACE_ME', 'ผู้ดูแลระบบ', 'ฝ่ายบริหาร', 'admin', '', new Date().toISOString()]);
    Logger.log('Added default admin. Please update lineUid with actual LINE User ID.');
  }

  Logger.log('Setup complete!');
}
