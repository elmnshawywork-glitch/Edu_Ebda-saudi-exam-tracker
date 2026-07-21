/**
 * =========================================================================
 *  نظام رصد بيانات وملاحظات الفحص السعودي - EBDA EDU
 *  Google Apps Script Backend  —  v11 (نسخة سريعة)
 * =========================================================================
 *  أهم التحسينات في هذه النسخة:
 *   1) فتح ملف الشيت مرة واحدة فقط لكل طلب (بدل عشرات المرات)
 *   2) التحقق من بنية الشيتات مرة واحدة فقط وليس مع كل قراءة
 *   3) تخزين مؤقت (Cache) للبيانات المرجعية وإحصائيات الداشبورد
 *   4) دالة bootstrap واحدة ترجّع كل ما تحتاجه الواجهة في طلب واحد
 * =========================================================================
 */

const TARGET_FOLDER_ID = '1o3fQe1zMmpLmGsvaOYpndlGGJXkuCgqk';
const DB_FILE_NAME = 'EBDA_Saudi_Exam_Tracker_DB';
const SESSION_DURATION = 6 * 60 * 60;

// غيّر هذا الرقم لو أضفت أعمدة جديدة، ليعيد النظام فحص بنية الشيتات مرة واحدة
const SCHEMA_VERSION = 'v11';

const SHEETS = {
  USERS: 'Users',
  PROFESSIONS: 'Professions',
  CENTERS: 'Centers',
  EVALUATORS: 'Evaluators',
  CANDIDATES: 'Candidates'
};

const RESULT_VALUES = ['ناجح', 'راسب', 'غائب'];

const SHEET_HEADERS = {
  Users: ['id', 'username', 'passwordHash', 'displayName', 'role'],
  Professions: ['id', 'name', 'notes', 'createdAt'],
  Centers: ['id', 'name', 'location', 'notes', 'createdAt'],
  Evaluators: ['id', 'name', 'professionIds', 'professionNames', 'notes', 'createdAt'],
  Candidates: [
    'id', 'date', 'centerId', 'centerName', 'professionId', 'professionName',
    'candidateName', 'idNumber', 'phone', 'result', 'notes', 'createdBy', 'createdAt',
    'evaluatorId', 'evaluatorName'
  ]
};

/* ================================================================== *
 *  1) ذاكرة مؤقتة داخل نفس الطلب  —  أكبر مصدر للبطء كان هنا
 * ================================================================== */
var __SS = null;          // ملف الشيت المفتوح
var __SHEETS = {};        // الشيتات المفتوحة
var __SCHEMA_OK = false;  // هل تم التحقق من البنية في هذا الطلب

function getOrCreateDatabase_() {
  if (__SS) return __SS;

  const props = PropertiesService.getScriptProperties();
  let fileId = props.getProperty('DB_FILE_ID');

  if (fileId) {
    try {
      __SS = SpreadsheetApp.openById(fileId);
      return __SS;
    } catch (e) {
      // الملف غير موجود، سيتم إنشاء ملف جديد بالأسفل
    }
  }

  const ss = SpreadsheetApp.create(DB_FILE_NAME);
  fileId = ss.getId();
  props.setProperty('DB_FILE_ID', fileId);
  props.deleteProperty('SCHEMA_VERSION');

  try {
    const file = DriveApp.getFileById(fileId);
    DriveApp.getFolderById(TARGET_FOLDER_ID).addFile(file);
    DriveApp.getRootFolder().removeFile(file);
  } catch (e) {
    Logger.log('تعذر نقل الملف إلى المجلد المحدد: ' + e);
  }

  __SS = ss;
  return ss;
}

// يتحقق من بنية كل الشيتات مرة واحدة فقط (ويحفظ ذلك في ScriptProperties)
function ensureSchemaOnce_(ss) {
  if (__SCHEMA_OK) return;
  const props = PropertiesService.getScriptProperties();
  if (props.getProperty('SCHEMA_VERSION') === SCHEMA_VERSION) {
    __SCHEMA_OK = true;
    return;
  }
  Object.keys(SHEET_HEADERS).forEach(function (name) {
    ensureSheet_(ss, name, SHEET_HEADERS[name]);
  });
  props.setProperty('SCHEMA_VERSION', SCHEMA_VERSION);
  __SCHEMA_OK = true;
}

function getSheet_(name) {
  if (__SHEETS[name]) return __SHEETS[name];
  const ss = getOrCreateDatabase_();
  ensureSchemaOnce_(ss);
  let sheet = ss.getSheetByName(name);
  if (!sheet) sheet = ensureSheet_(ss, name, SHEET_HEADERS[name]);
  __SHEETS[name] = sheet;
  return sheet;
}

function ensureSheet_(ss, name, headers) {
  let sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    sheet.appendRow(headers);
    sheet.setFrozenRows(1);
    sheet.getRange(1, 1, 1, headers.length)
      .setFontWeight('bold').setBackground('#57B1C9').setFontColor('#FFFFFF');
  } else {
    ensureColumns_(sheet, headers);
  }
  const defaultSheet = ss.getSheetByName('Sheet1');
  if (defaultSheet && ss.getSheets().length > 1) ss.deleteSheet(defaultSheet);
  return sheet;
}

function ensureColumns_(sheet, headers) {
  const lastCol = sheet.getLastColumn();
  const existing = lastCol > 0 ? sheet.getRange(1, 1, 1, lastCol).getValues()[0] : [];
  headers.forEach(function (h) {
    if (existing.indexOf(h) === -1) {
      const newCol = sheet.getLastColumn() + 1;
      sheet.getRange(1, newCol).setValue(h)
        .setFontWeight('bold').setBackground('#57B1C9').setFontColor('#FFFFFF');
    }
  });
}

/* ================================================================== *
 *  2) تخزين مؤقت بين الطلبات (CacheService) — يُلغى تلقائيًا عند أي كتابة
 * ================================================================== */
const CACHE_KEYS = ['REF_DATA', 'DASH_STATS'];

function cacheGet_(key) {
  try {
    const v = CacheService.getScriptCache().get(key);
    return v ? JSON.parse(v) : null;
  } catch (e) { return null; }
}

function cachePut_(key, obj) {
  try {
    CacheService.getScriptCache().put(key, JSON.stringify(obj), 3600);
  } catch (e) { /* الحجم أكبر من المسموح — نتجاهل بهدوء */ }
}

function invalidateCache_() {
  try { CacheService.getScriptCache().removeAll(CACHE_KEYS); } catch (e) {}
}

/* ================================================================== *
 *  أدوات عامة
 * ================================================================== */
function toText_(v) {
  if (v === null || v === undefined) return '';
  if (v instanceof Date) return Utilities.formatDate(v, Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm');
  return String(v);
}

function formatDate_(value) {
  if (value instanceof Date) {
    return Utilities.formatDate(value, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  }
  return toText_(value);
}

function hashPassword_(password) {
  const digest = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, password, Utilities.Charset.UTF_8);
  return digest.map(function (b) {
    return (b < 0 ? b + 256 : b).toString(16).padStart(2, '0');
  }).join('');
}

/* ================================================================== *
 *  الإعداد الأولي
 * ================================================================== */
function setup() {
  const ss = getOrCreateDatabase_();
  Object.keys(SHEET_HEADERS).forEach(function (name) {
    ensureSheet_(ss, name, SHEET_HEADERS[name]);
  });
  PropertiesService.getScriptProperties().setProperty('SCHEMA_VERSION', SCHEMA_VERSION);

  const usersSheet = ss.getSheetByName(SHEETS.USERS);
  if (usersSheet.getLastRow() < 2) {
    usersSheet.appendRow([
      Utilities.getUuid(), 'admin', hashPassword_('admin@2026'), 'مدير النظام', 'admin'
    ]);
  }
  SpreadsheetApp.flush();
  invalidateCache_();
  return 'تم الإعداد. المستخدم: admin — كلمة المرور: admin@2026';
}

/**
 * شغّل هذه الدالة يدويًا مرة واحدة من محرر Apps Script (زر Run)
 * لضبط حساب admin على كلمة المرور admin@2026 حتى لو كان الحساب موجودًا من قبل.
 */
function forceAdminPassword() {
  const sheet = getSheet_(SHEETS.USERS);
  const data = sheet.getDataRange().getValues();
  const hash = hashPassword_('admin@2026');

  for (let i = 1; i < data.length; i++) {
    if (toText_(data[i][1]) === 'admin') {
      sheet.getRange(i + 1, 3).setValue(hash);
      sheet.getRange(i + 1, 5).setValue('admin');
      SpreadsheetApp.flush();
      return 'تم ضبط كلمة مرور admin على admin@2026';
    }
  }

  sheet.appendRow([Utilities.getUuid(), 'admin', hash, 'مدير النظام', 'admin']);
  SpreadsheetApp.flush();
  return 'تم إنشاء حساب admin بكلمة المرور admin@2026';
}

/* ================================================================== *
 *  واجهة الويب
 * ================================================================== */
function doGet(e) {
  // createHtmlOutputFromFile أسرع من createTemplateFromFile لأن الصفحة لا تحتوي قوالب
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('نظام رصد الفحص السعودي - EBDA EDU')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/* ================================================================== *
 *  الجلسات والصلاحيات
 * ================================================================== */
function createSession_(user) {
  const token = Utilities.getUuid();
  CacheService.getScriptCache().put('session_' + token, JSON.stringify({
    id: user.id, username: user.username, displayName: user.displayName, role: user.role
  }), SESSION_DURATION);
  return token;
}

function getSession_(token) {
  if (!token) return null;
  const raw = CacheService.getScriptCache().get('session_' + token);
  return raw ? JSON.parse(raw) : null;
}

function requireRole_(token, allowedRoles) {
  const session = getSession_(token);
  if (!session) throw new Error('انتهت صلاحية الجلسة، من فضلك سجّل الدخول مرة أخرى');
  if (allowedRoles && allowedRoles.length > 0 && allowedRoles.indexOf(session.role) === -1) {
    throw new Error('ليس لديك صلاحية القيام بهذا الإجراء');
  }
  return session;
}

/* ================================================================== *
 *  تسجيل الدخول — يرجّع كل بيانات البداية في نفس الطلب (توفير رحلتين)
 * ================================================================== */
function login(username, password) {
  const data = getSheet_(SHEETS.USERS).getDataRange().getValues();
  const hash = hashPassword_(password);

  for (let i = 1; i < data.length; i++) {
    if (toText_(data[i][1]) === username && toText_(data[i][2]) === hash) {
      const user = {
        id: toText_(data[i][0]),
        username: toText_(data[i][1]),
        displayName: toText_(data[i][3]),
        role: toText_(data[i][4])
      };
      const token = createSession_(user);
      return {
        success: true,
        token: token,
        username: user.username,
        displayName: user.displayName,
        role: user.role,
        reference: readReferenceData_(),
        dashboard: readDashboardStats_()
      };
    }
  }
  return { success: false, message: 'اسم المستخدم أو كلمة المرور غير صحيحة' };
}

// تحديث كامل لبيانات الواجهة في طلب واحد
function bootstrap(token) {
  requireRole_(token, []);
  return { reference: readReferenceData_(), dashboard: readDashboardStats_() };
}

/* ================================================================== *
 *  القراءة الداخلية للبيانات المرجعية (مع الكاش)
 * ================================================================== */
function readReferenceData_() {
  const cached = cacheGet_('REF_DATA');
  if (cached) return cached;

  const data = {
    professions: readProfessions_(),
    centers: readCenters_(),
    evaluators: readEvaluators_()
  };
  cachePut_('REF_DATA', data);
  return data;
}

function readProfessions_() {
  const data = getSheet_(SHEETS.PROFESSIONS).getDataRange().getValues();
  const rows = [];
  for (let i = 1; i < data.length; i++) {
    const id = toText_(data[i][0]);
    if (!id) continue;
    rows.push({ id: id, name: toText_(data[i][1]), notes: toText_(data[i][2]), createdAt: toText_(data[i][3]) });
  }
  return rows;
}

function readCenters_() {
  const data = getSheet_(SHEETS.CENTERS).getDataRange().getValues();
  const rows = [];
  for (let i = 1; i < data.length; i++) {
    const id = toText_(data[i][0]);
    if (!id) continue;
    rows.push({
      id: id, name: toText_(data[i][1]), location: toText_(data[i][2]),
      notes: toText_(data[i][3]), createdAt: toText_(data[i][4])
    });
  }
  return rows;
}

function readEvaluators_() {
  const data = getSheet_(SHEETS.EVALUATORS).getDataRange().getValues();
  const rows = [];
  for (let i = 1; i < data.length; i++) {
    const id = toText_(data[i][0]);
    if (!id) continue;
    const ids = toText_(data[i][2]);
    const names = toText_(data[i][3]);
    rows.push({
      id: id,
      name: toText_(data[i][1]),
      professionIds: ids ? ids.split(',').filter(Boolean) : [],
      professionNames: names ? names.split(',').filter(Boolean) : [],
      notes: toText_(data[i][4]),
      createdAt: toText_(data[i][5])
    });
  }
  return rows;
}

// دوال عامة للتوافق مع أي استدعاء قديم
function getAllReferenceData(token) { requireRole_(token, []); return readReferenceData_(); }
function getProfessions(token) { requireRole_(token, []); return readReferenceData_().professions; }
function getCenters(token) { requireRole_(token, []); return readReferenceData_().centers; }
function getEvaluators(token) { requireRole_(token, []); return readReferenceData_().evaluators; }

/* ================================================================== *
 *  المهن
 * ================================================================== */
function addProfession(token, name, notes) {
  requireRole_(token, ['admin', 'editor']);
  if (!name) throw new Error('اسم المهنة مطلوب');
  getSheet_(SHEETS.PROFESSIONS).appendRow([Utilities.getUuid(), name, notes || '', new Date()]);
  SpreadsheetApp.flush();
  invalidateCache_();
  return readReferenceData_();
}

function deleteProfession(token, id) {
  requireRole_(token, ['admin']);
  deleteRowById_(SHEETS.PROFESSIONS, id);
  SpreadsheetApp.flush();
  invalidateCache_();
  return readReferenceData_();
}

/* ================================================================== *
 *  مراكز التدريب
 * ================================================================== */
function addCenter(token, name, location, notes) {
  requireRole_(token, ['admin', 'editor']);
  if (!name) throw new Error('اسم المركز مطلوب');
  getSheet_(SHEETS.CENTERS).appendRow([Utilities.getUuid(), name, location || '', notes || '', new Date()]);
  SpreadsheetApp.flush();
  invalidateCache_();
  return readReferenceData_();
}

function deleteCenter(token, id) {
  requireRole_(token, ['admin']);
  deleteRowById_(SHEETS.CENTERS, id);
  SpreadsheetApp.flush();
  invalidateCache_();
  return readReferenceData_();
}

/* ================================================================== *
 *  المقيّمون
 * ================================================================== */
function addEvaluator(token, name, professionIds, notes) {
  requireRole_(token, ['admin', 'editor']);
  if (!name) throw new Error('اسم المقيّم مطلوب');

  professionIds = professionIds || [];
  // نقرأ المهن مرة واحدة من الكاش بدل البحث في الشيت لكل مهنة
  const allProfessions = readReferenceData_().professions;
  const nameById = {};
  allProfessions.forEach(function (p) { nameById[p.id] = p.name; });
  const professionNames = professionIds.map(function (pid) { return nameById[pid] || ''; }).filter(Boolean);

  getSheet_(SHEETS.EVALUATORS).appendRow([
    Utilities.getUuid(), name, professionIds.join(','), professionNames.join(','), notes || '', new Date()
  ]);
  SpreadsheetApp.flush();
  invalidateCache_();
  return readReferenceData_();
}

function deleteEvaluator(token, id) {
  requireRole_(token, ['admin']);
  deleteRowById_(SHEETS.EVALUATORS, id);
  SpreadsheetApp.flush();
  invalidateCache_();
  return readReferenceData_();
}

/* ================================================================== *
 *  أدوات الحذف والبحث
 * ================================================================== */
function deleteRowById_(sheetName, id) {
  const sheet = getSheet_(sheetName);
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (toText_(data[i][0]) === id) {
      sheet.deleteRow(i + 1);
      return { success: true };
    }
  }
  return { success: false, message: 'العنصر غير موجود' };
}

/* ================================================================== *
 *  المتقدمون للفحص
 * ================================================================== */
function addCandidate(token, c) {
  requireRole_(token, ['admin', 'editor']);
  if (RESULT_VALUES.indexOf(c.result) === -1) throw new Error('نتيجة الفحص غير صحيحة');

  // كل الأسماء من الكاش في عملية واحدة بدل 3 قراءات منفصلة من الشيت
  const ref = readReferenceData_();
  const centerName = nameFromList_(ref.centers, c.centerId);
  const professionName = nameFromList_(ref.professions, c.professionId);
  const evaluatorName = c.evaluatorId ? nameFromList_(ref.evaluators, c.evaluatorId) : '';

  const id = Utilities.getUuid();
  getSheet_(SHEETS.CANDIDATES).appendRow([
    id, c.date, c.centerId, centerName, c.professionId, professionName,
    c.candidateName || '', c.idNumber || '', c.phone || '', c.result,
    c.notes || '', c.createdBy || '', new Date(),
    c.evaluatorId || '', evaluatorName
  ]);
  SpreadsheetApp.flush();
  try { CacheService.getScriptCache().remove('DASH_STATS'); } catch (e) {}

  return { success: true, id: id, dashboard: readDashboardStats_() };
}

function nameFromList_(list, id) {
  if (!id) return '';
  for (let i = 0; i < list.length; i++) {
    if (list[i].id === id) return list[i].name;
  }
  return '';
}

function deleteCandidate(token, id) {
  requireRole_(token, ['admin']);
  const res = deleteRowById_(SHEETS.CANDIDATES, id);
  SpreadsheetApp.flush();
  invalidateCache_();
  return res;
}

function getAllCandidates_() {
  const data = getSheet_(SHEETS.CANDIDATES).getDataRange().getValues();
  const rows = [];
  for (let i = 1; i < data.length; i++) {
    const d = data[i];
    if (!toText_(d[0])) continue;
    rows.push({
      id: toText_(d[0]),
      date: formatDate_(d[1]),
      centerId: toText_(d[2]),
      centerName: toText_(d[3]),
      professionId: toText_(d[4]),
      professionName: toText_(d[5]),
      candidateName: toText_(d[6]),
      idNumber: toText_(d[7]),
      phone: toText_(d[8]),
      result: toText_(d[9]),
      notes: toText_(d[10]),
      createdBy: toText_(d[11]),
      createdAt: toText_(d[12]),
      evaluatorId: toText_(d[13]),
      evaluatorName: toText_(d[14])
    });
  }
  return rows;
}

function getCandidates(token, filters) {
  requireRole_(token, []);
  filters = filters || {};
  return getAllCandidates_().filter(function (r) {
    if (filters.date && r.date !== filters.date) return false;
    if (filters.centerId && r.centerId !== filters.centerId) return false;
    if (filters.professionId && r.professionId !== filters.professionId) return false;
    if (filters.result && r.result !== filters.result) return false;
    if (filters.evaluatorId && r.evaluatorId !== filters.evaluatorId) return false;
    return true;
  }).sort(function (a, b) { return a.date < b.date ? 1 : -1; });
}

/* ================================================================== *
 *  التقارير
 * ================================================================== */
function bump_(map, key, result) {
  key = key || 'غير محدد';
  if (!map[key]) map[key] = { name: key, totalCandidates: 0, passed: 0, failed: 0, absent: 0 };
  map[key].totalCandidates++;
  if (result === 'ناجح') map[key].passed++;
  else if (result === 'راسب') map[key].failed++;
  else if (result === 'غائب') map[key].absent++;
}

function getMonthlyReport(token, year, month) {
  requireRole_(token, []);
  const ym = year + '-' + String(month).padStart(2, '0');
  const monthRows = getAllCandidates_().filter(function (r) {
    return r.date && r.date.indexOf(ym) === 0;
  });

  const byProfession = {}, byCenter = {}, notes = [];
  const totals = { totalCandidates: monthRows.length, passed: 0, failed: 0, absent: 0 };

  monthRows.forEach(function (r) {
    if (r.result === 'ناجح') totals.passed++;
    else if (r.result === 'راسب') totals.failed++;
    else if (r.result === 'غائب') totals.absent++;
    bump_(byProfession, r.professionName, r.result);
    bump_(byCenter, r.centerName, r.result);
    if (r.notes && r.notes.trim() !== '') {
      notes.push({
        date: r.date, center: r.centerName, profession: r.professionName,
        candidateName: r.candidateName, note: r.notes
      });
    }
  });

  return {
    year: year, month: month, totals: totals,
    byProfession: Object.values(byProfession),
    byCenter: Object.values(byCenter),
    notes: notes.sort(function (a, b) { return a.date < b.date ? -1 : 1; }),
    dailyRows: monthRows.sort(function (a, b) { return a.date < b.date ? -1 : 1; })
  };
}

function getCustomReport(token, filters) {
  requireRole_(token, []);
  filters = filters || {};

  const rows = getAllCandidates_().filter(function (r) {
    if (filters.dateFrom && r.date < filters.dateFrom) return false;
    if (filters.dateTo && r.date > filters.dateTo) return false;
    if (filters.centerIds && filters.centerIds.length && filters.centerIds.indexOf(r.centerId) === -1) return false;
    if (filters.professionIds && filters.professionIds.length && filters.professionIds.indexOf(r.professionId) === -1) return false;
    if (filters.evaluatorIds && filters.evaluatorIds.length && filters.evaluatorIds.indexOf(r.evaluatorId) === -1) return false;
    if (filters.results && filters.results.length && filters.results.indexOf(r.result) === -1) return false;
    return true;
  });

  const byProfession = {}, byCenter = {}, byEvaluator = {}, byDate = {}, notes = [];
  const totals = { totalCandidates: rows.length, passed: 0, failed: 0, absent: 0 };

  rows.forEach(function (r) {
    if (r.result === 'ناجح') totals.passed++;
    else if (r.result === 'راسب') totals.failed++;
    else if (r.result === 'غائب') totals.absent++;

    bump_(byProfession, r.professionName, r.result);
    bump_(byCenter, r.centerName, r.result);
    bump_(byEvaluator, r.evaluatorName || 'غير محدد', r.result);
    bump_(byDate, r.date, r.result);

    if (r.notes && r.notes.trim() !== '') {
      notes.push({
        date: r.date, center: r.centerName, profession: r.professionName,
        candidateName: r.candidateName, note: r.notes
      });
    }
  });

  return {
    totals: totals,
    byProfession: Object.values(byProfession),
    byCenter: Object.values(byCenter),
    byEvaluator: Object.values(byEvaluator),
    byDate: Object.values(byDate).sort(function (a, b) { return a.name < b.name ? -1 : 1; }),
    notes: notes.sort(function (a, b) { return a.date < b.date ? -1 : 1; }),
    rows: rows.sort(function (a, b) { return a.date < b.date ? 1 : -1; })
  };
}

/* ================================================================== *
 *  سجل المتقدمين الكامل (تبويب "سجل المتقدمين")
 *  يدعم الفلترة بعدة معايير مع بحث نصي حر
 * ------------------------------------------------------------------ *
 *  filters:
 *   dateFrom, dateTo   : نطاق تاريخ (yyyy-MM-dd)
 *   centerIds[]        : مراكز — فارغة = الكل
 *   professionIds[]    : مهن — فارغة = الكل
 *   evaluatorIds[]     : مقيّمون — فارغة = الكل
 *   results[]          : نتائج — فارغة = الكل
 *   q                  : بحث في الاسم / الرقم القومي / الهاتف / الملاحظات
 * ================================================================== */
function getRegistry(token, filters) {
  requireRole_(token, []);
  filters = filters || {};
  const q = String(filters.q || '').trim().toLowerCase();

  const rows = getAllCandidates_().filter(function (r) {
    if (filters.dateFrom && r.date < filters.dateFrom) return false;
    if (filters.dateTo && r.date > filters.dateTo) return false;
    if (filters.centerIds && filters.centerIds.length && filters.centerIds.indexOf(r.centerId) === -1) return false;
    if (filters.professionIds && filters.professionIds.length && filters.professionIds.indexOf(r.professionId) === -1) return false;
    if (filters.evaluatorIds && filters.evaluatorIds.length && filters.evaluatorIds.indexOf(r.evaluatorId) === -1) return false;
    if (filters.results && filters.results.length && filters.results.indexOf(r.result) === -1) return false;
    if (q) {
      const hay = (r.candidateName + ' ' + r.idNumber + ' ' + r.phone + ' ' + r.notes + ' ' +
                   r.centerName + ' ' + r.professionName + ' ' + r.evaluatorName).toLowerCase();
      if (hay.indexOf(q) === -1) return false;
    }
    return true;
  }).sort(function (a, b) { return a.date < b.date ? 1 : (a.date > b.date ? -1 : 0); });

  const totals = { totalCandidates: rows.length, passed: 0, failed: 0, absent: 0 };
  rows.forEach(function (r) {
    if (r.result === 'ناجح') totals.passed++;
    else if (r.result === 'راسب') totals.failed++;
    else if (r.result === 'غائب') totals.absent++;
  });

  return { rows: rows, totals: totals };
}

/* ================================================================== *
 *  إحصائيات الداشبورد (مع الكاش)
 * ================================================================== */
function getDashboardStats(token) {
  requireRole_(token, []);
  return readDashboardStats_();
}

function readDashboardStats_() {
  const cached = cacheGet_('DASH_STATS');
  if (cached) return cached;

  const all = getAllCandidates_();
  const ref = readReferenceData_();

  const totals = { totalCandidates: all.length, passed: 0, failed: 0, absent: 0, sessions: 0 };
  const byProfession = {}, byCenter = {}, byMonth = {}, sessionKeys = {};

  all.forEach(function (r) {
    if (r.result === 'ناجح') totals.passed++;
    else if (r.result === 'راسب') totals.failed++;
    else if (r.result === 'غائب') totals.absent++;

    sessionKeys[r.date + '|' + r.centerId] = 1;

    if (!byProfession[r.professionName]) {
      byProfession[r.professionName] = { name: r.professionName, totalCandidates: 0, passed: 0, failed: 0 };
    }
    byProfession[r.professionName].totalCandidates++;
    if (r.result === 'ناجح') byProfession[r.professionName].passed++;
    else if (r.result === 'راسب') byProfession[r.professionName].failed++;

    if (!byCenter[r.centerName]) {
      byCenter[r.centerName] = { name: r.centerName, totalCandidates: 0, passed: 0, failed: 0 };
    }
    byCenter[r.centerName].totalCandidates++;
    if (r.result === 'ناجح') byCenter[r.centerName].passed++;
    else if (r.result === 'راسب') byCenter[r.centerName].failed++;

    const ym = r.date ? r.date.substring(0, 7) : 'غير محدد';
    if (!byMonth[ym]) byMonth[ym] = { month: ym, totalCandidates: 0, passed: 0, failed: 0 };
    byMonth[ym].totalCandidates++;
    if (r.result === 'ناجح') byMonth[ym].passed++;
    else if (r.result === 'راسب') byMonth[ym].failed++;
  });

  totals.sessions = Object.keys(sessionKeys).length;

  const stats = {
    totals: totals,
    passRate: totals.totalCandidates > 0
      ? Math.round((totals.passed / totals.totalCandidates) * 1000) / 10 : 0,
    byProfession: Object.values(byProfession),
    byCenter: Object.values(byCenter),
    byMonth: Object.values(byMonth).sort(function (a, b) { return a.month < b.month ? -1 : 1; }),
    centersCount: ref.centers.length,
    professionsCount: ref.professions.length,
    evaluatorsCount: ref.evaluators.length
  };

  cachePut_('DASH_STATS', stats);
  return stats;
}

/* ================================================================== *
 *  إدارة المستخدمين (Admin فقط)
 * ================================================================== */
function getUsers(token) {
  requireRole_(token, ['admin']);
  const data = getSheet_(SHEETS.USERS).getDataRange().getValues();
  const rows = [];
  for (let i = 1; i < data.length; i++) {
    const id = toText_(data[i][0]);
    if (!id) continue;
    rows.push({
      id: id, username: toText_(data[i][1]),
      displayName: toText_(data[i][3]), role: toText_(data[i][4])
    });
  }
  return rows;
}

function addUser(token, username, password, displayName, role) {
  requireRole_(token, ['admin']);
  if (!username || !password) throw new Error('اسم المستخدم وكلمة المرور مطلوبان');
  const sheet = getSheet_(SHEETS.USERS);
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (toText_(data[i][1]) === username) throw new Error('اسم المستخدم موجود بالفعل');
  }
  const validRole = ['admin', 'editor', 'viewer'].indexOf(role) !== -1 ? role : 'viewer';
  sheet.appendRow([Utilities.getUuid(), username, hashPassword_(password), displayName || username, validRole]);
  SpreadsheetApp.flush();
  return { success: true };
}

function updateUserRole(token, id, role) {
  requireRole_(token, ['admin']);
  const validRole = ['admin', 'editor', 'viewer'].indexOf(role) !== -1 ? role : 'viewer';
  const sheet = getSheet_(SHEETS.USERS);
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (toText_(data[i][0]) === id) {
      sheet.getRange(i + 1, 5).setValue(validRole);
      SpreadsheetApp.flush();
      return { success: true };
    }
  }
  return { success: false, message: 'المستخدم غير موجود' };
}

function resetUserPassword(token, id, newPassword) {
  requireRole_(token, ['admin']);
  if (!newPassword) throw new Error('كلمة المرور الجديدة مطلوبة');
  const sheet = getSheet_(SHEETS.USERS);
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (toText_(data[i][0]) === id) {
      sheet.getRange(i + 1, 3).setValue(hashPassword_(newPassword));
      SpreadsheetApp.flush();
      return { success: true };
    }
  }
  return { success: false, message: 'المستخدم غير موجود' };
}

function deleteUser(token, id) {
  const session = requireRole_(token, ['admin']);
  if (session.id === id) throw new Error('لا يمكنك حذف حسابك الحالي');
  const res = deleteRowById_(SHEETS.USERS, id);
  SpreadsheetApp.flush();
  return res;
}

/* ================================================================== *
 *  معلومات تشخيصية (تُطلب عند الضغط على الزر فقط — ليست في كل تحميل)
 * ================================================================== */
function getSystemInfo(token) {
  requireRole_(token, []);
  const ss = getOrCreateDatabase_();
  const ref = readReferenceData_();
  return {
    spreadsheetId: ss.getId(),
    spreadsheetUrl: ss.getUrl(),
    allSheetNames: ss.getSheets().map(function (s) { return s.getName(); }),
    professionsCount: ref.professions.length,
    centersCount: ref.centers.length,
    evaluatorsCount: ref.evaluators.length,
    candidatesCount: Math.max(0, getSheet_(SHEETS.CANDIDATES).getLastRow() - 1)
  };
}

/**
 * مسح الذاكرة المؤقتة يدويًا (لو عدّلت بيانات مباشرة داخل Google Sheet
 * ولم تظهر في النموذج فورًا).
 */
function clearCache(token) {
  requireRole_(token, ['admin']);
  invalidateCache_();
  return { success: true };
}
