/**
 * API ABSENSI DIGITAL PJOK 5A
 * Backend Google Apps Script untuk frontend GitHub Pages.
 *
 * Spreadsheet harus terhubung ke project ini:
 * buka Google Sheets > Ekstensi > Apps Script.
 *
 * Sheet Siswa   : NISN | Nama | Kelas
 * Sheet Absensi : Tanggal | Waktu | NISN | Nama | Kelas | Status
 */

const CONFIG = {
  TIME_ZONE: 'Asia/Jakarta',
  STUDENTS_SHEET: 'Siswa',
  ATTENDANCE_SHEET: 'Absensi',
  STUDENTS_HEADERS: ['NISN', 'Nama', 'Kelas'],
  ATTENDANCE_HEADERS: ['Tanggal', 'Waktu', 'NISN', 'Nama', 'Kelas', 'Status']
};

/**
 * API JSONP untuk GitHub Pages.
 * Contoh: ?action=getStudents&callback=namaCallback
 * JSONP digunakan supaya browser tidak terhalang CORS.
 */
function doGet(e) {
  const params = (e && e.parameter) || {};
  const callback = String(params.callback || '');
  let response;

  try {
    const action = String(params.action || '');
    let data;

    switch (action) {
      case 'ping':
        data = { message: 'API Absensi PJOK aktif.' };
        break;
      case 'getStudents':
        data = getStudents();
        break;
      case 'getDashboardData':
        data = getDashboardData();
        break;
      case 'getReport':
        data = getReport(params.startDate, params.endDate);
        break;
      case 'recordScan':
        data = recordScan(params.qrText);
        break;
      default:
        throw new Error('Aksi API tidak dikenal: ' + action);
    }
    response = { success: true, data: data };
  } catch (err) {
    response = {
      success: false,
      message: err && err.message ? err.message : String(err)
    };
  }

  const json = JSON.stringify(response);
  // Callback hanya menerima karakter nama fungsi JS yang aman.
  if (callback) {
    if (!/^[A-Za-z_$][0-9A-Za-z_$\.]*$/.test(callback)) {
      return ContentService.createTextOutput('Invalid callback')
        .setMimeType(ContentService.MimeType.TEXT);
    }
    return ContentService.createTextOutput(callback + '(' + json + ');')
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return ContentService.createTextOutput(json)
    .setMimeType(ContentService.MimeType.JSON);
}

/** Jalankan satu kali untuk membuat sheet dan header. */
function setupSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) throw new Error('Buka Apps Script dari Google Spreadsheet tujuan.');

  ss.setSpreadsheetTimeZone(CONFIG.TIME_ZONE);

  let siswa = ss.getSheetByName(CONFIG.STUDENTS_SHEET);
  if (!siswa) siswa = ss.insertSheet(CONFIG.STUDENTS_SHEET);
  siswa.getRange(1, 1, 1, 3).setValues([CONFIG.STUDENTS_HEADERS]);
  siswa.setFrozenRows(1);
  siswa.getRange('A:A').setNumberFormat('@');

  let absensi = ss.getSheetByName(CONFIG.ATTENDANCE_SHEET);
  if (!absensi) absensi = ss.insertSheet(CONFIG.ATTENDANCE_SHEET);
  absensi.getRange(1, 1, 1, 6).setValues([CONFIG.ATTENDANCE_HEADERS]);
  absensi.setFrozenRows(1);
  absensi.getRange('A:A').setNumberFormat('yyyy-mm-dd');
  absensi.getRange('B:B').setNumberFormat('HH:mm:ss');
  absensi.getRange('C:C').setNumberFormat('@');

  return 'Sheet Siswa dan Absensi berhasil disiapkan.';
}

/** Mengambil siswa. NISN dibaca sebagai teks agar nol di depan tidak hilang. */
function getStudents() {
  const sheet = getSheet_(CONFIG.STUDENTS_SHEET);
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  return sheet.getRange(2, 1, lastRow - 1, 3).getDisplayValues()
    .filter(r => r[0] && r[1])
    .map(r => ({
      nisn: String(r[0]).trim(),
      nama: String(r[1]).trim(),
      kelas: String(r[2] || '').trim()
    }));
}

/** Mencatat satu scan QR sebagai Hadir; duplikat pada tanggal yang sama ditolak. */
function recordScan(qrText) {
  if (!qrText) throw new Error('QR Code kosong atau tidak terbaca.');

  let payload;
  try {
    payload = JSON.parse(String(qrText));
  } catch (err) {
    payload = { nisn: String(qrText).trim() };
  }

  const nisn = String(payload.nisn || '').trim();
  if (!nisn) throw new Error('QR tidak memuat NISN siswa.');

  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const students = getStudents();
    const student = students.find(s => s.nisn === nisn);
    if (!student) throw new Error('NISN ' + nisn + ' tidak ditemukan pada sheet Siswa.');

    const now = new Date();
    const today = Utilities.formatDate(now, CONFIG.TIME_ZONE, 'yyyy-MM-dd');
    const sheet = getSheet_(CONFIG.ATTENDANCE_SHEET);
    const lastRow = sheet.getLastRow();

    if (lastRow >= 2) {
      const rows = sheet.getRange(2, 1, lastRow - 1, 6).getValues();
      const exists = rows.some(r =>
        dateKey_(r[0]) === today && String(r[2]).trim() === nisn
      );
      if (exists) {
        return { success: true, message: student.nama + ' sudah tercatat hadir hari ini.' };
      }
    }

    const dateOnly = new Date(today + 'T00:00:00+07:00');
    sheet.appendRow([dateOnly, now, student.nisn, student.nama, student.kelas, 'Hadir']);
    return { success: true, message: 'Berhasil! ' + student.nama + ' (' + student.kelas + ') tercatat HADIR.' };
  } finally {
    lock.releaseLock();
  }
}

/** Statistik dashboard hari ini. */
function getDashboardData() {
  const students = getStudents();
  const today = Utilities.formatDate(new Date(), CONFIG.TIME_ZONE, 'yyyy-MM-dd');
  const sheet = getSheet_(CONFIG.ATTENDANCE_SHEET);
  const lastRow = sheet.getLastRow();

  const result = {
    total: students.length, hadir: 0, izin: 0, sakit: 0, alfa: 0,
    belum: students.length, latest: []
  };
  if (lastRow < 2) return result;

  const rows = sheet.getRange(2, 1, lastRow - 1, 6).getValues();
  const todayRows = rows.filter(r => dateKey_(r[0]) === today);

  todayRows.forEach(r => {
    const status = normalizeStatus_(r[5]);
    if (status === 'Hadir') result.hadir++;
    if (status === 'Izin') result.izin++;
    if (status === 'Sakit') result.sakit++;
    if (status === 'Alfa') result.alfa++;
  });

  result.belum = Math.max(0, students.length - todayRows.length);
  result.latest = todayRows.slice()
    .sort((a, b) => new Date(b[1]).getTime() - new Date(a[1]).getTime())
    .slice(0, 10)
    .map(r => ({
      nama: String(r[3] || ''),
      kelas: String(r[4] || ''),
      status: normalizeStatus_(r[5]),
      waktu: r[1] instanceof Date
        ? Utilities.formatDate(r[1], CONFIG.TIME_ZONE, 'HH:mm:ss')
        : String(r[1] || '')
    }));
  return result;
}

/** Rekap laporan untuk rentang tanggal. */
function getReport(startDate, endDate) {
  if (!startDate || !endDate || startDate > endDate) {
    throw new Error('Rentang tanggal tidak valid.');
  }

  const sheet = getSheet_(CONFIG.ATTENDANCE_SHEET);
  const lastRow = sheet.getLastRow();
  const result = { Hadir: 0, Izin: 0, Sakit: 0, Alfa: 0, rows: [] };
  if (lastRow < 2) return result;

  const rows = sheet.getRange(2, 1, lastRow - 1, 6).getValues()
    .filter(r => {
      const key = dateKey_(r[0]);
      return key && key >= startDate && key <= endDate;
    })
    .sort((a, b) => dateKey_(a[0]).localeCompare(dateKey_(b[0])) ||
      new Date(a[1]).getTime() - new Date(b[1]).getTime());

  rows.forEach(r => {
    const status = normalizeStatus_(r[5]);
    if (Object.prototype.hasOwnProperty.call(result, status)) result[status]++;
    result.rows.push({
      tanggal: dateKey_(r[0]),
      nama: String(r[3] || ''),
      kelas: String(r[4] || ''),
      status: status,
      waktu: r[1] instanceof Date
        ? Utilities.formatDate(r[1], CONFIG.TIME_ZONE, 'HH:mm:ss')
        : String(r[1] || '')
    });
  });
  return result;
}

/** Fungsi manual opsional untuk mengubah status siswa. */
function setAttendanceStatus(nisn, status, dateString) {
  const normalized = normalizeStatus_(status);
  if (!['Hadir', 'Izin', 'Sakit', 'Alfa'].includes(normalized)) {
    throw new Error('Status harus Hadir, Izin, Sakit, atau Alfa.');
  }
  const student = getStudents().find(s => s.nisn === String(nisn).trim());
  if (!student) throw new Error('NISN tidak ditemukan.');

  const key = dateString || Utilities.formatDate(new Date(), CONFIG.TIME_ZONE, 'yyyy-MM-dd');
  if (!/^\d{4}-\d{2}-\d{2}$/.test(key)) throw new Error('Format tanggal harus yyyy-MM-dd.');

  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const sheet = getSheet_(CONFIG.ATTENDANCE_SHEET);
    const lastRow = sheet.getLastRow();
    const dateOnly = new Date(key + 'T00:00:00+07:00');
    const now = new Date();

    if (lastRow >= 2) {
      const rows = sheet.getRange(2, 1, lastRow - 1, 6).getValues();
      for (let i = 0; i < rows.length; i++) {
        if (dateKey_(rows[i][0]) === key && String(rows[i][2]).trim() === student.nisn) {
          sheet.getRange(i + 2, 6).setValue(normalized);
          return 'Status ' + student.nama + ' diperbarui menjadi ' + normalized + '.';
        }
      }
    }
    sheet.appendRow([dateOnly, now, student.nisn, student.nama, student.kelas, normalized]);
    return 'Status ' + student.nama + ' dicatat sebagai ' + normalized + '.';
  } finally {
    lock.releaseLock();
  }
}

function getSheet_(name) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) throw new Error('Spreadsheet tidak terhubung. Buka Apps Script dari spreadsheet.');
  const sheet = ss.getSheetByName(name);
  if (!sheet) throw new Error('Sheet "' + name + '" belum ada. Jalankan setupSheets() dahulu.');
  return sheet;
}

function dateKey_(value) {
  if (!value) return '';
  if (value instanceof Date && !isNaN(value.getTime())) {
    return Utilities.formatDate(value, CONFIG.TIME_ZONE, 'yyyy-MM-dd');
  }
  const text = String(value).trim();
  if (/^\d{4}-\d{2}-\d{2}$/.test(text)) return text;
  const parsed = new Date(value);
  return isNaN(parsed.getTime()) ? '' :
    Utilities.formatDate(parsed, CONFIG.TIME_ZONE, 'yyyy-MM-dd');
}

function normalizeStatus_(value) {
  const s = String(value || '').trim().toLowerCase();
  if (s === 'hadir') return 'Hadir';
  if (s === 'izin') return 'Izin';
  if (s === 'sakit') return 'Sakit';
  if (s === 'alfa' || s === 'alpha') return 'Alfa';
  return s ? s.charAt(0).toUpperCase() + s.slice(1) : '';
}
