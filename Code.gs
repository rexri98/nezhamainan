/**
 * ================================================================
 *  NEZHAMAINAN — SHOPEE SUPER SUMMIT 2026
 *  Backend Google Apps Script (Code.gs)
 *  Sistem: JALUR PRIORITAS ACC PENGAJUAN SAMPLE GRATIS SHOPEE
 *  Status di Google Sheet:
 *    - "MENUNGGU ACC" (Status awal saat kreator mendaftar)
 *    - "SUDAH DI-ACC" (Diubah oleh Seller saat approve pengajuan sample di Shopee)
 * ================================================================
 */

/* ── Konfigurasi Spreadsheet ── */
var SPREADSHEET_ID = "1ho9r5WqV8tz_i1t7YQBKofN4gQWt0Bz5NjW4ItpQqCI";
var SHEET_NAME     = "Data Affiliate";


/**
 * doGet — Entry point utama untuk submit pengajuan dan cek status ACC realtime.
 */
function doGet(e) {
  if (!e || !e.parameter) {
    return ContentService
      .createTextOutput(JSON.stringify({
        success: false,
        message: 'Akses langsung via editor tidak didukung. Gunakan URL deployment.'
      }))
      .setMimeType(ContentService.MimeType.JSON);
  }

  var callback = e.parameter.callback;
  var action   = e.parameter.action;
  var result;

  try {
    /* Route 1: Cek status terbaru dari Spreadsheet (Pengunjung) */
    if (action === 'cekStatus') {
      result = cekStatusKlaim(e.parameter.kodeKlaim);
    }
    /* Route 2: Baca pengaturan publik (Pengunjung & Admin) */
    else if (action === 'getSettings') {
      result = bacaPengaturanPublik();
    }
    /* Route 3: Verifikasi login admin */
    else if (action === 'adminLogin') {
      result = adminVerifikasiLogin(e.parameter.pin);
    }
    /* Route 4: Ambil semua data pendaftar & statistik */
    else if (action === 'adminGetData') {
      result = adminAmbilData(e.parameter.pin);
    }
    /* Route 5: Ubah status pendaftar (ACC / Tolak / Pending) */
    else if (action === 'adminUpdateStatus') {
      result = adminUbahStatus(e.parameter);
    }
    /* Route 6: Ambil pengaturan admin lengkap */
    else if (action === 'adminGetSettings') {
      result = adminAmbilPengaturan(e.parameter.pin);
    }
    /* Route 7: Simpan pengaturan admin */
    else if (action === 'adminSaveSettings') {
      result = adminSimpanPengaturan(e.parameter);
    }
    /* Route Default: Submit Pendaftaran Baru */
    else {
      result = simpanData(e.parameter);
    }
  } catch (err) {
    result = {
      success: false,
      message: 'Terjadi kesalahan sistem: ' + err.toString()
    };
  }

  /* Kembalikan sebagai JSONP agar bebas CORS di semua browser HP */
  if (callback) {
    return ContentService
      .createTextOutput(callback + '(' + JSON.stringify(result) + ')')
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }

  return ContentService
    .createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}


/**
 * Cek status terkini dari Google Sheet berdasarkan Kode Prioritas (misal: NZH-4836)
 */
function cekStatusKlaim(kodeKlaim) {
  if (!kodeKlaim) {
    return { success: false, message: 'Harap masukkan kode prioritas, nomor WhatsApp, atau username.' };
  }

  var ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName(SHEET_NAME) || ss.getSheets()[0];
  if (!sheet) {
    return { success: false, message: 'Sheet tidak ditemukan' };
  }

  var lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    return { success: false, message: 'Data belum ada' };
  }

  var cleanQuery = String(kodeKlaim).trim();
  var cleanKode  = cleanQuery.toUpperCase();
  var cleanUser  = normalUsername(cleanQuery);
  var cleanPhone = normalPhone(cleanQuery);

  var rows = sheet.getRange(2, 1, lastRow - 1, 9).getValues();

  for (var i = 0; i < rows.length; i++) {
    var rowKode  = String(rows[i][1] || '').trim().toUpperCase(); // Kolom B: Kode Klaim
    var rowUser  = String(rows[i][4] || '').trim().toLowerCase(); // Kolom E: Username Normal
    var rowPhone = String(rows[i][6] || '').trim();               // Kolom G: No WA Normal

    var isMatch = false;
    if (rowKode && rowKode === cleanKode) {
      isMatch = true;
    } else if (cleanPhone && cleanPhone.length >= 8 && rowPhone && (rowPhone === cleanPhone || rowPhone.indexOf(cleanPhone) !== -1)) {
      isMatch = true;
    } else if (cleanUser && cleanUser.length >= 3 && rowUser && rowUser === cleanUser) {
      isMatch = true;
    }

    if (isMatch) {
      var currentStatus = String(rows[i][8] || 'MENUNGGU ACC').trim(); // Kolom I: Status
      var rowWaktu = rows[i][0];
      var waktuFormatted = '-';
      if (rowWaktu instanceof Date) {
        waktuFormatted = Utilities.formatDate(rowWaktu, 'GMT+7', 'dd/MM/yyyy HH:mm') + ' WIB';
      } else if (rowWaktu) {
        waktuFormatted = String(rowWaktu);
      }

      return {
        success:   true,
        kodeKlaim: rowKode,
        status:    currentStatus, // "MENUNGGU ACC", "SUDAH DI-ACC", atau "DITOLAK"
        nama:      rows[i][2] || '-',
        username:  rows[i][3] || '-',
        phone:     rows[i][5] || '-',
        produk:    rows[i][7] || '-',
        waktu:     waktuFormatted
      };
    }
  }

  return { success: false, message: 'Data pendaftaran tidak ditemukan. Pastikan Kode Klaim (NZH-xxxx), No. WA, atau Username Shopee yang kamu masukkan benar.' };
}


/**
 * Normalisasi nomor WhatsApp ke format 62xxxxxxxxxx
 */
function normalPhone(phone) {
  var clean = String(phone || '').replace(/\D/g, '');
  if (!clean) return '';
  if (clean.charAt(0) === '0') return '62' + clean.substring(1);
  if (clean.charAt(0) === '8') return '62' + clean;
  return clean;
}


/**
 * Normalisasi username Shopee
 */
function normalUsername(username) {
  return String(username || '')
    .trim()
    .toLowerCase()
    .replace(/^@+/, '')
    .replace(/\s+/g, '');
}


/**
 * Simpan pendaftaran baru & cek duplikasi (1 Akun Shopee / WA = 1 Pengajuan Prioritas)
 */
function simpanData(data) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(12000);
  } catch (e) {
    return {
      success: false,
      message: 'Server sedang sibuk. Silakan coba klik KIRIM DATA kembali.'
    };
  }

  try {
    var settings = bacaPengaturanPublik();
    if (!settings.isOpen) {
      return {
        success: false,
        message: 'Mohon maaf, pendaftaran sample gratis sedang dijeda sementara oleh seller.'
      };
    }

    var ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheetByName(SHEET_NAME) || ss.getSheets()[0];

    if (!sheet) {
      return {
        success: false,
        message: 'Sheet tidak ditemukan di Spreadsheet.'
      };
    }

    var lastRow = sheet.getLastRow();

    var cleanNama  = String(data.nama    || '').trim();
    var cleanUser  = normalUsername(data.username);
    var cleanPhone = normalPhone(data.phone);
    var produk     = String(data.produk  || '').trim() || '-';

    if (!cleanNama) {
      return { success: false, message: 'Nama lengkap wajib diisi!' };
    }
    if (!cleanUser) {
      return { success: false, message: 'Username Shopee wajib diisi!' };
    }
    if (!cleanPhone || cleanPhone.length < 9) {
      return { success: false, message: 'Nomor WhatsApp tidak valid. Periksa kembali.' };
    }

    /* Buat header jika kosong */
    if (lastRow < 1) {
      sheet.appendRow([
        'Timestamp',
        'KODE PRIORITAS',
        'Nama Lengkap',
        'USERNAME SHOPEE (Input)',
        'Username Normal',
        'No. WhatsApp (Input)',
        'No. WA Normal',
        'PRODUK SAMPLE',
        'STATUS ACC'
      ]);
      lastRow = 1;
    }

    /* Cek Duplikasi Nomor WA & Username Shopee */
    if (lastRow >= 2) {
      var rows = sheet.getRange(2, 1, lastRow - 1, 9).getValues();

      for (var i = 0; i < rows.length; i++) {
        var rowUser  = String(rows[i][4] || '').trim();
        var rowPhone = String(rows[i][6] || '').trim();

        if (rowUser && rowUser === cleanUser) {
          return {
            success:   false,
            duplicate: true,
            message:
              'Username Shopee @' + (data.username || '') +
              ' sudah pernah mengajukan sample sebelumnya.\n' +
              'Maksimal 1 pengajuan sample per akun affiliate.'
          };
        }

        if (rowPhone && rowPhone === cleanPhone) {
          return {
            success:   false,
            duplicate: true,
            message:
              'Nomor WhatsApp ' + (data.phone || '') +
              ' sudah terdaftar dalam antrean prioritas.\n' +
              'Maksimal 1 pengajuan per orang.'
          };
        }
      }
    }

    /* Generate Kode Prioritas Unik */
    var timestamp  = new Date();
    var randomCode = 1000 + Math.floor(Math.random() * 9000);
    var kodeKlaim  = 'NZH-' + randomCode;

    /* Simpan dengan status awal: MENUNGGU ACC */
    sheet.appendRow([
      timestamp,
      kodeKlaim,
      cleanNama,
      data.username || '',
      cleanUser,
      data.phone    || '',
      cleanPhone,
      produk,
      'MENUNGGU ACC'
    ]);

    var waktu = Utilities.formatDate(timestamp, 'GMT+7', 'dd/MM/yyyy HH:mm') + ' WIB';

    return {
      success:   true,
      kodeKlaim: kodeKlaim,
      status:    'MENUNGGU ACC',
      nama:      cleanNama,
      username:  data.username || '',
      phone:     data.phone    || '',
      produk:    produk,
      waktu:     waktu
    };

  } catch (err) {
    return {
      success: false,
      message: 'Terjadi kesalahan sistem: ' + err.toString()
    };
  } finally {
    lock.releaseLock();
  }
}


/**
 * ================================================================
 * FUNGSI BANTUAN SEKALI KLIK:
 * Pilih "pasangDropdownSeluruhKolom" di toolbar atas, lalu klik ▶ Run.
 * Seluruh Kolom I (baris 2 sampai 1000) akan langsung jadi Dropdown!
 * ================================================================
 */
function pasangDropdownSeluruhKolom() {
  var ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName(SHEET_NAME) || ss.getSheets()[0];
  var rule  = SpreadsheetApp.newDataValidation()
    .requireValueInList(['MENUNGGU ACC', 'SUDAH DI-ACC', 'DITOLAK'], true)
    .build();

  sheet.getRange('I2:I1000').setDataValidation(rule);
  Logger.log('✅ Berhasil! Seluruh Kolom I (baris 2-1000) sekarang sudah memiliki menu Dropdown (MENUNGGU ACC / SUDAH DI-ACC / DITOLAK).');
}


/**
 * ================================================================
 *  MODUL ADMIN PANEL & PENGATURAN (PROPERTIES SERVICE)
 * ================================================================
 */

var DEFAULT_PIN        = '2026';
var DEFAULT_SHOPEE_URL = 'https://shopee.co.id/nezhamainan';
var DEFAULT_CHIPS      = [
  'Mobil Remot RC',
  'Robot Action Figure',
  'Puzzle Edukasi Kayu',
  'Water / Bubble Gun',
  'Board Game Anak'
];

function getAdminPin() {
  var props = PropertiesService.getScriptProperties();
  return props.getProperty('ADMIN_PIN') || DEFAULT_PIN;
}

function adminVerifikasiLogin(pin) {
  var validPin = getAdminPin();
  if (String(pin || '').trim() === String(validPin).trim()) {
    return { success: true, message: 'Login admin berhasil!' };
  }
  return { success: false, message: 'PIN Admin salah. Silakan coba lagi.' };
}

function adminAmbilData(pin) {
  var auth = adminVerifikasiLogin(pin);
  if (!auth.success) return auth;

  var ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName(SHEET_NAME) || ss.getSheets()[0];
  if (!sheet) {
    return { success: false, message: 'Sheet tidak ditemukan' };
  }

  var lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    return {
      success: true,
      stats: { total: 0, pending: 0, approved: 0, rejected: 0 },
      data: []
    };
  }

  var rows = sheet.getRange(2, 1, lastRow - 1, 9).getValues();
  var list = [];
  var stats = { total: rows.length, pending: 0, approved: 0, rejected: 0 };

  for (var i = 0; i < rows.length; i++) {
    var rawStatus = String(rows[i][8] || 'MENUNGGU ACC').trim().toUpperCase();
    var statusCat = 'pending';
    if (rawStatus === 'SUDAH DI-ACC' || rawStatus === 'SUDAH DIAMBIL' || rawStatus === 'DISETUJUI') {
      statusCat = 'approved';
      stats.approved++;
    } else if (rawStatus === 'DITOLAK' || rawStatus === 'TIDAK DI-ACC' || rawStatus === 'REJECTED') {
      statusCat = 'rejected';
      stats.rejected++;
    } else {
      stats.pending++;
    }

    var rowWaktu = rows[i][0];
    var waktuFormatted = '-';
    var ts = 0;
    if (rowWaktu instanceof Date) {
      waktuFormatted = Utilities.formatDate(rowWaktu, 'GMT+7', 'dd/MM/yyyy HH:mm') + ' WIB';
      ts = rowWaktu.getTime();
    } else if (rowWaktu) {
      waktuFormatted = String(rowWaktu);
    }

    list.push({
      rowIndex:  i + 2,
      timestamp: ts,
      waktu:     waktuFormatted,
      kodeKlaim: String(rows[i][1] || '').trim(),
      nama:      String(rows[i][2] || '').trim(),
      username:  String(rows[i][3] || '').trim(),
      phone:     String(rows[i][5] || '').trim(),
      produk:    String(rows[i][7] || '').trim(),
      status:    rawStatus,
      statusCat: statusCat
    });
  }

  // Urutkan dari yang paling baru daftar
  list.sort(function(a, b) {
    return b.timestamp - a.timestamp;
  });

  return {
    success: true,
    stats:   stats,
    data:    list
  };
}

function adminUbahStatus(p) {
  var auth = adminVerifikasiLogin(p.pin);
  if (!auth.success) return auth;

  var kode = String(p.kodeKlaim || '').trim().toUpperCase();
  var statusBaru = String(p.statusBaru || '').trim().toUpperCase();

  if (!kode) {
    return { success: false, message: 'Kode klaim wajib diisi' };
  }
  if (!statusBaru) {
    return { success: false, message: 'Status baru wajib diisi' };
  }

  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
  } catch (e) {
    return { success: false, message: 'Server sibuk. Silakan coba lagi.' };
  }

  try {
    var ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheetByName(SHEET_NAME) || ss.getSheets()[0];
    if (!sheet) return { success: false, message: 'Sheet tidak ditemukan' };

    var lastRow = sheet.getLastRow();
    if (lastRow < 2) return { success: false, message: 'Data masih kosong' };

    var rows = sheet.getRange(2, 2, lastRow - 1, 1).getValues(); // Kolom B: Kode Klaim
    var targetRow = -1;

    for (var i = 0; i < rows.length; i++) {
      if (String(rows[i][0] || '').trim().toUpperCase() === kode) {
        targetRow = i + 2;
        break;
      }
    }

    if (targetRow === -1) {
      return { success: false, message: 'Data pendaftar dengan kode ' + kode + ' tidak ditemukan' };
    }

    // Update Kolom I (Kolom 9): Status ACC
    sheet.getRange(targetRow, 9).setValue(statusBaru);

    return {
      success:    true,
      message:    'Status ' + kode + ' berhasil diperbarui menjadi ' + statusBaru,
      kodeKlaim:  kode,
      statusBaru: statusBaru
    };

  } catch (err) {
    return { success: false, message: 'Gagal update status: ' + err.toString() };
  } finally {
    lock.releaseLock();
  }
}

function bacaPengaturanPublik() {
  var props = PropertiesService.getScriptProperties();
  var shopeeUrl    = props.getProperty('SHOPEE_URL') || DEFAULT_SHOPEE_URL;
  var isOpenStr    = props.getProperty('IS_OPEN');
  var isOpen       = isOpenStr !== null ? (isOpenStr !== 'false') : true;
  var announcement = props.getProperty('ANNOUNCEMENT') || '';

  var rawChips = props.getProperty('PROD_CHIPS');
  var chips = DEFAULT_CHIPS;
  if (rawChips) {
    try {
      var parsed = JSON.parse(rawChips);
      if (Array.isArray(parsed) && parsed.length > 0) chips = parsed;
    } catch (_) {}
  }

  return {
    success:      true,
    shopeeUrl:    shopeeUrl,
    isOpen:       isOpen,
    announcement: announcement,
    chips:        chips
  };
}

function adminAmbilPengaturan(pin) {
  var auth = adminVerifikasiLogin(pin);
  if (!auth.success) return auth;

  var pub = bacaPengaturanPublik();
  pub.adminPin = getAdminPin();
  return pub;
}

function adminSimpanPengaturan(p) {
  var auth = adminVerifikasiLogin(p.pin);
  if (!auth.success) return auth;

  var props = PropertiesService.getScriptProperties();

  if (p.newPin && String(p.newPin).trim().length >= 4) {
    props.setProperty('ADMIN_PIN', String(p.newPin).trim());
  }
  if (p.shopeeUrl) {
    props.setProperty('SHOPEE_URL', String(p.shopeeUrl).trim());
  }
  if (p.isOpen !== undefined) {
    props.setProperty('IS_OPEN', String(p.isOpen));
  }
  if (p.announcement !== undefined) {
    props.setProperty('ANNOUNCEMENT', String(p.announcement).trim());
  }
  if (p.chips) {
    var chipsArr = typeof p.chips === 'string' ? JSON.parse(p.chips) : p.chips;
    if (Array.isArray(chipsArr)) {
      props.setProperty('PROD_CHIPS', JSON.stringify(chipsArr));
    }
  }

  return {
    success: true,
    message: 'Pengaturan berhasil disimpan!'
  };
}
