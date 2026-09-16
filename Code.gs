/**
 * ================================================================
 *  NEZHAMAINAN — SHOPEE SUPER SUMMIT 2026
 *  Backend Google Apps Script (Code.gs) — Pilihan B
 *
 *  CARA PAKAI:
 *  1. Paste kode ini ke editor Google Apps Script.
 *  2. Klik 💾 Simpan → Deploy → Manage deployments →
 *     Edit deployment → Version: New version → Deploy.
 *  3. URL deployment tidak berubah selama kamu edit deployment yang sama.
 *
 *  CARA TEST DI EDITOR (tanpa error):
 *  → Pilih fungsi "testSimpanData" di dropdown atas, lalu klik ▶ Run.
 *  → JANGAN run "doGet" langsung dari editor — pasti error karena
 *    tidak ada HTTP request masuk (parameter "e" tidak tersedia).
 * ================================================================
 */

/* ── Konfigurasi Spreadsheet — sesuaikan jika diperlukan ── */
var SPREADSHEET_ID = "1ho9r5WqV8tz_i1t7YQBKofN4gQWt0Bz5NjW4ItpQqCI";
var SHEET_NAME     = "Data Affiliate";


/**
 * doGet — Entry point utama.
 * Dipanggil otomatis oleh Google saat ada HTTP GET ke URL deployment.
 * TIDAK bisa dijalankan dari tombol ▶ Run di editor — gunakan testSimpanData().
 */
function doGet(e) {
  /* Guard: jika dipanggil tanpa request (misal dari editor) → beri pesan jelas */
  if (!e || !e.parameter) {
    return ContentService
      .createTextOutput(JSON.stringify({
        success: false,
        message: 'doGet hanya bisa dipanggil via URL deployment, bukan dari editor. Gunakan fungsi testSimpanData() untuk test.'
      }))
      .setMimeType(ContentService.MimeType.JSON);
  }

  var callback = e.parameter.callback;
  var result;

  try {
    result = simpanData(e.parameter);
  } catch (err) {
    result = {
      success: false,
      message: 'Terjadi kesalahan sistem: ' + err.toString()
    };
  }

  /* Jika ada parameter callback → kembalikan sebagai JSONP */
  if (callback) {
    return ContentService
      .createTextOutput(callback + '(' + JSON.stringify(result) + ')')
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }

  /* Fallback tanpa callback → JSON biasa */
  return ContentService
    .createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}


/**
 * testSimpanData — Fungsi khusus untuk test dari editor Apps Script.
 * Pilih fungsi ini di dropdown lalu klik ▶ Run.
 * Akan menyimpan 1 baris data dummy ke spreadsheet untuk memverifikasi koneksi.
 */
function testSimpanData() {
  var hasilTest = simpanData({
    nama:     'TEST - Hapus Baris Ini',
    username: '@test_nezhamainan',
    phone:    '081200000000',
    produk:   'Lampu LED Test'
  });

  Logger.log('=== HASIL TEST ===');
  Logger.log(JSON.stringify(hasilTest, null, 2));

  if (hasilTest.success) {
    Logger.log('✅ BERHASIL! Baris test sudah masuk ke sheet. Hapus baris tersebut setelah test.');
  } else {
    Logger.log('❌ GAGAL: ' + hasilTest.message);
  }
}


/**
 * Normalisasi nomor WhatsApp ke format 62xxxxxxxxxx.
 * Aturan:
 *   0812...  → 62812...
 *   8123...  → 628123...
 *   62812... → tetap 62812...
 */
function normalPhone(phone) {
  var clean = String(phone || '').replace(/\D/g, ''); // ambil hanya angka
  if (!clean) return '';
  if (clean.charAt(0) === '0') return '62' + clean.substring(1);
  if (clean.charAt(0) === '8') return '62' + clean;
  return clean; // sudah 62... atau format lain
}


/**
 * Normalisasi username Shopee.
 * Hapus @, spasi, ubah ke huruf kecil semua.
 * Contoh: " @Nezha_Mainan " → "nezha_mainan"
 */
function normalUsername(username) {
  return String(username || '')
    .trim()
    .toLowerCase()
    .replace(/^@+/, '')
    .replace(/\s+/g, '');
}


/**
 * simpanData — Fungsi utama: validasi duplikat & simpan ke Google Sheet.
 * Menggunakan ScriptLock agar aman jika 2 orang submit bersamaan.
 */
function simpanData(data) {

  /* ── ScriptLock: antrian submit agar tidak race condition ── */
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(12000); // tunggu antrean max 12 detik
  } catch (e) {
    return {
      success: false,
      message: 'Server sedang sibuk memproses antrean. Silakan klik KIRIM DATA kembali.'
    };
  }

  try {
    /* ── Koneksi ke sheet menggunakan ID & nama yang sudah dikonfigurasi ── */
    var ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheetByName(SHEET_NAME);

    if (!sheet) {
      return {
        success: false,
        message: 'Sheet "' + SHEET_NAME + '" tidak ditemukan di Spreadsheet. Pastikan nama sheet sudah benar.'
      };
    }

    var lastRow = sheet.getLastRow();

    /* ── Bersihkan & normalisasi input ── */
    var cleanNama  = String(data.nama    || '').trim();
    var cleanUser  = normalUsername(data.username);
    var cleanPhone = normalPhone(data.phone);
    var produk     = String(data.produk  || '').trim() || '-';

    /* ── Validasi field wajib ── */
    if (!cleanNama) {
      return { success: false, message: 'Nama lengkap wajib diisi!' };
    }
    if (!cleanUser) {
      return { success: false, message: 'Username Shopee wajib diisi!' };
    }
    if (!cleanPhone || cleanPhone.length < 9) {
      return { success: false, message: 'Nomor WhatsApp tidak valid. Periksa kembali.' };
    }

    /* ── Buat header otomatis jika sheet masih kosong ── */
    if (lastRow < 1) {
      sheet.appendRow([
        'Timestamp',
        'Kode Klaim',
        'Nama Lengkap',
        'Username Shopee (Input)',
        'Username Normal',
        'No. WhatsApp (Input)',
        'No. WA Normal',
        'Produk Sample',
        'Status'
      ]);
      lastRow = 1;
    }

    /* ── CEK DUPLIKASI (LAPIS 2) ── */
    if (lastRow >= 2) {
      var rows = sheet.getRange(2, 1, lastRow - 1, 9).getValues();

      for (var i = 0; i < rows.length; i++) {
        var rowUser  = String(rows[i][4] || '').trim(); // Kolom E: Username Normal
        var rowPhone = String(rows[i][6] || '').trim(); // Kolom G: No. WA Normal

        /* Duplikat Username Shopee */
        if (rowUser && rowUser === cleanUser) {
          return {
            success:   false,
            duplicate: true,
            message:
              'Username Shopee @' + (data.username || '') +
              ' sudah pernah mengambil sample sebelumnya.\n' +
              'Maksimal 1 sample per akun affiliate.'
          };
        }

        /* Duplikat Nomor WhatsApp */
        if (rowPhone && rowPhone === cleanPhone) {
          return {
            success:   false,
            duplicate: true,
            message:
              'Nomor WhatsApp ' + (data.phone || '') +
              ' sudah terdaftar mengambil sample.\n' +
              'Maksimal 1 sample per orang.'
          };
        }
      }
    }

    /* ── LOLOS SEMUA CEK → Generate Kode Klaim Unik ── */
    var timestamp  = new Date();
    var randomCode = 1000 + Math.floor(Math.random() * 9000);
    var kodeKlaim  = 'NZH-' + randomCode;

    /* Simpan ke baris baru di Spreadsheet */
    sheet.appendRow([
      timestamp,
      kodeKlaim,
      cleanNama,
      data.username || '',
      cleanUser,
      data.phone    || '',
      cleanPhone,
      produk,
      'SUDAH DIAMBIL'
    ]);

    var waktu = Utilities.formatDate(timestamp, 'GMT+7', 'dd/MM/yyyy HH:mm') + ' WIB';

    return {
      success:   true,
      kodeKlaim: kodeKlaim,
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
