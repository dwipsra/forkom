/**
 * ============================================================================
 * BACKEND GOOGLE APPS SCRIPT - APLIKASI FORKOM RSML (PWA)
 * ============================================================================
 */

const SHEET_NAMES = {
  USERS: 'Users',
  JENIS_PIMPINAN: 'JenisPimpinan',
  MASTER_CABANG: 'MasterCabang',
  MASTER_TTD: 'MasterTTD',
  PESAN: 'Pesan',
  PENGAJUAN_AIK: 'PengajuanAIK',
  SURAT_TUGAS: 'SuratTugas',
  SERTIFIKAT: 'Sertifikat',
  PRESENSI: 'Presensi',
  PENGATURAN: 'Pengaturan',
  OTP_RESET: 'OtpReset'
};

const DRIVE_FOLDERS = {
  UMUM: 'FORKOM_RSML_Uploads',
  SURAT_TUGAS: 'FORKOM_RSML_Surat_Tugas',
  SERTIFIKAT: 'FORKOM_RSML_Sertifikat'
};

/**
 * Menyajikan Halaman Utama Index.html
 */
function doGet(e) {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('FORKOM - Forum Komunikasi Karyawan RSML')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1.0, maximum-scale=1.0, user-scalable=no')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/**
 * Mengambil Spreadsheet aktif (terikat pada container-bound script)
 */
function getActiveDb() {
  let ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) {
    const props = PropertiesService.getScriptProperties();
    let savedId = props.getProperty('FORKOM_DB_ID');
    if (savedId) {
      ss = SpreadsheetApp.openById(savedId);
    } else {
      ss = SpreadsheetApp.create('Database FORKOM RSML');
      props.setProperty('FORKOM_DB_ID', ss.getId());
    }
  }
  return ss;
}

/**
 * Helper mendapatkan atau membuat folder tujuan di Google Drive (terpisah per kategori & tahun)
 */
function getTargetDriveFolder(rootFolderName, subFolderYear) {
  const rootFolders = DriveApp.getFoldersByName(rootFolderName);
  const rootFolder = rootFolders.hasNext() ? rootFolders.next() : DriveApp.createFolder(rootFolderName);
  if (!subFolderYear) {
    return rootFolder;
  }
  const yearStr = String(subFolderYear).trim();
  const subFolders = rootFolder.getFoldersByName(yearStr);
  return subFolders.hasNext() ? subFolders.next() : rootFolder.createFolder(yearStr);
}

/**
 * Helper mengekstrak tahun (misal 2026 / 2027) dari tanggal atau nomor surat
 */
function extractTargetYear(dateOrNumberText) {
  const match = String(dateOrNumberText || '').match(/\b(20\d{2})\b/);
  if (match && match[1]) {
    return match[1];
  }
  return Utilities.formatDate(new Date(), 'Asia/Jakarta', 'yyyy');
}

/**
 * Helper menjaga angka 0 di awal nomor telepon agar tersimpan sebagai teks
 */
function formatPhoneAsText(phone) {
  const str = String(phone || '').trim();
  if (!str || str === '-') return '-';
  if (/^8\d{7,}$/.test(str)) {
    return '0' + str;
  }
  return str;
}

/**
 * Helper mengambil Blob gambar TTD baik dari Base64 Data URI maupun Link Google Drive / URL
 */
function resolveSignatureBlob(ttdSource, prefixName) {
  const raw = String(ttdSource || '').trim();
  if (!raw) return null;
  try {
    if (raw.indexOf('data:image') === 0) {
      const parts = raw.split(',');
      const contentType = parts[0].match(/:(.*?);/)[1];
      return Utilities.newBlob(Utilities.base64Decode(parts[1]), contentType, (prefixName || 'TTD') + '_' + Date.now() + '.png');
    }
    const driveIdMatch = raw.match(/\/d\/([a-zA-Z0-9_-]+)/) || raw.match(/[?&]id=([a-zA-Z0-9_-]+)/);
    if (driveIdMatch && driveIdMatch[1]) {
      return DriveApp.getFileById(driveIdMatch[1]).getBlob();
    }
    if (raw.indexOf('http://') === 0 || raw.indexOf('https://') === 0) {
      return UrlFetchApp.fetch(raw).getBlob();
    }
  } catch (e) {
    Logger.log('Gagal memuat blob TTD: ' + e.message);
  }
  return null;
}

/**
 * Helper upsert nilai pada sheet Pengaturan
 */
function upsertSettingKey(keyName, valStr) {
  const ss = getActiveDb();
  let sheet = ss.getSheetByName(SHEET_NAMES.PENGATURAN);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAMES.PENGATURAN);
    sheet.appendRow(['Key', 'Value']);
  }
  const rows = sheet.getDataRange().getDisplayValues();
  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][0]).trim() === keyName) {
      sheet.getRange(i + 1, 2).setValue(valStr);
      return;
    }
  }
  sheet.appendRow([keyName, valStr]);
}

/**
 * Inisialisasi Otomatis Seluruh Sheet & Akun Bawaan
 */
function setupDatabase() {
  const ss = getActiveDb();

  const schemas = {
    [SHEET_NAMES.USERS]: ['Username', 'Password', 'Role', 'NamaLengkap', 'NIK', 'UnitKerja', 'NoTelp', 'Email', 'ForkomDomisili', 'FotoUrl'],
    [SHEET_NAMES.JENIS_PIMPINAN]: ['Kode', 'Label'],
    [SHEET_NAMES.MASTER_CABANG]: ['ID', 'JenisPimpinan', 'NamaCabang'],
    [SHEET_NAMES.MASTER_TTD]: ['Nama', 'FileTTDUrl'],
    [SHEET_NAMES.PESAN]: ['ID', 'Tanggal', 'TujuanNIK', 'Judul', 'Isi', 'FileUrl'],
    [SHEET_NAMES.PENGAJUAN_AIK]: ['ID', 'Tanggal', 'NIK', 'Nama', 'UnitKerja', 'Email', 'Domisili', 'Periode', 'JenisPimpinan', 'AreaCabang', 'NamaPimpinan', 'AktivitasA', 'AktivitasB', 'AktivitasC', 'BuktiUrl', 'Status', 'HasilDokumenUrl', 'AlasanPenolakan'],
    [SHEET_NAMES.SURAT_TUGAS]: ['ID', 'TanggalTerbit', 'NomorSurat', 'NamaKetua', 'NIKKetua', 'JabatanKetua', 'TTDKetuaUrl', 'NIKPetugas', 'NamaPetugas', 'PeranTugas', 'NamaKegiatan', 'TanggalKegiatan', 'Waktu', 'Lokasi', 'FileUrl'],
    [SHEET_NAMES.SERTIFIKAT]: ['ID', 'TanggalTerbit', 'NomorSertifikat', 'NIKPetugas', 'NamaPetugas', 'Partisipasi', 'NamaKegiatan', 'KetuaPanitia', 'TTDPanitiaUrl', 'FileUrl'],
    [SHEET_NAMES.PRESENSI]: ['ID', 'Tanggal', 'NIK', 'NamaPetugas', 'NamaKegiatan', 'Lokasi', 'FotoBuktiUrl', 'Status'],
    [SHEET_NAMES.PENGATURAN]: ['Key', 'Value'],
    [SHEET_NAMES.OTP_RESET]: ['Email', 'OTP', 'Timestamp']
  };

  Object.keys(schemas).forEach(function(sheetName) {
    let sheet = ss.getSheetByName(sheetName);
    if (!sheet) {
      sheet = ss.insertSheet(sheetName);
    }
    if (sheet.getLastRow() === 0) {
      sheet.appendRow(schemas[sheetName]);
      sheet.getRange(1, 1, 1, schemas[sheetName].length).setFontWeight('bold').setBackground('#E6F4EA');
    }
  });

  const defaultSheet = ss.getSheetByName('Sheet1');
  if (defaultSheet && ss.getSheets().length > 1 && defaultSheet.getLastRow() === 0) {
    ss.deleteSheet(defaultSheet);
  }

  const userSheet = ss.getSheetByName(SHEET_NAMES.USERS);
  if (userSheet.getLastRow() === 1) {
    userSheet.getRange(2, 5, 3, 1).setNumberFormat('@');
    userSheet.getRange(2, 7, 3, 1).setNumberFormat('@');
    userSheet.appendRow(['admin', 'admin2026', 'admin', 'Administrator FORKOM', '999999', 'Bagian IT & SDI', '081234567890', 'admin@rsml.id', 'Lamongan', 'https://placehold.co/150x200/006837/FFFFFF?text=Admin']);
    userSheet.appendRow(['khomarudin', 'edudigital', 'petugas', 'KHOMARUDIN, Amd.RMIK.', '112233', 'Rekam Medis', '081233445566', 'khomarudin@rsml.id', 'Pucuk', 'https://placehold.co/160x200/E6F4EA/006837?text=Khomarudin']);
    userSheet.appendRow(['peserta', 'edudigital', 'petugas', 'M Dwi Hidayatulloh', '123456', 'Instalasi IT', '085812345678', 'dwi@rsml.id', 'Lamongan', 'https://placehold.co/160x200/E6F4EA/006837?text=M+Dwi+H']);
  }

  const jenisSheet = ss.getSheetByName(SHEET_NAMES.JENIS_PIMPINAN);
  if (jenisSheet.getLastRow() === 1) {
    jenisSheet.appendRow(['PRM', 'PRM (Ranting Muhammadiyah)']);
    jenisSheet.appendRow(['PCM', 'PCM (Cabang Muhammadiyah)']);
    jenisSheet.appendRow(['PRA', 'PRA (Ranting Aisyiyah)']);
    jenisSheet.appendRow(['PCA', 'PCA (Cabang Aisyiyah)']);
  }

  const cabangSheet = ss.getSheetByName(SHEET_NAMES.MASTER_CABANG);
  if (cabangSheet.getLastRow() === 1) {
    cabangSheet.appendRow(['C1', 'PCM', 'PCM Pucuk']);
    cabangSheet.appendRow(['C2', 'PRM', 'PRM Made']);
    cabangSheet.appendRow(['C3', 'PCA', 'PCA Lamongan']);
    cabangSheet.appendRow(['C4', 'PRA', 'PRA Pucuk']);
  }

  const ttdSheet = ss.getSheetByName(SHEET_NAMES.MASTER_TTD);
  if (ttdSheet.getLastRow() === 1) {
    ttdSheet.appendRow(['KHOMARUDIN, Amd.RMIK.', '']);
  }

  const setSheet = ss.getSheetByName(SHEET_NAMES.PENGATURAN);
  if (setSheet.getLastRow() === 1) {
    setSheet.appendRow(['REK_BANK', 'BSI (Bank Syariah Indonesia)']);
    setSheet.appendRow(['REK_NOMOR', '4337146830']);
    setSheet.appendRow(['REK_ATASNAMA', 'Latifatul Afifah']);
    setSheet.appendRow(['HERO_BANNER_URL', 'https://placehold.co/960x360/D7EEF7/006837?text=GEDUNG+RS+MUHAMMADIYAH+LAMONGAN']);
    setSheet.appendRow(['DEFAULT_KETUA_NAMA', 'KHOMARUDIN, Amd.RMIK.']);
  }

  invalidateAppCache();
  return 'Setup Database FORKOM RSML Berhasil! ID: ' + ss.getId();
}

const APP_DATA_CACHE_KEY = 'FORKOM_INITIAL_DATA_V1';

function invalidateAppCache() {
  try {
    CacheService.getScriptCache().remove(APP_DATA_CACHE_KEY);
  } catch (e) {}
}

function getCachedDriveFolder(rootName, subYear) {
  const props = PropertiesService.getScriptProperties();
  const cacheKey = 'DRV_FLD_' + rootName + (subYear ? '_' + subYear : '');
  const savedId = props.getProperty(cacheKey);
  if (savedId) {
    try {
      const f = DriveApp.getFolderById(savedId);
      if (!f.isTrashed()) return f;
    } catch (e) {}
  }
  const folder = getTargetDriveFolder(rootName, subYear);
  if (folder) {
    try { props.setProperty(cacheKey, folder.getId()); } catch (e) {}
  }
  return folder;
}

/**
 * Verifikasi Login Pengguna
 */
function loginUser(payload) {
  try {
    const ss = getActiveDb();
    let sheet = ss.getSheetByName(SHEET_NAMES.USERS);
    if (!sheet || sheet.getLastRow() < 2) {
      setupDatabase();
      sheet = ss.getSheetByName(SHEET_NAMES.USERS);
    }
    const rows = sheet.getDataRange().getDisplayValues();
    const uInput = String(payload.usernameOrNik || '').trim();
    const pInput = String(payload.password || '').trim();

    for (let i = 1; i < rows.length; i++) {
      const r = rows[i];
      if ((String(r[0]).trim() === uInput || String(r[4]).trim() === uInput) && String(r[1]).trim() === pInput) {
        return {
          ok: true,
          user: {
            username: r[0],
            password: r[1],
            role: String(r[2] || 'petugas').trim().toLowerCase(),
            namaLengkap: r[3],
            nik: r[4],
            unitKerja: r[5],
            noTelp: formatPhoneAsText(r[6]),
            email: r[7],
            forkomDomisili: r[8],
            fotoUrl: r[9]
          }
        };
      }
    }
    return { error: 'Username atau Password tidak sesuai!' };
  } catch (err) {
    return { error: err.message };
  }
}

/**
 * Ambil Seluruh Data Awal untuk Frontend (Dengan opsi forceRefresh saat aplikasi dibuka kembali)
 */
function getInitialAppData(forceRefresh) {
  try {
    const cache = CacheService.getScriptCache();
    if (!forceRefresh) {
      const cachedJson = cache.get(APP_DATA_CACHE_KEY);
      if (cachedJson) {
        try {
          return JSON.parse(cachedJson);
        } catch (e) {}
      }
    }

    const ss = getActiveDb();
    if (!ss.getSheetByName(SHEET_NAMES.USERS) || !ss.getSheetByName(SHEET_NAMES.MASTER_TTD) || !ss.getSheetByName(SHEET_NAMES.JENIS_PIMPINAN)) {
      setupDatabase();
    }

    const readSheetRows = function(name) {
      const sh = ss.getSheetByName(name);
      if (!sh || sh.getLastRow() < 2) return [];
      return sh.getDataRange().getDisplayValues().slice(1);
    };

    const users = readSheetRows(SHEET_NAMES.USERS).map(function(r) {
      return {
        username: r[0], password: r[1], role: String(r[2] || 'petugas').trim().toLowerCase(),
        namaLengkap: r[3], nik: r[4], unitKerja: r[5], noTelp: formatPhoneAsText(r[6]), email: r[7],
        forkomDomisili: r[8], fotoUrl: r[9]
      };
    });

    const jenisPimpinanList = readSheetRows(SHEET_NAMES.JENIS_PIMPINAN).map(function(r) {
      const kodeClean = String(r[0] || '').trim().toUpperCase();
      return { kode: kodeClean, label: String(r[1] || kodeClean).trim() };
    }).filter(function(item) { return item.kode !== ''; });

    const masterCabang = readSheetRows(SHEET_NAMES.MASTER_CABANG).map(function(r) {
      return {
        id: String(r[0] || '').trim(),
        jenis: String(r[1] || '').trim().toUpperCase(),
        nama: String(r[2] || '').trim()
      };
    }).filter(function(item) { return item.nama !== ''; });

    const masterTtd = readSheetRows(SHEET_NAMES.MASTER_TTD).map(function(r) {
      return { nama: r[0], ttdUrl: r[1] };
    });

    const pesanList = readSheetRows(SHEET_NAMES.PESAN).reverse().map(function(r) {
      return { id: r[0], tanggal: r[1], tujuan: r[2], judul: r[3], isi: r[4], fileUrl: r[5] };
    });

    const aikList = readSheetRows(SHEET_NAMES.PENGAJUAN_AIK).reverse().map(function(r) {
      return {
        id: r[0], tanggal: r[1], nik: r[2], nama: r[3], unitKerja: r[4], email: r[5],
        domisili: r[6], periode: r[7], jenisPimpinan: r[8], areaCabang: r[9],
        namaPimpinan: r[10], a1: r[11], b1: r[12], c1: r[13], buktiUrl: r[14],
        status: r[15], hasilDokumen: r[16], alasanPenolakan: r[17] || ''
      };
    });

    const suratTugasList = readSheetRows(SHEET_NAMES.SURAT_TUGAS).reverse().map(function(r) {
      return {
        id: r[0], tanggalTerbit: r[1], nomorSurat: r[2], nama_ketua: r[3],
        nik_ketua: r[4], jabatan_ketua: r[5], ttd_ketua: r[6], nikPetugas: r[7],
        nama_petugas: r[8], peran_tugas: r[9], nama_kegiatan: r[10],
        tanggal: r[11], waktu: r[12], lokasi: r[13], fileUrl: r[14]
      };
    });

    const sertifikatList = readSheetRows(SHEET_NAMES.SERTIFIKAT).reverse().map(function(r) {
      return {
        id: r[0], tanggalTerbit: r[1], sertifikat: r[2], nikPetugas: r[3],
        nama_petugas: r[4], partisipasi: r[5], kegiatan: r[6],
        ketua_panitia: r[7], ttd_panitia: r[8], fileUrl: r[9]
      };
    });

    const presensiList = readSheetRows(SHEET_NAMES.PRESENSI).reverse().map(function(r) {
      return {
        id: r[0], tanggal: r[1], nik: r[2], namaPetugas: r[3],
        namaKegiatan: r[4], lokasi: r[5], fotoBukti: r[6], status: r[7]
      };
    });

    const setRows = readSheetRows(SHEET_NAMES.PENGATURAN);
    let rekeningAIK = { bank: 'BSI (Bank Syariah Indonesia)', noRekening: '4337146830', atasNama: 'Latifatul Afifah' };
    let heroBannerUrl = 'https://placehold.co/960x360/D7EEF7/006837?text=GEDUNG+RS+MUHAMMADIYAH+LAMONGAN';
    let defaultKetuaNama = 'KHOMARUDIN, Amd.RMIK.';

    setRows.forEach(function(r) {
      if (r[0] === 'REK_BANK') rekeningAIK.bank = r[1];
      if (r[0] === 'REK_NOMOR') rekeningAIK.noRekening = r[1];
      if (r[0] === 'REK_ATASNAMA') rekeningAIK.atasNama = r[1];
      if (r[0] === 'HERO_BANNER_URL' && r[1]) heroBannerUrl = r[1];
      if (r[0] === 'DEFAULT_KETUA_NAMA' && r[1]) defaultKetuaNama = r[1];
    });

    const result = {
      users: users,
      jenisPimpinanList: jenisPimpinanList,
      masterCabang: masterCabang,
      masterTtd: masterTtd,
      pesanList: pesanList,
      aikList: aikList,
      suratTugasList: suratTugasList,
      sertifikatList: sertifikatList,
      presensiList: presensiList,
      rekeningAIK: rekeningAIK,
      heroBannerUrl: heroBannerUrl,
      defaultKetuaNama: defaultKetuaNama
    };

    try {
      const payloadStr = JSON.stringify(result);
      if (payloadStr.length < 95000) {
        cache.put(APP_DATA_CACHE_KEY, payloadStr, 120);
      }
    } catch (e) {}

    return result;
  } catch (err) {
    return { error: err.message };
  }
}

/**
 * Daftar Akun Baru
 */
function registerUser(newUser) {
  try {
    const ss = getActiveDb();
    const sheet = ss.getSheetByName(SHEET_NAMES.USERS);
    const rows = sheet.getDataRange().getDisplayValues();

    for (let i = 1; i < rows.length; i++) {
      if (rows[i][0] === newUser.username || rows[i][4] === newUser.nik) {
        return { error: 'Username sudah terdaftar!' };
      }
    }

    const cleanPhone = formatPhoneAsText(newUser.noTelp);
    sheet.appendRow([
      String(newUser.username || ''), String(newUser.password || ''), 'petugas', String(newUser.namaLengkap || ''),
      String(newUser.nik || ''), String(newUser.unitKerja || ''), cleanPhone, String(newUser.email || ''),
      String(newUser.forkomDomisili || ''), String(newUser.fotoUrl || '')
    ]);
    const lastRow = sheet.getLastRow();
    sheet.getRange(lastRow, 5).setNumberFormat('@');
    sheet.getRange(lastRow, 7).setNumberFormat('@').setValue(cleanPhone);
    invalidateAppCache();
    return { ok: true };
  } catch (err) {
    return { error: err.message };
  }
}

/**
 * Hapus Akun User Tidak Aktif (Admin)
 */
function hapusUserAdmin(payload) {
  try {
    const ss = getActiveDb();
    const sheet = ss.getSheetByName(SHEET_NAMES.USERS);
    const rows = sheet.getDataRange().getDisplayValues();
    const targetUsername = String(payload.username || '').trim();
    for (let i = 1; i < rows.length; i++) {
      if (String(rows[i][0]).trim() === targetUsername) {
        sheet.deleteRow(i + 1);
        invalidateAppCache();
        return { ok: true };
      }
    }
    return { error: 'User tidak ditemukan!' };
  } catch (err) {
    return { error: err.message };
  }
}

/**
 * Kirim Kode OTP 6 Digit Reset Password via Email
 */
function sendPasswordResetOtp(payload) {
  try {
    const ss = getActiveDb();
    const email = String(payload.email || '').trim().toLowerCase();
    const userRows = ss.getSheetByName(SHEET_NAMES.USERS).getDataRange().getDisplayValues();
    let found = false;

    for (let i = 1; i < userRows.length; i++) {
      if (String(userRows[i][7]).trim().toLowerCase() === email) {
        found = true;
        break;
      }
    }
    if (!found) {
      return { error: 'Email tidak terdaftar dalam sistem!' };
    }

    const otp = Math.floor(100000 + Math.random() * 900000).toString();
    const otpSheet = ss.getSheetByName(SHEET_NAMES.OTP_RESET);
    otpSheet.appendRow([email, otp, new Date().toISOString()]);

    MailApp.sendEmail({
      to: email,
      subject: 'Kode OTP Reset Password - FORKOM RSML',
      body: 'Assalamualaikum Wr. Wb.\n\nKode OTP 6 digit pemulihan password akun FORKOM RSML Anda adalah: ' + otp + '\n\nJangan bagikan kode ini kepada siapa pun.'
    });

    return { ok: true };
  } catch (err) {
    return { error: err.message };
  }
}

/**
 * Verifikasi OTP & Simpan Password Baru
 */
function verifyOtpAndResetPassword(payload) {
  try {
    const ss = getActiveDb();
    const email = String(payload.email || '').trim().toLowerCase();
    const otpInput = String(payload.otp || '').trim();
    const newPassword = String(payload.newPassword || '').trim();

    const otpRows = ss.getSheetByName(SHEET_NAMES.OTP_RESET).getDataRange().getDisplayValues();
    let validOtp = false;
    for (let i = otpRows.length - 1; i >= 1; i--) {
      if (String(otpRows[i][0]).trim().toLowerCase() === email && String(otpRows[i][1]).trim() === otpInput) {
        validOtp = true;
        break;
      }
    }

    if (!validOtp) {
      return { error: 'Kode OTP 6 digit tidak valid!' };
    }

    const userSheet = ss.getSheetByName(SHEET_NAMES.USERS);
    const userRows = userSheet.getDataRange().getDisplayValues();
    for (let j = 1; j < userRows.length; j++) {
      if (String(userRows[j][7]).trim().toLowerCase() === email) {
        userSheet.getRange(j + 1, 2).setValue(newPassword);
        break;
      }
    }

    invalidateAppCache();
    return { ok: true, verified: true };
  } catch (err) {
    return { error: err.message };
  }
}

/**
 * Perbarui Profil & Foto Avatar Petugas
 */
function updateProfilPetugas(updated) {
  try {
    const ss = getActiveDb();
    const sheet = ss.getSheetByName(SHEET_NAMES.USERS);
    const rows = sheet.getDataRange().getDisplayValues();
    let fotoUrl = updated.fotoUrl;

    if (fotoUrl && fotoUrl.indexOf('data:image') === 0) {
      fotoUrl = saveBase64ToDrive(fotoUrl, 'Avatar_' + updated.nik + '_' + Date.now() + '.jpg', DRIVE_FOLDERS.UMUM);
    }

    const cleanPhone = formatPhoneAsText(updated.noTelp);
    for (let i = 1; i < rows.length; i++) {
      if (rows[i][0] === updated.username) {
        const targetRange = sheet.getRange(i + 1, 4, 1, 7);
        sheet.getRange(i + 1, 5).setNumberFormat('@');
        sheet.getRange(i + 1, 7).setNumberFormat('@');
        targetRange.setValues([[
          String(updated.namaLengkap || ''), String(updated.nik || ''), String(updated.unitKerja || ''),
          cleanPhone, String(updated.email || ''), String(updated.forkomDomisili || ''), String(fotoUrl || '')
        ]]);
        break;
      }
    }
    invalidateAppCache();
    return { ok: true, fotoUrl: fotoUrl, noTelp: cleanPhone };
  } catch (err) {
    return { error: err.message };
  }
}

/**
 * Kirim Pengumuman Admin ke Sheet Pesan (Khusus Pengumuman)
 */
function kirimPengumumanAdmin(pesan) {
  try {
    const ss = getActiveDb();
    ss.getSheetByName(SHEET_NAMES.PESAN).appendRow([
      pesan.id, pesan.tanggal, pesan.tujuan, pesan.judul, pesan.isi, pesan.fileUrl || ''
    ]);
    invalidateAppCache();
    return { ok: true };
  } catch (err) {
    return { error: err.message };
  }
}

/**
 * Simpan Pengajuan Verifikasi AIK Petugas
 */
function submitAIKForm(aik) {
  try {
    const ss = getActiveDb();
    let buktiUrl = aik.buktiUrl;
    if (buktiUrl && buktiUrl.indexOf('data:') === 0) {
      const isPdf = buktiUrl.indexOf('application/pdf') !== -1;
      const ext = isPdf ? 'pdf' : 'jpg';
      buktiUrl = saveBase64ToDrive(buktiUrl, 'BuktiAIK_' + aik.nik + '_' + Date.now() + '.' + ext, DRIVE_FOLDERS.UMUM);
    }

    ss.getSheetByName(SHEET_NAMES.PENGAJUAN_AIK).appendRow([
      aik.id, aik.tanggal, aik.nik, aik.nama, aik.unitKerja, aik.email,
      aik.domisili, aik.periode, aik.jenisPimpinan, aik.areaCabang,
      aik.namaPimpinan, aik.a1, aik.b1, aik.c1, buktiUrl, 'Pending', '', ''
    ]);
    invalidateAppCache();
    return { ok: true };
  } catch (err) {
    return { error: err.message };
  }
}

/**
 * Tolak Pengajuan AIK & Simpan Keterangan Penolakan (Tidak dikirim ke Pesan)
 */
function tolakAIKSubmission(payload) {
  try {
    const ss = getActiveDb();
    const aikSheet = ss.getSheetByName(SHEET_NAMES.PENGAJUAN_AIK);
    const rows = aikSheet.getDataRange().getDisplayValues();
    for (let i = 1; i < rows.length; i++) {
      if (String(rows[i][0]) === String(payload.id)) {
        aikSheet.getRange(i + 1, 16, 1, 3).setValues([['Ditolak', '', String(payload.alasanPenolakan || 'Ditolak Admin')]]);
        break;
      }
    }
    invalidateAppCache();
    return { ok: true };
  } catch (err) {
    return { error: err.message };
  }
}

/**
 * Approve AIK & Otomatis Isi Template Google Docs (Status jadi Approved, TIDAK dikirim ke menu Pesan)
 */
function approveAIKAndGenerateDoc(payload) {
  try {
    const ss = getActiveDb();
    const aik = payload.aikData;
    const templateId = payload.templateId || '1an5T_aw-L1AMfUpxm0UMOSs8i1RdTI9DvGcltCq3eQU';

    let docUrl = 'https://docs.google.com/document/d/' + templateId + '/edit';
    try {
      const folder = getCachedDriveFolder(DRIVE_FOLDERS.UMUM, null);

      const templateFile = DriveApp.getFileById(templateId);
      const copyFile = templateFile.makeCopy('Verifikasi_AIK_' + aik.nama + '_' + aik.periode, folder);
      copyFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

      const doc = DocumentApp.openById(copyFile.getId());
      const body = doc.getBody();
      const tglSurat = Utilities.formatDate(new Date(), 'Asia/Jakarta', 'dd MMMM yyyy');

      body.replaceText('\uFDFD', 'بِسْمِ اللَّهِ الرَّحْمَنِ الرَّحِيم');

      body.replaceText('\\{\\{tanggal\\}\\}', tglSurat);
      body.replaceText('\\{\\{nama\\}\\}', String(aik.nama || ''));
      body.replaceText('\\{\\{nik\\}\\}', String(aik.nik || ''));
      body.replaceText('\\{\\{unit_kerja\\}\\}', String(aik.unitKerja || ''));
      body.replaceText('\\{\\{domisili\\}\\}', String(aik.domisili || ''));
      body.replaceText('\\{\\{periode\\}\\}', String(aik.periode || '2026'));

      body.replaceText('\\{\\{a1\\}\\}', aik.a1 === '>5' ? '☑' : '☐');
      body.replaceText('\\{\\{a2\\}\\}', aik.a1 === '4-5' ? '☑' : '☐');
      body.replaceText('\\{\\{a3\\}\\}', aik.a1 === '<4' ? '☑' : '☐');

      body.replaceText('\\{\\{b1\\}\\}', aik.b1 === '>12' ? '☑' : '☐');
      body.replaceText('\\{\\{b2\\}\\}', aik.b1 === '6-12' ? '☑' : '☐');
      body.replaceText('\\{\\{b3\\}\\}', aik.b1 === '<6' ? '☑' : '☐');

      body.replaceText('\\{\\{c1\\}\\}', aik.c1 === '>4' ? '☑' : '☐');
      body.replaceText('\\{\\{c2\\}\\}', aik.c1 === '2-4' ? '☑' : '☐');
      body.replaceText('\\{\\{c3\\}\\}', aik.c1 === '<2' ? '☑' : '☐');

      body.replaceText('\\{\\{jenis_pimpinan\\}\\}', String(aik.jenisPimpinan || ''));
      body.replaceText('\\{\\{area_cabang\\}\\}', String(aik.areaCabang || ''));
      body.replaceText('\\{\\{nama_pimpinan\\}\\}', String(aik.namaPimpinan || ''));
      body.replaceText('\\{\\{ketua_forkom\\}\\}', 'KHOMARUDIN, Amd.RMIK.');

      doc.saveAndClose();

      const pdfBlob = copyFile.getAs(MimeType.PDF).setName('Verifikasi_AIK_' + aik.nama + '_' + aik.periode + '.pdf');
      const pdfFile = folder.createFile(pdfBlob);
      pdfFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
      copyFile.setTrashed(true);
      docUrl = pdfFile.getUrl();
    } catch (docErr) {
      Logger.log('Fallback template AIK: ' + docErr.message);
    }

    const aikSheet = ss.getSheetByName(SHEET_NAMES.PENGAJUAN_AIK);
    const rows = aikSheet.getDataRange().getDisplayValues();
    for (let i = 1; i < rows.length; i++) {
      if (String(rows[i][0]) === String(aik.id)) {
        aikSheet.getRange(i + 1, 16, 1, 3).setValues([['Approved', docUrl, '']]);
        break;
      }
    }

    invalidateAppCache();
    return { ok: true, docUrl: docUrl };
  } catch (err) {
    return { error: err.message };
  }
}

/**
 * Helper internal: Isi Template Surat Tugas & Simpan PDF ke Folder Khusus Surat Tugas per Tahun
 */
function generateSuratTugasPdfFile(st, templateId, filePrefix) {
  const ss = getActiveDb();
  let unitKerjaPetugas = '-';
  const userRows = ss.getSheetByName(SHEET_NAMES.USERS).getDataRange().getDisplayValues();
  for (let i = 1; i < userRows.length; i++) {
    if (String(userRows[i][4]).trim() === String(st.nikPetugas).trim()) {
      unitKerjaPetugas = userRows[i][5] || '-';
      break;
    }
  }

  // Ambil TTD otomatis dari MasterTTD jika tidak dikirim langsung
  let ttdSource = st.ttd_ketua || '';
  if (!ttdSource && ss.getSheetByName(SHEET_NAMES.MASTER_TTD)) {
    const ttdRows = ss.getSheetByName(SHEET_NAMES.MASTER_TTD).getDataRange().getDisplayValues();
    for (let t = 1; t < ttdRows.length; t++) {
      if (String(ttdRows[t][0]).trim().toLowerCase() === String(st.nama_ketua || '').trim().toLowerCase()) {
        ttdSource = ttdRows[t][1];
        break;
      }
    }
  }
  const ttdBlob = resolveSignatureBlob(ttdSource, 'TTD_ST');

  // Folder khusus Surat Tugas dipisah berdasarkan tahun (misal: FORKOM_RSML_Surat_Tugas/2026)
  const targetYear = extractTargetYear(st.tanggal || st.nomorSurat);
  const folder = getCachedDriveFolder(DRIVE_FOLDERS.SURAT_TUGAS, targetYear);

  const templateFile = DriveApp.getFileById(templateId);
  const baseName = filePrefix + '_' + st.nama_petugas + '_' + String(st.nomorSurat || '').replace(/\//g, '-');
  const copyFile = templateFile.makeCopy(baseName, folder);
  copyFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

  const doc = DocumentApp.openById(copyFile.getId());
  const body = doc.getBody();
  const tglTerbit = Utilities.formatDate(new Date(), 'Asia/Jakarta', 'dd MMMM yyyy');

  body.replaceText('\\{\\{nomor_surat\\}\\}', String(st.nomorSurat || ''));
  body.replaceText('\\{\\{nama_ketua\\}\\}', String(st.nama_ketua || ''));
  body.replaceText('\\{\\{nik_ketua\\}\\}', String(st.nik_ketua || ''));
  body.replaceText('\\{\\{jabatan_ketua\\}\\}', String(st.jabatan_ketua || ''));
  body.replaceText('\\{\\{nama_petugas\\}\\}', String(st.nama_petugas || ''));
  body.replaceText('\\{\\{nik_petugas\\}\\}', String(st.nikPetugas || ''));
  body.replaceText('\\{\\{unit_kerja\\}\\}', String(unitKerjaPetugas));
  body.replaceText('\\{\\{peran_tugas\\}\\}', String(st.peran_tugas || ''));
  body.replaceText('\\{\\{nama_kegiatan\\}\\}', String(st.nama_kegiatan || ''));
  body.replaceText('\\{\\{tanggal\\}\\}', String(st.tanggal || ''));
  body.replaceText('\\{\\{waktu\\}\\}', String(st.waktu || ''));
  body.replaceText('\\{\\{lokasi\\}\\}', String(st.lokasi || ''));
  body.replaceText('\\{\\{tanggal_terbit\\}\\}', tglTerbit);

  const foundTtd = body.findText('\\{\\{ttd_ketua\\}\\}');
  if (foundTtd) {
    const el = foundTtd.getElement();
    const parent = el.getParent();
    el.asText().setText('');
    if (ttdBlob && parent.insertInlineImage) {
      const img = parent.insertInlineImage(0, ttdBlob);
      img.setWidth(110);
      img.setHeight(55);
    }
  }

  doc.saveAndClose();

  const pdfBlob = copyFile.getAs(MimeType.PDF).setName(baseName + '.pdf');
  const pdfFile = folder.createFile(pdfBlob);
  pdfFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  copyFile.setTrashed(true);
  return { pdfFile: pdfFile, ttdUrl: ttdSource };
}

/**
 * Buat Pratinjau PDF Surat Tugas
 */
function previewSuratTugasGoogleDoc(payload) {
  try {
    const st = payload.suratTugas;
    const templateId = payload.templateId || '1O_vqE-DJPdxxvhbuY-JUK_hJo_um9wPZTbN49cIrBHQ';
    const result = generateSuratTugasPdfFile(st, templateId, 'PREVIEW_ST_' + Date.now());
    return { ok: true, docUrl: result.pdfFile.getUrl(), docId: result.pdfFile.getId() };
  } catch (err) {
    return { error: 'Gagal membuat pratinjau Surat Tugas: ' + err.message };
  }
}

/**
 * Terbitkan Surat Tugas (Disimpan ke folder FORKOM_RSML_Surat_Tugas/{Tahun})
 */
function terbitkanSuratTugas(payload) {
  try {
    const ss = getActiveDb();
    const st = payload.suratTugas;
    const templateId = payload.templateId || '1O_vqE-DJPdxxvhbuY-JUK_hJo_um9wPZTbN49cIrBHQ';

    let docUrl = 'https://docs.google.com/document/d/' + templateId + '/edit';
    let ttdUrl = st.ttd_ketua || '';
    try {
      const result = generateSuratTugasPdfFile(st, templateId, 'Surat_Tugas');
      docUrl = result.pdfFile.getUrl();
      ttdUrl = result.ttdUrl || ttdUrl;
    } catch (docErr) {
      Logger.log('Fallback template ST: ' + docErr.message);
    }

    ss.getSheetByName(SHEET_NAMES.SURAT_TUGAS).appendRow([
      st.id, st.tanggalTerbit, st.nomorSurat, st.nama_ketua, st.nik_ketua,
      st.jabatan_ketua, ttdUrl, st.nikPetugas, st.nama_petugas, st.peran_tugas,
      st.nama_kegiatan, st.tanggal, st.waktu, st.lokasi, docUrl
    ]);
    invalidateAppCache();
    return { ok: true, docUrl: docUrl };
  } catch (err) {
    return { error: err.message };
  }
}

/**
 * Jalankan fungsi ini 1x di editor Apps Script jika ingin mengotorisasi izin Google Slides
 */
function otorisasiSlides() {
  SlidesApp.getActivePresentation();
  return 'Izin Google Slides berhasil diaktifkan!';
}

/**
 * Helper internal: Isi Template Sertifikat & Simpan PDF ke Folder Khusus Sertifikat per Tahun
 */
function generateSertifikatPdfFile(s, templateId, filePrefix) {
  const ss = getActiveDb();
  const namaKetuaVal = String(s.nama_ketua || s.ketua_panitia || 'KHOMARUDIN, Amd.RMIK.');

  // Ambil TTD otomatis dari MasterTTD jika tidak dikirim langsung
  let rawTtd = s.ttd_ketua || s.ttd_panitia || '';
  if (!rawTtd && ss.getSheetByName(SHEET_NAMES.MASTER_TTD)) {
    const ttdRows = ss.getSheetByName(SHEET_NAMES.MASTER_TTD).getDataRange().getDisplayValues();
    for (let t = 1; t < ttdRows.length; t++) {
      if (String(ttdRows[t][0]).trim().toLowerCase() === namaKetuaVal.trim().toLowerCase()) {
        rawTtd = ttdRows[t][1];
        break;
      }
    }
  }
  const ttdBlob = resolveSignatureBlob(rawTtd, 'TTD_SERT');

  // Folder khusus Sertifikat dipisah berdasarkan tahun (misal: FORKOM_RSML_Sertifikat/2026)
  const targetYear = extractTargetYear(s.sertifikat || s.tanggalTerbit);
  const folder = getCachedDriveFolder(DRIVE_FOLDERS.SERTIFIKAT, targetYear);

  const templateFile = DriveApp.getFileById(templateId);
  const mimeType = templateFile.getMimeType();
  const baseName = filePrefix + '_' + s.nama_petugas + '_' + String(s.sertifikat || '').replace(/\//g, '-');
  const copyFile = templateFile.makeCopy(baseName, folder);
  copyFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

  const tglTerbit = Utilities.formatDate(new Date(), 'Asia/Jakarta', 'dd MMMM yyyy');

  if (mimeType === MimeType.GOOGLE_DOCS) {
    const doc = DocumentApp.openById(copyFile.getId());
    const body = doc.getBody();
    const sections = [body];
    if (doc.getHeader()) sections.push(doc.getHeader());
    if (doc.getFooter()) sections.push(doc.getFooter());

    sections.forEach(function(sec) {
      sec.replaceText('\\{\\{\\s*nomor_sertifikat\\s*\\}\\}', String(s.sertifikat || ''));
      sec.replaceText('\\{\\{\\s*sertifikat\\s*\\}\\}', String(s.sertifikat || ''));
      sec.replaceText('\\{\\{\\s*nama_petugas\\s*\\}\\}', String(s.nama_petugas || ''));
      sec.replaceText('\\{\\{\\s*nama\\s*\\}\\}', String(s.nama_petugas || ''));
      sec.replaceText('\\{\\{\\s*partisipasi\\s*\\}\\}', String(s.partisipasi || ''));
      sec.replaceText('\\{\\{\\s*kegiatan\\s*\\}\\}', String(s.kegiatan || ''));
      sec.replaceText('\\{\\{\\s*nama_kegiatan\\s*\\}\\}', String(s.kegiatan || ''));
      sec.replaceText('\\{\\{\\s*nama_ketua\\s*\\}\\}', namaKetuaVal);
      sec.replaceText('\\{\\{\\s*ketua_panitia\\s*\\}\\}', namaKetuaVal);
      sec.replaceText('\\{\\{\\s*ketua\\s*\\}\\}', namaKetuaVal);
      sec.replaceText('\\{\\{\\s*tanggal_terbit\\s*\\}\\}', tglTerbit);
      sec.replaceText('\\{\\{\\s*tanggal\\s*\\}\\}', tglTerbit);

      const ttdTags = ['\\{\\{\\s*ttd_ketua\\s*\\}\\}', '\\{\\{\\s*ttd_panitia\\s*\\}\\}', '\\{\\{\\s*ttd\\s*\\}\\}'];
      ttdTags.forEach(function(tagPattern) {
        let foundTtd = sec.findText(tagPattern);
        while (foundTtd) {
          const el = foundTtd.getElement();
          const parent = el.getParent();
          el.asText().setText('');
          if (ttdBlob && parent.insertInlineImage) {
            const img = parent.insertInlineImage(0, ttdBlob);
            img.setWidth(115);
            img.setHeight(58);
          }
          foundTtd = sec.findText(tagPattern);
        }
      });
    });

    doc.saveAndClose();
  } else {
    const presentation = SlidesApp.openById(copyFile.getId());
    presentation.replaceAllText('{{nomor_sertifikat}}', String(s.sertifikat || ''));
    presentation.replaceAllText('{{sertifikat}}', String(s.sertifikat || ''));
    presentation.replaceAllText('{{nama_petugas}}', String(s.nama_petugas || ''));
    presentation.replaceAllText('{{nama}}', String(s.nama_petugas || ''));
    presentation.replaceAllText('{{partisipasi}}', String(s.partisipasi || ''));
    presentation.replaceAllText('{{kegiatan}}', String(s.kegiatan || ''));
    presentation.replaceAllText('{{nama_kegiatan}}', String(s.kegiatan || ''));
    presentation.replaceAllText('{{nama_ketua}}', namaKetuaVal);
    presentation.replaceAllText('{{ketua_panitia}}', namaKetuaVal);
    presentation.replaceAllText('{{ketua}}', namaKetuaVal);
    presentation.replaceAllText('{{tanggal_terbit}}', tglTerbit);
    presentation.replaceAllText('{{tanggal}}', tglTerbit);

    const slides = presentation.getSlides();
    slides.forEach(function(slide) {
      slide.getShapes().forEach(function(shape) {
        const txt = shape.getText().asString();
        if (txt.indexOf('{{ttd_ketua}}') !== -1 || txt.indexOf('{{ttd_panitia}}') !== -1 || txt.indexOf('{{ttd}}') !== -1) {
          shape.getText().setText('');
          if (ttdBlob) {
            slide.insertImage(ttdBlob, shape.getLeft(), shape.getTop(), shape.getWidth() || 120, shape.getHeight() || 60);
          }
        }
      });
    });
    presentation.saveAndClose();
  }

  const pdfBlob = copyFile.getAs(MimeType.PDF).setName(baseName + '.pdf');
  const pdfFile = folder.createFile(pdfBlob);
  pdfFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  copyFile.setTrashed(true);
  return { pdfFile: pdfFile, ttdUrl: rawTtd };
}

/**
 * Buat Pratinjau Sertifikat & Konversi ke PDF
 */
function previewSertifikatGoogleSlide(payload) {
  try {
    const s = payload.sertifikat;
    const templateId = payload.templateId || '1K4tut7fT3lXsrbiMNyyUI1tnjjlyuSxEj5xV83gGE_U';
    const res = generateSertifikatPdfFile(s, templateId, 'PREVIEW_SERT_' + Date.now());
    return { ok: true, fileUrl: res.pdfFile.getUrl(), slideId: res.pdfFile.getId() };
  } catch (err) {
    return { error: 'Gagal membuat pratinjau PDF Sertifikat: ' + err.message };
  }
}

/**
 * Terbitkan Sertifikat (Disimpan ke folder FORKOM_RSML_Sertifikat/{Tahun})
 */
function terbitkanSertifikat(payload) {
  try {
    const ss = getActiveDb();
    const s = payload.sertifikat;
    const templateId = payload.templateId || '1K4tut7fT3lXsrbiMNyyUI1tnjjlyuSxEj5xV83gGE_U';
    const namaKetuaVal = s.nama_ketua || s.ketua_panitia || 'KHOMARUDIN, Amd.RMIK.';

    let finalPdfUrl = '';
    let ttdUrl = s.ttd_ketua || s.ttd_panitia || '';
    try {
      const res = generateSertifikatPdfFile(s, templateId, 'Sertifikat');
      finalPdfUrl = res.pdfFile.getUrl();
      ttdUrl = res.ttdUrl || ttdUrl;
    } catch (docErr) {
      return { error: 'Gagal membuat PDF Sertifikat: ' + docErr.message };
    }

    ss.getSheetByName(SHEET_NAMES.SERTIFIKAT).appendRow([
      s.id, s.tanggalTerbit, s.sertifikat, s.nikPetugas, s.nama_petugas,
      s.partisipasi, s.kegiatan, namaKetuaVal, ttdUrl, finalPdfUrl
    ]);

    invalidateAppCache();
    return { ok: true, fileUrl: finalPdfUrl };
  } catch (err) {
    return { error: err.message };
  }
}

/**
 * Simpan Presensi Kegiatan Petugas
 */
function submitPresensiKegiatan(pr) {
  try {
    const ss = getActiveDb();
    let fotoUrl = pr.fotoBukti;
    if (fotoUrl && fotoUrl.indexOf('data:image') === 0) {
      fotoUrl = saveBase64ToDrive(fotoUrl, 'Presensi_' + pr.nik + '_' + Date.now() + '.jpg', DRIVE_FOLDERS.UMUM);
    }

    ss.getSheetByName(SHEET_NAMES.PRESENSI).appendRow([
      pr.id, pr.tanggal, pr.nik, pr.namaPetugas, pr.namaKegiatan, pr.lokasi, fotoUrl, pr.status
    ]);
    invalidateAppCache();
    return { ok: true, fotoUrl: fotoUrl };
  } catch (err) {
    return { error: err.message };
  }
}

/**
 * Kelola Master Jenis Pimpinan, Master Cabang, Rekening AIK, Hero Banner & Ketua Default
 */
function tambahJenisPimpinan(item) {
  try {
    const ss = getActiveDb();
    let sheet = ss.getSheetByName(SHEET_NAMES.JENIS_PIMPINAN);
    if (!sheet) {
      setupDatabase();
      sheet = ss.getSheetByName(SHEET_NAMES.JENIS_PIMPINAN);
    }
    const kodeClean = String(item.kode || '').trim().toUpperCase();
    const labelClean = String(item.label || kodeClean).trim();
    sheet.appendRow([kodeClean, labelClean]);
    invalidateAppCache();
    return { ok: true };
  } catch (err) {
    return { error: err.message };
  }
}

function hapusJenisPimpinan(payload) {
  try {
    const ss = getActiveDb();
    const sheet = ss.getSheetByName(SHEET_NAMES.JENIS_PIMPINAN);
    if (!sheet || sheet.getLastRow() < 2) {
      return { error: 'Data jenis pimpinan tidak ditemukan!' };
    }
    const targetKode = String(payload.kode || '').trim().toUpperCase();
    const rows = sheet.getRange(2, 1, sheet.getLastRow() - 1, 2).getDisplayValues();
    for (let i = 0; i < rows.length; i++) {
      if (String(rows[i][0] || '').trim().toUpperCase() === targetKode) {
        sheet.deleteRow(i + 2);
        invalidateAppCache();
        return { ok: true };
      }
    }
    return { error: 'Kode jenis pimpinan tidak ditemukan!' };
  } catch (err) {
    return { error: err.message };
  }
}

function tambahMasterCabang(cabang) {
  try {
    const ss = getActiveDb();
    const namaClean = String(cabang.nama || '').trim();
    const jenisClean = String(cabang.jenis || '').trim().toUpperCase();
    ss.getSheetByName(SHEET_NAMES.MASTER_CABANG).appendRow([cabang.id, jenisClean, namaClean]);
    invalidateAppCache();
    return { ok: true };
  } catch (err) {
    return { error: err.message };
  }
}

function hapusMasterCabang(payload) {
  try {
    const ss = getActiveDb();
    const sheet = ss.getSheetByName(SHEET_NAMES.MASTER_CABANG);
    const rows = sheet.getRange(2, 1, sheet.getLastRow() - 1, 3).getValues();
    for (let i = 0; i < rows.length; i++) {
      if (String(rows[i][0]) === String(payload.id)) {
        sheet.deleteRow(i + 2);
        break;
      }
    }
    invalidateAppCache();
    return { ok: true };
  } catch (err) {
    return { error: err.message };
  }
}

function updateRekeningAIK(rek) {
  try {
    upsertSettingKey('REK_BANK', rek.bank);
    upsertSettingKey('REK_NOMOR', rek.noRekening);
    upsertSettingKey('REK_ATASNAMA', rek.atasNama);
    invalidateAppCache();
    return { ok: true };
  } catch (err) {
    return { error: err.message };
  }
}

function updateHeroBanner(payload) {
  try {
    let bannerUrl = payload.bannerData || '';
    if (bannerUrl && bannerUrl.indexOf('data:image') === 0) {
      bannerUrl = saveBase64ToDrive(bannerUrl, 'HeroBanner_' + Date.now() + '.jpg', DRIVE_FOLDERS.UMUM);
    }
    upsertSettingKey('HERO_BANNER_URL', bannerUrl);
    invalidateAppCache();
    return { ok: true, bannerUrl: bannerUrl };
  } catch (err) {
    return { error: err.message };
  }
}

function updateDefaultKetua(payload) {
  try {
    upsertSettingKey('DEFAULT_KETUA_NAMA', String(payload.defaultKetuaNama || ''));
    invalidateAppCache();
    return { ok: true };
  } catch (err) {
    return { error: err.message };
  }
}

/**
 * Helper Upload File Base64 ke Google Drive
 */
function saveBase64ToDrive(base64DataUri, fileName, targetRootFolder) {
  try {
    const parts = base64DataUri.split(',');
    const meta = parts[0];
    const base64Str = parts[1];
    const contentType = meta.match(/:(.*?);/)[1];
    const blob = Utilities.newBlob(Utilities.base64Decode(base64Str), contentType, fileName);

    const folderName = targetRootFolder || DRIVE_FOLDERS.UMUM;
    const folder = getCachedDriveFolder(folderName, null);

    const file = folder.createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    if (contentType && contentType.indexOf('application/pdf') !== -1) {
      return 'https://drive.google.com/file/d/' + file.getId() + '/view';
    }
    return 'https://lh3.googleusercontent.com/d/' + file.getId();
  } catch (e) {
    return base64DataUri;
  }
}
