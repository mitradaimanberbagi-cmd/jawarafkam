/**
 * ============================================================================
 * MODUL ENTRY & MANAJEMEN LAPORAN KEGIATAN (Kegiatan.gs)
 * ============================================================================
 */

const KEGIATAN_CONFIG = {
  SHEET_KEGIATAN: "LAPORAN_AKTIVITAS",
  SHEET_KATEGORI: "MASTER_KATEGORI",
  DRIVE_FOLDER_NAME: "JAWARA_FOTO_KEGIATAN"
};

/**
 * Helper: Mengekstrak Google Drive File ID murni (baik dari ID langsung maupun URL Drive)
 */
function extractDriveFileId_(rawValue) {
  const str = String(rawValue || "").trim();
  if (!str) return "";
  const match = str.match(/[-\w]{25,}/);
  return match ? match[0] : str;
}

/**
 * Mengambil opsi dropdown Kategori Kegiatan untuk form Entry Laporan Kegiatan
 */
function getActivityFormOptions(token) {
  const user = validateSession(token);
  if (!user || !user.user_id) {
    throw new Error("Sesi tidak valid atau telah berakhir. Silakan login kembali.");
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(KEGIATAN_CONFIG.SHEET_KATEGORI);

  if (!sheet) {
    return {
      success: true,
      kategori: []
    };
  }

  const values = sheet.getDataRange().getValues();
  if (values.length <= 1) {
    return { success: true, kategori: [] };
  }

  const headers = values[0].map(h => String(h).trim().toLowerCase());
  const idxId = headers.indexOf("kategori_id") !== -1 ? headers.indexOf("kategori_id") : 0;
  const idxNama = headers.indexOf("nama_kategori") !== -1
    ? headers.indexOf("nama_kategori")
    : (headers.indexOf("nama") !== -1 ? headers.indexOf("nama") : 1);
  const idxStatus = headers.indexOf("status");

  const kategori = [];
  for (let i = 1; i < values.length; i++) {
    const row = values[i];
    const id = String(row[idxId] || "").trim();
    const nama = String(row[idxNama] || "").trim();
    if (!id || !nama) continue;

    if (idxStatus !== -1) {
      const status = String(row[idxStatus] || "").trim().toUpperCase();
      if (status && status !== "AKTIF") continue;
    }

    kategori.push({ id: id, nama: nama });
  }

  return {
    success: true,
    kategori: kategori
  };
}

/**
 * Menyimpan Entry Laporan Kegiatan baru + Upload 2-5 Foto ke Google Drive
 */
function createActivityEntry(token, payload) {
  const user = validateSession(token);
  if (!user || !user.user_id) {
    throw new Error("Sesi tidak valid. Silakan login kembali.");
  }
  if (user.role !== "USER") {
    throw new Error("Hanya akun Da'i (USER) yang dapat menginput laporan kegiatan.");
  }

  payload = payload || {};
  const tanggalKegiatan = String(payload.tanggal_kegiatan || "").trim();
  const kategoriId = String(payload.kategori_id || "").trim();
  const namaPihak = String(payload.nama_pihak || "").trim();
  const noHpKontak = String(payload.no_hp_kontak || "").trim();
  const lokasi = String(payload.lokasi || "").trim();
  const narasiKegiatan = String(payload.narasi_kegiatan || "").trim();
  const tindakLanjut = String(payload.tindak_lanjut || "").trim();
  const tanggalFollowup = String(payload.tanggal_followup || "").trim();
  const fotos = Array.isArray(payload.fotos) ? payload.fotos : [];

  // Validasi kolom wajib
  if (!tanggalKegiatan) throw new Error("Tanggal kegiatan wajib diisi.");
  if (!kategoriId) throw new Error("Kategori kegiatan wajib dipilih.");
  if (!namaPihak) throw new Error("Nama pihak / tokoh / lembaga wajib diisi.");
  if (!lokasi) throw new Error("Lokasi kegiatan wajib diisi.");
  if (!narasiKegiatan) throw new Error("Narasi kegiatan wajib diisi.");

  // Validasi jumlah foto (Minimal 2, Maksimal 5)
  if (fotos.length < 2) {
    throw new Error("Minimal upload 2 foto dokumentasi kegiatan.");
  }
  if (fotos.length > 5) {
    throw new Error("Maksimal upload 5 foto dokumentasi kegiatan.");
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(KEGIATAN_CONFIG.SHEET_KEGIATAN);

  // Buat sheet otomatis dengan susunan kolom persis sesuai spesifikasi jika belum ada
  if (!sheet) {
    sheet = ss.insertSheet(KEGIATAN_CONFIG.SHEET_KEGIATAN);
    sheet.appendRow([
      "laporan_id",
      "tanggal_kegiatan",
      "waktu_input",
      "user_id",
      "kategori_id",
      "nama_pihak",
      "no_hp_kontak",
      "lokasi",
      "narasi_kegiatan",
      "tindak_lanjut",
      "tanggal_followup",
      "foto_1_file_id",
      "foto_2_file_id",
      "foto_3_file_id",
      "foto_4_file_id",
      "foto_5_file_id",
      "jumlah_foto",
      "status_laporan",
      "cancelled_at",
      "cancelled_by",
      "cancel_reason"
    ]);
  }

  const now = new Date();
  const laporanId = generateLaporanId_(now);
  const folder = getOrCreateKegiatanDriveFolder_();
  const uploadedPhotoIds = ["", "", "", "", ""];
  const createdFiles = [];

  try {
    for (let i = 0; i < fotos.length; i++) {
      const foto = fotos[i];
      if (!foto || !foto.base64 || !foto.mimeType) {
        throw new Error("Data foto ke-" + (i + 1) + " tidak valid.");
      }

      const allowedMimes = ["image/jpeg", "image/jpg", "image/png", "image/webp"];
      if (allowedMimes.indexOf(String(foto.mimeType).toLowerCase()) === -1) {
        throw new Error("Format foto ke-" + (i + 1) + " harus JPG, PNG, atau WEBP.");
      }

      const ext = getExtensionFromMime_(foto.mimeType);
      const safeFileName = laporanId + "_FOTO_" + (i + 1) + ext;
      const bytes = Utilities.base64Decode(foto.base64);
      const blob = Utilities.newBlob(bytes, foto.mimeType, safeFileName);
      const file = folder.createFile(blob);

      // Izinkan file dapat dilihat melalui link oleh admin/sistem
      file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

      // Simpan File ID murni sesuai kolom foto_1_file_id .. foto_5_file_id
      uploadedPhotoIds[i] = file.getId();
      createdFiles.push(file);
    }

    // Susunan 21 kolom sesuai urutan sheet LAPORAN_AKTIVITAS
    const newRow = [
      laporanId,              // 1. laporan_id
      tanggalKegiatan,        // 2. tanggal_kegiatan
      now,                    // 3. waktu_input
      user.user_id,           // 4. user_id
      kategoriId,             // 5. kategori_id
      namaPihak,              // 6. nama_pihak
      noHpKontak,             // 7. no_hp_kontak
      lokasi,                 // 8. lokasi
      narasiKegiatan,         // 9. narasi_kegiatan
      tindakLanjut,           // 10. tindak_lanjut
      tanggalFollowup,        // 11. tanggal_followup
      uploadedPhotoIds[0],    // 12. foto_1_file_id
      uploadedPhotoIds[1],    // 13. foto_2_file_id
      uploadedPhotoIds[2],    // 14. foto_3_file_id
      uploadedPhotoIds[3],    // 15. foto_4_file_id
      uploadedPhotoIds[4],    // 16. foto_5_file_id
      fotos.length,           // 17. jumlah_foto
      "TERCATAT",             // 18. status_laporan
      "",                     // 19. cancelled_at
      "",                     // 20. cancelled_by
      ""                      // 21. cancel_reason
    ];

    sheet.appendRow(newRow);

    // Pastikan kolom no_hp_kontak (kolom ke-7) terformat sebagai teks agar angka 0 di depan tidak hilang
    const lastRow = sheet.getLastRow();
    if (noHpKontak) {
      sheet.getRange(lastRow, 7).setNumberFormat("@").setValue(noHpKontak);
    }

    return {
      success: true,
      laporan_id: laporanId,
      jumlah_foto: fotos.length,
      message: "Laporan kegiatan berhasil disimpan."
    };

  } catch (err) {
    // Rollback: hapus file foto yang terlanjur terunggah jika proses penyimpanan gagal
    createdFiles.forEach(function(file) {
      try {
        file.setTrashed(true);
      } catch (e) {}
    });
    throw new Error(err.message || "Gagal menyimpan laporan kegiatan.");
  }
}

/**
 * Helper: Membuat ID Laporan unik (Contoh: KEG-20261124-4821)
 */
function generateLaporanId_(dateObj) {
  const tz = Session.getScriptTimeZone() || "Asia/Jakarta";
  const stamp = Utilities.formatDate(dateObj, tz, "yyyyMMdd");
  const rand = Math.floor(1000 + Math.random() * 9000);
  return "KEG-" + stamp + "-" + rand;
}

/**
 * Helper: Mendapatkan atau membuat folder Drive penyimpanan foto kegiatan
 */
function getOrCreateKegiatanDriveFolder_() {
  const folders = DriveApp.getFoldersByName(KEGIATAN_CONFIG.DRIVE_FOLDER_NAME);
  if (folders.hasNext()) {
    return folders.next();
  }
  return DriveApp.createFolder(KEGIATAN_CONFIG.DRIVE_FOLDER_NAME);
}

/**
 * Helper: Ekstensi file dari MIME type
 */
function getExtensionFromMime_(mimeType) {
  const m = String(mimeType).toLowerCase();
  if (m === "image/png") return ".png";
  if (m === "image/webp") return ".webp";
  return ".jpg";
}

/**
 * ============================================================================
 * RETRIEVAL DATA KEGIATAN (History Kegiatan USER & Data Kegiatan ADMIN)
 * ============================================================================
 */

/**
 * Mengambil riwayat laporan kegiatan milik USER yang sedang login
 */
function getUserActivities(token, filters) {
  const user = validateSession(token);
  if (!user || !user.user_id) {
    throw new Error("Sesi tidak valid. Silakan login kembali.");
  }
  return fetchActivitiesData_(user, filters || {}, false);
}

/**
 * Mengambil seluruh data laporan kegiatan semua Da'i (Khusus ADMIN)
 */
function getAdminActivities(token, filters) {
  const user = validateSession(token);
  if (!user || !user.user_id || user.role !== "ADMIN") {
    throw new Error("Akses ditolak. Khusus administrator.");
  }
  return fetchActivitiesData_(user, filters || {}, true);
}

/**
 * Helper utama membaca dan memfilter data sheet Kegiatan
 */
function fetchActivitiesData_(currentUser, filters, isAdminView) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const tz = Session.getScriptTimeZone() || "Asia/Jakarta";

  // 1. Ambil Master Kategori Kegiatan
  const kategoriMap = {};
  const kategoriList = [];
  const katSheet = ss.getSheetByName(KEGIATAN_CONFIG.SHEET_KATEGORI);
  if (katSheet && katSheet.getLastRow() > 1) {
    const katValues = katSheet.getDataRange().getValues();
    const katHeaders = katValues[0].map(h => String(h).trim().toLowerCase());
    const idxId = katHeaders.indexOf("kategori_id") !== -1 ? katHeaders.indexOf("kategori_id") : 0;
    const idxNama = katHeaders.indexOf("nama_kategori") !== -1
      ? katHeaders.indexOf("nama_kategori")
      : (katHeaders.indexOf("nama") !== -1 ? katHeaders.indexOf("nama") : 1);

    for (let i = 1; i < katValues.length; i++) {
      const id = String(katValues[i][idxId] || "").trim();
      const nama = String(katValues[i][idxNama] || "").trim();
      if (id && nama) {
        kategoriMap[id] = nama;
        kategoriList.push({ id: id, nama: nama });
      }
    }
  }

  // 2. Ambil Master Users (untuk menampilkan nama Da'i pada tampilan Admin)
  const userMap = {};
  if (isAdminView) {
    const userSheet = ss.getSheetByName("Users");
    if (userSheet && userSheet.getLastRow() > 1) {
      const uValues = userSheet.getDataRange().getValues();
      const uHeaders = uValues[0].map(h => String(h).trim().toLowerCase());
      const idxUId = uHeaders.indexOf("user_id") !== -1 ? uHeaders.indexOf("user_id") : 0;
      const idxUNama = uHeaders.indexOf("nama") !== -1 ? uHeaders.indexOf("nama") : 1;

      for (let i = 1; i < uValues.length; i++) {
        const uId = String(uValues[i][idxUId] || "").trim();
        const uNama = String(uValues[i][idxUNama] || "").trim();
        if (uId) userMap[uId] = uNama || uId;
      }
    }
  }

  // 3. Baca Sheet Kegiatan
  const sheet = ss.getSheetByName(KEGIATAN_CONFIG.SHEET_KEGIATAN);
  if (!sheet || sheet.getLastRow() <= 1) {
    return {
      success: true,
      data: [],
      master: { kategori: kategoriList }
    };
  }

  const values = sheet.getDataRange().getValues();
  const headers = values[0].map(h => String(h).trim().toLowerCase());

  const col = function(name, fallbackIdx) {
    const idx = headers.indexOf(name);
    return idx !== -1 ? idx : fallbackIdx;
  };

  const startDate = String(filters.startDate || "").trim();
  const endDate = String(filters.endDate || "").trim();
  const kategoriFilter = String(filters.kategoriId || "").trim();
  const keyword = String(filters.search || "").trim().toLowerCase();

  const results = [];

  for (let i = 1; i < values.length; i++) {
    const row = values[i];
    const laporanId = String(row[col("laporan_id", 0)] || "").trim();
    if (!laporanId) continue;

    const rowUserId = String(row[col("user_id", 3)] || "").trim();
    if (!isAdminView && rowUserId !== String(currentUser.user_id).trim()) {
      continue;
    }

    const rawTgl = row[col("tanggal_kegiatan", 1)];
    const tglStr = rawTgl instanceof Date
      ? Utilities.formatDate(rawTgl, tz, "yyyy-MM-dd")
      : String(rawTgl || "").trim().slice(0, 10);

    if (startDate && tglStr && tglStr < startDate) continue;
    if (endDate && tglStr && tglStr > endDate) continue;

    const katId = String(row[col("kategori_id", 4)] || "").trim();
    if (kategoriFilter && katId !== kategoriFilter) continue;

    const namaPihak = String(row[col("nama_pihak", 5)] || "").trim();
    const noHpKontak = String(row[col("no_hp_kontak", 6)] || "").trim();
    const lokasi = String(row[col("lokasi", 7)] || "").trim();
    const narasi = String(row[col("narasi_kegiatan", 8)] || "").trim();
    const tindakLanjut = String(row[col("tindak_lanjut", 9)] || "").trim();

    const rawFollowup = row[col("tanggal_followup", 10)];
    const followupStr = rawFollowup instanceof Date
      ? Utilities.formatDate(rawFollowup, tz, "yyyy-MM-dd")
      : String(rawFollowup || "").trim().slice(0, 10);

    const rawWaktuInput = row[col("waktu_input", 2)];
    const waktuInputStr = rawWaktuInput instanceof Date
      ? Utilities.formatDate(rawWaktuInput, tz, "dd/MM/yyyy HH:mm")
      : String(rawWaktuInput || "").trim();

    const namaDai = isAdminView ? (userMap[rowUserId] || rowUserId) : currentUser.nama;
    const namaKategori = kategoriMap[katId] || katId;

    if (keyword) {
      const haystack = [
        laporanId, namaPihak, noHpKontak, lokasi,
        narasi, namaKategori, namaDai
      ].join(" ").toLowerCase();
      if (haystack.indexOf(keyword) === -1) continue;
    }

    // Normalisasi seluruh isi kolom foto menjadi ID murni (kompatibel untuk data lama berbentuk URL maupun ID baru)
    const fotoIds = [
      extractDriveFileId_(row[col("foto_1_file_id", 11)]),
      extractDriveFileId_(row[col("foto_2_file_id", 12)]),
      extractDriveFileId_(row[col("foto_3_file_id", 13)]),
      extractDriveFileId_(row[col("foto_4_file_id", 14)]),
      extractDriveFileId_(row[col("foto_5_file_id", 15)])
    ].filter(Boolean);

    results.push({
      laporan_id: laporanId,
      tanggal_kegiatan: tglStr,
      waktu_input: waktuInputStr,
      user_id: rowUserId,
      nama_dai: namaDai,
      kategori_id: katId,
      nama_kategori: namaKategori,
      nama_pihak: namaPihak,
      no_hp_kontak: noHpKontak,
      lokasi: lokasi,
      narasi_kegiatan: narasi,
      tindak_lanjut: tindakLanjut,
      tanggal_followup: followupStr,
      foto_ids: fotoIds,
      jumlah_foto: Number(row[col("jumlah_foto", 16)]) || fotoIds.length,
      status_laporan: String(row[col("status_laporan", 17)] || "TERCATAT").trim(),
      cancelled_at: String(row[col("cancelled_at", 18)] || "").trim(),
      cancel_reason: String(row[col("cancel_reason", 20)] || "").trim()
    });
  }

  // Urutkan dari tanggal kegiatan terbaru
  results.sort(function(a, b) {
    return (b.tanggal_kegiatan || "").localeCompare(a.tanggal_kegiatan || "");
  });

  return {
    success: true,
    data: results,
    master: { kategori: kategoriList }
  };
}