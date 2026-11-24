/**
 * JAWARA FKAM
 * Jelajah Dakwah Nusantara
 * Backend Utama
 */

const CONFIG = {
  APP_NAME: "JAWARA",
  ORGANIZATION: "FKAM",
  TIMEZONE: "Asia/Jakarta",
  DATABASE_ID: "1F1NQjAvCE1-2TU4K6Y-k_PrPLfUNLpKOBVsmIYYhLoU"
};

/**
 * Membuka database utama
 */
function getDatabase() {
  return SpreadsheetApp.openById(CONFIG.DATABASE_ID);
}

/**
 * Menguji koneksi database
 */
function testDatabase() {
  const ss = getDatabase();

  const sheets = ss.getSheets().map(sheet => ({
    name: sheet.getName(),
    rows: sheet.getLastRow(),
    columns: sheet.getLastColumn()
  }));

  Logger.log({
    application: CONFIG.APP_NAME,
    database: ss.getName(),
    totalSheets: sheets.length,
    sheets: sheets
  });

  return sheets;
}

/**
 * Memeriksa ketersediaan sheet
 */
function checkDatabase() {
  const ss = getDatabase();

  const requiredSheets = [
    "USERS",
    "DONASI",
    "DONATUR",
    "LAPORAN_AKTIVITAS",
    "MASTER_PROGRAM",
    "MASTER_KATEGORI",
    "MASTER_METODE",
    "MASTER_JENIS_DONASI",
    "LOG_AKTIVITAS",
    "SESSIONS",
    "PENGATURAN"
  ];

  return requiredSheets.map(name => ({
    sheet: name,
    available: ss.getSheetByName(name) !== null
  }));
}

/**
 * Membuka halaman utama aplikasi JAWARA FKAM.
 */
function doGet() {
  return HtmlService
    .createHtmlOutputFromFile("Index")
    .setTitle("JAWARA FKAM - Jelajah Dakwah Nusantara")
    .addMetaTag("viewport", "width=device-width, initial-scale=1");
}

/**
 * Helper internal untuk memetakan sheet menjadi array objek
 */
function getSheetRecords_(sheetName) {
  const ss = getDatabase();
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet || sheet.getLastRow() <= 1) return [];

  const values = sheet.getDataRange().getValues();
  const headers = values[0].map(h => String(h).trim().toLowerCase());
  const records = [];

  for (let i = 1; i < values.length; i++) {
    const rowObj = {};
    for (let j = 0; j < headers.length; j++) {
      if (headers[j]) {
        rowObj[headers[j]] = values[i][j];
      }
    }
    records.push(rowObj);
  }
  return records;
}

/**
 * Helper internal format tanggal YYYY-MM-DD
 */
function formatDateIso_(val) {
  if (!val) return "";
  if (val instanceof Date) {
    return Utilities.formatDate(val, CONFIG.TIMEZONE, "yyyy-MM-dd");
  }
  return String(val).trim().substring(0, 10);
}

/**
 * Helper internal format waktu dd/MM/yyyy HH:mm
 */
function formatDateTimeDisplay_(val) {
  if (!val) return "-";
  if (val instanceof Date) {
    return Utilities.formatDate(val, CONFIG.TIMEZONE, "dd/MM/yyyy HH:mm");
  }
  return String(val);
}

/**
 * Mengambil riwayat donasi khusus USER (Da'i) yang sedang login
 */
function getUserDonations(token, filters) {
  const user = validateSession(token);
  if (!user || user.role !== "USER") {
    throw new Error("Akses ditolak. Khusus pengguna Da'i.");
  }

  const donasiRows = getSheetRecords_("DONASI");
  const donaturRows = getSheetRecords_("DONATUR");
  const programRows = getSheetRecords_("MASTER_PROGRAM");
  const jenisRows = getSheetRecords_("MASTER_JENIS_DONASI");
  const metodeRows = getSheetRecords_("MASTER_METODE");

  // Buat peta relasi master agar nama & detail tampil lengkap meski di sheet DONASI hanya tersimpan ID
  const donaturMap = {};
  donaturRows.forEach(d => {
    const id = String(d.donatur_id || d.id || "").trim();
    if (id) donaturMap[id] = d;
  });

  const programMap = {};
  programRows.forEach(p => {
    const id = String(p.program_id || p.id || "").trim();
    if (id) programMap[id] = p;
  });

  const jenisMap = {};
  const activeJenisList = [];
  jenisRows.forEach(j => {
    const id = String(j.jenis_donasi_id || j.id || "").trim();
    const nama = String(j.nama_jenis || j.nama || j.jenis_donasi || "").trim();
    const status = String(j.status || "AKTIF").trim().toUpperCase();
    if (id) {
      jenisMap[id] = nama;
      if (status === "AKTIF") {
        activeJenisList.push({ id: id, nama: nama });
      }
    }
  });

  const metodeMap = {};
  metodeRows.forEach(m => {
    const id = String(m.metode_id || m.id || "").trim();
    const nama = String(m.nama_metode || m.nama || m.metode || "").trim();
    if (id) metodeMap[id] = nama;
  });

  const startDate = filters && filters.startDate ? String(filters.startDate).trim() : "";
  const endDate = filters && filters.endDate ? String(filters.endDate).trim() : "";
  const jenisFilter = filters && filters.jenisDonasiId ? String(filters.jenisDonasiId).trim() : "";
  const search = filters && filters.search ? String(filters.search).trim().toLowerCase() : "";

  const currentUserId = String(user.user_id).trim();
  const results = [];

  donasiRows.forEach(row => {
    const rowUserId = String(row.user_id || row.dai_id || row.created_by || "").trim();
    if (rowUserId !== currentUserId) return;

    const tglStr = formatDateIso_(row.tanggal_donasi);
    if (startDate && tglStr < startDate) return;
    if (endDate && tglStr > endDate) return;

    const jenisId = String(row.jenis_donasi_id || "").trim();
    if (jenisFilter && jenisId !== jenisFilter) return;

    const donaturId = String(row.donatur_id || "").trim();
    const donaturInfo = donaturMap[donaturId] || {};
    const panggilan = String(donaturInfo.panggilan || "").trim();
    const rawNamaDonatur = String(row.nama_donatur || donaturInfo.nama || "").trim();
    const namaDonatur = (panggilan && rawNamaDonatur && !rawNamaDonatur.toLowerCase().startsWith(panggilan.toLowerCase() + " "))
      ? (panggilan + " " + rawNamaDonatur)
      : (rawNamaDonatur || "-");

    const noHp = String(row.no_hp || donaturInfo.no_hp || "").trim();

    const programId = String(row.program_id || "").trim();
    const programInfo = programMap[programId] || {};
    const namaProgramBase = String(row.nama_program || programInfo.nama_program || programInfo.nama || "").trim();
    const subProgram = String(row.sub_program || programInfo.sub_program || "").trim();
    const namaProgramDisplay = namaProgramBase
      ? (subProgram ? namaProgramBase + " - " + subProgram : namaProgramBase)
      : "-";

    const jenisNama = String(row.jenis_donasi || jenisMap[jenisId] || "-").trim();
    const metodeId = String(row.metode_id || "").trim();
    const metodeNama = String(row.metode || row.metode_nama || metodeMap[metodeId] || "-").trim();

    if (search) {
      const hay = (namaDonatur + " " + noHp + " " + namaProgramDisplay + " " + jenisNama).toLowerCase();
      if (!hay.includes(search)) return;
    }

    results.push({
      donasi_id: String(row.donasi_id || row.id || ""),
      tanggal_donasi: tglStr,
      waktu_input: formatDateTimeDisplay_(row.waktu_input || row.created_at),
      nama_donatur: namaDonatur,
      no_hp: noHp || "-",
      jenis_donatur: String(row.jenis_donatur || "-"),
      jenis_donasi_id: jenisId,
      jenis_donasi: jenisNama,
      nama_program: namaProgramDisplay,
      nominal: Number(row.nominal || 0),
      metode: metodeNama,
      nama_dai: user.nama || "-",
      status_transaksi: String(row.status_transaksi || row.status || "TERCATAT"),
      keterangan: String(row.keterangan || ""),
      bukti_donasi_file_id: String(row.bukti_donasi_file_id || row.bukti_file_id || ""),
      cancel_reason: String(row.cancel_reason || ""),
      cancelled_at: formatDateTimeDisplay_(row.cancelled_at)
    });
  });

  // Urutkan transaksi terbaru di paling atas
  results.sort((a, b) => {
    if (a.tanggal_donasi !== b.tanggal_donasi) {
      return b.tanggal_donasi.localeCompare(a.tanggal_donasi);
    }
    return b.donasi_id.localeCompare(a.donasi_id);
  });

  return {
    success: true,
    data: results,
    master: {
      jenis_donasi: activeJenisList
    }
  };
}