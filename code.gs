
/**
 * JAWARA FKAM
 * Jelajah Dakwah Nusantara
 * Backend tahap awal
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