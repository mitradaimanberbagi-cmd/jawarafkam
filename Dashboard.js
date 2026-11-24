
/**************************************************
 * JAWARA FKAM
 * DASHBOARD DATA
 **************************************************/

/**
 * Mengambil ringkasan dashboard sesuai hak akses.
 *
 * ADMIN:
 * - Melihat seluruh data.
 * - Dapat memfilter berdasarkan tanggal.
 * - Dapat memfilter berdasarkan da'i.
 *
 * USER:
 * - Hanya melihat data miliknya sendiri.
 * - Dapat memfilter berdasarkan tanggal.
 */
function getDashboardData(token, filters) {

  const user = validateSession(token);
  const isAdmin = user.role === "ADMIN";

  filters = filters || {};

  // ==========================================
  // FILTER TANGGAL
  // ==========================================

  const startDate = filters.startDate
    ? new Date(filters.startDate + "T00:00:00")
    : null;

  const endDate = filters.endDate
    ? new Date(filters.endDate + "T23:59:59.999")
    : null;

  if (
    (startDate && isNaN(startDate.getTime())) ||
    (endDate && isNaN(endDate.getTime()))
  ) {
    throw new Error("Format tanggal filter tidak valid.");
  }

  if (startDate && endDate && startDate > endDate) {
    throw new Error("Tanggal awal tidak boleh melebihi tanggal akhir.");
  }

  // ==========================================
  // FILTER DA'I
  // ==========================================

  let selectedUserId = "";

  if (isAdmin) {
    selectedUserId = String(filters.userId || "").trim();
  } else {
    // USER tidak boleh memilih data pengguna lain.
    selectedUserId = String(user.user_id);
  }

  // ==========================================
  // KONEKSI DATABASE
  // ==========================================

  const ss = getDatabase();

  // ==========================================
  // DATA DONASI
  // ==========================================

  const donationSheet = ss.getSheetByName("DONASI");
  const donationRows = donationSheet.getDataRange().getValues();

  let totalDonasi = 0;
  let jumlahTransaksi = 0;

  const donorIds = new Set();

  for (let i = 1; i < donationRows.length; i++) {

    const row = donationRows[i];

    if (!row[0]) continue;

    const rowUserId = String(row[3]);
    const status = String(row[15] || "").toUpperCase();

    // Filter da'i.
    if (
      selectedUserId &&
      rowUserId !== selectedUserId
    ) {
      continue;
    }

    // Abaikan transaksi dibatalkan.
    if (status === "DIBATALKAN") continue;

    // Filter tanggal donasi.
    if (
      !isDashboardDateInRange_(
        row[1],
        startDate,
        endDate
      )
    ) {
      continue;
    }

    totalDonasi += Number(row[10]) || 0;
    jumlahTransaksi++;

    if (row[4]) {
      donorIds.add(String(row[4]));
    }
  }

  // ==========================================
  // DATA KEGIATAN
  // ==========================================

  const activitySheet = ss.getSheetByName("LAPORAN_AKTIVITAS");
  const activityRows = activitySheet.getDataRange().getValues();

  let totalKegiatan = 0;

  for (let i = 1; i < activityRows.length; i++) {

    const row = activityRows[i];

    if (!row[0]) continue;

    const rowUserId = String(row[3]);
    const status = String(row[17] || "").toUpperCase();

    // Filter da'i.
    if (
      selectedUserId &&
      rowUserId !== selectedUserId
    ) {
      continue;
    }

    // Abaikan laporan dibatalkan.
    if (status === "DIBATALKAN") continue;

    // Filter tanggal kegiatan.
    if (
      !isDashboardDateInRange_(
        row[1],
        startDate,
        endDate
      )
    ) {
      continue;
    }

    totalKegiatan++;
  }

  // ==========================================
  // JUMLAH DA'I AKTIF
  // ==========================================

  const userSheet = ss.getSheetByName("USERS");
  const userRows = userSheet.getDataRange().getValues();

  let jumlahDaiAktif = 0;

  if (isAdmin && !selectedUserId) {

    for (let i = 1; i < userRows.length; i++) {

      const row = userRows[i];

      if (
        row[0] &&
        row[5] === "USER" &&
        row[7] === "AKTIF"
      ) {
        jumlahDaiAktif++;
      }
    }
  }

  // ==========================================
  // JUMLAH DONATUR
  // ==========================================

  // Menghitung donatur unik berdasarkan transaksi
  // yang memenuhi filter tanggal dan filter da'i.

  const jumlahDonatur = donorIds.size;

  // ==========================================
  // DATA PENGGUNA TERPILIH
  // ==========================================

  let namaDaiTerpilih = "";

  if (selectedUserId) {

    for (let i = 1; i < userRows.length; i++) {

      if (
        String(userRows[i][0]) === selectedUserId &&
        userRows[i][5] === "USER"
      ) {
        namaDaiTerpilih = userRows[i][1];
        break;
      }
    }
  }

  // ==========================================
  // HASIL DASHBOARD
  // ==========================================

  return {
    success: true,

    user: {
      user_id: user.user_id,
      nama: user.nama,
      role: user.role
    },

    filter: {
      startDate: filters.startDate || "",
      endDate: filters.endDate || "",
      userId: selectedUserId,
      namaDai: namaDaiTerpilih
    },

    summary: {
      total_donasi: totalDonasi,
      jumlah_transaksi: jumlahTransaksi,
      total_kegiatan: totalKegiatan,
      jumlah_dai_aktif: jumlahDaiAktif,
      jumlah_donatur: jumlahDonatur
    }
  };
}


/**
 * Memeriksa apakah tanggal data masuk
 * ke dalam rentang tanggal filter.
 */
function isDashboardDateInRange_(
  value,
  startDate,
  endDate
) {

  if (!startDate && !endDate) return true;

  let date;

  if (value instanceof Date) {
    date = value;
  } else {
    date = new Date(value);
  }

  if (isNaN(date.getTime())) return false;

  if (startDate && date < startDate) {
    return false;
  }

  if (endDate && date > endDate) {
    return false;
  }

  return true;
}
