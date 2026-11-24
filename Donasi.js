/**************************************************
 * JAWARA FKAM
 * MODUL DATA DONASI + ENTRY DONASI
 **************************************************/


/**
 * Mengambil seluruh data transaksi donasi.
 * Akses hanya untuk ADMIN.
 *
 * Filter:
 * - startDate      : tanggal awal
 * - endDate        : tanggal akhir
 * - jenisDonasiId  : ID jenis donasi
 * - search         : nama donatur, nomor HP, atau nama da'i
 */
function getDonations(token, filters) {

  // 1. VALIDASI AKSES ADMIN
  const admin = validateSession(token);

  if (!admin || admin.role !== "ADMIN") {
    throw new Error("Akses ditolak. Fitur ini khusus ADMIN.");
  }

  filters = filters || {};

  // 2. VALIDASI FILTER TANGGAL
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
    throw new Error("Format tanggal tidak valid.");
  }

  if (startDate && endDate && startDate > endDate) {
    throw new Error("Tanggal awal tidak boleh melebihi tanggal akhir.");
  }

  // 3. FILTER PENCARIAN
  const jenisDonasiId = String(filters.jenisDonasiId || "").trim();
  const search = String(filters.search || "").trim().toLowerCase();

  // 4. MENGAMBIL DATA DATABASE
  const ss = getDatabase();

  const donationSheet = ss.getSheetByName("DONASI");
  const userSheet = ss.getSheetByName("USERS");
  const jenisSheet = ss.getSheetByName("MASTER_JENIS_DONASI");
  const programSheet = ss.getSheetByName("MASTER_PROGRAM");
  const metodeSheet = ss.getSheetByName("MASTER_METODE");

  if (!donationSheet || !userSheet || !jenisSheet ||
      !programSheet || !metodeSheet) {
    throw new Error("Salah satu sheet database tidak ditemukan.");
  }

  const donationRows = donationSheet.getDataRange().getValues();
  const userRows = userSheet.getDataRange().getValues();
  const jenisRows = jenisSheet.getDataRange().getValues();
  const programRows = programSheet.getDataRange().getValues();
  const metodeRows = metodeSheet.getDataRange().getValues();

  // 5. MEMBUAT PETA DATA MASTER
  const userMap = {};
  const jenisMap = {};
  const programMap = {};
  const metodeMap = {};

  userRows.slice(1).forEach(function(row) {
    if (row[0]) userMap[String(row[0])] = row[1];
  });

  jenisRows.slice(1).forEach(function(row) {
    if (row[0]) jenisMap[String(row[0])] = row[1];
  });

  programRows.slice(1).forEach(function(row) {
    if (row[0]) programMap[String(row[0])] = row[1];
  });

  metodeRows.slice(1).forEach(function(row) {
    if (row[0]) metodeMap[String(row[0])] = row[1];
  });

  // 6. MEMPROSES DATA DONASI
  const results = [];

  for (let i = 1; i < donationRows.length; i++) {
    const row = donationRows[i];
    if (!row[0]) continue;

    const tanggal = row[1];
    const namaDonatur = String(row[5] || "");
    const noHp = String(row[6] || "");
    const namaDai = String(userMap[String(row[3])] || "");
    const status = String(row[14] || "");

    if (!isDonationDateInRange_(tanggal, startDate, endDate)) {
      continue;
    }

    if (jenisDonasiId && String(row[8]) !== jenisDonasiId) {
      continue;
    }

    if (search) {
      const searchTarget = (
        namaDonatur + " " + noHp + " " + namaDai
      ).toLowerCase();

      if (!searchTarget.includes(search)) continue;
    }

    results.push({
      donasi_id: row[0],
      tanggal_donasi: formatDonationDate_(tanggal),
      waktu_input: formatDonationDateTime_(row[2]),
      user_id: row[3],
      nama_dai: namaDai,
      donatur_id: row[4],
      nama_donatur: namaDonatur,
      no_hp: noHp,
      jenis_donatur: row[7],
      jenis_donasi_id: row[8],
      jenis_donasi: jenisMap[String(row[8])] || "-",
      program_id: row[9],
      nama_program: programMap[String(row[9])] || "-",
      nominal: Number(row[10]) || 0,
      metode_id: row[11],
      metode: metodeMap[String(row[11])] || "-",
      keterangan: row[12] || "",
      bukti_donasi_file_id: row[13] || "",
      status_transaksi: status || "BELUM DITENTUKAN",
      created_at: formatDonationDateTime_(row[15]),
      cancelled_at: formatDonationDateTime_(row[16]),
      cancelled_by: row[17] || "",
      cancel_reason: row[18] || ""
    });
  }

  // 7. MENGURUTKAN DATA TERBARU
  results.sort(function(a, b) {
    return String(b.waktu_input).localeCompare(String(a.waktu_input));
  });

  // 8. MASTER JENIS DONASI AKTIF
  const jenisDonasi = jenisRows.slice(1)
    .filter(function(row) {
      return row[0] && String(row[2]).toUpperCase() === "AKTIF";
    })
    .map(function(row) {
      return { id: row[0], nama: row[1] };
    });

  // 9. MENGEMBALIKAN HASIL
  return {
    success: true,
    total: results.length,
    data: results,
    master: { jenis_donasi: jenisDonasi }
  };
}


/**
 * Memeriksa apakah tanggal donasi masuk
 * dalam rentang tanggal filter.
 */
function isDonationDateInRange_(value, startDate, endDate) {
  if (!startDate && !endDate) return true;

  let date = value instanceof Date ? value : new Date(value);
  if (isNaN(date.getTime())) return false;

  if (startDate && date < startDate) return false;
  if (endDate && date > endDate) return false;

  return true;
}


/**
 * Mengubah tanggal menjadi yyyy-MM-dd.
 */
function formatDonationDate_(value) {
  if (!value) return "";

  const date = value instanceof Date ? value : new Date(value);
  if (isNaN(date.getTime())) return "";

  return Utilities.formatDate(
    date,
    CONFIG.TIMEZONE,
    "yyyy-MM-dd"
  );
}


/**
 * Mengubah tanggal dan waktu menjadi yyyy-MM-dd HH:mm:ss.
 */
function formatDonationDateTime_(value) {
  if (!value) return "";

  const date = value instanceof Date ? value : new Date(value);
  if (isNaN(date.getTime())) return "";

  return Utilities.formatDate(
    date,
    CONFIG.TIMEZONE,
    "yyyy-MM-dd HH:mm:ss"
  );
}


/* ==================================================
   ENTRY DONASI - AKSES USER / DA'I
================================================== */


/**
 * Mengambil pilihan master untuk formulir entry donasi.
 * Akses hanya untuk USER.
 */
function getDonationFormOptions(token) {
  const user = validateSession(token);
  if (!user || user.role !== "USER") {
    throw new Error("Fitur ini hanya untuk akun da'i.");
  }

  const db = getDatabase();

  function getActiveOptions(sheetName, idColumn, nameColumn) {
    const sh = db.getSheetByName(sheetName);
    if (!sh || sh.getLastRow() < 2) return [];
    const values = sh.getDataRange().getValues();
    const headers = values[0];
    const idIndex = headers.indexOf(idColumn);
    const nameIndex = headers.indexOf(nameColumn);
    const statusIndex = headers.indexOf("status");
    if (idIndex < 0 || nameIndex < 0 || statusIndex < 0) {
      throw new Error("Struktur sheet " + sheetName + " tidak sesuai.");
    }
    return values.slice(1)
      .filter(function(row) {
        return String(row[statusIndex]).toUpperCase() === "AKTIF";
      })
      .map(function(row) {
        return { id: String(row[idIndex]), nama: String(row[nameIndex]) };
      });
  }

  // Hanya dua metode donasi yang diperbolehkan pada formulir.
  const allowedMethodNames = ["tunai", "transfer bank"];
  const allMethods = getActiveOptions("MASTER_METODE", "metode_id", "nama_metode");
  const allowedMethods = allMethods.filter(function(item) {
    return allowedMethodNames.includes(String(item.nama).trim().toLowerCase());
  });

  return {
    success: true,
    jenisDonasi: getActiveOptions(
      "MASTER_JENIS_DONASI", "jenis_donasi_id", "nama_jenis"
    ),
    program: getActiveOptions(
      "MASTER_PROGRAM", "program_id", "nama_program"
    ),
    metode: allowedMethods
  };
}

/**
 * Memastikan suatu ID tersedia dan aktif pada sheet master.
 */
function isActiveDonationMaster_(sheet, idColumn, id) {
  if (!id) return false;
  if (!sheet || sheet.getLastRow() < 2) return false;

  const rows = sheet.getDataRange().getValues();
  const headers = rows[0];

  const idIndex = headers.indexOf(idColumn);
  const statusIndex = headers.indexOf("status");

  if (idIndex < 0 || statusIndex < 0) {
    throw new Error("Struktur sheet master tidak sesuai.");
  }

  return rows.slice(1).some(function(row) {
    return String(row[idIndex]) === String(id) &&
      String(row[statusIndex]).toUpperCase() === "AKTIF";
  });
}


/**
 * Menyimpan entry donasi dari akun da'i.
 *
 * Data ditulis ke:
 * - DONATUR (jika donatur belum ditemukan)
 * - DONASI
 */
function createDonationEntry(token, payload) {
  const user = validateSession(token);
  if (!user || user.role !== "USER") {
    throw new Error("Hanya akun da'i yang dapat memasukkan donasi.");
  }

  payload = payload || {};
  const tanggal = String(payload.tanggal_donasi || "").trim();
  const namaDonatur = String(payload.nama_donatur || "").trim();
  const noHp = String(payload.no_hp || "").trim();
  const jenisDonatur = String(payload.jenis_donatur || "").trim();
  const jenisDonasiId = String(payload.jenis_donasi_id || "").trim();
  const programId = String(payload.program_id || "").trim();
  const metodeId = String(payload.metode_id || "").trim();
  const keterangan = String(payload.keterangan || "").trim();
  const nominal = Number(payload.nominal);

  if (!/^\d{4}-\d{2}-\d{2}$/.test(tanggal)) {
    throw new Error("Tanggal donasi tidak valid.");
  }
  const parts = tanggal.split("-").map(Number);
  const tanggalDonasi = new Date(parts[0], parts[1] - 1, parts[2]);
  if (tanggalDonasi.getFullYear() !== parts[0] ||
      tanggalDonasi.getMonth() !== parts[1] - 1 ||
      tanggalDonasi.getDate() !== parts[2]) {
    throw new Error("Tanggal donasi tidak valid.");
  }
  if (!namaDonatur) throw new Error("Nama donatur wajib diisi.");
  if (!["INDIVIDU", "LEMBAGA"].includes(jenisDonatur)) {
    throw new Error("Jenis donatur tidak valid.");
  }
  if (!jenisDonasiId || !metodeId) {
    throw new Error("Jenis donasi dan metode donasi wajib dipilih.");
  }
  if (!Number.isFinite(nominal) || nominal <= 0) {
    throw new Error("Nominal donasi harus lebih besar dari nol.");
  }

  const db = getDatabase();
  const jenisSheet = db.getSheetByName("MASTER_JENIS_DONASI");
  const metodeSheet = db.getSheetByName("MASTER_METODE");
  const programSheet = db.getSheetByName("MASTER_PROGRAM");
  const donorSheet = db.getSheetByName("DONATUR");
  const donationSheet = db.getSheetByName("DONASI");
  if (!jenisSheet || !metodeSheet || !programSheet || !donorSheet || !donationSheet) {
    throw new Error("Salah satu sheet database tidak ditemukan.");
  }

  if (!isActiveDonationMaster_(jenisSheet, "jenis_donasi_id", jenisDonasiId)) {
    throw new Error("Jenis donasi tidak ditemukan atau tidak aktif.");
  }
  if (!isActiveDonationMaster_(metodeSheet, "metode_id", metodeId)) {
    throw new Error("Metode donasi tidak ditemukan atau tidak aktif.");
  }
  if (programId && !isActiveDonationMaster_(programSheet, "program_id", programId)) {
    throw new Error("Program tidak ditemukan atau tidak aktif.");
  }

  // Pastikan ID metode memang Tunai atau Transfer Bank.
  const methodRows = metodeSheet.getDataRange().getValues();
  const methodHeaders = methodRows[0];
  const methodIdIndex = methodHeaders.indexOf("metode_id");
  const methodNameIndex = methodHeaders.indexOf("nama_metode");
  const methodRow = methodRows.slice(1).find(function(row) {
    return String(row[methodIdIndex]) === metodeId;
  });
  if (!methodRow) throw new Error("Metode donasi tidak ditemukan.");
  const methodName = String(methodRow[methodNameIndex] || "").trim().toLowerCase();
  if (!["tunai", "transfer bank"].includes(methodName)) {
    throw new Error("Metode donasi hanya boleh Tunai atau Transfer Bank.");
  }

  const proof = payload.bukti_transfer || null;
  if (methodName === "transfer bank" && !proof) {
    throw new Error("Bukti transfer wajib diunggah untuk metode Transfer Bank.");
  }

  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  let uploadedFile = null;
  try {
    let proofFileId = "";
    if (proof) {
      const allowedTypes = ["image/jpeg", "image/png", "image/webp", "application/pdf"];
      const mimeType = String(proof.mimeType || "").toLowerCase();
      const fileName = String(proof.fileName || "bukti-transfer").replace(/[\\/:*?"<>|]/g, "_");
      const base64 = String(proof.base64 || "");
      if (!allowedTypes.includes(mimeType)) {
        throw new Error("Format bukti transfer harus JPG, PNG, WEBP, atau PDF.");
      }
      if (!base64) throw new Error("File bukti transfer tidak terbaca.");
      const bytes = Utilities.base64Decode(base64);
      if (bytes.length > 5 * 1024 * 1024) {
        throw new Error("Ukuran bukti transfer maksimal 5 MB.");
      }
      const blob = Utilities.newBlob(bytes, mimeType, fileName);
      const folderId = PropertiesService.getScriptProperties().getProperty("JAWARA_UPLOAD_FOLDER_ID");
      uploadedFile = folderId
        ? DriveApp.getFolderById(folderId).createFile(blob)
        : DriveApp.createFile(blob);
      proofFileId = uploadedFile.getId();
    }

    // Cari donor berdasarkan nomor HP; jika tidak ditemukan, buat catatan donor.
    let donaturId = "";
    if (noHp && donorSheet.getLastRow() >= 2) {
      const donorData = donorSheet.getDataRange().getValues();
      const donorHeaders = donorData[0];
      const phoneIndex = donorHeaders.indexOf("no_hp");
      const donorIdIndex = donorHeaders.indexOf("donatur_id");
      if (phoneIndex >= 0 && donorIdIndex >= 0) {
        for (let i = 1; i < donorData.length; i++) {
          const savedPhone = String(donorData[i][phoneIndex] || "").replace(/\D/g, "");
          if (savedPhone && savedPhone === noHp.replace(/\D/g, "")) {
            donaturId = String(donorData[i][donorIdIndex]);
            break;
          }
        }
      }
    }

    if (!donaturId) {
      donaturId = "DNR-" + Utilities.getUuid().slice(0, 8).toUpperCase();
      donorSheet.appendRow([
        donaturId, namaDonatur, noHp, "", "", tanggalDonasi, "ENTRY_DA'I", new Date()
      ]);
    }

    const now = new Date();
    const donasiId = "DNS-" + Utilities.getUuid().slice(0, 8).toUpperCase();
    donationSheet.appendRow([
      donasiId, tanggalDonasi, now, user.user_id, donaturId,
      namaDonatur, noHp, jenisDonatur, jenisDonasiId, programId,
      nominal, metodeId, keterangan, proofFileId, "TERCATAT",
      now, "", "", ""
    ]);

    return {
      success: true,
      message: "Donasi berhasil disimpan.",
      donasi_id: donasiId,
      bukti_donasi_file_id: proofFileId
    };
  } catch (error) {
    if (uploadedFile) {
      try { uploadedFile.setTrashed(true); } catch (cleanupError) {}
    }
    throw error;
  } finally {
    lock.releaseLock();
  }
}
