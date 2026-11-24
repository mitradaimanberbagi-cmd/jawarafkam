/**************************************************
 * JAWARA FKAM - DATA DONATUR
 * Gabungan perbaikan pemuatan dan penyimpanan donatur.
 * Mendukung pembatasan kepemilikan data per USER (Da'i)
 * dan tampilan seluruh data beserta pengelola bagi ADMIN.
 **************************************************/

function validateDonorAccess_(token, userOnly) {
  const user = validateSession(token);
  if (!user || !user.user_id) {
    throw new Error("Sesi tidak valid. Silakan login kembali.");
  }

  const role = String(user.role || "").toUpperCase();
  if (role !== "ADMIN" && role !== "USER") {
    throw new Error("Akses tidak diizinkan.");
  }
  if (userOnly && role !== "USER") {
    throw new Error("Hanya USER yang dapat menambahkan donatur.");
  }
  return user;
}


/**
 * Membaca posisi kolom berdasarkan nama header.
 * Tidak bergantung pada urutan kolom di Google Sheets.
 */
function getDonorColumnMap_(sheet) {
  if (!sheet) throw new Error("Sheet DONATUR tidak ditemukan.");
  const lastColumn = sheet.getLastColumn();
  if (lastColumn < 1) throw new Error("Header sheet DONATUR belum tersedia.");

  const headers = sheet.getRange(1, 1, 1, lastColumn).getDisplayValues()[0]
    .map(function(value) {
      return String(value || "").trim().toLowerCase();
    });

  const map = {};
  headers.forEach(function(header, index) {
    if (header) map[header] = index;
  });

  const required = [
    "donatur_id", "nama", "no_hp", "alamat",
    "tanggal_bergabung", "sumber_data", "created_at",
    "panggilan", "program_id", "sub_program", "user_id"
  ];

  const missing = required.filter(function(header) {
    return map[header] === undefined;
  });

  if (missing.length) {
    throw new Error(
      "Kolom " + missing.join(", ") +
      " belum tersedia di sheet DONATUR. Tambahkan header '" + missing.join(", ") + "' pada baris pertama sheet DONATUR."
    );
  }

  return map;
}


/**
 * Mengambil peta user_id -> nama dari sheet USERS
 * untuk menampilkan nama pengelola donatur pada sisi ADMIN.
 */
function getUserNameMap_(db) {
  const userSheet = db.getSheetByName("USERS");
  const userMap = {};
  if (!userSheet || userSheet.getLastRow() < 2) return userMap;

  const userRows = userSheet.getRange(2, 1, userSheet.getLastRow() - 1, 2).getValues();
  userRows.forEach(function(row) {
    const id = String(row[0] || "").trim();
    const nama = String(row[1] || "").trim();
    if (id) userMap[id] = nama;
  });

  return userMap;
}


/**
 * Mengambil data donatur:
 * - USER hanya melihat donatur yang dia kelola (user_id sesuai).
 * - ADMIN melihat semua data donatur beserta nama pengelolanya.
 */
function getDonors(token, search) {
  try {
    const currentUser = validateDonorAccess_(token, false);
    const isAdmin = String(currentUser.role || "").toUpperCase() === "ADMIN";
    const currentUserId = String(currentUser.user_id || "").trim();

    const db = getDatabase();
    const sheet = db.getSheetByName("DONATUR");
    if (!sheet) throw new Error("Sheet DONATUR tidak ditemukan.");

    if (sheet.getLastRow() < 2) {
      return { success: true, data: [], total: 0 };
    }

    const donorMap = getDonorColumnMap_(sheet);
    const programResult = getProgramRows_();
    const programMap = {};
    const pm = programResult.map;

    programResult.values.forEach(function(row) {
      programMap[String(row[pm.program_id])] =
        String(row[pm.nama_program] || "");
    });

    const userNameMap = isAdmin ? getUserNameMap_(db) : {};
    const keyword = String(search || "").trim().toLowerCase();
    const values = sheet.getRange(
      2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()
    ).getValues();

    const data = values
      .filter(function(row) {
        // Jika bukan ADMIN, filter ketat hanya donatur milik USER yang sedang login
        if (!isAdmin) {
          const ownerId = String(row[donorMap.user_id] || "").trim();
          return ownerId !== "" && ownerId === currentUserId;
        }
        return true;
      })
      .map(function(row) {
        const programId = row[donorMap.program_id] || "";
        const ownerId = String(row[donorMap.user_id] || "").trim();
        const ownerName = ownerId ? (userNameMap[ownerId] || ownerId) : "Belum ditentukan";

        return {
          donatur_id: row[donorMap.donatur_id],
          nama: row[donorMap.nama],
          panggilan: row[donorMap.panggilan] || "",
          no_hp: row[donorMap.no_hp] || "",
          alamat: row[donorMap.alamat] || "",
          program_id: programId,
          nama_program: programMap[String(programId)] || "",
          sub_program: row[donorMap.sub_program] || "",
          user_id: ownerId,
          nama_pengelola: isAdmin ? ownerName : (currentUser.nama || ""),
          tanggal_bergabung: formatDonorDate_(row[donorMap.tanggal_bergabung]),
          sumber_data: row[donorMap.sumber_data] || "",
          created_at: formatDonorDate_(row[donorMap.created_at])
        };
      })
      .filter(function(donor) {
        if (!keyword) return true;
        const searchableFields = [
          donor.nama,
          donor.panggilan,
          donor.no_hp,
          donor.alamat,
          donor.nama_program,
          donor.sub_program
        ];
        if (isAdmin) {
          searchableFields.push(donor.nama_pengelola);
        }
        return searchableFields.join(" ").toLowerCase().includes(keyword);
      })
      .reverse();

    return {
      success: true,
      data: data,
      total: data.length
    };

  } catch (error) {
    console.error("getDonors error: " + (error.stack || error.message));
    return {
      success: false,
      message: error.message || String(error)
    };
  }
}

function formatDonorDate_(value) {
  if (!value) return "";
  const date = value instanceof Date ? value : new Date(value);
  if (isNaN(date.getTime())) return String(value);
  return Utilities.formatDate(date, CONFIG.TIMEZONE, "yyyy-MM-dd HH:mm:ss");
}


/**
 * Menambahkan donatur.
 * Hanya USER yang diperbolehkan, otomatis tercatat sebagai pengelola (user_id).
 */
function createDonor(token, payload) {
  try {
    const currentUser = validateDonorAccess_(token, true);
    const currentUserId = String(currentUser.user_id || "").trim();
    payload = payload || {};

    const nama = String(payload.nama || "").trim();
    const panggilan = String(payload.panggilan || "").trim();
    const noHp = String(payload.no_hp || "").trim();
    const alamat = String(payload.alamat || "").trim();
    const programId = String(payload.program_id || "").trim();
    const subProgram = String(payload.sub_program || "").trim();

    if (!nama) throw new Error("Nama donatur wajib diisi.");
    if (nama.length > 150) {
      throw new Error("Nama donatur maksimal 150 karakter.");
    }
    if (!["Bapak", "Ibu", "Mas", "Mbak", "Owner"].includes(panggilan)) {
      throw new Error("Silakan pilih panggilan donatur.");
    }
    if (noHp.length > 25) {
      throw new Error("Nomor WhatsApp maksimal 25 karakter.");
    }

    const db = getDatabase();
    const donorSheet = db.getSheetByName("DONATUR");
    if (!donorSheet) throw new Error("Sheet DONATUR tidak ditemukan.");

    const donorMap = getDonorColumnMap_(donorSheet);

    let verifiedSubProgram = "";
    if (programId) {
      const programResult = getProgramRows_();
      const pm = programResult.map;
      const match = programResult.values.find(function(row) {
        return String(row[pm.program_id]) === programId &&
          String(row[pm.status]).toUpperCase() === "AKTIF";
      });

      if (!match) {
        throw new Error("Program donatur tidak ditemukan atau tidak aktif.");
      }

      verifiedSubProgram = String(match[pm.sub_program] || "").trim();

      if (subProgram && verifiedSubProgram && subProgram !== verifiedSubProgram) {
        throw new Error("Sub program tidak sesuai dengan program yang dipilih.");
      }
    } else if (subProgram) {
      throw new Error("Pilih program terlebih dahulu sebelum mengisi sub program.");
    }

    const lock = LockService.getScriptLock();
    lock.waitLock(10000);

    try {
      const lastRow = donorSheet.getLastRow();

      if (noHp && lastRow >= 2) {
        const existingRows = donorSheet.getRange(
          2, 1, lastRow - 1, donorSheet.getLastColumn()
        ).getValues();

        const normalizedPhone = noHp.replace(/\D/g, "");
        const duplicate = existingRows.some(function(row) {
          return String(row[donorMap.no_hp] || "").replace(/\D/g, "") ===
            normalizedPhone;
        });

        if (duplicate) {
          throw new Error(
            "Nomor WhatsApp sudah terdaftar dalam database."
          );
        }
      }

      const newRow = new Array(donorSheet.getLastColumn()).fill("");
      const now = new Date();
      const donorId = "DON-" +
        Utilities.getUuid().substring(0, 8).toUpperCase();

      newRow[donorMap.donatur_id] = donorId;
      newRow[donorMap.nama] = nama;
      newRow[donorMap.panggilan] = panggilan;
      newRow[donorMap.no_hp] = noHp;
      newRow[donorMap.alamat] = alamat;
      newRow[donorMap.tanggal_bergabung] = now;
      newRow[donorMap.sumber_data] = "JAWARA";
      newRow[donorMap.created_at] = now;
      newRow[donorMap.program_id] = programId;
      newRow[donorMap.sub_program] = programId ? verifiedSubProgram : "";
      newRow[donorMap.user_id] = currentUserId;

      donorSheet.appendRow(newRow);

      return {
        success: true,
        message: "Data donatur berhasil ditambahkan.",
        data: {
          donatur_id: donorId,
          nama: nama,
          panggilan: panggilan,
          no_hp: noHp,
          program_id: programId,
          sub_program: programId ? verifiedSubProgram : "",
          user_id: currentUserId
        }
      };

    } finally {
      lock.releaseLock();
    }

  } catch (error) {
    console.error("createDonor error: " + (error.stack || error.message));
    return {
      success: false,
      message: error.message || String(error)
    };
  }
}