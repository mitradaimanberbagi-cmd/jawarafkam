
/**************************************************
 * JAWARA FKAM
 * MODUL DATA DONATUR
 **************************************************/

/**
 * Validasi hak akses Data Donatur.
 * ADMIN dan USER boleh melihat data.
 * Hanya USER boleh menambahkan data.
 */
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
 * Mengambil daftar donatur.
 * ADMIN: hanya melihat.
 * USER: melihat untuk kebutuhan pencatatan donasi.
 *
 * Parameter search bersifat opsional.
 */
function getDonors(token, search) {
  try {
    validateDonorAccess_(token, false);

    const db = getDatabase();
    const sheet = db.getSheetByName("DONATUR");

    if (!sheet) {
      throw new Error("Sheet DONATUR tidak ditemukan.");
    }

    const lastRow = sheet.getLastRow();

    if (lastRow < 2) {
      return {
        success: true,
        data: [],
        total: 0
      };
    }

    const keyword = String(search || "")
      .trim()
      .toLowerCase();

    const values = sheet
      .getRange(2, 1, lastRow - 1, 8)
      .getValues();

    const data = values
      .map(function(row) {
        return {
          donatur_id: row[0],
          nama: row[1],
          no_hp: row[2],
          email: row[3],
          alamat: row[4],
          tanggal_bergabung: row[5],
          sumber_data: row[6],
          created_at: row[7]
        };
      })
      .filter(function(donor) {
        if (!keyword) return true;

        const searchable = [
          donor.nama,
          donor.no_hp,
          donor.email,
          donor.alamat
        ].join(" ").toLowerCase();

        return searchable.includes(keyword);
      })
      .reverse();

    return {
      success: true,
      data: data,
      total: data.length
    };

  } catch (error) {
    return {
      success: false,
      message: error.message
    };
  }
}


/**
 * Menambahkan donatur baru.
 * Hanya USER yang boleh menjalankan fungsi ini.
 */
function createDonor(token, payload) {
  try {
    const user = validateDonorAccess_(token, true);

    if (!payload || typeof payload !== "object") {
      throw new Error("Data donatur tidak valid.");
    }

    const nama = String(payload.nama || "").trim();
    const noHp = String(payload.no_hp || "").trim();
    const email = String(payload.email || "").trim();
    const alamat = String(payload.alamat || "").trim();

    if (!nama) {
      throw new Error("Nama donatur wajib diisi.");
    }

    if (nama.length > 150) {
      throw new Error("Nama donatur maksimal 150 karakter.");
    }

    if (noHp.length > 25) {
      throw new Error("Nomor WhatsApp maksimal 25 karakter.");
    }

    if (email && !isValidEmail_(email)) {
      throw new Error("Format email tidak valid.");
    }

    const lock = LockService.getScriptLock();
    lock.waitLock(10000);

    try {
      const db = getDatabase();
      const sheet = db.getSheetByName("DONATUR");

      if (!sheet) {
        throw new Error("Sheet DONATUR tidak ditemukan.");
      }

      const lastRow = sheet.getLastRow();

      // Cegah nomor WhatsApp yang sama tercatat dua kali.
      if (noHp && lastRow >= 2) {
        const existingPhones = sheet
          .getRange(2, 3, lastRow - 1, 1)
          .getValues();

        const normalizedPhone = noHp.replace(/\D/g, "");

        const duplicate = existingPhones.some(function(row) {
          return String(row[0] || "").replace(/\D/g, "") ===
            normalizedPhone;
        });

        if (duplicate) {
          throw new Error(
            "Nomor WhatsApp sudah terdaftar. Silakan cari nama donatur pada daftar."
          );
        }
      }

      const donaturId = "DON-" +
        Utilities.getUuid().substring(0, 8).toUpperCase();

      const now = new Date();

      sheet.appendRow([
        donaturId,
        nama,
        noHp,
        email,
        alamat,
        now,
        "JAWARA",
        now
      ]);

      return {
        success: true,
        message: "Data donatur berhasil ditambahkan.",
        data: {
          donatur_id: donaturId,
          nama: nama,
          no_hp: noHp,
          email: email,
          alamat: alamat
        }
      };

    } finally {
      lock.releaseLock();
    }

  } catch (error) {
    return {
      success: false,
      message: error.message
    };
  }
}
