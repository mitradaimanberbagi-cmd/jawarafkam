
/**************************************************
 * JAWARA FKAM
 * MODUL MANAJEMEN PROGRAM
 **************************************************/


/**
 * Memastikan pengguna memiliki sesi aktif.
 */

function validateProgramAccess_(token, adminOnly) {
  const user = validateSession(token);

  if (!user || !user.user_id) {
    throw new Error("Sesi tidak valid. Silakan login kembali.");
  }

  const role = String(user.role || "").toUpperCase();

  if (role !== "ADMIN" && role !== "USER") {
    throw new Error("Akses tidak diizinkan.");
  }

  if (adminOnly && role !== "ADMIN") {
    throw new Error("Hanya ADMIN yang dapat mengelola program.");
  }

  return user;
}


/**
 * Mengambil seluruh data program untuk halaman
 * Manajemen Program.
 *
 * Akses: ADMIN
 */
function getProgramManagementData(token) {
  try {
    validateProgramAccess_(token, true);

    const db = getDatabase();
    const sheet = db.getSheetByName("MASTER_PROGRAM");

    if (!sheet) {
      throw new Error("Sheet MASTER_PROGRAM tidak ditemukan.");
    }

    const lastRow = sheet.getLastRow();

    if (lastRow < 2) {
      return {
        success: true,
        data: []
      };
    }

    const values = sheet
      .getRange(2, 1, lastRow - 1, 5)
      .getValues();

    const data = values.map(function(row) {
      return {
        program_id: row[0],
        nama_program: row[1],
        kategori_program: row[2],
        deskripsi: row[3],
        status: row[4]
      };
    });

    return {
      success: true,
      data: data
    };

  } catch (error) {
    return {
      success: false,
      message: error.message
    };
  }
}


/**
 * Mengambil daftar program AKTIF.
 *
 * Akses: ADMIN dan USER
 * Digunakan untuk pilihan program pada Entry Donasi.
 */
function getActivePrograms(token) {
  try {
    validateProgramAccess_(token, false);

    const db = getDatabase();
    const sheet = db.getSheetByName("MASTER_PROGRAM");

    if (!sheet) {
      throw new Error("Sheet MASTER_PROGRAM tidak ditemukan.");
    }

    const lastRow = sheet.getLastRow();

    if (lastRow < 2) {
      return {
        success: true,
        data: []
      };
    }

    const values = sheet
      .getRange(2, 1, lastRow - 1, 5)
      .getValues();

    const data = values
      .filter(function(row) {
        return String(row[4]).toUpperCase() === "AKTIF";
      })
      .map(function(row) {
        return {
          program_id: row[0],
          nama_program: row[1],
          kategori_program: row[2],
          deskripsi: row[3],
          status: row[4]
        };
      });

    return {
      success: true,
      data: data
    };

  } catch (error) {
    return {
      success: false,
      message: error.message
    };
  }
}


/**
 * Menambahkan program baru.
 *
 * Akses: ADMIN
 */
function createProgram(token, payload) {
  try {
    validateProgramAccess_(token, true);

    if (!payload || typeof payload !== "object") {
      throw new Error("Data program tidak valid.");
    }

    const nama = String(payload.nama_program || "").trim();
    const kategori = String(payload.kategori_program || "").trim();
    const deskripsi = String(payload.deskripsi || "").trim();
    const status = String(payload.status || "AKTIF")
      .trim()
      .toUpperCase();

    if (!nama) {
      throw new Error("Nama program wajib diisi.");
    }

    if (!kategori) {
      throw new Error("Kategori program wajib diisi.");
    }

    if (status !== "AKTIF" && status !== "NONAKTIF") {
      throw new Error("Status program tidak valid.");
    }

    const lock = LockService.getScriptLock();
    lock.waitLock(10000);

    try {
      const db = getDatabase();
      const sheet = db.getSheetByName("MASTER_PROGRAM");

      if (!sheet) {
        throw new Error("Sheet MASTER_PROGRAM tidak ditemukan.");
      }

      const lastRow = sheet.getLastRow();

      if (lastRow >= 2) {
        const existingNames = sheet
          .getRange(2, 2, lastRow - 1, 1)
          .getValues();

        const duplicate = existingNames.some(function(row) {
          return String(row[0]).trim().toLowerCase() ===
            nama.toLowerCase();
        });

        if (duplicate) {
          throw new Error("Nama program sudah terdaftar.");
        }
      }

      const programId = "PRG-" +
        Utilities.getUuid().substring(0, 8).toUpperCase();

      sheet.appendRow([
        programId,
        nama,
        kategori,
        deskripsi,
        status
      ]);

      return {
        success: true,
        message: "Program berhasil ditambahkan.",
        program_id: programId
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


/**
 * Mengedit data program.
 *
 * Akses: ADMIN
 */
function editProgram(token, payload) {
  try {
    validateProgramAccess_(token, true);

    if (!payload || typeof payload !== "object") {
      throw new Error("Data program tidak valid.");
    }

    const programId = String(payload.program_id || "").trim();
    const nama = String(payload.nama_program || "").trim();
    const kategori = String(payload.kategori_program || "").trim();
    const deskripsi = String(payload.deskripsi || "").trim();
    const status = String(payload.status || "")
      .trim()
      .toUpperCase();

    if (!programId) {
      throw new Error("ID program tidak ditemukan.");
    }

    if (!nama) {
      throw new Error("Nama program wajib diisi.");
    }

    if (!kategori) {
      throw new Error("Kategori program wajib diisi.");
    }

    if (status !== "AKTIF" && status !== "NONAKTIF") {
      throw new Error("Status program tidak valid.");
    }

    const lock = LockService.getScriptLock();
    lock.waitLock(10000);

    try {
      const db = getDatabase();
      const sheet = db.getSheetByName("MASTER_PROGRAM");

      if (!sheet) {
        throw new Error("Sheet MASTER_PROGRAM tidak ditemukan.");
      }

      const lastRow = sheet.getLastRow();

      if (lastRow < 2) {
        throw new Error("Data program belum tersedia.");
      }

      const values = sheet
        .getRange(2, 1, lastRow - 1, 5)
        .getValues();

      let targetRow = -1;

      for (let i = 0; i < values.length; i++) {
        const currentId = String(values[i][0]).trim();

        if (currentId === programId) {
          targetRow = i + 2;
          continue;
        }

        const currentName = String(values[i][1])
          .trim()
          .toLowerCase();

        if (currentName === nama.toLowerCase()) {
          throw new Error("Nama program sudah digunakan program lain.");
        }
      }

      if (targetRow === -1) {
        throw new Error("Program tidak ditemukan.");
      }

      sheet.getRange(targetRow, 2, 1, 4).setValues([[
        nama,
        kategori,
        deskripsi,
        status
      ]]);

      return {
        success: true,
        message: "Data program berhasil diperbarui."
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
