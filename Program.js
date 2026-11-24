/**************************************************
 * JAWARA FKAM - MANAJEMEN PROGRAM & SUB PROGRAM
 * Versi gabungan dan kompatibel dengan modul Donatur.
 *
 * MASTER_PROGRAM:
 * program_id | nama_program | kategori_program | deskripsi
 * | status | sub_program
 **************************************************/

/**
 * Validasi sesi dan hak akses.
 * ADMIN: mengelola program.
 * ADMIN dan USER: membaca program aktif.
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
 * Membaca posisi kolom berdasarkan nama header.
 * Tidak bergantung pada urutan kolom di Google Sheets.
 */
function getProgramColumnMap_(sheet) {
  if (!sheet) {
    throw new Error("Sheet MASTER_PROGRAM tidak ditemukan.");
  }

  const lastColumn = sheet.getLastColumn();
  if (lastColumn < 1) {
    throw new Error("Header MASTER_PROGRAM belum tersedia.");
  }

  const headers = sheet
    .getRange(1, 1, 1, lastColumn)
    .getDisplayValues()[0]
    .map(function(value) {
      return String(value || "").trim().toLowerCase();
    });

  const map = {};
  headers.forEach(function(header, index) {
    if (header) map[header] = index;
  });

  const required = [
    "program_id",
    "nama_program",
    "kategori_program",
    "deskripsi",
    "status",
    "sub_program"
  ];

  const missing = required.filter(function(header) {
    return map[header] === undefined;
  });

  if (missing.length) {
    throw new Error(
      "Kolom " + missing.join(", ") +
      " belum tersedia di MASTER_PROGRAM. " +
      "Jalankan upgradeJawaraDonorProgramSchema() satu kali."
    );
  }

  return map;
}


/**
 * Helper utama yang digunakan Donatur.gs.
 * Mengembalikan sheet, map, dan values.
 */
function getProgramRows_() {
  const sheet = getDatabase().getSheetByName("MASTER_PROGRAM");
  if (!sheet) {
    throw new Error("Sheet MASTER_PROGRAM tidak ditemukan.");
  }

  const map = getProgramColumnMap_(sheet);
  const values = sheet.getLastRow() < 2
    ? []
    : sheet.getRange(
        2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()
      ).getValues();

  return {
    sheet: sheet,
    map: map,
    values: values
  };
}


/**
 * ADMIN: mengambil seluruh program untuk halaman Manajemen Program.
 */
function getProgramManagementData(token) {
  try {
    validateProgramAccess_(token, true);

    const result = getProgramRows_();
    const m = result.map;

    const data = result.values.map(function(row) {
      return {
        program_id: row[m.program_id],
        nama_program: row[m.nama_program],
        sub_program: row[m.sub_program] || "",
        kategori_program: row[m.kategori_program],
        deskripsi: row[m.deskripsi],
        status: row[m.status]
      };
    });

    return {
      success: true,
      data: data
    };

  } catch (error) {
    console.error("getProgramManagementData: " + (error.stack || error.message));
    return {
      success: false,
      message: error.message || String(error)
    };
  }
}


/**
 * ADMIN dan USER: mengambil program aktif untuk dropdown.
 * Mengirim dua set nama properti agar kompatibel dengan
 * UI yang menggunakan format lama maupun format baru.
 */
function getActivePrograms(token) {
  try {
    validateProgramAccess_(token, false);

    const result = getProgramRows_();
    const m = result.map;

    const data = result.values
      .filter(function(row) {
        return String(row[m.status] || "").trim().toUpperCase() === "AKTIF";
      })
      .map(function(row) {
        const id = String(row[m.program_id] || "");
        const nama = String(row[m.nama_program] || "");
        const subProgram = String(row[m.sub_program] || "");
        const kategori = String(row[m.kategori_program] || "");

        return {
          // Properti format baru
          id: id,
          nama: nama,

          // Alias format lama agar UI lama tetap berfungsi
          program_id: id,
          nama_program: nama,

          sub_program: subProgram,
          kategori_program: kategori,
          status: String(row[m.status] || "")
        };
      });

    return {
      success: true,
      data: data
    };

  } catch (error) {
    console.error("getActivePrograms: " + (error.stack || error.message));
    return {
      success: false,
      message: error.message || String(error)
    };
  }
}


/**
 * ADMIN: menambahkan program dan sub program.
 */
function createProgram(token, payload) {
  try {
    validateProgramAccess_(token, true);
    payload = payload || {};

    const nama = String(payload.nama_program || "").trim();
    const subProgram = String(payload.sub_program || "").trim();
    const kategori = String(payload.kategori_program || "").trim();
    const deskripsi = String(payload.deskripsi || "").trim();
    const status = String(payload.status || "AKTIF").trim().toUpperCase();

    if (!nama) throw new Error("Nama program wajib diisi.");
    if (!kategori) throw new Error("Kategori program wajib diisi.");
    if (status !== "AKTIF" && status !== "NONAKTIF") {
      throw new Error("Status program tidak valid.");
    }

    const lock = LockService.getScriptLock();
    lock.waitLock(10000);

    try {
      const result = getProgramRows_();
      const m = result.map;

      const duplicate = result.values.some(function(row) {
        return String(row[m.nama_program] || "").trim().toLowerCase() === nama.toLowerCase() &&
          String(row[m.sub_program] || "").trim().toLowerCase() === subProgram.toLowerCase();
      });

      if (duplicate) {
        throw new Error("Kombinasi nama program dan sub program sudah terdaftar.");
      }

      const newRow = new Array(result.sheet.getLastColumn()).fill("");
      const programId = "PRG-" +
        Utilities.getUuid().substring(0, 8).toUpperCase();

      newRow[m.program_id] = programId;
      newRow[m.nama_program] = nama;
      newRow[m.sub_program] = subProgram;
      newRow[m.kategori_program] = kategori;
      newRow[m.deskripsi] = deskripsi;
      newRow[m.status] = status;

      result.sheet.appendRow(newRow);

      return {
        success: true,
        message: "Program berhasil ditambahkan.",
        program_id: programId
      };

    } finally {
      lock.releaseLock();
    }

  } catch (error) {
    console.error("createProgram: " + (error.stack || error.message));
    return {
      success: false,
      message: error.message || String(error)
    };
  }
}


/**
 * ADMIN: mengedit program dan sub program.
 */
function editProgram(token, payload) {
  try {
    validateProgramAccess_(token, true);
    payload = payload || {};

    const id = String(payload.program_id || "").trim();
    const nama = String(payload.nama_program || "").trim();
    const subProgram = String(payload.sub_program || "").trim();
    const kategori = String(payload.kategori_program || "").trim();
    const deskripsi = String(payload.deskripsi || "").trim();
    const status = String(payload.status || "").trim().toUpperCase();

    if (!id) throw new Error("ID program tidak ditemukan.");
    if (!nama) throw new Error("Nama program wajib diisi.");
    if (!kategori) throw new Error("Kategori program wajib diisi.");
    if (status !== "AKTIF" && status !== "NONAKTIF") {
      throw new Error("Status program tidak valid.");
    }

    const lock = LockService.getScriptLock();
    lock.waitLock(10000);

    try {
      const result = getProgramRows_();
      const m = result.map;
      let targetRow = -1;

      for (let i = 0; i < result.values.length; i++) {
        const row = result.values[i];
        const currentId = String(row[m.program_id] || "").trim();

        if (currentId === id) {
          targetRow = i + 2;
          continue;
        }

        const sameName = String(row[m.nama_program] || "").trim().toLowerCase() === nama.toLowerCase();
        const sameSubProgram = String(row[m.sub_program] || "").trim().toLowerCase() === subProgram.toLowerCase();

        if (sameName && sameSubProgram) {
          throw new Error("Kombinasi nama program dan sub program sudah digunakan.");
        }
      }

      if (targetRow < 0) {
        throw new Error("Program tidak ditemukan.");
      }

      const target = result.sheet
        .getRange(targetRow, 1, 1, result.sheet.getLastColumn())
        .getValues()[0];

      target[m.nama_program] = nama;
      target[m.sub_program] = subProgram;
      target[m.kategori_program] = kategori;
      target[m.deskripsi] = deskripsi;
      target[m.status] = status;

      result.sheet
        .getRange(targetRow, 1, 1, target.length)
        .setValues([target]);

      return {
        success: true,
        message: "Data program berhasil diperbarui."
      };

    } finally {
      lock.releaseLock();
    }

  } catch (error) {
    console.error("editProgram: " + (error.stack || error.message));
    return {
      success: false,
      message: error.message || String(error)
    };
  }
}


/**
 * Jalankan SATU KALI setelah memasang kode baru.
 * Menambahkan kolom baru di ujung sheet agar kolom lama
 * tidak bergeser dan data lama tidak dihapus.
 */
function upgradeJawaraDonorProgramSchema() {
  try {
    const db = getDatabase();

    const specs = [
      {
        name: "MASTER_PROGRAM",
        headers: ["sub_program"]
      },
      {
        name: "DONATUR",
        headers: ["panggilan", "program_id", "sub_program"]
      },
      {
        name: "DONASI",
        headers: ["sub_program"]
      }
    ];

    const report = [];

    specs.forEach(function(spec) {
      const sheet = db.getSheetByName(spec.name);
      if (!sheet) {
        throw new Error("Sheet " + spec.name + " tidak ditemukan.");
      }

      let lastColumn = sheet.getLastColumn();
      if (lastColumn < 1) {
        throw new Error("Header sheet " + spec.name + " belum tersedia.");
      }

      const headers = sheet
        .getRange(1, 1, 1, lastColumn)
        .getDisplayValues()[0]
        .map(function(value) {
          return String(value || "").trim().toLowerCase();
        });

      spec.headers.forEach(function(header) {
        if (headers.indexOf(header.toLowerCase()) === -1) {
          lastColumn += 1;
          sheet.getRange(1, lastColumn).setValue(header);
          headers.push(header.toLowerCase());
          report.push(spec.name + ": kolom " + header + " ditambahkan");
        } else {
          report.push(spec.name + ": kolom " + header + " sudah ada");
        }
      });
    });

    return {
      success: true,
      message: "Struktur database berhasil diperiksa.",
      details: report
    };

  } catch (error) {
    console.error("upgradeJawaraDonorProgramSchema: " + (error.stack || error.message));
    return {
      success: false,
      message: error.message || String(error)
    };
  }
}
