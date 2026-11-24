
/**************************************************
 * JAWARA FKAM
 * USER MANAGEMENT
 **************************************************/

/**
 * Validasi format email.
 */
function isValidEmail_(email) {
  return /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email);
}


/**
 * Admin membuat akun da'i.
 */
function createUser(token, payload) {
  const admin = validateSession(token);

  if (admin.role !== "ADMIN") {
    throw new Error("Akses ditolak. Khusus admin.");
  }

  payload = payload || {};

  const nama = String(payload.nama || "").trim();
  const email = String(payload.email || "").trim().toLowerCase();
  const password = String(payload.password || "");
  const noHp = String(payload.no_hp || "").trim();

  if (!nama || !email || !password) {
    throw new Error("Nama, email, dan password wajib diisi.");
  }

  if (!isValidEmail_(email)) {
    throw new Error("Format email tidak valid.");
  }

  if (password.length < 4) {
    throw new Error("Password minimal 4 karakter.");
  }

  const lock = LockService.getScriptLock();
  lock.waitLock(10000);

  try {
    if (findUserByEmail_(email)) {
      throw new Error("Email sudah digunakan.");
    }

    const sh = getDatabase().getSheetByName("USERS");

    const userId = "USR-" + Utilities.getUuid().slice(0, 8);
    const salt = Utilities.getUuid() + Utilities.getUuid();
    const passwordHash = hashPassword_(password, salt);
    const now = new Date();

    sh.appendRow([
      userId,
      nama,
      email,
      passwordHash,
      salt,
      "USER",
      noHp,
      "AKTIF",
      now,
      ""
    ]);

    return {
      success: true,
      message: "Akun da'i berhasil dibuat.",
      user_id: userId
    };

  } finally {
    lock.releaseLock();
  }
}


/**
 * Admin melihat daftar user.
 * Hanya menampilkan akun dengan role USER.
 */
function getUsers(token) {
  const admin = validateSession(token);

  if (admin.role !== "ADMIN") {
    throw new Error("Akses ditolak. Khusus admin.");
  }

  const sh = getDatabase().getSheetByName("USERS");
  const data = sh.getDataRange().getValues();

  return data
    .slice(1)
    .filter(row => row[0] && row[5] === "USER")
    .map(row => ({
      user_id: row[0],
      nama: row[1],
      email: row[2],
      role: row[5],
      no_hp: row[6],
      status: row[7]
    }));
}


/**
 * Admin mengedit data dan mengatur ulang password da'i.
 *
 * Password dikosongkan:
 * Password lama tetap berlaku.
 *
 * Password diisi:
 * Password baru akan disimpan.
 *
 * Edit hanya berlaku untuk akun USER.
 */
function editUser(token, payload) {
  const admin = validateSession(token);

  if (admin.role !== "ADMIN") {
    throw new Error("Akses ditolak. Khusus admin.");
  }

  payload = payload || {};

  const userId = String(payload.user_id || "").trim();
  const nama = String(payload.nama || "").trim();
  const email = String(payload.email || "").trim().toLowerCase();
  const noHp = String(payload.no_hp || "").trim();
  const status = String(payload.status || "").trim();
  const newPassword = String(payload.password || "");

  if (!userId || !nama || !email) {
    throw new Error("Data wajib belum lengkap.");
  }

  if (!isValidEmail_(email)) {
    throw new Error("Format email tidak valid.");
  }

  if (!["AKTIF", "NONAKTIF"].includes(status)) {
    throw new Error("Status akun tidak valid.");
  }

  if (newPassword && newPassword.length < 4) {
    throw new Error("Password baru minimal 4 karakter.");
  }

  const lock = LockService.getScriptLock();
  lock.waitLock(10000);

  try {
    const ss = getDatabase();
    const sh = ss.getSheetByName("USERS");
    const data = sh.getDataRange().getValues();

    let targetRow = -1;

    // Cari akun berdasarkan ID dan periksa email ganda.
    for (let i = 1; i < data.length; i++) {

      if (String(data[i][0]) === userId) {
        targetRow = i + 1;
      }

      if (
        String(data[i][2]).trim().toLowerCase() === email &&
        String(data[i][0]) !== userId
      ) {
        throw new Error("Email sudah digunakan akun lain.");
      }
    }

    if (targetRow === -1) {
      throw new Error("User tidak ditemukan.");
    }

    // Edit hanya diperbolehkan untuk akun USER.
    if (data[targetRow - 1][5] !== "USER") {
      throw new Error("Edit hanya berlaku untuk akun da'i.");
    }

    // Perbarui data profil.
    sh.getRange(targetRow, 2).setValue(nama);
    sh.getRange(targetRow, 3).setValue(email);
    sh.getRange(targetRow, 7).setValue(noHp);
    sh.getRange(targetRow, 8).setValue(status);

    // Reset password hanya jika kolom diisi.
    const passwordChanged = newPassword.length > 0;

    if (passwordChanged) {
      const newSalt = Utilities.getUuid() + Utilities.getUuid();
      const newHash = hashPassword_(newPassword, newSalt);

      sh.getRange(targetRow, 4).setValue(newHash);
      sh.getRange(targetRow, 5).setValue(newSalt);
    }

    // Cabut sesi jika akun dinonaktifkan atau password diganti.
    if (status === "NONAKTIF" || passwordChanged) {

      const sessions = ss.getSheetByName("SESSIONS");
      const sessionData = sessions.getDataRange().getValues();

      for (let i = 1; i < sessionData.length; i++) {

        if (
          String(sessionData[i][1]) === userId &&
          sessionData[i][5] === "AKTIF"
        ) {
          sessions.getRange(i + 1, 6).setValue("DICABUT");
        }
      }
    }

    return {
      success: true,
      message: passwordChanged
        ? "Data da'i dan password berhasil diperbarui."
        : "Data da'i berhasil diperbarui."
    };

  } finally {
    lock.releaseLock();
  }
}
