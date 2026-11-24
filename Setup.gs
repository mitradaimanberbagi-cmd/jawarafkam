
/**************************************************
 * JAWARA FKAM
 * INITIAL ADMIN SETUP
 **************************************************/

function createInitialAdmin() {
  const props = PropertiesService.getScriptProperties();

  const email = String(props.getProperty("ADMIN_EMAIL") || "")
    .trim().toLowerCase();

  const nama = String(props.getProperty("ADMIN_NAME") || "").trim();
  const password = String(props.getProperty("ADMIN_PASSWORD") || "");

  if (!email || !nama || !password) {
    throw new Error("Data admin belum lengkap di Script Properties.");
  }

  if (password.length < 4) {
    throw new Error("Password admin minimal 4 karakter.");
  }

  const lock = LockService.getScriptLock();
  lock.waitLock(10000);

  try {
    const sh = getDatabase().getSheetByName("USERS");

    const existingAdmins = sh.getDataRange().getValues()
      .slice(1)
      .filter(row => row[5] === "ADMIN");

    if (existingAdmins.length > 0) {
      throw new Error("Admin pertama sudah dibuat.");
    }

    if (findUserByEmail_(email)) {
      throw new Error("Email sudah terdaftar.");
    }

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
      "ADMIN",
      "",
      "AKTIF",
      now,
      ""
    ]);

    // Hapus data setup sementara setelah berhasil.
    props.deleteProperty("ADMIN_EMAIL");
    props.deleteProperty("ADMIN_NAME");
    props.deleteProperty("ADMIN_PASSWORD");

    return {
      success: true,
      message: "Akun ADMIN JAWARA berhasil dibuat.",
      user_id: userId,
      email: email
    };

  } finally {
    lock.releaseLock();
  }
}
