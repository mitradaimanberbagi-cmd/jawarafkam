
/**************************************************
 * JAWARA FKAM
 * AUTHENTICATION SYSTEM
 **************************************************/

const SESSION_DURATION_HOURS = 12;
const PASSWORD_ITERATIONS = 10000;

/**
 * Membuat hash password dengan salt.
 */
function hashPassword_(password, salt) {
  let hash = Utilities.computeDigest(
    Utilities.DigestAlgorithm.SHA_256,
    salt + ":" + password,
    Utilities.Charset.UTF_8
  );

  for (let i = 1; i < PASSWORD_ITERATIONS; i++) {
    hash = Utilities.computeDigest(
      Utilities.DigestAlgorithm.SHA_256,
      hash
    );
  }

  return hash.map(function(byte) {
    return ("0" + ((byte + 256) % 256).toString(16)).slice(-2);
  }).join("");
}

/**
 * Hash token sesi.
 */
function hashToken_(token) {
  const bytes = Utilities.computeDigest(
    Utilities.DigestAlgorithm.SHA_256,
    token,
    Utilities.Charset.UTF_8
  );

  return bytes.map(function(byte) {
    return ("0" + ((byte + 256) % 256).toString(16)).slice(-2);
  }).join("");
}

/**
 * Mencari akun berdasarkan email.
 */
function findUserByEmail_(email) {
  const sh = getDatabase().getSheetByName("USERS");
  const data = sh.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    if (
      String(data[i][2]).trim().toLowerCase() ===
      String(email).trim().toLowerCase()
    ) {
      return {
        row: i + 1,
        data: data[i]
      };
    }
  }

  return null;
}

/**
 * Login pengguna.
 */
function login(email, password) {
  email = String(email || "").trim().toLowerCase();
  password = String(password || "");

  if (!email || !password) {
    throw new Error("Email dan password wajib diisi.");
  }

  const found = findUserByEmail_(email);

  if (!found) {
    throw new Error("Email atau password tidak sesuai.");
  }

  const user = found.data;

  if (user[7] !== "AKTIF") {
    throw new Error("Akun tidak aktif. Hubungi admin.");
  }

  const calculatedHash = hashPassword_(
    password,
    String(user[4])
  );

  if (calculatedHash !== String(user[3])) {
    throw new Error("Email atau password tidak sesuai.");
  }

  const lock = LockService.getScriptLock();
  lock.waitLock(10000);

  try {
    const ss = getDatabase();
    const userSheet = ss.getSheetByName("USERS");
    const sessionSheet = ss.getSheetByName("SESSIONS");

    const token = Utilities.getUuid() + Utilities.getUuid();
    const now = new Date();
    const expires = new Date(
      now.getTime() + SESSION_DURATION_HOURS * 60 * 60 * 1000
    );

    sessionSheet.appendRow([
      Utilities.getUuid(),
      user[0],
      hashToken_(token),
      now,
      expires,
      "AKTIF"
    ]);

    userSheet.getRange(found.row, 10).setValue(now);

    return {
      success: true,
      token: token,
      user: {
        user_id: user[0],
        nama: user[1],
        email: user[2],
        role: user[5]
      }
    };

  } finally {
    lock.releaseLock();
  }
}

/**
 * Validasi sesi pengguna.
 */

/**
 * Validasi sesi pengguna.
 */
function validateSession(token) {
  if (!token) {
    throw new Error("Sesi tidak ditemukan.");
  }

  const sh = getDatabase().getSheetByName("SESSIONS");
  const data = sh.getDataRange().getValues();
  const tokenHash = hashToken_(token);
  const now = new Date();

  for (let i = 1; i < data.length; i++) {
    if (
      String(data[i][2]) === tokenHash &&
      data[i][5] === "AKTIF"
    ) {
      if (
        new Date(data[i][4]).getTime() <= now.getTime()
      ) {
        sh.getRange(i + 1, 6).setValue("DICABUT");
        throw new Error("Sesi sudah kedaluwarsa.");
      }

      const user = getUserById_(data[i][1]);

      if (!user || user.status !== "AKTIF") {
        sh.getRange(i + 1, 6).setValue("DICABUT");
        throw new Error("Akun tidak aktif.");
      }

      if (
        user.role !== "ADMIN" &&
        user.role !== "USER"
      ) {
        sh.getRange(i + 1, 6).setValue("DICABUT");
        throw new Error("Role akun tidak valid.");
      }

      return user;
    }
  }

  throw new Error("Sesi tidak valid.");
}

/**
 * Mencari user berdasarkan ID.
 */
function getUserById_(userId) {
  const sh = getDatabase().getSheetByName("USERS");
  const data = sh.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(userId)) {
      return {
        user_id: data[i][0],
        nama: data[i][1],
        email: data[i][2],
        role: data[i][5],
        status: data[i][7]
      };
    }
  }

  return null;
}

/**
 * Logout dan cabut sesi.
 */
function logout(token) {
  if (!token) {
    return { success: true };
  }

  const sh = getDatabase().getSheetByName("SESSIONS");
  const data = sh.getDataRange().getValues();
  const tokenHash = hashToken_(token);

  for (let i = 1; i < data.length; i++) {
    if (
      String(data[i][2]) === tokenHash &&
      data[i][5] === "AKTIF"
    ) {
      sh.getRange(i + 1, 6).setValue("DICABUT");
      break;
    }
  }

  return { success: true };

  
}