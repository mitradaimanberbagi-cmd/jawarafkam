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
  const user=validateSession(token);
  if(!user||user.role!=="USER") throw new Error("Fitur ini hanya untuk akun da'i.");
  const db=getDatabase();
  function getActiveOptions(sheetName,idColumn,nameColumn) {
    const sh=db.getSheetByName(sheetName);
    if(!sh||sh.getLastRow()<2) return [];
    const values=sh.getDataRange().getValues(), headers=values[0].map(function(v){return String(v).trim().toLowerCase();});
    const idIndex=headers.indexOf(idColumn), nameIndex=headers.indexOf(nameColumn), statusIndex=headers.indexOf("status");
    if(idIndex<0||nameIndex<0||statusIndex<0) throw new Error("Struktur sheet "+sheetName+" tidak sesuai.");
    return values.slice(1).filter(function(r){return String(r[statusIndex]).toUpperCase()==="AKTIF";})
      .map(function(r){return {id:String(r[idIndex]),nama:String(r[nameIndex])};});
  }
  const methods=getActiveOptions("MASTER_METODE","metode_id","nama_metode")
    .filter(function(x){return ["tunai","transfer bank"].includes(x.nama.trim().toLowerCase());});
  const programsResult=getProgramRows_(), pm=programsResult.map;
  const programs=programsResult.values.filter(function(r){return String(r[pm.status]).toUpperCase()==="AKTIF";})
    .map(function(r){return {id:String(r[pm.program_id]),nama:String(r[pm.nama_program]),sub_program:String(r[pm.sub_program]||"")};});
  return {
    success:true,
    jenisDonasi:getActiveOptions("MASTER_JENIS_DONASI","jenis_donasi_id","nama_jenis"),
    program:programs,
    metode:methods
  };
}

function isActiveDonationMaster_(sheet,idColumn,id) {
  if(!id) return false;
  if(!sheet||sheet.getLastRow()<2) return false;
  const rows=sheet.getDataRange().getValues(), headers=rows[0].map(function(v){return String(v).trim().toLowerCase();});
  const idIndex=headers.indexOf(idColumn), statusIndex=headers.indexOf("status");
  if(idIndex<0||statusIndex<0) throw new Error("Struktur sheet master tidak sesuai.");
  return rows.slice(1).some(function(r){return String(r[idIndex])===String(id)&&String(r[statusIndex]).toUpperCase()==="AKTIF";});
}

function createDonationEntry(token,payload) {
  const user=validateSession(token);
  if(!user||user.role!=="USER") throw new Error("Hanya akun da'i yang dapat memasukkan donasi.");
  payload=payload||{};
  const tanggal=String(payload.tanggal_donasi||"").trim();
  const donaturId=String(payload.donatur_id||"").trim();
  const jenisDonatur=String(payload.jenis_donatur||"").trim();
  const jenisDonasiId=String(payload.jenis_donasi_id||"").trim();
  let programId=String(payload.program_id||"").trim();
  let subProgram=String(payload.sub_program||"").trim();
  const metodeId=String(payload.metode_id||"").trim();
  const keterangan=String(payload.keterangan||"").trim();
  const nominal=Number(payload.nominal);

  if(!/^\d{4}-\d{2}-\d{2}$/.test(tanggal)) throw new Error("Tanggal donasi tidak valid.");
  const p=tanggal.split("-").map(Number), tanggalDonasi=new Date(p[0],p[1]-1,p[2]);
  if(tanggalDonasi.getFullYear()!==p[0]||tanggalDonasi.getMonth()!==p[1]-1||tanggalDonasi.getDate()!==p[2]) throw new Error("Tanggal donasi tidak valid.");
  if(!donaturId) throw new Error("Pilih donatur dari daftar pencarian.");
  if(!["INDIVIDU","LEMBAGA"].includes(jenisDonatur)) throw new Error("Jenis donatur tidak valid.");
  if(!jenisDonasiId||!metodeId) throw new Error("Jenis donasi dan metode donasi wajib dipilih.");
  if(!Number.isFinite(nominal)||nominal<=0) throw new Error("Nominal donasi harus lebih besar dari nol.");

  const db=getDatabase();
  const donorSheet=db.getSheetByName("DONATUR"), donationSheet=db.getSheetByName("DONASI");
  const jenisSheet=db.getSheetByName("MASTER_JENIS_DONASI"), metodeSheet=db.getSheetByName("MASTER_METODE");
  if(!donorSheet||!donationSheet||!jenisSheet||!metodeSheet) throw new Error("Salah satu sheet database tidak ditemukan.");
  const dm=getDonorColumnMap_(donorSheet);
  const donorRows=donorSheet.getDataRange().getValues();
  let donor=null;
  for(let i=1;i<donorRows.length;i++) if(String(donorRows[i][dm.donatur_id])===donaturId){donor=donorRows[i];break;}
  if(!donor) throw new Error("Data donatur yang dipilih tidak ditemukan. Muat ulang daftar donatur.");
  const namaDonatur=String(donor[dm.nama]||"");
  const noHp=String(donor[dm.no_hp]||"");
  const donorProgram=String(donor[dm.program_id]||"").trim();
  const donorSubProgram=String(donor[dm.sub_program]||"").trim();
  if(!programId) programId=donorProgram;
  if(!subProgram) subProgram=donorSubProgram;

  if(!isActiveDonationMaster_(jenisSheet,"jenis_donasi_id",jenisDonasiId)) throw new Error("Jenis donasi tidak ditemukan atau tidak aktif.");
  if(!isActiveDonationMaster_(metodeSheet,"metode_id",metodeId)) throw new Error("Metode donasi tidak ditemukan atau tidak aktif.");
  const methodRows=metodeSheet.getDataRange().getValues(), mh=methodRows[0].map(function(v){return String(v).trim().toLowerCase();});
  const mi=mh.indexOf("metode_id"), mn=mh.indexOf("nama_metode");
  const methodRow=methodRows.slice(1).find(function(r){return String(r[mi])===metodeId;});
  if(!methodRow) throw new Error("Metode donasi tidak ditemukan.");
  const methodName=String(methodRow[mn]||"").trim().toLowerCase();
  if(!["tunai","transfer bank"].includes(methodName)) throw new Error("Metode donasi hanya boleh Tunai atau Transfer Bank.");

  let programName="";
  if(programId) {
    const pr=getProgramRows_(), pm=pr.map;
    const selected=pr.values.find(function(r){return String(r[pm.program_id])===programId&&String(r[pm.status]).toUpperCase()==="AKTIF";});
    if(!selected) throw new Error("Program tidak ditemukan atau tidak aktif.");
    programName=String(selected[pm.nama_program]||"");
    if(!subProgram) subProgram=String(selected[pm.sub_program]||"");
  }

  const proof=payload.bukti_transfer||null;
  if(methodName==="transfer bank"&&!proof) throw new Error("Bukti transfer wajib diunggah untuk metode Transfer Bank.");
  const lock=LockService.getScriptLock(); lock.waitLock(10000);
  let uploadedFile=null;
  try {
    let proofFileId="";
    if(proof) {
      const allowed=["image/jpeg","image/png","image/webp","application/pdf"];
      const mime=String(proof.mimeType||"").toLowerCase();
      const name=String(proof.fileName||"bukti-transfer").replace(/[\\/:*?"<>|]/g,"_");
      const base64=String(proof.base64||"");
      if(!allowed.includes(mime)) throw new Error("Format bukti transfer harus JPG, PNG, WEBP, atau PDF.");
      if(!base64) throw new Error("File bukti transfer tidak terbaca.");
      const bytes=Utilities.base64Decode(base64);
      if(bytes.length>5*1024*1024) throw new Error("Ukuran bukti transfer maksimal 5 MB.");
      const blob=Utilities.newBlob(bytes,mime,name);
      const folderId=PropertiesService.getScriptProperties().getProperty("JAWARA_UPLOAD_FOLDER_ID");
      uploadedFile=folderId?DriveApp.getFolderById(folderId).createFile(blob):DriveApp.createFile(blob);
      proofFileId=uploadedFile.getId();
    }

    const now=new Date(), donasiId="DNS-"+Utilities.getUuid().slice(0,8).toUpperCase();
    const headers=donationSheet.getRange(1,1,1,donationSheet.getLastColumn()).getValues()[0]
      .map(function(v){return String(v).trim().toLowerCase();});
    const row=new Array(donationSheet.getLastColumn()).fill("");
    function put(header,value){const index=headers.indexOf(header);if(index<0)throw new Error("Kolom "+header+" belum tersedia pada sheet DONASI.");row[index]=value;}
    put("donasi_id",donasiId); put("tanggal_donasi",tanggalDonasi); put("waktu_input",now);
    put("user_id",user.user_id); put("donatur_id",donaturId); put("nama_donatur_snapshot",namaDonatur);
    put("no_hp_snapshot",noHp); put("jenis_donatur",jenisDonatur); put("jenis_donasi_id",jenisDonasiId);
    put("program_id",programId); put("nominal",nominal); put("metode_id",metodeId);
    put("keterangan",keterangan); put("bukti_donasi_file_id",proofFileId); put("status_transaksi","TERCATAT");
    put("created_at",now); put("cancelled_at",""); put("cancelled_by",""); put("cancel_reason","");
    if(headers.indexOf("sub_program")>=0) put("sub_program",subProgram);
    donationSheet.appendRow(row);
    return {success:true,message:"Donasi berhasil disimpan.",donasi_id:donasiId,bukti_donasi_file_id:proofFileId,nama_program:programName,sub_program:subProgram};
  } catch(error) {
    if(uploadedFile){try{uploadedFile.setTrashed(true);}catch(cleanupError){}}
    throw error;
  } finally {lock.releaseLock();}
}
