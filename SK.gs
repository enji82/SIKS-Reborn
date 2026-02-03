/* ======================================================================
   SK.gs - LOGIKA BACKEND SIABA SK
   Variabel Global (SPREADSHEET_IDS & FOLDER_CONFIG) diambil dari Code.gs
   ====================================================================== */

/* ======================================================================
   HELPER FUNCTIONS
   ====================================================================== */
function handleError(context, error) {
  Logger.log("ERROR [" + context + "]: " + error);
  rekamCCTV("ERROR " + context, error.toString()); // Integrasi CCTV
  return { success: false, message: error.message || error.toString() };
}

function getOrCreateFolder(parentFolder, folderName) {
  var folders = parentFolder.getFoldersByName(folderName);
  return folders.hasNext() ? folders.next() : parentFolder.createFolder(folderName);
}

/* ======================================================================
   CORE: PROSES SIMPAN DATA BARU (INSERT)
   ====================================================================== */
function processManualForm(formData) {
  try {
    // Menggunakan ID dari Code.gs
    const ss = SpreadsheetApp.openById(SPREADSHEET_IDS.SK_DATA);
    const sheet = ss.getSheetByName("Unggah_SK");
    
    // Setup Folder (Rapi dengan Subfolder)
    // PERBAIKAN: Menggunakan FOLDER_CONFIG.MAIN_SK
    const mainFolder = DriveApp.getFolderById(FOLDER_CONFIG.MAIN_SK);
    const folderTahun = getOrCreateFolder(mainFolder, formData.tahunAjaran.replace(/\//g, '-'));
    const targetFolder = getOrCreateFolder(folderTahun, formData.semester);
    
    // Penamaan File
    const namaFile = `${formData.namaSd} - ${formData.tahunAjaran.replace(/\//g,'-')} - ${formData.semester} - ${formData.kriteriaSk} - ${formData.nomorSk}.pdf`;
    
    const blob = Utilities.newBlob(Utilities.base64Decode(formData.fileData.data), formData.fileData.mimeType, namaFile);
    const file = targetFolder.createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

    // Insert Database (Sesuai Kolom A-Q)
    sheet.appendRow([
      new Date(),             // A: Tgl Unggah
      formData.namaSd,        // B
      formData.tahunAjaran,   // C
      formData.semester,      // D
      "'" + formData.nomorSk, // E (Paksa Text)
      "'" + formData.tanggalSk, // F (Paksa Text)
      formData.kriteriaSk,    // G
      file.getUrl(),          // H
      formData.userInput,     // I
      "Diproses",             // J
      "", "", "", "", ""      // K-O Kosong
    ]);

    return { success: true, message: "Data SK berhasil disimpan." };
  } catch (e) { return handleError('processManualForm', e); }
}

/* ======================================================================
   CORE: UPDATE DATA (EDIT) - FIX STATUS & FOLDER CONFIG
   ====================================================================== */
function simpanPerubahanSK(form) {
  try {
    rekamCCTV("START EDIT", "No SK: " + form.nomorSk);

    var ss = SpreadsheetApp.openById(SPREADSHEET_IDS.SK_DATA);
    var sheet = ss.getSheetByName("Unggah_SK");
    var rowIdx = parseInt(form.editRowId);

    if (isNaN(rowIdx)) throw "Row ID Invalid";

    // MAPPING KOLOM (SESUAI REQUEST BAPAK: A=1 ... J=10 ... Q=17)
    var KOLOM = {
      NAMA_SD:   2,  // B
      TAHUN:     3,  // C
      SEMESTER:  4,  // D
      NO_SK:     5,  // E
      TGL_SK:    6,  // F
      KRITERIA:  7,  // G
      FILE_URL:  8,  // H
      STATUS:    10, // J
      TGL_UPD:   11, // K
      USER_UPD:  12  // L
    };

    // 1. UPDATE IDENTITAS (SAFE UPDATE)
    if (form.namaSd && form.namaSd !== "") sheet.getRange(rowIdx, KOLOM.NAMA_SD).setValue(form.namaSd);
    if (form.tahunAjaran && form.tahunAjaran !== "") sheet.getRange(rowIdx, KOLOM.TAHUN).setValue(form.tahunAjaran);
    if (form.semester && form.semester !== "") sheet.getRange(rowIdx, KOLOM.SEMESTER).setValue(form.semester);

    // 2. UPDATE DATA SK (INTI)
    sheet.getRange(rowIdx, KOLOM.NO_SK).setValue(form.nomorSk);
    sheet.getRange(rowIdx, KOLOM.TGL_SK).setValue("'" + form.tanggalSk); 
    sheet.getRange(rowIdx, KOLOM.KRITERIA).setValue(form.kriteriaSk);

    // 3. UPDATE FILE (JIKA UPLOAD BARU)
    if (form.fileData && form.fileData.data) {
       // PERBAIKAN: Gunakan FOLDER_CONFIG.MAIN_SK (Bukan FOLDER_IDS)
       const mainFolder = DriveApp.getFolderById(FOLDER_CONFIG.MAIN_SK);
       
       // Logika Subfolder (Konsisten dengan Tambah)
       var thn = (form.tahunAjaran && form.tahunAjaran !== "") ? form.tahunAjaran : sheet.getRange(rowIdx, KOLOM.TAHUN).getValue();
       var sem = (form.semester && form.semester !== "") ? form.semester : sheet.getRange(rowIdx, KOLOM.SEMESTER).getValue();
       
       const folderTahun = getOrCreateFolder(mainFolder, thn.toString().replace(/\//g, '-'));
       const targetFolder = getOrCreateFolder(folderTahun, sem);

       // Penamaan File
       var namaSdFix = (form.namaSd && form.namaSd !== "") ? form.namaSd : sheet.getRange(rowIdx, KOLOM.NAMA_SD).getValue();
       const namaFile = `${namaSdFix} - ${thn.toString().replace(/\//g,'-')} - ${sem} - ${form.kriteriaSk} - ${form.nomorSk}.pdf`;

       var blob = Utilities.newBlob(Utilities.base64Decode(form.fileData.data), form.fileData.mimeType, namaFile);
       var file = targetFolder.createFile(blob);
       file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
       
       sheet.getRange(rowIdx, KOLOM.FILE_URL).setValue(file.getUrl());
       rekamCCTV("UPLOAD", "File baru tersimpan: " + file.getUrl());
    }

    // 4. RESET STATUS JADI DIPROSES
    sheet.getRange(rowIdx, KOLOM.STATUS).setValue("Diproses");

    // 5. METADATA UPDATE
    sheet.getRange(rowIdx, KOLOM.TGL_UPD).setValue(new Date());
    sheet.getRange(rowIdx, KOLOM.USER_UPD).setValue(form.userUpdate);

    rekamCCTV("SUKSES", "Data baris " + rowIdx + " berhasil diupdate.");
    return { success: true, message: "Data berhasil diperbarui." };

  } catch (e) {
    rekamCCTV("ERROR", e.toString());
    return { success: false, message: "Error Server: " + e.toString() };
  }
}

/* ======================================================================
   CORE: GET DATA LIST
   ====================================================================== */
function getDaftarSK() {
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_IDS.SK_DATA);
    var sheet = ss.getSheetByName("Unggah_SK");
    var data = sheet.getDataRange().getValues();
    var result = [];
    
    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      
      // Safety check tanggal & Null
      if (!row[1]) continue; // Skip jika Nama SD Kosong

      var tglUnggah = (row[0] instanceof Date) ? Utilities.formatDate(row[0], Session.getScriptTimeZone(), "dd-MM-yyyy HH:mm") : row[0];
      var tglUpdate = (row[10] instanceof Date) ? Utilities.formatDate(row[10], Session.getScriptTimeZone(), "dd-MM-yyyy HH:mm") : row[10];
      var tglVerval = (row[12] instanceof Date) ? Utilities.formatDate(row[12], Session.getScriptTimeZone(), "dd-MM-yyyy HH:mm") : row[12];
      
      var tglSkRaw = row[5];
      var tglSkISO = "", tglSkDisplay = "";
      
      if (tglSkRaw instanceof Date) {
          tglSkISO = Utilities.formatDate(tglSkRaw, Session.getScriptTimeZone(), "yyyy-MM-dd");
          tglSkDisplay = Utilities.formatDate(tglSkRaw, Session.getScriptTimeZone(), "dd-MM-yyyy");
      } else {
          tglSkISO = tglSkRaw; tglSkDisplay = tglSkRaw;
      }

      result.push({
        rowBaris: i + 1,
        tglUnggah: tglUnggah,
        namaSd: row[1], tahun: row[2], semester: row[3], noSk: row[4],
        tglSk: tglSkISO, tglSkDisplay: tglSkDisplay,
        kriteria: row[6], fileUrl: row[7], userInput: row[8], status: row[9],
        tglUpdate: tglUpdate, userUpdate: row[11],
        tglVerval: tglVerval, verifikator: row[13], keterangan: row[14]
      });
    }
    return result;
  } catch (e) { return []; }
}

/* ======================================================================
   HELPER: CEK DUPLIKAT, HAPUS, & VERIFIKASI
   ====================================================================== */
function cekDuplikatSK(nomorSk) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_IDS.SK_DATA);
    const sheet = ss.getSheetByName("Unggah_SK");
    var data = sheet.getDataRange().getValues();
    var target = String(nomorSk).toLowerCase().replace(/[^a-z0-9]/g, '');
    
    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      var dbSk = String(row[4] || "").toLowerCase().replace(/[^a-z0-9]/g, ''); 
      
      if (dbSk === target && dbSk !== "") {
        var status = String(row[9] || "").toLowerCase();
        var isLocked = (status.includes("ok") || status.includes("setuju"));
        
        var rawTgl = row[5];
        var fmtTgl = (rawTgl instanceof Date) ? Utilities.formatDate(rawTgl, Session.getScriptTimeZone(), "yyyy-MM-dd") : rawTgl;

        return { 
          found: true, isLocked: isLocked,
          data: {
            rowId: i + 1, namaSd: row[1], tahun: row[2], semester: row[3],
            noSk: row[4], tglSk: fmtTgl, kriteria: row[6], fileUrl: row[7], status: row[9]
          }
        };
      }
    }
    return { found: false };
  } catch (e) { return { found: false }; }
}

function hapusDataSK(form) {
  try {
    // Validasi Kode
    var KODE_RAHASIA = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyyMMdd");
    if (String(form.hapusKode).trim() !== KODE_RAHASIA) {
      return { success: false, message: "Kode Keamanan SALAH!" };
    }

    const ss = SpreadsheetApp.openById(SPREADSHEET_IDS.SK_DATA);
    const sheetSource = ss.getSheetByName("Unggah_SK");
    var sheetTrash = ss.getSheetByName("Trash_SK");
    
    // Buat Sheet Trash jika belum ada
    if (!sheetTrash) {
       sheetTrash = ss.insertSheet("Trash_SK");
       // Ambil Header A-O (15 Kolom)
       var headers = sheetSource.getRange("A1:O1").getValues()[0];
       // Tambah Header P, Q, R
       headers.push("TGL HAPUS", "USER HAPUS", "ALASAN");
       sheetTrash.appendRow(headers); 
    }

    var rowIdx = parseInt(form.hapusRowId);
    if (isNaN(rowIdx)) return { success: false, message: "Row ID Invalid" };

    // AMBIL DATA SUMBER
    var values = sheetSource.getRange(rowIdx, 1, 1, sheetSource.getLastColumn()).getValues()[0]; 
    
    // MOVE FILE (Kolom H / Index 7)
    var fileUrl = values[7];
    if (fileUrl && fileUrl.indexOf("drive.google.com") !== -1) {
        try {
          var fileIdMatch = fileUrl.match(/[-\w]{25,}/);
          if (fileIdMatch) DriveApp.getFileById(fileIdMatch[0]).moveTo(DriveApp.getFolderById(FOLDER_CONFIG.TRASH_SK));
        } catch (e) { rekamCCTV("ERR FILE", e.toString()); }
    }

    // --- PERBAIKAN LOGIKA KOLOM ---
    // Kita hanya ambil data A-O (15 Kolom Pertama)
    // Index 0 s/d 14
    var dataToTrash = values.slice(0, 15); 

    // Masukkan Metadata ke P, Q, R
    dataToTrash[15] = new Date();        // P: Tgl Hapus
    dataToTrash[16] = form.userDelete;   // Q: User Hapus
    dataToTrash[17] = form.hapusAlasan;  // R: Alasan

    // Simpan ke Trash
    sheetTrash.appendRow(dataToTrash);

    // Hapus dari Sumber
    sheetSource.deleteRow(rowIdx);

    return { success: true, message: "Data berhasil dihapus." };

  } catch (e) {
    return { success: false, message: "Gagal: " + e.toString() };
  }
}

function verifikasiDataSK(form) {
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_IDS.SK_DATA);
    var sheet = ss.getSheetByName("Unggah_SK");
    var rowIdx = parseInt(form.verifRowId);

    if (isNaN(rowIdx) || rowIdx < 2) return { success: false, message: "ID Baris tidak valid!" };

    // Update Status (Kolom J = 10)
    sheet.getRange(rowIdx, 10).setValue(form.verifStatus);

    // Update Meta Verval (M=13, N=14, O=15)
    sheet.getRange(rowIdx, 13).setValue(new Date()); 
    sheet.getRange(rowIdx, 14).setValue(form.verifikator); 
    sheet.getRange(rowIdx, 15).setValue(form.verifKeterangan);

    SpreadsheetApp.flush();
    return { success: true, message: "Data diverifikasi: " + form.verifStatus };
  } catch (e) { return { success: false, message: "Error Verifikasi: " + e.toString() }; }
}

/* ======================================================================
   HELPER: REKAM JEJAK CCTV (WAJIB ADA DI BAWAH)
   ====================================================================== */
function rekamCCTV(aktivitas, data) {
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_IDS.SK_DATA); 
    var sheet = ss.getSheetByName("Log_CCTV");
    if (!sheet) {
      sheet = ss.insertSheet("Log_CCTV");
      sheet.appendRow(["TIMESTAMP", "AKTIVITAS", "DATA MENTAH"]);
    }
    var dataString = (typeof data === 'object') ? JSON.stringify(data) : data;
    sheet.appendRow([new Date(), aktivitas, dataString]);
  } catch (e) { Logger.log("CCTV Error"); }
}

/* ======================================================================
   CORE: HAPUS DATA (SOFT DELETE & MOVE FILE)
   ====================================================================== */
function hapusDataSK(form) {
  try {
    // 1. VALIDASI KODE KEAMANAN (SERVER SIDE)
    // Menggunakan Timezone Jakarta/Server agar sinkron
    var KODE_RAHASIA = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyyMMdd");

    if (String(form.hapusKode).trim() !== KODE_RAHASIA) {
      return { success: false, message: "Kode Keamanan SALAH!" };
    }

    const ss = SpreadsheetApp.openById(SPREADSHEET_IDS.SK_DATA);
    const sheetSource = ss.getSheetByName("Unggah_SK");
    
    // Siapkan Sheet Trash
    var sheetTrash = ss.getSheetByName("Trash_SK");
    if (!sheetTrash) {
       sheetTrash = ss.insertSheet("Trash_SK");
       // Copy Header dari sheet utama
       var headers = sheetSource.getRange("A1:Q1").getValues();
       // Tambah Header metadata hapus
       headers[0].push("TGL HAPUS", "USER HAPUS", "ALASAN");
       sheetTrash.appendRow(headers[0]); 
    }

    var rowIdx = parseInt(form.hapusRowId);
    if (isNaN(rowIdx)) return { success: false, message: "Row ID Invalid" };

    // 2. AMBIL DATA YANG AKAN DIHAPUS
    var rangeData = sheetSource.getRange(rowIdx, 1, 1, sheetSource.getLastColumn());
    var values = rangeData.getValues()[0]; 
    
    // Ambil info penting untuk pemindahan file
    var tahun = values[2];   // Kolom C
    var semester = values[3]; // Kolom D
    var fileUrl = values[7];  // Kolom H (Link File)

    // 3. PINDAHKAN FILE FISIK KE TRASH (Agar folder aktif bersih)
    if (fileUrl && fileUrl.indexOf("drive.google.com") !== -1) {
        try {
          var fileIdMatch = fileUrl.match(/[-\w]{25,}/);
          if (fileIdMatch) {
             var file = DriveApp.getFileById(fileIdMatch[0]);
             var trashRoot = DriveApp.getFolderById(FOLDER_CONFIG.TRASH_SK);
             
             // Opsional: Buat Subfolder Tahun di Trash biar rapi
             var folderTahun = getOrCreateFolder(trashRoot, String(tahun).replace(/\//g, '-'));
             var folderSmt = getOrCreateFolder(folderTahun, String(semester));
             
             file.moveTo(folderSmt); // Pindahkan file
          }
        } catch (errFile) {
          rekamCCTV("ERROR HAPUS FILE", errFile.toString());
          // Lanjut saja, jangan batalkan penghapusan data hanya karena file gagal dipindah
        }
    }

    // 4. PINDAHKAN DATA KE SHEET TRASH
    // Clone array values agar tidak merusak referensi
    var trashValues = values.slice(); 
    
    // Tambahkan Metadata Penghapusan
    trashValues.push(new Date());        // Tgl Hapus
    trashValues.push(form.userDelete);   // Siapa yang hapus
    trashValues.push(form.hapusAlasan);  // Alasannya

    sheetTrash.appendRow(trashValues);

    // 5. HAPUS DARI SHEET SUMBER
    sheetSource.deleteRow(rowIdx);

    rekamCCTV("HAPUS DATA", "Menghapus Baris " + rowIdx + " oleh " + form.userDelete);
    return { success: true, message: "Data berhasil dihapus." };

  } catch (e) {
    rekamCCTV("ERROR HAPUS", e.toString());
    return { success: false, message: "Gagal menghapus: " + e.toString() };
  }
}

/* ======================================================================
   CORE: GET DATA SAMPAH (TRASH)
   ====================================================================== */
function getTrashSK() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_IDS.SK_DATA);
  const sheet = ss.getSheetByName("Trash_SK");
  if (!sheet) return [];

  var data = sheet.getDataRange().getValues();
  var result = [];
  
  var fmt = function(d) {
    try { return (d instanceof Date) ? Utilities.formatDate(d, Session.getScriptTimeZone(), "dd-MM-yyyy HH:mm") : String(d); } catch(e){ return ""; }
  };

  // Loop baris data
  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    if (!row[1]) continue; 

    result.push({
      rowBaris: i + 1,
      namaSd: row[1], // Kolom B
      noSk: row[4],   // Kolom E
      
      // BACA KOLOM P, Q, R (Index 15, 16, 17)
      tglHapus: fmt(row[15]),  // Kolom P
      userHapus: row[16],      // Kolom Q
      alasanHapus: row[17]     // Kolom R
    });
  }
  return result;
}

/* ======================================================================
   CORE: RESTORE DATA (PULIHKAN DARI TRASH)
   ====================================================================== */
function restoreDataSK(form) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_IDS.SK_DATA);
    const sheetTrash = ss.getSheetByName("Trash_SK");
    const sheetActive = ss.getSheetByName("Unggah_SK");
    
    var rowIdx = parseInt(form.rowId);
    var values = sheetTrash.getRange(rowIdx, 1, 1, sheetTrash.getLastColumn()).getValues()[0];
    
    // POTONG METADATA (Hanya ambil A-O / 15 Kolom pertama)
    // Karena P, Q, R adalah sampah, jangan dikembalikan ke tabel utama
    var cleanValues = values.slice(0, 15);
    
    // RESTORE FILE FISIK (Opsional)
    var fileUrl = cleanValues[7];
    if (fileUrl && fileUrl.indexOf("drive.google.com") !== -1) {
        try {
          var fileIdMatch = fileUrl.match(/[-\w]{25,}/);
          if (fileIdMatch) DriveApp.getFileById(fileIdMatch[0]).moveTo(DriveApp.getFolderById(FOLDER_CONFIG.MAIN_SK));
        } catch (e) {}
    }
    
    // Catat Siapa yang Restore di Kolom O (Keterangan) - Opsional
    // cleanValues[14] = "Dipulihkan oleh " + form.userRestore; 

    sheetActive.appendRow(cleanValues);
    sheetTrash.deleteRow(rowIdx);
    
    return { success: true, message: "Data berhasil dipulihkan." };
    
  } catch (e) {
    return { success: false, message: "Gagal Restore: " + e.toString() };
  }
}

/* ======================================================================
   MODULE: STATUS PENGIRIMAN SK (SIMPLE HEADER)
   ====================================================================== */
function getSiabaStatusData() {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_IDS.SK_DATA);
    const sheet = ss.getSheetByName("Status_SK");
    
    // Ambil Data Mentah
    var rawData = sheet.getDataRange().getDisplayValues();
    
    if (rawData.length < 2) return { error: "Data Status SK belum tersedia." };

    // Baris 1: Header
    var headers = rawData[0]; 
    
    // Baris 2 dst: Data
    var rows = rawData.slice(1); 

    // Ambil Sekolah untuk Filter (Kolom A)
    var listSekolah = [];
    rows.forEach(r => {
      if(r[0] && r[0] !== "" && r[0] !== "NAMA SEKOLAH" && !r[0].includes("Sem ")) {
         listSekolah.push(r[0]);
      }
    });
    listSekolah = [...new Set(listSekolah)].sort();

    return {
       headers: headers,
       rows: rows,
       schools: listSekolah
    };

  } catch (e) {
    return { error: "Gagal ambil data: " + e.toString() };
  }
}

/* ======================================================================
   MODULE: DASHBOARD SK (DENGAN LOGIC BELUM MENGIRIM)
   ====================================================================== */
function getDashboardSK(filterTahun, filterSemester) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_IDS.SK_DATA);
    
    // 1. AMBIL DATA SUDAH MASUK (Unggah_SK)
    const sheetData = ss.getSheetByName("Unggah_SK");
    var rawData = sheetData.getDataRange().getValues();
    var rows = rawData.slice(1); // Skip Header

    // 2. AMBIL DATA MASTER SEKOLAH (Wajib ada sheet 'Master_Sekolah')
    var masterSekolah = [];
    var sheetMaster = ss.getSheetByName("Master_Sekolah");
    if (sheetMaster) {
        var rawMaster = sheetMaster.getDataRange().getValues();
        // Asumsi Nama Sekolah ada di Kolom A
        rawMaster.forEach(r => { if(r[0]) masterSekolah.push(String(r[0]).trim()); });
    }

    // Init Stats
    var stats = {
      totalMasuk: 0,
      diproses: 0,
      revisi: 0,
      disetujui: 0,
      ditolak: 0,
      progress: 0,
      belumLaporCount: 0,
      belumLaporList: [], // Array nama sekolah
      recent: []
    };

    // Set Sekolah yang sudah lapor (Untuk Comparison)
    var sekolahSudahLapor = new Set();

    // 3. FILTER & HITUNG
    var filteredRows = rows.filter(function(r) {
      if (!r[1]) return false;
      
      var matchTahun = (filterTahun === "" || String(r[2]) === String(filterTahun));
      var matchSmt = (filterSemester === "" || String(r[3]) === String(filterSemester));
      
      if (matchTahun && matchSmt) {
          sekolahSudahLapor.add(String(r[1]).trim()); // Catat sekolah yg sudah lapor
          return true;
      }
      return false;
    });

    stats.totalMasuk = filteredRows.length;

    // Hitung Detail Status
    filteredRows.forEach(function(r) {
      var s = String(r[8] || "").toLowerCase(); // Kolom I/Status (Index 8 di array 0-based data slice?? Cek mapping)
      // Cek mapping: A=0, B=1, C=2... I=8, J=9 (Status) di Unggah_SK biasanya Kolom J (Index 9)
      // Mari kita pakai index 9 sesuai kode sebelumnya (Kolom J)
      s = String(r[9] || "").toLowerCase();

      if (s.includes("ok") || s.includes("setuju") || s.includes("valid")) {
        stats.disetujui++;
      } else if (s.includes("revisi")) {
        stats.revisi++;
      } else if (s.includes("tolak")) {
        stats.ditolak++;
      } else {
        stats.diproses++;
      }
    });

    // 4. HITUNG YANG BELUM LAPOR
    if (masterSekolah.length > 0) {
        // Filter Master yang TIDAK ADA di Set sekolahSudahLapor
        stats.belumLaporList = masterSekolah.filter(x => !sekolahSudahLapor.has(x)).sort();
        stats.belumLaporCount = stats.belumLaporList.length;
        
        // Hitung Progress Real (Disetujui / Total Master)
        stats.progress = Math.round((stats.disetujui / masterSekolah.length) * 100);
    } else {
        // Fallback jika Master belum dibuat
        stats.belumLaporCount = 0;
        stats.belumLaporList = ["Sheet 'Master_Sekolah' belum dibuat di Database."];
    }

    // 5. RECENT ACTIVITY
    // Sort by timestamp desc (Kolom A / Index 0)
    var sorted = filteredRows.sort(function(a, b) {
      return new Date(b[0]) - new Date(a[0]);
    }).slice(0, 5);

    stats.recent = sorted.map(function(r) {
      return {
        sekolah: r[1],
        status: r[9], // Kolom J
        waktu: (r[0] instanceof Date) ? Utilities.formatDate(r[0], Session.getScriptTimeZone(), "dd/MM HH:mm") : r[0]
      };
    });

    return stats;

  } catch (e) {
    return { error: e.toString() };
  }
}