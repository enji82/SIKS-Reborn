const SPREADSHEET_IDS = {
  DATABASE_USER: "1wiDKez4rL5UYnpP2-OZjYowvmt1nRx-fIMy9trJlhBA",
  SHEET_USER_NAME: "Data User",
  SK_DATA: "1AmvOJAhOfdx09eT54x62flWzBZ1xNQ8Sy5lzvT9zJA4",
  DROPDOWN_DATA: "1wiDKez4rL5UYnpP2-OZjYowvmt1nRx-fIMy9trJlhBA", 
  PAUD_DATA: "1an0oQQPdMh6wrUJIAzTGYk3DKFvYprK5SU7RmRXjIgs",
  SD_DATA: "1u4tNL3uqt5xHITXYwHnytK6Kul9Siam-vNYuzmdZB4s",
  LAPBUL_GABUNGAN: "1aKEIkhKApmONrCg-QQbMhXyeGDJBjCZrhR-fvXZFtJU",
  PTK_PAUD_DB: "1XetGkBymmN2NZQlXpzZ2MQyG0nhhZ0sXEPcNsLffhEU",
  PTK_SD_DB: "1HlyLv3Ai3_vKFJu3EKznqI9v8g0tfqiNg0UbIojNMQ0",
  DATA_SEKOLAH: "1qeOYVfqFQdoTpysy55UIdKwAJv3VHo4df3g6u6m72Bs",   
  FORM_OPTIONS_DB: "1prqqKQBYzkCNFmuzblNAZE41ag9rZTCiY2a0WvZCTvU",
  SIABA_DB: "1sfbvyIZurU04gictep8hI-NnvicGs0wrDqANssVXt6o",
  SIABA_TA_PA: "1tQsQY1-Ny1ie66GOZPTLtvZ7BiYCgFdNrX-AVGCtaHA",
  SIABA_SALAH_DB: "1TZGrMiTuyvh2Xbo44RhJuWlQnOC5LzClsgIoNKtRFkY",
  SIABA_DINAS_DB: "1I_2yUFGXnBJTCSW6oaT3D482YCs8TIRkKgQVBbvpa1M",
  SIABA_CUTI_DB: "1DhBjmLHFMuJqWM6yJHsm-1EKvHzG8U4zK2GuU-dIgn8",
  SIABA_REKAP_HELPER: "1wiDKez4rL5UYnpP2-OZjYowvmt1nRx-fIMy9trJlhBA",
  SIABA_SKP_SOURCE: "1ReJt2qoDE2f_8LeR8DXJbROB9EAHK8qP2kYp-ZZ3V9w", 
  SIABA_SKP_DB: "1T-AQ0jYJ_jXYEPxzu_KZauOlRTTforVtFEZ_1UrWHwk",
  SIABA_PNS_DB: "1wiDKez4rL5UYnpP2-OZjYowvmt1nRx-fIMy9trJlhBA",
  SIABA_PAK_DB: "1mAXwf7cHaOqIj2uf51Fup5tyyBzijTeIxVS8uO1E4dM",
};

const FOLDER_CONFIG = {
  MAIN_SK: "1GwIow8B4O1OWoq3nhpzDbMO53LXJJUKs", 
  TRASH_SK: "1OB2Mxa_zvpYl7Vru9NEddYmBlU5SfYHL",
  LAPBUL_KB: "18CxRT-eledBGRtHW1lFd2AZ8Bub6q5ra",
  LAPBUL_TK: "1WUNz_BSFmcwRVlrG67D2afm9oJ-bVI9H",
  LAPBUL_SD: "1I8DRQYpBbTt1mJwtD1WXVD6UK51TC8El",
  SIABA_LUPA: "10kwGuGfwO5uFreEt7zBJZUaDx1fUSXo9",
  SIABA_DINAS: "1uPeOU7F_mgjZVyOLSsj-3LXGdq9rmmWl",
  SIABA_CUTI_DOCS: "1fAmqJXpmGIfEHoUeVm4LjnWvnwVwOfNM",
  SIABA_REKAP_ARCHIVE: "1MoGuseJNrOIMnkZNoqkKcK282jZpUkAm",
  SIABA_SKP_DOCS: "1DGYC8AtJFCpCZ0ou2ae9-5fc2-bWl20G",
  SIABA_PAK_DOCS: "1cvn-pOufs-OIbFQfqhmxc3fcmFuox4Sc",
};

function doGet(e) {
  // Gunakan 'createTemplateFromFile' agar bisa membaca <?!= include ?>
  var template = HtmlService.createTemplateFromFile('index');
  
  // Wajib ada .evaluate() untuk merender template
  return template.evaluate()
      .setTitle('SIKS - REBORN')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

// PERBAIKAN VITAL: Menggunakan Template agar kode <?!= di dalam file diproses server
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

function checkLogin(username, password) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_IDS.DATABASE_USER);
  const sheet = ss.getSheetByName(SPREADSHEET_IDS.SHEET_USER_NAME);
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim() == username && String(data[i][1]).trim() == password) {
      const userObj = {
        fullName: data[i][2], role: data[i][3], photo: data[i][4] || "", isLoggedIn: true
      };
      PropertiesService.getUserProperties().setProperty('currentUser', JSON.stringify(userObj));
      return userObj;
    }
  }
  return null;
}

function getCurrentUser() {
  const user = PropertiesService.getUserProperties().getProperty('currentUser');
  return user ? JSON.parse(user) : null;
}

function processLogin(formObject) {
  // 1. Ambil ID dari konstanta yang sudah Anda buat di atas
  var ss = SpreadsheetApp.openById(SPREADSHEET_IDS.DATABASE_USER); 
  
  // 2. Ambil Nama Sheet dari konstanta (tadi tertulis "Users", harusnya "Data User")
  var sheetName = SPREADSHEET_IDS.SHEET_USER_NAME; 
  var sheet = ss.getSheetByName(sheetName);
  
  if (!sheet) {
    return { status: "error", message: "Sheet '" + sheetName + "' tidak ditemukan!" };
  }

  var data = sheet.getDataRange().getValues();
  
  // Ambil input dari user (Pastikan name="username" di HTML ada)
  var inputUser = formObject.username ? formObject.username.toString().trim() : "";
  var inputPass = formObject.password ? formObject.password.toString().trim() : "";

  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    
    // Sesuaikan kolom A=0 (User), B=1 (Pass)
    var dbUser = row[0] ? row[0].toString().trim() : "";
    var dbPass = row[1] ? row[1].toString().trim() : "";
    
    if (dbUser === inputUser && dbPass === inputPass) {
      // Simpan juga ke PropertiesService agar fungsi checkSession() lama tetap jalan
      var userObj = {
        fullName: row[2], 
        role: row[3], 
        photo: row[4] || "", 
        isLoggedIn: true
      };
      PropertiesService.getUserProperties().setProperty('currentUser', JSON.stringify(userObj));

      return {
        status: "success",
        username: dbUser,
        nama: row[2], // Nama Lengkap
        role: row[3], 
        foto: row[4]  // ID Foto Drive
      };
    }
  }

  return { status: "error", message: "Username atau Password salah" };
}

function processLogout() {
  PropertiesService.getUserProperties().deleteProperty('currentUser');
}

function getHalaman(namaFile) {
  try {
    const prefix = "page_";
    const realName = namaFile.startsWith(prefix) ? namaFile : prefix + namaFile;
    return HtmlService.createTemplateFromFile(realName).evaluate().getContent();
  } catch (err) {
    return '<div class="p-4"><div class="alert alert-warning">File <b>' + namaFile + '</b> belum ada.</div></div>';
  }
}

function getScriptUrl() {
  return ScriptApp.getService().getUrl();
}

function prosesUnggahSK(formData) {
  try {
    // 1. TANGKAP USERNAME DARI BROWSER (Ini kuncinya!)
    // Jika formData.username kosong, fallback ke email login
    const usernameKirim = formData.username || Session.getActiveUser().getEmail();
    const usernameCari = usernameKirim.toString().toLowerCase().trim();

    // 2. Buka Database User untuk Mencari Nama Lengkap
    const ssUser = SpreadsheetApp.openById(SPREADSHEET_IDS.DATABASE_USER);
    const sheetUser = ssUser.getSheetByName(SPREADSHEET_IDS.SHEET_USER_NAME);
    const dataUser = sheetUser.getDataRange().getValues();
    
    let namaLengkapFinal = usernameKirim; // Default nama user jika tidak ketemu

    // 3. Loop Cari Username di Kolom A
    for (let i = 1; i < dataUser.length; i++) {
      // Pastikan kolom A ada isinya
      if (dataUser[i][0]) {
        let dbUsername = dataUser[i][0].toString().toLowerCase().trim(); 
        
        // Jika Username Cocok
        if (dbUsername === usernameCari) {
          namaLengkapFinal = dataUser[i][2]; // AMBIL NAMA LENGKAP (KOLOM C)
          break;
        }
      }
    }

    // 4. Proses Simpan File ke Drive
    const ssSK = SpreadsheetApp.openById(SPREADSHEET_IDS.SK_DATA);
    const sheetSK = ssSK.getSheetByName("Unggah_SK");
    const parentFolder = DriveApp.getFolderById(FOLDER_CONFIG.MAIN_SK);
    
    // Buat Folder Tahun & Semester jika belum ada
    const folderTahun = getSubFolder_(parentFolder, formData.tahunAjaran.replace(/\//g, "-"));
    const folderSemester = getSubFolder_(folderTahun, formData.semester);
    
    // Buat File PDF
    const fileBlob = Utilities.newBlob(
      Utilities.base64Decode(formData.fileData), 
      "application/pdf", 
      formData.namaSd + " - " + formData.kriteriaSk
    );
    const newFile = folderSemester.createFile(fileBlob);
    const fileUrl = newFile.getUrl();

    // 5. Masukkan Data ke Spreadsheet
    // Urutan Kolom: A=Tanggal, B=NamaSD, C=Tahun, D=Semester, E=NoSK, F=TglSK, G=Kriteria, H=Link, I=UserInput
    const rowData = [
      new Date(),             
      formData.namaSd,        
      formData.tahunAjaran,   
      formData.semester,      
      formData.nomorSk,       
      formData.tanggalSk,     
      formData.kriteriaSk,    
      fileUrl,                
      namaLengkapFinal,       // <--- NAMA LENGKAP HASIL PENCARIAN
      "Diproses",             
      "", "", "", "", ""      
    ];

    sheetSK.appendRow(rowData);
    return { success: true, message: "Dokumen berhasil disimpan atas nama: " + namaLengkapFinal };
    
  } catch (e) {
    return { success: false, message: "Error Server: " + e.toString() };
  }
}

/**
 * FUNGSI BANTUAN: Handle Folder Drive
 */
function getSubFolder_(parent, folderName) {
  const folders = parent.getFoldersByName(folderName);
  if (folders.hasNext()) {
    return folders.next();
  } else {
    return parent.createFolder(folderName);
  }
}

/* ======================================================================
   AMBIL DAFTAR SK (FIX KOLOM USER INPUT & TGL UNGGAH)
   ====================================================================== */
function getDaftarSK() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_IDS.SK_DATA);
  const sheet = ss.getSheetByName("Unggah_SK");
  var data = sheet.getDataRange().getValues();
  
  var result = [];
  
  // Helper Format
  var fmtFull = function(d) {
    try {
      if (d instanceof Date) return Utilities.formatDate(d, Session.getScriptTimeZone(), "dd-MM-yyyy HH:mm:ss");
      return String(d); 
    } catch(e) { return ""; }
  };

  var fmtDate = function(d) {
    try {
      if (d instanceof Date) return Utilities.formatDate(d, Session.getScriptTimeZone(), "dd-MM-yyyy");
      return String(d);
    } catch(e) { return ""; }
  };

  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    
    // Skip jika Nama SD kosong
    if (!row[1] || row[1] === "") continue; 

    result.push({
      rowBaris: i + 1,        
      namaSd: String(row[1]),         // B
      tahun: String(row[2]),          // C
      semester: String(row[3]),       // D
      noSk: String(row[4]),           // E
      tglSk: fmtDate(row[5]),         // F
      kriteria: String(row[6]),       // G
      fileUrl: String(row[7]),        // H
      
      // --- PERBAIKAN MAPPING DISINI ---
      
      // row[8] (Kolom I) ternyata adalah User Input
      userInput: String(row[8]),      
      
      status: String(row[9]),         // J (Status)
      tglUpdate: fmtFull(row[10]),    // K
      userUpdate: String(row[11]),    // L
      tglVerval: fmtFull(row[12]),    // M
      verifikator: String(row[13]),   // N
      keterangan: String(row[14]),    // O
      
      // row[15] (Kolom P) kita coba ambil sebagai Tanggal Unggah.
      // Jika kosong, alternatifnya bisa ambil row[0] (Kolom A) jika itu Timestamp.
      tglUnggah: fmtFull(row[15] || row[0]) 
      
      // --------------------------------
    });
  }
  
  return result;
}

// FUNGSI BANTUAN FORMAT TANGGAL (PENTING AGAR DATA MUNCUL)
function formatDate_(dateObj) {
  if (!dateObj || dateObj === "") return "-";
  try {
    // Ubah objek tanggal jadi teks "05-01-2026"
    return Utilities.formatDate(new Date(dateObj), "Asia/Jakarta", "dd-MM-yyyy");
  } catch (e) {
    return String(dateObj); // Kalau gagal, kembalikan aslinya
  }
}

function hapusDataSK(rowBaris) {
  /* ... (Fungsi hapus tetap sama seperti sebelumnya) ... */
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_IDS.SK_DATA);
    const sheet = ss.getSheetByName("Unggah_SK");
    sheet.deleteRow(parseInt(rowBaris));
    return { success: true, message: "Data berhasil dihapus!" };
  } catch (e) {
    return { success: false, message: "Gagal: " + e.toString() };
  }
}

/* ======================================================================
   FUNGSI UPDATE DATA SK (REVISI LOGIKA STATUS LEBIH KUAT)
   ====================================================================== */
function updateDataSK(form) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_IDS.SK_DATA);
    const sheet = ss.getSheetByName("Unggah_SK");
    var rowIdx = parseInt(form.editRowId);
    
    if (rowIdx < 2) return { success: false, message: "ID Baris tidak valid!" };

    // 1. UPDATE DATA TEKS UTAMA
    sheet.getRange(rowIdx, 5).setValue(form.editNoSk);     
    sheet.getRange(rowIdx, 6).setValue(form.editTglSk);    
    sheet.getRange(rowIdx, 7).setValue(form.editKriteria); 

    // 2. LOGIKA STATUS OTOMATIS (VERSI KEBAL)
    // Ambil status saat ini (Kolom J / 10)
    var rangeStatus = sheet.getRange(rowIdx, 10);
    var rawStatus = rangeStatus.getValue();
    
    // Bersihkan data: Ubah ke String, Trim spasi, Kecilkan huruf
    var statusLama = String(rawStatus).trim().toLowerCase();

    // Cek: jika "ditolak" (huruf besar/kecil/spasi tidak masalah)
    if (statusLama === 'ditolak') {
        rangeStatus.setValue('Revisi'); // Paksa ubah jadi Revisi
    }
    // Jika Diproses atau Revisi, biarkan saja.

    // 3. CEK APAKAH ADA FILE BARU?
    if (form.fileContent) {
       var cleanSD = form.editNamaSd.replace(/[^a-zA-Z0-9 ]/g, "").replace(/\s+/g, "_");
       var cleanTh = form.editTahun.replace("/", "-"); 
       var namaFileBaru = "SK_" + cleanSD + "_" + cleanTh + "_Smt" + form.editSemester + "_" + form.editKriteria + ".pdf";
       
       var data = Utilities.base64Decode(form.fileContent.split(',')[1]);
       var blob = Utilities.newBlob(data, form.mimeType, namaFileBaru);
       
       var folderTujuan;
       try { folderTujuan = DriveApp.getFolderById(FOLDER_CONFIG.MAIN_SK); } 
       catch(e) { folderTujuan = DriveApp.getRootFolder(); }
       
       var newFile = folderTujuan.createFile(blob); 
       newFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
       
       try {
         var urlLama = sheet.getRange(rowIdx, 8).getValue();
         if(urlLama && urlLama.includes("drive.google.com")) {
            var idLama = urlLama.match(/[-\w]{25,}/);
            if(idLama) DriveApp.getFileById(idLama[0]).setTrashed(true);
         }
       } catch(e) { }

       sheet.getRange(rowIdx, 8).setValue(newFile.getUrl());
    }

    // 4. CATAT LOG
    sheet.getRange(rowIdx, 11).setValue(new Date()); 
    
    var currentUser = "Admin"; 
    try {
       var props = PropertiesService.getUserProperties().getProperty('currentUser');
       if(props) {
         var userObj = JSON.parse(props);
         currentUser = userObj.nama || userObj.fullName || userObj.username || "User";
       }
    } catch(e){}
    
    sheet.getRange(rowIdx, 12).setValue(currentUser);

    return { success: true, message: "Data berhasil diperbarui!" };
    
  } catch (e) {
    return { success: false, message: "Error Update: " + e.toString() };
  }
}

/* ======================================================================
   FUNGSI VERIFIKASI DATA SK (REVISI KOLOM)
   ====================================================================== */
function verifikasiDataSK(form) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_IDS.SK_DATA);
    const sheet = ss.getSheetByName("Unggah_SK");
    var rowIdx = parseInt(form.verifRowId);
    
    if (rowIdx < 2) return { success: false, message: "ID Baris tidak valid!" };

    // 1. UPDATE STATUS -> Kolom J (Indeks 10)
    sheet.getRange(rowIdx, 10).setValue(form.verifStatus);
    
    // 2. UPDATE TANGGAL VERVAL -> Kolom M (Indeks 13)
    sheet.getRange(rowIdx, 13).setValue(new Date());

    // 3. UPDATE NAMA VERIFIKATOR -> Kolom N (Indeks 14)
    var currentUser = "Admin"; 
    try {
       var props = PropertiesService.getUserProperties().getProperty('currentUser');
       if(props) {
         var userObj = JSON.parse(props);
         currentUser = userObj.nama || userObj.fullName || userObj.username || "Verifikator";
       }
    } catch(e){}
    
    sheet.getRange(rowIdx, 14).setValue(currentUser);

    // 4. UPDATE KETERANGAN -> Kolom O (Indeks 15)
    sheet.getRange(rowIdx, 15).setValue(form.verifKeterangan);

    return { success: true, message: "Data berhasil diverifikasi!" };
    
  } catch (e) {
    return { success: false, message: "Error Verval: " + e.toString() };
  }
}

/* ======================================================================
   FUNGSI HAPUS DATA SK (SOFT DELETE & ARSIP FILE)
   ====================================================================== */
function hapusDataSK(form) {
  try {
    // 1. GENERATE KODE RAHASIA (YYYYMMDD)
    // Menggunakan Timezone Script (WIB) agar sinkron dengan hari user
    var KODE_RAHASIA = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyyMMdd");

    // 2. VALIDASI KODE
    // Pastikan kode yang dikirim user sama persis dengan tanggal hari ini
    if (String(form.hapusKode).trim() !== KODE_RAHASIA) {
      return { success: false, message: "Kode Konfirmasi SALAH! Gunakan format tanggal hari ini (YYYYMMDD)." };
    }

    const ss = SpreadsheetApp.openById(SPREADSHEET_IDS.SK_DATA);
    const sheetSource = ss.getSheetByName("Unggah_SK");
    const sheetTrash = ss.getSheetByName("Trash_SK");
    
    if (!sheetTrash) ss.insertSheet("Trash_SK");

    var rowIdx = parseInt(form.hapusRowId);
    
    // Ambil Data Baris
    var rangeData = sheetSource.getRange(rowIdx, 1, 1, sheetSource.getLastColumn());
    var values = rangeData.getValues()[0]; 

    var namaSD = values[1]; 
    var tahun = values[2];  
    var semester = values[3]; 
    var fileUrl = values[7]; 

    // 3. PINDAHKAN FILE KE TRASH
    if (fileUrl && fileUrl.includes("drive.google.com")) {
       try {
         var fileIdMatch = fileUrl.match(/[-\w]{25,}/);
         if (fileIdMatch) {
            var file = DriveApp.getFileById(fileIdMatch[0]);
            var trashRoot = DriveApp.getFolderById(FOLDER_CONFIG.TRASH_SK);
            
            // Subfolder Tahun
            var folderTahun;
            var iterTahun = trashRoot.getFoldersByName(tahun);
            folderTahun = iterTahun.hasNext() ? iterTahun.next() : trashRoot.createFolder(tahun);
            
            // Subfolder Semester
            var folderSmt;
            var namaSmt = "Semester " + semester;
            var iterSmt = folderTahun.getFoldersByName(namaSmt);
            folderSmt = iterSmt.hasNext() ? iterSmt.next() : folderTahun.createFolder(namaSmt);
            
            file.moveTo(folderSmt);
         }
       } catch (errFile) {
         console.warn("Gagal pindah file: " + errFile.toString());
       }
    }

    // 4. PINDAH DATA KE SHEET TRASH
    var dataTrash = values.slice(0, 15); 
    
    dataTrash[15] = new Date(); // Tgl Hapus
    
    var currentUser = "Admin"; 
    try {
       var props = PropertiesService.getUserProperties().getProperty('currentUser');
       if(props) currentUser = JSON.parse(props).fullName || "User";
    } catch(e){}
    dataTrash[16] = currentUser; // User Hapus

    dataTrash[17] = form.hapusAlasan; // Alasan

    sheetTrash.appendRow(dataTrash);
    sheetSource.deleteRow(rowIdx);

    return { success: true, message: "Data berhasil dipindahkan ke Trash." };

  } catch (e) {
    return { success: false, message: "Error Hapus: " + e.toString() };
  }
}

/* ======================================================================
   AMBIL DATA STATUS SK (KOLOM DINAMIS)
   ====================================================================== */
function getStatusPengiriman() {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_IDS.SK_DATA);
    const sheet = ss.getSheetByName("Status_SK");
    
    // Ambil semua data (Header + Isi) sebagai String (DisplayValues)
    // Agar format di Sheet terjaga
    var rawData = sheet.getDataRange().getDisplayValues();
    
    if (rawData.length === 0) return { headers: [], rows: [] };

    // Pisahkan Header (Baris 0) dan Data (Baris 1 s/d Akhir)
    var headers = rawData.shift(); // Ambil elemen pertama sebagai header
    var rows = rawData;            // Sisanya adalah data
    
    return { 
      headers: headers, 
      rows: rows 
    };

  } catch (e) {
    throw new Error("Gagal mengambil status: " + e.message);
  }
}

function getArsipData(folderId) {
  try {
    // Jika folderId null/kosong, gunakan MAIN_SK (Root)
    var targetId = folderId || FOLDER_CONFIG.MAIN_SK;
    var folder = DriveApp.getFolderById(targetId);
    
    // Cek apakah ini Root Folder (untuk sembunyikan tombol Back)
    var isRoot = (targetId === FOLDER_CONFIG.MAIN_SK);
    
    // Ambil Parent ID (untuk tombol Back)
    // Jika isRoot, parent null. Jika tidak, ambil parentnya.
    var parents = folder.getParents();
    var parentId = parents.hasNext() ? parents.next().getId() : null;

    var items = [];

    // 1. AMBIL SUB-FOLDER
    var subFolders = folder.getFolders();
    while (subFolders.hasNext()) {
      var f = subFolders.next();
      items.push({
        id: f.getId(),
        name: f.getName(),
        type: 'folder',
        mimeType: 'application/vnd.google-apps.folder',
        size: '-',
        date: Utilities.formatDate(f.getLastUpdated(), Session.getScriptTimeZone(), "dd-MM-yyyy HH:mm"),
        url: f.getUrl()
      });
    }

    // 2. AMBIL FILES
    var files = folder.getFiles();
    while (files.hasNext()) {
      var file = files.next();
      items.push({
        id: file.getId(),
        name: file.getName(),
        type: 'file',
        mimeType: file.getMimeType(),
        size: formatBytes(file.getSize()),
        date: Utilities.formatDate(file.getLastUpdated(), Session.getScriptTimeZone(), "dd-MM-yyyy HH:mm"),
        url: file.getUrl()
      });
    }

    // Sort: Folder dulu, baru File. Lalu urut abjad.
    items.sort(function(a, b) {
      if (a.type === b.type) {
        return a.name.localeCompare(b.name);
      }
      return a.type === 'folder' ? -1 : 1;
    });

    return {
      currentId: targetId,
      currentName: folder.getName(),
      parentId: parentId,
      isRoot: isRoot,
      items: items
    };

  } catch (e) {
    throw new Error("Gagal akses Drive: " + e.message);
  }
}

// Helper: Format Ukuran File (KB, MB, GB)
function formatBytes(bytes, decimals = 2) {
    if (bytes === 0) return '0 B';
    const k = 1024;
    const dm = decimals < 0 ? 0 : decimals;
    const sizes = ['B', 'KB', 'MB', 'GB', 'TB'];
    const i = Math.floor(Math.log(bytes) / Math.log(k));
    return parseFloat((bytes / Math.pow(k, i)).toFixed(dm)) + ' ' + sizes[i];
}

/* ======================================================================
   AMBIL DATA TRASH SK (SAMPAH)
   ====================================================================== */
function getTrashSK() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_IDS.SK_DATA);
  const sheet = ss.getSheetByName("Trash_SK");
  
  // Jika sheet belum ada (belum pernah ada yang dihapus), return kosong
  if (!sheet) return [];

  var data = sheet.getDataRange().getValues();
  var result = [];
  
  // Formatter
  var fmtFull = function(d) {
    try { return (d instanceof Date) ? Utilities.formatDate(d, Session.getScriptTimeZone(), "dd-MM-yyyy HH:mm:ss") : String(d); } catch(e){ return ""; }
  };
  var fmtDate = function(d) {
    try { return (d instanceof Date) ? Utilities.formatDate(d, Session.getScriptTimeZone(), "dd-MM-yyyy") : String(d); } catch(e){ return ""; }
  };

  // Loop mulai baris ke-2
  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    
    // Pastikan Nama SD tidak kosong
    if (!row[1]) continue; 

    result.push({
      // ID Baris di Trash (Penting jika nanti mau fitur Restore/Permanent Delete)
      rowBaris: i + 1,        
      
      // Data Asli
      namaSd: String(row[1]),         
      tahun: String(row[2]),          
      semester: String(row[3]),       
      noSk: String(row[4]),           
      tglSk: fmtDate(row[5]),
      kriteria: String(row[6]),
      fileUrl: String(row[7]),

      // Data Penghapusan (Kolom P, Q, R -> Index 15, 16, 17)
      tglHapus: fmtFull(row[15]),
      userHapus: String(row[16]),
      alasanHapus: String(row[17])
    });
  }
  
  return result;
}

/* ======================================================================
   RESTORE DATA SK (PULIHKAN DARI SAMPAH KE UNGGAH_SK)
   ====================================================================== */
function restoreDataSK(data) {
  try {
    const rowId = data.rowId;
    const userRestore = data.userRestore; // Nama admin yang memulihkan

    const ss = SpreadsheetApp.openById(SPREADSHEET_IDS.SK_DATA);
    const sheetTrash = ss.getSheetByName("Trash_SK");
    const sheetData = ss.getSheetByName("Unggah_SK"); // PERBAIKAN: Target Sheet
    
    // Config Folder (Pastikan ID ini ada di Global Variable)
    const targetFolderId = FOLDER_CONFIG.MAIN_SK; 

    // Validasi Baris
    var lastRow = sheetTrash.getLastRow();
    if (rowId > lastRow) {
      return { success: false, message: "Data tidak ditemukan (mungkin sudah berubah posisi)." };
    }

    // 1. AMBIL DATA DARI TRASH
    // Ambil semua kolom di baris tersebut
    var range = sheetTrash.getRange(rowId, 1, 1, sheetTrash.getLastColumn());
    var values = range.getValues()[0];

    // Struktur Trash (Asumsi):
    // Index 0-15: Data Asli (16 Kolom)
    // Index 16-18: Info Hapus (Tgl, User, Alasan) -> KITA BUANG
    
    // Potong data asli saja (16 kolom pertama)
    // Sesuaikan angka 16 dengan jumlah kolom di sheet Unggah_SK Boss
    var restoredRow = values.slice(0, 16); 
    
    // 2. UPDATE INFO (OPSIONAL)
    // Kita ubah kolom 'Keterangan' (Index 15 / Kolom P)
    // Agar ketahuan ini data pulihan
    restoredRow[15] = "Dipulihkan oleh " + userRestore + " pada " + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd-MM-yyyy");

    // 3. PINDAHKAN FILE FISIK DI DRIVE
    var fileUrl = restoredRow[6]; // Asumsi URL ada di Index 6 (Kolom G) - Sesuaikan jika beda
    // Jika URL kosong, lewati langkah ini
    if (fileUrl && fileUrl.indexOf("drive.google.com") !== -1) {
       try {
         var fileId = fileUrl.match(/[-\w]{25,}/);
         if (fileId) {
             var file = DriveApp.getFileById(fileId[0]);
             var targetFolder = DriveApp.getFolderById(targetFolderId);
             file.moveTo(targetFolder); // Pindahkan fisik file
         }
       } catch (e) {
         // Jika gagal pindah file (misal file sudah dihapus permanen), data tetap kita restore tapi beri info
         restoredRow[15] += " (File Gagal Dipulihkan)";
       }
    }

    // 4. MASUKKAN KE SHEET UNGGAH_SK
    sheetData.appendRow(restoredRow);
    
    // 5. HAPUS DARI SHEET TRASH
    sheetTrash.deleteRow(rowId);
    
    return { success: true, message: "Data & File berhasil dipulihkan ke Arsip Aktif." };

  } catch (e) {
    return { success: false, message: "Gagal Restore: " + e.message };
  }
}