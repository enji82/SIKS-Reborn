/* ======================================================================
   CODE.GS - KONFIGURASI GLOBAL & SISTEM UTAMA
   Berisi: ID Database, ID Folder, Login, & Routing Halaman
   ====================================================================== */

// 1. DATABASE CONFIG (Digunakan oleh semua file .gs lainnya)
const SPREADSHEET_IDS = {
  DATABASE_USER: "1wiDKez4rL5UYnpP2-OZjYowvmt1nRx-fIMy9trJlhBA",
  SHEET_USER_NAME: "Data User",
  SK_DATA: "1AmvOJAhOfdx09eT54x62flWzBZ1xNQ8Sy5lzvT9zJA4", // ID Database SK
  
  // ID Lainnya (Biarkan saja jika nanti dipakai modul lain)
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

// 2. FOLDER CONFIG (Digunakan oleh semua file .gs lainnya)
const FOLDER_CONFIG = {
  MAIN_SK: "1GwIow8B4O1OWoq3nhpzDbMO53LXJJUKs", // Folder Utama SK
  TRASH_SK: "1OB2Mxa_zvpYl7Vru9NEddYmBlU5SfYHL", // Folder Sampah SK
  
  // Folder Lainnya
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

// ==========================================
// 2. CORE WEB APP (DoGet & Routing)
// ==========================================
function doGet(e) {
  var template = HtmlService.createTemplateFromFile('index');
  return template.evaluate()
      .setTitle('SIKS - REBORN')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

function getScriptUrl() {
  return ScriptApp.getService().getUrl();
}

// Routing Halaman (KEMBALI KE NAMA ASLI 'getHalaman')
function getHalaman(namaFile) {
  try {
    const prefix = "page_";
    const realName = namaFile.startsWith(prefix) ? namaFile : prefix + namaFile;
    return HtmlService.createTemplateFromFile(realName).evaluate().getContent();
  } catch (err) {
    return '<div class="alert alert-danger p-3">Halaman <b>' + namaFile + '</b> belum dibuat atau nama file salah.</div>';
  }
}

// Alias untuk loadPage (jaga-jaga jika ada script lain yang memanggil)
function loadPage(namaFile) { return getHalaman(namaFile); }

// ==========================================
// 3. AUTH SYSTEM (MANUAL LOGIN)
// ==========================================

// A. PROSES CEK PASSWORD (SAAT TOMBOL LOGIN DITEKAN)
function processLogin(formObj) {
  try {
    // Normalisasi input (bisa objek form atau parameter terpisah)
    var inputUser = "";
    var inputPass = "";
    
    if (typeof formObj === 'object' && formObj.username) {
      inputUser = String(formObj.username).trim();
      inputPass = String(formObj.password).trim();
    } else {
      // Jika dipanggil manual processLogin('admin', '123')
      inputUser = String(arguments[0]).trim();
      inputPass = String(arguments[1]).trim();
    }

    var ss = SpreadsheetApp.openById(SPREADSHEET_IDS.DATABASE_USER); 
    var sheet = ss.getSheetByName(SPREADSHEET_IDS.SHEET_USER_NAME);
    var data = sheet.getDataRange().getValues();

    // Loop Database (Mulai baris 1)
    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      // Kolom A (0) = Username, Kolom B (1) = Password
      if (String(row[0]).trim() == inputUser && String(row[1]).trim() == inputPass) {
        
        // LOGIN SUKSES!
        var userObj = {
          username: row[0],
          fullName: row[2], // Kolom C: Nama Lengkap
          role: row[3],     // Kolom D: Role (Admin/User)
          photo: row[4] || "", // Kolom E: Foto
          isLoggedIn: true
        };
        
        // SIMPAN SESI KE USER PROPERTIES (Aman per akun Google)
        PropertiesService.getUserProperties().setProperty('currentUser', JSON.stringify(userObj));
        
        return { status: 'success', message: 'Login Berhasil' };
      }
    }

    return { status: 'error', message: 'Username atau Password Salah.' };

  } catch (e) {
    return { status: 'error', message: 'Error Server: ' + e.toString() };
  }
}

// B. AMBIL DATA USER (DIPANGGIL OLEH HOME)
function getCurrentUser() {
  try {
    // Ambil data dari penyimpanan sementara (UserProperties)
    var userStr = PropertiesService.getUserProperties().getProperty('currentUser');
    if (userStr) {
      return JSON.parse(userStr);
    }
    return null; // Belum login
  } catch(e) {
    return null;
  }
}

// C. LOGOUT
function processLogout() {
  PropertiesService.getUserProperties().deleteProperty('currentUser');
  return { status: 'success' };
}


// ==========================================
// 4. VISITOR COUNTER & SETTING
// ==========================================
function getVisitorStats() {
  var props = PropertiesService.getScriptProperties();
  var today = new Date().toLocaleDateString("id-ID"); 
  
  // Statistik
  var totalHits = Number(props.getProperty('TOTAL_HITS')) || 0;
  var lastDate = props.getProperty('LAST_DATE_HIT');
  var todayHits = Number(props.getProperty('TODAY_HITS')) || 0;

  if (lastDate !== today) {
    todayHits = 0;
    props.setProperty('LAST_DATE_HIT', today);
  }

  totalHits++;
  todayHits++;
  props.setProperty('TOTAL_HITS', totalHits.toString());
  props.setProperty('TODAY_HITS', todayHits.toString());

  // Running Text
  var totalUsers = 0;
  var infoText = "Selamat Datang di SIKS-REBORN";

  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_IDS.DATABASE_USER);
    // Hitung User
    var sheetUser = ss.getSheetByName(SPREADSHEET_IDS.SHEET_USER_NAME);
    if(sheetUser) totalUsers = sheetUser.getLastRow() - 1;

    // Ambil Running Text
    var sheetSetting = ss.getSheetByName("SETTING");
    if (sheetSetting) infoText = sheetSetting.getRange("B1").getValue();

  } catch (e) {
    infoText = "Maintenance Mode";
  }

  return { total: totalHits, today: todayHits, users: totalUsers, info: infoText };
}

function saveRunningText(textBaru) {
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_IDS.DATABASE_USER);
    var sheet = ss.getSheetByName("SETTING");
    if (!sheet) {
      sheet = ss.insertSheet("SETTING");
      sheet.getRange("A1").setValue("RUNNING_TEXT");
    }
    sheet.getRange("B1").setValue(textBaru);
    return { status: 'success', message: 'Berhasil disimpan!' };
  } catch (e) {
    return { status: 'error', message: 'Gagal: ' + e.message };
  }
}

// Untuk memuat halaman Setting di Sidebar
function loadPageSetting() {
  return HtmlService.createTemplateFromFile('page_setting').evaluate().getContent();
}

// ==========================================
// 5. MONITORING SYSTEM (CCTV)
// ==========================================
function logUserVisit(userData) {
  if (!userData) return;

  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_IDS.DATABASE_USER);
    var sheet = ss.getSheetByName("LOG_ACCESS");
    
    var now = new Date();
    var timestamp = Utilities.formatDate(now, "Asia/Jakarta", "dd/MM/yyyy HH:mm:ss");
    var tgalOnly  = Utilities.formatDate(now, "Asia/Jakarta", "yyyy-MM-dd");
    var blnOnly   = Utilities.formatDate(now, "Asia/Jakarta", "yyyy-MM");
    
    // Cek Hari Libur
    var dayIndex = now.getDay(); 
    var jenisHari = "Hari Efektif";
    var ketHari = "Reguler";

    // 1. Cek Weekend
    if (dayIndex === 0 || dayIndex === 6) {
      jenisHari = "Hari Libur";
      ketHari = (dayIndex === 0) ? "Minggu" : "Sabtu";
    }

    // 2. Cek Kalender Libur (DATA_LIBUR)
    var sheetLibur = ss.getSheetByName("DATA_LIBUR");
    if (sheetLibur && sheetLibur.getLastRow() > 1) {
      var dataLibur = sheetLibur.getRange(2, 1, sheetLibur.getLastRow()-1, 2).getValues();
      for (var i = 0; i < dataLibur.length; i++) {
        var tglLibur = Utilities.formatDate(new Date(dataLibur[i][0]), "Asia/Jakarta", "yyyy-MM-dd");
        if (tglLibur === tgalOnly) {
          jenisHari = "Hari Libur";
          ketHari = dataLibur[i][1];
          break;
        }
      }
    }

    // 3. Simpan
    if (sheet) {
      sheet.appendRow([timestamp, tgalOnly, blnOnly, userData.fullName, userData.role, jenisHari + " (" + ketHari + ")"]);
    }
    
  } catch (e) {
    console.log("Log Error: " + e.message);
  }
}

function getMonitoringStats() {
  var ss = SpreadsheetApp.openById(SPREADSHEET_IDS.DATABASE_USER);
  
  // Ambil Data Log
  var sheetLog = ss.getSheetByName("LOG_ACCESS");
  var dataLog = [];
  if (sheetLog && sheetLog.getLastRow() > 1) {
    dataLog = sheetLog.getRange(2, 1, sheetLog.getLastRow() - 1, 6).getValues();
  }

  var stats = {
    total: dataLog.length,
    kerja: 0,
    libur: 0,
    userCounts: {}, 
    daily: {},
    weekly: {},
    monthly: {}
  };

  dataLog.forEach(function(row) {
    var timestamp = row[0];
    var tgal = row[1];
    var nama = row[3];
    var jenis = row[5];

    if (String(jenis).includes("Libur")) stats.libur++;
    else stats.kerja++;

    stats.userCounts[nama] = (stats.userCounts[nama] || 0) + 1;
    stats.daily[tgal] = (stats.daily[tgal] || 0) + 1;

    var dateObj = new Date(timestamp);
    var namaBulan = Utilities.formatDate(dateObj, "Asia/Jakarta", "MMMM yyyy");
    stats.monthly[namaBulan] = (stats.monthly[namaBulan] || 0) + 1;

    var weekNum = Utilities.formatDate(dateObj, "Asia/Jakarta", "w");
    var weekLabel = "Minggu ke-" + weekNum;
    stats.weekly[weekLabel] = (stats.weekly[weekLabel] || 0) + 1;
  });

  var rankingUser = [];
  Object.keys(stats.userCounts).forEach(function(name){
    rankingUser.push({ name: name, count: stats.userCounts[name] });
  });
  rankingUser.sort(function(a, b){ return b.count - a.count });

  // Cari User Pasif
  var sheetUser = ss.getSheetByName(SPREADSHEET_IDS.SHEET_USER_NAME);
  var userPasif = [];
  if (sheetUser) {
    // Ambil kolom Nama (C)
    var allUsers = sheetUser.getRange(2, 3, sheetUser.getLastRow()-1, 1).getValues(); 
    allUsers.forEach(function(u){
      var uName = u[0];
      if (uName && !stats.userCounts[uName]) userPasif.push(uName);
    });
  }

  return {
    summary: { total: stats.total, kerja: stats.kerja, libur: stats.libur },
    topUsers: rankingUser.slice(0, 10),
    passiveUsers: userPasif,
    chartData: { daily: stats.daily, weekly: stats.weekly, monthly: stats.monthly }
  };
}