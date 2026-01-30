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

/* ======================================================================
   CORE WEB APP: DO GET & INCLUDE
   ====================================================================== */
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

/* ======================================================================
   ROUTING HALAMAN
   ====================================================================== */
function getHalaman(namaFile) {
  try {
    const prefix = "page_";
    // Cek apakah nama file sudah ada prefix 'page_' atau belum
    const realName = namaFile.startsWith(prefix) ? namaFile : prefix + namaFile;
    return HtmlService.createTemplateFromFile(realName).evaluate().getContent();
  } catch (err) {
    return '<div class="p-4"><div class="alert alert-warning">Halaman <b>' + namaFile + '</b> belum tersedia / file tidak ditemukan.</div></div>';
  }
}

/* ======================================================================
   SISTEM LOGIN (AUTH)
   ====================================================================== */
function checkLogin(username, password) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_IDS.DATABASE_USER);
  const sheet = ss.getSheetByName(SPREADSHEET_IDS.SHEET_USER_NAME);
  const data = sheet.getDataRange().getValues();
  
  for (let i = 1; i < data.length; i++) {
    // Kolom A=Username, B=Password
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
  var ss = SpreadsheetApp.openById(SPREADSHEET_IDS.DATABASE_USER); 
  var sheetName = SPREADSHEET_IDS.SHEET_USER_NAME; 
  var sheet = ss.getSheetByName(sheetName);
  
  if (!sheet) {
    return { status: "error", message: "Sheet '" + sheetName + "' tidak ditemukan!" };
  }

  var data = sheet.getDataRange().getValues();
  var inputUser = formObject.username ? formObject.username.toString().trim() : "";
  var inputPass = formObject.password ? formObject.password.toString().trim() : "";

  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    var dbUser = row[0] ? row[0].toString().trim() : "";
    var dbPass = row[1] ? row[1].toString().trim() : "";
    
    if (dbUser === inputUser && dbPass === inputPass) {
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
        nama: row[2], 
        role: row[3], 
        foto: row[4]  
      };
    }
  }

  return { status: "error", message: "Username atau Password salah" };
}

function processLogout() {
  PropertiesService.getUserProperties().deleteProperty('currentUser');
}