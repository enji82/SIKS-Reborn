/* ======================================================================
   MODUL: KELOLA PTK SD
   Spreadsheet ID: 1t0-Lmy0YD_GxHzimFWJGh5R5x6RhGL13uqKeVwWoCYE
   Sheet: Master Data GTK
   ====================================================================== */

var ID_DB_PTK = "1t0-Lmy0YD_GxHzimFWJGh5R5x6RhGL13uqKeVwWoCYE";
var SHEET_PTK = "Master Data GTK";

// 1. AMBIL OPSI FILTER (UNIT & STATUS)
function getFilterOptionsPTK() {
  try {
    var ss = SpreadsheetApp.openById(ID_DB_PTK);
    var sheet = ss.getSheetByName(SHEET_PTK);
    if (!sheet) return JSON.stringify({ units: [], statuses: [] });
    
    // Ambil Kolom C (Unit) dan S (Status)
    var lastRow = sheet.getLastRow();
    if(lastRow < 2) return JSON.stringify({ units: [], statuses: [] });

    // Ambil data Unit (C/Index 2) dan Status (S/Index 18)
    // Kita ambil range besar sekalian biar 1x call
    var data = sheet.getRange(2, 1, lastRow - 1, 19).getValues(); 
    
    var unitSet = new Set();
    var statusSet = new Set();
    
    for(var i=0; i<data.length; i++){
        if(data[i][2]) unitSet.add(String(data[i][2]).trim());
        if(data[i][18]) statusSet.add(String(data[i][18]).trim());
    }
    
    return JSON.stringify({
        units: Array.from(unitSet).sort(),
        statuses: Array.from(statusSet).sort()
    });
  } catch(e) { return JSON.stringify({ error: e.message }); }
}

// 2. AMBIL DATA UTAMA (OPTIMASI DISPLAY VALUES)
function getDataPTKSD(filterUnit, filterStatus) {
  var ss = SpreadsheetApp.openById(ID_DB_PTK);
  var sheet = ss.getSheetByName(SHEET_PTK);
  var data = sheet.getDataRange().getValues();
  data.shift(); 
  
  var result = [];
  
  for (var i = 0; i < data.length; i++) {
    var row = data[i];
    
    var tglLahirISO = parseIndoDate(row[9]);
    var tmtJabISO   = parseIndoDate(row[20]);
    var tmtGolISO   = parseIndoDate(row[22]);

    result.push({
      id: row[0],
      npsn: row[1],
      unit: row[2],
      gelar_depan: row[3],
      nama_no_gelar: row[4],
      gelar_belakang: row[5],
      nama_lengkap: row[6],
      nip: row[7],
      tmp_lahir: row[8],
      tgl_lahir: tglLahirISO,
      nik: row[10],
      lp: row[11],
      agama: row[12],
      pendidikan: row[13],
      jurusan: row[14],
      thn_lulus: row[15],
      alamat: row[16],
      hp: row[17],
      status_peg: row[18],
      jabatan: row[19],
      tmt_jabatan: tmtJabISO,
      pangkat: row[21],
      tmt_gol: tmtGolISO,
      mkg: row[23],
      kelas_jab: row[24],
      tugas: row[25],
      nuptk: row[26],
      serdik: row[27],
      dapodik: row[28],
      tugtam: row[29],
      jabatan_guru: row[30], // AE: Sekarang Jabatan Guru
      diinput: row[31] ? Utilities.formatDate(new Date(row[31]), Session.getScriptTimeZone(), "dd/MM/yy HH:mm") : "",
      user_input: row[32],
      diedit: row[33] ? Utilities.formatDate(new Date(row[33]), Session.getScriptTimeZone(), "dd/MM/yy HH:mm") : "",
      user_edit: row[34]
    });
  }
  
  return JSON.stringify(result);
}

// 3. UPDATE DATA PTK
function updateDataPTK(form) {
  try {
    var ss = SpreadsheetApp.openById(ID_DB_PTK);
    var sheet = ss.getSheetByName(SHEET_PTK);
    var data = sheet.getDataRange().getValues();
    
    var rowIndex = -1;
    // Cari baris berdasarkan ID (Kolom A)
    for(var i=1; i<data.length; i++){
        if(String(data[i][0]) === String(form.id)){
            rowIndex = i + 1; // +1 karena index mulai 0
            break;
        }
    }
    
    if(rowIndex === -1) return "Error: ID PTK tidak ditemukan.";

    var now = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd-MM-yyyy HH:mm:ss");
    var user = form.user_login || "Admin";

    // Update Kolom Tertentu (Mapping form ke Kolom Excel)
    // Kita update data utama saja sesuai form edit
    // Note: getRange(row, col) -> col mulai dari 1 (A=1)
    
    sheet.getRange(rowIndex, 7).setValue(form.nama_lengkap); // G
    sheet.getRange(rowIndex, 8).setValue("'"+form.nip);      // H (Pakai petik biar text)
    sheet.getRange(rowIndex, 9).setValue(form.tmp_lahir);    // I
    sheet.getRange(rowIndex, 10).setValue("'"+form.tgl_lahir); // J
    sheet.getRange(rowIndex, 11).setValue("'"+form.nik);     // K
    
    sheet.getRange(rowIndex, 14).setValue(form.pendidikan);  // N
    sheet.getRange(rowIndex, 15).setValue(form.jurusan);     // O
    sheet.getRange(rowIndex, 16).setValue(form.thn_lulus);   // P
    
    sheet.getRange(rowIndex, 18).setValue(form.status_peg);  // R (Hp) -> S (Status)
    sheet.getRange(rowIndex, 19).setValue(form.status_peg);  // S
    
    sheet.getRange(rowIndex, 20).setValue(form.jabatan);     // T
    sheet.getRange(rowIndex, 21).setValue("'"+form.tmt_jabatan); // U
    
    sheet.getRange(rowIndex, 22).setValue(form.pangkat);     // V
    sheet.getRange(rowIndex, 23).setValue("'"+form.tmt_gol); // W
    
    sheet.getRange(rowIndex, 24).setValue(form.kelas_jab);   // X
    sheet.getRange(rowIndex, 25).setValue(form.tugas);       // Y
    sheet.getRange(rowIndex, 26).setValue("'"+form.nuptk);   // Z
    sheet.getRange(rowIndex, 27).setValue(form.serdik);      // AA
    sheet.getRange(rowIndex, 28).setValue(form.dapodik);     // AB
    sheet.getRange(rowIndex, 29).setValue(form.tugtam);      // AC

    // Log Edit (Kolom AG=33, AH=34)
    sheet.getRange(rowIndex, 33).setValue(now);
    sheet.getRange(rowIndex, 34).setValue(user);

    return "Sukses";
  } catch(e) { return "Error: " + e.message; }
}

/* ======================================================================
   MODUL: REFERENSI & INSERT PTK (AUTO FILL)
   ====================================================================== */

// 1. AMBIL DATA REFERENSI (JABATAN, PANGKAT, TUGAS)
function getReferensiPTK() {
  var ss = SpreadsheetApp.openById(ID_DB_PTK);
  
  function getList(sheetName) {
    var s = ss.getSheetByName(sheetName);
    if (!s) return [];
    var last = s.getLastRow();
    if (last < 2) return [];
    // Ambil Kolom A, B, C (Untuk mapping Jabatan Guru butuh 3 kolom)
    return s.getRange(2, 1, last - 1, 3).getValues(); 
  }

  return JSON.stringify({
    jabatan: getList("data_jabatan"), // [Nama, Kelas]
    pangkat: getList("data_pangkat"), // [Nama]
    tugas: getList("data_tugas"),     // [Nama]
    // Sheet jabatan_guru: Col A (Tugas), Col B (Jabatan), Col C (Hasil)
    mapping_jabatan: getList("jabatan_guru") 
  });
}

// 2. INSERT DATA PTK (AUTO FILL LOGIC)
function insertDataPTK(form) {
  var ss = SpreadsheetApp.openById("1t0-Lmy0YD_GxHzimFWJGh5R5x6RhGL13uqKeVwWoCYE"); // ID Database
  var sheet = ss.getSheetByName("Master Data GTK");
  if (!sheet) return "Error: Sheet 'Master Data GTK' tidak ditemukan.";

  // --- VALIDASI DUPLIKAT NIP (KHUSUS SDN) ---
  var inputNip = String(form.nip).trim().replace(/[^0-9]/g, ''); // Ambil angka saja

  // Hanya validasi jika User mengisi NIP (Jika '-' atau kosong, lewati)
  if (inputNip !== "" && inputNip !== "-") {
      var data = sheet.getDataRange().getValues();
      
      // Loop semua baris (Mulai index 1)
      for (var i = 1; i < data.length; i++) {
        // Kolom H (NIP) ada di index 7 (0-based)
        var rowNip = String(data[i][7]).replace(/[^0-9]/g, ''); 
    
        if (rowNip === inputNip) {
          // Jika NIP sama, ambil Nama Pemilik (Kolom G / Index 6)
          var namaPemilik = data[i][6]; 
          
          return "NIP " + inputNip + " sudah terdaftar atas nama " + namaPemilik + ", hubungi admin Korwil untuk melanjutkan.";
        }
      }
  }
  // -------------------------------------------

  // Generate ID
  var newId = "GTK-" + new Date().getTime();
  
  var namaFull = (form.gelar_depan ? form.gelar_depan + " " : "") + 
                 form.nama_lengkap + 
                 (form.gelar_belakang ? ", " + form.gelar_belakang : "");
                 
  var now = new Date();
  var timestamp = Utilities.formatDate(now, "Asia/Jakarta", "dd/MM/yyyy HH:mm:ss");

  // Array Data (Sesuaikan dengan urutan kolom SDN)
  var rowData = [
    newId,                  
    form.npsn_login || "",  
    form.unit_login || "",  
    form.gelar_depan,       
    form.nama_lengkap,      
    form.gelar_belakang,    
    namaFull,               
    "'" + form.nip,         // Validasi dilakukan di atas
    form.tmp_lahir,         
    form.tgl_lahir,         
    "'" + form.nik,         
    form.lp,                
    form.agama,             
    form.pendidikan,        
    form.jurusan,           
    form.thn_lulus,         
    form.alamat,            
    "'" + form.hp,          
    form.status_peg,        
    form.jabatan,           
    form.tmt_jabatan,       
    form.pangkat,           // SDN pakai Pangkat/Gol
    form.tmt_pangkat,       
    form.masa_kerja_thn,
    form.masa_kerja_bln,
    form.gaji_pokok,
    "'" + form.nuptk,       
    form.serdik,            
    form.dapodik,           
    form.tugtam,            
    timestamp,              
    form.user_login,        
    "",                     
    ""                      
  ];

  sheet.appendRow(rowData);
  return "Sukses";
}

/* ======================================================================
   HELPER: PARSE TANGGAL CERDAS (ISO, SLASH, INDO TEXT)
   ====================================================================== */
function parseIndoDate(dateStr) {
  if (!dateStr || dateStr === "-" || dateStr === "") return "";
  
  var str = String(dateStr).trim();

  // 1. Cek jika sudah format ISO (yyyy-MM-dd) -> Cocok untuk HTML Date
  if (str.match(/^\d{4}-\d{2}-\d{2}$/)) return str;

  // 2. Cek format Slash (dd/MM/yyyy) -> Contoh: 31/12/1990
  var slashMatch = str.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
  if (slashMatch) {
    var day = slashMatch[1].length === 1 ? "0" + slashMatch[1] : slashMatch[1];
    var month = slashMatch[2].length === 1 ? "0" + slashMatch[2] : slashMatch[2];
    var year = slashMatch[3];
    return year + "-" + month + "-" + day;
  }

  // 3. Cek format Indo Teks (dd MMMM yyyy) -> Contoh: 17 Agustus 1945
  var months = {
    'Januari': '01', 'Februari': '02', 'Maret': '03', 'April': '04', 'Mei': '05', 'Juni': '06',
    'Juli': '07', 'Agustus': '08', 'September': '09', 'Oktober': '10', 'November': '11', 'Desember': '12',
    'Jan': '01', 'Feb': '02', 'Mar': '03', 'Apr': '04', 'Jun': '06', 'Jul': '07', 'Agu': '08', 'Sep': '09', 'Okt': '10', 'Nov': '11', 'Des': '12' 
  };

  var parts = str.split(' '); 
  if (parts.length >= 3) {
    // Ambil bagian angka pertama sebagai tanggal (buang karakter non-digit jika ada)
    var dayRaw = parts[0].replace(/[^0-9]/g, ''); 
    var day = dayRaw.length === 1 ? "0" + dayRaw : dayRaw;
    
    var monthName = parts[1];
    var year = parts[2];
    
    var month = months[monthName];
    
    if (month && year.match(/^\d{4}$/)) {
        return year + "-" + month + "-" + day;
    }
  }
    
  // 4. Fallback: Coba Parse sebagai Object Date (Excel Serial Number)
  try {
    var d = new Date(dateStr);
    if (!isNaN(d.getTime())) {
      // Pastikan timezone sesuai script (Jakarta)
      return Utilities.formatDate(d, Session.getScriptTimeZone(), "yyyy-MM-dd");
    }
  } catch(e) {}
    
  return ""; // Nyerah, balikin kosong
}

/* ======================================================================
   MODUL: HAPUS DATA PTK (MOVE TO NON-AKTIF)
   ====================================================================== */
function moveDataPTKToNonAktif(id, reason, userLogin) {
  try {
    var ss = SpreadsheetApp.openById(ID_DB_PTK);
    var sheetSource = ss.getSheetByName(SHEET_PTK);
    var sheetTarget = ss.getSheetByName("gtk_non_aktif"); // Pastikan sheet ini ada!
    
    // Jika sheet target belum ada, buat baru (Opsional/Safety)
    if (!sheetTarget) {
      sheetTarget = ss.insertSheet("gtk_non_aktif");
      // Copy header dari source
      var headers = sheetSource.getRange(1, 1, 1, sheetSource.getLastColumn()).getValues();
      // Tambah header pelengkap
      headers[0].push("Alasan Hapus", "Tanggal Hapus", "User Hapus");
      sheetTarget.getRange(1, 1, 1, headers[0].length).setValues(headers);
    }

    var data = sheetSource.getDataRange().getValues();
    var rowIndex = -1;

    // Cari Baris berdasarkan ID (Kolom A / Index 0)
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][0]) === String(id)) {
        rowIndex = i;
        break;
      }
    }

    if (rowIndex === -1) return "Data tidak ditemukan.";

    // Ambil Data Baris Tersebut
    var rowData = data[rowIndex];
    
    // Tambahkan Info Penghapusan
    var deleteTime = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd-MM-yyyy HH:mm:ss");
    rowData.push(reason, deleteTime, userLogin);

    // 1. Simpan ke Sheet Non Aktif
    sheetTarget.appendRow(rowData);

    // 2. Hapus dari Sheet Utama (Perhatikan +1 karena array 0-based vs sheet 1-based)
    sheetSource.deleteRow(rowIndex + 1);

    return "Sukses";

  } catch (e) {
    return "Error: " + e.message;
  }
}

function getDataKeadaanGTK() {
  var id = "1t0-Lmy0YD_GxHzimFWJGh5R5x6RhGL13uqKeVwWoCYE"; // ID Spreadsheet
  var ss = SpreadsheetApp.openById(id);
  var sheet = ss.getSheetByName("Keadaan GTK");
  if (!sheet) return [];
  
  // Asumsi Data mulai dari Baris 3 (Karena header bertingkat 2 baris)
  // Kolom A sampai BD (56 Kolom)
  var lastRow = sheet.getLastRow();
  if (lastRow < 3) return [];
  
  // Ambil Range A3:BD_LastRow
  var data = sheet.getRange(3, 1, lastRow - 2, 56).getDisplayValues();
  return data;
}

function getDataKebutuhanGuru() {
  var id = "1t0-Lmy0YD_GxHzimFWJGh5R5x6RhGL13uqKeVwWoCYE"; 
  var ss = SpreadsheetApp.openById(id);
  var sheet = ss.getSheetByName("Kebutuhan Guru");
  if (!sheet) return [];
  
  // Data mulai baris 3 (Header 2 baris)
  var lastRow = sheet.getLastRow();
  if (lastRow < 3) return [];
  
  // Ambil A3:AP
  // A=1, AP=42
  var data = sheet.getRange(3, 1, lastRow - 2, 42).getDisplayValues();
  return data;
}

// =============================================================
// BACKEND: KELOLA DATA PTK SD SWASTA (SDS) - REVISI ID
// =============================================================

var ID_SPREADSHEET_PTK = "1t0-Lmy0YD_GxHzimFWJGh5R5x6RhGL13uqKeVwWoCYE"; // ID Database Utama

/**
 * 1. GET DATA (READ)
 */
function getDataPTKSDS() {
  try {
    var ss = SpreadsheetApp.openById(ID_SPREADSHEET_PTK); // FIX: Pakai OpenById
    var sheet = ss.getSheetByName("Master Data GTK SDS");
    if (!sheet) return JSON.stringify([]);

    var lastRow = sheet.getLastRow();
    if (lastRow < 2) return JSON.stringify([]); 

    // Ambil Data A2:AE (Index 1 s.d 31)
    var data = sheet.getRange(2, 1, lastRow - 1, 31).getDisplayValues();
    
    var result = [];
    for (var i = 0; i < data.length; i++) {
      var row = data[i];
      if(row[0] === "") continue; 

      result.push({
        id: row[0],             
        npsn: row[1],           
        unit: row[2],           
        gelar_depan: row[3],    
        nama_no_gelar: row[4],  
        gelar_belakang: row[5], 
        nama_lengkap: row[6],   
        niy: row[7],            
        tmp_lahir: row[8],      
        tgl_lahir: row[9],      
        nik: row[10],           
        lp: row[11],            
        agama: row[12],         
        pendidikan: row[13],    
        jurusan: row[14],       
        thn_lulus: row[15],     
        alamat: row[16],        
        hp: row[17],            
        status_peg: row[18],    
        jabatan: row[19],       
        tmt_jabatan: row[20],   
        inpassing: row[21],     
        tmt_inpassing: row[22], 
        nuptk: row[23],         
        serdik: row[24],        
        dapodik: row[25],       
        tugtam: row[26],        
        diinput: row[27],       
        user_input: row[28],    
        diedit: row[29],        
        user_edit: row[30]      
      });
    }
    return JSON.stringify(result);
    
  } catch(e) {
    return JSON.stringify([]); // Return array kosong jika error koneksi
  }
}

/**
 * 2. INSERT DATA (CREATE)
 */
function insertDataPTKSDS(form) {
  var ss = SpreadsheetApp.openById(ID_SPREADSHEET_PTK);
  var sheet = ss.getSheetByName("Master Data GTK SDS");
  if (!sheet) return "Error: Sheet SDS tidak ditemukan.";

  // --- VALIDASI DUPLIKAT NIK (Added Logic) ---
  var data = sheet.getDataRange().getValues();
  var inputNik = String(form.nik).trim(); // NIK yang diinput user

  // Loop semua baris (Mulai index 1 untuk lewati header)
  for (var i = 1; i < data.length; i++) {
    // Kolom K (NIK) ada di index 10 (0-based)
    // Kita hapus tanda kutip (') jika ada, agar perbandingan akurat
    var rowNik = String(data[i][10]).replace(/'/g, "").trim(); 

    if (rowNik === inputNik) {
      // Jika NIK sama, ambil Nama Pemilik (Kolom G / Index 6)
      var namaPemilik = data[i][6]; 
      
      // Return Pesan Error Spesifik (Proses Simpan Dibatalkan)
      return "NIK " + inputNik + " sudah terdaftar atas nama " + namaPemilik + ", hubungi admin Korwil untuk melanjutkan.";
    }
  }
  // -------------------------------------------

  // Jika Lolos Validasi, Lanjut Simpan
  var newId = "SDS-" + new Date().getTime();
  
  var namaFull = (form.gelar_depan ? form.gelar_depan + " " : "") + 
                 form.nama_lengkap + 
                 (form.gelar_belakang ? ", " + form.gelar_belakang : "");
                 
  var now = new Date();
  var timestamp = Utilities.formatDate(now, "Asia/Jakarta", "dd/MM/yyyy HH:mm:ss");

  var rowData = [
    newId,                  
    form.npsn_login || "",  
    form.unit_login || "",  
    form.gelar_depan,       
    form.nama_lengkap,      
    form.gelar_belakang,    
    namaFull,               
    form.niy,               
    form.tmp_lahir,         
    form.tgl_lahir,         
    "'" + form.nik,         
    form.lp,                
    form.agama,             
    form.pendidikan,        
    form.jurusan,           
    form.thn_lulus,         
    form.alamat,            
    "'" + form.hp,          
    form.status_peg,        
    form.jabatan,           
    form.tmt_jabatan,       
    form.inpassing,         
    form.tmt_inpassing,     
    "'" + form.nuptk,       
    form.serdik,            
    form.dapodik,           
    form.tugtam,            
    timestamp,              
    form.user_login,        
    "",                     
    ""                      
  ];

  sheet.appendRow(rowData);
  return "Sukses";
}

/**
 * 3. UPDATE DATA (EDIT)
 */
function updateDataPTKSDS(form) {
  var ss = SpreadsheetApp.openById(ID_SPREADSHEET_PTK); // FIX
  var sheet = ss.getSheetByName("Master Data GTK SDS");
  var data = sheet.getDataRange().getValues();
  
  var rowIdx = -1;
  for (var i = 0; i < data.length; i++) {
    if (data[i][0] == form.id) {
      rowIdx = i + 1; 
      break;
    }
  }
  
  if (rowIdx == -1) return "Error: ID tidak ditemukan.";

  var namaFull = (form.gelar_depan ? form.gelar_depan + " " : "") + 
                 form.nama_lengkap + 
                 (form.gelar_belakang ? ", " + form.gelar_belakang : "");

  var now = new Date();
  var timestamp = Utilities.formatDate(now, "Asia/Jakarta", "dd/MM/yyyy HH:mm:ss");

  // Update Kolom D (4) s.d AA (27)
  var range = sheet.getRange(rowIdx, 4, 1, 24); 
  var updateValues = [[
    form.gelar_depan,       
    form.nama_lengkap,      
    form.gelar_belakang,    
    namaFull,               
    form.niy,               
    form.tmp_lahir,         
    form.tgl_lahir,         
    "'" + form.nik,         
    form.lp,                
    form.agama,             
    form.pendidikan,        
    form.jurusan,           
    form.thn_lulus,         
    form.alamat,            
    "'" + form.hp,          
    form.status_peg,        
    form.jabatan,           
    form.tmt_jabatan,       
    form.inpassing,         
    form.tmt_inpassing,     
    "'" + form.nuptk,       
    form.serdik,            
    form.dapodik,           
    form.tugtam             
  ]];
  
  range.setValues(updateValues);

  // Update Info Diedit (AD & AE)
  sheet.getRange(rowIdx, 30).setValue(timestamp);      
  sheet.getRange(rowIdx, 31).setValue(form.user_login); 

  return "Sukses";
}

/**
 * 4. DELETE DATA
 */
function deleteDataPTKSDS(id, alasan, userLogin) {
  var ss = SpreadsheetApp.openById(ID_SPREADSHEET_PTK); // FIX
  var sheet = ss.getSheetByName("Master Data GTK SDS");
  var sheetArsip = ss.getSheetByName("Arsip PTK Non Aktif"); 
  
  if (!sheetArsip) {
    sheetArsip = ss.insertSheet("Arsip PTK Non Aktif");
    sheetArsip.appendRow(["ID", "Nama", "Unit", "Alasan Hapus", "User Hapus", "Waktu Hapus", "Data Asli JSON"]);
  }

  var data = sheet.getDataRange().getValues();
  var rowIdx = -1;
  var rowData = [];

  for (var i = 0; i < data.length; i++) {
    if (data[i][0] == id) {
      rowIdx = i + 1;
      rowData = data[i];
      break;
    }
  }

  if (rowIdx == -1) return "Error: Data tidak ditemukan.";

  var now = new Date();
  var timestamp = Utilities.formatDate(now, "Asia/Jakarta", "dd/MM/yyyy HH:mm:ss");

  sheetArsip.appendRow([
    rowData[0], 
    rowData[6], 
    rowData[2], 
    alasan,
    userLogin,
    timestamp,
    JSON.stringify(rowData) 
  ]);

  sheet.deleteRow(rowIdx);

  return "Sukses";
}

// ==========================================
// DATA KEADAAN GTK SDS (Untuk Halaman Laporan)
// ==========================================

function getDataKeadaanGTKSDS() {
  var id = "1t0-Lmy0YD_GxHzimFWJGh5R5x6RhGL13uqKeVwWoCYE"; // ID Spreadsheet Database
  var ss = SpreadsheetApp.openById(id);
  var sheet = ss.getSheetByName("Keadaan GTK SDS");
  if (!sheet) return [];
  
  var lastRow = sheet.getLastRow();
  if (lastRow < 3) return []; // Header 2 baris
  
  // Ambil Range A3:AA
  // A=1, AA=27
  var data = sheet.getRange(3, 1, lastRow - 2, 27).getDisplayValues();
  return data;
}

// ==========================================
// DATA KEBUTUHAN GURU SDS
// ==========================================

function getDataKebutuhanGuruSDS() {
  var id = "1t0-Lmy0YD_GxHzimFWJGh5R5x6RhGL13uqKeVwWoCYE"; 
  var ss = SpreadsheetApp.openById(id);
  var sheet = ss.getSheetByName("Kebutuhan Guru SDS");
  if (!sheet) return [];
  
  var lastRow = sheet.getLastRow();
  if (lastRow < 3) return [];
  
  // Ambil A3:AA
  // A=1, AA=27
  var data = sheet.getRange(3, 1, lastRow - 2, 27).getDisplayValues();
  return data;
}

// =============================================================
// BACKEND: KELOLA DATA PTK PAUD
// ID Spreadsheet: 1XetGkBymmN2NZQlXpzZ2MQyG0nhhZ0sXEPcNsLffhEU
// =============================================================

var ID_SPREADSHEET_PAUD = "1XetGkBymmN2NZQlXpzZ2MQyG0nhhZ0sXEPcNsLffhEU";

/**
 * 1. GET DATA PTK PAUD
 */
function getDataPTKPAUD() {
  try {
    var ss = SpreadsheetApp.openById(ID_SPREADSHEET_PAUD);
    var sheet = ss.getSheetByName("Master Data GTK PAUD");
    if (!sheet) return JSON.stringify([]);

    var lastRow = sheet.getLastRow();
    if (lastRow < 2) return JSON.stringify([]); 

    // Ambil Data A2:AF (Index 1 s.d 32)
    var data = sheet.getRange(2, 1, lastRow - 1, 32).getDisplayValues();
    
    var result = [];
    for (var i = 0; i < data.length; i++) {
      var row = data[i];
      if(row[0] === "") continue; 

      result.push({
        id: row[0],             
        npsn: row[1],           
        unit: row[2],
        jenjang: row[3],        // D: Jenjang (Baru)
        gelar_depan: row[4],    
        nama_no_gelar: row[5],  
        gelar_belakang: row[6], 
        nama_lengkap: row[7],   
        niy: row[8],            
        tmp_lahir: row[9],      
        tgl_lahir: row[10],      
        nik: row[11],           // L: NIK
        lp: row[12],            
        agama: row[13],         
        pendidikan: row[14],    
        jurusan: row[15],       
        thn_lulus: row[16],     
        alamat: row[17],        
        hp: row[18],            
        status_peg: row[19],    
        jabatan: row[20],       
        tmt_jabatan: row[21],   
        inpassing: row[22],     
        tmt_inpassing: row[23], 
        nuptk: row[24],         
        serdik: row[25],        
        dapodik: row[26],       
        tugtam: row[27],        
        diinput: row[28],       
        user_input: row[29],    
        diedit: row[30],        
        user_edit: row[31]      
      });
    }
    return JSON.stringify(result);
    
  } catch(e) {
    return JSON.stringify([]); 
  }
}

/**
 * 2. INSERT DATA PTK PAUD (VALIDASI NIK)
 */
function insertDataPTKPAUD(form) {
  var ss = SpreadsheetApp.openById(ID_SPREADSHEET_PAUD);
  var sheet = ss.getSheetByName("Master Data GTK PAUD");
  if (!sheet) return "Error: Sheet PAUD tidak ditemukan.";

  // --- VALIDASI DUPLIKAT NIK (Kolom L / Index 11) ---
  var data = sheet.getDataRange().getValues();
  var inputNik = String(form.nik).trim(); 

  for (var i = 1; i < data.length; i++) {
    var rowNik = String(data[i][11]).replace(/'/g, "").trim(); 
    if (rowNik === inputNik) {
      var namaPemilik = data[i][7]; // Nama di Kolom H (Index 7)
      return "NIK " + inputNik + " sudah terdaftar atas nama " + namaPemilik + ", hubungi admin Korwil untuk melanjutkan.";
    }
  }

  var newId = "PAUD-" + new Date().getTime();
  
  var namaFull = (form.gelar_depan ? form.gelar_depan + " " : "") + 
                 form.nama_lengkap + 
                 (form.gelar_belakang ? ", " + form.gelar_belakang : "");
                 
  var now = new Date();
  var timestamp = Utilities.formatDate(now, "Asia/Jakarta", "dd/MM/yyyy HH:mm:ss");

  var rowData = [
    newId,                  
    form.npsn_login || "",  
    form.unit_login || "",
    form.jenjang || "",     // D: Jenjang
    form.gelar_depan,       
    form.nama_lengkap,      
    form.gelar_belakang,    
    namaFull,               
    form.niy,               
    form.tmp_lahir,         
    form.tgl_lahir,         
    "'" + form.nik,         
    form.lp,                
    form.agama,             
    form.pendidikan,        
    form.jurusan,           
    form.thn_lulus,         
    form.alamat,            
    "'" + form.hp,          
    form.status_peg,        
    form.jabatan,           
    form.tmt_jabatan,       
    form.inpassing,         
    form.tmt_inpassing,     
    "'" + form.nuptk,       
    form.serdik,            
    form.dapodik,           
    form.tugtam,            
    timestamp,              
    form.user_login,        
    "",                     
    ""                      
  ];

  sheet.appendRow(rowData);
  return "Sukses";
}

/**
 * 3. UPDATE DATA PTK PAUD
 */
function updateDataPTKPAUD(form) {
  var ss = SpreadsheetApp.openById(ID_SPREADSHEET_PAUD);
  var sheet = ss.getSheetByName("Master Data GTK PAUD");
  var data = sheet.getDataRange().getValues();
  
  var rowIdx = -1;
  for (var i = 0; i < data.length; i++) {
    if (data[i][0] == form.id) {
      rowIdx = i + 1; 
      break;
    }
  }
  
  if (rowIdx == -1) return "Error: ID tidak ditemukan.";

  var namaFull = (form.gelar_depan ? form.gelar_depan + " " : "") + 
                 form.nama_lengkap + 
                 (form.gelar_belakang ? ", " + form.gelar_belakang : "");

  var now = new Date();
  var timestamp = Utilities.formatDate(now, "Asia/Jakarta", "dd/MM/yyyy HH:mm:ss");

  // Update Kolom D (4) s.d AB (28)
  var range = sheet.getRange(rowIdx, 4, 1, 25); 
  var updateValues = [[
    form.jenjang,           // D
    form.gelar_depan,       // E
    form.nama_lengkap,      // F
    form.gelar_belakang,    // G
    namaFull,               // H
    form.niy,               // I
    form.tmp_lahir,         // J
    form.tgl_lahir,         // K
    "'" + form.nik,         // L
    form.lp,                // M
    form.agama,             // N
    form.pendidikan,        // O
    form.jurusan,           // P
    form.thn_lulus,         // Q
    form.alamat,            // R
    "'" + form.hp,          // S
    form.status_peg,        // T
    form.jabatan,           // U
    form.tmt_jabatan,       // V
    form.inpassing,         // W
    form.tmt_inpassing,     // X
    "'" + form.nuptk,       // Y
    form.serdik,            // Z
    form.dapodik,           // AA
    form.tugtam             // AB
  ]];
  
  range.setValues(updateValues);

  // Update Info Diedit (AE & AF)
  sheet.getRange(rowIdx, 31).setValue(timestamp);      
  sheet.getRange(rowIdx, 32).setValue(form.user_login); 

  return "Sukses";
}

/**
 * 4. DELETE DATA PTK PAUD
 */
function deleteDataPTKPAUD(id, alasan, userLogin) {
  var ss = SpreadsheetApp.openById(ID_SPREADSHEET_PAUD);
  var sheet = ss.getSheetByName("Master Data GTK PAUD");
  var sheetArsip = ss.getSheetByName("gtk_non_aktif"); 
  
  if (!sheetArsip) {
    sheetArsip = ss.insertSheet("gtk_non_aktif");
    sheetArsip.appendRow(["ID", "Nama", "Unit", "Alasan Hapus", "User Hapus", "Waktu Hapus", "Data Asli JSON"]);
  }

  var data = sheet.getDataRange().getValues();
  var rowIdx = -1;
  var rowData = [];

  for (var i = 0; i < data.length; i++) {
    if (data[i][0] == id) {
      rowIdx = i + 1;
      rowData = data[i];
      break;
    }
  }

  if (rowIdx == -1) return "Error: Data tidak ditemukan.";

  var now = new Date();
  var timestamp = Utilities.formatDate(now, "Asia/Jakarta", "dd/MM/yyyy HH:mm:ss");

  sheetArsip.appendRow([
    rowData[0], 
    rowData[7], // Nama Full (Col H)
    rowData[2], // Unit
    alasan,
    userLogin,
    timestamp,
    JSON.stringify(rowData) 
  ]);

  sheet.deleteRow(rowIdx);

  return "Sukses";
}

// =============================================================
// HELPER: AMBIL JENJANG DARI DATABASE SEKOLAH (VALIDASI)
// =============================================================

function getJenjangByNPSN(npsn) {
  // PENTING: Ganti ID ini dengan ID Spreadsheet dimana sheet "Database Sekolah" berada
  var id = "1t0-Lmy0YD_GxHzimFWJGh5R5x6RhGL13uqKeVwWoCYE"; 
  
  try {
    var ss = SpreadsheetApp.openById(id);
    var sheet = ss.getSheetByName("Database Sekolah");
    if (!sheet) return "Sheet Tidak Ditemukan"; 

    var lastRow = sheet.getLastRow();
    // Ambil semua data (A:C) biar aman
    var data = sheet.getRange(2, 1, lastRow - 1, 3).getDisplayValues();
    
    var searchNpsn = String(npsn).trim();

    for (var i = 0; i < data.length; i++) {
      var rowNpsn = String(data[i][0]).trim(); // Kolom A: NPSN
      var rowJenjang = String(data[i][1]).trim(); // Kolom B: Jenjang
      
      if (rowNpsn === searchNpsn) {
        return rowJenjang; // KETEMU! Kembalikan Jenjang (TK/KB/SPS/TPA)
      }
    }
    return ""; // Tidak ketemu di list
  } catch (e) {
    return ""; // Error Spreadsheet
  }
}

// ==========================================
// DATA KEADAAN GTK PAUD
// ID: 1XetGkBymmN2NZQlXpzZ2MQyG0nhhZ0sXEPcNsLffhEU
// ==========================================

function getDataKeadaanGTKPAUD() {
  var id = "1XetGkBymmN2NZQlXpzZ2MQyG0nhhZ0sXEPcNsLffhEU"; 
  var ss = SpreadsheetApp.openById(id);
  var sheet = ss.getSheetByName("Keadaan GTK PAUD");
  if (!sheet) return [];
  
  var lastRow = sheet.getLastRow();
  if (lastRow < 3) return []; // Header 2 baris
  
  // Ambil Range A3:AB (A=1, AB=28)
  var data = sheet.getRange(3, 1, lastRow - 2, 28).getDisplayValues();
  return data;
}