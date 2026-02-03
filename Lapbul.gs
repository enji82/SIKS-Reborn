/* ======================================================================
   LAPBUL.GS - BACKEND UTAMA
   Update: Mapping SD Fixed Column & Auto-Sum Rombel
   ====================================================================== */

/* ======================================================================
   BAGIAN 1: GET DATA TABEL KELOLA (UPDATE MAPPING METADATA SD)
   ====================================================================== */
function getLapbulKelolaData(filterJenjang, filterBulan, filterTahun, filterStatus, keyword) {
  var result = [];
  
  if (typeof SPREADSHEET_IDS === 'undefined') return [];

  var reqJenjang = String(filterJenjang || "").toUpperCase().trim();
  var reqBulan = String(filterBulan || "").toLowerCase().trim();
  var reqTahun = String(filterTahun || "").toLowerCase().trim();
  
  // --- FUNGSI PENGAMBIL DATA ---
  var fetchDataFromSource = function(spreadsheetId, sheetName, sourceLabel) {
      var sourceResult = [];
      try {
          var ss = SpreadsheetApp.openById(spreadsheetId);
          var sheet = ss.getSheetByName(sheetName);
          if (!sheet) return [];

          var data = sheet.getDataRange().getValues();
          if (data.length < 2) return [];

          var headers = data[0];

          // FINDER
          var findIdx = function(keywords) {
              if (!Array.isArray(keywords)) keywords = [keywords];
              // Level 1: Persis
              for (var h = 0; h < headers.length; h++) {
                  var head = String(headers[h]).toLowerCase().trim();
                  for (var k = 0; k < keywords.length; k++) {
                      if (head === keywords[k].toLowerCase()) return h;
                  }
              }
              // Level 2: Mirip
              for (var h = 0; h < headers.length; h++) {
                  var head = String(headers[h]).toLowerCase().trim();
                  for (var k = 0; k < keywords.length; k++) {
                      if (head.includes(keywords[k].toLowerCase())) return h;
                  }
              }
              return -1;
          };

          // MAPPING INDEX (UPDATE DI SINI)
          var idx = {
             npsn: findIdx(["npsn", "nomor pokok"]),
             nama: findIdx(["nama sekolah", "nama_sekolah", "nama lembaga", "lembaga"]),
             jenjang: findIdx(["jenjang", "bentuk"]),
             statusSekolah: findIdx(["status sekolah", "status lembaga"]), 
             bulan: findIdx(["bulan"]),
             tahun: findIdx(["tahun"]),
             rombel: findIdx(["rombel", "jumlah rombel", "kelompok"]),
             status: findIdx(["status data", "status_data", "status laporan"]), 
             userKirim: findIdx(["user kirim", "pengirim", "email"]),
             
             // --- UPDATE DISINI: TAMBAHKAN "TANGGAL UNGGAH" ---
             tglKirim: findIdx(["tanggal unggah", "tgl unggah", "unggah", "timestamp", "waktu", "tanggal kirim", "tgl kirim"]), 
             // -------------------------------------------------
             
             tglEdit: findIdx(["tanggal edit", "tgl edit", "update"]),
             userEdit: findIdx(["user edit", "penyunting"]),
             file: findIdx(["dokumen", "file", "link"])
          };

          // HELPER TIMESTAMP
          var parseTS = function(v) {
               if(v instanceof Date) return v.getTime();
               if(typeof v === 'string') {
                   try { v=v.replace(/'/g,""); var p=v.split(/[\s/:]/); 
                   if(p.length>=3) return new Date(p[2],p[1]-1,p[0],p[3]||0,p[4]||0).getTime(); } catch(e){}
               }
               return 0;
          };

          // LOOPING BARIS
          for (var i = 1; i < data.length; i++) {
              var row = data[i];
              
              // Filter
              var valBulan = String(row[idx.bulan] || "").toLowerCase().trim();
              var valTahun = String(row[idx.tahun] || "").toLowerCase().trim();
              if (reqBulan && valBulan !== reqBulan) continue;
              if (reqTahun && valTahun !== reqTahun) continue;

              // Filter Jenjang Spesifik
              var valJenjang = (idx.jenjang > -1) ? String(row[idx.jenjang]).toUpperCase() : sourceLabel;
              if (reqJenjang && reqJenjang !== "SD" && reqJenjang !== "PAUD" && !valJenjang.includes(reqJenjang)) {
                  if (!reqJenjang.includes("PAUD")) continue; 
              }

              // Tanggal Logic
              var tglKirimVal = (idx.tglKirim > -1) ? row[idx.tglKirim] : "";
              var tglEditVal = (idx.tglEdit > -1) ? row[idx.tglEdit] : "";
              
              var tsKirim = parseTS(tglKirimVal);
              var tsEdit = parseTS(tglEditVal);
              var sortTime = (tsEdit > 0) ? tsEdit : tsKirim;
              if (sortTime === 0) sortTime = i;

              // Format Tampilan Tanggal
              var txtTglKirim = "-";
              if (tglKirimVal instanceof Date) {
                  txtTglKirim = Utilities.formatDate(tglKirimVal, Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm");
              } else if (tglKirimVal && String(tglKirimVal).length > 5) {
                  txtTglKirim = String(tglKirimVal).replace(/'/g,""); // Jika text, tampilkan apa adanya
              }

              var txtTglEdit = (tsEdit > 0) ? Utilities.formatDate(new Date(tsEdit), Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm") : "-";

              sourceResult.push({
                  rowId: (i + 1).toString(),
                  namaSekolah: (idx.nama > -1) ? row[idx.nama] : "Tanpa Nama",
                  npsn: (idx.npsn > -1) ? row[idx.npsn] : "-",
                  bulan: (idx.bulan > -1) ? row[idx.bulan] : "-",
                  tahun: (idx.tahun > -1) ? row[idx.tahun] : "-",
                  statusData: (idx.status > -1) ? row[idx.status] : "Diproses",
                  jenjang: (idx.jenjang > -1 && row[idx.jenjang]) ? row[idx.jenjang] : sourceLabel,
                  rombel: (idx.rombel > -1) ? row[idx.rombel] : "0",
                  statusSekolah: (idx.statusSekolah > -1) ? row[idx.statusSekolah] : "-", 
                  fileUrl: (idx.file > -1) ? row[idx.file] : "", 
                  
                  userKirim: (idx.userKirim > -1) ? row[idx.userKirim] : "-",
                  tglKirim: txtTglKirim, // <--- Ini yang diperbaiki
                  
                  userEdit: (idx.userEdit > -1 && row[idx.userEdit]) ? String(row[idx.userEdit]).replace(/'/g,"") : "-",
                  tglEdit: txtTglEdit,
                  
                  verifikator: "-", tglVerif: "-", 
                  sortTime: sortTime,
                  source: sourceLabel 
              });
          }
      } catch (e) {
          Logger.log("Error reading " + sourceLabel + ": " + e.toString());
      }
      return sourceResult;
  };

  // EXECUTE
  if (reqJenjang === "" || reqJenjang.includes("SD")) {
      result = result.concat(fetchDataFromSource(SPREADSHEET_IDS.SD_DATA, "Input SD", "SD"));
  }
  if (reqJenjang === "" || reqJenjang.includes("PAUD") || reqJenjang.includes("TK") || reqJenjang.includes("KB")) {
      result = result.concat(fetchDataFromSource(SPREADSHEET_IDS.PAUD_DATA, "Input PAUD", "PAUD"));
  }

  return result;
}

function parseDateStr(str) {
  if(!str) return new Date(0);
  try {
    var parts = str.split(' '); var d = parts[0].split('/'); var t = (parts[1]||"00:00:00").split(':');
    return new Date(d[2], d[1]-1, d[0], t[0], t[1], t[2]);
  } catch(e) { return new Date(0); }
}

// --- 2. MASTER DATA (AUTO FILL) ---
function getSekolahByNPSN(npsn) {
  try {
    const ss = SpreadsheetApp.openById("1wiDKez4rL5UYnpP2-OZjYowvmt1nRx-fIMy9trJlhBA");
    const sheet = ss.getSheetByName("Data_Sekolah");
    const data = sheet.getDataRange().getValues();
    
    for (var i = 1; i < data.length; i++) {
      // Kolom A: NPSN, B: Jenjang, C: Nama, D: Status
      if (String(data[i][0]).trim() === String(npsn).trim()) {
        return {
          found: true,
          npsn: data[i][0],
          jenjang: data[i][1],
          nama_sekolah: data[i][2],
          status_sekolah: data[i][3]
        };
      }
    }
    return { found: false };
  } catch (e) { return { error: e.toString() }; }
}

// --- 3. SIMPAN DATA SD (AUTO SUM ROMBEL) ---
function simpanLapbulSD_Complex(form, fileData) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_IDS.SD_DATA);
    const sheet = ss.getSheetByName("Input SD");
    
    // Validasi Ukuran 200KB
    if (fileData && fileData.data.length > 280000) {
         return { success: false, message: "Ukuran file terlalu besar (Max 200KB)!" };
    }

    var fileUrl = "";
    if (fileData && fileData.data) {
       var folderId = "1I8DRQYpBbTt1mJwtD1WXVD6UK51TC8El"; 
       var fileName = "Laporan Bulan - " + form.nama_sekolah + " - " + form.bulan + " - " + form.tahun;
       fileUrl = uploadFileToDrive(fileData, folderId, fileName);
    }

    var timestamp = new Date();
    
    // Helper ambil value angka
    const val = (key) => parseInt(form[key]) || 0;

    // --- LOGIKA HITUNG TOTAL ROMBEL (AUTOMATIC) ---
    // Dijumlahkan dari Rombel 1 (Col I) + Rombel 2 (Col AD) + ... + Rombel 6 (Col DJ)
    // Sesuai input name di form: rombel_1, rombel_2, ...
    var totalRombelHitung = val('rombel_1') + val('rombel_2') + val('rombel_3') + val('rombel_4') + val('rombel_5') + val('rombel_6');

    // MAPPING DATA KE ARRAY (A - HP)
    var rowData = [
      timestamp,          // A: Tanggal Kirim
      form.bulan,         // B: Bulan
      form.tahun,         // C: Tahun
      form.npsn,          // D: NPSN
      form.status_sekolah,// E: Status Sekolah
      totalRombelHitung,  // F: Rombel (HASIL PENJUMLAHAN OTOMATIS)
      form.jenjang,       // G: Jenjang
      form.nama_sekolah,  // H: Nama Sekolah
      
      // --- KELAS 1 (I - AC) ---
      val('rombel_1'),    // I: Rombel 1
      val('k1_l'), val('k1_p'), 
      val('k1a_l'), val('k1a_p'), val('k1b_l'), val('k1b_p'), val('k1c_l'), val('k1c_p'), 
      val('k1_islam_l'), val('k1_islam_p'), val('k1_kristen_l'), val('k1_kristen_p'), 
      val('k1_katolik_l'), val('k1_katolik_p'), val('k1_hindu_l'), val('k1_hindu_p'), 
      val('k1_buddha_l'), val('k1_buddha_p'), val('k1_konghucu_l'), val('k1_konghucu_p'),

      // --- KELAS 2 (AD - AX) ---
      val('rombel_2'),    // AD: Rombel 2
      val('k2_l'), val('k2_p'), 
      val('k2a_l'), val('k2a_p'), val('k2b_l'), val('k2b_p'), val('k2c_l'), val('k2c_p'),
      val('k2_islam_l'), val('k2_islam_p'), val('k2_kristen_l'), val('k2_kristen_p'), 
      val('k2_katolik_l'), val('k2_katolik_p'), val('k2_hindu_l'), val('k2_hindu_p'), 
      val('k2_buddha_l'), val('k2_buddha_p'), val('k2_konghucu_l'), val('k2_konghucu_p'),

      // --- KELAS 3 (AY - BS) ---
      val('rombel_3'),    // AY: Rombel 3
      val('k3_l'), val('k3_p'), 
      val('k3a_l'), val('k3a_p'), val('k3b_l'), val('k3b_p'), val('k3c_l'), val('k3c_p'),
      val('k3_islam_l'), val('k3_islam_p'), val('k3_kristen_l'), val('k3_kristen_p'), 
      val('k3_katolik_l'), val('k3_katolik_p'), val('k3_hindu_l'), val('k3_hindu_p'), 
      val('k3_buddha_l'), val('k3_buddha_p'), val('k3_konghucu_l'), val('k3_konghucu_p'),

      // --- KELAS 4 (BT - CN) ---
      val('rombel_4'),    // BT: Rombel 4
      val('k4_l'), val('k4_p'), 
      val('k4a_l'), val('k4a_p'), val('k4b_l'), val('k4b_p'), val('k4c_l'), val('k4c_p'),
      val('k4_islam_l'), val('k4_islam_p'), val('k4_kristen_l'), val('k4_kristen_p'), 
      val('k4_katolik_l'), val('k4_katolik_p'), val('k4_hindu_l'), val('k4_hindu_p'), 
      val('k4_buddha_l'), val('k4_buddha_p'), val('k4_konghucu_l'), val('k4_konghucu_p'),

      // --- KELAS 5 (CO - DI) ---
      val('rombel_5'),    // CO: Rombel 5
      val('k5_l'), val('k5_p'), 
      val('k5a_l'), val('k5a_p'), val('k5b_l'), val('k5b_p'), val('k5c_l'), val('k5c_p'),
      val('k5_islam_l'), val('k5_islam_p'), val('k5_kristen_l'), val('k5_kristen_p'), 
      val('k5_katolik_l'), val('k5_katolik_p'), val('k5_hindu_l'), val('k5_hindu_p'), 
      val('k5_buddha_l'), val('k5_buddha_p'), val('k5_konghucu_l'), val('k5_konghucu_p'),

      // --- KELAS 6 (DJ - ED) ---
      val('rombel_6'),    // DJ: Rombel 6
      val('k6_l'), val('k6_p'), 
      val('k6a_l'), val('k6a_p'), val('k6b_l'), val('k6b_p'), val('k6c_l'), val('k6c_p'),
      val('k6_islam_l'), val('k6_islam_p'), val('k6_kristen_l'), val('k6_kristen_p'), 
      val('k6_katolik_l'), val('k6_katolik_p'), val('k6_hindu_l'), val('k6_hindu_p'), 
      val('k6_buddha_l'), val('k6_buddha_p'), val('k6_konghucu_l'), val('k6_konghucu_p'),

      // --- PTK (EE - HI) ---
      val('ks_pns'), val('ks_pppk'), val('ks_nonasn'), 
      
      val('gk_pns'), val('gk_pppk'), val('gk_pppk_pw'), val('gk_gty'), val('gk_gtt'), 
      val('gpai_pns'), val('gpai_pppk'), val('gpai_pppk_pw'), val('gpai_gty'), val('gpai_gtt'), 
      val('gpjok_pns'), val('gpjok_pppk'), val('gpjok_pppk_pw'), val('gpjok_gty'), val('gpjok_gtt'), 
      val('gkris_pns'), val('gkris_pppk'), val('gkris_pppk_pw'), val('gkris_gty'), val('gkris_gtt'), 
      val('gkat_pns'), val('gkat_pppk'), val('gkat_pppk_pw'), val('gkat_gty'), val('gkat_gtt'), 
      
      val('gbing_pns'), val('gbing_pppk'), val('gbing_pppk_pw'), val('gbing_gty'), val('gbing_gtt'), 
      val('gmap_pns'), val('gmap_pppk'), val('gmap_pppk_pw'), val('gmap_gty'), val('gmap_gtt'), 
      
      val('puo_pns'), val('puo_pppk'), val('puo_pppk_pw'), val('puo_pty'), val('puo_ptt'), 
      val('olo_pns'), val('olo_pppk'), val('olo_pppk_pw'), val('olo_pty'), val('olo_ptt'), 
      val('plo_pns'), val('plo_pppk'), val('plo_pppk_pw'), val('plo_pty'), val('plo_ptt'), 
      val('ptlo_pns'), val('ptlo_pppk'), val('ptlo_pppk_pw'), val('ptlo_pty'), val('ptlo_ptt'), 
      val('adm_pns'), val('adm_pppk'), val('adm_pppk_pw'), val('adm_pty'), val('adm_ptt'), 
      val('pjg_pns'), val('pjg_pppk'), val('pjg_pppk_pw'), val('pjg_pty'), val('pjg_ptt'), 
      val('tas_pns'), val('tas_pppk'), val('tas_pppk_pw'), val('tas_pty'), val('tas_ptt'), 
      val('pust_pns'), val('pust_pppk'), val('pust_pppk_pw'), val('pust_pty'), val('pust_ptt'), 
      val('lain_pns'), val('lain_pppk'), val('lain_pppk_pw'), val('lain_pty'), val('lain_ptt'), 

      // --- METADATA (HJ - HP) ---
      fileUrl,          // HJ: Dokumen (Index 217)
      "Diproses",       // HK: Status (Index 218)
      form.user_login,  // HL: User Kirim (Index 219)
      "",               // HM: Tanggal Edit (Index 220)
      "",               // HN: User Edit (Index 221)
      "",               // HO: Tanggal Verif (Index 222)
      ""                // HP: Admin Verif (Index 223)
    ];

    sheet.appendRow(rowData);
    updateDataSekolahMaster(form);

    return { success: true, message: "Laporan Disimpan & Data Sekolah Diperbarui" };

  } catch (e) { return { success: false, message: "Gagal Simpan SD: " + e.toString() }; }
}

function updateDataSekolahMaster(form) {
  try {
    const ss = SpreadsheetApp.openById("1wiDKez4rL5UYnpP2-OZjYowvmt1nRx-fIMy9trJlhBA");
    const sheet = ss.getSheetByName("Data_Sekolah");
    const data = sheet.getDataRange().getValues();
    
    // Cari Baris Berdasarkan NPSN (Kolom A / Index 0)
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][0]).trim() === String(form.npsn).trim()) {
        // Ketemu! Update Kolom E (Index 4) s.d P (Index 15)
        // Ingat: getRange pakai index 1-based. Baris = i+1.
        // Kolom E adalah kolom ke-5.
        
        var updateRange = sheet.getRange(i + 1, 5, 1, 12); // Mulai col 5, update 12 kolom
        var updateValues = [[
          form.yayasan,           // E
          form.no_sk_pendirian,   // F
          form.tgl_pendirian,     // G
          form.no_sk_ijin,        // H
          form.tgl_ijin,          // I
          form.akreditasi,        // J
          form.skor,              // K
          form.no_sertifikat,     // L
          form.tgl_sertifikat,    // M
          form.alamat,            // N
          form.telepon,           // O
          form.email              // P
        ]];
        
        updateRange.setValues(updateValues);
        break; // Stop loop setelah ketemu
      }
    }
  } catch(e) {
    // Silent fail (jangan gagalkan laporan jika update master error, atau bisa di log)
    Logger.log("Update Master Error: " + e.toString());
  }
}

// --- 4. SIMPAN DATA PAUD (MAPPING FIXED) ---
function simpanLapbulPAUD(form, fileData) {
  try {
    const ss = SpreadsheetApp.openById("1an0oQQPdMh6wrUJIAzTGYk3DKFvYprK5SU7RmRXjIgs"); // ID Sesuai Request
    const sheet = ss.getSheetByName("Input PAUD");
    
    // 1. Upload File
    if (fileData && fileData.data.length > 280000) {
         return { success: false, message: "Ukuran file terlalu besar (Max 200KB)!" };
    }
    
    // Tentukan Folder berdasarkan Jenjang (Opsional, bisa disatukan)
    var folderId = "18CxRT-eledBGRtHW1lFd2AZ8Bub6q5ra"; // Default Folder PAUD
    
    var fileUrl = "";
    if (fileData && fileData.data) {
       var fileName = "Laporan Bulan - " + form.nama_sekolah + " - " + form.bulan + " - " + form.tahun;
       fileUrl = uploadFileToDrive(fileData, folderId, fileName);
    }

    var timestamp = new Date();
    const val = (key) => parseInt(form[key]) || 0;

    // 2. Mapping Data ke Kolom A - AW (Index 0 - 48)
    // A=0, B=1, ... AW=48
    var rowData = [
      timestamp,          // A: Tanggal Kirim
      form.bulan,         // B: Bulan
      form.tahun,         // C: Tahun
      form.npsn,          // D: NPSN
      form.status_sekolah,// E: Status Sekolah
      val('jumlah_rombel'),// F: Rombel
      form.jenjang,       // G: Jenjang
      form.nama_sekolah,  // H: Nama Sekolah
      
      // --- MENURUT USIA (I - V) ---
      val('u01_l'), val('u01_p'), // I, J
      val('u12_l'), val('u12_p'), // K, L
      val('u23_l'), val('u23_p'), // M, N
      val('u34_l'), val('u34_p'), // O, P
      val('u45_l'), val('u45_p'), // Q, R
      val('u56_l'), val('u56_p'), // S, T
      val('u6_l'), val('u6_p'),   // U, V

      // --- MENURUT ROMBEL (TK) (W - Z) ---
      val('kel_a_l'), val('kel_a_p'), // W, X
      val('kel_b_l'), val('kel_b_p'), // Y, Z

      // --- KEPALA SEKOLAH (AA - AD) ---
      val('ks_gty'), val('ks_gtt'), val('ks_pns'), val('ks_pppk'),

      // --- GURU KELAS (AE - AH) ---
      val('gk_gty'), val('gk_gtt'), val('gk_pns'), val('gk_pppk'),

      // --- GURU PENDAMPING (AI - AL) ---
      val('gp_gty'), val('gp_gtt'), val('gp_pns'), val('gp_pppk'),

      // --- TENDIK (AM - AP) ---
      val('td_penjaga'), // AM
      val('td_adm'),     // AN
      val('td_perpus'),  // AO
      val('td_lain'),    // AP

      // --- METADATA (AQ - AW) ---
      fileUrl,           // AQ: Dokumen
      form.user_login,   // AR: User Kirim
      "",                // AS: Update (Tgl Edit)
      "",                // AT: User Edit
      "",                // AU: Verif
      "",                // AV: Admin
      "Diproses"         // AW: Status
    ];

    sheet.appendRow(rowData);

    // 3. Update Master Data (Fungsi Global yang sudah dibuat sebelumnya)
    updateDataSekolahMaster(form);

    return { success: true, message: "Laporan PAUD Berhasil Disimpan" };

  } catch (e) { return { success: false, message: "Gagal Simpan PAUD: " + e.toString() }; }
}

function uploadFileToDrive(fileData, folderId, fileName) {
  try {
    var folder = DriveApp.getFolderById(folderId);
    var blob = Utilities.newBlob(Utilities.base64Decode(fileData.data), fileData.mimeType, fileName);
    var file = folder.createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    return file.getUrl();
  } catch (e) { return "Error Upload"; }
}

/* ======================================================================
   TAMBAHAN: SEARCH ENGINE DATA SEKOLAH
   ====================================================================== */
function getAllSchoolsList() {
  try {
    const ss = SpreadsheetApp.openById("1wiDKez4rL5UYnpP2-OZjYowvmt1nRx-fIMy9trJlhBA");
    const sheet = ss.getSheetByName("Data_Sekolah");
    
    // Ambil Data dari Baris 2, Kolom 1 (A) sampai Kolom 16 (P)
    // A=0, B=1 ... P=15
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return [];

    const data = sheet.getRange(2, 1, lastRow - 1, 16).getValues();
    
    // Helper Format Tanggal untuk HTML input type="date" (YYYY-MM-DD)
    const fmtDate = (d) => {
      if (!d || !(d instanceof Date)) return "";
      return Utilities.formatDate(d, Session.getScriptTimeZone(), "yyyy-MM-dd");
    };

    return data.map(function(r) {
      return {
        // Data Tetap (A-D)
        npsn: String(r[0]).trim(),
        jenjang: r[1],
        nama: r[2],
        status: r[3],
        
        // Data Dinamis (E-P)
        yayasan: r[4],
        no_sk_pendirian: r[5],
        tgl_pendirian: fmtDate(r[6]),
        no_sk_ijin: r[7],
        tgl_ijin: fmtDate(r[8]),
        akreditasi: r[9],
        skor: r[10],
        no_sertifikat: r[11],
        tgl_sertifikat: fmtDate(r[12]),
        alamat: r[13],
        telepon: r[14],
        email: r[15],

        // Kunci Pencarian
        search_key: (String(r[0]) + " " + String(r[2])).toLowerCase()
      };
    });
  } catch (e) {
    return [];
  }
}

/* ======================================================================
   FITUR EDIT: AMBIL DATA DETAIL (SMART MAPPING)
   ====================================================================== */
function getDetailRowSD(rowId) {
  var result = {};
  
  try {
    // ==========================================
    // 1. AMBIL DATA LAPORAN (DARI SPREADSHEET 'INPUT SD')
    // ==========================================
    var idBaris = parseInt(rowId);
    
    // Buka Spreadsheet Laporan
    var ssLaporan = SpreadsheetApp.openById(SPREADSHEET_IDS.SD_DATA);
    var sheetInput = ssLaporan.getSheetByName("Input SD");
    
    if (!sheetInput) return { error: "Sheet 'Input SD' tidak ditemukan!" };

    var lastCol = sheetInput.getLastColumn();
    var headersInput = sheetInput.getRange(1, 1, 1, lastCol).getValues()[0];
    var dataInput = sheetInput.getRange(idBaris, 1, 1, lastCol).getValues()[0];
    
    // Mapping Data Laporan
    for (var i = 0; i < headersInput.length; i++) {
       var key = String(headersInput[i]).trim();
       var val = dataInput[i];
       if (val instanceof Date) {
         try { val = Utilities.formatDate(val, Session.getScriptTimeZone(), "yyyy-MM-dd"); } catch(e) {}
       }
       result[key] = val;
    }
    result.ROW_ID = idBaris;

    // Ambil NPSN target
    var targetNPSN = String(result["NPSN"] || result["npsn"] || "").trim();
    
    if (!targetNPSN) {
        Logger.log("NPSN Kosong di Data Laporan Baris " + rowId);
        return result; 
    }

    // ==========================================
    // 2. AMBIL DATA MASTER (DARI SPREADSHEET 'DATA SEKOLAH')
    // ==========================================
    
    // ID Spreadsheet Master (Sesuai info Bapak)
    var ID_MASTER = "1wiDKez4rL5UYnpP2-OZjYowvmt1nRx-fIMy9trJlhBA";
    
    try {
        // Buka Spreadsheet Master
        var ssMaster = SpreadsheetApp.openById(ID_MASTER);
        var sheetMaster = ssMaster.getSheetByName("Data_Sekolah"); 
        
        if (sheetMaster) {
            var dataMaster = sheetMaster.getDataRange().getValues();
            if (dataMaster.length > 0) {
                var headMaster = dataMaster[0]; // Header
                
                // A. Cari Posisi Kolom NPSN di Master
                var colIndexNPSN = -1;
                for (var c = 0; c < headMaster.length; c++) {
                    if (String(headMaster[c]).toLowerCase().trim() === "npsn") {
                        colIndexNPSN = c;
                        break;
                    }
                }

                // B. Jika Kolom NPSN Ketemu, Cari Barisnya
                if (colIndexNPSN > -1) {
                    var dataKetemu = false;
                    for (var m = 1; m < dataMaster.length; m++) {
                        var rowM = dataMaster[m];
                        var npsnMaster = String(rowM[colIndexNPSN] || "").trim();
                        
                        // Cek Kesamaan NPSN
                        if (npsnMaster === targetNPSN) {
                            dataKetemu = true;
                            
                            // C. Gabungkan Data Master ke Result
                            for (var h = 0; h < headMaster.length; h++) {
                                var keyM = String(headMaster[h]).trim();
                                var valM = rowM[h];
                                
                                // Masukkan data master HANYA JIKA di result belum ada / kosong
                                // (Supaya data laporan utama tidak tertimpa)
                                if (result[keyM] === undefined || result[keyM] === "") {
                                    if (valM instanceof Date) {
                                       try { valM = Utilities.formatDate(valM, Session.getScriptTimeZone(), "yyyy-MM-dd"); } catch(e) {}
                                    }
                                    result[keyM] = valM;
                                }
                            }
                            break; // Stop loop
                        }
                    }
                    if (!dataKetemu) Logger.log("NPSN " + targetNPSN + " tidak ditemukan di Master.");
                } else {
                    Logger.log("Error: Tidak ada kolom 'NPSN' di Master Data_Sekolah");
                }
            }
        } else {
            Logger.log("Error: Sheet 'Data_Sekolah' tidak ditemukan di ID Spreadsheet Master.");
        }
    } catch (errMaster) {
        Logger.log("Gagal membuka Spreadsheet Master: " + errMaster.toString());
    }
    
    return result;

  } catch (e) {
    Logger.log("CRITICAL ERROR: " + e.toString());
    return { error: "Error Backend: " + e.toString() };
  }
}

/* ======================================================================
   FITUR EDIT: UPDATE DATA (TIMPA BARIS LAMA)
   ====================================================================== */
function updateLapbulSD(form, fileData) {
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_IDS.SD_DATA);
    var sheet = ss.getSheetByName("Input SD");
    var rowId = parseInt(form.EDIT_ROW_ID);
    
    // ... (KODE UPLOAD FILE TETAP SAMA) ...
    var fileUrl = form.file_url_lama || ""; 
    if (fileData && fileData.data) {
       var folderId = "1I8DRQYpBbTt1mJwtD1WXVD6UK51TC8El"; 
       var fileName = "Laporan SD - " + (form.nama_sekolah||"SD") + " - " + form.bulan + " " + form.tahun + " (Revisi)";
       fileUrl = uploadFileToDrive(fileData, folderId, fileName);
    }

    // SIAPKAN DATA BARU
    var lastCol = sheet.getLastColumn();
    var headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
    var currentRowData = sheet.getRange(rowId, 1, 1, lastCol).getValues()[0];
    var newRowData = [];

    // --- BUAT DATA WAKTU SEKARANG ---
    var now = new Date();
    var strTglEdit = Utilities.formatDate(now, Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm:ss");
    var strUserEdit = form.user_login || "Admin";
    // --------------------------------

    for (var i = 0; i < headers.length; i++) {
        var rawHeader = String(headers[i]).toLowerCase().trim();
        var keyForm = rawHeader.replace(/\s+/g, '_'); 
        
        if (rawHeader.includes("tgl edit") || rawHeader.includes("tanggal edit") || rawHeader.includes("update")) {
            newRowData.push("'" + strTglEdit); // SIMPAN YANG BARU
        } 
        else if (rawHeader.includes("user edit") || rawHeader.includes("penyunting")) {
            newRowData.push("'" + strUserEdit); 
        }
        else if (rawHeader.includes("status data") || rawHeader === "status") {
            newRowData.push("Diproses"); 
        }
        else if (rawHeader.includes("dokumen") || rawHeader.includes("file")) {
            newRowData.push(fileUrl);
        }
        else if (form[keyForm] !== undefined) {
             var val = form[keyForm];
             if (rawHeader.includes("tgl") || rawHeader.includes("tanggal")) {
                 newRowData.push("'" + val); 
             } else {
                 newRowData.push(val);
             }
        }
        else {
             newRowData.push(currentRowData[i]);
        }
    }

    // SIMPAN & FLUSH
    sheet.getRange(rowId, 1, 1, newRowData.length).setValues([newRowData]);
    SpreadsheetApp.flush(); 
    
    // --- RETURN DATA LENGKAP KE FRONTEND ---
    return { 
        success: true, 
        message: "Data berhasil diperbarui!",
        // Kirim balik data ini agar frontend bisa update tabel TANPA reload
        newData: {
            tglEdit: Utilities.formatDate(now, Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm"),
            userEdit: strUserEdit,
            statusData: "Diproses",
            sortTime: now.getTime(), // Timestamp terbaru agar naik ke atas
            fileUrl: fileUrl
        }
    };

  } catch (e) { 
    return { success: false, message: "Gagal Update: " + e.toString() }; 
  }
}

// --- 1. AMBIL DETAIL PAUD ---
function getDetailRowPAUD(rowId) {
  // Gunakan logika yang sama dengan SD, tapi arahkan ke Sheet PAUD
  // Kita bisa reuse kode SD dengan sedikit modifikasi string, 
  // tapi agar aman dan tidak saling ganggu, kita buat fungsi terpisah.
  
  var result = {};
  try {
    var idBaris = parseInt(rowId);
    var ssLaporan = SpreadsheetApp.openById(SPREADSHEET_IDS.PAUD_DATA); // Pastikan ada ID ini di config
    var sheetInput = ssLaporan.getSheetByName("Input PAUD"); // <--- BEDA DISINI
    
    if (!sheetInput) return { error: "Sheet 'Input PAUD' tidak ditemukan!" };

    var lastCol = sheetInput.getLastColumn();
    var headersInput = sheetInput.getRange(1, 1, 1, lastCol).getValues()[0];
    var dataInput = sheetInput.getRange(idBaris, 1, 1, lastCol).getValues()[0];
    
    // Mapping Data Laporan
    for (var i = 0; i < headersInput.length; i++) {
       var key = String(headersInput[i]).trim();
       var val = dataInput[i];
       if (val instanceof Date) {
         try { val = Utilities.formatDate(val, Session.getScriptTimeZone(), "yyyy-MM-dd"); } catch(e) {}
       }
       result[key] = val;
    }
    result.ROW_ID = idBaris;

    // AMBIL DATA MASTER (SAMA SEPERTI SD)
    var targetNPSN = String(result["NPSN"] || result["npsn"] || "").trim();
    var ID_MASTER = "1wiDKez4rL5UYnpP2-OZjYowvmt1nRx-fIMy9trJlhBA"; // ID MASTER SEKOLAH
    
    if (targetNPSN) {
        try {
            var ssMaster = SpreadsheetApp.openById(ID_MASTER);
            var sheetMaster = ssMaster.getSheetByName("Data_Sekolah"); 
            if (sheetMaster) {
                var dataMaster = sheetMaster.getDataRange().getValues();
                var headMaster = dataMaster[0];
                var colIndexNPSN = -1;
                
                // Cari Kolom NPSN
                for (var c = 0; c < headMaster.length; c++) {
                    if (String(headMaster[c]).toLowerCase().trim() === "npsn") { colIndexNPSN = c; break; }
                }

                if (colIndexNPSN > -1) {
                    for (var m = 1; m < dataMaster.length; m++) {
                        var rowM = dataMaster[m];
                        if (String(rowM[colIndexNPSN] || "").trim() === targetNPSN) {
                            // Merge Data
                            for (var h = 0; h < headMaster.length; h++) {
                                var keyM = String(headMaster[h]).trim();
                                var valM = rowM[h];
                                if (result[keyM] === undefined || result[keyM] === "") {
                                    if (valM instanceof Date) {
                                       try { valM = Utilities.formatDate(valM, Session.getScriptTimeZone(), "yyyy-MM-dd"); } catch(e) {}
                                    }
                                    result[keyM] = valM;
                                }
                            }
                            break;
                        }
                    }
                }
            }
        } catch (e) {}
    }
    return result;
  } catch (e) { return { error: "Backend PAUD Error: " + e.toString() }; }
}

// --- 2. UPDATE DATA PAUD ---
function updateLapbulPAUD(form, fileData) {
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_IDS.PAUD_DATA);
    var sheet = ss.getSheetByName("Input PAUD"); // <--- BEDA DISINI
    var rowId = parseInt(form.EDIT_ROW_ID);
    
    // Upload File
    var fileUrl = form.file_url_lama || ""; 
    if (fileData && fileData.data) {
       var folderId = "1I8DRQYpBbTt1mJwtD1WXVD6UK51TC8El"; // Samakan atau bedakan foldernya
       var fileName = "Laporan PAUD - " + (form.nama_sekolah||"PAUD") + " - " + form.bulan + " " + form.tahun + " (Revisi)";
       fileUrl = uploadFileToDrive(fileData, folderId, fileName);
    }

    // Persiapan Data
    var lastCol = sheet.getLastColumn();
    var headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
    var currentRowData = sheet.getRange(rowId, 1, 1, lastCol).getValues()[0];
    var newRowData = [];

    // Metadata Time
    var now = new Date();
    var strTglEdit = Utilities.formatDate(now, Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm:ss");
    var strUserEdit = form.user_login || "Admin";

    // Mapping Header (Logika "Smart Header" SD dicopy kesini)
    for (var i = 0; i < headers.length; i++) {
        var rawHeader = String(headers[i]).toLowerCase().trim();
        var keyForm = rawHeader.replace(/\s+/g, '_'); 
        
        if (rawHeader.includes("tgl edit") || rawHeader.includes("tanggal edit") || rawHeader.includes("update")) {
            newRowData.push("'" + strTglEdit);
        } 
        else if (rawHeader.includes("user edit") || rawHeader.includes("penyunting")) {
            newRowData.push("'" + strUserEdit); 
        }
        else if (rawHeader.includes("status data") || rawHeader === "status") {
            newRowData.push("Diproses"); 
        }
        else if (rawHeader.includes("dokumen") || rawHeader.includes("file")) {
            newRowData.push(fileUrl);
        }
        else if (form[keyForm] !== undefined) {
             var val = form[keyForm];
             if (rawHeader.includes("tgl") || rawHeader.includes("tanggal")) {
                 newRowData.push("'" + val); 
             } else {
                 newRowData.push(val);
             }
        }
        else {
             newRowData.push(currentRowData[i]);
        }
    }

    // Simpan & Flush
    sheet.getRange(rowId, 1, 1, newRowData.length).setValues([newRowData]);
    SpreadsheetApp.flush(); 
    
    // Return untuk Update Lokal
    return { 
        success: true, 
        message: "Data PAUD berhasil diperbarui!",
        newData: {
            tglEdit: Utilities.formatDate(now, Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm"),
            userEdit: strUserEdit,
            statusData: "Diproses",
            sortTime: now.getTime(),
            fileUrl: fileUrl
        }
    };
  } catch (e) { return { success: false, message: "Gagal Update PAUD: " + e.toString() }; }
}