/* ======================================================================
   LAPBUL.GS - VERSI ULTRA LIGHT (PARTIAL FETCH)
   Hanya mengambil potongan data terbawah agar loading super cepat.
   ====================================================================== */

function getLapbulKelolaData(filterJenjang, filterBulan, filterTahun, filterStatus, keyword) {
  var result = [];
  
  // KONFIGURASI LIMIT
  // Jika loading awal (tanpa search), cukup ambil 300 data terbaru per sheet.
  // Jika sedang search, kita perbesar jangkauan scan ke 2000 baris terakhir.
  var isSearching = (keyword && keyword.length > 2);
  var LIMIT_PER_SHEET = isSearching ? 2000 : 300; 

  var IDS = (typeof SPREADSHEET_IDS !== 'undefined') ? SPREADSHEET_IDS : {
      PAUD_DATA: "1an0oQQPdMh6wrUJIAzTGYk3DKFvYprK5SU7RmRXjIgs", 
      SD_DATA: "1u4tNL3uqt5xHITXYwHnytK6Kul9Siam-vNYuzmdZB4s"    
  };

  // Normalisasi Filter
  var reqJenjang = String(filterJenjang || "").toUpperCase().trim();
  var reqBulan = String(filterBulan || "").toLowerCase().trim();
  var reqTahun = String(filterTahun || "").toLowerCase().trim();
  var reqStatus = String(filterStatus || "").toLowerCase().trim();
  var reqKey = String(keyword || "").toLowerCase().trim();

  // Formatter Cepat (Native JS)
  var fastFormat = function(val) {
      if (!val || val === "" || val === "-") return "-";
      var d = (val instanceof Date) ? val : new Date(val);
      if (isNaN(d.getTime())) return "-";
      var pad = function(n) { return n < 10 ? '0' + n : n; };
      return pad(d.getDate()) + '/' + pad(d.getMonth() + 1) + '/' + d.getFullYear() + ' ' + 
             pad(d.getHours()) + ':' + pad(d.getMinutes()) + ':' + pad(d.getSeconds());
  };

  var fetchDataSmart = function(spreadsheetId, sheetName, sourceLabel) {
      var sourceResult = [];
      try {
          var ss = SpreadsheetApp.openById(spreadsheetId);
          var sheet = ss.getSheetByName(sheetName);
          if (!sheet) return [];

          var lastRow = sheet.getLastRow();
          if (lastRow < 2) return []; // Hanya header atau kosong

          // --- LANGKAH 1: AMBIL HEADER SAJA (Sangat Ringan) ---
          var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
          headers = headers.map(function(h) { return String(h).toLowerCase(); });

          // Mapping Index Kolom
          var idx = {
              nama: headers.indexOf("nama sekolah"),
              npsn: headers.indexOf("npsn"),
              bulan: headers.indexOf("bulan"),
              tahun: headers.indexOf("tahun"),
              jenjang: headers.indexOf("jenjang"),
              statusSekolah: headers.findIndex(h => h.includes("status sekolah") || h === "status"),
              rombel: headers.findIndex(function(h) { return h.includes("rombel") || h.includes("jml") || h.includes("total"); }),
              file: headers.findIndex(h => h.includes("file") || h.includes("dokumen"))
          };

          // Index Activity (Manual karena struktur fix)
          var col = (sourceLabel === 'PAUD') ? 
                    { tglKirim:0, userKirim:43, tglEdit:44, userEdit:45, tglVerif:46, userVerif:47, statusData:48, ket:49 } : 
                    { tglKirim:0, userKirim:219, tglEdit:220, userEdit:221, tglVerif:222, userVerif:223, statusData:218, ket:224 };

          // --- LANGKAH 2: HITUNG KOORDINAT POTONGAN DATA ---
          // Kita hanya ambil 'LIMIT_PER_SHEET' baris dari bawah
          var startRow = Math.max(2, lastRow - LIMIT_PER_SHEET + 1); 
          var numRows = (lastRow - startRow + 1);
          
          if (numRows < 1) return [];

          // --- LANGKAH 3: AMBIL DATA POTONGAN SAJA (Cepat) ---
          // Mengambil 300 baris jauh lebih cepat daripada 3000 baris
          var data = sheet.getRange(startRow, 1, numRows, sheet.getLastColumn()).getValues();

          // --- LANGKAH 4: LOOPING MUNDUR (Dari data terbaru di potongan itu) ---
          for (var i = data.length - 1; i >= 0; i--) {
              var row = data[i];
              var realRowNumber = startRow + i; // Nomor baris asli di Excel

              // Filter Cepat
              if (reqBulan && String(row[idx.bulan]||"").toLowerCase() !== reqBulan) continue;
              if (reqTahun && String(row[idx.tahun]||"").toLowerCase() !== reqTahun) continue;
              if (reqJenjang && String(row[idx.jenjang]||"").toUpperCase() !== reqJenjang) continue;

              var rStatusData = String(row[col.statusData] || "Diproses");
              if (rStatusData.toLowerCase().includes("hapus")) continue;
              if (reqStatus && !rStatusData.toLowerCase().includes(reqStatus)) continue;

              var rNama = (idx.nama > -1) ? String(row[idx.nama]) : "Tanpa Nama";
              var rNpsn = (idx.npsn > -1) ? String(row[idx.npsn]) : "";
              if (reqKey && !rNama.toLowerCase().includes(reqKey) && !rNpsn.includes(reqKey)) continue;

              // Format Data
              var item = {
                  rowId: realRowNumber, // PENTING: ID baris harus sesuai aslinya untuk Edit/Hapus
                  source: sourceLabel,
                  namaSekolah: rNama,
                  npsn: rNpsn,
                  bulan: String(row[idx.bulan]||""),
                  tahun: String(row[idx.tahun]||""),
                  jenjang: String(row[idx.jenjang]||""),
                  statusSekolah: (idx.statusSekolah > -1) ? row[idx.statusSekolah] : "",
                  rombel: (idx.rombel > -1) ? (parseInt(row[idx.rombel]) || 0) : 0,
                  fileUrl: (idx.file > -1) ? row[idx.file] : "",
                  
                  tglKirim: fastFormat(row[col.tglKirim]),
                  userKirim: row[col.userKirim] || "-",
                  tglEdit: fastFormat(row[col.tglEdit]),
                  userEdit: row[col.userEdit] || "-",
                  tglVerif: fastFormat(row[col.tglVerif]),
                  verifikator: row[col.userVerif] || "-",
                  statusData: rStatusData,
                  keterangan: row[col.ket] || ""
              };
              sourceResult.push(item);
          }
      } catch (e) {
          console.log("Error fetch " + sourceLabel + ": " + e.toString());
      }
      return sourceResult;
  };

  // EKSEKUSI PARALEL (Pseudocode - tetap serial di GAS tapi optimized)
  var dataPAUD = fetchDataSmart(IDS.PAUD_DATA, "Input PAUD", "PAUD");
  var dataSD = fetchDataSmart(IDS.SD_DATA, "Input SD", "SD");
  
  result = dataPAUD.concat(dataSD);
  
  // Karena kita mengambil potongan dari 2 file berbeda, 
  // kita perlu sort lagi sedikit agar gabungan PAUD & SD urut waktu secara sempurna
  // (Opsional, tapi bagus untuk UX)
  // result.sort(function(a, b) { ... }); 
  
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

/* ======================================================================
   LAPBUL.GS - VERSI FINAL (SMART MAPPING)
   Menangani SD & PAUD dengan logika pencarian kolom otomatis.
   ====================================================================== */

// --- 1. FUNGSI PEMANGGIL UTAMA (SD) ---
function simpanLapbulSD_Complex(form, fileData) {
  // Pastikan ID ini benar milik Spreadsheet SD Anda
  var ID_SD = "1u4tNL3uqt5xHITXYwHnytK6Kul9Siam-vNYuzmdZB4s"; 
  return prosesSimpanLengkap(ID_SD, "Input SD", "SD", form, fileData);
}

// --- 2. FUNGSI PEMANGGIL UTAMA (PAUD) ---
function simpanLapbulPAUD(form, fileData) {
  // Pastikan ID ini benar milik Spreadsheet PAUD Anda
  var ID_PAUD = "1an0oQQPdMh6wrUJIAzTGYk3DKFvYprK5SU7RmRXjIgs"; 
  return prosesSimpanLengkap(ID_PAUD, "Input PAUD", "PAUD", form, fileData);
}

// --- 3. MESIN PINTAR (CORE ENGINE) ---
/* ======================================================================
   MESIN PENYIMPANAN PINTAR - VERSI LENGKAP (SD & PAUD)
   Update: Menambahkan Mapping Khusus untuk Kolom PAUD
   ====================================================================== */
function prosesSimpanLengkap(idSpreadsheet, namaSheet, source, form, fileData) {
  try {
    var ss = SpreadsheetApp.openById(idSpreadsheet);
    var sheet = ss.getSheetByName(namaSheet);
    
    // 1. UPLOAD FILE
    var fileUrl = "";
    if (fileData && fileData.data) {
      // ID Folder Drive (Sesuaikan jika perlu)
      var folderId = "1I8DRQYpBbTt1mJwtD1WXVD6UK51TC8El"; 
      var blob = Utilities.newBlob(Utilities.base64Decode(fileData.data), fileData.mimeType, fileData.name);
      var file = (folderId) ? DriveApp.getFolderById(folderId).createFile(blob) : DriveApp.createFile(blob);
      file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
      fileUrl = file.getUrl();
    }

    // 2. BACA HEADER & SIAPKAN ARRAY
    // Baca sampai kolom 300 agar kolom jauh (User Kirim) terjangkau
    var headers = sheet.getRange(1, 1, 1, 300).getValues()[0].map(function(h) { 
      return String(h).toLowerCase().trim(); 
    });
    
    var rowData = new Array(headers.length).fill(""); 

    // Helper: Fungsi Pengisi Kolom (Bisa terima 1 nama atau Array nama)
    var isi = function(daftarNama, nilai) {
      if (!Array.isArray(daftarNama)) daftarNama = [daftarNama];
      
      var nilaiFinal = (nilai === null || nilai === undefined) ? "" : String(nilai);
      
      // Cek satu per satu kemungkinan nama header
      for (var i = 0; i < daftarNama.length; i++) {
        var keyword = daftarNama[i].toLowerCase();
        var idx = headers.indexOf(keyword);
        
        if (idx > -1) {
          rowData[idx] = nilaiFinal;
          return; // Stop jika sudah ketemu
        }
      }
    };

    // 3. MAPPING DATA UMUM (SD & PAUD)
    isi(["nama sekolah", "nama"], form.nama_sekolah);
    isi(["npsn"], form.npsn);
    isi(["bulan"], form.bulan);
    isi(["tahun"], form.tahun);
    isi(["jenjang"], form.jenjang);
    isi(["status sekolah", "status"], form.status_sekolah);
    isi(["rombel", "total rombel", "jumlah rombel"], form.total_rombel || form.jumlah_rombel);

    // --- MAPPING KHUSUS PAUD (SOLUSI MASALAH ANDA) ---
    if (source === "PAUD") {
        // A. DATA MURID (USIA)
        isi("0-1 L", form.u01_l); isi("0-1 P", form.u01_p);
        isi("1-2 L", form.u12_l); isi("1-2 P", form.u12_p);
        isi("2-3 L", form.u23_l); isi("2-3 P", form.u23_p);
        isi("3-4 L", form.u34_l); isi("3-4 P", form.u34_p);
        isi("4-5 L", form.u45_l); isi("4-5 P", form.u45_p);
        isi("5-6 L", form.u56_l); isi("5-6 P", form.u56_p);
        isi(["> 6 L", ">6 l"], form.u6_l);  
        isi(["> 6 P", ">6 p"], form.u6_p);

        // B. DATA KELOMPOK (TK)
        isi("A L", form.kel_a_l); isi("A P", form.kel_a_p);
        isi("B L", form.kel_b_l); isi("B P", form.kel_b_p);

        // C. DATA PTK (KEPALA SEKOLAH)
        isi("KS GTY", form.ks_gty);
        isi("KS GTT", form.ks_gtt);
        isi("KS PNS", form.ks_pns);
        isi("KS PPPK", form.ks_pppk);

        // D. DATA PTK (GURU KELAS)
        isi("GK GTY", form.gk_gty);
        isi("GK GTT", form.gk_gtt);
        isi("GK PNS", form.gk_pns);
        isi("GK PPPK", form.gk_pppk);

        // E. DATA PTK (GURU PENDAMPING)
        isi("GP GTY", form.gp_gty);
        isi("GP GTT", form.gp_gtt);
        isi("GP PNS", form.gp_pns);
        isi("GP PPPK", form.gp_pppk);

        // F. DATA TENDIK (NAMA KOLOM BEDA DIKIT)
        isi(["Penjaga", "Penjaga Sekolah"], form.td_penjaga);
        isi(["TAS", "Tenaga Administrasi", "Adm"], form.td_adm);
        isi(["Pustakawan", "Tenaga Perpustakaan"], form.td_perpus);
        isi(["Tendik Lain", "Tendik Lainnya"], form.td_lain);
    } 
    else {
        // --- JIKA SD (Tetap gunakan Looping Otomatis) ---
        for (var key in form) {
           isi([key, key.replace(/_/g, " ")], form[key]);
        }
    }

    // 4. METADATA SYSTEM & FILE
    isi(["dokumen", "file laporan", "link file"], fileUrl);
    
    var now = Utilities.formatDate(new Date(), "Asia/Jakarta", "dd/MM/yyyy HH:mm:ss");
    isi(["waktu kirim", "tgl kirim", "tanggal kirim", "timestamp"], now);
    isi(["status data", "status"], "Diproses");

    // 5. USER KIRIM (DENGAN FALLBACK)
    var userLogin = form.user_login || "Admin";
    // Cari kolom User Kirim, User Input, atau Pengirim
    isi(["user kirim", "user input", "pengirim"], userLogin);

    // 6. SIMPAN
    sheet.appendRow(rowData);
    
    // Update Data Master jika perlu
    if (typeof updateDataSekolahMaster === 'function') {
        updateDataSekolahMaster(form);
    }

    return { success: true, message: "Laporan berhasil disimpan! (" + userLogin + ")" };

  } catch (e) {
    return { success: false, message: "Error Server: " + e.toString() };
  }
}

// --- 4. FUNGSI UPDATE MASTER (Tetap Pertahankan yang Lama) ---
function updateDataSekolahMaster(form) {
  try {
    // GANTI ID SPREADSHEET MASTER DATA ANDA
    const ss = SpreadsheetApp.openById("1wiDKez4rL5UYnpP2-OZjYowvmt1nRx-fIMy9trJlhBA");
    const sheet = ss.getSheetByName("Data_Sekolah");
    const data = sheet.getDataRange().getValues();
    
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][0]).trim() === String(form.npsn).trim()) {
        // Update Kolom E (4) s.d P (15)
        var updateRange = sheet.getRange(i + 1, 5, 1, 12); 
        var updateValues = [[
          form.yayasan, form.no_sk_pendirian, form.tgl_pendirian, 
          form.no_sk_ijin, form.tgl_ijin, form.akreditasi, 
          form.skor, form.no_sertifikat, form.tgl_sertifikat, 
          form.alamat, form.telepon, form.email
        ]];
        updateRange.setValues(updateValues);
        break;
      }
    }
  } catch(e) { Logger.log("Update Master Error: " + e); }
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

/* ======================================================================
   FITUR HAPUS DATA (SOFT DELETE)
   ====================================================================== */
function softDeleteLapbulSD(rowId, userLogin) {
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_IDS.SD_DATA);
    var sheet = ss.getSheetByName("Input SD");
    
    // Validasi Baris (Pastikan rowId valid)
    var r = parseInt(rowId);
    if (isNaN(r) || r < 2) return { success: false, message: "ID Baris tidak valid" };

    // Update Metadata Delete
    // Kolom HK (Index 219) = Status Data -> Ubah jadi "Dihapus"
    // Kolom HM (Index 221) = Tanggal Edit
    // Kolom HN (Index 222) = User Edit (Siapa yang menghapus)
    
    var now = new Date();
    var strTgl = Utilities.formatDate(now, Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm:ss");

    // Array update: [Status, User Kirim (skip), Tgl Edit, User Edit]
    // Posisi di Sheet:
    // HK (219) -> Status
    // HM (221) -> Tgl Edit
    // HN (222) -> User Edit
    
    // Kita tembak langsung sel-nya agar akurat
    sheet.getRange(r, 219).setValue("Dihapus"); // HK
    sheet.getRange(r, 221).setValue("'" + strTgl); // HM
    sheet.getRange(r, 222).setValue("'" + userLogin); // HN
    
    return { success: true, message: "Data berhasil dihapus dari database." };
    
  } catch (e) {
    return { success: false, message: "Gagal Hapus: " + e.toString() };
  }
}

/* ======================================================================
   FITUR HAPUS DATA (MOVE TO TRASH & ARCHIVE)
   Update: SD & PAUD Support
   ====================================================================== */
function processDeleteData(source, rowId, inputCode, userLogin) {
  try {
    // 1. VALIDASI KODE KEAMANAN (Server Side)
    var now = new Date();
    var y = now.getFullYear();
    var m = String(now.getMonth() + 1).padStart(2, '0');
    var d = String(now.getDate()).padStart(2, '0');
    var serverCode = y + m + d; // Pola: 20260203

    if (String(inputCode).trim() !== serverCode) {
      return { success: false, message: "Kode Hapus Salah! Hubungi Admin Korwil." };
    }

    // 2. KONFIGURASI SESUAI JENJANG
    var config = {};
    if (source === 'SD') {
      config = {
        ssId: SPREADSHEET_IDS.SD_DATA, // 1u4tNL3uqt5xHITXYwHnytK6Kul9Siam-vNYuzmdZB4s
        sheetName: "Input SD",
        trashSheetName: "Trash",
        trashFolderId: "1MpEgpCDrTX-SHjdNIa3aUpKUyYZpejrb"
      };
    } else {
      config = {
        ssId: SPREADSHEET_IDS.PAUD_DATA, // 1an0oQQPdMh6wrUJIAzTGYk3DKFvYprK5SU7RmRXjIgs
        sheetName: "Input PAUD",
        trashSheetName: "Trash",
        trashFolderId: "1EUIOthRbotJQlSphxVZ-QAdewe17UCOU"
      };
    }

    var ss = SpreadsheetApp.openById(config.ssId);
    var sheetMain = ss.getSheetByName(config.sheetName);
    var sheetTrash = ss.getSheetByName(config.trashSheetName);
    
    if (!sheetMain || !sheetTrash) {
      return { success: false, message: "Sheet Database/Trash tidak ditemukan!" };
    }

    var r = parseInt(rowId);
    var lastCol = sheetMain.getLastColumn();
    
    // 3. AMBIL DATA BARIS
    var rowRange = sheetMain.getRange(r, 1, 1, lastCol);
    var rowValues = rowRange.getValues()[0];
    
    // Cari URL File (Biasanya ada kata 'http') untuk dipindahkan
    // Kita cari kolom yang isinya link drive
    var fileUrl = "";
    var colIndexFile = -1;
    
    // Header check untuk memastikan kolom file
    var headers = sheetMain.getRange(1, 1, 1, lastCol).getValues()[0];
    for(var h=0; h<headers.length; h++) {
       var head = String(headers[h]).toLowerCase();
       if(head.includes("dokumen") || head.includes("file") || head.includes("link")) {
           fileUrl = rowValues[h];
           colIndexFile = h;
           break;
       }
    }

    // 4. PINDAHKAN FILE KE FOLDER TRASH
    var moveStatus = "File tidak ditemukan";
    if (fileUrl && fileUrl.includes("drive.google.com")) {
       try {
         var fileId = fileUrl.match(/[-\w]{25,}/); // Regex ambil ID
         if (fileId) {
            var file = DriveApp.getFileById(fileId[0]);
            var folderTrash = DriveApp.getFolderById(config.trashFolderId);
            file.moveTo(folderTrash); // Pindah Folder
            moveStatus = "File dipindahkan ke Trash";
         }
       } catch (errFile) {
         moveStatus = "Gagal pindah file: " + errFile.message;
       }
    }

    // 5. TAMBAHKAN METADATA PENGHAPUSAN (Di kolom paling belakang Trash)
    // Format Trash Sheet sebaiknya sama kolomnya, lalu kita append info hapus
    var trashData = rowValues.slice(); // Copy array
    trashData.push("Dihapus oleh: " + userLogin);
    trashData.push("Tgl Hapus: " + Utilities.formatDate(now, Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm:ss"));
    trashData.push(moveStatus);

    // 6. SIMPAN KE SHEET TRASH
    sheetTrash.appendRow(trashData);

    // 7. HAPUS DARI SHEET UTAMA
    sheetMain.deleteRow(r);

    return { success: true, message: "Data berhasil dihapus dari database." };

  } catch (e) {
    return { success: false, message: "Error System: " + e.toString() };
  }
}

/* ======================================================================
   FITUR VERIFIKASI DATA (SD & PAUD)
   Update: Mapping Kolom Spesifik
   ====================================================================== */
function processVerifikasiLapbul(source, rowId, status, keterangan, userLogin) {
  try {
    var config = {};
    
    // 1. TENTUKAN KONFIGURASI KOLOM
    if (source === 'SD') {
      config = {
        ssId: SPREADSHEET_IDS.SD_DATA,
        sheetName: "Input SD",
        // SD: HK(219), HQ(225), HO(223), HP(224)
        colStatus: 219,    
        colKet: 225,       
        colTglVerif: 223,  
        colUserVerif: 224  
      };
    } else {
      config = {
        ssId: SPREADSHEET_IDS.PAUD_DATA,
        sheetName: "Input PAUD",
        
        // --- MAPPING KOLOM PAUD (PENTING!) ---
        // Kolom E (5)  = Status Sekolah (JANGAN DIGANGGU)
        // Kolom AW (49) = Status Data/Verifikasi (TARGET KITA)
        
        // Hitungan: A-Z(26) + AA-AV(22) = 48. Maka AW = 49.
        colStatus: 49,     // AW (Status Data: Disetujui/Revisi/dll)
        colKet: 50,        // AX (Keterangan)
        colTglVerif: 47,   // AU (Tanggal Verif)
        colUserVerif: 48   // AV (Verifikator)
      };
    }

    var ss = SpreadsheetApp.openById(config.ssId);
    var sheet = ss.getSheetByName(config.sheetName);
    var r = parseInt(rowId);

    if (!sheet || isNaN(r)) return { success: false, message: "Referensi Data Salah" };

    // 2. SIAPKAN DATA
    var now = new Date();
    var strTgl = Utilities.formatDate(now, Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm");

    // 3. TULIS KE DATABASE
    // Kita tembak spesifik ke kolom AW, AX, AU, AV
    sheet.getRange(r, config.colStatus).setValue(status);           
    sheet.getRange(r, config.colKet).setValue(keterangan);          
    sheet.getRange(r, config.colTglVerif).setValue("'" + strTgl);   
    sheet.getRange(r, config.colUserVerif).setValue(userLogin);     

    return { 
      success: true, 
      message: "Data berhasil diverifikasi: " + status,
      newData: {
        status: status, // Ini akan mengupdate Status Data di Cache (AW)
        ket: keterangan,
        tgl: strTgl,
        user: userLogin
      }
    };

  } catch (e) {
    return { success: false, message: "Gagal Verifikasi: " + e.toString() };
  }
}

/* ======================================================================
   REKAP STATUS LAPORAN BULAN (LOAD ALL DATA FOR CLIENT SIDE FILTER)
   ====================================================================== */
function getRekapLapbulStatus() {
  // KITA TIDAK PAKAI PARAMETER FILTER LAGI, AMBIL SEMUA
  var result = { headers: [], rows: [] };
  
  var ID_SD = "1u4tNL3uqt5xHITXYwHnytK6Kul9Siam-vNYuzmdZB4s";
  var ID_PAUD = "1an0oQQPdMh6wrUJIAzTGYk3DKFvYprK5SU7RmRXjIgs";

  // HEADER: Kita tetapkan statis saja biar rapi, karena kolom dinamis
  result.headers = [
    "Nama Sekolah", "NPSN", "Jenjang", 
    "Jan", "Feb", "Mar", "Apr", "Mei", "Jun", 
    "Jul", "Agu", "Sep", "Okt", "Nov", "Des"
  ];

  var fetchData = function(id, sheetName, defaultJenjang) {
    var temp = [];
    try {
      var ss = SpreadsheetApp.openById(id);
      var sheet = ss.getSheetByName(sheetName);
      if (!sheet) return [];
      
      var data = sheet.getDataRange().getValues();
      if (data.length < 2) return [];

      for (var i = 1; i < data.length; i++) {
        var row = data[i];
        
        // A=0(Nama), B=1(NPSN), C=2(Jenjang), D=3(Tahun)
        var rTahun = String(row[3] || "").trim(); 
        var rJenjang = String(row[2] || defaultJenjang).toUpperCase().trim();

        // KITA MASUKKAN TAHUN KE DALAM ARRAY (INDEX 3) UNTUK FILTER DI CLIENT
        // Structure: [0:Nama, 1:NPSN, 2:Jenjang, 3:TAHUN, 4:Jan ... 15:Des]
        var cleanRow = [
          row[0], 
          row[1], 
          rJenjang, 
          rTahun, // <--- PENTING: TAHUN DISIMPAN DI SINI
          row[4], row[5], row[6], row[7], row[8], row[9],
          row[10], row[11], row[12], row[13], row[14], row[15]
        ];
        temp.push(cleanRow);
      }
    } catch (e) {
      // ignore
    }
    return temp;
  };

  var rowsSD = fetchData(ID_SD, "Status SD", "SD");
  var rowsPAUD = fetchData(ID_PAUD, "Status PAUD", "PAUD");

  result.rows = rowsSD.concat(rowsPAUD);

  // Sorting default
  result.rows.sort(function(a, b) {
    if (a[2] === b[2]) return (a[0] < b[0]) ? -1 : 1;
    return (a[2] < b[2]) ? -1 : 1;
  });

  return result;
}

/* ======================================================================
   DASHBOARD LAPBUL (OPTIMIZED SINGLE PASS)
   ====================================================================== */

// Konfigurasi ID Spreadsheet
var CONF_LAPBUL = {
  sd:   "1u4tNL3uqt5xHITXYwHnytK6Kul9Siam-vNYuzmdZB4s",
  paud: "1an0oQQPdMh6wrUJIAzTGYk3DKFvYprK5SU7RmRXjIgs"
};

// 1. FUNGSI KHUSUS SD (Cepat)
function getLapbulMetric_SD(tahun, bulan) {
  return processSheet(CONF_LAPBUL.sd, "Status SD", tahun, bulan, ["SD"]);
}

// 2. FUNGSI KHUSUS PAUD (Single Pass Loop untuk TK, KB, SPS)
function getLapbulMetric_PAUD(tahun, bulan) {
  // Kita ambil sekaligus TK, KB, dan SPS dalam satu kali buka file
  return processSheet(CONF_LAPBUL.paud, "Status PAUD", tahun, bulan, ["TK", "KB", "SPS"]);
}

// CORE PROCESSOR (UPDATED: RETURN LIST BELUM LAPOR)
function processSheet(idSS, sheetName, tahun, bulan, targetJenjangArray) {
  // Struktur Result Dinamis
  var result = { recent: [] };
  
  targetJenjangArray.forEach(function(j) {
    // Tambahkan array 'listBelum' di sini
    result[j.toLowerCase()] = { 
        total:0, sudah:0, belum:0, persen:0, 
        disetujui:0, diproses:0, revisi:0, ditolak:0,
        listBelum: [] // <--- ARRAY PENAMPUNG NAMA
    };
  });

  try {
    var ss = SpreadsheetApp.openById(idSS);
    var sheet = ss.getSheetByName(sheetName);
    if (!sheet) return JSON.stringify(result);

    var data = sheet.getDataRange().getValues();
    if (data.length < 2) return JSON.stringify(result);

    var idxStatus = parseInt(bulan) + 3; 

    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      var rTahun = String(row[3]).trim(); 
      var rJenjang = String(row[2]).trim().toUpperCase();

      if (rTahun !== String(tahun)) continue;
      
      // Fix: Pastikan jenjang cocok persis (misal 'SD' tidak boleh masuk ke 'SPS')
      var indexJenjang = targetJenjangArray.indexOf(rJenjang);
      if (indexJenjang === -1) continue;

      var key = rJenjang.toLowerCase(); 
      var stats = result[key];
      
      stats.total++;

      var rawStatus = String(row[idxStatus] || "").trim();
      var st = rawStatus.toLowerCase();

      // LOGIKA BELUM LAPOR
      if (st === "" || st === "-" || st === "0") {
        stats.belum++;
        // MASUKKAN NAMA SEKOLAH KE LIST (Kolom A / Index 0)
        stats.listBelum.push(row[0]); 
      } else {
        stats.sudah++;
        
        if (st.includes('revisi') || st.includes('perbaiki')) stats.revisi++;
        else if (st.includes('tolak') || st.includes('x') || st.includes('salah')) stats.ditolak++;
        else if (st.includes('ok') || st.includes('setuju') || st.includes('valid')) stats.disetujui++;
        else stats.diproses++;

        if (result.recent.length < 10) {
           result.recent.push({ sekolah: row[0], jenjang: rJenjang, status: rawStatus });
        }
      }
    }

    targetJenjangArray.forEach(function(j) {
       var k = j.toLowerCase();
       var s = result[k];
       s.persen = s.total === 0 ? 0 : Math.round((s.sudah / s.total) * 100);
       // Sortir nama sekolah biar rapi (A-Z)
       s.listBelum.sort(); 
    });

  } catch (e) { result.error = e.toString(); }

  return JSON.stringify(result);
}

/* ======================================================================
   LAPBUL.GS - VERSI TEXT-ONLY (GET DISPLAY VALUES)
   Mengambil tampilan teks layar agar aman dari error Date/Formula.
   ====================================================================== */

function getLapbulDataSD(bulan, tahun, status, keyword) {
  // CONFIG
  var ID_SD = "1u4tNL3uqt5xHITXYwHnytK6Kul9Siam-vNYuzmdZB4s"; 
  var NAMA_SHEET = "Input SD";

  // SANITASI INPUT
  var qBulan = (bulan) ? String(bulan).toLowerCase() : "";
  var qTahun = (tahun) ? String(tahun) : "";
  var qStatus = (status) ? String(status).toLowerCase() : "";
  var qKey = (keyword) ? String(keyword).toLowerCase() : "";
  var LIMIT = (qKey.length > 2) ? 1000 : 300; 

  return fetchDataDisplay(ID_SD, NAMA_SHEET, "SD", LIMIT, qBulan, qTahun, qStatus, qKey);
}

function getLapbulDataPAUD(bulan, tahun, status, keyword) {
  var ID_PAUD = "1an0oQQPdMh6wrUJIAzTGYk3DKFvYprK5SU7RmRXjIgs"; 
  var NAMA_SHEET = "Input PAUD"; 

  var qBulan = (bulan) ? String(bulan).toLowerCase() : "";
  var qTahun = (tahun) ? String(tahun) : "";
  var qStatus = (status) ? String(status).toLowerCase() : "";
  var qKey = (keyword) ? String(keyword).toLowerCase() : "";
  var LIMIT = (qKey.length > 2) ? 1000 : 300;

  return fetchDataDisplay(ID_PAUD, NAMA_SHEET, "PAUD", LIMIT, qBulan, qTahun, qStatus, qKey);
}

// --- ENGINE: MENGGUNAKAN getDisplayValues() ---
function fetchDataDisplay(id, sheetName, sourceLabel, limit, reqBulan, reqTahun, reqStatus, reqKey) {
  var result = [];
  
  try {
    var ss = SpreadsheetApp.openById(id);
    var sheet = ss.getSheetByName(sheetName);
    
    // ERROR CHECKING
    if (!sheet) return [{ error: "Sheet '" + sheetName + "' tidak ditemukan!" }];

    var lastRow = sheet.getLastRow();
    if (lastRow < 2) return []; // Kosong

    // 1. TENTUKAN RANGE (Potong Data)
    var startRow = Math.max(2, lastRow - limit + 1);
    var numRows = (lastRow - startRow + 1);
    
    // 2. AMBIL HEADER (Baris 1)
    // getValues() aman untuk header karena biasanya string
    var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map(function(h) { return String(h).toLowerCase(); });
    
    // 3. MAPPING KOLOM
    var idx = {
        nama: headers.indexOf("nama sekolah"),
        npsn: headers.indexOf("npsn"),
        bulan: headers.indexOf("bulan"),
        tahun: headers.indexOf("tahun"),
        jenjang: headers.indexOf("jenjang"),
        status: headers.findIndex(function(h) { return h.includes("status sekolah") || h === "status"; }),
        rombel: (sourceLabel === 'PAUD') ? headers.indexOf("jumlah rombel") : headers.indexOf("total rombel"),
        file: headers.findIndex(function(h) { return h.includes("file") || h.includes("dokumen"); })
    };

    if (idx.nama === -1) return [{ error: "Kolom 'Nama Sekolah' tidak ditemukan di baris 1." }];

    // Index Activity (Manual - Sesuaikan jika geser)
    // PAUD: Col 43 (AR) start | SD: Col 219 (HM) start
    // Ingat: Array index dimulai dari 0. Jadi Col A = 0.
    var col = (sourceLabel === 'PAUD') ? 
              { tglKirim:0, userKirim:43, tglEdit:44, userEdit:45, tglVerif:46, userVerif:47, statusData:48, ket:49 } : 
              { tglKirim:0, userKirim:219, tglEdit:220, userEdit:221, tglVerif:222, userVerif:223, statusData:218, ket:224 };

    // 4. AMBIL DATA UTAMA SEBAGAI STRING (The Fix!)
    // getDisplayValues() mengubah semua tanggal/angka/error menjadi String persis tampilan excel
    var data = sheet.getRange(startRow, 1, numRows, sheet.getLastColumn()).getDisplayValues();

    // 5. LOOPING
    // Data diambil dari array 'data', index 0 adalah baris 'startRow'
    for (var i = data.length - 1; i >= 0; i--) {
        var row = data[i]; // row isinya sudah PASTI string semua
        
        // Filter
        var rBulan = row[idx.bulan].toLowerCase();
        var rTahun = row[idx.tahun];
        
        if (reqBulan && rBulan !== reqBulan) continue;
        if (reqTahun && rTahun !== reqTahun) continue;

        var rStatusData = row[col.statusData];
        if (rStatusData.toLowerCase().includes("hapus")) continue;
        if (reqStatus && !rStatusData.toLowerCase().includes(reqStatus)) continue;

        var rNama = row[idx.nama];
        var rNpsn = row[idx.npsn];
        if (reqKey && !rNama.toLowerCase().includes(reqKey) && !rNpsn.includes(reqKey)) continue;

        // Masukkan Data
        result.push({
            rowId: startRow + i, // ID Asli baris Excel
            source: sourceLabel,
            namaSekolah: rNama || "Tanpa Nama",
            npsn: rNpsn,
            bulan: row[idx.bulan],
            tahun: row[idx.tahun],
            jenjang: row[idx.jenjang],
            statusSekolah: row[idx.status],
            rombel: row[idx.rombel] || "0",
            fileUrl: row[idx.file], // URL File
            
            statusData: rStatusData || "Diproses",
            keterangan: row[col.ket],
            
            // Tanggal sudah jadi string (misal "06/02/2026"), aman dikirim
            tglKirim: row[col.tglKirim], 
            userKirim: row[col.userKirim],
            tglEdit: row[col.tglEdit],
            userEdit: row[col.userEdit],
            tglVerif: row[col.tglVerif],
            verifikator: row[col.userVerif],
            
            // Untuk sorting, kita pakai index saja biar cepat
            sortTime: i 
        });
    }

  } catch (e) {
    return [{ error: "SERVER ERROR: " + e.message }];
  }
  
  return result;
}