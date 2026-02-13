/* ======================================================================
   SIABA PRESENSI HARIAN - DISPLAY VALUES VERSION
   Menggunakan .getDisplayValues() untuk menjamin data sesuai teks asli spreadsheet.
   ====================================================================== */

function getSiabaFilters() {
  const ID_DB = "1wiDKez4rL5UYnpP2-OZjYowvmt1nRx-fIMy9trJlhBA";
  try {
    const ss = SpreadsheetApp.openById(ID_DB);
    const sheet = ss.getSheetByName("Lookup Siaba");
    if (!sheet) return JSON.stringify({ error: "Sheet 'Lookup Siaba' tidak ditemukan." });
    
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return JSON.stringify({ years: [], months: [] });
    
    const data = sheet.getRange(2, 1, lastRow - 1, 2).getDisplayValues();
    
    let years = new Set();
    let months = new Set();
    
    data.forEach(row => {
      if (row[0]) years.add(row[0]); 
      if (row[1]) months.add(row[1]); 
    });

    const URUTAN_BULAN = ["Januari", "Februari", "Maret", "April", "Mei", "Juni", "Juli", "Agustus", "September", "Oktober", "November", "Desember"];
    let sortedMonths = Array.from(months).sort((a, b) => URUTAN_BULAN.indexOf(a) - URUTAN_BULAN.indexOf(b));

    return JSON.stringify({
      years: Array.from(years).sort().reverse(),
      months: sortedMonths
    });
  } catch (e) {
    return JSON.stringify({ error: e.message });
  }
}

function getSiabaPresensiHarian(filterTahun, filterBulan, filterUnit) {
  const ID_DB = "1wiDKez4rL5UYnpP2-OZjYowvmt1nRx-fIMy9trJlhBA"; 
  
  try {
    // 1. CARI FILE TARGET
    var ssLookup = SpreadsheetApp.openById(ID_DB);
    var sheetLookup = ssLookup.getSheetByName("Lookup Siaba");
    var dataLookup = sheetLookup.getDataRange().getDisplayValues();
    var targetId = "", customSheet = "";
    
    for (var i = 1; i < dataLookup.length; i++) {
        if (dataLookup[i][0] == filterTahun && dataLookup[i][1] == filterBulan) {
            targetId = dataLookup[i][2];
            customSheet = dataLookup[i][3];     
            break; 
        }
    }
    
    if (!targetId) return JSON.stringify({ error: "Data Periode " + filterBulan + " " + filterTahun + " belum tersedia." });

    // 2. BUKA SHEET
    var ssTarget = SpreadsheetApp.openById(targetId);
    var sheetTarget = customSheet ? ssTarget.getSheetByName(customSheet) : ssTarget.getSheets()[0];
    if (!sheetTarget) sheetTarget = ssTarget.getSheetByName("Data Siaba");

    // 3. AMBIL DATA
    var lastRow = sheetTarget.getLastRow();
    var lastCol = sheetTarget.getLastColumn(); 
    if (lastCol < 87) return JSON.stringify({ error: "Format kolom sheet tidak sesuai." });

    var allData = sheetTarget.getRange(1, 1, lastRow, lastCol).getDisplayValues();
    var headerRow = allData[0].slice(3, 87); // Header D s.d CI
    var rawRows = allData.slice(1);

    var cleanRows = [];
    for (var i = 0; i < rawRows.length; i++) {
        var r = rawRows[i];
        // Filter Server Side (Hanya jika diminta spesifik, tapi biasanya SEMUA)
        if (filterUnit === "SEMUA" || r[2] === filterUnit) {
            cleanRows.push(r);
        }
    }

    // 4. SORTING BERTINGKAT (TP > TA > PLA > LA)
    cleanRows.sort(function(a, b) {
        var tpA = parseInt(a[5]) || 0; var tpB = parseInt(b[5]) || 0;
        if (tpB !== tpA) return tpB - tpA; 
        
        var taA = parseInt(a[20]) || 0; var taB = parseInt(b[20]) || 0;
        if (taB !== taA) return taB - taA; 

        var plaA = parseInt(a[22]) || 0; var plaB = parseInt(b[22]) || 0;
        if (plaB !== plaA) return plaB - plaA; 

        var laA = parseInt(a[24]) || 0; var laB = parseInt(b[24]) || 0;
        return laB - laA; 
    });

    // 5. DATA MAPPING (D-CI + UNIT HIDDEN)
    // Kita tambahkan Unit Kerja (r[2]) di indeks terakhir array agar bisa difilter di frontend
    var finalData = cleanRows.map(function(row) {
        var dataD_CI = row.slice(3, 87); // Kolom D s.d CI
        var unitMeta = row[2];           // Kolom C (Unit)
        return dataD_CI.concat([unitMeta]); // Gabung: [D...CI, UNIT]
    });

    return JSON.stringify({
      headers: headerRow,
      rows: finalData
    });

  } catch (e) {
    return JSON.stringify({ error: "Error Server: " + e.message });
  }
}

/* ======================================================================
   SIABA APEL & UPACARA 
   ====================================================================== */

function getSiabaApelFilters() {
  const ID_DB = "1wiDKez4rL5UYnpP2-OZjYowvmt1nRx-fIMy9trJlhBA"; 
  try {
    const ss = SpreadsheetApp.openById(ID_DB);
    const sheet = ss.getSheetByName("Lookup Siaba");
    if (!sheet) return JSON.stringify({ error: "Sheet 'Lookup Siaba' tidak ditemukan." });
    
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return JSON.stringify({ years: [], months: [] });
    
    const data = sheet.getRange(2, 1, lastRow - 1, 2).getDisplayValues();
    
    let years = new Set();
    let months = new Set();
    
    data.forEach(row => {
      if (row[0]) years.add(row[0]); 
      if (row[1]) months.add(row[1]); 
    });

    const URUTAN_BULAN = ["Januari", "Februari", "Maret", "April", "Mei", "Juni", "Juli", "Agustus", "September", "Oktober", "November", "Desember"];
    let sortedMonths = Array.from(months).sort((a, b) => URUTAN_BULAN.indexOf(a) - URUTAN_BULAN.indexOf(b));

    return JSON.stringify({
      years: Array.from(years).sort().reverse(),
      months: sortedMonths
    });
  } catch (e) {
    return JSON.stringify({ error: e.message });
  }
}

function getSiabaDataApel(filterTahun, filterBulan, filterUnit) {
  const ID_DB = "1wiDKez4rL5UYnpP2-OZjYowvmt1nRx-fIMy9trJlhBA";
  
  try {
    // 1. CARI FILE ID DI LOOKUP
    const ssLookup = SpreadsheetApp.openById(ID_DB);
    const sheetLookup = ssLookup.getSheetByName("Lookup Siaba");
    const dataLookup = sheetLookup.getDataRange().getDisplayValues();
    
    let targetId = "";
    for (let i = 1; i < dataLookup.length; i++) {
        if (dataLookup[i][0] == filterTahun && dataLookup[i][1] == filterBulan) {
            targetId = dataLookup[i][2]; 
            break; 
        }
    }

    if (!targetId) return JSON.stringify({ error: `Data Apel ${filterBulan} ${filterTahun} tidak ditemukan.` });

    // 2. BUKA FILE TARGET
    const ssTarget = SpreadsheetApp.openById(targetId);
    const sheetTarget = ssTarget.getSheetByName("Data Apel");
    if (!sheetTarget) return JSON.stringify({ error: `Sheet "Data Apel" tidak ditemukan.` });

    const allData = sheetTarget.getDataRange().getDisplayValues();
    
    // Header & Data: Ambil dari Kolom D (Index 3) sampai AP (Index 41)
    const headerData = allData[0].slice(3, 42); 
    
    allData.shift(); // Hapus baris header
    
    let result = [];
    
    for (let i = 0; i < allData.length; i++) {
        let row = allData[i];
        if (row.length < 3) continue;
        
        let rowUnit = row[2]; // Kolom C = Unit Kerja
        
        // Filter Unit: Kirim SEMUA data (Client Side Cache)
        if (filterUnit === "SEMUA" || rowUnit == filterUnit) {
             let dataCells = row.slice(3, 42); // Data D s.d AP
             
             // Tambahkan Unit di paling belakang (Hidden) untuk filter
             result.push(dataCells.concat([rowUnit]));
        }
    }
    
    return JSON.stringify({
      headers: headerData,
      rows: result
    });

  } catch (e) {
    return JSON.stringify({ error: "SYSTEM ERROR: " + e.message });
  }
}

/* ======================================================================
   SIABA TIDAK PRESENSI (SMART CACHE SUPPORT)
   ====================================================================== */

function getSiabaTidakFilters() {
  const ID_DB = "1wiDKez4rL5UYnpP2-OZjYowvmt1nRx-fIMy9trJlhBA"; 
  try {
    const ss = SpreadsheetApp.openById(ID_DB);
    const sheet = ss.getSheetByName("Lookup Siaba");
    if (!sheet) return JSON.stringify({ error: "Sheet 'Lookup Siaba' tidak ditemukan." });
    
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return JSON.stringify({ years: [], months: [] });
    
    const data = sheet.getRange(2, 1, lastRow - 1, 2).getDisplayValues();
    
    let years = new Set();
    let months = new Set();
    
    data.forEach(row => {
      if (row[0]) years.add(row[0]); 
      if (row[1]) months.add(row[1]); 
    });

    const URUTAN_BULAN = ["Januari", "Februari", "Maret", "April", "Mei", "Juni", "Juli", "Agustus", "September", "Oktober", "November", "Desember"];
    let sortedMonths = Array.from(months).sort((a, b) => URUTAN_BULAN.indexOf(a) - URUTAN_BULAN.indexOf(b));

    return JSON.stringify({
      years: Array.from(years).sort().reverse(),
      months: sortedMonths
    });
  } catch (e) {
    return JSON.stringify({ error: e.message });
  }
}

function getSiabaTidakData(filterTahun, filterBulan, filterUnit) {
  const ID_DB = "1wiDKez4rL5UYnpP2-OZjYowvmt1nRx-fIMy9trJlhBA";
  
  try {
    // 1. CEK LOOKUP
    let ssLookup;
    try { ssLookup = SpreadsheetApp.openById(ID_DB); } 
    catch(e) { return JSON.stringify({ error: "Gagal buka Database Lookup." }); }

    const sheetLookup = ssLookup.getSheetByName("Lookup Siaba");
    if (!sheetLookup) return JSON.stringify({ error: "Sheet Lookup Siaba hilang." });

    const dataLookup = sheetLookup.getDataRange().getDisplayValues();
    let targetId = "";
    
    for (let i = 1; i < dataLookup.length; i++) {
        if (dataLookup[i][0] == filterTahun && dataLookup[i][1] == filterBulan) {
            targetId = dataLookup[i][2]; 
            break; 
        }
    }

    if (!targetId) return JSON.stringify({ error: `Data ${filterBulan} ${filterTahun} belum ada di Lookup.` });

    // 2. BUKA FILE TARGET
    let ssTarget;
    try { ssTarget = SpreadsheetApp.openById(targetId); }
    catch(e) { return JSON.stringify({ error: `Gagal akses File ID: ...${targetId.substr(-5)}` }); }

    const TARGET_SHEET_NAME = "Data Alpa";
    const sheetTarget = ssTarget.getSheetByName(TARGET_SHEET_NAME);

    if (!sheetTarget) return JSON.stringify({ error: `Sheet "${TARGET_SHEET_NAME}" tidak ditemukan di file target.` });

    const maxCol = sheetTarget.getLastColumn();
    if (maxCol < 4) return JSON.stringify({ error: `Sheet Data Alpa kolom < 4.` });

    const allData = sheetTarget.getDataRange().getDisplayValues();
    const headerData = allData[0].slice(3); // Ambil header mulai kolom D
    
    allData.shift(); // Hapus header row
    
    let result = [];
    
    for (let i = 0; i < allData.length; i++) {
        let row = allData[i];
        if (row.length < 3) continue;
        
        let rowUnit = row[2]; // Kolom C = Unit Kerja
        
        // Logika Smart Cache: Kirim SEMUA data unit ke browser
        if (filterUnit === "SEMUA" || rowUnit == filterUnit) {
            let rowData = row.slice(3, 3 + headerData.length);
            
            // PENTING: Tambahkan Unit di elemen terakhir array (Hidden) untuk filter di browser
            rowData.push(rowUnit);
            
            result.push(rowData);
        }
    }

    // Sort berdasarkan kolom ke-3 (Index 2) -> Biasanya jumlah Alpha/Lupa
    if (result.length > 0) {
        result.sort((a, b) => {
            const valA = parseInt(a[2]) || 0;
            const valB = parseInt(b[2]) || 0;
            return valB - valA; 
        });
    }
    
    return JSON.stringify({
      headers: headerData,
      rows: result
    });

  } catch (e) {
    return JSON.stringify({ error: "SYSTEM ERROR: " + e.message });
  }
}

/* ======================================================================
   SIABA TERLAMBAT
   ====================================================================== */

function getSiabaTerlambatFilters() {
  const ID_DB = "1tQsQY1-Ny1ie66GOZPTLtvZ7BiYCgFdNrX-AVGCtaHA"; 
  try {
    const ss = SpreadsheetApp.openById(ID_DB);
    const sheet = ss.getSheetByName("Rekap_Terlambat");
    if (!sheet) return JSON.stringify({ years: [] }); // Return empty array biar gak error
    
    const lastRow = sheet.getLastRow();
    if (lastRow < 3) return JSON.stringify({ years: [] });
    
    // Ambil Kolom A (Tahun)
    const data = sheet.getRange(3, 1, lastRow - 2, 1).getDisplayValues();
    
    let years = new Set();
    data.forEach(row => {
      if (row[0]) years.add(row[0]); 
    });

    return JSON.stringify({
      years: Array.from(years).sort().reverse()
    });
  } catch (e) {
    return JSON.stringify({ error: e.message });
  }
}

function getSiabaTerlambatData(filterTahun, filterUnit) {
  const ID_DB = "1tQsQY1-Ny1ie66GOZPTLtvZ7BiYCgFdNrX-AVGCtaHA";
  
  try {
    const ss = SpreadsheetApp.openById(ID_DB);
    const sheet = ss.getSheetByName("Rekap_Terlambat");
    if (!sheet) return JSON.stringify({ error: "Sheet 'Rekap_Terlambat' tidak ditemukan." });

    const maxCol = sheet.getLastColumn(); 
    const lastRow = sheet.getLastRow();
    
    // Header (Baris 1 & 2, Mulai Kolom C)
    const headerRange = sheet.getRange(1, 3, 2, maxCol - 2).getDisplayValues();
    const headerTop = headerRange[0]; 
    const headerSub = headerRange[1]; 

    if (lastRow < 3) return JSON.stringify({ error: "Data Kosong" });

    // Ambil Semua Data (Mulai Baris 3)
    const rawData = sheet.getRange(3, 1, lastRow - 2, maxCol).getDisplayValues();
    
    let result = [];
    
    // Logic Filter
    // Kita abaikan filterUnit di server side jika requestnya "SEMUA", 
    // tapi kita TETAP kirim data Unitnya ke frontend (hidden) agar bisa difilter di sana.
    
    for (let i = 0; i < rawData.length; i++) {
        let row = rawData[i];
        
        let rowTahun = String(row[0]).trim(); 
        let rowUnit  = String(row[1]).toUpperCase().trim(); 
        
        // Filter Tahun Wajib Server Side
        if (rowTahun == String(filterTahun).trim()) {
             // Ambil Data Tampilan (Mulai Kolom C / Index 2)
             let rowDisplay = row.slice(2); 
             
             // [PENTING] Sisipkan Unit Kerja di array paling belakang untuk filter Frontend
             rowDisplay.push(rowUnit); 
             
             result.push(rowDisplay);
        }
    }

    // Sorting (Berdasarkan Kolom Terakhir Data Tampilan - Sebelum Unit disisipkan)
    // Index Total adalah length - 2 (karena length-1 sekarang adalah Unit)
    if (result.length > 0) {
        result.sort((a, b) => {
            // Kolom Total adalah index kedua dari belakang
            let idxTotal = a.length - 2; 
            let valA = parseInt(String(a[idxTotal]).replace(/\./g,'')) || 0;
            let valB = parseInt(String(b[idxTotal]).replace(/\./g,'')) || 0;
            return valB - valA; // Descending
        });
    }

    return JSON.stringify({
      headerTop: headerTop,
      headerSub: headerSub,
      rows: result
    });

  } catch (e) {
    return JSON.stringify({ error: "SYSTEM ERROR: " + e.message });
  }
}

/* ======================================================================
   SIABA PULANG AWAL
   ====================================================================== */

function getSiabaPulangFilters() {
  const ID_DB = "1tQsQY1-Ny1ie66GOZPTLtvZ7BiYCgFdNrX-AVGCtaHA"; 
  try {
    const ss = SpreadsheetApp.openById(ID_DB);
    const sheet = ss.getSheetByName("Rekap_Pulang_Awal"); 
    if (!sheet) return JSON.stringify({ years: [] });
    
    const lastRow = sheet.getLastRow();
    if (lastRow < 3) return JSON.stringify({ years: [] });
    
    // Ambil Kolom A (Tahun)
    const data = sheet.getRange(3, 1, lastRow - 2, 1).getDisplayValues();
    let years = new Set();
    data.forEach(row => { if (row[0]) years.add(row[0]); });

    return JSON.stringify({
      years: Array.from(years).sort().reverse()
    });
  } catch (e) {
    return JSON.stringify({ error: e.message });
  }
}

function getSiabaPulangData(filterTahun, filterUnit) {
  const ID_DB = "1tQsQY1-Ny1ie66GOZPTLtvZ7BiYCgFdNrX-AVGCtaHA";
  
  try {
    const ss = SpreadsheetApp.openById(ID_DB);
    const sheet = ss.getSheetByName("Rekap_Pulang_Awal"); 
    if (!sheet) return JSON.stringify({ error: "Sheet 'Rekap_Pulang_Awal' tidak ditemukan." });

    const maxCol = sheet.getLastColumn(); 
    const lastRow = sheet.getLastRow();

    // Header (Baris 1 & 2, Mulai Kolom C)
    const headerRange = sheet.getRange(1, 3, 2, maxCol - 2).getDisplayValues();
    const headerTop = headerRange[0]; 
    const headerSub = headerRange[1]; 

    if (lastRow < 3) return JSON.stringify({ error: "Data Kosong" });

    // Ambil Semua Data (Mulai Baris 3)
    const rawData = sheet.getRange(3, 1, lastRow - 2, maxCol).getDisplayValues();
    
    let result = [];
    
    // Logic Filter Server Side (Hanya Tahun, Unit dikirim semua untuk cache)
    for (let i = 0; i < rawData.length; i++) {
        let row = rawData[i];
        
        let rowTahun = String(row[0]).trim(); 
        let rowUnit  = String(row[1]).toUpperCase().trim(); 
        
        if (rowTahun == String(filterTahun).trim()) {
             // Ambil Data Tampilan (Mulai Kolom C / Index 2)
             let rowDisplay = row.slice(2); 
             
             // [PENTING] Sisipkan Unit Kerja di array paling belakang untuk filter Frontend
             rowDisplay.push(rowUnit);
             
             result.push(rowDisplay);
        }
    }

    // Sorting (Berdasarkan Kolom Total - Descending)
    if (result.length > 0) {
        result.sort((a, b) => {
            // Kolom Total adalah index kedua dari belakang (karena index terakhir adalah Unit)
            let idxTotal = a.length - 2; 
            let valA = parseInt(String(a[idxTotal]).replace(/\./g,'')) || 0;
            let valB = parseInt(String(b[idxTotal]).replace(/\./g,'')) || 0;
            return valB - valA; 
        });
    }

    return JSON.stringify({
      headerTop: headerTop,
      headerSub: headerSub,
      rows: result
    });

  } catch (e) {
    return JSON.stringify({ error: "SYSTEM ERROR: " + e.message });
  }
}

/* ======================================================================
   SIABA UNDUH REKAP
   ====================================================================== */

function getSiabaUnduhData(folderId) {
  const ROOT_ID = "1MoGuseJNrOIMnkZNoqkKcK282jZpUkAm"; 
  let targetId = folderId || ROOT_ID;
  let folder;

  try {
    folder = DriveApp.getFolderById(targetId);
  } catch(e) {
    return { error: "Folder tidak ditemukan atau akses ditolak." };
  }

  let parentId = null;
  let isRoot = (targetId === ROOT_ID);
  
  if (!isRoot) {
    let parents = folder.getParents();
    if (parents.hasNext()) parentId = parents.next().getId();
  }

  let items = [];

  let subfolders = folder.getFolders();
  while (subfolders.hasNext()) {
     let f = subfolders.next();
     items.push({
       id: f.getId(),
       name: f.getName(),
       type: 'folder',
       mimeType: 'application/vnd.google-apps.folder',
       size: '-',
       date: Utilities.formatDate(f.getLastUpdated(), Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm"),
       url: f.getUrl()
     });
  }

  let files = folder.getFiles();
  while (files.hasNext()) {
     let f = files.next();
     let size = (f.getSize() / 1024).toFixed(0) + " KB";
     if (f.getSize() > 1024 * 1024) size = (f.getSize() / (1024*1024)).toFixed(1) + " MB";

     items.push({
       id: f.getId(),
       name: f.getName(),
       type: 'file',
       mimeType: f.getMimeType(),
       size: size,
       date: Utilities.formatDate(f.getLastUpdated(), Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm"),
       url: f.getUrl()
     });
  }

  const URUTAN_BULAN = [
      "JANUARI", "FEBRUARI", "MARET", "APRIL", "MEI", "JUNI", 
      "JULI", "AGUSTUS", "SEPTEMBER", "OKTOBER", "NOVEMBER", "DESEMBER"
  ];

  items.sort((a, b) => {
     if (a.type !== b.type) {
         return a.type === 'folder' ? -1 : 1;
     }

     let nameA = a.name.toUpperCase().trim();
     let nameB = b.name.toUpperCase().trim();

     let idxA = URUTAN_BULAN.indexOf(nameA);
     let idxB = URUTAN_BULAN.indexOf(nameB);

     if (idxA > -1 && idxB > -1) {
         return idxA - idxB;
     }

     return nameA.localeCompare(nameB, undefined, {numeric: true, sensitivity: 'base'});
  });

  return {
    currentId: targetId,
    currentName: folder.getName(),
    parentId: parentId,
    isRoot: isRoot,
    items: items
  };
}

/* ======================================================================
   SIABA SALAH PRESENSI 
   ====================================================================== */

function getSiabaSalahFilters() {
  const ID_DB = "1TZGrMiTuyvh2Xbo44RhJuWlQnOC5LzClsgIoNKtRFkY";
  try {
    const ss = SpreadsheetApp.openById(ID_DB);
    const sheet = ss.getSheetByName("Salah_Absen");
    if (!sheet || sheet.getLastRow() < 2) return JSON.stringify({ years: [], units: [], statuses: [] });
    
    const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 9).getDisplayValues();
    
    let years = new Set();
    let units = new Set();
    let statuses = new Set();
    
    years.add(String(new Date().getFullYear()));

    data.forEach(row => {
      let unit = row[0];
      let tgl = row[3];
      let stat = row[8]; 

      if (unit) units.add(unit.toUpperCase().trim());
      if (stat) statuses.add(stat.trim());
      
      if (tgl) {
        let match = tgl.match(/\d{4}/);
        if (match) years.add(match[0]);
      }
    });

    return JSON.stringify({
      years: Array.from(years).sort().reverse(),
      units: Array.from(units).sort(),
      statuses: Array.from(statuses).sort()
    });
  } catch (e) {
    return JSON.stringify({ error: e.message });
  }
}

function getSiabaSalahData(filterTahun, filterBulan, filterUnit, filterStatus) {
  const ID_DB = "1TZGrMiTuyvh2Xbo44RhJuWlQnOC5LzClsgIoNKtRFkY";
  
  try {
    const ss = SpreadsheetApp.openById(ID_DB);
    const sheet = ss.getSheetByName("Salah_Absen");
    if (!sheet) return JSON.stringify({ error: "Sheet Salah_Absen tidak ditemukan." });

    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return JSON.stringify({ rows: [] });

    const data = sheet.getRange(2, 1, lastRow - 1, 14).getDisplayValues();
    let result = [];
    
    let fTahun = String(filterTahun);
    let fUnit = filterUnit ? String(filterUnit).toUpperCase().trim() : "SEMUA";
    let fStatus = filterStatus ? String(filterStatus).trim() : "SEMUA";
    
    const MAP_BULAN = {"Januari":0,"Februari":1,"Maret":2,"April":3,"Mei":4,"Juni":5,"Juli":6,"Agustus":7,"September":8,"Oktober":9,"November":10,"Desember":11};
    let fBulanIdx = (filterBulan && filterBulan !== "SEMUA") ? MAP_BULAN[filterBulan] : -1;

    for (let i = 0; i < data.length; i++) {
        let row = data[i];
        
        let rowUnit = String(row[0]).toUpperCase().trim();
        let rowTglStr = row[3];
        let rowStatus = String(row[8]).trim(); 
        
        let rowThn = "";
        let rowBln = -1;
        let d = new Date(rowTglStr);
        if(isNaN(d.getTime())) {
           let parts = rowTglStr.split('/'); 
           if(parts.length === 3) d = new Date(parts[2], parts[1]-1, parts[0]);
        }
        if (!isNaN(d.getTime())) {
            rowThn = String(d.getFullYear());
            rowBln = d.getMonth();
        }

        let matchTahun = (rowThn === fTahun);
        let matchBulan = (fBulanIdx === -1 || rowBln === fBulanIdx);
        let matchUnit  = (fUnit === "SEMUA" || rowUnit === fUnit);
        let matchStatus = (fStatus === "SEMUA" || rowStatus === fStatus);

        if (matchTahun && matchBulan && matchUnit && matchStatus) {
            row.push(i + 2); 
            result.push(row);
        }
    }

    result.reverse(); 
    return JSON.stringify({ rows: result });

  } catch (e) {
    return JSON.stringify({ error: "SYSTEM ERROR: " + e.message });
  }
}

function getDaftarSalahPresensi(tahun, bulan) {
  var ID_DB = "1TZGrMiTuyvh2Xbo44RhJuWlQnOC5LzClsgIoNKtRFkY"; 
  var SHEET_NAME = "Salah_Presensi"; 

  try {
    var ss = SpreadsheetApp.openById(ID_DB);
    var sheet = ss.getSheetByName(SHEET_NAME);
    
    if (!sheet) return JSON.stringify([]);

    var data = sheet.getDataRange().getDisplayValues(); 
    var result = [];

    // Filter Tahun & Bulan
    var fTahun  = (tahun) ? String(tahun).trim() : "";
    var mapBulan = {
        "Januari": "01", "Februari": "02", "Maret": "03", "April": "04",
        "Mei": "05", "Juni": "06", "Juli": "07", "Agustus": "08",
        "September": "09", "Oktober": "10", "November": "11", "Desember": "12"
    };
    var fBulanAngka = mapBulan[bulan] || ""; 

    // Loop dari BAWAH ke ATAS (Data terbaru di atas)
    for (var i = data.length - 1; i >= 1; i--) {
      var row = data[i];
      
      // Validasi: Skip jika Nama (Kolom B / Index 1) kosong
      if (!row[1]) continue; 

      var txtTgl = String(row[3]); // Kolom D: Tanggal

      // --- LOGIC FILTER (Tahun & Bulan) ---
      // Format Tanggal di Log Anda: "02-01-2026" (dd-mm-yyyy)
      
      if (fTahun !== "") {
          if (txtTgl.indexOf(fTahun) === -1) continue;
      }
      
      if (fBulanAngka !== "") {
          var cekStrip = "-" + fBulanAngka + "-"; // misal "-01-"
          var cekSlash = "/" + fBulanAngka + "/"; // misal "/01/"
          if (txtTgl.indexOf(cekStrip) === -1 && txtTgl.indexOf(cekSlash) === -1) continue;
      }

      // === MAPPING DATA (SESUAI LOG HEADER ANDA) ===
      /*
         0: Unit Kerja
         1: Nama ASN
         2: NIP
         3: Tanggal
         4: Jam
         5: Jenis
         6: Tgl Pengajuan
         7: User Input
         8: Status
         9: Keterangan (Alasan)
         10: Tgl Edit
         11: User Edit
         12: Tgl Verif
         13: Admin Verif
      */

      result.push({
        rowBaris: i + 1,
        
        unit:     row[0],  // A: Unit
        nama:     row[1],  // B: Nama
        nip:      row[2],  // C: NIP
        tanggal:  row[3],  // D: Tanggal Salah
        jam:      row[4],  // E: Jam Salah
        jenis:    row[5],  // F: Jenis (Datang/Pulang)
        
        tglKirim: row[6],  // G: Tgl Pengajuan (Index 6)
        userInput:row[7],  // H: User Input (Index 7)
        
        status:   row[8],  // I: Status (Index 8)
        ket:      row[9],  // J: Keterangan/Alasan (Index 9) - INI YANG PENTING
        
        tglEdit:    row[10], // K
        userEdit:   row[11], // L
        tglVerif:   row[12], // M
        adminVerif: row[13]  // N
      });
    }
    
    return JSON.stringify(result);

  } catch (e) {
    // Return Error agar terlihat di console
    return JSON.stringify([]);
  }
}

/* =========================
   FUNGSI DATABASE PEGAWAI
   ========================= */

function getDatabasePegawai() {
  // ID Spreadsheet Sumber Baru
  var ID_DB = "160IjN8aiDAgDYXjgDLStS4nCZLKn3Ny-dq3BOFAfDrU"; 
  var SHEET_NAME = "Database";

  try {
    var ss = SpreadsheetApp.openById(ID_DB);
    var sheet = ss.getSheetByName(SHEET_NAME);
    var data = sheet.getDataRange().getDisplayValues();
    var result = [];

    // Loop mulai baris ke-2 (Index 1) untuk melewati Header
    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      
      // Validasi sederhana: Skip jika NIP (Col B) atau Nama (Col C) kosong
      if (!row[1] || !row[2]) continue; 

      result.push({
        unit: String(row[0]).trim(), // Kolom A: Unit Kerja
        nip:  String(row[1]).trim(), // Kolom B: NIP
        nama: String(row[2]).trim()  // Kolom C: Nama ASN
      });
    }
    
    // Opsional: Urutkan berdasarkan Nama (A-Z) agar rapi di dropdown
    result.sort(function(a, b) {
      var nA = a.nama.toUpperCase();
      var nB = b.nama.toUpperCase();
      return (nA < nB) ? -1 : (nA > nB) ? 1 : 0;
    });

    return result;

  } catch (e) {
    Logger.log("Error getDatabasePegawai: " + e.toString());
    return [];
  }
}

/* --- UPDATE FUNGSI SIMPAN (AGAR ALASAN TERSIMPAN) --- */
function simpanSalahAbsen(form) {
  const ID_DB = "1TZGrMiTuyvh2Xbo44RhJuWlQnOC5LzClsgIoNKtRFkY";

  try {
    var ss = SpreadsheetApp.openById(ID_DB);
    var sheet = ss.getSheetByName("Salah_Presensi"); // <-- PASTIKAN NAMA SHEET BENAR
    
    if (!sheet) throw new Error("Sheet 'Salah_Presensi' tidak ditemukan!");
    
    // Format Tanggal (yyyy-mm-dd -> dd-mm-yyyy)
    var tglSimpan = "";
    if (form.tanggal) {
       var parts = form.tanggal.split('-'); 
       tglSimpan = parts[2] + '-' + parts[1] + '-' + parts[0]; 
    }
    
    var jamSimpan = String(form.waktu); 
    var namaUser = form.user_login || "Guest";
    var tglKirim = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm:ss");
    
    // Perbaikan: Menambahkan Alasan di Kolom J (Index 9)
    // Urutan: A, B, C, D, E, F, G, H, I, J
    var barisBaru = [
      form.unit_kerja, 
      form.nama_asn,   
      "'"+form.nip_asn, 
      tglSimpan,        
      jamSimpan,        
      form.jenis,      
      tglKirim,         
      namaUser,        
      "Diproses",       // Status Awal
      form.alasan || "-" // Keterangan / Alasan
    ];

    sheet.appendRow(barisBaru);
    return "SUKSES: Data berhasil disimpan.";
    
  } catch (e) {
    throw new Error("Gagal simpan: " + e.message);
  }
}

/* --- UPDATE FUNGSI UPDATE (AGAR ALASAN & STATUS TERUPDATE) --- */
function updateSalahAbsen(form) {
  const ID_DB = "1TZGrMiTuyvh2Xbo44RhJuWlQnOC5LzClsgIoNKtRFkY"; 
  try {
    var ss = SpreadsheetApp.openById(ID_DB);
    var sheet = ss.getSheetByName("Salah_Presensi");
    if (!sheet) throw new Error("Sheet Salah_Presensi tidak ditemukan");

    var targetNip = String(form.nip_lama).trim();
    var targetTgl = String(form.tgl_lama).trim();
    var targetJam = String(form.jam_lama).trim();

    var data = sheet.getDataRange().getDisplayValues();
    var barisKetemu = -1;
    var statusLama = "";

    for (var i = 1; i < data.length; i++) {
       var sheetNip = String(data[i][2]).trim();
       var sheetTgl = String(data[i][3]).trim();
       var sheetJam = String(data[i][4]).trim(); // Compare string langsung

       if (sheetNip === targetNip && sheetTgl === targetTgl && sheetJam === targetJam) {
          barisKetemu = i + 1;
          statusLama = String(data[i][8]).trim(); 
          break;
       }
    }

    if (barisKetemu === -1) throw new Error("Data asli tidak ditemukan.");

    // Cek Status (Lock)
    if (statusLama.toLowerCase().includes("ok")) return "Gagal: Data sudah Disetujui.";

    // Logic Status Baru
    var statusBaru = "Diproses";
    if (statusLama.toLowerCase().includes("tolak")) statusBaru = "Revisi";

    // Format Data Baru
    var tglBaru = "";
    if (form.tanggal) {
       var p = form.tanggal.split('-'); 
       tglBaru = p[2] + '-' + p[1] + '-' + p[0];
    }
    
    var tglEdit = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm:ss");

    // Update Cell
    sheet.getRange(barisKetemu, 4).setValue("'" + tglBaru);        // Tanggal
    sheet.getRange(barisKetemu, 5).setValue("'" + form.waktu);     // Jam
    sheet.getRange(barisKetemu, 6).setValue(form.jenis);           // Jenis
    sheet.getRange(barisKetemu, 9).setValue(statusBaru);           // Status
    sheet.getRange(barisKetemu, 10).setValue(form.alasan || "-");  // Alasan (Kolom J)
    sheet.getRange(barisKetemu, 11).setValue("'" + tglEdit);       // Tgl Edit
    sheet.getRange(barisKetemu, 12).setValue("'" + form.user_login); // User Edit

    return "SUKSES: Data diperbarui.";
  } catch (e) {
    throw new Error(e.message);
  }
}

/* --- FUNGSI SOFT DELETE (REVISI: FORMAT TETAP UTUH) --- */
function softDeleteSalahAbsen(dataKirim) {
  const ID_DB = "1TZGrMiTuyvh2Xbo44RhJuWlQnOC5LzClsgIoNKtRFkY";
  try {
    var ss = SpreadsheetApp.openById(ID_DB);
    var sheetSource = ss.getSheetByName("Salah_Absen");
    var sheetTrash = ss.getSheetByName("Trash");
    if (!sheetTrash) sheetTrash = ss.insertSheet("Trash");
    if (!sheetSource) throw new Error("Sheet Salah_Absen tidak ditemukan");

    // 1. DATA TARGET (Kunci Pencarian)
    var targetNip = String(dataKirim.nip).trim();
    var targetTgl = String(dataKirim.tgl).trim();
    var targetJam = String(dataKirim.jam).trim();

    // 2. CARI BARIS (Pakai Display Values agar akurat)
    var data = sheetSource.getDataRange().getDisplayValues();
    var barisKetemu = -1;

    for (var i = 1; i < data.length; i++) {
       var sheetNip = String(data[i][2]).trim();
       var sheetTgl = String(data[i][3]).trim().replace(/\//g, '-');
       
       // Normalisasi Jam Sheet (misal 7:15 -> 07:15)
       var sheetJam = String(data[i][4]).trim().replace(/'/g, "").substring(0, 5);
       if (/^\d:\d{2}/.test(sheetJam)) sheetJam = "0" + sheetJam;

       if (sheetNip === targetNip && sheetTgl === targetTgl && sheetJam === targetJam) {
          barisKetemu = i + 1;
          break;
       }
    }

    if (barisKetemu === -1) {
      throw new Error("Data tidak ditemukan (Cek kecocokan Jam/Tanggal).");
    }

    // 3. AMBIL DATA ASLI SEBAGAI TEKS (PENTING: getDisplayValues)
    // Agar format tanggal/jam terambil persis seperti yang tertulis di sel
    var rowRange = sheetSource.getRange(barisKetemu, 1, 1, 14);
    var rowValues = rowRange.getDisplayValues()[0]; 

    // 4. KUNCI FORMAT DENGAN PETIK SATU (')
    // Kita paksa kolom sensitif menjadi string agar tidak berubah format di Trash
    rowValues[2] = "'" + rowValues[2]; // NIP
    rowValues[3] = "'" + rowValues[3]; // Tanggal
    rowValues[4] = "'" + rowValues[4]; // Jam
    rowValues[6] = "'" + rowValues[6]; // Tgl Kirim

    // 5. SIAPKAN METADATA HAPUS
    var tglHapus = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd-MM-yyyy HH:mm:ss");
    var userHapus = "Guest";
    try {
        var cu = getCurrentUser();
        if (cu && cu.fullName) userHapus = cu.fullName;
        else if (dataKirim.user) userHapus = dataKirim.user;
    } catch(e) { 
        userHapus = dataKirim.user || "Guest"; 
    }

    var alasan = dataKirim.alasan || "-";
    
    // Gabungkan data asli + metadata hapus
    var trashRow = rowValues.concat([tglHapus, userHapus, alasan]);

    // 6. PINDAHKAN
    sheetTrash.appendRow(trashRow); 
    sheetSource.deleteRow(barisKetemu);

    return "Sukses";
  } catch (e) {
    throw new Error(e.message);
  }
}

/* FUNGSI VERIFIKASI (FIX NAMA SHEET) */
function processVerifikasiSalahAbsen(dataKirim) {
  const ID_DB = "1TZGrMiTuyvh2Xbo44RhJuWlQnOC5LzClsgIoNKtRFkY";
  const SHEET_NAME = "Salah_Presensi"; // <--- SUDAH DIGANTI JADI Salah_Presensi

  try {
    var ss = SpreadsheetApp.openById(ID_DB);
    var sheet = ss.getSheetByName(SHEET_NAME);
    
    // Validasi Ekstra: Cek nama sheet
    if (!sheet) {
      throw new Error("Sheet '" + SHEET_NAME + "' tidak ditemukan! Cek nama tab di Spreadsheet.");
    }

    // Kita gunakan RecId (Nomor Baris)
    var rowIndex = parseInt(dataKirim.recId);
    
    // Validasi Baris Data: Cek apakah baris itu ada isinya
    var cekData = sheet.getRange(rowIndex, 1).getValue(); 
    if (!cekData) throw new Error("Baris data ke-" + rowIndex + " kosong/tidak ditemukan.");

    // Update Status (Kolom I / Index 9) -> Kolom ke-9
    sheet.getRange(rowIndex, 9).setValue(dataKirim.status);
    
    // Update Catatan (Kolom J / Index 10) -> Kolom ke-10
    // Gunakan petik satu (') untuk mengunci format teks
    sheet.getRange(rowIndex, 10).setValue("'" + dataKirim.ket);
    
    // Metadata Verifikasi
    var tglVerif = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd-MM-yyyy HH:mm:ss");
    
    // Update Tgl Verif (Kolom M / Index 13) -> Kolom ke-13
    sheet.getRange(rowIndex, 13).setValue("'" + tglVerif);
    
    // Update Admin Verif (Kolom N / Index 14) -> Kolom ke-14
    sheet.getRange(rowIndex, 14).setValue("'" + dataKirim.admin);

    return "Sukses Verifikasi";
  } catch (e) {
    throw new Error(e.message);
  }
}

function getDaftarSampah() {
  const ID_DB = "1TZGrMiTuyvh2Xbo44RhJuWlQnOC5LzClsgIoNKtRFkY";
  const SHEET_TRASH = "Sampah_Salah";
  try {
    var ss = SpreadsheetApp.openById(ID_DB);
    var sheet = ss.getSheetByName(SHEET_TRASH);
    if (!sheet || sheet.getLastRow() < 2) return [];

    var data = sheet.getDataRange().getDisplayValues();
    var result = [];
    for(var i=1; i<data.length; i++){
        result.push(data[i]);
    }
    return result.reverse(); 
  } catch (e) {
    return [];
  }
}

/* --- FUNGSI RESTORE (REVISI: LOCK FORMAT DARI TRASH KE SOURCE) --- */
function prosesRestoreSalahAbsen(nip, tgl, jam) {
  const ID_DB = "1TZGrMiTuyvh2Xbo44RhJuWlQnOC5LzClsgIoNKtRFkY";
  const SHEET_MAIN = "Salah_Presensi"; // <--- SUDAH DIGANTI
  const SHEET_TRASH = "Sampah_Salah";

  try {
    var ss = SpreadsheetApp.openById(ID_DB);
    var sheetTrash = ss.getSheetByName(SHEET_TRASH);
    var sheetSource = ss.getSheetByName(SHEET_MAIN);
    
    if (!sheetTrash || !sheetSource) throw new Error("Sheet Database tidak ditemukan.");

    var dataDisplay = sheetTrash.getDataRange().getDisplayValues();
    var barisKetemu = -1;

    var targetNip = String(nip).trim();
    var targetTgl = String(tgl).trim();
    var targetJam = String(jam).trim();

    for (var i = 1; i < dataDisplay.length; i++) {
       var sheetNip = String(dataDisplay[i][2]).trim();
       var sheetTgl = String(dataDisplay[i][3]).trim().replace(/\//g, '-');
       var sheetJam = String(dataDisplay[i][4]).trim().replace(/'/g, "").substring(0, 5);
       if (/^\d:\d{2}/.test(sheetJam)) sheetJam = "0" + sheetJam;

       if (sheetNip === targetNip && sheetTgl === targetTgl && sheetJam === targetJam) {
          barisKetemu = i + 1;
          break;
       }
    }

    if (barisKetemu === -1) throw new Error("Data tidak ditemukan di Trash.");

    var rowValues = sheetTrash.getRange(barisKetemu, 1, 1, 14).getDisplayValues()[0];

    rowValues[2] = "'" + rowValues[2]; 
    rowValues[3] = "'" + rowValues[3]; 
    rowValues[4] = "'" + rowValues[4]; 
    
    sheetSource.appendRow(rowValues);
    sheetTrash.deleteRow(barisKetemu);

    return "Sukses Restore Data";
  } catch (e) {
    throw new Error(e.message);
  }
}

/* ======================================================================
   MODUL LUPA PRESENSI (VERSI LOKAL - LEBIH STABIL)
   Semua ID didefinisikan di dalam fungsi agar tidak ada konflik global.
   ====================================================================== */

function cekBentrokLupa(nipBaru, tglBaruStr, jenisBaru, rowIdPengecualian) {
  var ID_DB = "160IjN8aiDAgDYXjgDLStS4nCZLKn3Ny-dq3BOFAfDrU";
  var SHEET_NAME = "Lupa_Presensi";
  
  var ss = SpreadsheetApp.openById(ID_DB);
  var sheet = ss.getSheetByName(SHEET_NAME);
  // Ambil data kolom C (NIP), D (Tgl), F (Jenis), K (Status) untuk efisiensi
  // Tapi getDataRange().getValues() lebih mudah dibaca
  var data = sheet.getDataRange().getValues();
  
  // 1. STANDARISASI INPUT BARU -> YYYY-MM-DD
  // Contoh: '2026-02-12' tetap '2026-02-12'
  var tglBaruYMD = normalizeToYMD(tglBaruStr);
  var jenisBaruClean = String(jenisBaru).trim().toLowerCase();

  // Loop Data (Mulai baris 2 / Index 1)
  for (var i = 1; i < data.length; i++) {
    // Skip jika sedang edit baris ini sendiri
    if (rowIdPengecualian && (i + 1) == rowIdPengecualian) continue;

    var rowNip = String(data[i][2]).replace(/'/g, "").trim(); 
    var rowStatus = String(data[i][10]).toLowerCase();
    
    // Cek NIP & Status Aktif (Bukan Ditolak)
    if (rowNip === String(nipBaru).trim() && !rowStatus.includes("tolak")) {
       
       // 2. STANDARISASI TANGGAL DATABASE -> YYYY-MM-DD
       // Contoh: '12-02-2026' akan diubah jadi '2026-02-12'
       var rowTglRaw = data[i][3];
       var rowTglYMD = normalizeToYMD(rowTglRaw);
       
       var rowJenis = String(data[i][5]).trim().toLowerCase();
       
       // 3. BANDINGKAN FORMAT YANG SUDAH SAMA
       if (rowTglYMD === tglBaruYMD && rowJenis === jenisBaruClean) {
           var tglDisplay = String(rowTglRaw).replace(/'/g,""); // Tampilkan tanggal asli biar user paham
           return "Gagal: Data ganda! Anda sudah mengajukan Lupa Presensi (" + data[i][5] + ") pada tanggal " + tglDisplay + ".";
       }
    }
  }
  return null; // Aman
}

// --- HELPER BARU: PENERJEMAH TANGGAL ---
function normalizeToYMD(val) {
  if (!val) return "";
  
  // 1. Jika Data Asli Excel (Date Object) -> Ubah ke YYYY-MM-DD
  if (val instanceof Date) {
    return Utilities.formatDate(val, Session.getScriptTimeZone(), "yyyy-MM-dd");
  }
  
  var s = String(val).replace(/'/g, "").trim();
  
  // 2. Jika format sudah YYYY-MM-DD (Input HTML) -> Biarkan
  if (s.match(/^\d{4}-\d{2}-\d{2}$/)) {
    return s;
  }
  
  // 3. Jika format DD-MM-YYYY atau DD/MM/YYYY (Format Indo) -> Balik jadi YYYY-MM-DD
  var parts = s.split(/[-/]/); // Pisahkan berdasarkan - atau /
  if (parts.length === 3) {
     // Asumsi: Bagian pertama adalah Tanggal (DD), Ketiga adalah Tahun (YYYY)
     // Cek panjang string: DD (2 digit), YYYY (4 digit)
     if (parts[0].length <= 2 && parts[2].length === 4) {
        var dd = parts[0].padStart(2, '0');
        var mm = parts[1].padStart(2, '0');
        var yyyy = parts[2];
        return yyyy + "-" + mm + "-" + dd;
     }
  }
  
  return s; // Kembalikan apa adanya jika format aneh
}

// 1. FILTER TAHUN
function getLupaFilters() {
  var ID_DB = "160IjN8aiDAgDYXjgDLStS4nCZLKn3Ny-dq3BOFAfDrU"; // ID Spreadsheet
  var SHEET_NAME = "Lupa_Presensi";

  try {
    var ss = SpreadsheetApp.openById(ID_DB);
    var sheet = ss.getSheetByName(SHEET_NAME);
    if (!sheet || sheet.getLastRow() < 2) return JSON.stringify({ years: [] });
    
    var data = sheet.getRange(2, 4, sheet.getLastRow() - 1, 1).getDisplayValues();
    var years = new Set();
    
    data.forEach(function(r) {
      var match = String(r[0]).match(/\d{4}/);
      if(match) years.add(match[0]);
    });
    
    return JSON.stringify({ years: Array.from(years).sort().reverse() });
  } catch (e) { return JSON.stringify({ error: e.message }); }
}

function getDaftarLupaPresensi(tahun, bulan) {
  // CATATAN: Parameter Unit & Status DIHAPUS karena filter dilakukan di Browser
  var ID_DB = "160IjN8aiDAgDYXjgDLStS4nCZLKn3Ny-dq3BOFAfDrU"; 
  var SHEET_NAME = "Lupa_Presensi";

  try {
    var ss = SpreadsheetApp.openById(ID_DB);
    var sheet = ss.getSheetByName(SHEET_NAME);
    // Ambil Data dari Bawah ke Atas (Optimasi Sorting Last Activity Alami)
    var data = sheet.getDataRange().getDisplayValues(); 
    var result = [];

    var fTahun  = (tahun) ? String(tahun).trim() : "";
    
    var mapBulan = {
        "Januari": "01", "Februari": "02", "Maret": "03", "April": "04",
        "Mei": "05", "Juni": "06", "Juli": "07", "Agustus": "08",
        "September": "09", "Oktober": "10", "November": "11", "Desember": "12"
    };
    var fBulanAngka = mapBulan[bulan] || ""; 

    // Loop dari baris terbawah (Data Terbaru) ke atas
    for (var i = data.length - 1; i >= 1; i--) {
      var row = data[i];
      if (!row[1]) continue; 

      var txtTgl = String(row[3]); // Kolom D

      // 1. FILTER TAHUN (Wajib di Server)
      if (fTahun !== "" && txtTgl.indexOf(fTahun) === -1) continue;

      // 2. FILTER BULAN (Wajib di Server)
      if (fBulanAngka !== "") {
          var cekStrip = "-" + fBulanAngka + "-";
          var cekSlash = "/" + fBulanAngka + "/";
          if (txtTgl.indexOf(cekStrip) === -1 && txtTgl.indexOf(cekSlash) === -1) continue;
      }

      // 3. MASUKKAN SEMUA DATA (Unit & Status JANGAN difilter disini)
      result.push({
        rowBaris: i + 1,       
        unit: row[0], nama: row[1], nip: row[2],           
        tanggal: row[3], jam: row[4], jenis: row[5], komulatif: row[6],     
        tglKirim: row[7], userInput: row[8], fileUrl: row[9], status: row[10],       
        tglEdit: row[11], userEdit: row[12], tglVerif: row[13], adminVerif: row[14], ket: row[15]           
      });
    }
    
    return JSON.stringify(result);

  } catch (e) {
    return JSON.stringify([]);
  }
}

// 3. GET DATA BY ID (VERSI FINAL BERSIH)
function getLupaById(rowIndex) {
  var ID_DB = "160IjN8aiDAgDYXjgDLStS4nCZLKn3Ny-dq3BOFAfDrU";
  var SHEET_NAME = "Lupa_Presensi";

  if (!rowIndex) return null;
  try {
    var ss = SpreadsheetApp.openById(ID_DB);
    var sheet = ss.getSheetByName(SHEET_NAME);
    var idx = parseInt(rowIndex);
    
    // 1. AMBIL DATA BARIS (A-P) SEBAGAI TEKS
    var range = sheet.getRange(idx, 1, 1, 16);
    var row = range.getDisplayValues()[0]; 

    // 2. EKSTRAKSI LINK FILE (KOLOM J)
    // Teknik ini membaca Smart Chip, Hyperlink Formula, dan Teks Biasa
    var cellJ = sheet.getRange(idx, 10); 
    
    // a. Cek Link Tersembunyi (Smart Chip / Insert Link)
    var richText = cellJ.getRichTextValue();
    var linkRich = richText ? richText.getLinkUrl() : "";
    
    // b. Cek Formula (=HYPERLINK)
    var formula = cellJ.getFormula();
    var linkFormula = "";
    if (formula && formula.includes("http")) {
       var match = formula.match(/"(https?:\/\/[^"]+)"/);
       if (match) linkFormula = match[1];
    }
    
    // c. Cek Teks Biasa (Paste URL mentah)
    var linkText = row[9];

    // Prioritas: Link Rich -> Link Formula -> Link Teks
    var finalUrl = linkRich || linkFormula || linkText;

    return {
      id: idx,
      unit: row[0],  
      nama: row[1],  
      nip: row[2],   
      tanggal: row[3], 
      waktu: row[4],   
      jenis: row[5],   
      komulatif: row[6], 
      fileUrl: finalUrl, // URL bersih
      status: row[10], 
      ket: row[15]     
    };
  } catch (e) { return null; }
}

// 4. VERIFIKASI ADMIN
function processVerifikasiLupaPresensi(dataKirim) {
  var ID_DB = "160IjN8aiDAgDYXjgDLStS4nCZLKn3Ny-dq3BOFAfDrU";
  var SHEET_NAME = "Lupa_Presensi";

  try {
    var ss = SpreadsheetApp.openById(ID_DB);
    var sheet = ss.getSheetByName(SHEET_NAME);
    var baris = parseInt(form.recId);

    // Update Status (Kolom K / Index 11)
    sheet.getRange(baris, 11).setValue(form.status);

    // Update Timestamp Verif (Kolom N / Index 14)
    var now = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd-MM-yyyy HH:mm:ss");
    sheet.getRange(baris, 14).setValue(now);

    // Update Admin Verifikator (Kolom O / Index 15)
    sheet.getRange(baris, 15).setValue(form.user_verif);

    // Update Keterangan (Kolom P / Index 16)
    sheet.getRange(baris, 16).setValue(form.keterangan);

    return "Sukses verifikasi.";
  } catch (e) {
    throw new Error(e.message);
  }
}

// 5. UPDATE DATA (EDIT)
function updateLupaPresensi(form, fileData) {
  var ID_DB = "160IjN8aiDAgDYXjgDLStS4nCZLKn3Ny-dq3BOFAfDrU";
  var SHEET_NAME = "Lupa_Presensi";
  var DRIVE_ID = "1h8LcyYYrdVmd-fDPdcZ47hT9--rLQ7Fa";

  try {
    var baris = parseInt(form.recId);
    var targetNip = form.nip_asn || form.nip_lama; 
    
    // Normalisasi Tanggal
    var tglSimpan = form.tanggal; // yyyy-mm-dd

    // --- VALIDASI BENTROK ---
    var err = cekBentrokLupa(targetNip, tglSimpan, form.jenis, baris);
    if (err) return err;
    // ------------------------

    var ss = SpreadsheetApp.openById(ID_DB);
    var sheet = ss.getSheetByName(SHEET_NAME);
    
    // Ambil Data Lama untuk Cek File
    var rangeLama = sheet.getRange(baris, 1, 1, 16); 
    var valLama = rangeLama.getValues()[0];
    var finalUrl = valLama[9]; 
    
    // ... (Logika file sama seperti sebelumnya) ...
    // Format Nama File Baru
    var namaFileBaru = targetNip + " - " + tglSimpan + " - " + form.jenis + ".pdf";
    var targetFolder = getFolderTahunBulan(DRIVE_ID, tglSimpan);

    if (fileData && fileData.data) {
       // Upload Baru
       var blob = Utilities.newBlob(Utilities.base64Decode(fileData.data), fileData.mimeType, namaFileBaru);
       var newFile = targetFolder.createFile(blob);
       newFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
       finalUrl = newFile.getUrl();
    } else {
       // Rename Lama jika perlu
       var tglLamaSheet = String(valLama[3]).replace(/'/g, "");
       if (tglSimpan !== tglLamaSheet || form.jenis !== valLama[5]) {
           try { 
             var idFile = finalUrl.match(/[-\w]{25,}/);
             if(idFile) {
                 var fileDrive = DriveApp.getFileById(idFile[0]);
                 fileDrive.setName(namaFileBaru);
                 if (tglSimpan !== tglLamaSheet) fileDrive.moveTo(targetFolder);
             }
           } catch(e) {}
       }
    }

    var jamSimpan = form.waktu; // Format HH:mm

    sheet.getRange(baris, 4).setValue("'" + tglSimpan);      
    sheet.getRange(baris, 5).setValue("'" + jamSimpan);      
    sheet.getRange(baris, 6).setValue(form.jenis);   
    sheet.getRange(baris, 7).setValue("'" + form.komulatif); 
    sheet.getRange(baris, 10).setValue(finalUrl);    
    sheet.getRange(baris, 12).setValue(Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd-MM-yyyy HH:mm:ss"));        
    sheet.getRange(baris, 13).setValue(form.user_login);

    return "Sukses Data Berhasil Diupdate";
  } catch(e) { return "Error: " + e.message; }
}

function softDeleteLupaPresensi(dataKirim) {
  var ID_DB = "160IjN8aiDAgDYXjgDLStS4nCZLKn3Ny-dq3BOFAfDrU";
  var SHEET_MAIN = "Lupa_Presensi";
  var SHEET_TRASH = "Trash"; // Pastikan sheet ini ada

  try {
    var ss = SpreadsheetApp.openById(ID_DB);
    var sheetMain = ss.getSheetByName(SHEET_MAIN);
    
    // Cek Sheet Sampah, buat jika belum ada
    var sheetTrash = ss.getSheetByName(SHEET_TRASH);
    if (!sheetTrash) {
       sheetTrash = ss.insertSheet(SHEET_TRASH);
       sheetTrash.appendRow(["Unit","Nama","NIP","Tanggal","Jam","Jenis","Komulatif","Tgl Kirim","User","File","Status","Ket","...","...","...","...","Waktu Hapus","User Hapus","Alasan"]);
    }

    var rowIdx = parseInt(dataKirim.recId);
    if (isNaN(rowIdx)) throw new Error("ID Baris tidak valid.");

    // Cek Kode Hapus (Server Side Validation)
    var now = new Date();
    var validCode = Utilities.formatDate(now, Session.getScriptTimeZone(), "yyyyMMdd");
    if(String(dataKirim.kode).trim() !== validCode) {
       throw new Error("KODE_SALAH"); // Lempar error khusus
    }

    // Ambil Data Baris tsb (Kolom A s.d P / 1 s.d 16)
    var range = sheetMain.getRange(rowIdx, 1, 1, 16);
    var rowValues = range.getValues()[0];

    // Format ulang tanggal/jam biar ada petiknya saat masuk tong sampah (biar format terjaga)
    rowValues[3] = "'" + String(rowValues[3]).replace(/'/g, ""); // Tanggal
    rowValues[4] = "'" + String(rowValues[4]).replace(/'/g, ""); // Jam
    rowValues[6] = "'" + String(rowValues[6]).replace(/'/g, ""); // Komulatif

    // Tambah Metadata Hapus
    var tglHapus = Utilities.formatDate(now, Session.getScriptTimeZone(), "dd-MM-yyyy HH:mm:ss");
    var userHapus = dataKirim.user_login || "Admin";
    var alasan = dataKirim.alasan || "-";

    var trashRow = rowValues.concat([tglHapus, userHapus, alasan]);

    // Pindah ke Sampah & Hapus dari Utama
    sheetTrash.appendRow(trashRow);
    sheetMain.deleteRow(rowIdx);

    return "Sukses";

  } catch (e) {
    if(e.message === "KODE_SALAH") return "KODE_SALAH";
    return "Error Server: " + e.message;
  }
}

// 6. SOFT DELETE (HAPUS KE TRASH)
function softDeleteSalahAbsen(dataKirim) {
  const ID_DB = "1TZGrMiTuyvh2Xbo44RhJuWlQnOC5LzClsgIoNKtRFkY";
  const SHEET_MAIN = "Salah_Presensi"; // <--- SUDAH DIGANTI
  const SHEET_TRASH = "Sampah_Salah";  

  try {
    var ss = SpreadsheetApp.openById(ID_DB);
    var sheetSource = ss.getSheetByName(SHEET_MAIN);
    
    // Buat sheet sampah jika belum ada
    var sheetTrash = ss.getSheetByName(SHEET_TRASH);
    if (!sheetTrash) {
       sheetTrash = ss.insertSheet(SHEET_TRASH);
       sheetTrash.appendRow(["Unit","Nama","NIP","Tanggal","Jam","Jenis","Tgl Ajuan","User","Status","Ket","Edit","UserEdit","Verif","AdminVerif","Waktu Hapus","User Hapus","Alasan Hapus"]);
    }
    
    if (!sheetSource) throw new Error("Sheet '" + SHEET_MAIN + "' tidak ditemukan");

    var rowIndex = parseInt(dataKirim.recId);
    
    // Ambil Data Baris tsb (Kolom A s.d N / 1 s.d 14)
    var rowRange = sheetSource.getRange(rowIndex, 1, 1, 14);
    var rowValues = rowRange.getDisplayValues()[0]; 

    // Kunci format penting
    rowValues[2] = "'" + rowValues[2]; // NIP
    rowValues[3] = "'" + rowValues[3]; // Tanggal
    rowValues[4] = "'" + rowValues[4]; // Jam

    // Siapkan Metadata Hapus
    var tglHapus = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd-MM-yyyy HH:mm:ss");
    var userHapus = dataKirim.user || "Guest";
    var alasan = dataKirim.alasan || "-";
    
    var trashRow = rowValues.concat([tglHapus, userHapus, alasan]);

    sheetTrash.appendRow(trashRow); 
    sheetSource.deleteRow(rowIndex);

    return "Sukses";
  } catch (e) {
    throw new Error(e.message);
  }
}

// 7. SIMPAN DATA BARU
function simpanLupaPresensi(dataKirim) {
  var ID_DB = "160IjN8aiDAgDYXjgDLStS4nCZLKn3Ny-dq3BOFAfDrU";
  var SHEET_NAME = "Lupa_Presensi";
  var DRIVE_ID = "1h8LcyYYrdVmd-fDPdcZ47hT9--rLQ7Fa"; 

  try {
    // Normalisasi Tanggal untuk Validasi & Simpan
    var tglSimpan = "";
    if (dataKirim.tanggal && dataKirim.tanggal.includes("-")) {
       var parts = dataKirim.tanggal.split("-");
       // Input HTML date: yyyy-mm-dd -> Kita simpan yyyy-mm-dd juga biar konsisten
       tglSimpan = parts[0] + "-" + parts[1] + "-" + parts[2]; 
       // KOREKSI: Jika logic sebelumnya membalik tanggal, sesuaikan disini. 
       // Biasanya database spreadsheet lebih aman pakai yyyy-mm-dd atau 'dd-mm-yyyy text.
       // Mari gunakan format text 'yyyy-mm-dd sesuai input HTML agar match stringnya mudah.
    } else { tglSimpan = dataKirim.tanggal; }

    // --- VALIDASI BENTROK ---
    // Kirim tglSimpan yang sudah dinormalisasi
    var err = cekBentrokLupa(dataKirim.nip_asn, tglSimpan, dataKirim.jenis, null);
    if (err) return err; 
    // ------------------------

    var ss = SpreadsheetApp.openById(ID_DB);
    var sheet = ss.getSheetByName(SHEET_NAME);

    // Normalisasi Jam
    var jamSimpan = dataKirim.waktu;
    if (jamSimpan && jamSimpan.includes(":")) {
       var jamParts = jamSimpan.split(":");
       jamSimpan = String(jamParts[0]).padStart(2, '0') + ":" + String(jamParts[1]).padStart(2, '0');
    }

    // Simpan File
    var targetFolder = getFolderTahunBulan(DRIVE_ID, tglSimpan);
    var fileExt = dataKirim.file.name.split('.').pop();
    var fileNameBaru = dataKirim.nip_asn + " - " + tglSimpan + " - " + dataKirim.jenis + "." + fileExt;
    var fileBlob = Utilities.newBlob(Utilities.base64Decode(dataKirim.file.data), dataKirim.file.mimeType, dataKirim.file.name);
    var newFile = targetFolder.createFile(fileBlob).setName(fileNameBaru);
    newFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    var fileUrl = newFile.getUrl();

    var timestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd-MM-yyyy HH:mm:ss");
    
    // Simpan dengan tanda petik agar terbaca TEXT
    var rowData = [
      dataKirim.unit_kerja, dataKirim.nama_asn, dataKirim.nip_asn,
      "'" + tglSimpan, "'" + jamSimpan, dataKirim.jenis, "'" + dataKirim.komulatif,
      timestamp, dataKirim.user_login, fileUrl, "Diproses",
      "", "", "", "", ""
    ];
    sheet.appendRow(rowData);
    return "Sukses Data Berhasil Disimpan";
    
  } catch (e) { return "Error: " + e.message; }
}

// HELPER (Tetap sama, tidak perlu ID DB)
function parseDateForInput(val) {
    if(!val) return "";
    if(val instanceof Date) return Utilities.formatDate(val, Session.getScriptTimeZone(), "yyyy-MM-dd");
    var str = String(val).trim();
    if(str.match(/^\d{2}-\d{2}-\d{4}/)) { var p = str.split('-'); return p[2]+"-"+p[1]+"-"+p[0]; }
    return str.substring(0,10);
}
function parseTimeForInput(val) {
    if(!val) return "";
    if(val instanceof Date) return Utilities.formatDate(val, Session.getScriptTimeZone(), "HH:mm");
    return String(val).substring(0,5);
}

/* ======================================================================
   FITUR SAMPAH & RESTORE (KHUSUS LUPA PRESENSI)
   ====================================================================== */

// 1. LIHAT DAFTAR SAMPAH
function getDaftarSampahLupa() {
  var ID_DB = "160IjN8aiDAgDYXjgDLStS4nCZLKn3Ny-dq3BOFAfDrU";
  var SHEET_TRASH = "Trash";
  
  try {
    var ss = SpreadsheetApp.openById(ID_DB);
    var sheet = ss.getSheetByName(SHEET_TRASH);
    if (!sheet) return [];
    
    // Ambil semua data (Termasuk kolom Q, R, S)
    var data = sheet.getDataRange().getDisplayValues();
    var result = [];
    
    // Loop mulai baris 1 (lewati header)
    for (var i = 1; i < data.length; i++) {
       var row = data[i];
       // Pastikan baris memiliki minimal 19 kolom (A sampai S)
       if (row.length < 19) continue;

       result.push({
         // Data Utama untuk identifikasi
         nip: row[2],
         nama: row[1],
         tgl: row[3].replace(/'/g, ""), // Bersihkan petik untuk tampilan
         jam: row[4].replace(/'/g, ""),
         
         // Data Tambahan (Metadata Hapus)
         tglHapus: row[16], // Kolom Q (Index 16)
         userDel: row[17],  // Kolom R (Index 17)
         alasan: row[18]    // Kolom S (Index 18)
       });
    }
    // Urutkan dari yang baru dihapus (paling bawah di sheet trash)
    return result.reverse();
    
  } catch (e) { return []; }
}

// 2. RESTORE DATA (KEMBALIKAN KE UTAMA)
function restoreLupaPresensi(nip, tgl, jam) {
  var ID_DB = "160IjN8aiDAgDYXjgDLStS4nCZLKn3Ny-dq3BOFAfDrU";
  var SHEET_MAIN = "Lupa_Presensi";
  var SHEET_TRASH = "Trash";
  // Masukkan ID Folder Utama jika ingin file dikembalikan ke folder asal (Opsional)
  var MAIN_FOLDER_ID = "1h8LcyYYrdVmd-fDPdcZ47hT9--rLQ7Fa"; 

  try {
    var ss = SpreadsheetApp.openById(ID_DB);
    var sheetTrash = ss.getSheetByName(SHEET_TRASH);
    var sheetMain = ss.getSheetByName(SHEET_MAIN);
    
    var dataDisplay = sheetTrash.getDataRange().getDisplayValues();
    var barisKetemu = -1;
    var rowDataFull = []; // Ini berisi 19 Kolom

    // Bersihkan parameter pencarian
    var fNip = String(nip).trim();
    var fTgl = String(tgl).replace(/'/g, "").trim(); 
    var fJam = String(jam).replace(/'/g, "").trim();

    // Loop cari data di Trash
    for (var i = 1; i < dataDisplay.length; i++) {
       var rNip = String(dataDisplay[i][2]).trim();
       var rTgl = String(dataDisplay[i][3]).replace(/'/g, "").trim();
       var rJam = String(dataDisplay[i][4]).replace(/'/g, "").trim();

       if (rNip === fNip && rTgl === fTgl && rJam === fJam) {
          barisKetemu = i + 1; 
          rowDataFull = dataDisplay[i];
          break;
       }
    }

    if (barisKetemu === -1) throw new Error("Data tidak ditemukan diampah.");

    // 1. POTONG DATA (AMBIL 16 KOLOM PERTAMA SAJA)
    // Kita buang kolom Q, R, S (Index 16, 17, 18)
    var rowRestore = rowDataFull.slice(0, 16); 

    // 2. FORCE STRING ULANG (Safety)
    // Pastikan Tgl & Jam tetap ada tanda petik
    if(!rowRestore[3].startsWith("'")) rowRestore[3] = "'" + rowRestore[3];
    if(!rowRestore[4].startsWith("'")) rowRestore[4] = "'" + rowRestore[4];
    if(!rowRestore[6].startsWith("'")) rowRestore[6] = "'" + rowRestore[6];

    // 3. KEMBALIKAN FILE (Opsional: Pindahkan balik ke folder utama)
    var fileUrl = rowRestore[9];
    if (fileUrl && String(fileUrl).includes("drive") && MAIN_FOLDER_ID) {
        try {
            var fid = fileUrl.match(/[-\w]{25,}/);
            if(fid) DriveApp.getFileById(fid[0]).moveTo(DriveApp.getFolderById(MAIN_FOLDER_ID));
        } catch(e){}
    }

    // 4. SIMPAN KE UTAMA & HAPUS DARI TRASH
    sheetMain.appendRow(rowRestore);
    sheetTrash.deleteRow(barisKetemu);

    return "Sukses Data Berhasil Dipulihkan";
  } catch (e) { throw new Error(e.message); }
}

function verifikasiLupaPresensi(form) {
  var ID_DB = "160IjN8aiDAgDYXjgDLStS4nCZLKn3Ny-dq3BOFAfDrU"; // ID Spreadsheet Anda
  var SHEET_NAME = "Lupa_Presensi";

  try {
    var ss = SpreadsheetApp.openById(ID_DB);
    var sheet = ss.getSheetByName(SHEET_NAME);
    
    // Pastikan ID Baris Valid
    var baris = parseInt(form.recId);
    if (isNaN(baris) || baris < 2) throw new Error("ID Baris tidak valid.");

    // 1. Update Status (Kolom K / Index 11)
    sheet.getRange(baris, 11).setValue(form.status);

    // 2. Update Tanggal Verif (Kolom N / Index 14)
    var now = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd-MM-yyyy HH:mm:ss");
    sheet.getRange(baris, 14).setValue(now);

    // 3. Update Admin Verifikator (Kolom O / Index 15)
    sheet.getRange(baris, 15).setValue(form.user_verif);

    // 4. Update Keterangan (Kolom P / Index 16)
    sheet.getRange(baris, 16).setValue(form.keterangan);

    return "Sukses data berhasil diverifikasi.";
    
  } catch (e) {
    throw new Error("Gagal Verifikasi: " + e.message);
  }
}

/* =================================================================
   FUNGSI PERBAIKAN MASSAL (JALANKAN SEKALI SAJA DARI EDITOR)
   Fungsinya: Mengubah semua file lama menjadi Public (bisa dipreview)
   ================================================================= */
function fixPerizinanFileLama() {
  // ID Folder "Lupa Presensi" Anda (diambil dari kode sebelumnya)
  var FOLDER_ID = "1h8LcyYYrdVmd-fDPdcZ47hT9--rLQ7Fa"; 
  
  try {
    var folder = DriveApp.getFolderById(FOLDER_ID);
    var files = folder.getFiles();
    var count = 0;
    
    console.log("Mulai memperbaiki izin file...");
    
    while (files.hasNext()) {
      var file = files.next();
      // Set Permission menjadi: Anyone with Link -> Viewer
      file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
      count++;
    }
    
    console.log("SUKSES! Berhasil memperbaiki " + count + " file.");
    return "Selesai. " + count + " file sekarang sudah bisa dilihat.";
    
  } catch (e) {
    console.error("Gagal: " + e.message);
    return "Eror: " + e.message;
  }
}

/* =================================================================
   FUNGSI ONE-TIME FIX: MEMBUKA KUNCI SEMUA FILE LAMA
   Cara Pakai: 
   1. Pilih fungsi 'bukaKunciSemuaFile' di toolbar atas.
   2. Klik tombol Run (Segitiga).
   3. Tunggu sampai log "Selesai" muncul.
   ================================================================= */
function bukaKunciSemuaFile() {
  // ID FOLDER LUPA PRESENSI (Pastikan ID ini benar sesuai folder Anda)
  var FOLDER_ID = "1h8LcyYYrdVmd-fDPdcZ47hT9--rLQ7Fa"; 
  
  try {
    var folder = DriveApp.getFolderById(FOLDER_ID);
    var files = folder.getFiles();
    var hitung = 0;
    
    console.log("Memulai proses buka kunci file...");
    
    while (files.hasNext()) {
      var file = files.next();
      // Ubah jadi Public (Viewer) agar bisa tampil di Iframe
      file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
      hitung++;
    }
    
    console.log("SUKSES! " + hitung + " file telah dibuka kuncinya.");
    return "Selesai. " + hitung + " file diperbaiki.";
    
  } catch (e) {
    console.error("Gagal: " + e.message);
    return "Error: " + e.message;
  }
}

function getFolderTahunBulan(parentId, strTgl) {
  // strTgl format: "dd-mm-yyyy" (Misal: 17-01-2026)
  var parts = strTgl.split("-");
  var year = parts[2];
  var monthIdx = parseInt(parts[1], 10) - 1; // 0-11
  
  var arrBulan = ["Januari", "Februari", "Maret", "April", "Mei", "Juni", 
                  "Juli", "Agustus", "September", "Oktober", "November", "Desember"];
  var monthName = arrBulan[monthIdx];
  
  // 1. Cek Folder Tahun
  var parentFolder = DriveApp.getFolderById(parentId);
  var yearFolder;
  var yearIter = parentFolder.getFoldersByName(year);
  if (yearIter.hasNext()) {
    yearFolder = yearIter.next();
  } else {
    yearFolder = parentFolder.createFolder(year);
  }
  
  // 2. Cek Folder Bulan (di dalam folder Tahun)
  var targetFolder;
  var monthIter = yearFolder.getFoldersByName(monthName);
  if (monthIter.hasNext()) {
    targetFolder = monthIter.next();
  } else {
    targetFolder = yearFolder.createFolder(monthName);
  }
  
  return targetFolder;
}

/* ====================================================================== */
/* MODUL: PERJALANAN DINAS (SIABA) - FULL BACKEND                         */
/* ====================================================================== */

var ID_SS_DINAS = "1I_2yUFGXnBJTCSW6oaT3D482YCs8TIRkKgQVBbvpa1M"; 
var ID_FOLDER_DINAS = "1uPeOU7F_mgjZVyOLSsj-3LXGdq9rmmWl";

/* TIMPA FUNGSI INI DI SIABA.GS */
function getDaftarDinas(tahun, bulan, status, _cb) {
  try {
    SpreadsheetApp.flush();
    var ss = SpreadsheetApp.openById(ID_SS_DINAS);
    var sheet = ss.getSheetByName("Perjalanan_Dinas");
    if (!sheet) return JSON.stringify([]);

    var data = sheet.getDataRange().getValues();
    var result = [];
    
    // Filter
    var fTahun = (tahun == null) ? "" : String(tahun).trim();
    var fBulan = (bulan == null) ? "" : String(bulan).trim();
    var fStatus = (status == null) ? "" : String(status).trim();

    console.log("FILTER -> Thn: [" + fTahun + "], Bln: [" + fBulan + "]");

    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      if (String(row[1]).trim() === "") continue; 

      // Parsing Tanggal (Smart Parser)
      var valTgl = row[3];
      var rowTahun = "", rowBulan = "";

      if (valTgl instanceof Date) {
        rowTahun = String(valTgl.getFullYear());
        rowBulan = String(valTgl.getMonth() + 1);
      } else {
        var s = String(valTgl).replace(/'/g, "").trim();
        var parts = s.split(/[-/]/); 
        if (parts.length === 3) {
           if(parts[2].length === 4) { rowTahun = String(parts[2]); rowBulan = String(parseInt(parts[1], 10)); }
           else if (parts[0].length === 4) { rowTahun = String(parts[0]); rowBulan = String(parseInt(parts[1], 10)); }
        }
      }

      var matchTahun = (fTahun === "") || (rowTahun === fTahun);
      var matchBulan = (fBulan === "") || (rowBulan === fBulan);
      var matchStatus = (fStatus === "") || (String(row[9]) == fStatus);

      if (matchTahun && matchBulan && matchStatus) {
        // AMBIL LAST ACTIVITY (MAX DATE)
        // Kolom L(11), N(13), P(15)
        var t1 = parseTime(row[11]); // Tgl Kirim
        var t2 = parseTime(row[13]); // Last Update
        var t3 = parseTime(row[15]); // Tgl Verif
        var lastActivity = Math.max(t1, t2, t3);

        result.push({
          rowBaris: i + 1,
          jenis: row[0], noSpt: row[1], tglSpt: cleanDate(row[2]), tglMulai: cleanDate(row[3]), tglSelesai: cleanDate(row[4]),
          tujuan: row[5], kegiatan: row[6], jmlAsn: row[7], dokumen: row[8], status: row[9], jenisDok: row[10],
          tglKirim: cleanDate(row[11]), userKirim: row[12], lastUpdate: cleanDate(row[13]), lastUser: row[14],
          tglVerif: cleanDate(row[15]), verifikator: row[16], keterangan: row[17],
          timestamp: lastActivity // Simpan untuk sorting
        });
      }
    }
    
    // SORTING BERDASARKAN LAST ACTIVITY TERBARU
    result.sort(function(a, b) { return b.timestamp - a.timestamp; });
    
    return JSON.stringify(result);
  } catch (e) { return JSON.stringify([]); }
}

function simpanSptUnified(payload) {
  try {
    var ss = SpreadsheetApp.openById(ID_SS_DINAS);
    var sheetMaster = ss.getSheetByName("Perjalanan_Dinas");
    var sheetDetail = ss.getSheetByName("Perjalanan_Dinas_Peserta");
    
    if (!sheetDetail) {
      sheetDetail = ss.insertSheet("Perjalanan_Dinas_Peserta");
      sheetDetail.appendRow(["No SPT", "NIP", "Nama", "Unit", "Status", "Keterangan", "Waktu Input"]);
    }

    var now = new Date();
    var sysDateStr = "'" + Utilities.formatDate(now, Session.getScriptTimeZone(), "dd-MM-yyyy HH:mm");
    var userName = payload.user_login || "User Web";

    // Format Header
    var tglSptTxt = toTextDate(payload.header.tglSpt);
    var tglMulaiTxt = toTextDate(payload.header.tglMulai);
    var tglSelesaiTxt = toTextDate(payload.header.tglSelesai);

    // Proses File
    var fileUrl = "";
    if (payload.fileData && payload.fileName) {
      var folder = DriveApp.getFolderById(ID_FOLDER_DINAS);
      var blob = Utilities.newBlob(Utilities.base64Decode(payload.fileData), payload.mimeType, payload.fileName);
      var file = folder.createFile(blob);
      file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
      fileUrl = file.getUrl();
    }

    if (payload.isNewSpt) {
      // INSERT BARU
      sheetMaster.appendRow([
        payload.header.jenis, payload.header.noSpt, tglSptTxt, tglMulaiTxt, tglSelesaiTxt,
        payload.header.tujuan, payload.header.kegiatan, payload.listPeserta.length, fileUrl, "Diproses", 
        payload.header.jenisDok, sysDateStr, userName, sysDateStr, userName, "", "", ""
      ]);
    } else {
      // EDIT / UPDATE
      var dataM = sheetMaster.getDataRange().getValues();
      var found = false;
      for(var j=1; j<dataM.length; j++){
        if(String(dataM[j][1]).trim() === String(payload.header.noSpt).trim()) {
          var r = j + 1;
          // Update Header
          if(payload.header.jenis) sheetMaster.getRange(r, 1).setValue(payload.header.jenis);
          if(payload.header.tglSpt) sheetMaster.getRange(r, 3).setValue(tglSptTxt);
          if(payload.header.tglMulai) sheetMaster.getRange(r, 4).setValue(tglMulaiTxt);
          if(payload.header.tglSelesai) sheetMaster.getRange(r, 5).setValue(tglSelesaiTxt);
          if(payload.header.tujuan) sheetMaster.getRange(r, 6).setValue(payload.header.tujuan);
          if(payload.header.kegiatan) sheetMaster.getRange(r, 7).setValue(payload.header.kegiatan);
          
          if(fileUrl !== "") sheetMaster.getRange(r, 9).setValue(fileUrl);
          if(payload.header.jenisDok) sheetMaster.getRange(r, 11).setValue(payload.header.jenisDok);

          // Update Log (Kolom N & O)
          var curJml = parseInt(dataM[j][7] || 0);
          sheetMaster.getRange(r, 8).setValue(curJml + payload.listPeserta.length);
          sheetMaster.getRange(r, 14).setValue(sysDateStr); // Last Update
          sheetMaster.getRange(r, 15).setValue(userName);   // Last User
          found = true;
          break;
        }
      }
      if(!found) return "Error: Data tidak ditemukan.";
    }

    // Simpan Peserta
    var rowsPeserta = [];
    payload.listPeserta.forEach(function(p){
      rowsPeserta.push([payload.header.noSpt, p.nip, p.nama, p.unit, "Diproses", "", sysDateStr]);
    });
    if(rowsPeserta.length > 0) {
      sheetDetail.getRange(sheetDetail.getLastRow() + 1, 1, rowsPeserta.length, 7).setValues(rowsPeserta);
    }

    SpreadsheetApp.flush();
    return "Sukses";
  } catch (e) { return "Error: " + e.toString(); }
}

/* 2. SIMPAN DATA (UNIFIED) */
function renderTabelDinas(jsonString) {
    var data = []; try { data = JSON.parse(jsonString); } catch(e) { return; }
    console.log("Rendering " + data.length + " data.");
    
    // MATIKAN DATATABLES DULU
    if ($.fn.DataTable.isDataTable('#tabelDinas')) {
        $('#tabelDinas').DataTable().destroy();
    }

    $('#bodyTabelDinas').empty();

    if(!data || data.length === 0) { 
        initDataTableStandard('#tabelDinas'); 
        return; 
    }

    var isAdmin = false; 
    try { var user = JSON.parse(localStorage.getItem("siksUser")); if (user && (user.role.toLowerCase().includes('admin') || user.role.toLowerCase().includes('super'))) isAdmin = true; } catch(e) {}
    
    var html = "";
    for (var i = 0; i < data.length; i++) {
        var row = data[i]; 
        DATA_DINAS[row.rowBaris] = row;
        
        var safeNoSpt = cleanText(row.noSpt); 
        var safeUrl = (row.dokumen||"").replace(/'/g, "%27").replace(/"/g, "%22");
        
        var btnDetail = `<button class="btn btn-outline-info btn-aksi-table" onclick="tambahDataDinasWithVal('${safeNoSpt}')"><i class="fas fa-pencil-alt mr-1"></i> Detail / Edit</button>`;
        
        // --- UPDATE TOMBOL DOKUMEN DISINI (PREVIEW MODAL) ---
        var btnFile = '-';
        if(safeUrl.length > 5) {
             var jns = (row.jenisDok || "").toUpperCase();
             var iconBtn = (jns === "SPT") ? 'btn-icon-info' : 'btn-icon-warning';
             var iconFa = (jns === "SPT") ? 'fa-eye' : 'fa-exclamation-triangle';
             
             // Panggil fungsi previewDokumen() alih-alih window.open()
             btnFile = `<button class="btn-icon-sultan ${iconBtn}" onclick="previewDokumen('${safeUrl}')" title="Lihat Dokumen"><i class="fas ${iconFa}" style="font-size:12px;"></i></button>`;
        }
        
        var btnAdmin = '-';
        if (isAdmin) { 
            btnAdmin = `<div class="d-flex justify-content-center" style="gap: 5px;">
                <button class="btn-icon-sultan btn-icon-success" onclick="bukaModalVerif('${row.rowBaris}')" title="Verifikasi"><i class="fas fa-check" style="font-size:12px;"></i></button>
                <button class="btn-icon-sultan btn-icon-danger" onclick="hapusDataDinas('${row.rowBaris}')" title="Hapus"><i class="fas fa-trash" style="font-size:12px;"></i></button>
            </div>`; 
        }

        html += `<tr class="tr-animasi" style="animation: slideInUp 0.4s ease-out forwards; animation-delay: ${i * 0.05}s;">
            <td class="align-middle bg-white border-right">${btnDetail}</td>
            <td class="align-middle text-black-sultan">${cleanText(row.jenis)}</td>
            <td class="align-middle font-weight-bold text-black-sultan">${safeNoSpt}</td>
            <td class="align-middle text-black-sultan">${row.tglSpt}</td>
            <td class="align-middle text-black-sultan">${row.tglMulai}</td>
            <td class="align-middle text-black-sultan">${row.tglSelesai}</td>
            <td class="col-wrap-sultan text-black-sultan">${cleanText(row.tujuan)}</td>
            <td class="col-wrap-sultan text-black-sultan small">${cleanText(row.kegiatan)}</td>
            <td class="text-center align-middle font-weight-bold text-black-sultan">${row.jmlAsn||'0'}</td>
            <td class="text-center align-middle">${btnFile}</td>
            <td class="align-middle small text-black-sultan">${cleanText(row.jenisDok)}</td>
            <td class="text-center align-middle">${renderBadgeSultan(row.status, row.keterangan)}</td>
            <td class="text-center align-middle">${btnAdmin}</td>
            <td class="align-middle small text-muted">${row.tglKirim||'-'}</td>
            <td class="align-middle small text-muted">${cleanText(row.userKirim)}</td>
            <td class="align-middle small text-muted">${row.lastUpdate||'-'}</td>
            <td class="align-middle small text-muted">${cleanText(row.lastUser)}</td>
            <td class="align-middle small text-muted">${row.tglVerif||'-'}</td>
            <td class="align-middle small text-muted">${cleanText(row.verifikator)}</td>
        </tr>`;
    }
    
    $('#bodyTabelDinas').html(html);
    initDataTableStandard('#tabelDinas');
}

/* 3. FUNGSI PENCARIAN PEGAWAI (YANG DIPERBAIKI) */
function cariPegawaiDatabase(keyword) {
  var ss = SpreadsheetApp.openById(ID_SS_DINAS);
  var sheet = ss.getSheetByName("Database"); 
  if(!sheet) return JSON.stringify([]);

  var data = sheet.getDataRange().getValues();
  var result = []; // Nama variabel konsisten
  var k = keyword.toLowerCase();

  for(var i=1; i<data.length; i++) {
    var nip = String(data[i][1]).toLowerCase(); 
    var nama = String(data[i][2]).toLowerCase();
    
    if(nama.includes(k) || nip.includes(k)) {
       // Ambil Unit, NIP, Nama
       result.push({ unit: data[i][0], nip: data[i][1], nama: data[i][2] });
       
       // Batasi 10 hasil agar tidak berat
       if(result.length >= 10) break;
    }
  }
  return JSON.stringify(result);
}

/* 5. FUNGSI PENDUKUNG LAINNYA */
function verifikasiDataDinas(rowBaris, status, keterangan, userVerifikator) {
  try {
    var ss = SpreadsheetApp.openById(ID_SS_DINAS);
    var sheet = ss.getSheetByName("Perjalanan_Dinas");
    var row = parseInt(rowBaris);
    var verifikator = userVerifikator || "Admin";
    var now = new Date();
    var sysDateStr = "'" + Utilities.formatDate(now, Session.getScriptTimeZone(), "dd-MM-yyyy HH:mm");

    sheet.getRange(row, 10).setValue(status);
    sheet.getRange(row, 14).setValue(sysDateStr);
    sheet.getRange(row, 15).setValue(verifikator);
    sheet.getRange(row, 16).setValue(sysDateStr);
    sheet.getRange(row, 17).setValue(verifikator);

    if (keterangan) {
        sheet.getRange(row, 18).setValue(keterangan);
    } else if (status === 'Disetujui') {
        sheet.getRange(row, 18).setValue("");
    }
    
    SpreadsheetApp.flush();
    Utilities.sleep(1500); 
    return "Sukses";
  } catch(e) { return "Error: " + e.toString(); }
}

function hapusDataDinas(rowBaris) {
  try {
    var ss = SpreadsheetApp.openById(ID_SS_DINAS);
    var sheet = ss.getSheetByName("Perjalanan_Dinas");
    sheet.deleteRow(parseInt(rowBaris));
    SpreadsheetApp.flush();
    Utilities.sleep(1500); 
    return "Sukses";
  } catch(e) { return "Error: " + e.toString(); }
}

function cekInfoSpt(noSpt) {
  try {
    var ss = SpreadsheetApp.openById(ID_SS_DINAS);
    var sheet = ss.getSheetByName("Perjalanan_Dinas");
    if (!sheet) return JSON.stringify({ found: false });
    var data = sheet.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][1]).trim().toUpperCase() === String(noSpt).trim().toUpperCase()) {
        return JSON.stringify({
          found: true,
          data: {
            jenis: data[i][0],
            tglSpt: toHtmlDate(data[i][2]),
            tglMulai: toHtmlDate(data[i][3]),
            tglSelesai: toHtmlDate(data[i][4]),
            tujuan: data[i][5],
            kegiatan: data[i][6],
            status: data[i][9],
            jenisDok: data[i][10]
          }
        });
      }
    }
    return JSON.stringify({ found: false });
  } catch(e) { return JSON.stringify({ found: false }); }
}

function getPesertaDinas(noSpt) {
  var ss = SpreadsheetApp.openById(ID_SS_DINAS);
  var sheet = ss.getSheetByName("Perjalanan_Dinas_Peserta");
  if (!sheet) return JSON.stringify([]);
  var data = sheet.getDataRange().getValues();
  var result = [];
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim().toUpperCase() === String(noSpt).trim().toUpperCase()) {
      result.push({ nip: data[i][1], nama: data[i][2], unit: data[i][3], status: data[i][4] });
    }
  }
  return JSON.stringify(result);
}

/* ======================================================================
   MODUL: DATA CUTI (SIABA)
   Spreadsheet ID: 1UYG80gGxuC19ieaVBzJaUV8bhlS2q5gExr0-Yl7upKo
   Sheet: Form Cuti
   ====================================================================== */

var ID_SS_CUTI = "1UYG80gGxuC19ieaVBzJaUV8bhlS2q5gExr0-Yl7upKo";
var ID_FOLDER_CUTI = "1uPeOU7F_mgjZVyOLSsj-3LXGdq9rmmWl"; // Gunakan folder yang sama atau buat baru

/* 1. GET DATA CUTI (FILTER & SORT) */
/* ======================================================================
   MODUL: DATA CUTI PEGAWAI
   Spreadsheet ID: 1UYG80gGxuC19ieaVBzJaUV8bhlS2q5gExr0-Yl7upKo
   Sheet: Form Cuti
   ====================================================================== */

var ID_SS_CUTI = "1UYG80gGxuC19ieaVBzJaUV8bhlS2q5gExr0-Yl7upKo";

function getDaftarCuti(tahun, bulan, unit, status) {
  try {
    var ss = SpreadsheetApp.openById(ID_SS_CUTI);
    var sheet = ss.getSheetByName("Form Cuti");
    if (!sheet) return JSON.stringify([]);

    var data = sheet.getDataRange().getDisplayValues(); // Gunakan DisplayValues agar tanggal terbaca sebagai teks
    var result = [];

    // Filter Variables
    var fTahun  = (tahun && String(tahun).trim() !== "") ? String(tahun).trim() : null;
    var fBulan  = (bulan && String(bulan).trim() !== "") ? String(bulan).trim() : null;
    var fStatus = (status && String(status).trim() !== "" && status !== "SEMUA") ? String(status).trim() : null;
    var arrBulan = ["Januari", "Februari", "Maret", "April", "Mei", "Juni", "Juli", "Agustus", "September", "Oktober", "November", "Desember"];

    // Loop data (Mulai baris 2 / Index 1)
    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      if (!row[1]) continue; 

      // 1. FILTER TANGGAL MULAI (Kolom E / Index 4)
      var tglMulai = String(row[4]); 
      var passTgl = true;
      
      if (fTahun || fBulan) {
          // Parsing Tanggal Indonesia (dd MMMM yyyy atau dd/mm/yyyy)
          var dateObj = parseDateIndo(tglMulai);
          if (dateObj) {
              if (fTahun && String(dateObj.getFullYear()) !== fTahun) passTgl = false;
              if (fBulan && arrBulan[dateObj.getMonth()] !== fBulan) passTgl = false;
          } else {
             passTgl = false; // Tanggal tidak valid/kosong
          }
      }

      // Filter Status (Kolom K / Index 10)
      if (fStatus && String(row[10]) !== fStatus) continue;

      if (!passTgl) continue;

      // 2. LOGIKA SORTING (LAST ACTIVITY)
      // Kolom N (Input) = Index 13
      // Kolom P (Edit)  = Index 15
      // Kolom R (Verif) = Index 17
      
      // Menggunakan fungsi parseTime yang SUDAH ADA di file Siaba.gs Anda
      var tInput = parseTime(row[13]); 
      var tEdit  = parseTime(row[15]); 
      var tVerif = parseTime(row[17]); 
      
      var lastActivity = Math.max(tInput, tEdit, tVerif);

      result.push({
        rowBaris: i + 1,
        unit: row[0], nama: row[1], nip: row[2], jenis: row[3],
        tglMulai: row[4], tglSelesai: row[5], jumlah: row[6],
        alasan: row[7], alamat: row[8], telepon: row[9],
        status: row[10], ket: row[11], fileUrl: row[12],
        tglInput: row[13], userInput: row[14],
        tglEdit: row[15], userEdit: row[16],
        tglVerif: row[17], verifikator: row[18],
        sisaCt: row[19], ambilCt: row[20],
        timestamp: lastActivity // Key sorting
      });
    }

    // Sort Descending (Terbaru paling atas)
    result.sort(function(a, b) { return b.timestamp - a.timestamp; });
    
    return JSON.stringify(result);
  } catch (e) { return JSON.stringify([]); }
}

/* 2. SIMPAN / UPDATE CUTI */
function simpanDataCuti(payload) {
  try {
    var ss = SpreadsheetApp.openById(ID_SS_CUTI);
    var sheet = ss.getSheetByName("Form Cuti");
    if (!sheet) return "Error: Sheet tidak ditemukan.";

    var now = new Date();
    var sysDateStr = "'" + Utilities.formatDate(now, Session.getScriptTimeZone(), "dd-MM-yyyy HH:mm:ss");
    var userName = payload.user_login || "User Web";

    // A. Handle File Upload
    var fileUrl = "";
    if (payload.fileData && payload.fileName) {
      var folder = DriveApp.getFolderById(ID_FOLDER_DINAS); // Gunakan folder dinas atau folder khusus cuti
      var blob = Utilities.newBlob(Utilities.base64Decode(payload.fileData), payload.mimeType, payload.fileName);
      var file = folder.createFile(blob);
      file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
      fileUrl = file.getUrl();
    }

    // B. Mode Insert vs Edit
    if (payload.isNew) {
        // INSERT
        sheet.appendRow([
            payload.unit,       // A
            payload.nama,       // B
            "'" + payload.nip,  // C
            payload.jenis,      // D
            "'" + payload.tglMulai,   // E
            "'" + payload.tglSelesai, // F
            payload.jumlah,     // G
            payload.alasan,     // H
            payload.alamat,     // I
            "'" + payload.telepon, // J
            "Diproses",         // K (Status Awal)
            "",                 // L (Ket)
            fileUrl,            // M
            sysDateStr,         // N (Tgl Input)
            userName,           // O (User Input)
            "", "", "", "", "", "" // P-U Kosong
        ]);
    } else {
        // EDIT
        var rowIdx = parseInt(payload.recId);
        // Validasi Status sebelum edit
        var statusLama = sheet.getRange(rowIdx, 11).getValue();
        if (String(statusLama).includes("Disetujui") || String(statusLama).includes("OK")) {
            return "Error: Data sudah disetujui, tidak dapat diedit.";
        }

        // Update Data Utama
        // Kolom A-J
        sheet.getRange(rowIdx, 4).setValue(payload.jenis);
        sheet.getRange(rowIdx, 5).setValue("'" + payload.tglMulai);
        sheet.getRange(rowIdx, 6).setValue("'" + payload.tglSelesai);
        sheet.getRange(rowIdx, 7).setValue(payload.jumlah);
        sheet.getRange(rowIdx, 8).setValue(payload.alasan);
        sheet.getRange(rowIdx, 9).setValue(payload.alamat);
        sheet.getRange(rowIdx, 10).setValue("'" + payload.telepon);
        
        // Jika ada file baru, update kolom M (13)
        if(fileUrl !== "") {
            sheet.getRange(rowIdx, 13).setValue(fileUrl);
        }

        // Log Update: P (16) & Q (17)
        sheet.getRange(rowIdx, 16).setValue(sysDateStr);
        sheet.getRange(rowIdx, 17).setValue(userName);
        
        // Reset Status ke Diproses jika sebelumnya Revisi/Ditolak
        sheet.getRange(rowIdx, 11).setValue("Diproses");
    }

    SpreadsheetApp.flush();
    return "Sukses";
  } catch (e) { return "Error: " + e.toString(); }
}

/* 3. HAPUS CUTI */
function hapusDataCuti(rowId) {
  try {
    var ss = SpreadsheetApp.openById(ID_SS_CUTI);
    var sheet = ss.getSheetByName("Form Cuti");
    sheet.deleteRow(parseInt(rowId));
    return "Sukses";
  } catch(e) { return "Error: " + e.toString(); }
}

/* 4. VERIFIKASI CUTI */
function verifikasiCuti(payload) {
  try {
    var ss = SpreadsheetApp.openById(ID_SS_CUTI);
    var sheet = ss.getSheetByName("Form Cuti");
    var row = parseInt(payload.recId);
    
    var now = new Date();
    var sysDateStr = "'" + Utilities.formatDate(now, Session.getScriptTimeZone(), "dd-MM-yyyy HH:mm:ss");

    // Status (K/11), Ket (L/12)
    sheet.getRange(row, 11).setValue(payload.status);
    sheet.getRange(row, 12).setValue(payload.keterangan);
    
    // Log Verif: R(18) & S(19)
    sheet.getRange(row, 18).setValue(sysDateStr);
    sheet.getRange(row, 19).setValue(payload.verifikator);

    return "Sukses";
  } catch(e) { return "Error: " + e.toString(); }
}

// 1. MASTER PARSER WAKTU (Menangani Sorting & Timestamp)
function parseTime(val) {
  if (!val) return 0;
  
  // A. Jika tipe datanya sudah Date object (Format Date native Excel)
  if (val instanceof Date) return val.getTime();

  // B. Bersihkan tanda petik (') dan spasi
  var s = String(val).replace(/'/g, "").trim();
  if (s === "") return 0;

  // C. Cek Format Mesin ISO (yyyy-mm-dd)
  // Contoh: 2026-01-20 14:30
  var isoCheck = s.split("-");
  if (isoCheck.length === 3 && isoCheck[0].length === 4) {
      return new Date(s).getTime();
  }

  // D. Parsing Format Indonesia (dd-mm-yyyy atau dd/mm/yyyy)
  var parts = s.split(" "); // Pisahkan Tanggal dan Jam
  var dateStr = parts[0];
  var timeStr = (parts.length > 1) ? parts[1] : "00:00:00";

  // Deteksi pemisah: strip (-) atau miring (/)
  var separator = dateStr.includes("-") ? "-" : "/";
  var dParts = dateStr.split(separator);
  
  // Pastikan ada 3 bagian (tgl, bln, thn)
  if (dParts.length !== 3) return 0;

  var tgl = parseInt(dParts[0], 10);
  var bln = parseInt(dParts[1], 10) - 1; // JS Month mulai dari 0
  var thn = parseInt(dParts[2], 10);

  // Parsing Jam (Support HH, HH:mm, HH:mm:ss)
  var h = 0, min = 0, sec = 0;
  var tParts = timeStr.split(":");
  if (tParts.length >= 1) h = parseInt(tParts[0], 10);
  if (tParts.length >= 2) min = parseInt(tParts[1], 10);
  if (tParts.length >= 3) sec = parseInt(tParts[2], 10);

  // Validasi Angka
  if (isNaN(tgl) || isNaN(bln) || isNaN(thn)) return 0;

  return new Date(thn, bln, tgl, h, min, sec).getTime();
}

// 2. PARSER TANGGAL INDO (Khusus Filter: "20 Januari 2026")
function parseDateIndo(str) {
    if(!str) return null;
    var months = ["januari","februari","maret","april","mei","juni","juli","agustus","september","oktober","november","desember"];
    
    str = String(str).toLowerCase().replace(/,/g, "");
    var parts = str.split(" ");
    
    // Format: 20 Januari 2026
    if (parts.length >= 3) {
        var d = parseInt(parts[0]);
        var mIdx = months.indexOf(parts[1]);
        var y = parseInt(parts[2]);
        if(mIdx > -1) return new Date(y, mIdx, d);
    } 
    
    // Fallback: 20/01/2026
    var p2 = str.split("/");
    if (p2.length === 3) return new Date(p2[2], p2[1]-1, p2[0]);
    
    return null;
}

// 3. HELPER TEXT SEDERHANA (Legacy untuk Modul Perjalanan Dinas)
function cleanDate(val) { 
  if (val instanceof Date) return Utilities.formatDate(val, Session.getScriptTimeZone(), "dd-MM-yyyy"); 
  return String(val).replace(/'/g, "").trim(); 
}

function toTextDate(isoDate) { 
  // Input: 2026-01-20 -> Output: '20-01-2026
  if(!isoDate) return ""; 
  var parts = isoDate.split("-"); 
  if(parts.length !== 3) return "'" + isoDate; 
  return "'" + parts[2] + "-" + parts[1] + "-" + parts[0]; 
}

function toHtmlDate(textDate) { 
  // Input: '20-01-2026 -> Output: 2026-01-20 (untuk value input type="date")
  var s = String(textDate).replace(/'/g, "").trim(); 
  var parts = s.split("-"); 
  if(parts.length !== 3) return ""; 
  return parts[2] + "-" + parts[1] + "-" + parts[0]; 
}

/* ======================================================================
   MODUL: PENGAJUAN CUTI (BACKEND) - FIXED
   ====================================================================== */

var ID_SS_CUTI = "1UYG80gGxuC19ieaVBzJaUV8bhlS2q5gExr0-Yl7upKo"; // Ganti dengan ID Spreadsheet Cuti Anda

/* 1. AMBIL DATABASE REFERENSI (Optimized) */
function getDatabaseCutiOptions() {
  try {
    var ss = SpreadsheetApp.openById(ID_SS_CUTI);
    var sheet = ss.getSheetByName("Database Cuti");
    if (!sheet) return JSON.stringify([]);

    var data = sheet.getDataRange().getValues();
    var result = [];

    for (var i = 1; i < data.length; i++) { 
      if (data[i][0] && data[i][2]) {
        result.push({
          nip: String(data[i][0]),
          unit: String(data[i][1]),
          nama: String(data[i][2]),
          status: String(data[i][3]),
          alamat: String(data[i][8]),
          hp: String(data[i][9])
        });
      }
    }
    return JSON.stringify(result);
  } catch (e) { return JSON.stringify([]); }
}

/* ======================================================================
   HELPER VALIDASI: CEK BENTROK TANGGAL CUTI
   Return: null jika aman, string error jika bentrok
   ====================================================================== */
function cekBentrokCuti(nipBaru, tglMulaiBaruStr, tglSelesaiBaruStr, rowIdPengecualian) {
  var ss = SpreadsheetApp.openById(ID_SS_CUTI);
  var sheet = ss.getSheetByName("Form Cuti");
  var data = sheet.getDataRange().getValues();
  
  var dMulaiBaru = new Date(tglMulaiBaruStr); // Format yyyy-mm-dd
  var dSelesaiBaru = new Date(tglSelesaiBaruStr);
  
  // Reset jam agar komparasi murni tanggal
  dMulaiBaru.setHours(0,0,0,0);
  dSelesaiBaru.setHours(0,0,0,0);

  // Loop semua data (Skip header baris 1)
  for (var i = 1; i < data.length; i++) {
    // Skip jika sedang edit baris ini sendiri
    if (rowIdPengecualian && (i + 1) == rowIdPengecualian) continue;

    var rowNip = String(data[i][2]).replace(/'/g, "").trim(); // Kolom C
    var rowStatus = String(data[i][10]).toLowerCase();        // Kolom K
    
    // Cek NIP sama & Status Aktif (Bukan Ditolak/Dibatalkan)
    if (rowNip === String(nipBaru).trim() && !rowStatus.includes("tolak") && !rowStatus.includes("batal")) {
      
      // Ambil Tanggal Lama (Kolom E & F) - Parsing manual jika format teks Indonesia
      var tglMulaiLama = parseDateIndo(data[i][4]) || new Date(data[i][4]);
      var tglSelesaiLama = parseDateIndo(data[i][5]) || new Date(data[i][5]);
      
      if (tglMulaiLama && tglSelesaiLama) {
        tglMulaiLama.setHours(0,0,0,0);
        tglSelesaiLama.setHours(0,0,0,0);
        
        // RUMUS TABRAKAN TANGGAL:
        // (StartA <= EndB) and (EndA >= StartB)
        if (dMulaiBaru <= tglSelesaiLama && dSelesaiBaru >= tglMulaiLama) {
           var conflictDate = Utilities.formatDate(tglMulaiLama, "GMT+7", "dd/MM/yyyy") + " s.d " + 
                              Utilities.formatDate(tglSelesaiLama, "GMT+7", "dd/MM/yyyy");
           return "Gagal: Tanggal bentrok dengan pengajuan aktif (" + conflictDate + ")";
        }
      }
    }
  }
  return null; // Aman
}

/* 2. SIMPAN PENGAJUAN CUTI */
function simpanPengajuanCuti(payload) {
  try {
    var errorBentrok = cekBentrokCuti(payload.nip, payload.tglMulai, payload.tglSelesai, null);
    if (errorBentrok) return errorBentrok;
    var ss = SpreadsheetApp.openById(ID_SS_CUTI);
    var sheet = ss.getSheetByName("Form Cuti");
    if (!sheet) return "Error: Sheet 'Form Cuti' tidak ditemukan.";

    var now = new Date();
    var sysDateStr = "'" + Utilities.formatDate(now, Session.getScriptTimeZone(), "dd-MM-yyyy HH:mm:ss");
    var userName = payload.userInput || "User Web";

    // 1. FORMAT TANGGAL
    var tglMulaiIndo   = formatIndoText(payload.tglMulai);
    var tglSelesaiIndo = formatIndoText(payload.tglSelesai);
    
    // --- PERBAIKAN: Gunakan payload.tglPengajuan ---
    var tglPengajuanFormat = formatTglIndo(payload.tglPengajuan);

    // 2. AMBIL DATA DETIL PEGAWAI
    var dbData = getDetailPegawaiByNip(payload.nip); 
    var empGol = dbData ? dbData.golongan : ""; 
    var empJab = dbData ? dbData.jabatan : ""; 
    
    // 3. LOOKUP PEJABAT STRUKTURAL
    var pejabat = lookupPejabatStruktural(payload.jenisCuti, payload.unit, empGol, empJab);
    
    var final_kepada, final_nama_atasan, final_nip_atasan, final_jab_atasan, final_nama_setuju, final_nip_setuju, final_jab_setuju;

    if (pejabat) {
        final_kepada      = pejabat.kepada;
        final_nama_atasan = pejabat.nama_atasan;
        final_nip_atasan  = pejabat.nip_atasan;
        final_jab_atasan  = pejabat.jabatan_atasan;
        final_nama_setuju = pejabat.nama_setuju;
        final_nip_setuju  = pejabat.nip_setuju;
        final_jab_setuju  = pejabat.jabatan_setuju;
    } else {
        final_kepada      = dbData ? dbData.fullRow[19] : ""; 
        final_nama_atasan = dbData ? dbData.fullRow[13] : ""; 
        final_nip_atasan  = dbData ? dbData.fullRow[14] : ""; 
        final_jab_atasan  = dbData ? dbData.fullRow[15] : ""; 
        final_nama_setuju = dbData ? dbData.fullRow[16] : ""; 
        final_nip_setuju  = dbData ? dbData.fullRow[17] : ""; 
        final_jab_setuju  = dbData ? dbData.fullRow[18] : ""; 
    }

    // 4. LOGIKA SISA CUTI
    var thnMulai = parseInt(payload.tglMulai.split("-")[0]);
    var sisaN2 = "-", sisaN1 = "-", sisaN = "-";
    if (dbData) {
        if (thnMulai === 2023) { sisaN2 = dbData.fullRow[26]||"0"; sisaN1 = dbData.fullRow[28]||"0"; sisaN = dbData.fullRow[30]||"0"; }
        else if (thnMulai === 2024) { sisaN2 = dbData.fullRow[38]||"0"; sisaN1 = dbData.fullRow[40]||"0"; sisaN = dbData.fullRow[42]||"0"; }
        else if (thnMulai === 2025) { sisaN2 = dbData.fullRow[50]||"0"; sisaN1 = dbData.fullRow[52]||"0"; sisaN = dbData.fullRow[54]||"0"; }
        else if (thnMulai === 2026) { sisaN2 = dbData.fullRow[62]||"0"; sisaN1 = dbData.fullRow[64]||"0"; sisaN = dbData.fullRow[66]||"0"; }
    }

    // 5. CHECKLIST JENIS CUTI
    var j = String(payload.jenisCuti).toLowerCase();
    var c = { ct:"", cs:"", cap:"", cb:"", cm:"", cltn:"" }; 
    var CHECK = "✓"; 
    
    if (j.includes("tahunan") || j.includes("umroh")) c.ct = CHECK;
    if (j.includes("sakit")) c.cs = CHECK;
    else if (j.includes("penting")) c.cap = CHECK;
    else if (j.includes("besar")) c.cb = CHECK;
    else if (j.includes("melahirkan")) c.cm = CHECK;
    else if (j.includes("luar") || j.includes("tanggungan")) c.cltn = CHECK;

    // 6. SUSUN DATA PDF
    var pdfData = {
        // Gunakan Tanggal Pengajuan dari User untuk Tanggal Surat
        tanggal: tglPengajuanFormat, 
        kepada: final_kepada,
        
        asn: payload.nama,
        nip: payload.nip,
        jabatan: dbData ? dbData.jabatan : "", 
        masa_kerja: dbData ? dbData.masaKerja : "", 
        unit: dbData ? dbData.unitLengkap : payload.unit, 
        
        alasan: payload.alasan,
        jumlah: payload.jumlahHari,
        tmc: tglMulaiIndo,
        tsc: tglSelesaiIndo,
        
        "N-2": sisaN2, "N-1": sisaN1, "N": sisaN,
        
        alamat: payload.alamat,
        telp: payload.hp,
        
        jabatan_atasan: final_jab_atasan,
        nama_atasan: final_nama_atasan,
        nip_atasan: final_nip_atasan,
        
        jabatan_setuju: final_jab_setuju,
        nama_setuju: final_nama_setuju,
        nip_setuju: final_nip_setuju,
        
        ct: c.ct, cs: c.cs, cap: c.cap, cb: c.cb, cm: c.cm, cltn: c.cltn,

        jenisCutiRaw: payload.jenisCuti, 
        tglMulaiRaw: payload.tglMulai    
    };
    
    // GENERATE PDF
    var linkPdf = generatePdfCuti(pdfData); 

    // 7. SIMPAN SPREADSHEET
    var spacer = ["", "", "", "", "", ""]; 
    var rowData = [
      payload.unit, payload.nama, "'" + payload.nip, payload.jenisCuti, 
      tglMulaiIndo, tglSelesaiIndo, payload.jumlahHari, payload.alasan, 
      payload.alamat, "'" + payload.hp, "Diproses", "", linkPdf, 
      sysDateStr, userName,
      
      ...spacer, 

      tglPengajuanFormat, // Kolom V
      pdfData.jabatan,    // W
      pdfData.masa_kerja, // X
      pdfData.unit,       // Y
      c.ct, c.cb, c.cs, c.cm, c.cap, c.cltn, 
      sisaN2, sisaN1, sisaN,
      final_jab_atasan, final_nama_atasan, final_nip_atasan,
      final_jab_setuju, final_nama_setuju, final_nip_setuju,
      final_kepada
    ];

    sheet.appendRow(rowData);
    SpreadsheetApp.flush();
    return "Sukses";
    
  } catch (e) { return "Error: " + e.toString(); }
}

/* 3. GENERATE PDF */
function generatePdfCuti(data) {
  var ID_TEMPLATE = "1k5KmEZj5nikuUV-MLnY4c6Tn-jFIhmOMGwhjvqaUSzk"; 
  var ID_FOLDER_INDUK = "1suNhGklZ931kT6Y5wbp5x_92ZCtlWfQz"; 
  var ID_IMAGE_CHECK = "1AbFps5ZiyeBH9hVa_XTYvfnoO77DxFle";

  try {
    var templateFile = DriveApp.getFileById(ID_TEMPLATE);
    var indukFolder = DriveApp.getFolderById(ID_FOLDER_INDUK);
    var checkImgBlob = DriveApp.getFileById(ID_IMAGE_CHECK).getBlob();

    var parts = data.tglMulaiRaw.split("-");
    var year = parts[0]; 
    var monthIndex = parseInt(parts[1]) - 1; 
    var monthNames = ["Januari", "Februari", "Maret", "April", "Mei", "Juni", "Juli", "Agustus", "September", "Oktober", "November", "Desember"];
    var monthName = monthNames[monthIndex];

    var yearFolder = getOrCreateSubfolder(indukFolder, year);
    var targetFolder = getOrCreateSubfolder(yearFolder, monthName);

    var fileName = data.jenisCutiRaw + " - " + data.asn + " - " + data.tmc + ".pdf";

    var tempFile = templateFile.makeCopy(fileName, targetFolder);
    var tempDoc = DocumentApp.openById(tempFile.getId());
    var body = tempDoc.getBody();
    
    for (var key in data) {
      if (data.hasOwnProperty(key)) {
        var val = data[key];
        if (["ct","cs","cb","cm","cap","cltn"].indexOf(key) > -1) {
            if (val === "✓") {
                replaceTextWithImage(body, "{{" + key + "}}", checkImgBlob);
            } else {
                body.replaceText("{{" + key + "}}", ""); 
            }
        } 
        else if (key !== "jenisCutiRaw" && key !== "tglMulaiRaw") {
            var txt = val == null ? "" : String(val);
            body.replaceText("{{" + key + "}}", txt);
        }
      }
    }
    
    tempDoc.saveAndClose();
    var pdfBlob = tempFile.getAs(MimeType.PDF);
    var pdfFile = targetFolder.createFile(pdfBlob).setName(fileName);
    pdfFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    tempFile.setTrashed(true);
    
    return pdfFile.getUrl();
  } catch (e) { return "Error PDF: " + e.toString(); }
}

function replaceTextWithImage(body, placeholder, imgBlob) {
  var next = body.findText(placeholder);
  while (next) {
    var element = next.getElement();
    var start = next.getStartOffset();
    var end = next.getEndOffsetInclusive();
    element.deleteText(start, end);
    var img = element.getParent().asParagraph().insertInlineImage(start, imgBlob);
    img.setWidth(11).setHeight(11); 
    next = body.findText(placeholder);
  }
}

function getDetailPegawaiByNip(targetNip) {
  var ss = SpreadsheetApp.openById(ID_SS_CUTI);
  var sheet = ss.getSheetByName("Database Cuti");
  if (!sheet) return null;
  
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    var rowNip = String(data[i][0]).trim(); 
    if (rowNip === String(targetNip).trim()) {
      return {
        golongan:  data[i][4],  
        jabatan:   data[i][5],  
        unitLengkap: data[i][6], 
        masaKerja: data[i][7],  
        fullRow: data[i]        
      };
    }
  }
  return null;
}

function lookupPejabatStruktural(jenisCuti, unitUser, golUser, tugasUser) {
  var ss = SpreadsheetApp.openById(ID_SS_CUTI);
  var sheet = ss.getSheetByName("Data Atasan");
  if (!sheet) return null;

  var data = sheet.getDataRange().getValues();
  
  var j = String(jenisCuti).toLowerCase().trim();
  var t = String(tugasUser).toLowerCase().trim(); 
  var g = String(golUser).toLowerCase().trim();
  var u = String(unitUser).toLowerCase().trim();

  // 1. UMROH
  if (j === "cuti umroh") {
    for (var i = 1; i < data.length; i++) {
       if (String(data[i][0]).toLowerCase().trim() === "cuti umroh") return mapRow(data[i]);
    }
  }

  // 2. GOLONGAN IV
  if (g.includes("iv/") || g === "iv") {
    for (var i = 1; i < data.length; i++) {
       var ruleGol = String(data[i][2]).toLowerCase().trim();
       if (ruleGol !== "" && g.includes(ruleGol)) return mapRow(data[i]);
    }
  }

  // 3. KEPALA SD
  for (var i = 1; i < data.length; i++) {
     var ruleTugas = String(data[i][1]).toLowerCase().trim(); 
     if (ruleTugas !== "" && t.includes(ruleTugas)) return mapRow(data[i]);
  }

  // 4. UNIT KERJA
  for (var i = 1; i < data.length; i++) {
     var ruleUnit = String(data[i][3]).toLowerCase().trim();
     if (ruleUnit !== "" && ruleUnit === u) return mapRow(data[i]);
  }

  return null; 
}

function mapRow(row) {
  return {
     nama_atasan:    row[4], 
     nip_atasan:     row[5], 
     jabatan_atasan: row[6], 
     nama_setuju:    row[7], 
     nip_setuju:     row[8], 
     jabatan_setuju: row[9], 
     kepada:         row[10] 
  };
}

function formatIndoText(isoDate) {
  if (!isoDate) return "";
  var parts = isoDate.split("-");
  if (parts.length !== 3) return isoDate;
  var months = ["Januari", "Februari", "Maret", "April", "Mei", "Juni", "Juli", "Agustus", "September", "Oktober", "November", "Desember"];
  var y = parts[0];
  var m = parseInt(parts[1], 10) - 1;
  var d = parseInt(parts[2], 10);
  return d + " " + months[m] + " " + y;
}

function getOrCreateSubfolder(parentFolder, folderName) {
  var folders = parentFolder.getFoldersByName(folderName);
  return folders.hasNext() ? folders.next() : parentFolder.createFolder(folderName);
}

/* ======================================================================
   MODUL: UPDATE / EDIT CUTI - FIXED
   ====================================================================== */

function updatePengajuanCuti(payload) {
  try {
    var rowIndex = parseInt(payload.rowBaris);
    var errorBentrok = cekBentrokCuti(payload.nip, payload.tglMulai, payload.tglSelesai, rowIndex);
    if (errorBentrok) return errorBentrok;
    var ss = SpreadsheetApp.openById(ID_SS_CUTI);
    var sheet = ss.getSheetByName("Form Cuti");
    if (!sheet) return "Error: Sheet 'Form Cuti' tidak ditemukan.";
    
    if (!rowIndex || rowIndex < 2) return "Error: Baris data tidak valid.";

    var now = new Date();
    var tglEditStr = "'" + Utilities.formatDate(now, Session.getScriptTimeZone(), "dd-MM-yyyy HH:mm:ss");
    var userEdit = payload.userInput || "User Web";

    // 1. FORMAT TANGGAL
    var tglMulaiIndo   = formatIndoText(payload.tglMulai);
    var tglSelesaiIndo = formatIndoText(payload.tglSelesai);
    
    // --- PERBAIKAN: Gunakan payload.tglPengajuan ---
    var tglPengajuanFormat = formatTglIndo(payload.tglPengajuan);

    // 2. AMBIL DATA DETIL PEGAWAI
    var dbData = getDetailPegawaiByNip(payload.nip); 
    var empGol = dbData ? dbData.golongan : ""; 
    var empJab = dbData ? dbData.jabatan : ""; 
    
    // 3. LOOKUP PEJABAT
    var pejabat = lookupPejabatStruktural(payload.jenisCuti, payload.unit, empGol, empJab);
    
    var final_kepada, final_nama_atasan, final_nip_atasan, final_jab_atasan, final_nama_setuju, final_nip_setuju, final_jab_setuju;

    if (pejabat) {
        final_kepada      = pejabat.kepada;
        final_nama_atasan = pejabat.nama_atasan;
        final_nip_atasan  = pejabat.nip_atasan;
        final_jab_atasan  = pejabat.jabatan_atasan;
        final_nama_setuju = pejabat.nama_setuju;
        final_nip_setuju  = pejabat.nip_setuju;
        final_jab_setuju  = pejabat.jabatan_setuju;
    } else {
        final_kepada      = dbData ? dbData.fullRow[19] : ""; 
        final_nama_atasan = dbData ? dbData.fullRow[13] : ""; 
        final_nip_atasan  = dbData ? dbData.fullRow[14] : ""; 
        final_jab_atasan  = dbData ? dbData.fullRow[15] : ""; 
        final_nama_setuju = dbData ? dbData.fullRow[16] : ""; 
        final_nip_setuju  = dbData ? dbData.fullRow[17] : ""; 
        final_jab_setuju  = dbData ? dbData.fullRow[18] : ""; 
    }

    // 4. LOGIKA SISA CUTI
    var thnMulai = parseInt(payload.tglMulai.split("-")[0]);
    var sisaN2 = "-", sisaN1 = "-", sisaN = "-";
    if (dbData) {
        if (thnMulai === 2023) { sisaN2 = dbData.fullRow[26]||"0"; sisaN1 = dbData.fullRow[28]||"0"; sisaN = dbData.fullRow[30]||"0"; }
        else if (thnMulai === 2024) { sisaN2 = dbData.fullRow[38]||"0"; sisaN1 = dbData.fullRow[40]||"0"; sisaN = dbData.fullRow[42]||"0"; }
        else if (thnMulai === 2025) { sisaN2 = dbData.fullRow[50]||"0"; sisaN1 = dbData.fullRow[52]||"0"; sisaN = dbData.fullRow[54]||"0"; }
        else if (thnMulai === 2026) { sisaN2 = dbData.fullRow[62]||"0"; sisaN1 = dbData.fullRow[64]||"0"; sisaN = dbData.fullRow[66]||"0"; }
    }

    // 5. CHECKLIST JENIS CUTI
    var j = String(payload.jenisCuti).toLowerCase().trim();
    var c = { ct:"", cs:"", cap:"", cb:"", cm:"", cltn:"" }; 
    var CHECK = "✓"; 
    
    if (j.includes("sakit")) c.cs = CHECK;
    else if (j.includes("penting")) c.cap = CHECK;
    else if (j.includes("besar")) c.cb = CHECK;
    else if (j.includes("melahirkan")) c.cm = CHECK;
    else if (j.includes("luar") || j.includes("tanggungan")) c.cltn = CHECK;
    else c.ct = CHECK;

    // 6. SUSUN DATA PDF
    var pdfData = {
        // Gunakan Tanggal Pengajuan dari User untuk Tanggal Surat
        tanggal: tglPengajuanFormat, 
        kepada: final_kepada,
        
        asn: payload.nama,
        nip: payload.nip,
        jabatan: dbData ? dbData.jabatan : "", 
        masa_kerja: dbData ? dbData.masaKerja : "", 
        unit: dbData ? dbData.unitLengkap : payload.unit, 
        
        alasan: payload.alasan,
        jumlah: payload.jumlahHari,
        tmc: tglMulaiIndo,
        tsc: tglSelesaiIndo,
        
        "N-2": sisaN2, "N-1": sisaN1, "N": sisaN,
        
        alamat: payload.alamat,
        telp: payload.hp,
        
        jabatan_atasan: final_jab_atasan,
        nama_atasan: final_nama_atasan,
        nip_atasan: final_nip_atasan,
        
        jabatan_setuju: final_jab_setuju,
        nama_setuju: final_nama_setuju,
        nip_setuju: final_nip_setuju,
        
        ct: c.ct, cs: c.cs, cap: c.cap, cb: c.cb, cm: c.cm, cltn: c.cltn,
        
        jenisCutiRaw: payload.jenisCuti, 
        tglMulaiRaw: payload.tglMulai    
    };
    
    // GENERATE PDF
    var linkPdf = generatePdfCuti(pdfData); 

    // 7. UPDATE SPREADSHEET
    var rangeUtama = sheet.getRange(rowIndex, 1, 1, 10);
    rangeUtama.setValues([[
        payload.unit, payload.nama, "'" + payload.nip, payload.jenisCuti, 
        tglMulaiIndo, tglSelesaiIndo, payload.jumlahHari, payload.alasan, 
        payload.alamat, "'" + payload.hp
    ]]);

    sheet.getRange(rowIndex, 11, 1, 3).setValues([["Diproses", "", linkPdf]]);
    sheet.getRange(rowIndex, 16, 1, 2).setValues([[tglEditStr, userEdit]]);

    var rangeExtra = sheet.getRange(rowIndex, 22, 1, 20);
    rangeExtra.setValues([[
      tglPengajuanFormat, // Kolom V
      pdfData.jabatan,    // W
      pdfData.masa_kerja, // X
      pdfData.unit,       // Y
      c.ct, c.cb, c.cs, c.cm, c.cap, c.cltn, 
      sisaN2, sisaN1, sisaN,
      final_jab_atasan, final_nama_atasan, final_nip_atasan,
      final_jab_setuju, final_nama_setuju, final_nip_setuju,
      final_kepada       
    ]]);

    SpreadsheetApp.flush();
    return "Sukses";
    
  } catch (e) { return "Error Update: " + e.toString(); }
}

function hapusPengajuanCuti(rowBaris, kodeInput, userDelete) {
  try {
    var now = new Date();
    var validCode = Utilities.formatDate(now, Session.getScriptTimeZone(), "yyyyMMdd");
    
    if (String(kodeInput).trim() !== validCode) {
      return "KODE_SALAH";
    }

    var ss = SpreadsheetApp.openById(ID_SS_CUTI);
    var sheet = ss.getSheetByName("Form Cuti");
    if (!sheet) return "Error: Sheet tidak ditemukan.";
    
    var row = parseInt(rowBaris);
    if (isNaN(row) || row < 2) return "Error: Baris data tidak valid.";
    
    sheet.deleteRow(row);
    
    return "Sukses";
    
  } catch (e) { return "Error Hapus: " + e.toString(); }
}

function verifikasiPengajuan(rowBaris, status, catatan, adminName) {
  try {
    var ss = SpreadsheetApp.openById(ID_SS_CUTI);
    var sheet = ss.getSheetByName("Form Cuti");
    if (!sheet) return "Error: Sheet tidak ditemukan.";
    
    var row = parseInt(rowBaris);
    if (isNaN(row) || row < 2) return "Error: Baris tidak valid.";

    sheet.getRange(row, 11).setValue(status);
    sheet.getRange(row, 12).setValue(catatan);
    
    var now = new Date();
    var sysDateStr = "'" + Utilities.formatDate(now, Session.getScriptTimeZone(), "dd-MM-yyyy HH:mm:ss");
    
    sheet.getRange(row, 18).setValue(sysDateStr);
    sheet.getRange(row, 19).setValue(adminName || "Admin");
    
    return "Sukses";
    
  } catch (e) { return "Error Verif: " + e.toString(); }
}

function getUnitOptions() {
  var ss = SpreadsheetApp.openById(ID_SS_CUTI);
  var sheet = ss.getSheetByName("Database Cuti"); 
  if (!sheet) return [];
  
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  
  var data = sheet.getRange(2, 2, lastRow - 1, 1).getValues();
  var uniqueUnits = [];
  var seen = {};
  
  for (var i = 0; i < data.length; i++) {
    var unit = String(data[i][0]).trim();
    if (unit !== "" && !seen[unit]) {
      uniqueUnits.push(unit);
      seen[unit] = true;
    }
  }
  
  uniqueUnits.sort();
  return uniqueUnits;
}

function getDataCuti(tahun, bulan, unitFilter) { 
  var ss = SpreadsheetApp.openById(ID_SS_CUTI);
  var sheet = ss.getSheetByName("Form Cuti");
  if (!sheet) return JSON.stringify([]);
  
  var dataRaw = sheet.getDataRange().getValues();
  var dataDisplay = sheet.getDataRange().getDisplayValues(); 
  var result = [];
  
  var fTahun = tahun ? String(tahun).trim() : "";
  var fBulan = bulan ? String(bulan).toLowerCase().trim() : "";
  var fUnit  = unitFilter ? String(unitFilter).toLowerCase().trim() : "";

  for (var i = 1; i < dataRaw.length; i++) {
    var row = dataRaw[i];       
    var rowTxt = dataDisplay[i];
    
    var rowUnitRaw = String(row[0]).toLowerCase();
    var tglMulaiTxt = String(rowTxt[4]).trim(); 
    
    var rTahun = "";
    var rBulan = "";
    var parts = tglMulaiTxt.split(" ");
    if (parts.length >= 3) {
       rBulan = parts[1].toLowerCase();
       rTahun = parts[2];
    } else if (row[4] instanceof Date) {
       rTahun = String(row[4].getFullYear());
       var mNames = ["januari","februari","maret","april","mei","juni","juli","agustus","september","oktober","november","desember"];
       rBulan = mNames[row[4].getMonth()];
    }

    var matchTahun = (fTahun === "") || (rTahun === fTahun);
    var matchBulan = (fBulan === "") || (rBulan === fBulan);
    var matchUnit  = (fUnit === "")  || (rowUnitRaw.indexOf(fUnit) > -1);

    if (matchTahun && matchBulan && matchUnit) {
      result.push({
        rowBaris: i + 1,
        unit: rowTxt[0],
        nama: rowTxt[1],
        nip:  rowTxt[2],
        jenis: rowTxt[3],
        tglMulai: rowTxt[4],
        tglSelesai: rowTxt[5],
        jumlah: rowTxt[6],
        alasan: rowTxt[7],
        alamat: rowTxt[8],
        telepon: rowTxt[9],
        status: rowTxt[10],
        ket: rowTxt[11],
        fileUrl: rowTxt[12],
        
        tglInput: rowTxt[13], 
        userInput: rowTxt[14],
        tglEdit: rowTxt[15],  
        userEdit: rowTxt[16],
        tglVerif: rowTxt[17], 
        verifikator: rowTxt[18],
        // Tambahkan Kolom V (Tanggal Pengajuan) ke response jika perlu ditampilkan di tabel
        tanggal: rowTxt[21] 
      });
    }
  }
  
  result.sort(function(a, b) {
      function getMs(str) {
          if (!str || str.length < 10) return 0;
          try {
            var parts = str.split(" ");     
            var d = parts[0].split("-");    
            var t = parts[1].split(":");    
            return new Date(d[2], d[1]-1, d[0], t[0], t[1], t[2]).getTime();
          } catch(e) { return 0; }
      }
      var maxA = Math.max(getMs(a.tglInput), getMs(a.tglEdit), getMs(a.tglVerif));
      var maxB = Math.max(getMs(b.tglInput), getMs(b.tglEdit), getMs(b.tglVerif));
      return maxB - maxA;
  });
  
  return JSON.stringify(result);
}

// --- HELPER FORMAT TANGGAL INDO (FINAL) ---
function formatTglIndo(strDate) {
  if(!strDate) return "";
  var d = new Date(strDate);
  var months = ["Januari", "Februari", "Maret", "April", "Mei", "Juni", "Juli", "Agustus", "September", "Oktober", "November", "Desember"];
  return d.getDate() + " " + months[d.getMonth()] + " " + d.getFullYear();
}

/* ======================================================================
   MODUL: SISA CUTI TAHUNAN
   ====================================================================== */
function getSisaCutiData() {
  var ID_SS_SISA = "1UYG80gGxuC19ieaVBzJaUV8bhlS2q5gExr0-Yl7upKo";
  
  try {
    var ss = SpreadsheetApp.openById(ID_SS_SISA);
    var sheet = ss.getSheetByName("Sisa CT");
    if (!sheet) return JSON.stringify({ error: "Sheet 'Sisa CT' tidak ditemukan." });
    
    var lastRow = sheet.getLastRow();
    if (lastRow < 1) return JSON.stringify({ headers: [], data: [] });
    
    // Ambil Data dari Kolom A sampai M (13 Kolom)
    // Gunakan getDisplayValues() agar data persis seperti tampilan di Excel
    var range = sheet.getRange(1, 1, lastRow, 13); 
    var rawValues = range.getDisplayValues(); 
    
    // Pisahkan Header (Baris 1) dan Data (Baris 2 dst)
    var headers = rawValues[0];
    var rows = rawValues.slice(1);
    
    // URUTKAN ABJAD BERDASARKAN NAMA (Kolom B -> Index 1)
    rows.sort(function(a, b) {
       var valA = String(a[1]).toLowerCase();
       var valB = String(b[1]).toLowerCase();
       if (valA < valB) return -1;
       if (valA > valB) return 1;
       return 0;
    });

    return JSON.stringify({
      headers: headers,
      data: rows
    });
    
  } catch (e) {
    return JSON.stringify({ error: "Error Server: " + e.toString() });
  }
}

/* ======================================================================
   MODUL: REKAP CUTI (MULTI-SHEET)
   ====================================================================== */

function getRekapYears() {
  var ID_MASTER = "1UYG80gGxuC19ieaVBzJaUV8bhlS2q5gExr0-Yl7upKo"; // ID Spreadsheet Master
  try {
    var ss = SpreadsheetApp.openById(ID_MASTER);
    var sheet = ss.getSheetByName("Jumlah Cuti");
    if (!sheet) return [];
    
    // Ambil Kolom A (Tahun) & B (Sheet ID/Nama Tab) mulai baris 2
    var lastRow = sheet.getLastRow();
    if (lastRow < 2) return [];
    
    var data = sheet.getRange(2, 1, lastRow - 1, 2).getDisplayValues();
    // Filter baris kosong
    var result = data.filter(function(row) { return row[0] !== "" && row[1] !== ""; });
    
    return result.map(function(r) { return { tahun: r[0], id: r[1] }; });
  } catch (e) { return []; }
}

function getRekapData(targetInput) {
  var ID_MASTER = "1UYG80gGxuC19ieaVBzJaUV8bhlS2q5gExr0-Yl7upKo";
  var ss, sheet;

  // 1. Coba cari sebagai Nama Tab di dalam File Master dulu
  try {
    ss = SpreadsheetApp.openById(ID_MASTER);
    sheet = ss.getSheetByName(targetInput);
  } catch(e) {}

  // 2. Jika tidak ketemu, anggap input adalah ID Spreadsheet terpisah
  if (!sheet) {
      try {
        ss = SpreadsheetApp.openById(targetInput);
        sheet = ss.getSheets()[0]; // Ambil sheet pertama
      } catch(e) {
        return JSON.stringify({ error: "Gagal membuka data. Pastikan Nama Tab atau ID File benar." });
      }
  }
  
  try {
    var lastRow = sheet.getLastRow();
    if (lastRow < 3) return JSON.stringify({ h1:[], h2:[], data: [] });
    
    // Ambil Data A sampai O (15 Kolom)
    // Baris 1: Header Utama, Baris 2: Sub Header, Baris 3+: Data
    var range = sheet.getRange(1, 1, lastRow, 15);
    var rawValues = range.getDisplayValues();
    
    var h1 = rawValues[0];
    var h2 = rawValues[1];
    var dataRows = rawValues.slice(2);
    
    return JSON.stringify({ h1: h1, h2: h2, data: dataRows });
  } catch (e) {
    return JSON.stringify({ error: "Error: " + e.toString() });
  }
}

/* ======================================================================
   MODUL: UNGGAH SURAT CUTI (UPDATE LOGIC)
   ====================================================================== */

function getUnitOptionsUnggah() {
  var ID_SS = "1UYG80gGxuC19ieaVBzJaUV8bhlS2q5gExr0-Yl7upKo";
  try {
    var ss = SpreadsheetApp.openById(ID_SS);
    var sheet = ss.getSheetByName("Form Cuti"); // Spesifik Sheet Form Cuti
    if (!sheet) return [];
    
    var lastRow = sheet.getLastRow();
    if (lastRow < 2) return [];
    
    // Ambil Kolom A (Unit Kerja) dari Form Cuti
    var data = sheet.getRange(2, 1, lastRow - 1, 1).getDisplayValues();
    var unique = {};
    var result = [];
    
    for (var i = 0; i < data.length; i++) {
      var unit = String(data[i][0]).trim();
      if (unit && !unique[unit]) {
        unique[unit] = true;
        result.push(unit);
      }
    }
    
    result.sort();
    return result;
  } catch (e) {
    return [];
  }
}

function getDaftarUnggahCuti(tahun, bulan, unit, status) {
  var ID_SS = "1UYG80gGxuC19ieaVBzJaUV8bhlS2q5gExr0-Yl7upKo";
  var SHEET_NAME = "Form Cuti";
  
  try {
    var ss = SpreadsheetApp.openById(ID_SS);
    var sheet = ss.getSheetByName(SHEET_NAME);
    if (!sheet) return JSON.stringify([]);

    var lastRow = sheet.getLastRow();
    if (lastRow < 2) return JSON.stringify([]);
    
    // Ambil Data A s.d AX (Index 0 s.d 49)
    var data = sheet.getRange(2, 1, lastRow - 1, 50).getDisplayValues();
    var result = [];

    // Filter Helper
    var fTahun  = (tahun && String(tahun).trim() !== "") ? String(tahun).trim() : null;
    var fBulan  = (bulan && String(bulan).trim() !== "") ? String(bulan).trim() : null;
    var fUnit   = (unit && String(unit).trim() !== "" && unit !== "SEMUA") ? String(unit).trim() : null;
    var fStatus = (status && String(status).trim() !== "" && status !== "SEMUA") ? String(status).trim() : null;
    var arrBulanIndo = ["Januari", "Februari", "Maret", "April", "Mei", "Juni", "Juli", "Agustus", "September", "Oktober", "November", "Desember"];

    // --- HELPER: Parse Segala Jenis Tanggal ke Timestamp ---
    // Menerima format: "12 Januari 2026", "22-01-2026 14:00:00", atau "2026-01-01"
    function parseToTs(strDate) {
      if (!strDate || strDate === "") return 0;
      strDate = String(strDate).trim();

      // 1. Cek Format System Timestamp (dd-MM-yyyy HH:mm:ss) -> Kolom AR, AT, AV
      if (strDate.indexOf(":") > -1 && strDate.indexOf("-") > -1 && strDate.length > 10) {
         // Contoh: 22-01-2026 14:30:15
         var parts = strDate.split(" ");
         var dPart = parts[0].split("-");
         var tPart = parts[1].split(":");
         return new Date(dPart[2], dPart[1]-1, dPart[0], tPart[0], tPart[1], tPart[2]).getTime();
      }

      // 2. Cek Format Teks Indonesia (dd MMMM yyyy) -> Kolom E
      if (strDate.indexOf(" ") > -1) {
         var p = strDate.split(" ");
         if (p.length >= 3) {
             var thn = 0, bln = 0, tgl = 0;
             // Cari Tahun (4 digit)
             for(var i=0; i<p.length; i++) {
                 if(!isNaN(p[i]) && p[i].length === 4) thn = parseInt(p[i]);
                 else if(arrBulanIndo.indexOf(p[i]) > -1) bln = arrBulanIndo.indexOf(p[i]);
                 else if(!isNaN(p[i]) && p[i].length <= 2) tgl = parseInt(p[i]);
             }
             if (thn > 0) return new Date(thn, bln, tgl).getTime();
         }
      }

      // 3. Fallback ke Date Parsing standar
      var std = new Date(strDate).getTime();
      return isNaN(std) ? 0 : std;
    }

    for (var i = 0; i < data.length; i++) {
        var row = data[i];
        
        // Validasi Dasar
        if (!row[1]) continue; 
        if (String(row[10]).trim().toLowerCase() !== "disetujui") continue;

        // --- FILTER ---
        var rawTgl = String(row[4]).trim(); // Tgl Mulai (Col E)
        var rTahun = "", rBulan = "";
        
        // Ekstrak Tahun/Bulan dari Tgl Mulai untuk Filter
        if (rawTgl.indexOf(" ") > -1) {
             var p = rawTgl.split(" ");
             if (p.length >= 3) {
                 for(var x=0; x<p.length; x++) {
                     if(!isNaN(p[x]) && p[x].length === 4) rTahun = p[x];
                     if(arrBulanIndo.indexOf(p[x]) > -1) rBulan = p[x];
                 }
             }
        }
        
        if (fTahun && rTahun !== fTahun) continue;
        if (fBulan && rBulan !== fBulan) continue;
        if (fUnit && String(row[0]).trim() !== fUnit) continue;
        
        var stUnggah = String(row[42]).trim(); 
        if (fStatus) {
            if (fStatus === "Belum" && stUnggah !== "") continue;
            if (fStatus !== "Belum" && stUnggah !== fStatus) continue;
        }

        // --- HITUNG LAST ACTIVITY ---
        // Kita cari waktu paling maksimum dari 4 kejadian:
        var tsMulai  = parseToTs(row[4]);  // E (Tanggal Mulai Cuti)
        var tsUnggah = parseToTs(row[43]); // AR (Tanggal Upload)
        var tsEdit   = parseToTs(row[45]); // AT (Tanggal Edit)
        var tsVerif  = parseToTs(row[47]); // AV (Tanggal Verif)

        // Ambil nilai terbesar (terbaru)
        var lastActivityTs = Math.max(tsMulai, tsUnggah, tsEdit, tsVerif);

        result.push({
            rowBaris: i + 2,
            unit: row[0],
            nama: row[1],
            nip: row[2],
            jenis: row[3],
            tglMulai: row[4],
            tglSelesai: row[5],
            jumlah: row[6],
            
            fileUrl: row[41],      
            statusUnggah: row[42], 
            
            tglUnggah: row[43],    
            userUnggah: row[44],   
            tglEdit: row[45],      
            userEdit: row[46],     
            tglVerif: row[47],     
            verifikator: row[48],  
            ket: row[49],
            
            // Simpan timestamp untuk sorting
            lastActivity: lastActivityTs
        });
    }

    // --- SORTING FINAL: Last Activity Descending (Terbaru di Atas) ---
    result.sort(function(a, b) { 
        // Jika timestamp sama, fallback ke urutan baris (ID)
        if (b.lastActivity === a.lastActivity) {
            return b.rowBaris - a.rowBaris;
        }
        return b.lastActivity - a.lastActivity; 
    });

    return JSON.stringify(result);

  } catch (e) {
    return JSON.stringify([{ error: e.toString() }]);
  }
}

function hapusUnggahCuti(recId, userLogin) {
  var ID_SS = "1UYG80gGxuC19ieaVBzJaUV8bhlS2q5gExr0-Yl7upKo";
  var SHEET_NAME = "Form Cuti";
  try {
    var ss = SpreadsheetApp.openById(ID_SS);
    var sheet = ss.getSheetByName(SHEET_NAME);
    var row = parseInt(recId);
    // Hapus konten kolom AP s.d AX (9 kolom)
    sheet.getRange(row, 42, 1, 9).clearContent();
    return "Sukses";
  } catch (e) { throw new Error(e.message); }
}

function simpanUnggahSurat(form, fileData) {
  var ID_SS = "1UYG80gGxuC19ieaVBzJaUV8bhlS2q5gExr0-Yl7upKo";
  var ID_FOLDER = "1uPeOU7F_mgjZVyOLSsj-3LXGdq9rmmWl"; // Folder Cuti
  var SHEET_NAME = "Form Cuti";

  try {
    var ss = SpreadsheetApp.openById(ID_SS);
    var sheet = ss.getSheetByName(SHEET_NAME);
    var row = parseInt(form.recId);
    
    // Validasi Row
    if (isNaN(row) || row < 2) throw new Error("Data tidak valid.");

    var now = new Date();
    var sysDateStr = "'" + Utilities.formatDate(now, Session.getScriptTimeZone(), "dd-MM-yyyy HH:mm:ss");
    var userName = form.user_login || "User Web";

    // 1. PROSES FILE
    var fileUrl = "";
    if (fileData && fileData.data) {
        var folder = DriveApp.getFolderById(ID_FOLDER);
        // Format Nama File: SURAT_CUTI - NAMA - JENIS - TGL.pdf
        var namaFile = "SURAT_CUTI - " + form.nama + " - " + form.jenis + ".pdf";
        var blob = Utilities.newBlob(Utilities.base64Decode(fileData.data), fileData.mimeType, namaFile);
        var file = folder.createFile(blob);
        file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
        fileUrl = file.getUrl();
    } else {
        throw new Error("File wajib diunggah.");
    }

    // 2. TENTUKAN KOLOM UPDATE
    // AP=42 (index 42 di getRange?), AQ=43...
    // AP adalah kolom ke-42.
    
    // Cek apakah ini Update (Edit) atau Baru
    var oldStatus = sheet.getRange(row, 43).getValue(); // AQ (Status)
    var isEdit = (oldStatus !== "" && oldStatus !== null);

    if (isEdit) {
        // UPDATE MODE: Update File (AP), Status (AQ), Log Edit (AT, AU)
        sheet.getRange(row, 42).setValue(fileUrl);      // AP
        sheet.getRange(row, 43).setValue("Diproses");   // AQ (Reset Status)
        sheet.getRange(row, 46).setValue(sysDateStr);   // AT (Tgl Edit)
        sheet.getRange(row, 47).setValue(userName);     // AU (User Edit)
        // Kosongkan Verif
        sheet.getRange(row, 48).setValue("");           // AV
        sheet.getRange(row, 49).setValue("");           // AW
        sheet.getRange(row, 50).setValue(""); // AX (Hapus Keterangan Lama)
    } else {
        // NEW MODE: Update File (AP), Status (AQ), Log Unggah (AR, AS)
        sheet.getRange(row, 42).setValue(fileUrl);      // AP
        sheet.getRange(row, 43).setValue("Diproses");   // AQ
        sheet.getRange(row, 44).setValue(sysDateStr);   // AR (Tgl Unggah)
        sheet.getRange(row, 45).setValue(userName);     // AS (User Unggah)
    }

    return "Sukses";

  } catch (e) {
    throw new Error("Gagal Unggah: " + e.message);
  }
}

function verifikasiUnggahSurat(form) {
  var ID_SS = "1UYG80gGxuC19ieaVBzJaUV8bhlS2q5gExr0-Yl7upKo";
  var SHEET_NAME = "Form Cuti";

  try {
    var ss = SpreadsheetApp.openById(ID_SS);
    var sheet = ss.getSheetByName(SHEET_NAME);
    var row = parseInt(form.recId);

    var now = new Date();
    var sysDateStr = "'" + Utilities.formatDate(now, Session.getScriptTimeZone(), "dd-MM-yyyy HH:mm:ss");

    // Update Status (AQ / Col 43)
    sheet.getRange(row, 43).setValue(form.status);
    
    // Update Verif Log (AV / 48, AW / 49, AX / 50)
    sheet.getRange(row, 48).setValue(sysDateStr);      // AV
    sheet.getRange(row, 49).setValue(form.user_verif); // AW
    sheet.getRange(row, 50).setValue(form.ket);        // AX (Keterangan)

    return "Sukses";
  } catch (e) {
    throw new Error("Gagal Verifikasi: " + e.message);
  }
}

/* ======================================================================
   DASHBOARD SIABA: MODULAR & PARALLEL (FAST LOAD)
   ====================================================================== */

// Helper Configuration
function getConfigSiaba() {
  return {
    cuti:  { id: "1UYG80gGxuC19ieaVBzJaUV8bhlS2q5gExr0-Yl7upKo", sheet: "Form Cuti",      dateCol: 4, statCol: 10, nameCol: 1 },
    dinas: { id: "1I_2yUFGXnBJTCSW6oaT3D482YCs8TIRkKgQVBbvpa1M", sheet: "Perjalanan_Dinas", dateCol: 3, statCol: 9,  nameCol: 1 },
    lupa:  { id: "160IjN8aiDAgDYXjgDLStS4nCZLKn3Ny-dq3BOFAfDrU", sheet: "Lupa_Presensi",    dateCol: 3, statCol: 10, nameCol: 1 },
    salah: { id: "1TZGrMiTuyvh2Xbo44RhJuWlQnOC5LzClsgIoNKtRFkY", sheet: "Salah_Presensi",   dateCol: 3, statCol: 8,  nameCol: 1 },
    rekap: { id: "1tQsQY1-Ny1ie66GOZPTLtvZ7BiYCgFdNrX-AVGCtaHA" }
  };
}

// 1. FUNGSI GET METRIC (Dipanggil 4x secara paralel oleh Frontend)
function getSiabaMetric(type) {
  var conf = getConfigSiaba()[type];
  if (!conf) return JSON.stringify({ error: "Invalid Type" });

  var cache = CacheService.getScriptCache();
  var cacheKey = "dash_metric_v2_" + type;
  var cached = cache.get(cacheKey);
  if (cached) return cached;

  var result = { 
    type: type,
    total: 0, bulanIni: 0, setuju: 0, tolak: 0, proses: 0, revisi: 0 
  };

  try {
    var ss = SpreadsheetApp.openById(conf.id);
    var sheet = ss.getSheetByName(conf.sheet);
    if (!sheet) return JSON.stringify(result);

    var data = sheet.getDataRange().getValues();
    var now = new Date();
    var curYear = now.getFullYear();
    var curMonth = now.getMonth();

    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      if (!row[conf.nameCol]) continue; // Skip nama kosong

      var rawDate = row[conf.dateCol];
      var eventDate = parseEventDateSimple(rawDate); // Gunakan parser ringan

      if (!eventDate || eventDate.getFullYear() !== curYear) continue;

      var status = String(row[conf.statCol] || "").toLowerCase();

      result.total++;
      if (eventDate.getMonth() === curMonth) {
        result.bulanIni++;
      }

      if (status.includes("setuju") || status.includes("ok") || status.includes("acc")) result.setuju++;
      else if (status.includes("tolak")) result.tolak++;
      else if (status.includes("revisi") || status.includes("ubah")) result.revisi++;
      else result.proses++;
    }
  } catch (e) { result.error = e.toString(); }

  var json = JSON.stringify(result);
  cache.put(cacheKey, json, 120); // Cache 2 menit
  return json;
}

// 2. FUNGSI GET CHART TREN (Dipanggil terpisah)
function getSiabaChartTrend() {
  var cache = CacheService.getScriptCache();
  var cacheKey = "dash_chart_trend_v2";
  var cached = cache.get(cacheKey);
  if (cached) return cached;

  var conf = getConfigSiaba()['rekap'];
  var result = {
    labels: ["Jan", "Feb", "Mar", "Apr", "Mei", "Jun", "Jul", "Ags", "Sep", "Okt", "Nov", "Des"],
    terlambat: [0,0,0,0,0,0,0,0,0,0,0,0],
    pulangAwal: [0,0,0,0,0,0,0,0,0,0,0,0]
  };

  try {
    var ss = SpreadsheetApp.openById(conf.id);
    var curYearStr = new Date().getFullYear().toString();

    function processSheet(sheetName, targetKey) {
      var sheet = ss.getSheetByName(sheetName);
      if(!sheet) return;
      var data = sheet.getDataRange().getDisplayValues();
      for (var i = 2; i < data.length; i++) {
        if (String(data[i][0]).trim() === curYearStr) {
          for (var m = 0; m < 12; m++) {
            var val = data[i][4 + (m * 2)];
            if (val && val !== "0" && val !== "-" && val.trim() !== "") {
               result[targetKey][m]++;
            }
          }
        }
      }
    }

    processSheet("Rekap_Terlambat", "terlambat");
    processSheet("Rekap_Pulang_Awal", "pulangAwal");

  } catch (e) { result.error = e.toString(); }

  var json = JSON.stringify(result);
  cache.put(cacheKey, json, 300); // Cache 5 menit
  return json;
}

// Helper Parser Ringan (Tanpa Regex Berat)
function parseEventDateSimple(raw) {
  if (raw instanceof Date) return raw;
  if (!raw) return null;
  var str = String(raw).trim();
  
  // Deteksi Format: "21 Januari 2026"
  var months = ["januari","februari","maret","april","mei","juni","juli","agustus","september","oktober","november","desember"];
  var parts = str.split(' ');
  if (parts.length >= 3) {
    var d = parseInt(parts[0]);
    var mStr = parts[1].toLowerCase();
    var y = parseInt(parts[2]);
    var m = months.indexOf(mStr);
    if (m > -1 && !isNaN(d) && !isNaN(y)) return new Date(y, m, d);
  }
  
  // Fallback ke standard date
  var dObj = new Date(raw);
  return isNaN(dObj.getTime()) ? null : dObj;
}

/* ======================================================================
   DASHBOARD BACKEND: ASYNCHRONOUS / PARALLEL MODE
   Setiap fungsi mengambil data spesifik agar frontend tidak menunggu lama.
   ====================================================================== */

// 1. CONFIG ID DATABASE
var DB_ID = {
  CUTI:  "1UYG80gGxuC19ieaVBzJaUV8bhlS2q5gExr0-Yl7upKo",
  DINAS: "1I_2yUFGXnBJTCSW6oaT3D482YCs8TIRkKgQVBbvpa1M",
  LUPA:  "160IjN8aiDAgDYXjgDLStS4nCZLKn3Ny-dq3BOFAfDrU",
  SALAH: "1TZGrMiTuyvh2Xbo44RhJuWlQnOC5LzClsgIoNKtRFkY",
  REKAP: "1tQsQY1-Ny1ie66GOZPTLtvZ7BiYCgFdNrX-AVGCtaHA"
};

// 2. HELPER PARSER TANGGAL
function parseDateIso(v) {
  if(!v) return null;
  if(v instanceof Date) return v;
  var s = String(v).trim().replace(/'/g,'');
  var m = s.match(/^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{4})/);
  if(m) return new Date(m[3], m[2]-1, m[1]);
  return new Date(s);
}

// 3. API: AMBIL DATA PER MODUL (Dipanggil Terpisah oleh Frontend)
function getSiabaMetric(type) {
  var cache = CacheService.getScriptCache();
  var cacheKey = "metric_v2_" + type;
  var cached = cache.get(cacheKey);
  if (cached) return cached;

  var now = new Date();
  var curYear = now.getFullYear();
  var curMonth = now.getMonth();

  // Config per tipe
  var config = {};
  if (type == 'cuti')  config = { id: DB_ID.CUTI,  tab: "Form Cuti",        idxTgl: 4, idxStat: 10 };
  if (type == 'dinas') config = { id: DB_ID.DINAS, tab: "Perjalanan_Dinas", idxTgl: 3, idxStat: 9 };
  if (type == 'lupa')  config = { id: DB_ID.LUPA,  tab: "Lupa_Presensi",    idxTgl: 3, idxStat: 10 };
  if (type == 'salah') config = { id: DB_ID.SALAH, tab: "Salah_Presensi",   idxTgl: 3, idxStat: 8 };

  var res = {
    type: type,
    total: 0, 
    bulanIni: 0,
    proses: 0, revisi: 0, setuju: 0, tolak: 0, // Untuk Pie Chart & List
    trend: { // Untuk Grafik Bar per Modul
      proses: new Array(12).fill(0),
      revisi: new Array(12).fill(0),
      setuju: new Array(12).fill(0),
      tolak:  new Array(12).fill(0)
    }
  };

  try {
    var ss = SpreadsheetApp.openById(config.id);
    var sh = ss.getSheetByName(config.tab);
    if (sh) {
      var data = sh.getDataRange().getValues();
      for (var i = 1; i < data.length; i++) {
        var row = data[i];
        var tgl = parseDateIso(row[config.idxTgl]);
        
        if (!tgl || tgl.getFullYear() !== curYear) continue;

        var bln = tgl.getMonth();
        var st = String(row[config.idxStat]||"").toLowerCase();

        res.total++;
        if (bln === curMonth) res.bulanIni++;

        if (st.includes("setuju") || st.includes("ok") || st.includes("disetujui")) {
          res.setuju++; res.trend.setuju[bln]++;
        } else if (st.includes("tolak") || st.includes("tidak")) {
          res.tolak++; res.trend.tolak[bln]++;
        } else if (st.includes("revisi") || st.includes("ubah")) {
          res.revisi++; res.trend.revisi[bln]++;
        } else {
          res.proses++; res.trend.proses[bln]++;
        }
      }
    }
  } catch (e) { res.error = e.message; }

  var json = JSON.stringify(res);
  cache.put(cacheKey, json, 120); // Cache 2 menit per modul
  return json;
}

// 4. API: AMBIL DATA GRAFIK TREN (Terpisah)
function getSiabaChartTrend() {
  var cache = CacheService.getScriptCache();
  var cached = cache.get("chart_trend_v2");
  if (cached) return cached;

  var curYear = new Date().getFullYear();
  var res = {
    labels: ["Jan", "Feb", "Mar", "Apr", "Mei", "Jun", "Jul", "Ags", "Sep", "Okt", "Nov", "Des"],
    terlambat: new Array(12).fill(0),
    pulangAwal: new Array(12).fill(0)
  };

  try {
    var ss = SpreadsheetApp.openById(DB_ID.REKAP);
    ["Rekap_Terlambat", "Rekap_Pulang_Awal"].forEach(function(nm) {
      var sh = ss.getSheetByName(nm);
      if (sh) {
        var data = sh.getDataRange().getDisplayValues();
        var key = nm.includes("Terlambat") ? "terlambat" : "pulangAwal";
        for (var i = 2; i < data.length; i++) {
          if (String(data[i][0]).trim() == curYear) {
            for (var m = 0; m < 12; m++) {
              res[key][m] += (parseInt(data[i][4 + (m * 2)]) || 0);
            }
          }
        }
      }
    });
  } catch (e) {}

  var json = JSON.stringify(res);
  cache.put("chart_trend_v2", json, 300);
  return json;
}

/* ======================================================================
   [ADD-ON] KHUSUS DATA TREN BULANAN
   Fungsi ini berdiri sendiri agar tidak mengganggu kode utama.
   ====================================================================== */
function getTrenBulananData() {
  var cache = CacheService.getScriptCache();
  var cacheKey = "siaba_trend_addon_v1"; 
  var cachedResult = cache.get(cacheKey);
  if (cachedResult) return cachedResult;

  var now = new Date();
  var curYear = now.getFullYear(); 

  // Inisialisasi Struktur Data
  function initArr() { return [0,0,0,0,0,0,0,0,0,0,0,0]; }
  var result = {
    cuti:  { proses: initArr(), revisi: initArr(), setuju: initArr(), tolak: initArr() },
    dinas: { proses: initArr(), revisi: initArr(), setuju: initArr(), tolak: initArr() },
    lupa:  { proses: initArr(), revisi: initArr(), setuju: initArr(), tolak: initArr() },
    salah: { proses: initArr(), revisi: initArr(), setuju: initArr(), tolak: initArr() }
  };

  // Helper Tanggal
  function parseTgl(v) {
    if(!v) return null;
    if(v instanceof Date) return v;
    var s = String(v).trim().replace(/'/g,'');
    var m = s.match(/^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{4})/);
    if(m) return new Date(m[3], m[2]-1, m[1]);
    return new Date(s);
  }

  // Logic Processor
  function processTrend(idSS, tabName, key, colTgl, colStat) {
    try {
      var ss = SpreadsheetApp.openById(idSS);
      var sheet = ss.getSheetByName(tabName);
      if(!sheet) return;
      var data = sheet.getDataRange().getValues();
      
      for(var i=1; i<data.length; i++) {
        var row = data[i];
        var tgl = parseTgl(row[colTgl]);
        if(!tgl || tgl.getFullYear() !== curYear) continue;

        var bln = tgl.getMonth(); // 0-11
        var st = String(row[colStat]||"").toLowerCase();
        var target = result[key];

        if(st.includes("setuju") || st.includes("ok") || st.includes("disetujui")) target.setuju[bln]++;
        else if(st.includes("tolak") || st.includes("tidak")) target.tolak[bln]++;
        else if(st.includes("revisi") || st.includes("ubah")) target.revisi[bln]++;
        else target.proses[bln]++;
      }
    } catch(e) {}
  }

  // ID DATABASE (Gunakan ID yang sama dengan di file Anda)
  var ID_DB_CUTI  = "1UYG80gGxuC19ieaVBzJaUV8bhlS2q5gExr0-Yl7upKo";
  var ID_DB_DINAS = "1I_2yUFGXnBJTCSW6oaT3D482YCs8TIRkKgQVBbvpa1M"; 
  var ID_DB_LUPA  = "160IjN8aiDAgDYXjgDLStS4nCZLKn3Ny-dq3BOFAfDrU";
  var ID_DB_SALAH = "1TZGrMiTuyvh2Xbo44RhJuWlQnOC5LzClsgIoNKtRFkY";

  // Eksekusi
  processTrend(ID_DB_CUTI, "Form Cuti", "cuti", 4, 10);
  processTrend(ID_DB_DINAS, "Perjalanan_Dinas", "dinas", 3, 9);
  processTrend(ID_DB_LUPA, "Lupa_Presensi", "lupa", 3, 10);
  processTrend(ID_DB_SALAH, "Salah_Presensi", "salah", 3, 8);

  var json = JSON.stringify(result);
  cache.put(cacheKey, json, 120);
  return json;
}

/* ======================================================================
   [ADD-ON] DATA TREN BULANAN (PARALLEL FETCHING)
   ====================================================================== */
function getTrenBulananData() {
  var cache = CacheService.getScriptCache();
  var cacheKey = "siaba_trend_full_v2"; 
  var cachedResult = cache.get(cacheKey);
  
  if (cachedResult) return cachedResult;

  var now = new Date();
  var curYear = now.getFullYear(); 

  // Init Array 0-11
  function initArr() { return [0,0,0,0,0,0,0,0,0,0,0,0]; }
  
  var result = {
    cuti:  { proses: initArr(), revisi: initArr(), setuju: initArr(), tolak: initArr() },
    dinas: { proses: initArr(), revisi: initArr(), setuju: initArr(), tolak: initArr() },
    lupa:  { proses: initArr(), revisi: initArr(), setuju: initArr(), tolak: initArr() },
    salah: { proses: initArr(), revisi: initArr(), setuju: initArr(), tolak: initArr() }
  };

  // Helper Parser Tanggal
  function parseDate(v) {
    if(!v) return null;
    if(v instanceof Date) return v;
    var s = String(v).trim().replace(/'/g,'');
    var m = s.match(/^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{4})/);
    if(m) return new Date(m[3], m[2]-1, m[1]);
    return new Date(s);
  }

  // Logic Processor
  function processData(id, tab, key, colTgl, colStat) {
    try {
      var ss = SpreadsheetApp.openById(id);
      var sh = ss.getSheetByName(tab);
      if(!sh) return;
      var data = sh.getDataRange().getValues();
      
      for(var i=1; i<data.length; i++) {
        var row = data[i];
        var tgl = parseDate(row[colTgl]);
        
        if(!tgl || tgl.getFullYear() !== curYear) continue;

        var bln = tgl.getMonth();
        var st = String(row[colStat]||"").toLowerCase();
        var target = result[key];

        if(st.includes("setuju") || st.includes("ok") || st.includes("disetujui")) target.setuju[bln]++;
        else if(st.includes("tolak") || st.includes("tidak")) target.tolak[bln]++;
        else if(st.includes("revisi") || st.includes("ubah")) target.revisi[bln]++;
        else target.proses[bln]++;
      }
    } catch(e){}
  }

  // ID DATABASE (Sesuaikan dengan file Anda)
  var ID_DB_CUTI  = "1UYG80gGxuC19ieaVBzJaUV8bhlS2q5gExr0-Yl7upKo";
  var ID_DB_DINAS = "1I_2yUFGXnBJTCSW6oaT3D482YCs8TIRkKgQVBbvpa1M"; 
  var ID_DB_LUPA  = "160IjN8aiDAgDYXjgDLStS4nCZLKn3Ny-dq3BOFAfDrU";
  var ID_DB_SALAH = "1TZGrMiTuyvh2Xbo44RhJuWlQnOC5LzClsgIoNKtRFkY";

  // Eksekusi
  processData(ID_DB_CUTI, "Form Cuti", "cuti", 4, 10);
  processData(ID_DB_DINAS, "Perjalanan_Dinas", "dinas", 3, 9);
  processData(ID_DB_LUPA, "Lupa_Presensi", "lupa", 3, 10);
  processData(ID_DB_SALAH, "Salah_Presensi", "salah", 3, 8);

  var json = JSON.stringify(result);
  cache.put(cacheKey, json, 120); // Cache 2 menit
  return json;
}