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
    
    // Gunakan getDisplayValues juga disini agar tahun/bulan terbaca sebagai teks
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
    // 1. LOOKUP
    let ssLookup;
    try { ssLookup = SpreadsheetApp.openById(ID_DB); } 
    catch(e) { return JSON.stringify({ error: "Gagal buka Database Lookup." }); }

    const sheetLookup = ssLookup.getSheetByName("Lookup Siaba");
    if (!sheetLookup) return JSON.stringify({ error: "Sheet Lookup Siaba hilang." });

    // Pakai getDisplayValues agar perbandingan string aman
    const dataLookup = sheetLookup.getDataRange().getDisplayValues();
    let targetId = "", customSheet = "";
    
    for (let i = 1; i < dataLookup.length; i++) {
        if (dataLookup[i][0] == filterTahun && dataLookup[i][1] == filterBulan) {
            targetId = dataLookup[i][2]; 
            customSheet = dataLookup[i][3];     
            break; 
        }
    }

    if (!targetId) return JSON.stringify({ error: `Data ${filterBulan} ${filterTahun} tidak ada di Lookup.` });

    // 2. TARGET FILE
    let ssTarget;
    try { ssTarget = SpreadsheetApp.openById(targetId); }
    catch(e) { return JSON.stringify({ error: `Gagal buka file ID: ...${targetId.substr(-5)}` }); }

    // 3. TARGET SHEET
    let sheetTarget = null;
    let usedName = "";

    if (customSheet) { sheetTarget = ssTarget.getSheetByName(customSheet); usedName = customSheet; }
    if (!sheetTarget) { sheetTarget = ssTarget.getSheetByName("Data Siaba"); usedName = "Data Siaba"; }
    if (!sheetTarget) { 
        const all = ssTarget.getSheets(); 
        if (all.length > 0) { sheetTarget = all[0]; usedName = sheetTarget.getName(); }
    }

    if (!sheetTarget) return JSON.stringify({ error: "File target kosong." });

    // 4. GET DATA (KUNCI PERBAIKAN DISINI)
    const maxCol = sheetTarget.getLastColumn();
    if (maxCol < 4) return JSON.stringify({ error: `Sheet '${usedName}' kolom < 4.` });

    // --- PERUBAHAN UTAMA: getDisplayValues() ---
    // Ini mengambil semua data sebagai STRING (Teks) persis tampilan di Excel
    // Tidak akan ada lagi tanggal aneh (1899, dll)
    const allData = sheetTarget.getDataRange().getDisplayValues();
    
    // Header ada di baris pertama (Index 0), mulai kolom D (Index 3)
    const headerData = allData[0].slice(3);

    allData.shift(); // Hapus baris header dari data
    
    let result = [];
    
    for (let i = 0; i < allData.length; i++) {
        let row = allData[i];
        if (row.length < 3) continue;
        
        let rowUnit = row[2]; // Kolom C
        
        if (filterUnit === "SEMUA" || rowUnit == filterUnit) {
            // Ambil data mulai kolom D (index 3)
            // Karena sudah getDisplayValues, tidak perlu konversi tanggal lagi
            let rowData = row.slice(3, 3 + headerData.length);
            result.push(rowData);
        }
    }

    // 5. SORTING
    // Meskipun data string, parseInt tetap bisa membaca angka di dalamnya untuk sorting
    if (result.length > 0) {
        result.sort((a, b) => {
            const getVal = (val) => val === "" ? 0 : (parseInt(val) || 0);
            // Index Mapping: TP(2), TA(17), PLA(19), LA(21)
            
            if (a.length < 22 || b.length < 22) return 0;

            const tpA = getVal(a[2]), tpB = getVal(b[2]);
            if (tpB !== tpA) return tpB - tpA;
            const taA = getVal(a[17]), taB = getVal(b[17]);
            if (taB !== taA) return taB - taA; 
            const plaA = getVal(a[19]), plaB = getVal(b[19]);
            if (plaB !== plaA) return plaB - plaA; 
            const laA = getVal(a[21]), laB = getVal(b[21]);
            return laB - laA; 
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
   SIABA APEL & UPACARA - LOOKUP & JSON VERSION
   Target Sheet: "Data Apel"
   ====================================================================== */

function getSiabaApelFilters() {
  // Gunakan ID Database yang sama
  const ID_DB = "1wiDKez4rL5UYnpP2-OZjYowvmt1nRx-fIMy9trJlhBA"; 
  try {
    const ss = SpreadsheetApp.openById(ID_DB);
    const sheet = ss.getSheetByName("Lookup Siaba");
    if (!sheet) return JSON.stringify({ error: "Sheet 'Lookup Siaba' tidak ditemukan." });
    
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return JSON.stringify({ years: [], months: [] });
    
    // Ambil sebagai DisplayValues
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
    // 1. LOOKUP ID
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

    if (!targetId) return JSON.stringify({ error: `Data Apel ${filterBulan} ${filterTahun} tidak ditemukan di Lookup.` });

    // 2. BUKA FILE TARGET
    let ssTarget;
    try { ssTarget = SpreadsheetApp.openById(targetId); }
    catch(e) { return JSON.stringify({ error: `Gagal akses File ID: ...${targetId.substr(-5)}` }); }

    // 3. CARI SHEET "Data Apel"
    const TARGET_SHEET_NAME = "Data Apel";
    const sheetTarget = ssTarget.getSheetByName(TARGET_SHEET_NAME);

    if (!sheetTarget) return JSON.stringify({ error: `Sheet "${TARGET_SHEET_NAME}" tidak ditemukan di file target.` });

    // 4. AMBIL DATA (TEXT MODE)
    const maxCol = sheetTarget.getLastColumn();
    if (maxCol < 4) return JSON.stringify({ error: `Sheet Data Apel kolom < 4.` });

    const allData = sheetTarget.getDataRange().getDisplayValues();
    
    // Header (Baris 1, Kolom D ke kanan)
    const headerData = allData[0].slice(3);
    
    allData.shift(); // Buang header
    
    let result = [];
    
    for (let i = 0; i < allData.length; i++) {
        let row = allData[i];
        if (row.length < 3) continue;
        
        let rowUnit = row[2]; // Kolom C
        
        if (filterUnit === "SEMUA" || rowUnit == filterUnit) {
            // Ambil data mulai kolom D (index 3)
            let rowData = row.slice(3, 3 + headerData.length);
            result.push(rowData);
        }
    }

    // Tidak ada sorting khusus yang diminta untuk Apel, tampilkan apa adanya
    
    return JSON.stringify({
      headers: headerData,
      rows: result
    });

  } catch (e) {
    return JSON.stringify({ error: "SYSTEM ERROR: " + e.message });
  }
}

/* ======================================================================
   SIABA TIDAK PRESENSI - LOOKUP & JSON VERSION
   Target Sheet: "Data Alpa"
   Sorting: Kolom F (Index 2 dalam slice) Descending
   ====================================================================== */

function getSiabaTidakFilters() {
  // Gunakan ID Database yang sama
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
    // 1. LOOKUP ID
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

    // 3. CARI SHEET "Data Alpa"
    const TARGET_SHEET_NAME = "Data Alpa";
    const sheetTarget = ssTarget.getSheetByName(TARGET_SHEET_NAME);

    if (!sheetTarget) return JSON.stringify({ error: `Sheet "${TARGET_SHEET_NAME}" tidak ditemukan di file target.` });

    // 4. AMBIL DATA (TEXT MODE)
    const maxCol = sheetTarget.getLastColumn();
    if (maxCol < 4) return JSON.stringify({ error: `Sheet Data Alpa kolom < 4.` });

    const allData = sheetTarget.getDataRange().getDisplayValues();
    
    // Header (Baris 1, Kolom D ke kanan)
    const headerData = allData[0].slice(3);
    
    allData.shift(); // Buang header
    
    let result = [];
    
    for (let i = 0; i < allData.length; i++) {
        let row = allData[i];
        if (row.length < 3) continue;
        
        let rowUnit = row[2]; // Kolom C
        
        // Handle filter Default "SEMUA" (berjaga-jaga jika frontend kirim kosong)
        if (!filterUnit) filterUnit = "SEMUA";

        if (filterUnit === "SEMUA" || rowUnit == filterUnit) {
            // Ambil data mulai kolom D (index 3)
            let rowData = row.slice(3, 3 + headerData.length);
            result.push(rowData);
        }
    }

    // 5. SORTING DESCENDING BERDASARKAN KOLOM F
    // Kolom D=0, E=1, F=2 (Index 2 dalam rowData)
    if (result.length > 0) {
        result.sort((a, b) => {
            const valA = parseInt(a[2]) || 0;
            const valB = parseInt(b[2]) || 0;
            return valB - valA; // Descending (Terbanyak di atas)
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
   SIABA TERLAMBAT - NESTED HEADERS (NO MONTH FILTER)
   Sheet: "Rekap_Terlambat"
   Filter: Tahun (Col A) & Unit (Frontend/Col D)
   ====================================================================== */

function getSiabaTerlambatFilters() {
  const ID_DB = "1tQsQY1-Ny1ie66GOZPTLtvZ7BiYCgFdNrX-AVGCtaHA"; 
  try {
    const ss = SpreadsheetApp.openById(ID_DB);
    const sheet = ss.getSheetByName("Rekap_Terlambat");
    if (!sheet) return JSON.stringify({ error: "Sheet 'Rekap_Terlambat' tidak ditemukan." });
    
    const lastRow = sheet.getLastRow();
    if (lastRow < 3) return JSON.stringify({ years: [] });
    
    // Ambil Kolom A (Tahun) saja
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

/* ======================================================================
   SIABA TERLAMBAT - FINAL FIX
   ====================================================================== */

function getSiabaTerlambatData(filterTahun, filterUnit) {
  const ID_DB = "1tQsQY1-Ny1ie66GOZPTLtvZ7BiYCgFdNrX-AVGCtaHA";
  
  try {
    const ss = SpreadsheetApp.openById(ID_DB);
    const sheet = ss.getSheetByName("Rekap_Terlambat");
    if (!sheet) return JSON.stringify({ error: "Sheet 'Rekap_Terlambat' tidak ditemukan." });

    const maxCol = sheet.getLastColumn(); 

    // 1. HEADER (Baris 1 & 2)
    const headerRange = sheet.getRange(1, 3, 2, maxCol - 2).getDisplayValues();
    const headerTop = headerRange[0]; 
    const headerSub = headerRange[1]; 

    // 2. DATA (Mulai Baris 3)
    const lastRow = sheet.getLastRow();
    if (lastRow < 3) return JSON.stringify({ error: "Data Kosong" });

    const rawData = sheet.getRange(3, 1, lastRow - 2, maxCol).getDisplayValues();
    
    let result = [];
    
    // Normalisasi Filter Unit
    // Jika kosong/null, anggap SEMUA (atau frontend yang handle, tapi di sini kita jaga-jaga)
    const targetUnit = filterUnit ? String(filterUnit).toUpperCase().trim() : "SEMUA";

    for (let i = 0; i < rawData.length; i++) {
        let row = rawData[i];
        
        let rowTahun = String(row[0]).trim(); // Kolom A
        let rowUnit  = String(row[1]).toUpperCase().trim(); // Kolom B (Unit Kerja)
        
        // --- LOGIKA FILTER ---
        if (rowTahun == String(filterTahun).trim()) {
             
             // Jika Filter Unit TIDAK "SEMUA" dan TIDAK KOSONG, maka harus cocok
             if (targetUnit !== "SEMUA" && targetUnit !== "" && rowUnit !== targetUnit) {
                 continue; 
             }

             let rowDisplay = row.slice(2); 
             result.push(rowDisplay);
        }
    }

    // 3. SORTING DESCENDING
    if (result.length > 0) {
        result.sort((a, b) => {
            let idxLast = a.length - 1; 
            let valA = parseInt(a[idxLast].replace(/\./g,'')) || 0;
            let valB = parseInt(b[idxLast].replace(/\./g,'')) || 0;
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
   SIABA PULANG AWAL - LOGIC COPY FROM TERLAMBAT
   Sheet: "Rekap_Pulang_Awal"
   ====================================================================== */

function getSiabaPulangFilters() {
  const ID_DB = "1tQsQY1-Ny1ie66GOZPTLtvZ7BiYCgFdNrX-AVGCtaHA"; 
  try {
    const ss = SpreadsheetApp.openById(ID_DB);
    const sheet = ss.getSheetByName("Rekap_Pulang_Awal"); // <-- BEDA SHEET
    if (!sheet) return JSON.stringify({ error: "Sheet 'Rekap_Pulang_Awal' tidak ditemukan." });
    
    const lastRow = sheet.getLastRow();
    if (lastRow < 3) return JSON.stringify({ years: [] });
    
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
    const sheet = ss.getSheetByName("Rekap_Pulang_Awal"); // <-- BEDA SHEET
    if (!sheet) return JSON.stringify({ error: "Sheet 'Rekap_Pulang_Awal' tidak ditemukan." });

    const maxCol = sheet.getLastColumn(); 

    // 1. HEADER (Baris 1 & 2)
    const headerRange = sheet.getRange(1, 3, 2, maxCol - 2).getDisplayValues();
    const headerTop = headerRange[0]; 
    const headerSub = headerRange[1]; 

    // 2. DATA (Mulai Baris 3)
    const lastRow = sheet.getLastRow();
    if (lastRow < 3) return JSON.stringify({ error: "Data Kosong" });

    const rawData = sheet.getRange(3, 1, lastRow - 2, maxCol).getDisplayValues();
    
    let result = [];
    
    // Normalisasi Filter Unit
    const targetUnit = filterUnit ? String(filterUnit).toUpperCase().trim() : "SEMUA";

    for (let i = 0; i < rawData.length; i++) {
        let row = rawData[i];
        let rowTahun = String(row[0]).trim(); 
        let rowUnit  = String(row[1]).toUpperCase().trim(); // Kolom B
        
        // Filter
        if (rowTahun == String(filterTahun).trim()) {
             if (targetUnit !== "SEMUA" && targetUnit !== "" && rowUnit !== targetUnit) {
                 continue; 
             }
             let rowDisplay = row.slice(2); 
             result.push(rowDisplay);
        }
    }

    // 3. SORTING DESCENDING
    if (result.length > 0) {
        result.sort((a, b) => {
            let idxLast = a.length - 1; 
            let valA = parseInt(a[idxLast].replace(/\./g,'')) || 0;
            let valB = parseInt(b[idxLast].replace(/\./g,'')) || 0;
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
   SIABA UNDUH REKAP - SORTING BULAN & UNDUH
   Root Folder: 1MoGuseJNrOIMnkZNoqkKcK282jZpUkAm
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

  // Cek Parent
  let parentId = null;
  let isRoot = (targetId === ROOT_ID);
  
  if (!isRoot) {
    let parents = folder.getParents();
    if (parents.hasNext()) parentId = parents.next().getId();
  }

  let items = [];

  // 1. AMBIL SUB-FOLDER
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

  // 2. AMBIL FILE
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

  // 3. LOGIKA SORTING (FOLDER BULAN KRONOLOGIS)
  const URUTAN_BULAN = [
      "JANUARI", "FEBRUARI", "MARET", "APRIL", "MEI", "JUNI", 
      "JULI", "AGUSTUS", "SEPTEMBER", "OKTOBER", "NOVEMBER", "DESEMBER"
  ];

  items.sort((a, b) => {
     // A. Prioritas: Folder selalu di atas File
     if (a.type !== b.type) {
         return a.type === 'folder' ? -1 : 1;
     }

     // B. Sorting Nama
     let nameA = a.name.toUpperCase().trim();
     let nameB = b.name.toUpperCase().trim();

     // Cek apakah keduanya adalah Nama Bulan?
     let idxA = URUTAN_BULAN.indexOf(nameA);
     let idxB = URUTAN_BULAN.indexOf(nameB);

     // Jika keduanya adalah Bulan, urutkan berdasarkan Index Bulan (Kronologis)
     if (idxA > -1 && idxB > -1) {
         return idxA - idxB;
     }

     // Jika tidak (atau salah satu bukan bulan), urutkan Abjad biasa
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