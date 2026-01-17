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
    // 1. LOOKUP
    let ssLookup;
    try { ssLookup = SpreadsheetApp.openById(ID_DB); } 
    catch(e) { return JSON.stringify({ error: "Gagal buka Database Lookup." }); }

    const sheetLookup = ssLookup.getSheetByName("Lookup Siaba");
    if (!sheetLookup) return JSON.stringify({ error: "Sheet Lookup Siaba hilang." });

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

    // 4. GET DATA
    const maxCol = sheetTarget.getLastColumn();
    if (maxCol < 4) return JSON.stringify({ error: `Sheet '${usedName}' kolom < 4.` });

    const allData = sheetTarget.getDataRange().getDisplayValues();
    const headerData = allData[0].slice(3);

    allData.shift(); 
    
    let result = [];
    
    for (let i = 0; i < allData.length; i++) {
        let row = allData[i];
        if (row.length < 3) continue;
        
        let rowUnit = row[2]; 
        
        if (filterUnit === "SEMUA" || rowUnit == filterUnit) {
            let rowData = row.slice(3, 3 + headerData.length);
            result.push(rowData);
        }
    }

    // 5. SORTING
    if (result.length > 0) {
        result.sort((a, b) => {
            const getVal = (val) => val === "" ? 0 : (parseInt(val) || 0);
            
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

    let ssTarget;
    try { ssTarget = SpreadsheetApp.openById(targetId); }
    catch(e) { return JSON.stringify({ error: `Gagal akses File ID: ...${targetId.substr(-5)}` }); }

    const TARGET_SHEET_NAME = "Data Apel";
    const sheetTarget = ssTarget.getSheetByName(TARGET_SHEET_NAME);

    if (!sheetTarget) return JSON.stringify({ error: `Sheet "${TARGET_SHEET_NAME}" tidak ditemukan di file target.` });

    const maxCol = sheetTarget.getLastColumn();
    if (maxCol < 4) return JSON.stringify({ error: `Sheet Data Apel kolom < 4.` });

    const allData = sheetTarget.getDataRange().getDisplayValues();
    const headerData = allData[0].slice(3);
    
    allData.shift(); 
    
    let result = [];
    
    for (let i = 0; i < allData.length; i++) {
        let row = allData[i];
        if (row.length < 3) continue;
        
        let rowUnit = row[2]; 
        
        if (filterUnit === "SEMUA" || rowUnit == filterUnit) {
            let rowData = row.slice(3, 3 + headerData.length);
            result.push(rowData);
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
   SIABA TIDAK PRESENSI
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

    let ssTarget;
    try { ssTarget = SpreadsheetApp.openById(targetId); }
    catch(e) { return JSON.stringify({ error: `Gagal akses File ID: ...${targetId.substr(-5)}` }); }

    const TARGET_SHEET_NAME = "Data Alpa";
    const sheetTarget = ssTarget.getSheetByName(TARGET_SHEET_NAME);

    if (!sheetTarget) return JSON.stringify({ error: `Sheet "${TARGET_SHEET_NAME}" tidak ditemukan di file target.` });

    const maxCol = sheetTarget.getLastColumn();
    if (maxCol < 4) return JSON.stringify({ error: `Sheet Data Alpa kolom < 4.` });

    const allData = sheetTarget.getDataRange().getDisplayValues();
    const headerData = allData[0].slice(3);
    
    allData.shift(); 
    
    let result = [];
    
    for (let i = 0; i < allData.length; i++) {
        let row = allData[i];
        if (row.length < 3) continue;
        
        let rowUnit = row[2]; 
        
        if (!filterUnit) filterUnit = "SEMUA";

        if (filterUnit === "SEMUA" || rowUnit == filterUnit) {
            let rowData = row.slice(3, 3 + headerData.length);
            result.push(rowData);
        }
    }

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
    if (!sheet) return JSON.stringify({ error: "Sheet 'Rekap_Terlambat' tidak ditemukan." });
    
    const lastRow = sheet.getLastRow();
    if (lastRow < 3) return JSON.stringify({ years: [] });
    
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

    const headerRange = sheet.getRange(1, 3, 2, maxCol - 2).getDisplayValues();
    const headerTop = headerRange[0]; 
    const headerSub = headerRange[1]; 

    const lastRow = sheet.getLastRow();
    if (lastRow < 3) return JSON.stringify({ error: "Data Kosong" });

    const rawData = sheet.getRange(3, 1, lastRow - 2, maxCol).getDisplayValues();
    
    let result = [];
    const targetUnit = filterUnit ? String(filterUnit).toUpperCase().trim() : "SEMUA";

    for (let i = 0; i < rawData.length; i++) {
        let row = rawData[i];
        
        let rowTahun = String(row[0]).trim(); 
        let rowUnit  = String(row[1]).toUpperCase().trim(); 
        
        if (rowTahun == String(filterTahun).trim()) {
             if (targetUnit !== "SEMUA" && targetUnit !== "" && rowUnit !== targetUnit) {
                 continue; 
             }
             let rowDisplay = row.slice(2); 
             result.push(rowDisplay);
        }
    }

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
   SIABA PULANG AWAL
   ====================================================================== */

function getSiabaPulangFilters() {
  const ID_DB = "1tQsQY1-Ny1ie66GOZPTLtvZ7BiYCgFdNrX-AVGCtaHA"; 
  try {
    const ss = SpreadsheetApp.openById(ID_DB);
    const sheet = ss.getSheetByName("Rekap_Pulang_Awal"); 
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
    const sheet = ss.getSheetByName("Rekap_Pulang_Awal"); 
    if (!sheet) return JSON.stringify({ error: "Sheet 'Rekap_Pulang_Awal' tidak ditemukan." });

    const maxCol = sheet.getLastColumn(); 

    const headerRange = sheet.getRange(1, 3, 2, maxCol - 2).getDisplayValues();
    const headerTop = headerRange[0]; 
    const headerSub = headerRange[1]; 

    const lastRow = sheet.getLastRow();
    if (lastRow < 3) return JSON.stringify({ error: "Data Kosong" });

    const rawData = sheet.getRange(3, 1, lastRow - 2, maxCol).getDisplayValues();
    
    let result = [];
    const targetUnit = filterUnit ? String(filterUnit).toUpperCase().trim() : "SEMUA";

    for (let i = 0; i < rawData.length; i++) {
        let row = rawData[i];
        let rowTahun = String(row[0]).trim(); 
        let rowUnit  = String(row[1]).toUpperCase().trim(); 
        
        if (rowTahun == String(filterTahun).trim()) {
             if (targetUnit !== "SEMUA" && targetUnit !== "" && rowUnit !== targetUnit) {
                 continue; 
             }
             let rowDisplay = row.slice(2); 
             result.push(rowDisplay);
        }
    }

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

/* --- FUNGSI GET DATA SALAH PRESENSI (VERSI FILTER SERVER-SIDE) --- */
function getDaftarSalahPresensi(tahun, bulan, unit, status) {
  var ID_DB = "1TZGrMiTuyvh2Xbo44RhJuWlQnOC5LzClsgIoNKtRFkY";
  var SHEET_NAME = "Salah_Absen";

  try {
    var ss = SpreadsheetApp.openById(ID_DB);
    var sheet = ss.getSheetByName(SHEET_NAME);
    if (!sheet) return JSON.stringify([]);

    // Ambil Data
    var data = sheet.getDataRange().getDisplayValues();
    var output = [];

    // Sanitasi Filter
    var fTahun  = (tahun && String(tahun).trim() !== "") ? String(tahun).trim() : null;
    var fBulan  = (bulan && String(bulan).trim() !== "") ? String(bulan).trim() : null;
    var fUnit   = (unit && String(unit).trim() !== "" && unit !== "SEMUA") ? String(unit).trim() : null;
    var fStatus = (status && String(status).trim() !== "" && status !== "SEMUA") ? String(status).trim() : null;
    var arrBulan = ["Januari", "Februari", "Maret", "April", "Mei", "Juni", "Juli", "Agustus", "September", "Oktober", "November", "Desember"];

    // Loop (Mulai baris 2, index 1)
    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      if (!row[1]) continue; // Skip jika Nama kosong

      // --- LOGIKA FILTER ---
      
      // 1. Filter Tanggal (Kolom Index 3: dd-mm-yyyy atau dd/mm/yyyy)
      if (fTahun || fBulan) {
          var tglStr = String(row[3]).replace(/\//g, '-'); // Ubah / jadi -
          var tglParts = tglStr.split('-');
          if (tglParts.length === 3) {
             // Asumsi format: 12-05-2026 (dd-mm-yyyy) -> part[2] = tahun
             // Jika format sheet yyyy-mm-dd, sesuaikan indexnya.
             // Default Apps Script biasanya dd-mm-yyyy jika locale Indonesia.
             var thnData = (tglParts[0].length === 4) ? tglParts[0] : tglParts[2]; // Cek mana yg tahun
             var blnData = parseInt((tglParts[0].length === 4) ? tglParts[1] : tglParts[1]);

             if (fTahun && thnData !== fTahun) continue;
             if (fBulan) {
                 if (!isNaN(blnData) && blnData >= 1 && blnData <= 12) {
                     if (arrBulan[blnData-1] !== fBulan) continue;
                 } else { continue; }
             }
          }
      }

      // 2. Filter Unit (Kolom Index 0)
      if (fUnit && String(row[0]) !== fUnit) continue;

      // 3. Filter Status (Kolom Index 8)
      if (fStatus && String(row[8]) !== fStatus) continue;

      // --- MAPPING DATA ---
      var rowData = {
        rowBaris: i + 1,     
        unitKerja: row[0],   
        namaAsn: row[1],     
        nip: row[2],         
        tanggal: row[3],     
        jam: row[4],         
        jenis: row[5],       
        tglAjuan: row[6],    
        userInput: row[7],   
        status: row[8],      
        ket: row[9],         
        tglEdit: row[10],    
        userEdit: row[11],   
        tglVerif: row[12],   
        adminVerif: row[13]  
      };
      output.push(rowData);
    }
    
    return JSON.stringify(output);

  } catch (e) {
    return JSON.stringify([{ status: 'Error', nama: e.toString() }]);
  }
}

/* =================================================================
   FUNGSI DATABASE PEGAWAI (SUMBER PUSAT - SINGLE SOURCE OF TRUTH)
   Spreadsheet Pusat: 1ReJt2qoDE2f_8LeR8DXJbROB9EAHK8qP2kYp-ZZ3V9w
   Sheet: Database_ASN_SIKS
   ================================================================= */

function getDatabasePegawai() {
  // ID Spreadsheet PUSAT
  const ID_DB_PUSAT = "1ReJt2qoDE2f_8LeR8DXJbROB9EAHK8qP2kYp-ZZ3V9w";
  
  try {
    var ss = SpreadsheetApp.openById(ID_DB_PUSAT);
    var sheet = ss.getSheetByName("Database_ASN_SIKS"); 
    
    if (!sheet) throw new Error("Sheet 'Database_ASN_SIKS' tidak ditemukan di file Pusat!");

    var lastRow = sheet.getLastRow();
    if (lastRow < 2) return []; 

    // Ambil Kolom A, B, C (Unit, NIP, Nama)
    // Asumsi di File Pusat:
    // Col A (Indeks 0) = Unit Kerja
    // Col B (Indeks 1) = NIP
    // Col C (Indeks 2) = Nama
    
    var data = sheet.getRange(2, 1, lastRow - 1, 3).getValues();
    
    // Mapping Data
    return data.map(function(r) {
      return {
        unit: r[0], // Kolom A
        nip:  r[1], // Kolom B
        nama: r[2]  // Kolom C
      };
    });

  } catch (e) {
    // Fallback: Jika gagal akses pusat, throw error agar ketahuan
    throw new Error("Gagal ambil Database Pusat: " + e.message);
  }
}

/* --- FUNGSI SIMPAN DATA BARU (FORMAT TEXT MANUAL) --- */
function simpanSalahAbsen(form) {
  const ID_DB = "1TZGrMiTuyvh2Xbo44RhJuWlQnOC5LzClsgIoNKtRFkY";

  try {
    var ss = SpreadsheetApp.openById(ID_DB);
    var sheet = ss.getSheetByName("Salah_Absen");
    
    if (!sheet) throw new Error("Sheet 'Salah_Absen' tidak ditemukan!");
    
    // 1. FORMAT TANGGAL (dd-mm-yyyy)
    // Kita ubah input yyyy-mm-dd menjadi dd-mm-yyyy secara manual
    // agar tersimpan sebagai string yang konsisten.
    var tglSimpan = "";
    if (form.tanggal) {
       var parts = form.tanggal.split('-'); // input: 2026-01-11
       tglSimpan = parts[2] + '-' + parts[1] + '-' + parts[0]; // hasil: 11-01-2026
    }
    
    // 2. FORMAT JAM (HH:mm)
    // Kita pastikan tersimpan sebagai string "14:00"
    // Tanpa tanda kutip, tapi format terjaga.
    var jamSimpan = String(form.waktu); 

    // 3. GET USER
    var namaUser = "Guest";
    try {
       var currentUser = getCurrentUser();
       if (currentUser && currentUser.fullName) {
          namaUser = currentUser.fullName;
       }
    } catch (err) {
       namaUser = "Guest (Error User)";
    }
    
    // 4. HISTORY (Lengkap dengan Detik)
    // Untuk history, kita pakai format lengkap dd-mm-yyyy HH:mm:ss
    var tglKirim = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd-MM-yyyy HH:mm:ss");
    var status = "Diproses";
    
    var barisBaru = [
      form.unit_kerja, 
      form.nama_asn,   
      "'"+form.nip_asn, // NIP wajib kutip agar 0 depan aman
      tglSimpan,        // Kolom D: String "11-01-2026"
      jamSimpan,        // Kolom E: String "14:00"
      form.jenis,      
      tglKirim,         // Kolom G: String "11-01-2026 14:05:00"
      namaUser,        
      status           
    ];

    sheet.appendRow(barisBaru);
    return "SUKSES";
    
  } catch (e) {
    throw new Error("Gagal simpan: " + e.message);
  }
}

/* --- FUNGSI UPDATE DATA (VERSI ROBUST / TAHAN BANTING) --- */
function updateSalahAbsen(form) {
  const ID_DB = "1TZGrMiTuyvh2Xbo44RhJuWlQnOC5LzClsgIoNKtRFkY";
  try {
    var ss = SpreadsheetApp.openById(ID_DB);
    var sheet = ss.getSheetByName("Salah_Absen");
    if (!sheet) throw new Error("Sheet Salah_Absen tidak ditemukan");

    // 1. DATA TARGET (KEY PENCARIAN DARI CLIENT)
    var targetNip = String(form.nip_lama).trim();
    var targetTgl = String(form.tgl_lama).trim(); // dd-mm-yyyy
    var targetJam = String(form.jam_lama).trim(); // HH:mm

    // 2. CARI BARIS (GUNAKAN DISPLAY VALUES AGAR AKURAT)
    var data = sheet.getDataRange().getDisplayValues();
    var barisKetemu = -1;
    var statusSaatIni = "";

    for (var i = 1; i < data.length; i++) {
       var sheetNip = String(data[i][2]).trim();
       
       // Normalisasi Tanggal Sheet (jaga-jaga format miring /)
       var sheetTgl = String(data[i][3]).trim().replace(/\//g, '-');
       
       // Normalisasi Jam Sheet (Handle 7:15 -> 07:15 dan buang kutip/detik)
       var sheetJam = String(data[i][4]).trim().replace(/'/g, "").substring(0, 5);
       if (/^\d:\d{2}/.test(sheetJam)) sheetJam = "0" + sheetJam;

       if (sheetNip === targetNip && sheetTgl === targetTgl && sheetJam === targetJam) {
          barisKetemu = i + 1; // Index array + 1 = Nomor Baris Excel
          statusSaatIni = String(data[i][8]).trim(); // Ambil Status (Kolom I / Index 8)
          break;
       }
    }

    if (barisKetemu === -1) {
      throw new Error(`Data tidak ditemukan. Format Jam/Tanggal mungkin berbeda.`);
    }

    // 3. LOGIKA STATUS OTOMATIS
    var statusBaru = statusSaatIni;
    if (statusSaatIni === "Ditolak") statusBaru = "Revisi";
    else if (statusSaatIni === "Revisi") statusBaru = "Diproses";
    else if (statusSaatIni === "Diproses") statusBaru = "Diproses";

    // 4. PERSIAPAN DATA BARU (SANITASI FORMAT)
    // Pastikan Tanggal Baru tersimpan sebagai dd-mm-yyyy
    var tglBaruIndo = "";
    if (form.tanggal) {
       var parts = form.tanggal.split('-'); // input: yyyy-mm-dd
       tglBaruIndo = parts[2] + '-' + parts[1] + '-' + parts[0]; 
    }
    
    // Pastikan Jam Baru tersimpan sebagai HH:mm (07:15)
    var jamBaru = String(form.waktu).trim();
    if (/^\d:\d{2}/.test(jamBaru)) jamBaru = "0" + jamBaru;

    // 5. GET USER EDIT
    var userEdit = "Guest";
    try {
       var currentUser = getCurrentUser(); 
       if (currentUser && currentUser.fullName) userEdit = currentUser.fullName;
    } catch (e) {}
    var tglEdit = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd-MM-yyyy HH:mm:ss");

    // 6. EKSEKUSI UPDATE KE SEL
    // Gunakan setValue string agar format terjaga
    sheet.getRange(barisKetemu, 4).setValue(tglBaruIndo); // Kolom D
    sheet.getRange(barisKetemu, 5).setValue(jamBaru);     // Kolom E
    sheet.getRange(barisKetemu, 6).setValue(form.jenis);  // Kolom F
    sheet.getRange(barisKetemu, 9).setValue(statusBaru);  // Kolom I
    sheet.getRange(barisKetemu, 11).setValue(tglEdit);    // Kolom K
    sheet.getRange(barisKetemu, 12).setValue(userEdit);   // Kolom L

    return "Data Berhasil Diperbarui!";
  } catch (e) {
    throw new Error(e.message);
  }
}

/* --- FUNGSI SOFT DELETE (FIX JAM 0 DIGIT) --- */
function softDeleteSalahAbsen(dataKirim) {
  const ID_DB = "1TZGrMiTuyvh2Xbo44RhJuWlQnOC5LzClsgIoNKtRFkY";

  try {
    var ss = SpreadsheetApp.openById(ID_DB);
    var sheetSource = ss.getSheetByName("Salah_Absen");
    var sheetTrash = ss.getSheetByName("Trash");

    if (!sheetTrash) sheetTrash = ss.insertSheet("Trash");
    if (!sheetSource) throw new Error("Sheet Salah_Absen tidak ditemukan");

    // 1. DATA TARGET (String Bersih dari Client)
    var targetNip = String(dataKirim.nip).trim();
    var targetTgl = String(dataKirim.tgl).trim(); // "dd-mm-yyyy"
    var targetJam = String(dataKirim.jam).trim(); // "07:15" (Pasti 2 digit depan)

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

    // 3. AMBIL DATA SUMBER
    var rowRange = sheetSource.getRange(barisKetemu, 1, 1, 14);
    var rowValues = rowRange.getValues()[0]; 

    // --- BAGIAN PENTING: SANITASI SEBELUM MASUK TRASH ---
    // Jangan pakai nilai rowValues[4] mentah, karena bisa jadi isinya Date Object atau 7:15.
    // Kita TIMPA dengan targetJam (07:15) yang sudah pasti benar formatnya.
    
    rowValues[3] = targetTgl; // Paksa format dd-mm-yyyy
    rowValues[4] = targetJam; // Paksa format HH:mm (07:15)
    
    // ----------------------------------------------------

    var tglHapus = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd-MM-yyyy HH:mm:ss");
    var userHapus = "Guest";
    try {
        var cu = getCurrentUser(); 
        if (cu && cu.fullName) userHapus = cu.fullName;
        else if (dataKirim.user) userHapus = dataKirim.user;
    } catch(e) { userHapus = dataKirim.user || "Guest"; }

    var alasan = dataKirim.alasan;
    var trashRow = rowValues.concat([tglHapus, userHapus, alasan]);

    // Simpan ke Trash
    sheetTrash.appendRow(trashRow); 
    // Hapus dari Sumber
    sheetSource.deleteRow(barisKetemu);

    return "Sukses";

  } catch (e) {
    throw new Error(e.message);
  }
}

/* --- FUNGSI VERIFIKASI (VERSI ROBUST / TAHAN BANTING) --- */
function processVerifikasiSalahAbsen(dataKirim) {
  const ID_DB = "1TZGrMiTuyvh2Xbo44RhJuWlQnOC5LzClsgIoNKtRFkY";
  try {
    var ss = SpreadsheetApp.openById(ID_DB);
    var sheet = ss.getSheetByName("Salah_Absen");
    if (!sheet) throw new Error("Sheet Salah_Absen tidak ditemukan");

    // 1. DATA TARGET
    var targetNip = String(dataKirim.nip).trim();
    var targetTgl = String(dataKirim.tgl).trim();
    var targetJam = String(dataKirim.jam).trim();

    // 2. CARI BARIS (GUNAKAN DISPLAY VALUES)
    var data = sheet.getDataRange().getDisplayValues();
    var barisKetemu = -1;

    for (var i = 1; i < data.length; i++) {
       var sheetNip = String(data[i][2]).trim();
       
       // Normalisasi Tgl
       var sheetTgl = String(data[i][3]).trim().replace(/\//g, '-');
       
       // Normalisasi Jam (KUNCI UTAMA: Handle 7:15 -> 07:15)
       var sheetJam = String(data[i][4]).trim().replace(/'/g, "").substring(0, 5);
       if (/^\d:\d{2}/.test(sheetJam)) sheetJam = "0" + sheetJam;

       if (sheetNip === targetNip && sheetTgl === targetTgl && sheetJam === targetJam) {
          barisKetemu = i + 1;
          break;
       }
    }

    if (barisKetemu === -1) {
      throw new Error("Data tidak ditemukan saat verifikasi. Coba refresh tabel.");
    }

    // 3. UPDATE DATA VERIFIKASI
    var tglVerif = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd-MM-yyyy HH:mm:ss");
    
    sheet.getRange(barisKetemu, 9).setValue(dataKirim.status);  // Status (Kolom I)
    sheet.getRange(barisKetemu, 10).setValue(dataKirim.ket);    // Ket (Kolom J)
    sheet.getRange(barisKetemu, 13).setValue(tglVerif);         // Tgl Verif (Kolom M)
    sheet.getRange(barisKetemu, 14).setValue(dataKirim.admin);  // Admin (Kolom N)

    return "Sukses";
  } catch (e) {
    throw new Error(e.message);
  }
}

/* --- FITUR SAMPAH & RESTORE --- */

function getDaftarSampah() {
  try {
    var ss = SpreadsheetApp.openById("1TZGrMiTuyvh2Xbo44RhJuWlQnOC5LzClsgIoNKtRFkY");
    var sheet = ss.getSheetByName("Trash");
    if (!sheet || sheet.getLastRow() < 2) return [];

    // Ambil semua data trash (tanpa header)
    // Urutkan dari yg terakhir dihapus (Logika array reverse di JS client atau di sini)
    var data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 17).getDisplayValues();
    return data.reverse(); // Yang baru dihapus ada di atas
  } catch (e) {
    return [];
  }
}

/* --- FUNGSI RESTORE (AUTO REPAIR JAM 7:15 -> 07:15) --- */
function prosesRestoreSalahAbsen(nip, tgl, jam) {
  const ID_DB = "1TZGrMiTuyvh2Xbo44RhJuWlQnOC5LzClsgIoNKtRFkY";
  try {
    var ss = SpreadsheetApp.openById(ID_DB);
    var sheetTrash = ss.getSheetByName("Trash");
    var sheetSource = ss.getSheetByName("Salah_Absen");
    if (!sheetTrash || !sheetSource) throw new Error("Database error.");

    // Pakai DisplayValues agar pencarian mudah
    var dataDisplay = sheetTrash.getDataRange().getDisplayValues();
    var barisKetemu = -1;

    // 1. NORMALISASI TARGET
    var targetNip = String(nip).trim();
    var targetTgl = String(tgl).trim();
    var targetJam = String(jam).trim(); // "07:15"

    // 2. CARI DI TRASH
    for (var i = 1; i < dataDisplay.length; i++) {
       var sheetNip = String(dataDisplay[i][2]).trim();
       var sheetTgl = String(dataDisplay[i][3]).trim().replace(/\//g, '-');
       
       // Normalisasi Jam Trash (Jaga-jaga di trash tersimpan 7:15)
       var sheetJam = String(dataDisplay[i][4]).trim().replace(/'/g, "").substring(0, 5);
       if (/^\d:\d{2}/.test(sheetJam)) sheetJam = "0" + sheetJam;

       if (sheetNip === targetNip && sheetTgl === targetTgl && sheetJam === targetJam) {
          barisKetemu = i + 1;
          break;
       }
    }

    if (barisKetemu === -1) throw new Error(`Data Trash tidak ditemukan.`);

    // 3. AMBIL DATA ASLI (GetValues untuk ambil object aslinya)
    var fullRow = sheetTrash.getRange(barisKetemu, 1, 1, 14).getValues()[0];
    
    // --- SANITASI / REPAIR DATA SEBELUM KEMBALI KE UTAMA ---
    
    // Repair Tanggal -> Paksa string "dd-mm-yyyy"
    if (fullRow[3] instanceof Date) {
        fullRow[3] = Utilities.formatDate(fullRow[3], Session.getScriptTimeZone(), "dd-MM-yyyy");
    } else {
        fullRow[3] = String(fullRow[3]).trim().replace(/\//g, '-');
    }

    // Repair Jam -> Paksa string "HH:mm" (Menangani kasus 7:15 -> 07:15)
    var rawJam = fullRow[4];
    if (rawJam instanceof Date) {
        fullRow[4] = Utilities.formatDate(rawJam, Session.getScriptTimeZone(), "HH:mm");
    } else {
        var strJam = String(rawJam).replace(/'/g, "").trim().substring(0, 5);
        // INI KUNCINYA: Jika depannya cuma 1 digit angka lalu titik dua, tambah 0
        if (/^\d:\d{2}/.test(strJam)) {
            strJam = "0" + strJam;
        }
        fullRow[4] = strJam;
    }
    // -------------------------------------------------------

    sheetSource.appendRow(fullRow);
    sheetTrash.deleteRow(barisKetemu);

    return "Sukses";
  } catch (e) {
    throw new Error(e.message);
  }
}

/* ======================================================================
   MODUL LUPA PRESENSI (VERSI LOKAL - LEBIH STABIL)
   Semua ID didefinisikan di dalam fungsi agar tidak ada konflik global.
   ====================================================================== */

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

/* --- FUNGSI PENCARIAN (FIX: FILTER UNIT & STATUS DIAKTIFKAN) --- */
function getDaftarLupaPresensi(tahun, bulan, unit, status) {
  var ID_DB = "160IjN8aiDAgDYXjgDLStS4nCZLKn3Ny-dq3BOFAfDrU"; 
  var SHEET_NAME = "Lupa_Presensi";

  try {
    var ss = SpreadsheetApp.openById(ID_DB);
    var sheet = ss.getSheetByName(SHEET_NAME);
    if (!sheet) throw new Error("Sheet tidak ditemukan.");

    var data = sheet.getDataRange().getDisplayValues(); 
    var result = [];

    // Sanitasi Filter
    var fTahun  = (tahun && String(tahun).trim() !== "") ? String(tahun).trim() : null;
    var fBulan  = (bulan && String(bulan).trim() !== "") ? String(bulan).trim() : null;
    var fUnit   = (unit && String(unit).trim() !== "" && unit !== "SEMUA") ? String(unit).trim() : null;
    var fStatus = (status && String(status).trim() !== "" && status !== "SEMUA") ? String(status).trim() : null;

    var arrBulan = ["Januari", "Februari", "Maret", "April", "Mei", "Juni", "Juli", "Agustus", "September", "Oktober", "November", "Desember"];

    // Loop data
    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      if (!row[1]) continue; 

      // 1. FILTER TAHUN & BULAN
      if (fTahun || fBulan) {
          var tglParts = String(row[3]).split('-'); 
          if (tglParts.length === 3) {
             var thnData = tglParts[2].trim();
             var blnData = parseInt(tglParts[1].trim());

             if (fTahun && thnData !== fTahun) continue;
             if (fBulan) {
                 if (!isNaN(blnData) && blnData >= 1 && blnData <= 12) {
                     if (arrBulan[blnData-1] !== fBulan) continue;
                 } else { continue; }
             }
          } else if (fTahun || fBulan) { continue; }
      }

      // 2. FILTER UNIT (FIX: SEKARANG BERFUNGSI)
      // Kolom 0 = Unit Kerja
      if (fUnit && String(row[0]) !== fUnit) continue;

      // 3. FILTER STATUS (FIX: SEKARANG BERFUNGSI)
      // Kolom 10 = Status
      if (fStatus && String(row[10]) !== fStatus) continue;

      // Masukkan Data
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
    return JSON.stringify([{ status: 'Error', nama: e.message }]);
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
    var ss = SpreadsheetApp.openById(ID_DB);
    var sheet = ss.getSheetByName(SHEET_NAME);
    var baris = parseInt(form.recId);
    
    if (isNaN(baris)) throw new Error("ID Data tidak valid");
    var rangeLama = sheet.getRange(baris, 1, 1, 16); 
    var valLama = rangeLama.getValues()[0];
    
    var stLama = valLama[10]; 
    if(stLama === "OK" || stLama === "Disetujui") throw new Error("Data sudah disetujui, tidak bisa diedit.");
    
    var stBaru = (stLama === "Ditolak" || stLama === "Revisi") ? "Diproses" : "Diproses";

    // Format Tanggal Baru
    var tglSimpan = "";
    if (form.tanggal && form.tanggal.includes("-")) {
       var parts = form.tanggal.split("-"); 
       tglSimpan = parts[2] + "-" + parts[1] + "-" + parts[0];
    } else { tglSimpan = form.tanggal; }

    var jamSimpan = "";
    if (form.waktu) {
        var jamParts = form.waktu.split(":");
        jamSimpan = String(jamParts[0]).padStart(2, '0') + ":" + String(jamParts[1]).padStart(2, '0');
    }

    // --- LOGIKA FILE UPDATE (FORMAT NIP) ---
    var finalUrl = valLama[9]; 
    var nipUser = valLama[2]; // Ambil NIP dari database (Kolom C / Index 2)
    
    // FORMAT BARU: <NIP> - <Tanggal> - <Jenis>.pdf
    var namaFileBaru = nipUser + " - " + tglSimpan + " - " + form.jenis + ".pdf";

    var targetFolder = getFolderTahunBulan(DRIVE_ID, tglSimpan);

    if (fileData && fileData.data) {
       // A. UPLOAD FILE BARU
       var blob = Utilities.newBlob(Utilities.base64Decode(fileData.data), fileData.mimeType, namaFileBaru);
       var newFile = targetFolder.createFile(blob);
       newFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
       finalUrl = newFile.getUrl();
       
    } else {
       // B. RENAME FILE LAMA (Jika tanggal/jenis berubah)
       var tglLamaSheet = String(valLama[3]).replace(/'/g, "");
       
       // Cek perubahan (Tanggal, Jenis, atau Nama File belum sesuai NIP)
       if (tglSimpan !== tglLamaSheet || form.jenis !== valLama[5]) {
           try { 
             var idFile = finalUrl.match(/[-\w]{25,}/);
             if(idFile) {
                 var fileDrive = DriveApp.getFileById(idFile[0]);
                 fileDrive.setName(namaFileBaru); // Rename jadi format NIP
                 
                 // Pindah folder jika tanggal berubah
                 if (tglSimpan !== tglLamaSheet) {
                     fileDrive.moveTo(targetFolder);
                 }
             }
           } catch(e) {}
       }
    }

    // Simpan ke Sheet
    sheet.getRange(baris, 4).setValue("'" + tglSimpan);      
    sheet.getRange(baris, 5).setValue("'" + jamSimpan);      
    sheet.getRange(baris, 6).setValue(form.jenis);   
    sheet.getRange(baris, 7).setValue("'" + form.komulatif); 
    sheet.getRange(baris, 10).setValue(finalUrl);    
    sheet.getRange(baris, 11).setValue(stBaru);      
    sheet.getRange(baris, 12).setValue(Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd-MM-yyyy HH:mm:ss"));        
    sheet.getRange(baris, 13).setValue(form.user_login);

    return "Sukses Data Berhasil Diupdate";
  } catch(e) { throw new Error("Gagal Update: " + e.message); }
}

// 6. SOFT DELETE (HAPUS KE TRASH)
function softDeleteLupaPresensi(form) {
  var ID_DB = "160IjN8aiDAgDYXjgDLStS4nCZLKn3Ny-dq3BOFAfDrU";
  var SHEET_NAME = "Lupa_Presensi";
  var SHEET_TRASH = "Trash";
  var TRASH_ROOT_ID = "1Hop5S8iFazx3I3pX9SJILNLBkn-eBNfP"; 

  try {
    var KODE_RAHASIA = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyyMMdd");
    if (String(form.kode).trim() !== KODE_RAHASIA) throw new Error("Kode Salah.");

    var ss = SpreadsheetApp.openById(ID_DB);
    var sheetSource = ss.getSheetByName(SHEET_NAME);
    var sheetTrash = ss.getSheetByName(SHEET_TRASH);
    if (!sheetTrash) sheetTrash = ss.insertSheet(SHEET_TRASH);

    var baris = parseInt(form.recId);
    var range = sheetSource.getRange(baris, 1, 1, 16);
    var values = range.getDisplayValues()[0]; 

    if (!values[2]) throw new Error("Data kosong.");

    // --- PINDAHKAN FILE KE TRASH (STRUKTUR TAHUN > BULAN) ---
    var fileUrl = values[9];
    var tglData = String(values[3]).replace(/'/g, ""); // "dd-mm-yyyy"
    
    if (fileUrl && String(fileUrl).includes("drive")) {
       try {
         var fid = fileUrl.match(/[-\w]{25,}/);
         if(fid) {
            var file = DriveApp.getFileById(fid[0]);
            
            // Cari/Buat Folder Tahun > Bulan di dalam TRASH
            var targetTrashFolder = getFolderTahunBulan(TRASH_ROOT_ID, tglData);
            
            file.moveTo(targetTrashFolder); // Pindah
         }
       } catch(e){}
    }

    // Persiapan Data Trash
    values[3] = "'" + values[3]; values[4] = "'" + values[4]; values[6] = "'" + values[6];
    var tglHapus = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd-MM-yyyy HH:mm:ss");
    var userHapus = form.user_login || "Guest";
    var alasanHapus = form.alasan || "-";
    
    var trashRow = values.concat([tglHapus, userHapus, alasanHapus]);

    sheetTrash.appendRow(trashRow); 
    sheetSource.deleteRow(baris);      

    return "Data berhasil dipindahkan ke Trash.";
  } catch (e) { throw new Error(e.message); }
}

// 7. SIMPAN DATA BARU
function simpanLupaPresensi(dataKirim) {
  var ID_DB = "160IjN8aiDAgDYXjgDLStS4nCZLKn3Ny-dq3BOFAfDrU";
  var SHEET_NAME = "Lupa_Presensi";
  var DRIVE_ID = "1h8LcyYYrdVmd-fDPdcZ47hT9--rLQ7Fa"; 

  try {
    var ss = SpreadsheetApp.openById(ID_DB);
    var sheet = ss.getSheetByName(SHEET_NAME);
    
    // Normalisasi Tanggal
    var tglSimpan = "";
    if (dataKirim.tanggal && dataKirim.tanggal.includes("-")) {
       var parts = dataKirim.tanggal.split("-");
       tglSimpan = parts[2] + "-" + parts[1] + "-" + parts[0];
    } else { tglSimpan = dataKirim.tanggal; }

    // Normalisasi Jam
    var jamSimpan = dataKirim.waktu;
    if (jamSimpan && jamSimpan.includes(":")) {
       var jamParts = jamSimpan.split(":");
       jamSimpan = String(jamParts[0]).padStart(2, '0') + ":" + String(jamParts[1]).padStart(2, '0');
    }

    // --- LOGIKA FILE (UPDATE NIP) ---
    var targetFolder = getFolderTahunBulan(DRIVE_ID, tglSimpan);
    
    // FORMAT BARU: <NIP> - <Tanggal> - <Jenis>.pdf
    var fileExt = dataKirim.file.name.split('.').pop();
    var fileNameBaru = dataKirim.nip_asn + " - " + tglSimpan + " - " + dataKirim.jenis + "." + fileExt;

    // Simpan File
    var fileBlob = Utilities.newBlob(Utilities.base64Decode(dataKirim.file.data), dataKirim.file.mimeType, dataKirim.file.name);
    var newFile = targetFolder.createFile(fileBlob).setName(fileNameBaru);
    
    // Set Public
    newFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    var fileUrl = newFile.getUrl();
    // --------------------------------

    var timestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd-MM-yyyy HH:mm:ss");
    var rowData = [
      dataKirim.unit_kerja, dataKirim.nama_asn, dataKirim.nip_asn,
      "'" + tglSimpan, "'" + jamSimpan, dataKirim.jenis, dataKirim.komulatif,
      timestamp, dataKirim.user_login, fileUrl, "Diproses",
      "", "", "", "", ""
    ];
    sheet.appendRow(rowData);
    return "Sukses Data Berhasil Disimpan";
    
  } catch (e) { throw new Error("Gagal Simpan: " + e.message); }
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