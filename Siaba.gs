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

/* --- FUNGSI UPDATE DATA (REVISI: LOGIKA STATUS DINAMIS) --- */
function updateSalahAbsen(form) {
  const ID_DB = "1TZGrMiTuyvh2Xbo44RhJuWlQnOC5LzClsgIoNKtRFkY"; // Pastikan ID benar
  try {
    var ss = SpreadsheetApp.openById(ID_DB);
    var sheet = ss.getSheetByName("Salah_Absen");
    if (!sheet) throw new Error("Sheet Salah_Absen tidak ditemukan");

    // 1. DATA TARGET
    var targetNip = String(form.nip_lama).trim();
    var targetTgl = String(form.tgl_lama).trim();
    var targetJam = String(form.jam_lama).trim();

    // 2. CARI BARIS & AMBIL STATUS LAMA
    var data = sheet.getDataRange().getDisplayValues();
    var barisKetemu = -1;
    var statusLama = "";

    for (var i = 1; i < data.length; i++) {
       var sheetNip = String(data[i][2]).trim();
       var sheetTgl = String(data[i][3]).trim().replace(/\//g, '-');
       var sheetJam = String(data[i][4]).trim().replace(/'/g, "").substring(0, 5);
       if (/^\d:\d{2}/.test(sheetJam)) sheetJam = "0" + sheetJam;

       if (sheetNip === targetNip && sheetTgl === targetTgl && sheetJam === targetJam) {
          barisKetemu = i + 1;
          statusLama = String(data[i][8]).trim(); // Ambil Status saat ini (Kolom I)
          break;
       }
    }

    if (barisKetemu === -1) {
      throw new Error("Data asli tidak ditemukan. Pastikan data belum berubah.");
    }

    // 3. LOGIKA PERUBAHAN STATUS (SESUAI REQUEST)
    var st = statusLama.toLowerCase();
    
    // Cek jika sudah OK/Disetujui, tolak edit dari sisi server (Double Protection)
    if (st.includes("ok") || st.includes("setuju") || st.includes("acc")) {
        return "Gagal: Data sudah Disetujui, tidak dapat diedit.";
    }

    var statusBaru = "Diproses"; // Default fallback

    if (st.includes("tolak") || st.includes("ditolak")) {
        // Ditolak -> Revisi
        statusBaru = "Revisi";
    } else if (st.includes("revisi")) {
        // Revisi -> Diproses
        statusBaru = "Diproses";
    } else if (st.includes("proses") || st.includes("diproses")) {
        // Diproses -> Tetap Diproses
        statusBaru = "Diproses";
    }

    // 4. PERSIAPAN DATA BARU
    var tglBaruIndo = "";
    if (form.tanggal) {
       var parts = form.tanggal.split('-'); 
       tglBaruIndo = parts[2] + '-' + parts[1] + '-' + parts[0];
    }
    var jamBaru = String(form.waktu).trim();
    if (/^\d:\d{2}/.test(jamBaru)) jamBaru = "0" + jamBaru;

    var userEdit = "Guest";
    try { var cu = getCurrentUser(); if (cu) userEdit = cu.fullName || "Guest"; } catch(e){}
    var tglEdit = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd-MM-yyyy HH:mm:ss");

    // 5. UPDATE KE SHEET (Lock Format dengan ')
    sheet.getRange(barisKetemu, 3).setValue("'" + form.nip_asn);   // NIP
    sheet.getRange(barisKetemu, 4).setValue("'" + tglBaruIndo);    // Tgl
    sheet.getRange(barisKetemu, 5).setValue("'" + jamBaru);        // Jam
    sheet.getRange(barisKetemu, 6).setValue(form.jenis);           // Jenis
    
    // Update Status Baru
    sheet.getRange(barisKetemu, 9).setValue(statusBaru);           // Status
    
    // Metadata
    sheet.getRange(barisKetemu, 11).setValue("'" + tglEdit);       // Tgl Edit
    sheet.getRange(barisKetemu, 12).setValue("'" + userEdit);      // User Edit

    return "Sukses! Data diperbarui menjadi status: " + statusBaru;
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

/* --- FUNGSI VERIFIKASI (REVISI: LOCK METADATA VERIFIKASI) --- */
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

    // 2. CARI BARIS (DISPLAY VALUES)
    var data = sheet.getDataRange().getDisplayValues();
    var barisKetemu = -1;

    for (var i = 1; i < data.length; i++) {
       var sheetNip = String(data[i][2]).trim();
       var sheetTgl = String(data[i][3]).trim().replace(/\//g, '-');
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

    // 3. UPDATE HANYA KOLOM STATUS & VERIFIKATOR (DATA UTAMA AMAN)
    var tglVerif = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd-MM-yyyy HH:mm:ss");
    
    // Update Status (Kolom I / Index 9)
    sheet.getRange(barisKetemu, 9).setValue(dataKirim.status);
    
    // Update Keterangan (Kolom J / Index 10)
    // Pakai tanda petik untuk teks bebas, jaga-jaga ada karakter aneh
    sheet.getRange(barisKetemu, 10).setValue("'" + dataKirim.ket);
    
    // Update Tgl Verif (Kolom M / Index 13) - Pakai Petik
    sheet.getRange(barisKetemu, 13).setValue("'" + tglVerif);
    
    // Update Admin Verif (Kolom N / Index 14) - Pakai Petik
    sheet.getRange(barisKetemu, 14).setValue("'" + dataKirim.admin);

    return "Sukses Verifikasi";
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

/* --- FUNGSI RESTORE (REVISI: LOCK FORMAT DARI TRASH KE SOURCE) --- */
function prosesRestoreSalahAbsen(nip, tgl, jam) {
  const ID_DB = "1TZGrMiTuyvh2Xbo44RhJuWlQnOC5LzClsgIoNKtRFkY";
  try {
    var ss = SpreadsheetApp.openById(ID_DB);
    var sheetTrash = ss.getSheetByName("Trash");
    var sheetSource = ss.getSheetByName("Salah_Absen");
    if (!sheetTrash || !sheetSource) throw new Error("Sheet database tidak ditemukan.");

    // Gunakan DisplayValues agar membaca teks apa adanya dari Trash
    var dataDisplay = sheetTrash.getDataRange().getDisplayValues();
    var barisKetemu = -1;

    // 1. NORMALISASI TARGET PENCARIAN
    var targetNip = String(nip).trim();
    var targetTgl = String(tgl).trim();
    var targetJam = String(jam).trim();

    // 2. CARI DATA DI TRASH
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

    // 3. AMBIL DATA BARIS SEBAGAI TEKS (PENTING: getDisplayValues)
    // Kita ambil 14 kolom pertama (sesuai struktur tabel Salah_Absen)
    // Kolom selanjutnya di Trash adalah metadata hapus (tidak perlu dikembalikan)
    var rowValues = sheetTrash.getRange(barisKetemu, 1, 1, 14).getDisplayValues()[0];

    // 4. KUNCI FORMAT DATA (TAMBAHKAN PETIK SATU)
    // Ini menjamin NIP "001" tetap "001", bukan "1"
    // Dan Tanggal "02-01-2026" tetap teks, bukan Date Object
    rowValues[2] = "'" + rowValues[2]; // NIP
    rowValues[3] = "'" + rowValues[3]; // Tanggal
    rowValues[4] = "'" + rowValues[4]; // Jam
    rowValues[6] = "'" + rowValues[6]; // Tgl Kirim
    
    // Jika ada tanggal edit/verif, kunci juga
    if(rowValues[10]) rowValues[10] = "'" + rowValues[10]; // Tgl Edit
    if(rowValues[12]) rowValues[12] = "'" + rowValues[12]; // Tgl Verif

    // 5. KEMBALIKAN KE SHEET UTAMA
    sheetSource.appendRow(rowValues);
    
    // 6. HAPUS DARI TRASH
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
   MODUL: PENGAJUAN CUTI (BACKEND)
   ====================================================================== */

/* 1. AMBIL DATABASE REFERENSI (Optimized) */
function getDatabaseCutiOptions() {
  try {
    var ss = SpreadsheetApp.openById(ID_SS_CUTI); // ID Spreadsheet Cuti
    var sheet = ss.getSheetByName("Database Cuti");
    if (!sheet) return JSON.stringify([]);

    // Ambil semua data sekaligus (Cache friendly)
    // Kolom A=NIP, B=Unit, C=Nama, D=Status, ..., I=Alamat, J=HP
    var data = sheet.getDataRange().getValues();
    var result = [];

    for (var i = 1; i < data.length; i++) { // Skip Header
      // Pastikan ada NIP dan Nama
      if (data[i][0] && data[i][2]) {
        result.push({
          nip: String(data[i][0]),      // Kolom A
          unit: String(data[i][1]),     // Kolom B
          nama: String(data[i][2]),     // Kolom C
          status: String(data[i][3]),   // Kolom D (Status Kepegawaian)
          alamat: String(data[i][8]),   // Kolom I (Index 8)
          hp: String(data[i][9])        // Kolom J (Index 9)
        });
      }
    }
    return JSON.stringify(result);
  } catch (e) { return JSON.stringify([]); }
}

/* 2. SIMPAN PENGAJUAN CUTI */
function simpanPengajuanCuti(payload) {
  try {
    var ss = SpreadsheetApp.openById(ID_SS_CUTI);
    var sheet = ss.getSheetByName("Form Cuti");
    if (!sheet) return "Error: Sheet 'Form Cuti' tidak ditemukan.";

    var now = new Date();
    var sysDateStr = "'" + Utilities.formatDate(now, Session.getScriptTimeZone(), "dd-MM-yyyy HH:mm:ss");
    var userName = payload.userInput || "User Web";

    // 1. FORMAT TANGGAL
    var tglMulaiIndo   = formatIndoText(payload.tglMulai);
    var tglSelesaiIndo = formatIndoText(payload.tglSelesai);
    var tglSuratIndo   = formatIndoText(Utilities.formatDate(now, Session.getScriptTimeZone(), "yyyy-MM-dd"));

    // 2. AMBIL DATA DETIL PEGAWAI
    // Kita gunakan dbData.jabatan dan dbData.golongan yang sudah pasti benar mappingnya
    var dbData = getDetailPegawaiByNip(payload.nip); 
    
    // --- PERBAIKAN DI SINI (AMBIL DARI PROPERTI OBJECT, JANGAN FULLROW MANUAL) ---
    // Pastikan getDetailPegawaiByNip mappingnya: jabatan = data[i][5] (Kolom F)
    var empGol = dbData ? dbData.golongan : ""; 
    var empJab = dbData ? dbData.jabatan : "";  // INI KUNCI PERBAIKAN KEPALA SD
    
    // 3. LOOKUP PEJABAT STRUKTURAL
    // Kirim jabatan yang benar (Col F) ke fungsi lookup
    var pejabat = lookupPejabatStruktural(payload.jenisCuti, payload.unit, empGol, empJab);
    
    var final_kepada, final_nm_ats, final_nip_ats, final_jab_ats, final_nm_stj, final_nip_stj, final_jab_stj;

    if (pejabat) {
        // KASUS 1: Ditemukan Aturan Khusus di Sheet "Data Atasan"
        final_kepada      = pejabat.kepada;
        final_nama_atasan = pejabat.nama_atasan;
        final_nip_atasan  = pejabat.nip_atasan;
        final_jab_atasan  = pejabat.jabatan_atasan;
        final_nama_setuju = pejabat.nama_setuju;
        final_nip_setuju  = pejabat.nip_setuju;
        final_jab_setuju  = pejabat.jabatan_setuju;
    } else {
        // KASUS 2: Default (Ambil Atasan Langsung dari Database Cuti)
        // Col 20(T), 14(N), 15(O), 16(P), 17(Q), 18(R), 19(S)
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

    // 5. CHECKLIST JENIS CUTI (UPDATE: UMROH JUGA CENTANG CT)
    var j = String(payload.jenisCuti).toLowerCase();
    var c = { ct:"", cs:"", cap:"", cb:"", cm:"", cltn:"" }; 
    var CHECK = "✓"; 
    
    // Logika CT: Jika Tahunan ATAU Umroh -> Centang CT
    if (j.includes("tahunan") || j.includes("umroh")) c.ct = CHECK;
    
    if (j.includes("sakit")) c.cs = CHECK;
    else if (j.includes("penting")) c.cap = CHECK;
    else if (j.includes("besar")) c.cb = CHECK;
    else if (j.includes("melahirkan")) c.cm = CHECK;
    else if (j.includes("luar") || j.includes("tanggungan")) c.cltn = CHECK;

    // 6. SUSUN DATA PDF
    var pdfData = {
        tanggal: tglSuratIndo,
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

        jenisCutiRaw: payload.jenisCuti, // Contoh: "Cuti Tahunan"
        tglMulaiRaw: payload.tglMulai    // Contoh: "2026-01-21" (Format yyyy-mm-dd)
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

      tglSuratIndo,      // V
      pdfData.jabatan,   // W
      pdfData.masa_kerja,// X
      pdfData.unit,      // Y (Gunakan Unit Lengkap)
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

/* 2. UPDATE FUNGSI INI (Folder & Filename Logic) */
function generatePdfCuti(data) {
  // --- KONFIGURASI ---
  var ID_TEMPLATE = "1k5KmEZj5nikuUV-MLnY4c6Tn-jFIhmOMGwhjvqaUSzk"; 
  var ID_FOLDER_INDUK = "1suNhGklZ931kT6Y5wbp5x_92ZCtlWfQz"; // Folder Induk Anda
  var ID_IMAGE_CHECK = "1AbFps5ZiyeBH9hVa_XTYvfnoO77DxFle";

  try {
    var templateFile = DriveApp.getFileById(ID_TEMPLATE);
    var indukFolder = DriveApp.getFolderById(ID_FOLDER_INDUK);
    var checkImgBlob = DriveApp.getFileById(ID_IMAGE_CHECK).getBlob();

    // --- LOGIKA SUBFOLDER (TAHUN > BULAN) ---
    // 1. Ambil Tahun dan Bulan dari tglMulaiRaw (yyyy-mm-dd)
    var parts = data.tglMulaiRaw.split("-"); // ["2026", "01", "21"]
    var year = parts[0]; 
    var monthIndex = parseInt(parts[1]) - 1; // 0-11
    var monthNames = ["Januari", "Februari", "Maret", "April", "Mei", "Juni", "Juli", "Agustus", "September", "Oktober", "November", "Desember"];
    var monthName = monthNames[monthIndex];

    // 2. Cek/Buat Folder Tahun di dalam Folder Induk
    var yearFolder = getOrCreateSubfolder(indukFolder, year);

    // 3. Cek/Buat Folder Bulan di dalam Folder Tahun
    var targetFolder = getOrCreateSubfolder(yearFolder, monthName);

    // --- LOGIKA NAMA FILE BARU ---
    // Format: <Jenis Cuti> - <Nama ASN> - <Tanggal Mulai Indo>
    // Contoh: Cuti Tahunan - Budi - 21 Januari 2026.pdf
    var fileName = data.jenisCutiRaw + " - " + data.asn + " - " + data.tmc + ".pdf";

    // --- PROSES PEMBUATAN PDF ---
    var tempFile = templateFile.makeCopy(fileName, targetFolder);
    var tempDoc = DocumentApp.openById(tempFile.getId());
    var body = tempDoc.getBody();
    
    // Loop Replace Data
    for (var key in data) {
      if (data.hasOwnProperty(key)) {
        var val = data[key];

        // Replace Centang dengan Gambar
        if (["ct","cs","cb","cm","cap","cltn"].indexOf(key) > -1) {
            if (val === "✓") {
                replaceTextWithImage(body, "{{" + key + "}}", checkImgBlob);
            } else {
                body.replaceText("{{" + key + "}}", ""); 
            }
        } 
        // Jangan replace key raw (hanya internal script)
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
  // Cari penemuan pertama
  var next = body.findText(placeholder);
  
  // Lakukan Loop selama placeholder masih ditemukan di dokumen
  while (next) {
    var element = next.getElement();
    var start = next.getStartOffset();
    var end = next.getEndOffsetInclusive();
    
    // 1. Hapus Teks Placeholder {{ct}} / {{cs}} dll
    element.deleteText(start, end);
    
    // 2. Sisipkan Gambar di posisi tersebut
    var img = element.getParent().asParagraph().insertInlineImage(start, imgBlob);
    
    // 3. ATUR UKURAN LEBIH KECIL (Revisi User)
    // Ukuran 11x11 atau 12x12 biasanya pas untuk kotak centang
    img.setWidth(11).setHeight(11); 
    
    // 4. Cari lagi placeholder berikutnya (agar bagian V juga kena)
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
        // Mapping: A=0, ..., F=5 (Jabatan), G=6 (Unit Lengkap)
        golongan:  data[i][4],  // Col E
        jabatan:   data[i][5],  // Col F (JABATAN YANG BENAR)
        unitLengkap: data[i][6], // Col G
        masaKerja: data[i][7],  // Col H
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
  
  // Normalisasi Input User
  var j = String(jenisCuti).toLowerCase().trim();
  var t = String(tugasUser).toLowerCase().trim(); // Jabatan dari DB Cuti Col F
  var g = String(golUser).toLowerCase().trim();
  var u = String(unitUser).toLowerCase().trim();

  // --- LOGIKA PRIORITAS (SESUAI REQUEST) ---

  // 1. CEK CUTI UMROH (Prioritas Tertinggi)
  if (j === "cuti umroh") {
    // Cari baris yang Kolom A = "Cuti Umroh"
    for (var i = 1; i < data.length; i++) {
       if (String(data[i][0]).toLowerCase().trim() === "cuti umroh") return mapRow(data[i]);
    }
  }

  // 2. CEK GOLONGAN IV (IV/a, IV/b, IV/c)
  // Syarat: Bukan Umroh (sudah lewat di atas) & Golongan mengandung "iv/"
  if (g.includes("iv/") || g === "iv") {
    // Cari baris di Excel yang Kolom C-nya diisi "IV" atau "IV/"
    // Kita cari baris yang Kolom C-nya cocok dengan golongan user
    for (var i = 1; i < data.length; i++) {
       var ruleGol = String(data[i][2]).toLowerCase().trim();
       // Jika Excel C="iv" dan User="iv/a" -> COCOK (User includes Rule)
       if (ruleGol !== "" && g.includes(ruleGol)) return mapRow(data[i]);
    }
  }

  // 3. CEK TUGAS "KEPALA SD" (Prioritas Ketiga)
  // Syarat: Bukan Gol IV (sudah lewat) & Jabatan mengandung "Kepala"
  // Logika: Kita cari baris di Excel yg Kolom B-nya tidak kosong.
  // Jika Jabatan User MENGANDUNG apa yang tertulis di Kolom B, maka MATCH.
  for (var i = 1; i < data.length; i++) {
     var ruleTugas = String(data[i][1]).toLowerCase().trim(); // Kolom B
     
     // Contoh: Rule="kepala sd", User="kepala sd negeri 1" -> MATCH
     if (ruleTugas !== "" && t.includes(ruleTugas)) {
        return mapRow(data[i]);
     }
  }

  // 4. CEK UNIT KERJA (Default)
  // Syarat: Unit User cocok dengan Kolom D
  for (var i = 1; i < data.length; i++) {
     var ruleUnit = String(data[i][3]).toLowerCase().trim(); // Kolom D
     if (ruleUnit !== "" && ruleUnit === u) return mapRow(data[i]);
  }

  return null; // Fallback ke Atasan Langsung (Database Cuti)
}

// Helper Mapping (Tetap sama, pastikan ada di bawah)
function mapRow(row) {
  return {
     nama_atasan:    row[4], // E
     nip_atasan:     row[5], // F
     jabatan_atasan: row[6], // G
     nama_setuju:    row[7], // H
     nip_setuju:     row[8], // I
     jabatan_setuju: row[9], // J
     kepada:         row[10] // K
  };
}

// Helper Pencarian Baris di Data Atasan
function cariBarisDataAtasan(allData, kriteria, value) {
  for (var i = 1; i < allData.length; i++) { // Skip Header
    var r = allData[i];
    // A=0(Jenis), B=1(Tugas), C=2(Gol), D=3(Unit)
    
    if (kriteria === "jenis" && String(r[0]).toLowerCase().includes(value)) return mapRow(r);
    if (kriteria === "gol"   && String(r[2]).toLowerCase() !== "" && value.includes(String(r[2]).toLowerCase())) return mapRow(r);
    if (kriteria === "tugas" && String(r[1]).toLowerCase().includes(value)) return mapRow(r);
    if (kriteria === "unit"  && String(r[3]).toLowerCase() === value) return mapRow(r);
  }
  return null;
}

/* HELPER: KONVERSI TANGGAL "YYYY-MM-DD" KE "d MMMM yyyy" (Indo) */
function formatIndoText(isoDate) {
  if (!isoDate) return "";
  var parts = isoDate.split("-"); // [2026, 01, 20]
  if (parts.length !== 3) return isoDate;
  
  var months = ["Januari", "Februari", "Maret", "April", "Mei", "Juni", "Juli", "Agustus", "September", "Oktober", "November", "Desember"];
  var y = parts[0];
  var m = parseInt(parts[1], 10) - 1;
  var d = parseInt(parts[2], 10);
  
  return d + " " + months[m] + " " + y;
}

function getOrCreateSubfolder(parentFolder, folderName) {
  var folders = parentFolder.getFoldersByName(folderName);
  
  if (folders.hasNext()) {
    return folders.next(); // Folder sudah ada, pakai yang itu
  } else {
    return parentFolder.createFolder(folderName); // Belum ada, buat baru
  }
}

/* ======================================================================
   MODUL: UPDATE / EDIT CUTI
   ====================================================================== */

function updatePengajuanCuti(payload) {
  try {
    var ss = SpreadsheetApp.openById(ID_SS_CUTI);
    var sheet = ss.getSheetByName("Form Cuti");
    if (!sheet) return "Error: Sheet 'Form Cuti' tidak ditemukan.";
    
    var rowIndex = parseInt(payload.rowBaris);
    if (!rowIndex || rowIndex < 2) return "Error: Baris data tidak valid.";

    var now = new Date();
    var tglEditStr = "'" + Utilities.formatDate(now, Session.getScriptTimeZone(), "dd-MM-yyyy HH:mm:ss");
    var userEdit = payload.userInput || "User Web";

    // 1. FORMAT TANGGAL
    var tglMulaiIndo   = formatIndoText(payload.tglMulai);
    var tglSelesaiIndo = formatIndoText(payload.tglSelesai);
    var tglSuratIndo   = formatIndoText(Utilities.formatDate(now, Session.getScriptTimeZone(), "yyyy-MM-dd"));

    // 2. AMBIL DATA DETIL PEGAWAI
    var dbData = getDetailPegawaiByNip(payload.nip); 
    var empGol = dbData ? dbData.golongan : ""; 
    var empJab = dbData ? dbData.jabatan : ""; 
    
    // 3. LOOKUP PEJABAT
    var pejabat = lookupPejabatStruktural(payload.jenisCuti, payload.unit, empGol, empJab);
    
    var final_kepada, final_nm_ats, final_nip_ats, final_jab_ats, final_nm_stj, final_nip_stj, final_jab_stj;

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

    // 5. CHECKLIST JENIS CUTI (PERBAIKAN LOGIKA MUTLAK/STRICT)
    // Pastikan string bersih dari spasi
    var j = String(payload.jenisCuti).toLowerCase().trim();
    var c = { ct:"", cs:"", cap:"", cb:"", cm:"", cltn:"" }; 
    var CHECK = "✓"; 
    
    // Gunakan ELSE IF agar hanya SATU kondisi yang terpenuhi
    if (j.includes("sakit")) {
        c.cs = CHECK;
    } else if (j.includes("penting")) {
        c.cap = CHECK;
    } else if (j.includes("besar")) {
        c.cb = CHECK;
    } else if (j.includes("melahirkan")) {
        c.cm = CHECK;
    } else if (j.includes("luar") || j.includes("tanggungan")) {
        c.cltn = CHECK;
    } else {
        // Default / Fallback: Jika mengandung "tahunan" atau "umroh", atau sisanya
        c.ct = CHECK;
    }

    // 6. SUSUN DATA PDF
    var pdfData = {
        tanggal: tglSuratIndo,
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
        
        // Data Raw untuk Penamaan File
        jenisCutiRaw: payload.jenisCuti, 
        tglMulaiRaw: payload.tglMulai    
    };
    
    // GENERATE PDF BARU
    var linkPdf = generatePdfCuti(pdfData); 

    // 7. UPDATE SPREADSHEET
    
    // A. Update Data Utama
    var rangeUtama = sheet.getRange(rowIndex, 1, 1, 10);
    rangeUtama.setValues([[
        payload.unit, payload.nama, "'" + payload.nip, payload.jenisCuti, 
        tglMulaiIndo, tglSelesaiIndo, payload.jumlahHari, payload.alasan, 
        payload.alamat, "'" + payload.hp
    ]]);

    // B. Update Status & Link PDF (Mereset Status jadi Diproses)
    sheet.getRange(rowIndex, 11, 1, 3).setValues([["Diproses", "", linkPdf]]);

    // C. Update Log Edit
    sheet.getRange(rowIndex, 16, 1, 2).setValues([[tglEditStr, userEdit]]);

    // D. Update Data Kolom Belakang (V - AO)
    var rangeExtra = sheet.getRange(rowIndex, 22, 1, 20);
    rangeExtra.setValues([[
      tglSuratIndo,      // V
      pdfData.jabatan,   // W
      pdfData.masa_kerja,// X
      pdfData.unit,      // Y
      c.ct, c.cb, c.cs, c.cm, c.cap, c.cltn, // Z - AE
      sisaN2, sisaN1, sisaN,
      final_jab_atasan, final_nama_atasan, final_nip_atasan,
      final_jab_setuju, final_nama_setuju, final_nip_setuju,
      final_kepada       // AO
    ]]);

    SpreadsheetApp.flush();
    return "Sukses";
    
  } catch (e) { return "Error Update: " + e.toString(); }
}

/* ======================================================================
   MODUL: HAPUS DATA (DENGAN KODE KEAMANAN HARIAN)
   ====================================================================== */

function hapusPengajuanCuti(rowBaris, kodeInput) {
  try {
    // 1. GENERATE KODE RAHASIA HARI INI (YYYYMMDD)
    var now = new Date();
    var validCode = Utilities.formatDate(now, Session.getScriptTimeZone(), "yyyyMMdd");
    
    // 2. VALIDASI INPUT USER
    // Pastikan input string dan bersih dari spasi
    if (String(kodeInput).trim() !== validCode) {
      return "KODE_SALAH";
    }

    // 3. PROSES HAPUS
    var ss = SpreadsheetApp.openById(ID_SS_CUTI);
    var sheet = ss.getSheetByName("Form Cuti");
    if (!sheet) return "Error: Sheet tidak ditemukan.";
    
    // Validasi Baris (Jangan sampai menghapus Header di baris 1)
    var row = parseInt(rowBaris);
    if (isNaN(row) || row < 2) return "Error: Baris data tidak valid.";
    
    sheet.deleteRow(row);
    
    return "Sukses";
    
  } catch (e) {
    return "Error Hapus: " + e.toString();
  }
}

/* ======================================================================
   MODUL: VERIFIKASI ADMIN (UBAH STATUS)
   ====================================================================== */

function verifikasiPengajuan(rowBaris, status, catatan, adminName) {
  try {
    var ss = SpreadsheetApp.openById(ID_SS_CUTI);
    var sheet = ss.getSheetByName("Form Cuti");
    if (!sheet) return "Error: Sheet tidak ditemukan.";
    
    var row = parseInt(rowBaris);
    if (isNaN(row) || row < 2) return "Error: Baris tidak valid.";

    // Update Data
    // Kolom K (11) = Status
    // Kolom L (12) = Catatan (Keterangan)
    sheet.getRange(row, 11).setValue(status);
    sheet.getRange(row, 12).setValue(catatan);
    
    // Update Metadata Verifikasi
    // Kolom R (18) = Tgl Verif
    // Kolom S (19) = Verifikator
    var now = new Date();
    var sysDateStr = "'" + Utilities.formatDate(now, Session.getScriptTimeZone(), "dd-MM-yyyy HH:mm:ss");
    
    sheet.getRange(row, 18).setValue(sysDateStr);
    sheet.getRange(row, 19).setValue(adminName || "Admin");
    
    return "Sukses";
    
  } catch (e) { return "Error Verif: " + e.toString(); }
}

/* 1. UPDATE FUNGSI OPSI UNIT (SUMBER: DATABASE CUTI KOLOM B) */
function getUnitOptions() {
  var ss = SpreadsheetApp.openById(ID_SS_CUTI);
  var sheet = ss.getSheetByName("Database Cuti"); // GANTI SHEET
  if (!sheet) return [];
  
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  
  // AMBIL KOLOM B (Index 1) sesuai permintaan
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

/* 2. UPDATE FUNGSI GET DATA (LOGIKA FILTER LEBIH AMAN) */
function getDataCuti(tahun, bulan, unitFilter) { 
  var ss = SpreadsheetApp.openById(ID_SS_CUTI);
  var sheet = ss.getSheetByName("Form Cuti");
  if (!sheet) return JSON.stringify([]);
  
  // Ambil semua data & display values (agar tanggal terbaca sebagai teks apa adanya)
  var data = sheet.getDataRange().getValues();
  // Opsional: Gunakan getDisplayValues jika ingin format teks persis seperti di layar Excel
  var displayData = sheet.getDataRange().getDisplayValues(); 
  
  var result = [];
  
  // Normalisasi Filter
  var fTahun = tahun ? String(tahun).trim() : "";
  var fBulan = bulan ? String(bulan).toLowerCase().trim() : "";
  var fUnit  = unitFilter ? String(unitFilter).toLowerCase().trim() : "";

  // Loop mulai baris ke-2 (Index 1)
  for (var i = 1; i < data.length; i++) {
    var row = data[i];     // Data asli (mungkin object date)
    var rowTxt = displayData[i]; // Data teks (tampilan)
    
    // Index Kolom: A=0(Unit), E=4(TglMulai)
    var rowUnitRaw = String(row[0]).toLowerCase();
    
    // Ambil teks tanggal dari displayData agar lebih akurat sesuai tampilan
    var tglMulaiTxt = String(rowTxt[4]).trim(); // "21 Januari 2026"
    
    // Deteksi Tahun & Bulan
    var rTahun = "";
    var rBulan = "";
    
    // Coba parsing format "dd MMMM yyyy" (Spasi)
    var parts = tglMulaiTxt.split(" ");
    if (parts.length >= 3) {
       rBulan = parts[1].toLowerCase(); // "januari"
       rTahun = parts[2]; // "2026"
    } 
    // Fallback: Jika gagal split spasi, mungkin format "yyyy-mm-dd" atau lainnya
    else {
       // Coba baca dari object Date jika ada
       if (row[4] instanceof Date) {
          rTahun = String(row[4].getFullYear());
          var mIndex = row[4].getMonth();
          var mNames = ["januari","februari","maret","april","mei","juni","juli","agustus","september","oktober","november","desember"];
          rBulan = mNames[mIndex];
       }
    }

    // --- LOGIKA FILTER AMAN ---
    
    // 1. Cek Tahun: Lolos jika filter kosong ATAU tahun cocok ATAU tahun tidak terdeteksi (tampilkan saja drpd hilang)
    var matchTahun = (fTahun === "") || (rTahun === fTahun);
    
    // 2. Cek Bulan: Lolos jika filter kosong ATAU bulan cocok
    var matchBulan = (fBulan === "") || (rBulan === fBulan);
    
    // 3. Cek Unit: Lolos jika filter kosong ATAU teks unit mengandung kata kunci
    var matchUnit = (fUnit === "") || (rowUnitRaw.indexOf(fUnit) > -1);

    // Gabungkan
    if (matchTahun && matchBulan && matchUnit) {
      result.push({
        rowBaris: i + 1,
        unit: rowTxt[0],   // Gunakan rowTxt agar format terjaga
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
        fileUrl: rowTxt[12], // Link PDF
        tglInput: rowTxt[13],
        userInput: rowTxt[14],
        tglEdit: rowTxt[15],
        userEdit: rowTxt[16],
        tglVerif: rowTxt[17],
        verifikator: rowTxt[18]
      });
    }
  }
  
  // Urutkan Terbaru di Atas
  result.reverse();
  return JSON.stringify(result);
}

// Helper Format Tanggal Pendek (dd/MM/yy HH:mm)
function formatDateShort(dateObj) {
  if (!dateObj) return "";
  return Utilities.formatDate(new Date(dateObj), Session.getScriptTimeZone(), "dd/MM/yy HH:mm");
}