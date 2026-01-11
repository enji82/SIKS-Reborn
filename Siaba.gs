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

// --- BAGIAN SERVER CRUD SALAH ABSEN ---

function getDaftarSalahPresensi() {
  try {
    var ss = SpreadsheetApp.openById("1TZGrMiTuyvh2Xbo44RhJuWlQnOC5LzClsgIoNKtRFkY");
    var sheet = ss.getSheetByName("Salah_Absen"); 
    
    if (!sheet) return JSON.stringify([]); 

    var lastRow = sheet.getLastRow();
    if (lastRow < 2) return JSON.stringify([]);

    var dataRange = sheet.getRange(2, 1, lastRow - 1, 14); 
    var displayValues = dataRange.getDisplayValues(); 

    var output = [];

    for (var i = 0; i < displayValues.length; i++) {
      var row = displayValues[i];
      if (row[1] === "") continue;

      var rowData = {
        rowBaris: i + 2,     
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
    return JSON.stringify([{ error: true, message: e.toString() }]);
  }
}

function getDatabasePegawai() {
  const ID_DB = "1TZGrMiTuyvh2Xbo44RhJuWlQnOC5LzClsgIoNKtRFkY";
  
  try {
    var ss = SpreadsheetApp.openById(ID_DB);
    var sheet = ss.getSheetByName("Database"); 
    
    if (!sheet) throw new Error("Sheet 'Database' tidak ditemukan!");

    var lastRow = sheet.getLastRow();
    if (lastRow < 2) return []; 

    var data = sheet.getRange(2, 1, lastRow - 1, 3).getValues();
    
    return data.map(function(r) {
      return {
        unit: r[0], 
        nip:  r[1], 
        nama: r[2]  
      };
    });
  } catch (e) {
    throw new Error("Gagal ambil data: " + e.message);
  }
}

function simpanSalahAbsen(form) {
  const ID_DB = "1TZGrMiTuyvh2Xbo44RhJuWlQnOC5LzClsgIoNKtRFkY";

  try {
    var ss = SpreadsheetApp.openById(ID_DB);
    var sheet = ss.getSheetByName("Salah_Absen");
    
    if (!sheet) throw new Error("Sheet 'Salah_Absen' tidak ditemukan!");
    
    var tglIndo = "";
    if (form.tanggal) {
       var parts = form.tanggal.split('-'); 
       tglIndo = parts[2] + '-' + parts[1] + '-' + parts[0]; 
    }
    
    // Paksa menjadi string agar 09:30 tidak jadi 9:30
    var jam = "'" + form.waktu; 

    var namaUser = "Guest";
    try {
       var currentUser = getCurrentUser(); 
       if (currentUser && currentUser.fullName) {
          namaUser = currentUser.fullName;
       }
    } catch (err) {
       namaUser = "Guest (Error User)";
    }
    
    var tglKirim = new Date();
    var status = "Diproses";

    var barisBaru = [
      form.unit_kerja, 
      form.nama_asn,   
      "'"+form.nip_asn,
      tglIndo,         
      jam,             
      form.jenis,      
      tglKirim,        
      namaUser,        
      status           
    ];
    
    sheet.appendRow(barisBaru);
    return "SUKSES";
    
  } catch (e) {
    throw new Error("Gagal simpan: " + e.message);
  }
}

/* --- FUNGSI UPDATE DATA (VERSI FINAL: AUTO UPDATE STATUS) --- */
function updateSalahAbsen(form) {
  const ID_DB = "1TZGrMiTuyvh2Xbo44RhJuWlQnOC5LzClsgIoNKtRFkY"; 

  try {
    var ss = SpreadsheetApp.openById(ID_DB);
    var sheet = ss.getSheetByName("Salah_Absen");
    if (!sheet) throw new Error("Sheet Salah_Absen tidak ditemukan");

    // 1. DATA TARGET (DARI CLIENT)
    var targetNip = String(form.nip_lama).trim();
    var targetTgl = String(form.tgl_lama).trim();
    var targetJam = String(form.jam_lama).trim();

    // 2. DATA BARU (UNTUK DISIMPAN)
    var parts = form.tanggal.split('-'); 
    var tglBaruIndo = parts[2] + '-' + parts[1] + '-' + parts[0]; 
    
    var userEdit = "Guest";
    try {
       var currentUser = getCurrentUser(); 
       if (currentUser && currentUser.fullName) userEdit = currentUser.fullName;
    } catch (e) {}
    var tglEdit = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd-MM-yyyy HH:mm:ss"); 

    // 3. LOOPING PENCARIAN
    var data = sheet.getDataRange().getValues();
    var barisKetemu = -1;
    var statusSaatIni = ""; // Variabel untuk menyimpan status lama

    for (var i = 1; i < data.length; i++) {
       var sheetNip = String(data[i][2]).replace(/'/g, "").trim(); 
       
       var sheetTglRaw = data[i][3];
       var sheetTglStr = (sheetTglRaw instanceof Date) ? 
           Utilities.formatDate(sheetTglRaw, Session.getScriptTimeZone(), "dd-MM-yyyy") : String(sheetTglRaw).trim();
       
       var sheetJamRaw = data[i][4];
       var sheetJamStr = "";
       if (sheetJamRaw instanceof Date) {
         sheetJamStr = Utilities.formatDate(sheetJamRaw, Session.getScriptTimeZone(), "HH:mm");
       } else {
         sheetJamStr = String(sheetJamRaw).trim();
         if (/^\d:\d{2}/.test(sheetJamStr)) sheetJamStr = "0" + sheetJamStr;
       }

       if (sheetNip === targetNip && sheetTglStr === targetTgl && sheetJamStr === targetJam) {
          barisKetemu = i + 1; 
          statusSaatIni = String(data[i][8]).trim(); // Ambil Status (Kolom I / Index 8)
          break;
       }
    }

    if (barisKetemu === -1) {
      throw new Error(`Data tidak ditemukan. Mungkin data telah berubah.`);
    }

    // 4. LOGIKA PERUBAHAN STATUS OTOMATIS
    // Ditolak -> Revisi
    // Revisi -> Diproses
    // Diproses -> Tetap Diproses
    
    var statusBaru = statusSaatIni; // Default sama

    if (statusSaatIni === "Ditolak") {
        statusBaru = "Revisi";
    } else if (statusSaatIni === "Revisi") {
        statusBaru = "Diproses";
    } else if (statusSaatIni === "Diproses") {
        statusBaru = "Diproses";
    }

    // 5. EKSEKUSI UPDATE
    // Update Tanggal (Kolom D / 4)
    sheet.getRange(barisKetemu, 4).setValue(tglBaruIndo);
    // Update Jam (Kolom E / 5)
    sheet.getRange(barisKetemu, 5).setValue("'" + form.waktu);
    // Update Jenis (Kolom F / 6)
    sheet.getRange(barisKetemu, 6).setValue(form.jenis);
    
    // Update Status (Kolom I / 9) - HASIL LOGIKA BARU
    sheet.getRange(barisKetemu, 9).setValue(statusBaru);

    // Log Edit
    sheet.getRange(barisKetemu, 11).setValue(tglEdit);
    sheet.getRange(barisKetemu, 12).setValue(userEdit);

    return "Data Berhasil Diperbarui!";

  } catch (e) {
    throw new Error(e.message);
  }
}

function simpanSalahAbsen(form) {
  const ID_DB = "1TZGrMiTuyvh2Xbo44RhJuWlQnOC5LzClsgIoNKtRFkY";

  try {
    var ss = SpreadsheetApp.openById(ID_DB);
    var sheet = ss.getSheetByName("Salah_Absen");
    
    if (!sheet) throw new Error("Sheet 'Salah_Absen' tidak ditemukan!");
    
    var tglIndo = "";
    if (form.tanggal) {
       var parts = form.tanggal.split('-'); 
       tglIndo = parts[2] + '-' + parts[1] + '-' + parts[0]; 
    }
    
    // Paksa menjadi string agar 09:30 tidak jadi 9:30
    var jam = "'" + form.waktu; 

    var namaUser = "Guest";
    try {
       var currentUser = getCurrentUser(); 
       if (currentUser && currentUser.fullName) {
          namaUser = currentUser.fullName;
       }
    } catch (err) {
       namaUser = "Guest (Error User)";
    }
    
    var tglKirim = new Date();
    var status = "Diproses";

    var barisBaru = [
      form.unit_kerja, 
      form.nama_asn,   
      "'"+form.nip_asn,
      tglIndo,         
      jam,             
      form.jenis,      
      tglKirim,        
      namaUser,        
      status           
    ];
    
    sheet.appendRow(barisBaru);
    return "SUKSES";
    
  } catch (e) {
    throw new Error("Gagal simpan: " + e.message);
  }
}

/* --- FUNGSI SOFT DELETE (PINDAH KE TRASH) --- */
function softDeleteSalahAbsen(dataKirim) {
  const ID_DB = "1TZGrMiTuyvh2Xbo44RhJuWlQnOC5LzClsgIoNKtRFkY"; // ID Database Salah Absen

  try {
    var ss = SpreadsheetApp.openById(ID_DB);
    var sheetSource = ss.getSheetByName("Salah_Absen");
    var sheetTrash = ss.getSheetByName("Trash");

    // Jika sheet Trash belum ada, buat baru
    if (!sheetTrash) {
      sheetTrash = ss.insertSheet("Trash");
      // Opsional: Buat Header Trash jika baru dibuat
      // sheetTrash.appendRow(["Unit", "Nama", "NIP", "Tanggal", "Jam", "Jenis", "Tgl Ajuan", "User Input", "Status", "Ket", "Tgl Edit", "User Edit", "Tgl Verif", "Admin Verif", "TGL HAPUS", "USER HAPUS", "ALASAN"]);
    }

    if (!sheetSource) throw new Error("Sheet Salah_Absen tidak ditemukan");

    // 1. DATA TARGET
    var targetNip = String(dataKirim.nip).trim();
    var targetTgl = String(dataKirim.tgl).trim();
    var targetJam = String(dataKirim.jam).trim();

    // 2. CARI BARIS (LOGIKA SAMA DENGAN UPDATE AGAR AKURAT)
    var data = sheetSource.getDataRange().getValues();
    var barisKetemu = -1;

    for (var i = 1; i < data.length; i++) {
       var sheetNip = String(data[i][2]).replace(/'/g, "").trim(); 
       
       var sheetTglRaw = data[i][3];
       var sheetTglStr = (sheetTglRaw instanceof Date) ? 
           Utilities.formatDate(sheetTglRaw, Session.getScriptTimeZone(), "dd-MM-yyyy") : String(sheetTglRaw).trim();
       
       var sheetJamRaw = data[i][4];
       var sheetJamStr = "";
       if (sheetJamRaw instanceof Date) {
         sheetJamStr = Utilities.formatDate(sheetJamRaw, Session.getScriptTimeZone(), "HH:mm");
       } else {
         sheetJamStr = String(sheetJamRaw).trim();
         if (/^\d:\d{2}/.test(sheetJamStr)) sheetJamStr = "0" + sheetJamStr;
       }

       if (sheetNip === targetNip && sheetTglStr === targetTgl && sheetJamStr === targetJam) {
          barisKetemu = i + 1; 
          break;
       }
    }

    if (barisKetemu === -1) {
      throw new Error("Data yang akan dihapus tidak ditemukan/sudah berubah.");
    }

    // 3. AMBIL DATA BARIS TERSEBUT
    // Kita ambil range sesuai lebar kolom (misal kolom A sampai N = 14 kolom)
    var rowRange = sheetSource.getRange(barisKetemu, 1, 1, 14);
    var rowValues = rowRange.getValues()[0]; // Array 1D

    // 4. SIAPKAN DATA TRASH
    var tglHapus = new Date();
    var userHapus = dataKirim.user;
    var alasan = dataKirim.alasan;

    // Gabungkan: Data Asli + Tgl Hapus (Kolom O) + User (Kolom P) + Alasan (Kolom Q)
    var trashRow = rowValues.concat([tglHapus, userHapus, alasan]);

    // 5. PINDAHKAN
    sheetTrash.appendRow(trashRow); // Tulis ke Trash
    sheetSource.deleteRow(barisKetemu); // Hapus dari Sumber

    return "Sukses";

  } catch (e) {
    throw new Error(e.message);
  }
}

/* --- FUNGSI PROSES VERIFIKASI (ADMIN) --- */
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

    // 2. CARI BARIS (LOGIKA ROBUST)
    var data = sheet.getDataRange().getValues();
    var barisKetemu = -1;

    for (var i = 1; i < data.length; i++) {
       var sheetNip = String(data[i][2]).replace(/'/g, "").trim(); 
       
       var sheetTglRaw = data[i][3];
       var sheetTglStr = (sheetTglRaw instanceof Date) ? 
           Utilities.formatDate(sheetTglRaw, Session.getScriptTimeZone(), "dd-MM-yyyy") : String(sheetTglRaw).trim();
       
       var sheetJamRaw = data[i][4];
       var sheetJamStr = "";
       if (sheetJamRaw instanceof Date) {
         sheetJamStr = Utilities.formatDate(sheetJamRaw, Session.getScriptTimeZone(), "HH:mm");
       } else {
         sheetJamStr = String(sheetJamRaw).trim();
         if (/^\d:\d{2}/.test(sheetJamStr)) sheetJamStr = "0" + sheetJamStr;
       }

       if (sheetNip === targetNip && sheetTglStr === targetTgl && sheetJamStr === targetJam) {
          barisKetemu = i + 1; 
          break;
       }
    }

    if (barisKetemu === -1) {
      throw new Error("Data tidak ditemukan saat verifikasi. Mungkin data telah berubah.");
    }

    // 3. UPDATE DATA
    var tglVerif = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd-MM-yyyy HH:mm:ss");
    
    // --- KOREKSI MAPPING KOLOM ---
    // Kolom I (Status) -> Kolom ke-9 (Index 9)
    // Kolom J (Keterangan) -> Kolom ke-10 (Index 10)
    // Kolom M (Tgl Verif) -> Kolom ke-13 (Index 13)
    // Kolom N (Admin Verif) -> Kolom ke-14 (Index 14)
    
    sheet.getRange(barisKetemu, 9).setValue(dataKirim.status);  // FIX: Masuk Kolom I
    sheet.getRange(barisKetemu, 10).setValue(dataKirim.ket);    // Masuk Kolom J
    sheet.getRange(barisKetemu, 13).setValue(tglVerif);         // Masuk Kolom M
    sheet.getRange(barisKetemu, 14).setValue(dataKirim.admin);  // Masuk Kolom N

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

/* --- FUNGSI RESTORE DATA (VERSI FINAL: ANTI-GAGAL FORMAT) --- */
function prosesRestoreSalahAbsen(nip, tgl, jam) {
  const ID_DB = "1TZGrMiTuyvh2Xbo44RhJuWlQnOC5LzClsgIoNKtRFkY"; 
  try {
    var ss = SpreadsheetApp.openById(ID_DB);
    var sheetTrash = ss.getSheetByName("Trash");
    var sheetSource = ss.getSheetByName("Salah_Absen");

    if (!sheetTrash || !sheetSource) throw new Error("Database error.");

    var data = sheetTrash.getDataRange().getValues();
    var barisKetemu = -1;

    // --- 1. NORMALISASI TARGET (DARI CLIENT) ---
    var targetNip = String(nip).trim();
    
    // Normalisasi Tanggal: Ganti semua garis miring (/) jadi strip (-)
    // Contoh: "07/01/2026" -> "07-01-2026"
    var targetTgl = String(tgl).trim().replace(/\//g, '-');
    
    // Normalisasi Jam: "7:47:00" -> "07:47"
    var rawJamTarget = String(jam).trim();
    var targetJam = "";
    var matchJam = rawJamTarget.match(/(\d{1,2}):(\d{2})/);
    if (matchJam) {
       var h = matchJam[1].length === 1 ? "0" + matchJam[1] : matchJam[1]; // Tambah nol depan
       var m = matchJam[2];
       targetJam = h + ":" + m;
    } else {
       targetJam = rawJamTarget;
    }

    // --- 2. LOOPING PENCARIAN ---
    for (var i = 1; i < data.length; i++) {
       // A. NIP
       var sheetNip = String(data[i][2]).replace(/'/g, "").trim(); 
       
       // B. TANGGAL (SHEET)
       var sheetTglRaw = data[i][3];
       var sheetTglStr = "";
       
       if (sheetTglRaw instanceof Date) {
           // Jika format Date Object, ubah ke dd-MM-yyyy
           sheetTglStr = Utilities.formatDate(sheetTglRaw, Session.getScriptTimeZone(), "dd-MM-yyyy");
       } else {
           // Jika String, ganti / jadi - juga biar ADIL
           sheetTglStr = String(sheetTglRaw).trim().replace(/\//g, '-');
       }
       
       // C. JAM (SHEET)
       var sheetJamRaw = data[i][4];
       var sheetJamStr = "";
       if (sheetJamRaw instanceof Date) {
         sheetJamStr = Utilities.formatDate(sheetJamRaw, Session.getScriptTimeZone(), "HH:mm");
       } else {
         var txtJam = String(sheetJamRaw).trim();
         var matchSheet = txtJam.match(/(\d{1,2}):(\d{2})/);
         if (matchSheet) {
            var hS = matchSheet[1].length === 1 ? "0" + matchSheet[1] : matchSheet[1];
            var mS = matchSheet[2];
            sheetJamStr = hS + ":" + mS;
         } else {
            sheetJamStr = txtJam;
         }
       }

       // --- DEBUG (Opsional: Nyalakan jika masih gagal) ---
       // if (sheetNip === targetNip) {
       //    console.log(`Cek: Tgl[${sheetTglStr}] vs [${targetTgl}] | Jam[${sheetJamStr}] vs [${targetJam}]`);
       // }

       // D. BANDINGKAN
       if (sheetNip === targetNip && sheetTglStr === targetTgl && sheetJamStr === targetJam) {
          barisKetemu = i + 1; 
          break;
       }
    }

    if (barisKetemu === -1) {
        // Pesan error lebih detail
        throw new Error(`Data tidak ditemukan. Format Sheet: [${sheetTglStr || '?'}], Dicari: [${targetTgl}]`);
    }

    // 3. AMBIL & PINDAHKAN
    var fullRow = sheetTrash.getRange(barisKetemu, 1, 1, 17).getValues()[0];
    var restoreRow = fullRow.slice(0, 14); // Ambil 14 kolom pertama (buang data sampah)

    sheetSource.appendRow(restoreRow);
    sheetTrash.deleteRow(barisKetemu);

    return "Sukses";

  } catch (e) {
    throw new Error(e.message);
  }
}