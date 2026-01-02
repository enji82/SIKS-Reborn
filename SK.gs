function getSkArsipFolderIds() {
  try {
    return {
      'MAIN_SK': FOLDER_CONFIG.MAIN_SK
    };
  } catch (e) {
    return handleError('getSkArsipFolderIds', e);
  }
}

function getMasterSkOptions() {
  // Kunci cache unik
  const cacheKey = 'master_sk_options_v1';
  
  // Gunakan fungsi cache yang sudah ada
  return getCachedData(cacheKey, function() {
    try {
      const ss = SpreadsheetApp.openById(SPREADSHEET_CONFIG.DROPDOWN_DATA.id);
      const getValuesFromSheet = (sheetName) => {
        const sheet = ss.getSheetByName(sheetName);
        if (!sheet) return [];
        return sheet.getRange('A2:A' + sheet.getLastRow()).getValues()
                    .flat()
                    .filter(value => String(value).trim() !== '');
      };

      return {
        'Nama SD': getValuesFromSheet('Nama SD').sort(),
        'Tahun Ajaran': getValuesFromSheet('Tahun Ajaran').sort().reverse(),
        'Semester': getValuesFromSheet('Semester').sort(),
        'Kriteria SK': getValuesFromSheet('Kriteria SK').sort()
      };
    } catch (e) {
      // Saat caching, kita lempar error agar tidak menyimpan cache yang rusak
      throw new Error(`Gagal mengambil SK Options: ${e.message}`);
    }
  });
}

function processManualForm(formData) {
  try {
    Logger.log("--- [DEBUG] MULAI UPLOAD SK ---");
    Logger.log("User Input yang diterima dari Form: " + formData.userInput);

    const targetSheetName = "Unggah_SK"; 
    const config = SPREADSHEET_CONFIG.SK_FORM_RESPONSES; 
    const ss = SpreadsheetApp.openById(config.id);
    const sheet = ss.getSheetByName(targetSheetName);

    if (!sheet) throw new Error(`Sheet "${targetSheetName}" tidak ditemukan.`);

    const mainFolder = DriveApp.getFolderById(FOLDER_CONFIG.MAIN_SK);  
    const tahunAjaranFolderName = formData.tahunAjaran.replace(/\//g, '-');
    const tahunAjaranFolder = getOrCreateFolder(mainFolder, tahunAjaranFolderName);
    const semesterFolderName = formData.semester;
    const targetFolder = getOrCreateFolder(tahunAjaranFolder, semesterFolderName);

    const newFilename = `${formData.namaSD} - ${tahunAjaranFolderName} - ${formData.semester} - ${formData.kriteriaSK}.pdf`;
    
    const decodedData = Utilities.base64Decode(formData.fileData.data);
    const blob = Utilities.newBlob(decodedData, formData.fileData.mimeType, newFilename);
    const newFile = targetFolder.createFile(blob);
    const fileUrl = newFile.getUrl();
    
    // LOGIKA PENYIMPANAN
    const newRow = [ 
      new Date(),                   
      formData.namaSD,              
      formData.tahunAjaran,         
      formData.semester,            
      formData.nomorSK,             
      new Date(formData.tanggalSK), 
      formData.kriteriaSK,          
      fileUrl,                      
      formData.userInput,           // <--- Titik Kritis 1
      "Diproses"                    
    ];
    
    Logger.log("Data yang akan disimpan ke Sheet: " + JSON.stringify(newRow));
    sheet.appendRow(newRow);
    
    const lastRow = sheet.getLastRow();
    sheet.getRange(lastRow, 6).setNumberFormat("dd-MM-yyyy");

    Logger.log("--- [DEBUG] SELESAI UPLOAD ---");
    return "Dokumen SK berhasil diunggah dan status 'Diproses'.";
  } catch (e) {
    Logger.log("ERROR processManualForm: " + e.message);
    return handleError('processManualForm', e);
  }
}

function getSKRiwayatData() {
  try {
    // === KONFIGURASI ===
    // Pastikan SPREADSHEET_CONFIG sudah ada di file Config.gs atau Code.gs
    // Jika belum ada, ganti baris di bawah dengan ID manual: 
    // const sheetId = "ID_SPREADSHEET_ANDA_DISINI";
    
    const config = SPREADSHEET_CONFIG.SK_FORM_RESPONSES; 
    const targetSheetName = "Unggah_SK"; 
    
    const ss = SpreadsheetApp.openById(config.id);
    const sheet = ss.getSheetByName(targetSheetName);

    if (!sheet) {
      return { headers: [], rows: [] }; 
    }

    const allData = sheet.getDataRange().getValues();
    if (allData.length < 2) return { headers: [], rows: [] };

    // Mapping Header (Lowercase)
    const originalHeaders = allData[0].map(h => String(h).trim().toLowerCase());
    const headerMap = {};
    originalHeaders.forEach((h, index) => { headerMap[h] = index; });

    const dataRows = allData.slice(1);
    
    // Sort Data (Terbaru di atas)
    const timestampIndex = headerMap['tanggal unggah'];
    if (timestampIndex !== undefined) {
        dataRows.sort((a, b) => {
            const dateA = a[timestampIndex] instanceof Date ? a[timestampIndex].getTime() : 0;
            const dateB = b[timestampIndex] instanceof Date ? b[timestampIndex].getTime() : 0;
            return dateB - dateA; 
        });
    }

    // Ambil Data
    let structuredRows = dataRows.map(row => {
      const getVal = (key) => {
         const idx = headerMap[key];
         return (idx !== undefined) ? row[idx] : null;
      };

      const rowObj = {};
      rowObj['Nama SD']      = getVal('nama sd') || '-';
      rowObj['Tahun Ajaran'] = getVal('tahun ajaran') || '-';
      rowObj['Semester']     = getVal('semester') || '-';
      rowObj['Nomor SK']     = getVal('nomor sk') || '-';
      rowObj['Kriteria SK']  = getVal('kriteria sk') || '-';
      
      // Handle Link Dokumen
      const dokVal = getVal('link dokumen') || getVal('dokumen');
      rowObj['Dokumen'] = dokVal || '#';

      // Handle User Input
      const userVal = getVal('user input') || getVal('userinput');
      rowObj['User Input'] = userVal || '-';

      // Handle Status
      rowObj['Status'] = getVal('status') || 'Diproses';

      // Format Tanggal SK
      const tglSK = getVal('tanggal sk');
      rowObj['Tanggal SK'] = (tglSK instanceof Date) ? 
          Utilities.formatDate(tglSK, Session.getScriptTimeZone(), "dd/MM/yyyy") : (tglSK || '');

      return rowObj;
    });

    return { headers: [], rows: structuredRows };
  } catch (e) {
    throw new Error("Backend Error (getSKRiwayatData): " + e.message);
  }
}

function getSKStatusData() {
  try {
    const config = SPREADSHEET_CONFIG.SK_FORM_RESPONSES; 
    const targetSheetName = "Status_SK"; 
    
    const ss = SpreadsheetApp.openById(config.id);
    const sheet = ss.getSheetByName(targetSheetName);

    if (!sheet) return { headers: [], rows: [] }; 

    // Ambil Semua Data
    const data = sheet.getDataRange().getDisplayValues(); // getDisplayValues agar tanggal/angka sesuai tampilan sheet
    if (data.length < 2) return { headers: [], rows: [] };

    // BARIS 1: Header
    const headers = data[0]; // ["No", "Nama Sekolah", "2021 Ganjil", "2021 Genap", ...]

    // BARIS 2 dst: Isi Data
    const rows = data.slice(1);

    return { 
      headers: headers,
      rows: rows 
    };
    
  } catch (e) {
    throw new Error("Gagal mengambil data Status: " + e.message);
  }
}

function getArsipData(folderId) {
  try {
    // 1. AMBIL ID DARI CONFIG TERPUSAT (Code.gs)
    // Pastikan FOLDER_CONFIG.MAIN_SK sudah didefinisikan di Code.gs
    const rootId = FOLDER_CONFIG.MAIN_SK; 
    
    if (!rootId) {
      throw new Error("ID Folder belum diset di FOLDER_CONFIG.MAIN_SK (Code.gs)");
    }
    
    // Jika frontend tidak kirim ID (saat pertama buka), pakai Root ID dari Config
    const targetId = folderId || rootId; 
    
    const folder = DriveApp.getFolderById(targetId);
    if (!folder) throw new Error("Folder tidak ditemukan di Google Drive");

    let items = [];

    // 2. AMBIL FOLDER (Sub-folder)
    const subFolders = folder.getFolders();
    while (subFolders.hasNext()) {
      let f = subFolders.next();
      items.push({
        id: f.getId(),
        name: f.getName(),
        type: 'folder',
        mimeType: 'application/vnd.google-apps.folder',
        date: f.getLastUpdated(),
        size: '-'
      });
    }

    // 3. AMBIL FILE
    const files = folder.getFiles();
    while (files.hasNext()) {
      let f = files.next();
      // Format ukuran file (KB/MB)
      let sizeBytes = f.getSize();
      let sizeStr = (sizeBytes / 1024).toFixed(1) + " KB";
      if (sizeBytes > 1024 * 1024) sizeStr = (sizeBytes / (1024 * 1024)).toFixed(1) + " MB";

      items.push({
        id: f.getId(),
        name: f.getName(),
        type: 'file',
        mimeType: f.getMimeType(),
        url: f.getUrl(), // Link untuk buka file
        date: f.getLastUpdated(),
        size: sizeStr
      });
    }

    // 4. SORTING: Folder di atas, File di bawah. Lalu urut abjad.
    items.sort((a, b) => {
      if (a.type === b.type) return a.name.localeCompare(b.name);
      return a.type === 'folder' ? -1 : 1;
    });

    // 5. BREADCRUMB (Jalur Navigasi)
    let parentId = null;
    // Cek apakah kita sedang berada di dalam sub-folder (bukan di root)
    if (targetId !== rootId) {
      const parents = folder.getParents();
      if (parents.hasNext()) parentId = parents.next().getId();
    }

    return {
      currentId: targetId,
      currentName: folder.getName(),
      parentId: parentId,
      isRoot: (targetId === rootId),
      items: items.map(i => ({
         ...i,
         date: Utilities.formatDate(i.date, Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm")
      }))
    };

  } catch (e) {
    throw new Error("Gagal akses Drive: " + e.message);
  }
}

function getSKKelolaData() {
  try {
    const config = SPREADSHEET_CONFIG.SK_FORM_RESPONSES;
    const sheet = SpreadsheetApp.openById(config.id).getSheetByName(config.sheet);
    if (!sheet || sheet.getLastRow() < 2) {
      return { headers: [], rows: [] };
    }

    const originalData = sheet.getDataRange().getValues();
    const originalHeaders = originalData[0].map(h => String(h).trim());
    const dataRows = originalData.slice(1);
    
    const parseDate = (value) => value instanceof Date && !isNaN(value) ? value : new Date(0);

    const indexedData = dataRows.map((row, index) => ({
      row: row,
      originalIndex: index + 2
    }));
    
    const updateIndex = originalHeaders.indexOf('Update');
    const timestampIndex = originalHeaders.indexOf('Tanggal Unggah');

    indexedData.sort((a, b) => {
      const dateB_update = parseDate(b.row[updateIndex]);
      const dateA_update = parseDate(a.row[updateIndex]);
      if (dateB_update.getTime() !== dateA_update.getTime()) {
        return dateB_update - dateA_update;
      }
      const dateB_timestamp = parseDate(b.row[timestampIndex]);
      const dateA_timestamp = parseDate(a.row[timestampIndex]);
      return dateB_timestamp - dateA_timestamp;
    });

    const structuredRows = indexedData.map(item => {
      const rowObject = {
        _rowIndex: item.originalIndex,
        _source: 'SK'
      };
      originalHeaders.forEach((header, i) => {
      let cell = item.row[i];
      // MODIFIKASI DIMULAI DI SINI
      if (header === 'Tanggal SK' && cell instanceof Date) {
      rowObject[header] = Utilities.formatDate(cell, Session.getScriptTimeZone(), "dd/MM/yyyy");
      } else if ((header === 'Tanggal Unggah' || header === 'Update') && cell instanceof Date) {
      rowObject[header] = Utilities.formatDate(cell, Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm:ss");
      // MODIFIKASI SELESAI
      } else {
      rowObject[header] = cell;
    }
  });
  return rowObject;
    });
    
    const desiredHeaders = ["Nama SD", "Tahun Ajaran", "Semester", "Nomor SK", "Kriteria SK", "Dokumen", "Aksi", "Tanggal Unggah", "Update"];

    return {
      headers: desiredHeaders,
      rows: structuredRows
    };
  } catch (e) {
    return handleError("getSKKelolaData", e);
  }
}

function getSKDataByRow(rowIndex) {
  try {
    const config = SPREADSHEET_CONFIG.SK_FORM_RESPONSES;
    const sheet = SpreadsheetApp.openById(config.id).getSheetByName(config.sheet);
    
    // Ambil nilai mentah (RAW) untuk mendapatkan objek Date asli
    const rawValues = sheet.getRange(rowIndex, 1, 1, sheet.getLastColumn()).getValues()[0];
    // Ambil nilai tampilan (DISPLAY) untuk konsistensi string/angka
    const displayValues = sheet.getRange(rowIndex, 1, 1, sheet.getLastColumn()).getDisplayValues()[0];
    
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map(h => h.trim());
    
    const rowData = {};
    headers.forEach((header, i) => {
      // KUNCI PERBAIKAN: Format Tanggal SK ke YYYY-MM-DD
      if (header === 'Tanggal SK' && rawValues[i] instanceof Date) {
        // Format yang wajib untuk HTML input type="date"
        rowData[header] = Utilities.formatDate(rawValues[i], "UTC", "yyyy-MM-dd");
      } else {
        // Gunakan display value untuk field lain (Nomor SK, dll.)
        rowData[header] = displayValues[i];
      }
    });
    return rowData;
  } catch (e) {
    return handleError("getSKDataByRow", e);
  }
}

function updateSKData(formData) {
  try {
    const config = SPREADSHEET_CONFIG.SK_FORM_RESPONSES;
    const sheet = SpreadsheetApp.openById(config.id).getSheetByName(config.sheet);
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map(h => h.trim());
    
    const range = sheet.getRange(formData.rowIndex, 1, 1, headers.length);
    const existingRowValues = range.getDisplayValues()[0];
    const existingRowObject = {};
    headers.forEach((header, i) => { existingRowObject[header] = existingRowValues[i]; });

    const mainFolder = DriveApp.getFolderById(FOLDER_CONFIG.MAIN_SK);
    const tahunAjaranFolderName = existingRowObject['Tahun Ajaran'].replace(/\//g, '-');
    const tahunAjaranFolder = getOrCreateFolder(mainFolder, tahunAjaranFolderName);
    
    let fileUrl = existingRowObject['Dokumen'];
    const fileUrlIndex = headers.indexOf('Dokumen');

    const newSemesterFolderName = formData['Semester'];
    const newTargetFolder = getOrCreateFolder(tahunAjaranFolder, newSemesterFolderName);
    const newFilename = `${existingRowObject['Nama SD']} - ${tahunAjaranFolderName} - ${newSemesterFolderName} - ${formData['Kriteria SK']}.pdf`;

    if (formData.fileData && formData.fileData.data) {
      if (fileUrlIndex > -1 && existingRowObject['Dokumen']) {
        try {
          const fileId = existingRowObject['Dokumen'].match(/[-\w]{25,}/);
          if (fileId) DriveApp.getFileById(fileId[0]).setTrashed(true);
        } catch (e) {
          Logger.log(`Gagal menghapus file lama saat upload baru: ${e.message}`);
        }
      }
      
      const decodedData = Utilities.base64Decode(formData.fileData.data);
      const blob = Utilities.newBlob(decodedData, formData.fileData.mimeType, newFilename);
      const newFile = newTargetFolder.createFile(blob);
      fileUrl = newFile.getUrl();

    } else if (fileUrlIndex > -1 && existingRowObject['Dokumen']) {
        const fileIdMatch = existingRowObject['Dokumen'].match(/[-\w]{25,}/);
        if (fileIdMatch) {
            const fileId = fileIdMatch[0];
            const file = DriveApp.getFileById(fileId);
            const currentFileName = file.getName();
            const currentParentFolder = file.getParents().next();

            if (currentFileName !== newFilename || currentParentFolder.getName() !== newSemesterFolderName) {
                file.moveTo(newTargetFolder);
                file.setName(newFilename);
                fileUrl = file.getUrl();
            }
        }
    }
    
    formData['Dokumen'] = fileUrl;
    formData['Update'] = new Date();

    const newRowValuesForSheet = headers.map(header => {
      return formData.hasOwnProperty(header) ? formData[header] : existingRowObject[header];
    });

    sheet.getRange(formData.rowIndex, 1, 1, headers.length).setValues([newRowValuesForSheet]);
    
    const tanggalSKIndex = headers.indexOf('Tanggal SK');
    if (tanggalSKIndex !== -1) {
      sheet.getRange(formData.rowIndex, tanggalSKIndex + 1).setNumberFormat("dd-MM-yyyy");
    }
    
    return "Data berhasil diperbarui!";
  } catch (e) {
    return handleError('updateSKData', e);
  }
}

function deleteSKData(rowIndex, deleteCode) {
  try {
    const todayCode = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyyMMdd");
    if (String(deleteCode).trim() !== todayCode) throw new Error("Kode Hapus salah.");

    const config = SPREADSHEET_CONFIG.SK_FORM_RESPONSES;
    const sheet = SpreadsheetApp.openById(config.id).getSheetByName(config.sheet);
    
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const fileUrlIndex = headers.findIndex(h => h.trim().toLowerCase() === 'dokumen');
    
    if (fileUrlIndex !== -1) {
        const fileUrl = sheet.getRange(rowIndex, fileUrlIndex + 1).getValue();
        if (fileUrl && typeof fileUrl === 'string') {
            const fileId = fileUrl.match(/[-\w]{25,}/);
            if (fileId) {
                try {
                    DriveApp.getFileById(fileId[0]).setTrashed(true);
                } catch (err) {
                    Logger.log(`Gagal menghapus file dengan ID ${fileId[0]}: ${err.message}`);
                }
            }
        }
    }
    
    sheet.deleteRow(rowIndex);
    return "Data dan file terkait berhasil dihapus.";
  } catch (e) {
    return handleError("deleteSKData", e);
  }
}