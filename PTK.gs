/* ======================================================================
   PTK.GS - KHUSUS PENGELOLAAN DATA PTK SD
   Sheet: Master Data GTK | ID: 1t0-Lmy0YD_GxHzimFWJGh5R5x6RhGL13uqKeVwWoCYE
   ====================================================================== */

const CONF_PTK_SD = {
  ID: "1t0-Lmy0YD_GxHzimFWJGh5R5x6RhGL13uqKeVwWoCYE", 
  SHEET: "Master Data GTK"
};

/* 1. AMBIL DATA (READ) */
function getPTK_SD() {
  try {
    var ss = SpreadsheetApp.openById(CONF_PTK_SD.ID);
    var sheet = ss.getSheetByName(CONF_PTK_SD.SHEET);
    if (!sheet) return [];

    var lastRow = sheet.getLastRow();
    if (lastRow < 2) return [];

    // Ambil Data A2:X (24 Kolom)
    var data = sheet.getRange(2, 1, lastRow - 1, 24).getValues();
    
    // FORMAT DATA AGAR AMAN DIKIRIM KE JSON
    // Tanggal (Date Object) harus jadi String agar tidak error di client
    var result = data.map(function(row) {
      return {
        id: row[0],             // A
        npsn: row[1],           // B
        unit: row[2],           // C
        nama: row[3],           // D
        nip: row[4],            // E
        jabatan: row[5],        // F
        gol: row[6],            // G
        tmp_lahir: row[7],      // H
        tgl_lahir: formatDate(row[8]), // I
        pendidikan: row[9],     // J
        jurusan: row[10],       // K
        status_peg: row[11],    // L
        tugas: row[12],         // M
        serdik: row[13],        // N
        dapodik: row[14],       // O
        tmt_nonasn: formatDate(row[15]), // P
        keaktifan: row[16],     // Q
        file_awal: row[17],     // R
        file_akhir: row[18],    // S
        jenis_sk: row[19],      // T
        tgl_input: formatDateTime(row[20]), // U
        user_input: row[21],    // V
        tgl_edit: formatDateTime(row[22]),  // W
        user_edit: row[23]      // X
      };
    });

    return JSON.stringify(result); // Kirim sebagai String JSON

  } catch (e) {
    return JSON.stringify({ error: e.message });
  }
}

/* 2. SIMPAN DATA (CREATE / UPDATE) */
function savePTK_SD(form) {
  try {
    var ss = SpreadsheetApp.openById(CONF_PTK_SD.ID);
    var sheet = ss.getSheetByName(CONF_PTK_SD.SHEET);
    
    if (!sheet) {
      sheet = ss.insertSheet(CONF_PTK_SD.SHEET);
      // Header A-X
      sheet.appendRow(["ID","NPSN","Unit Kerja","Nama","NIP","Jabatan","Pangkat/Gol","Tempat Lahir","Tanggal Lahir","Pendidikan","Jurusan","Kepegawaian","Tugas","Serdik","Dapodik","TMT Non ASN","Keaktifan","File Awal","File Akhir","Jenis SK Akhir","Tgl Input","User Input","Tgl Edit","User Edit"]);
    }

    var now = new Date();
    var user = form.user || "Admin";
    var tglLahir = form.tgl_lahir ? new Date(form.tgl_lahir) : "";
    var tmt = form.tmt_nonasn ? new Date(form.tmt_nonasn) : "";

    // Cek Mode Edit
    if (form.id && form.id !== "") {
      var data = sheet.getDataRange().getValues();
      var rowIndex = -1;
      
      for (var i = 1; i < data.length; i++) {
        if (String(data[i][0]) === String(form.id)) { rowIndex = i + 1; break; }
      }

      if (rowIndex === -1) return "Gagal: ID tidak ditemukan.";

      // Update Kolom B-T (Index 2-20)
      var rowData = [
        form.npsn, form.unit, form.nama, "'"+form.nip, form.jabatan, form.gol,
        form.tmp_lahir, tglLahir, form.pendidikan, form.jurusan, form.status_peg,
        form.tugas, form.serdik, form.dapodik, tmt, form.keaktifan,
        form.file_awal, form.file_akhir, form.jenis_sk_akhir
      ];
      sheet.getRange(rowIndex, 2, 1, 19).setValues([rowData]);
      
      // Update Log (W-X)
      sheet.getRange(rowIndex, 23).setValue(now);
      sheet.getRange(rowIndex, 24).setValue(user);

      return "Sukses: Data berhasil diperbarui.";

    } else {
      // Mode Baru
      var newId = Utilities.getUuid();
      var newRow = [
        newId, form.npsn, form.unit, form.nama, "'"+form.nip, form.jabatan, form.gol,
        form.tmp_lahir, tglLahir, form.pendidikan, form.jurusan, form.status_peg,
        form.tugas, form.serdik, form.dapodik, tmt, form.keaktifan,
        form.file_awal, form.file_akhir, form.jenis_sk_akhir,
        now, user, "", ""
      ];
      sheet.appendRow(newRow);
      return "Sukses: Data baru tersimpan.";
    }

  } catch (e) {
    return "Error: " + e.message;
  }
}

/* 3. HAPUS DATA */
function deletePTK_SD(id) {
  try {
    var ss = SpreadsheetApp.openById(CONF_PTK_SD.ID);
    var sheet = ss.getSheetByName(CONF_PTK_SD.SHEET);
    var data = sheet.getDataRange().getValues();
    
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][0]) === String(id)) {
        sheet.deleteRow(i + 1);
        return "Sukses: Data dihapus.";
      }
    }
    return "Gagal: ID tidak ditemukan.";
  } catch (e) {
    return "Error: " + e.message;
  }
}

/* HELPER DATE FORMATTER */
function formatDate(dateObj) {
  if (!dateObj || !(dateObj instanceof Date)) return "";
  return Utilities.formatDate(dateObj, Session.getScriptTimeZone(), "yyyy-MM-dd");
}
function formatDateTime(dateObj) {
  if (!dateObj || !(dateObj instanceof Date)) return "";
  return Utilities.formatDate(dateObj, Session.getScriptTimeZone(), "dd-MM-yyyy HH:mm");
}