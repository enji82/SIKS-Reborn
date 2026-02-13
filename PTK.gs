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
  try {
    var ss = SpreadsheetApp.openById(ID_DB_PTK);
    var sheet = ss.getSheetByName(SHEET_PTK);
    
    // Generate ID
    var newId = "";
    var nipClean = String(form.nip).replace(/[^0-9]/g, "");
    var nikClean = String(form.nik).replace(/[^0-9]/g, "");
    if (nipClean.length === 18) newId = nipClean;
    else if (nikClean.length === 16) newId = nikClean;
    else newId = "PTK-" + new Date().getTime();

    // Gabung Nama
    var gelarDepan = (form.gelar_depan && form.gelar_depan !== "-") ? form.gelar_depan + " " : "";
    var gelarBelakang = (form.gelar_belakang && form.gelar_belakang !== "-") ? ", " + form.gelar_belakang : "";
    var namaGabungan = gelarDepan + form.nama_lengkap.toUpperCase() + gelarBelakang;

    // Lookup Kelas Jabatan
    var kelasJabatan = "-";
    var refJabatan = ss.getSheetByName("data_jabatan").getDataRange().getValues();
    for (var i = 1; i < refJabatan.length; i++) {
        if (String(refJabatan[i][0]).trim() === String(form.jabatan).trim()) {
            kelasJabatan = refJabatan[i][1];
            break;
        }
    }

    // Format MKG
    var mkgStr = "";
    if (form.mkg_thn || form.mkg_bln) {
        var thn = form.mkg_thn || "00";
        var bln = form.mkg_bln || "00";
        mkgStr = thn + " Tahun " + bln + " Bulan";
    }

    var now = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd-MM-yyyy HH:mm:ss");
    var user = form.user_login || "Admin";

    var rowData = [
        newId,                  // A
        "'" + form.npsn_login,  // B
        form.unit_login,        // C
        form.gelar_depan,       // D
        form.nama_lengkap.toUpperCase(), // E
        form.gelar_belakang,    // F
        namaGabungan,           // G
        "'" + form.nip,         // H
        form.tmp_lahir,         // I
        "'" + form.tgl_lahir,   // J
        "'" + form.nik,         // K
        form.lp,                // L
        form.agama,             // M
        form.pendidikan,        // N
        form.jurusan,           // O
        form.thn_lulus,         // P
        form.alamat,            // Q
        "'" + form.hp,          // R
        form.status_peg,        // S
        form.jabatan,           // T
        "'" + form.tmt_jabatan, // U
        form.pangkat,           // V
        "'" + form.tmt_gol,     // W
        mkgStr,                 // X: MKG
        kelasJabatan,           // Y: Kelas Jab.
        form.tugas,             // Z: Tugas
        "'" + form.nuptk,       // AA
        form.serdik,            // AB
        form.dapodik,           // AC
        form.tugtam,            // AD
        form.jabatan_guru_pegawai, // AE: JABATAN GURU (Pengganti Keaktifan)
        now,                    // AF
        user,                   // AG
        "",                     // AH
        ""                      // AI
    ];

    sheet.appendRow(rowData);
    return "Sukses";
    
  } catch (e) { return "Error: " + e.message; }
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