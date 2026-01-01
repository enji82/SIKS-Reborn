/**
 * ===================================================================
 * ===================== MODUL LOGIN & OTENTIKASI ====================
 * ===================================================================
 */

// Konfigurasi Database User
const USER_DB_CONFIG = {
  id: "1wiDKez4rL5UYnpP2-OZjYowvmt1nRx-fIMy9trJlhBA", // ID Spreadsheet Anda
  sheet: "DATA_USERS" // Nama Sheet
};

/**
 * Fungsi Utama: Cek Username & Password
 */
function checkLoginCredentials(form) {
  try {
    // 1. Validasi Input Dasar
    if (!form.username || !form.password) {
      return { success: false, message: "Username dan Password wajib diisi." };
    }

    // 2. Buka Spreadsheet Database
    const ss = SpreadsheetApp.openById(USER_DB_CONFIG.id);
    const sheet = ss.getSheetByName(USER_DB_CONFIG.sheet);
    
    if (!sheet) {
      throw new Error(`Database Error: Sheet '${USER_DB_CONFIG.sheet}' tidak ditemukan. Hubungi Admin.`);
    }

    // 3. Ambil Semua Data User (Cache-friendly)
    // Ambil semua data (A:E) untuk meminimalkan panggilan server
    const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 5).getValues(); 

    // 4. Cari Kecocokan (Case Insensitive untuk Username)
    const inputUser = String(form.username).trim().toLowerCase();
    const inputPass = String(form.password).trim();
    
    let userFound = null;

    for (let i = 0; i < data.length; i++) {
      const row = data[i];
      const dbUser = String(row[0]).trim().toLowerCase(); // Kolom A: Username
      const dbPass = String(row[1]).trim(); // Kolom B: Password
      
      // Cek kecocokan
      if (dbUser === inputUser && dbPass === inputPass) {
        userFound = {
          username: row[0],
          nama: row[2], // Kolom C: Nama Lengkap
          role: row[3], // Kolom D: Role
          unit: row[4]  // Kolom E: Unit Kerja (Opsional)
        };
        break; // Stop looping jika ketemu
      }
    }

    // 5. Kembalikan Hasil
    if (userFound) {
      return { 
        success: true, 
        message: "Login Berhasil", 
        user: userFound 
      };
    } else {
      return { 
        success: false, 
        message: "Username atau Password salah!" 
      };
    }

  } catch (e) {
    Logger.log("Login Error: " + e.message);
    return { success: false, message: "Terjadi kesalahan sistem: " + e.message };
  }
}