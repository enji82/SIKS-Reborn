# 🔧 PATCH LOG - Bug Fixes SIKS-Reborn

**Date:** March 22, 2026  
**Status:** ✅ COMPLETED - 6 Critical & High Priority Bugs Fixed

---

## 📋 Summary of Patches Applied

### ✅ **PATCH #1: Password Security (CRITICAL)**
**File:** `code.gs`  
**Bug:** Plain text password storage  
**Fix Applied:**
- ✅ Added `hashPassword()` function using SHA-256 encryption
- ✅ Added `hashPasswordBase64()` for secure storage in Sheets
- ✅ Updated `processLogin()` to use hashed password comparison
- ✅ Added `migrateHashPasswords()` helper function to bulk-hash existing passwords

**Before:**
```javascript
if (String(row[1]).trim() == inputPass) {  // Plain text ❌
```

**After:**
```javascript
var inputPassHash = hashPasswordBase64(inputPass);
if (storedPassHash === inputPassHash) {  // Hashed ✅
```

**ACTION REQUIRED:**
1. Go to Apps Script: Tools → Script Editor
2. Run `migrateHashPasswords()` function (Execute)
3. This will hash all existing passwords in Sheets (one-time only)
4. Verify in Sheets that column B passwords are now long hashed strings

---

### ✅ **PATCH #2: Brute Force Protection (HIGH)**
**File:** `code.gs`  
**Bug:** No rate limiting on login attempts  
**Fix Applied:**
- ✅ Added `checkRateLimit()` function
- ✅ Added `recordFailedLogin()` function with 5-strike lockout
- ✅ Added `clearFailedLogin()` function for successful login
- ✅ 15-minute account lockout after 5 failed attempts
- ✅ Updated `processLogin()` to enforce rate limiting

**Behavior:**
- User tries login 5 times with wrong password → Account locked for 15 minutes
- After lockout expires, counter resets automatically
- Uses Google Apps Script Properties Service (no database needed)

---

### ✅ **PATCH #3: Duplikat Kolom Unit (HIGH)**
**File:** `code.gs`  
**Bug:** `photo` dan `unit` ambil dari kolom yang sama  
**Fix Applied:**
- ✅ Changed `unit` from `row[4]` to correct mapping
- ✅ Changed `photo` from `row[4]` to `row[5]`
- ✅ Updated spreadsheet column mapping: A=Username, B=Password, C=Nama, D=Role, E=Unit, F=Photo

**Before:**
```javascript
photo: row[4] || "",  // ❌ Kolom E
unit: row[4] || "",   // ❌ DUPLIKAT!
```

**After:**
```javascript
unit: String(row[4] || "").trim(),    // ✅ Kolom E (Unit)
photo: String(row[5] || "").trim(),   // ✅ Kolom F (Photo)
```

---

### ✅ **PATCH #4: XSS Injection Prevention (CRITICAL)**
**File:** `code.gs`  
**Bug:** Unescaped HTML in error message allows JavaScript injection  
**Fix Applied:**
- ✅ Added `HtmlService.htmlEscape()` on user input in `getHalaman()`
- ✅ Prevents malicious script injection via page name parameter

**Before:**
```javascript
return '<div class="alert"><b>' + namaFile + '</b> belum dibuat</div>';
// Attack: getHalaman('<img src=x onerror="alert(1)">')
// Result: XSS WORKS! ❌
```

**After:**
```javascript
var safeFileName = HtmlService.htmlEscape(namaFile);  // ✅ Sanitized
return '<div class="alert"><b>' + safeFileName + '</b> belum dibuat</div>';
// Attack attempt now displays as plain text, not executed
```

---

### ✅ **PATCH #5: Null Sheet Reference Validation (HIGH)**
**File:** `SK.gs`  
**Bug:** No validation when `getSheetByName()` returns null  
**Fix Applied:**
- ✅ Added null checks for spreadsheet and sheet before operations
- ✅ Returns meaningful error messages instead of crashing
- ✅ Updated `processManualForm()` with comprehensive validation
- ✅ Updated `getDaftarSK()` with null checks

**Before:**
```javascript
const sheet = ss.getSheetByName("Unggah_SK");
sheet.appendRow([...]);  // CRASH if sheet is null ❌
```

**After:**
```javascript
const sheet = ss.getSheetByName("Unggah_SK");
if (!sheet) return { success: false, message: 'Sheet "Unggah_SK" tidak ditemukan.' };
sheet.appendRow([...]);  // ✅ Safe
```

---

### ✅ **PATCH #6: Race Condition Fix & Performance Optimization (MEDIUM)**
**File:** `SK.gs` & `code.gs`  
**Bugs:** 
1. Race condition on visitor counter (multiple concurrent users)
2. `getDataRange()` timeout on large datasets

**Fixes Applied:**

#### a) Visitor Counter Race Condition:
- ✅ Added `LockService.getScriptLock()` for atomic operations
- ✅ Prevents concurrent read-modify-write conflicts
- ✅ 5-second wait timeout with always-release guarantee

**Before:**
```javascript
var totalHits = Number(props.getProperty('TOTAL_HITS')) || 0;
totalHits++;  // ❌ Race condition if 2 users access same time
props.setProperty('TOTAL_HITS', totalHits.toString());
```

**After:**
```javascript
var lock = LockService.getScriptLock();
lock.waitLock(5000);  // ✅ Atomic
try {
  var totalHits = Number(props.getProperty('TOTAL_HITS')) || 0;
  totalHits++;
  props.setProperty('TOTAL_HITS', totalHits.toString());
} finally {
  lock.releaseLock();  // Always release
}
```

#### b) Large Dataset Timeout:
- ✅ Changed `getDataRange()` to `getRange()` with batch limit
- ✅ Only fetch last 500 rows (prevents 30-second timeout on large sheets)
- ✅ Added `getLastRow()` check for data existence
- ✅ Simplified date parsing with native `Date()` constructor

**Before:**
```javascript
var data = sheet.getDataRange().getDisplayValues();  // ❌ Ambil SEMUA!
// Jika 10,000 rows → timeout error
```

**After:**
```javascript
var lastRow = sheet.getLastRow();
if (lastRow < 2) return [];  // Early exit
var startRow = Math.max(2, lastRow - 499);  // ✅ Hanya 500 baris terakhir
var data = sheet.getRange(startRow, 1, numRows, sheet.getLastColumn()).getDisplayValues();
```

---

## 📊 Impact Summary

| Bug | Severity | Type | Fixed |
|-----|----------|------|-------|
| Plain Text Password | 🔴 CRITICAL | Security | ✅ Hashed with SHA-256 |
| XSS Injection | 🔴 CRITICAL | Security | ✅ HTML escaped |
| Duplikat Kolom | 🟠 HIGH | Logic | ✅ Corrected mapping |
| Null References | 🟠 HIGH | Runtime | ✅ Validation added |
| Brute Force | 🟡 MEDIUM | Security | ✅ Rate limiting |
| Race Condition | 🟡 MEDIUM | Concurrency | ✅ Lock service |
| Dataset Timeout | 🟡 MEDIUM | Performance | ✅ Pagination |

---

## ⚙️ DEPLOYMENT INSTRUCTIONS

### Step 1: Backup (IMPORTANT!)
```
1. Download backup of all Google Sheets (Download as .xlsx)
2. Take screenshot of current DATA_USER sheet structure
3. Note current password values
```

### Step 2: Update Code
```
1. Go to Apps Script: Script Editor
2. Replace code.gs content with patched version
3. Replace SK.gs content with patched version
4. Click Save (Ctrl+S)
```

### Step 3: Verify Functions
```
1. Run → migrateHashPasswords()
   - Should show "X passwords berhasil di-hash"
   - Check Sheets column B - passwords should be ~88 character hashes
2. If error: Revert to backup and check sheet column names
```

### Step 4: Test Login
```
1. Open deployment script URL (App → New Deployment → Web App)
2. Test login with existing credentials
   - Should work after migration
   - If fails: Check if passwords were hashed correctly
```

### Step 5: Clean Up
```
1. Delete migrateHashPasswords() function (optional - can keep as helper)
2. Verify no console errors in deployment
```

---

## ⚠️ IMPORTANT NOTES

### Password Migration
- **Before deploying to production**, run `migrateHashPasswords()` once in Script Editor
- This converts all existing plain-text passwords to SHA-256 hashes
- After migration, passwords in Sheets will be unreadable (correct behavior!)
- Users can still login normally - the app hashes their input for comparison

### New Password Creation
- If admins need to add new user:
  1. Get plain text password from user
  2. In Script Editor, run: `hashPasswordBase64("password123")`
  3. Copy hash result into Sheets column B

### Spreadsheet Structure Check
**Required columns in DATA_USER sheet:**
| Col | Name | Type | Example |
|-----|------|------|---------|
| A | Username | String | admin |
| B | Password | String | `[88-char hash]` |
| C | Nama Lengkap | String | Admin Sistem |
| D | Role | String | Administrator |
| E | Unit | String | Korwil Secang |
| F | Foto | String | https://... |

---

## 🧪 Testing Checklist

- [ ] `migrateHashPasswords()` successfully runs without error
- [ ] Passwords in Sheets column B are now ~88 character hashes
- [ ] Login works with correct username/password
- [ ] Login fails with wrong password
- [ ] After 5 failed attempts, account locked for 15 minutes
- [ ] Error messages no longer show raw HTML
- [ ] SK data loads without timeout (even if 1000+ rows)
- [ ] Visitor counter accurate (test with multiple simultaneous users if possible)

---

## 📞 Troubleshooting

### Issue: "Cannot read property 'getContentText' of null"
**Cause:** Sheet not found  
**Fix:** Check sheet name in SPREADSHEET_IDS matches Sheets exactly

### Issue: Password hashing fails
**Cause:** Row doesn't have 6 columns  
**Fix:** Add "Unit" column (E) before "Foto" column (F) in Sheets

### Issue: "Execution timeout"
**Cause:** Dataset still too large  
**Fix:** Archive old records, keep only last 2000 rows active

### Issue: Login always fails after migration
**Cause:** Passwords not properly hashed  
**Fix:** Re-run `migrateHashPasswords()` and verify Sheets output

---

**Patch Created:** March 22, 2026  
**Applied By:** GitHub Copilot  
**Status:** Ready for Production Testing ✅

