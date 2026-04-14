const FOLDER_ID = "11Uz5gfuzV-X5m83MsjKx90yfX8Cv_xR-"; 
const ID_FEEDBACK_PATIENT = "1rUKIsFHHWJIq885eRyHtJWSvzahW7lszQLr4e68Fbjw"; 
const ID_FEEDBACK_CLINIC = "14tqRzsZWtXL1ciBW_Zas4VKsijrWEaqAZAe-AfFiI3o";
const ID_FEEDBACK_STAF = "1xdWVGZE8GGxtG9tHHQdfXp4IXhaaKK-p0cnGL4wKtOs";

// --- CORE FUNCTIONS ---

function doGet() {
  return HtmlService.createTemplateFromFile('Index').evaluate()
      .setTitle('Ara Locum Hub')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
      .addMetaTag('viewport', 'width=device-width, initial-scale=1.0, maximum-scale=1.0, user-scalable=no');
}

function lightHash(text) {
  try { return Utilities.base64Encode(text); } catch(e) { return text; }
}

function loginUser(phone, password) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Users");
  var data = sheet.getDataRange().getDisplayValues();
  var encodedInput = lightHash(password); 
  
  for (var i = 1; i < data.length; i++) {
    var storedPass = data[i][1].trim();
    if (data[i][0].trim() == phone.trim() && (storedPass == password.trim() || storedPass == encodedInput)) {
      return { 
        success: true, 
        nama: data[i][2], 
        role: data[i][3], 
        phone: data[i][0],
        email: data[i][4], 
        mmc: data[i][5], 
        apc: data[i][6], 
        indemnity: data[i][7], 
        tempatKerja: data[i][8],
        points: data[i][10] ? Number(data[i][10]) : 0,
        badges: data[i][11] || "" // <--- WAJIB TAMBAH INI supaya App boleh baca badges
      };
    }
  }
  return { success: false, message: "Invalid Phone Number or Password!" };
}

function changePassword(phone, newPassword) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Users");
  var data = sheet.getDataRange().getValues();
  var encodedPass = lightHash(newPassword); 
  
  for (var i = 1; i < data.length; i++) {
    if (data[i][0].toString().trim() == phone.trim()) {
      sheet.getRange(i + 1, 2).setValue(encodedPass);
      return "Password successfully updated & secured!";
    }
  }
  return "Error: User not found.";
}

function getInitialData() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Slots");
  if (!sheet) return [];
  var data = sheet.getDataRange().getDisplayValues().slice(1);
  return data.map(r => ({
    id: r[0], tarikh: r[1], masa: r[2], cawangan: r[3], status: r[4], 
    dr: r[5], phone: r[7], gaji: r[8]
  }));
}

function getMyApplications(phone) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Slots");
  if (!sheet) return [];
  var data = sheet.getDataRange().getDisplayValues().slice(1);
  
  // Tapis slot ikut phone Dr dan pastikan ID (r[0]) dihantar sekali
  return data.filter(r => r[7].trim() == phone.trim()).map(r => ({
    id: r[0],        // <--- INI KUNCI DIA! Dr tertinggal baris ni tadi
    tarikh: r[1], 
    masa: r[2], 
    cawangan: r[3], 
    status: r[4]
  }));
}

function bookSlot(slotId, locumName, locumPhone) {
  if(!locumName || !locumPhone) return "Error: Invalid user session.";
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Slots");
  var data = sheet.getDataRange().getValues();
  
  for (var i = 1; i < data.length; i++) {
    if (data[i][0] == slotId) {
      if(data[i][4] !== "Available") return "Slot is no longer available.";
      
      // Kemaskini Status, Nama, dan No Telefon
      sheet.getRange(i + 1, 5).setValue("Pending");
      sheet.getRange(i + 1, 6).setValue(locumName);
      sheet.getRange(i + 1, 8).setValue(locumPhone);
      
      // --- TAMBAHAN: REKOD MASA BOOKING ---
      // Kita ambil masa sekarang (Waktu Malaysia GMT+8)
      var timestamp = Utilities.formatDate(new Date(), "GMT+8", "dd/MM HH:mm:ss");
      sheet.getRange(i + 1, 12).setValue(timestamp); // Simpan di Kolum L (12)
      
      return "Application successfully submitted at " + timestamp;
    }
  }
  return "Error: Slot not found.";
}

// --- ADMIN FUNCTIONS ---

function adminApproveSlot(id) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Slots");
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (data[i][0] == id) {
      sheet.getRange(i + 1, 5).setValue("Approved");
      var drName = data[i][5];
      var drPhone = data[i][7];
      var tarikh = data[i][1];
      var masa = data[i][2];
      var cawangan = data[i][3];
      
      logAdminActivity("APPROVED: Slot " + id + " for Dr " + drName);

      var userData = ss.getSheetByName("Users").getDataRange().getValues();
      var drEmail = "";
      for(var j=1; j<userData.length; j++){
        if(userData[j][0].toString() == drPhone.toString()) { drEmail = userData[j][4]; break; }
      }
      
      if(drEmail) {
        var subject = "LOCUM SLOT CONFIRMATION: " + cawangan;
        var message = "Hi/salam Dr " + drName + ",\n\nThank you for taking up the slot, details as below:\n\n" +
                      "Date: " + tarikh + "\n" +
                      "Time: " + masa + "\n" +
                      "Branch: " + cawangan + "\n\n" +
                      "Please check the app for more info.\n\nThank you.";
        try { MailApp.sendEmail(drEmail, subject, message); } catch(e) { console.log("Email failed: " + e.message); }
      }
      return "Slot Approved & Email Sent!";
    }
  }
}

/** FUNGSI SUPER ADMIN: REPLACE, CANCEL, DELETE (VERSION 2.1 - LOG AWARE) **/
function adminManageSlot(action, id, newDrPhone, manualName) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Slots");
  var logSheet = ss.getSheetByName("Activity Log"); // Tab log untuk Reward Scanner
  var data = sheet.getDataRange().getValues();
  
  for (var i = 1; i < data.length; i++) {
    if (data[i][0].toString() === id.toString()) {
      
      // Ambil nama doktor asal sebelum dipadam (PENTING!)
      var drAsal = data[i][5] ? data[i][5].toString() : "Unknown";
      var branch = data[i][3] ? data[i][3].toString() : "N/A";

      // PILIHAN 1: DELETE (Padam terus baris dari sheet)
      if (action === 'DELETE') {
        sheet.deleteRow(i + 1);
        logAdminActivity("ADMIN: PERMANENT DELETE - Slot ID " + id);
        return "✅ Slot successfully deleted permanently.";
      }
      
      // PILIHAN 2: CANCEL (Bagi Available balik)
      if (action === 'CANCEL') {
        // --- LOGIK PENTING UNTUK REWARD SCANNER ---
        // Walaupun admin yang buat, kita rekodkan dlm Activity Log supaya scanner detect dia pernah cancel
        if (logSheet && drAsal !== "" && drAsal !== "Doctor") {
           logSheet.appendRow([
             new Date(), 
             drAsal, 
             "CANCEL SLOT", 
             "(Admin Assisted) ID: " + id + " | Branch: " + branch
           ]);
        }

        sheet.getRange(i + 1, 5).setValue("Available"); // Status (E)
        sheet.getRange(i + 1, 6).setValue("");          // Nama (F) dipadam selepas dicatat dlm log
        sheet.getRange(i + 1, 8).setValue("");          // Phone (H)
        sheet.getRange(i + 1, 12).setValue("");         // bookedAt (L)
        
        logAdminActivity("ADMIN: CANCEL & RESET - Slot ID " + id + " (Dr: " + drAsal + ")");
        return "✅ Slot reset to Available. Cancellation recorded for Dr. " + drAsal;
      }
      
      // PILIHAN 3: REPLACE (Tukar kepada doktor lain)
      if (action === 'REPLACE') {
        var userSheet = ss.getSheetByName("Users");
        var userData = userSheet.getDataRange().getValues();
        
        var finalName = "";
        var finalPhone = "";

        // 1. Semak jika ada nama manual
        if (manualName && manualName.trim() !== "") {
          finalName = manualName.trim() + " (External)";
          finalPhone = "MANUAL";
        } 
        // 2. Jika tiada, cari dlm database (Laluan asal Dr)
        else {
          for (var j = 1; j < userData.length; j++) {
            if (userData[j][0].toString().trim() === newDrPhone.toString().trim()) {
              finalName = userData[j][2];
              finalPhone = userData[j][0];
              break;
            }
          }
        }

        // 3. Proses kemaskini Sheet jika nama dijumpai
        if (finalName) {
          // Rekod pembatalan Dr Asal dlm Log (Kekal logik asal Dr)
          if (logSheet && drAsal !== "" && drAsal !== "Doctor") {
            logSheet.appendRow([
              new Date(), 
              drAsal, 
              "CANCEL SLOT", 
              "(Admin Replaced) ID: " + id + " | Branch: " + branch
            ]);
          }

          sheet.getRange(i + 1, 5).setValue("Approved");
          sheet.getRange(i + 1, 6).setValue(finalName);
          sheet.getRange(i + 1, 8).setValue(finalPhone);
          
          logAdminActivity("ADMIN: REPLACE - Slot " + id + " (Old: " + drAsal + " -> New: " + finalName + ")");
          return "✅ Successfully replaced with: " + finalName;
        }
        return "❌ Error: Doctor not found.";
      } // Tutup if REPLACE
    } // Tutup if matching ID
  } // Tutup for loop
  return "❌ Error: Slot ID not found.";
} // Tutup keseluruhan fungsi

function getDoctorListForAdmin() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Users");
  var data = sheet.getDataRange().getValues();
  // Ambil hanya yang role 'Doctor' (Kolum D / Index 3)
  return data.filter(r => r[3] === "Doctor").map(r => ({nama: r[2], phone: r[0]}));
}

function adminCreateBulkSlots(obj) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Slots");
  var rows = [];
  var now = new Date().getTime();

  // Kita bina senarai baris (array) dalam memori dulu, bukan terus tulis kat sheet
  obj.dates.forEach(function(d, index) {
    // Tukar format YYYY-MM-DD kepada DD/MM/YYYY
    var parts = d.split('-');
    var formattedDate = parts[2] + "/" + parts[1] + "/" + parts[0];
    
    // Jana ID Unik simple
    var id = "SLOT" + (now + index);
    
    // Susunan Kolum: A=ID, B=Tarikh, C=Masa, D=Cawangan, E=Status, F-H=Kosong, I=Gaji
    rows.push([id, formattedDate, obj.masa, obj.cawangan, "Available", "", "", "", obj.gaji]);
  });

  try {
    if (rows.length > 0) {
      // TEKNIK TURBO: Tulis semua baris dalam satu arahan sahaja
      var startRow = sheet.getLastRow() + 1;
      sheet.getRange(startRow, 1, rows.length, 9).setValues(rows);
      
      return "✅ Berjaya tambah " + rows.length + " slot!";
    }
  } catch (e) {
    return "❌ Error Server: " + e.message;
  }
  return "❌ Tiada tarikh dipilih.";
}

function getAnnouncement() {
  return PropertiesService.getScriptProperties().getProperty('announcement') || "No current announcements.";
}

/** 3. FUNGSI MANAGE ANNOUNCEMENT **/
function saveAnnouncement(text) {
  PropertiesService.getScriptProperties().setProperty('announcement', text);
  logAdminActivity("UPDATED ANNOUNCEMENT: " + text.substring(0, 20) + "...");
  return "Announcement updated!";
}

function getAdminDashboardData(selectedMonth, selectedYear) {
  // ZASSS: Jika Dr tak pilih bulan/tahun (cth: masa mula-mula load), guna tarikh harini
  var now = new Date();
  if (!selectedMonth) selectedMonth = (now.getMonth() + 1).toString().padStart(2, '0');
  if (!selectedYear) selectedYear = now.getFullYear().toString();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  // Guna getDisplayValues supaya format RM dlm Sheet (Kolum I, J, K) kekal tepat
  var slots = ss.getSheetByName("Slots").getDataRange().getDisplayValues().slice(1);
  var users = ss.getSheetByName("Users").getDataRange().getDisplayValues().slice(1);
  
  var today = new Date();
  today.setHours(0,0,0,0);
  var filteredByMonth = slots.filter(r => {
    if (!r[1]) return false;
    // ZASSS: Guna logik yang boleh pecah '/' dan '-' sekaligus
    var p = r[1].includes('-') ? r[1].split('-') : r[1].split('/');
    var day = r[1].includes('-') ? p[2] : p[0];
    var month = r[1].includes('-') ? p[1] : p[1];
    var year = r[1].includes('-') ? p[0] : p[2];
    
    return (month.padStart(2, '0') === selectedMonth && year.toString() === selectedYear);
  });

  // 1. Data untuk Past Slots History (Approved & Tarikh Lepas)
  var pastSlots = slots.filter(r => {
    if (!r[1]) return false;
    var dateParts = r[1].split('/');
    var slotDate = dateParts.length === 3 ? new Date(dateParts[2], dateParts[1] - 1, dateParts[0]) : new Date(r[1]);
    
    // Feature: Tapis ikut filter Bulan & Tahun di Dashboard
    var m = (slotDate.getMonth() + 1).toString().padStart(2, '0');
    var y = slotDate.getFullYear().toString();
    
    return slotDate < today && r[4] === "Approved" && m === selectedMonth && y === selectedYear; 
  }).map(r => ({
      id: r[0],        
      tarikh: r[1],    
      masa: r[2],      
      cawangan: r[3],  
      status: r[4],    
      dr: r[5],        
      bayaran: r[8],   // Kolum I (PAY RM) <-- Ini ubat supaya tak hilang
      sales: r[9],     // Kolum J (SALES)  <-- Ini ubat supaya tak hilang
      pesakit: r[10]   // Kolum K (PTS)    <-- Ini ubat supaya tak hilang
  })).reverse(); 

  // 2. Data untuk Pending Tasks (Guna trim() dan toUpperCase() supaya kebal ejaan)
  var pendingTasks = slots.filter(r => {
    var status = r[4] ? r[4].toString().trim().toUpperCase() : "";
    return status === "PENDING"; 
  }).map(r => ({
    id: r[0], 
    tarikh: r[1], 
    cawangan: r[3], 
    dr: r[5], 
    phone: r[7], 
    masa: r[2],
    bookedAt: r[11] 
  }));

  // --- TAMBAHAN 2: KIRA ARACOINS BULANAN ---
  var monthlyLeaderboard = users.filter(u => u[3] === "Doctor").map(u => {
    var pointsMonth = 0;
    var badgeStr = u[11] || "";
    
    if (badgeStr) {
      badgeStr.split(',').forEach(item => {
        var match = item.match(/\((.*?)\)/);
        var dateBadgeId = selectedMonth + "/" + selectedYear; 
        
        if (match && match[1].trim() === dateBadgeId) {
          // Kod pengiraan point Dr bermula di sini...
          // 1. Ambil angka selepas titik bertindih (contoh :2 atau :3)
          var parts = item.split(':');
          var count = (parts.length === 2) ? (parseInt(parts[1]) || 1) : 1; 

          // 2. Tentukan point ikut jenis badge (Semua 6 Jenis)
          if (item.includes("Team Favorite")) {
            pointsMonth += (20 * count);
          } 
          else if (item.includes("Heart Winner") || item.includes("Savior") || item.includes("Last Minute Savior")) {
            pointsMonth += (15 * count);
          } 
          else if (item.includes("Iron Doctor") || item.includes("Unstoppable") || item.includes("Diligent")) {
            pointsMonth += (10 * count);
          }
          else {
            pointsMonth += (10 * count); 
          }
        }
      }); // Tutup forEach
    } // Tutup if (badgeStr)

    return { 
      nama: u[2], 
      phone: u[0], 
      points: pointsMonth, 
      complete: (u[4] && u[5] && u[6]) ? true : false 
    };
  }).sort((a, b) => b.points - a.points); // Terus susun tanpa tapis

  // ... (kod atas sama) ...
  
  return {
    total: filteredByMonth.length, // Ini akan kira semua baris dlm bulan tu (Auto + Manual)
    pending: filteredByMonth.filter(r => r[4] === "Pending").length,
    available: filteredByMonth.filter(r => r[4] === "Available").length,
    approved: filteredByMonth.filter(r => r[4] === "Approved").length, // Tambah ni untuk rujukan
    pendingApprovals: pendingTasks,
    directory: monthlyLeaderboard, // Ini yang hantar senarai semua doktor ke App
    pastSlots: pastSlots           // Ini yang hantar history slot ke App
  };
}
function getAdminFeedbackAll() {
  function getData(id) { try { return SpreadsheetApp.openById(id).getSheets()[0].getDataRange().getDisplayValues(); } catch(e) { return [["No Data"]]; } }
  return {
    stafLocum: getData(ID_FEEDBACK_STAF),
    locumClinic: getData(ID_FEEDBACK_CLINIC),
    patientDoc: getData(ID_FEEDBACK_PATIENT)
  };
}

function getDoctorFeedback(doctorName) {
  try {
    var extSS = SpreadsheetApp.openById("1rUKIsFHHWJIq885eRyHtJWSvzahW7lszQLr4e68Fbjw");
    var searchName = doctorName.toLowerCase().trim();
    var results = [];

    // --- 1. PROSES TAB PERTAMA (Form Responses) ---
    var sheetForm = extSS.getSheets()[0];
    var dataForm = sheetForm.getDataRange().getDisplayValues().slice(1);
    
    dataForm.forEach(r => {
      var drCol9 = r[9] ? r[9].toLowerCase().trim() : ""; // Kolum J (Nama Dr)
      if (drCol9.includes(searchName)) {
        // Kira rating purata dari kolum 4,5,6,7 (Likert Scale)
        var rating5Star = hitungRating(r); 
        results.push({
          tarikh: r[0],
          pesakit: r[1] || "Patient",
          cawangan: r[3] || "Klinik ARA",
          komen: r[8] || "Tiada ulasan spesifik.", // Kolum I (Maklum balas lain-lain)
          rating: rating5Star
        });
      }
    });

    // --- 2. PROSES TAB KEDUA (manual feedbac) ---
    var sheetManual = extSS.getSheetByName("manual feedbac");
    if (sheetManual) {
      var dataManual = sheetManual.getDataRange().getDisplayValues().slice(1);
      dataManual.forEach(r => {
        var drCol6 = r[6] ? r[6].toLowerCase().trim() : ""; // Kolum G (Nama Dr)
        if (drCol6.includes(searchName)) {
          results.push({
            tarikh: r[0] || "Manual",
            pesakit: r[1] || "Patient", // Kolum B
            cawangan: r[3] || "Klinik ARA", // Kolum D
            komen: r[5] || "Tiada ulasan.", // Kolum F (Ulasan/Komen)
            rating: parseFloat(r[4] || 5).toFixed(1) // Kolum E (Rating 1-5)
          });
        }
      });
    }

    // --- FUNGSI PEMBANTU: TUKAR LIKERT KE 5-STAR (VERSI FIX) ---
    function hitungRating(r) {
      function k(t) {
        if(!t) return 0;
        var s = t.toUpperCase().trim();
        if(s === "SANGAT SETUJU") return 4;
        if(s === "SETUJU") return 3;
        if(s === "TIDAK SETUJU") return 2;
        if(s === "SANGAT TIDAK SETUJU") return 1;
        return isNaN(t) ? 0 : parseFloat(t);
      }
      
      var scores = [k(r[4]), k(r[5]), k(r[6]), k(r[7])].filter(s => s > 0);
      if(scores.length === 0) return "5.0";
      
      var avg = scores.reduce((a,b) => a+b) / scores.length;
      return ((avg / 4) * 5).toFixed(1);
    }

    // --- BAHAGIAN INI KEKAL (JANGAN BUANG) ---
    return results.sort((a, b) => new Date(b.tarikh) - new Date(a.tarikh));

  } catch(e) { 
    return []; 
  }
}

function updateProfile(phone, email, mmc, apc, bank, tempatKerja, mmcFile, apcFile, insFile, indStatus) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Users");
  var folder = DriveApp.getFolderById(FOLDER_ID);
  var data = sheet.getDataRange().getValues();
  var doctorName = "DOCTOR"; 
  var rowIndex = -1;

  // 1. Cari row user & ambil nama doktor untuk penamaan fail
  for (var i = 1; i < data.length; i++) {
    if (data[i][0].toString().trim() == phone.trim()) {
      doctorName = data[i][2]; // Nama doktor di kolum C (index 2)
      rowIndex = i + 1; 
      break;
    }
  }

  if (rowIndex == -1) return "Error: User profile not found.";

  // 2. Fungsi upload ikut gaya asal awak (guna Nama Doktor)
  function upload(fileData, label) {
    if (!fileData || !fileData.contents) return null;
    var fileName = doctorName.toUpperCase().replace(/\s+/g, '_') + "_" + label;
    var blob = Utilities.newBlob(Utilities.base64Decode(fileData.contents), fileData.mimeType, fileName);
    var file = folder.createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW); // Supaya admin boleh buka fail
    return file.getUrl();
  }

  // 3. Proses muat naik fail
  var mmcUrl = upload(mmcFile, "MMC");
  var apcUrl = upload(apcFile, "APC2026"); 
  var insUrl = upload(insFile, "INDEMNITY"); // Muat naik fail indemnity

  // 4. Update data ke dalam Sheet mengikut kolum yang betul
  sheet.getRange(rowIndex, 5).setValue(email); // Kolum E
  
  if(mmcUrl) sheet.getRange(rowIndex, 6).setValue(mmc + " | " + mmcUrl); // Kolum F
  else if(mmc) sheet.getRange(rowIndex, 6).setValue(mmc); // Jika tukar teks tapi tak upload fail baru

  if(apcUrl) sheet.getRange(rowIndex, 7).setValue(apc + " | " + apcUrl); // Kolum G
  else if(apc) sheet.getRange(rowIndex, 7).setValue(apc);

  // Simpan Status Indemnity & Link ke Kolum H
  if(indStatus) {
    var indValue = insUrl ? indStatus + " | " + insUrl : indStatus;
    sheet.getRange(rowIndex, 8).setValue(indValue); // Kolum H
  }

  sheet.getRange(rowIndex, 9).setValue(tempatKerja); // Kolum I

  return "Profile Successfully Updated!";
}

function logAdminActivity(action) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var logSheet = ss.getSheetByName("AdminLogs") || ss.insertSheet("AdminLogs");
  logSheet.appendRow([new Date(), action]);
}

function adminGivePoints(phone, points, awardName) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var userSheet = ss.getSheetByName('Users');
  var data = userSheet.getDataRange().getValues();
  
  for (var i = 1; i < data.length; i++) {
    if (data[i][0].toString().trim() == phone.trim()) { 
      
      // 1. TAMBAH POINT (KOLUM K)
      var currentPoints = parseInt(data[i][10]) || 0;
      userSheet.getRange(i + 1, 11).setValue(currentPoints + parseInt(points)); 

      // 2. LOGIK ASINGKAN NAMA & ID [R]
      var cleanBadgeName = awardName.split(' [')[0].trim(); // "Heart Winner (03/2026)"
      var rowTag = awardName.match(/\[R\d+\]/) ? awardName.match(/\[R\d+\]/)[0] : ""; // "[R3]"

      // 3. UPDATE BADGES (KOLUM L) - CANTUMKAN :2, :3
      var badgeString = data[i][11] ? data[i][11].toString() : "";
      var badgeMap = {};

      if (badgeString) {
        badgeString.split(',').forEach(function(item) {
          var parts = item.split(':');
          if (parts.length == 2) {
            var key = parts[0].trim().split(' [')[0]; // Repair nama lama
            badgeMap[key] = (badgeMap[key] || 0) + parseInt(parts[1]);
          }
        });
      }
      badgeMap[cleanBadgeName] = (badgeMap[cleanBadgeName] || 0) + 1;

      var updatedBadgeString = Object.keys(badgeMap).map(k => k + ":" + badgeMap[k]).join(', ');
      userSheet.getRange(i + 1, 12).setValue(updatedBadgeString);
      
      // 4. SIMPAN LOCK (KOLUM M) - SUPAYA REVIEW LAIN TAK HILANG
      if (rowTag) {
        var currentLock = userSheet.getRange(i + 1, 13).getValue().toString();
        if (currentLock.indexOf(rowTag) === -1) {
          userSheet.getRange(i + 1, 13).setValue(currentLock + rowTag);
        }
      }

      logAdminActivity("AWARD GIVEN: " + cleanBadgeName + " to " + data[i][2]);
      return "✅ Berjaya! " + cleanBadgeName + " kini Total: " + badgeMap[cleanBadgeName];
    }
  }
  return "Error: User tak jumpa.";
}

function getLatestUserProfile(phone) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var data = ss.getSheetByName("Users").getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (data[i][0].toString().trim() == phone.toString().trim()) {
      return {
        points: data[i][10] ? Number(data[i][10]) : 0,
        badges: data[i][11] || ""
      };
    }
  }
  return null;
}
function completeSlotAndAwardPoints(slotId, sales, patients, payment, period) { 
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Slots");
  var userSheet = ss.getSheetByName("Users");
  var data = sheet.getDataRange().getValues();
  
  if (!slotId || slotId.toString().trim() === "") {
    return "❌ ERROR: ID Slot kosong.";
  }

  for (var i = 1; i < data.length; i++) {
    if (data[i][0].toString().trim() === slotId.toString().trim()) {
      
      // 1. UPDATE DATA PERFORMANCE
      sheet.getRange(i + 1, 9).setValue(payment);  
      sheet.getRange(i + 1, 10).setValue(sales);   
      sheet.getRange(i + 1, 11).setValue(patients);
      SpreadsheetApp.flush();
      
      var drNameInSlot = data[i][5] ? data[i][5].toString().trim().toUpperCase() : ""; 
      var slotTimeRaw = data[i][2] ? data[i][2].toString().toLowerCase() : ""; 
      var branchRaw = data[i][3] ? data[i][3].toString().toUpperCase() : "";
      var slotDateRaw = data[i][1];
      var bookedAtRaw = data[i][11];
      
      var doctorsToReward = [];
      if (drNameInSlot === "CME") {
        doctorsToReward = patients.toString().split(",").map(function(n) { return n.trim().toUpperCase(); }).filter(function(n) { return n !== ""; });
      } else if (drNameInSlot !== "") {
        doctorsToReward.push(drNameInSlot);
      }

      if (doctorsToReward.length === 0) return "✅ Performance saved.";
      var globalFoundDoctors = [];

      // 2. LOOP SETIAP DOKTOR
      doctorsToReward.forEach(function(currentDr) {
        var badgesToUpdate = []; 
        var dateLockId = "(" + slotDateRaw + ")";

        // LOGIK IRON DOCTOR
        var numbersOnly = slotTimeRaw.replace(/[^0-9]/g, ""); 
        if (/8.*8|9.*9|10.*10/.test(numbersOnly) || slotTimeRaw.includes("12h") || slotTimeRaw.includes("12jam")) {
          badgesToUpdate.push("Iron Doctor");
        }

        // LOGIK CME
        if (branchRaw.includes("CME") || branchRaw.includes("BRIEFING")) {
          badgesToUpdate.push("The Diligent Doc");
        }

        // LOGIK SAVIOR
        if (bookedAtRaw && slotDateRaw) {
          try {
            function getValidDate(d) {
              if (d instanceof Date) return d;
              var str = d.toString();
              var parts = str.split(/[/\s:-]/);
              if (parts.length >= 3) return new Date(parts[2], parts[1]-1, parts[0], parts[3]||0, parts[4]||0, parts[5]||0);
              return new Date(str);
            }
            var sDate = getValidDate(slotDateRaw);
            var bDate = getValidDate(bookedAtRaw);
            var diffInHours = (sDate.getTime() - bDate.getTime()) / (1000 * 60 * 60);
            if (diffInHours > 0 && diffInHours < 48) badgesToUpdate.push("Last Minute Savior");
          } catch(e) { console.log(e.message); }
        }

        // 3. UPDATE TAB USERS (ANTI-DOUBLE)
        if (badgesToUpdate.length > 0) {
          var userData = userSheet.getDataRange().getValues();
          for (var j = 1; j < userData.length; j++) {
            if (userData[j][2].toString().trim().toUpperCase() === currentDr) {
              
              // POINT 2: SUPER LOCKING LOGIC (Anti-Double Entry)
              var lock = LockService.getScriptLock();
              try {
                lock.waitLock(30000); 

                // Ambil data segar dari Kolum K, L, dan M
                var latestUserData = userSheet.getRange(j + 1, 11, 1, 3).getValues()[0]; 
                var currentPoints = parseInt(latestUserData[0]) || 0; 
                var badgeString = latestUserData[1] ? latestUserData[1].toString() : "";
                var lockHistory = latestUserData[2] ? latestUserData[2].toString() : ""; 
                
                var currentLockId = "[" + slotId + "]";
                var monthlyId = "(" + period + ")"; 

                // PAGAR UTAMA: Jika Slot ID ini tiada dalam sejarah (LockHistory)
                if (lockHistory.indexOf(currentLockId) === -1) {

                  // Tapis badge yang layak
                  var validBadges = badgesToUpdate.filter(function(bName) {
                    return badgeString.indexOf(bName + " " + monthlyId) === -1 || lockHistory.indexOf(currentLockId) === -1;
                  });

                  if (validBadges.length > 0) {
                    // 1. TAMBAH POINT
                    userSheet.getRange(j + 1, 11).setValue(currentPoints + (validBadges.length * 10));

                    // 2. TANDA LOCK (Simpan ID Slot dlm Kolum M)
                    userSheet.getRange(j + 1, 13).setValue(lockHistory + currentLockId);

                    // 3. KEMASKINI STRING AWARD (Kolum L)
                    var badgeMap = {};
                    if (badgeString) {
                      badgeString.split(',').forEach(function(item) {
                        var parts = item.split(':');
                        if (parts.length == 2) {
                          badgeMap[parts[0].trim()] = parseInt(parts[1]);
                        }
                      });
                    }

                    validBadges.forEach(function(bName) {
                      var cleanKey = bName + " " + monthlyId; 
                      badgeMap[cleanKey] = (badgeMap[cleanKey] || 0) + 1;
                    });

                    var updatedBadgeStr = Object.keys(badgeMap).map(function(k) { 
                      return k + ":" + badgeMap[k]; 
                    }).join(', ');

                    userSheet.getRange(j + 1, 12).setValue(updatedBadgeStr);
                    
                    SpreadsheetApp.flush(); // Paksa simpan terus
                    globalFoundDoctors.push(currentDr);
                  } 
                } else {
                   console.log("Slot ID " + slotId + " already has points. Skipping.");
                }
              } catch (e) {
                console.log("Lock error: " + e.message);
              } finally {
                lock.releaseLock(); 
              }
              break; 
            }
          }
        }
      }); // Tutup forEach

      return globalFoundDoctors.length > 0 ? "✅ Success! Awarded to: " + globalFoundDoctors.join(", ") : "✅ Performance saved.";
    }
  }
  return "❌ Slot ID not found.";
}

function processMonthlyUnstoppable(selectedMonth, selectedYear) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var slotSheet = ss.getSheetByName("Slots");
  var userSheet = ss.getSheetByName("Users");
  
  // --- INTEGRASI ACTIVITY LOG (TAMBAHAN) ---
  var logSheet = ss.getSheetByName("Activity Log");
  var logData = logSheet ? logSheet.getDataRange().getValues() : [];
  // ------------------------------------------

  var slotData = slotSheet.getDataRange().getValues();
  var userData = userSheet.getDataRange().getValues();
  var userRange = userSheet.getRange(2, 11, userData.length - 1, 2); 
  var userUpdateValues = userRange.getValues(); 
  
  var summary = {}; 
  var dateBadgeId = "(" + selectedMonth + "/" + selectedYear + ")";

  // --- KOD ASAL DR (SCAN SLOTS) ---
  for (var i = slotData.length - 1; i >= 1; i--) {
    var rawDate = slotData[i][1];
    if (!rawDate) continue;

    var dateVal;
    if (rawDate instanceof Date) { dateVal = rawDate; } 
    else {
      var parts = rawDate.toString().split(/[/-]/);
      dateVal = new Date(parts[2], parts[1]-1, parts[0]);
    }

    if (isNaN(dateVal.getTime())) continue;

    var m = (dateVal.getMonth() + 1).toString().padStart(2, '0');
    var y = dateVal.getFullYear().toString();
    
    if (m === selectedMonth && y === selectedYear) {
      var dr = slotData[i][5] ? slotData[i][5].toString().trim().toUpperCase() : "";
      var status = slotData[i][4] ? slotData[i][4].toString().trim() : "";
      
      if (dr) {
        if (!summary[dr]) summary[dr] = { approved: 0, cancelled: 0 };
        if (status === "Approved") summary[dr].approved++;
        if (status === "Cancelled") summary[dr].cancelled++;
      }
    }
    
    var diffMonth = (parseInt(selectedYear) * 12 + parseInt(selectedMonth)) - (dateVal.getFullYear() * 12 + (dateVal.getMonth() + 1));
    if (diffMonth > 4) break; 
  }

  // --- INTEGRASI ACTIVITY LOG (TAMBAHAN LOGIK TANPA UBAH STRUKTUR ASAL) ---
  // Kita scan log pulak untuk cari siapa yang cancel tapi slot dah jadi 'Available' semula
  for (var k = 1; k < logData.length; k++) {
    var logDate = logData[k][0];
    if (!(logDate instanceof Date)) logDate = new Date(logDate);
    
    var lm = (logDate.getMonth() + 1).toString().padStart(2, '0');
    var ly = logDate.getFullYear().toString();

    if (lm === selectedMonth && ly === selectedYear) {
      var logAction = logData[k][2] ? logData[k][2].toString() : "";
      if (logAction === "CANCEL SLOT") {
        var logDr = logData[k][1] ? logData[k][1].toString().trim().toUpperCase() : "";
        if (logDr) {
          if (!summary[logDr]) summary[logDr] = { approved: 0, cancelled: 0 };
          summary[logDr].cancelled++; // Tambah rekod cancel dari log ke dalam summary sedia ada
        }
      }
    }
  }
  // ------------------------------------------------------------------------

  var awardedList = [];
  var skippedList = [];
  var changeMade = false;

  // --- KOD ASAL DR (AWARDING LOGIC) ---
  for (var drName in summary) {
    if (summary[drName].approved >= 2 && summary[drName].cancelled === 0) {
      for (var j = 1; j < userData.length; j++) {
        var sheetDrName = userData[j][2] ? userData[j][2].toString().trim().toUpperCase() : "";
        if (sheetDrName === drName) {
          var currentPoints = parseInt(userUpdateValues[j-1][0]) || 0;
          var badgeStr = userUpdateValues[j-1][1] ? userUpdateValues[j-1][1].toString() : "";
          
          if (badgeStr.indexOf("The Unstoppable " + dateBadgeId) === -1) {
            userUpdateValues[j-1][0] = currentPoints + 10;
            var badgeMap = {};
            if (badgeStr) {
              badgeStr.split(',').forEach(function(item) {
                var p = item.split(':');
                if (p.length == 2) {
                  var clean = p[0].split(' (')[0].trim();
                  badgeMap[clean] = (badgeMap[clean] || 0) + parseInt(p[1]);
                }
              });
            }
            badgeMap["The Unstoppable"] = (badgeMap["The Unstoppable"] || 0) + 1;
            userUpdateValues[j-1][1] = Object.keys(badgeMap).map(k => k + " " + dateBadgeId + ":" + badgeMap[k]).join(', ');
            awardedList.push(drName);
            changeMade = true;
          } else {
            skippedList.push(drName);
          }
          break;
        }
      }
    }
  }

  if (changeMade) userRange.setValues(userUpdateValues);

  var result = "";
  if (awardedList.length > 0) result += "✅ SUCCESS! Awarded: " + awardedList.join(", ");
  if (skippedList.length > 0) result += (result ? "\n" : "") + "ℹ️ ALREADY AWARDED: " + skippedList.join(", ");
  if (result === "") result = "❌ No eligible doctors for " + selectedMonth + "/" + selectedYear;
  
  return result;
}

function getManualHeartCandidates() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var userData = ss.getSheetByName("Users").getDataRange().getValues();
    
    var extSS = SpreadsheetApp.openById("1rUKIsFHHWJIq885eRyHtJWSvzahW7lszQLr4e68Fbjw");
    var sheetManual = extSS.getSheetByName("manual feedbac");
    if (!sheetManual) return [];

    var manualData = sheetManual.getDataRange().getDisplayValues(); // Ambil semua termasuk header
    var candidates = [];

    // Kita start dari i=1 (skip header)
    for (var i = 1; i < manualData.length; i++) {
      var r = manualData[i];
      var tarikhReview = r[0]; // Kolum A
      var rating = parseFloat(r[4]); // Kolum E
      var drName = r[6] ? r[6].trim().toUpperCase() : ""; // Kolum G
      
      if (rating === 5 && drName !== "") {
        // 1. Guna rowId (ID Baris) sebagai "cap jari" unik
        var rowId = "R" + (i + 1);
        
        // 2. Proses tarikh (Jika Dr taip "03/2026" terus dlm Sheet)
        var formattedMonth = "(MM/YYYY)"; 
        if (tarikhReview) {
          var dateStr = tarikhReview.toString().trim();
          
          // Jika Dr letak "03/2026", kita terus balut dengan kurungan
          if (dateStr.includes('/') && dateStr.length <= 7) {
            formattedMonth = "(" + dateStr + ")";
          } 
          // Jika Dr letak tarikh penuh "DD/MM/YYYY", kita ambil belakang dia je
          else if (dateStr.includes('/')) {
            var parts = dateStr.split('/');
            formattedMonth = "(" + parts[1] + "/" + parts[2] + ")";
          }
        }
        
        // 3. KUNCI: Gunakan ID Unik yang ada Row ID sekali
        var specificBadgeId = "Heart Winner " + formattedMonth + " [" + rowId + "]";

        var alreadyAwarded = false;
        for (var j = 1; j < userData.length; j++) {
          if (userData[j][2].toString().trim().toUpperCase() === drName) {
            // Kita semak Kolum L (Index 11) dan Kolum M (Index 12)
            var badges = userData[j][11] ? userData[j][11].toString() : "";
            var lockHistory = userData[j][12] ? userData[j][12].toString() : "";
            
            // Dr, dia hanya akan 'skip' kalau JUMPA specificBadgeId yang ada [R...] tu
            if (badges.indexOf(specificBadgeId) !== -1 || lockHistory.indexOf("[" + rowId + "]") !== -1) {
              alreadyAwarded = true;
            }
            break;
          }
        }

        if (!alreadyAwarded) {
          candidates.push({ 
            name: drName, 
            date: tarikhReview, 
            row: rowId,
            badgeId: specificBadgeId // Ini yang akan dihantar ke adminGivePoints
          });
        }
      }
    } // <--- Dr perlukan kurungan ini untuk tutup loop 'for' di atas

    return candidates.reverse(); 
  } catch(e) {
    return [];
  }
}

/** * FUNGSI BATAL SLOT (BACKEND)
 * Dr pastikan dlm HTML, google.script.run hantar 6 data ikut urutan ini
 */
function doctorCancelSlot(slotId, locumPhone, statusAsal, drName, slotDate, branchName) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var slotSheet = ss.getSheetByName("Slots"); 
  var logSheet = ss.getSheetByName("Activity Log");
  
  var data = slotSheet.getDataRange().getValues();
  var found = false;
  var locumName = "";
  var branch = "";

  // Bersihkan input
  var searchId = String(slotId).trim();

  for (var i = 1; i < data.length; i++) {
    var sheetId = String(data[i][0]).trim(); 

    // PADANKAN GUNA ID
    if (sheetId === searchId) {
      locumName = data[i][5] ? String(data[i][5]) : "Doctor"; 
      branch = data[i][3] ? String(data[i][3]) : "N/A";

      // PROSES KEMASKINI SHEET SLOTS
      slotSheet.getRange(i + 1, 5).setValue("Available"); // Status (E)
      slotSheet.getRange(i + 1, 6).setValue("");          // Nama (F)
      slotSheet.getRange(i + 1, 8).setValue("");          // Phone (H)
      slotSheet.getRange(i + 1, 12).setValue("");         // Timestamp (L)
      
      found = true;
      break;
    }
  }

  if (found) {
    // 1. LOGIK EMEL NOTIFIKASI
    var checkStatus = (statusAsal || "").toString().toUpperCase();
    if (checkStatus === "APPROVED") {
      try {
        // Memanggil fungsi emel (Pastikan fungsi sendCancelNotification ada dlm Code.gs)
        sendCancelNotification(drName, slotDate, branchName);
      } catch (e) {
        Logger.log("Email error: " + e.message);
      }
    }

    // 2. REKOD DALAM ACTIVITY LOG
    if (logSheet) {
      logSheet.appendRow([new Date(), locumName, "CANCEL SLOT", "ID: " + searchId + " | Branch: " + branch + " | Status: " + statusAsal]);
    }

    return "✅ Success! Slot kini Available.";
  } else {
    return "❌ Error: ID " + searchId + " tidak dijumpai dlm Sheet. Sila refresh App.";
  }
}

/**
 * LIVE PROFILE CHECKER
 * Checks the latest data in the "Users" sheet directly.
 */
function checkLiveProfileStatus(phone) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName("Users");
    const data = sheet.getDataRange().getValues();
    
    // Convert phone to string for accurate matching
    const searchPhone = phone.toString().trim();

    for (let i = 1; i < data.length; i++) {
      if (data[i][0].toString().trim() === searchPhone) {
        const email = (data[i][4] || "").toString(); // Column E
        const mmc = (data[i][5] || "").toString();   // Column F
        const apc = (data[i][6] || "").toString();   // Column G
        
        // The 3 Pillars (Rukun)
        const hasEmail = email.includes('@');
        const hasMmc = mmc.length > 2;
        const hasApc = apc.includes('http');
        
        return (hasEmail && hasMmc && hasApc);
      }
    }
    return false; // User not found
  } catch (e) {
    Logger.log("Error in checkLiveProfileStatus: " + e.message);
    return false;
  }
}

function checkLiveProfileStatus(phone) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName("Users");
    const data = sheet.getDataRange().getValues();
    const searchPhone = phone.toString().trim();

    for (let i = 1; i < data.length; i++) {
      if (data[i][0].toString().trim() === searchPhone) {
        const email = (data[i][4] || "").toString(); 
        const mmc = (data[i][5] || "").toString();   
        const apc = (data[i][6] || "").toString();   
        
        const hasEmail = email.includes('@');
        const hasMmc = mmc.length > 2;
        const hasApc = apc.includes('http');
        
        return (hasEmail && hasMmc && hasApc);
      }
    }
    return false;
  } catch (e) {
    return false;
  }
}

function getNewLocumApplications() {
  var ss = SpreadsheetApp.openById("1JhLEA8DjNyt0-fIVybtUY5MCuaP2XsN0UftHlYfe6lM");
  var sheet = ss.getSheetByName("Form responses 1");
  var data = sheet.getDataRange().getDisplayValues().slice(1); // Ambil data kecuali header 
  
  // Susun ikut yang terbaru (Timestamp) 
  return data.reverse().slice(0, 5); // Ambil 5 yang paling terbaru
}

function sendCancelNotification(drName, slotDate, branchName) {
  const adminEmail = "operation@hsohealthcare.com";
  
  const subject = "URGENT: Slot Cancellation - " + drName;
  
  const message = 
    "Dear Operations Team,\n\n" +
    "This is an automated notification to inform you that a slot has been cancelled.\n\n" +
    "DETAILS:\n" +
    "-----------------------------------\n" +
    "Doctor Name: " + drName + "\n" +
    "Slot Date  : " + slotDate + "\n" +
    "Branch     : " + branchName + "\n" +
    "-----------------------------------\n\n" +
    "Please update the schedule and notify the clinic accordingly.\n\n" +
    "Best regards,\n" +
    "HSO Healthcare System";

  MailApp.sendEmail({
    to: adminEmail,
    subject: subject,
    body: message
  });
}
