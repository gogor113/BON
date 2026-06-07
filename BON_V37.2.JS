// ==================== BON UTANG WARUNG v37.2 ====================
// Upgrade dari v37.1 ke v37.2:
// 1. FITUR BARU: Auto Cleanup Log di Google Sheet (max 50 baris)
// 2. Menambahkan endpoint 'cleanOldLogs' untuk membersihkan log otomatis
// 3. Setiap sinkronisasi dan test koneksi dapat memicu cleanup log
// 4. Versi dinaikkan menjadi 37.2 untuk kompatibilitas dengan frontend
// WebApp URL: Setelah deploy, copy URL ke hidden config di HTML

const SHEET_NAME_BONS = "DataBon";
const SHEET_NAME_PAYMENTS = "DataPembayaran";
const SHEET_NAME_USERS = "DataUsers";
const SHEET_NAME_SYNC_LOG = "SyncLog";
const VERSION = "37.2";
const MAX_LOG_LINES = 50; // Maksimal baris log (tidak termasuk header)

function doGet() {
  return doPost({
    postData: { contents: "" }
  });
}

function doPost(e) {
  try {
    const data = e.parameter || {};
    const action = data.action;
    
    if (!action) {
      return ContentService
        .createTextOutput(JSON.stringify({ success: false, error: "No action specified" }))
        .setMimeType(ContentService.MimeType.JSON);
    }
    
    // Initialize sheets if not exist
    initializeSheets();
    
    // Route actions
    if (action === "testConnection") {
      return handleTestConnection();
    } 
    else if (action === "mergeBackupV35" || action === "restoreV35") {
      const jsonData = JSON.parse(data.data || "{}");
      if (action === "mergeBackupV35") {
        return handleMergeBackupV372(jsonData);
      } else {
        return handleRestoreV372(data.username);
      }
    }
    else if (action === "syncUserAuth") {
      const authData = JSON.parse(data.data || "{}");
      return handleSyncUserAuth(authData);
    }
    else if (action === "getUserAuth") {
      return handleGetUserAuth(data.username);
    }
    else if (action === "syncBonV35") {
      return handleSyncBonV372(data.username, JSON.parse(data.bonData || "{}"));
    }
    else if (action === "syncPaymentV35") {
      return handleSyncPaymentV372(data.username, JSON.parse(data.paymentData || "{}"));
    }
    else if (action === "cleanupLunasOtomatis") {
      return handleCleanupLunasOtomatisV372(data.username);
    }
    else if (action === "deleteCustomerData") {
      const customerName = data.customerName;
      const username = data.username;
      return handleDeleteCustomerDataV372(username, customerName);
    }
    // ========== FITUR BARU v37.2 ==========
    else if (action === "cleanOldLogs") {
      const username = data.username || "system";
      const maxLines = parseInt(data.maxLines) || MAX_LOG_LINES;
      return handleCleanOldLogs(username, maxLines);
    }
    
    return ContentService
      .createTextOutput(JSON.stringify({ success: false, error: "Unknown action" }))
      .setMimeType(ContentService.MimeType.JSON);
      
  } catch (error) {
    console.error("Error in doPost:", error);
    return ContentService
      .createTextOutput(JSON.stringify({ success: false, error: error.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function initializeSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  if (!ss.getSheetByName(SHEET_NAME_BONS)) {
    const sheet = ss.insertSheet(SHEET_NAME_BONS);
    sheet.appendRow(["uniqueId", "username", "namaPelanggan", "total", "items", "waktu", "lastModified", "isActive", "version"]);
  }
  
  if (!ss.getSheetByName(SHEET_NAME_PAYMENTS)) {
    const sheet = ss.insertSheet(SHEET_NAME_PAYMENTS);
    sheet.appendRow(["uniqueId", "username", "namaPelanggan", "jumlah", "waktu", "lastModified", "version"]);
  }
  
  if (!ss.getSheetByName(SHEET_NAME_USERS)) {
    const sheet = ss.insertSheet(SHEET_NAME_USERS);
    sheet.appendRow(["username", "password", "security_pet", "security_hobby", "security_food", "registeredAt", "lastLogin", "isActive", "version"]);
  }
  
  if (!ss.getSheetByName(SHEET_NAME_SYNC_LOG)) {
    const sheet = ss.insertSheet(SHEET_NAME_SYNC_LOG);
    sheet.appendRow(["timestamp", "username", "action", "details", "version"]);
  }
}

// ==================== HANDLE CLEAN OLD LOGS (v37.2) ====================
function handleCleanOldLogs(username, maxLines) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME_SYNC_LOG);
    const data = sheet.getDataRange().getValues();
    
    if (data.length <= maxLines + 1) {
      // Tidak perlu dibersihkan
      return ContentService
        .createTextOutput(JSON.stringify({ 
          success: true, 
          message: "Log already within limit", 
          currentRows: data.length - 1,
          maxLines: maxLines,
          deletedCount: 0,
          version: VERSION
        }))
        .setMimeType(ContentService.MimeType.JSON);
    }
    
    // Hitung jumlah baris yang akan dihapus (baris data tertua)
    const rowsToDelete = data.length - (maxLines + 1); // +1 untuk header
    if (rowsToDelete <= 0) {
      return ContentService
        .createTextOutput(JSON.stringify({ 
          success: true, 
          message: "No cleanup needed", 
          deletedCount: 0,
          version: VERSION
        }))
        .setMimeType(ContentService.MimeType.JSON);
    }
    
    // Hapus baris dari bawah (baris tertua adalah setelah header)
    // Baris ke-2 (index 1) adalah data tertua
    for (let i = 0; i < rowsToDelete; i++) {
      sheet.deleteRow(2); // hapus baris kedua (setelah header) berulang
    }
    
    addSyncLog(username, "cleanOldLogs", `Deleted ${rowsToDelete} old log rows, kept max ${maxLines} rows`, VERSION);
    
    return ContentService
      .createTextOutput(JSON.stringify({ 
        success: true, 
        message: `Log cleaned, deleted ${rowsToDelete} old entries`, 
        deletedCount: rowsToDelete,
        currentRows: maxLines,
        maxLines: maxLines,
        version: VERSION
      }))
      .setMimeType(ContentService.MimeType.JSON);
      
  } catch (error) {
    console.error("Error cleaning logs:", error);
    return ContentService
      .createTextOutput(JSON.stringify({ success: false, error: error.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function handleTestConnection() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  // Panggil cleanup log saat test koneksi juga
  try {
    handleCleanOldLogs("system", MAX_LOG_LINES);
  } catch(e) {
    // Abaikan error cleanup jika gagal
  }
  
  return ContentService
    .createTextOutput(JSON.stringify({ 
      success: true, 
      message: "Cloud Backup v37.2 ACTIVE - Auto Log Cleanup Ready (Max 50 lines)", 
      spreadsheetId: ss.getId(),
      sheets: ss.getSheets().map(s => s.getName()),
      version: VERSION,
      maxLogLines: MAX_LOG_LINES
    }))
    .setMimeType(ContentService.MimeType.JSON);
}

function handleMergeBackupV372(data) {
  const username = data.username;
  const remoteBons = data.semuaBon || [];
  const remotePayments = data.pembayaran || [];
  
  // Process Bons
  const existingBons = getAllBonsByUser(username);
  const mergedBons = mergeBonDataV372(existingBons, remoteBons);
  saveAllBonsToSheet(username, mergedBons);
  
  // Process Payments
  const existingPayments = getAllPaymentsByUser(username);
  const mergedPayments = mergePaymentDataV372(existingPayments, remotePayments);
  saveAllPaymentsToSheet(username, mergedPayments);
  
  // Auto cleanup LUNAS data (sisa <= 0)
  const cleanupResult = cleanupLunasDataV372(username);
  
  // Auto cleanup log setelah merge (opsional)
  try {
    handleCleanOldLogs(username, MAX_LOG_LINES);
  } catch(e) {}
  
  addSyncLog(username, "mergeBackup", `Merged ${mergedBons.length} bons, ${mergedPayments.length} payments, cleaned ${cleanupResult.count} customers`, VERSION);
  
  return ContentService
    .createTextOutput(JSON.stringify({ 
      success: true, 
      message: "Data merged successfully v37.2",
      cleanedUp: cleanupResult,
      version: VERSION
    }))
    .setMimeType(ContentService.MimeType.JSON);
}

function handleRestoreV372(username) {
  const allBons = getAllBonsByUser(username);
  const allPayments = getAllPaymentsByUser(username);
  
  // Filter hanya active bons (sisa utang > 0) - TIDAK MENGAMBIL DATA LUNAS
  const activeBons = filterActiveBonsV372(allBons, allPayments);
  
  return ContentService
    .createTextOutput(JSON.stringify({ 
      success: true, 
      semuaBon: activeBons,
      pembayaran: allPayments,
      lastModified: new Date().toISOString(),
      version: VERSION
    }))
    .setMimeType(ContentService.MimeType.JSON);
}

function handleSyncBonV372(username, bonData) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME_BONS);
  const existingRow = findRowByUniqueId(sheet, bonData.uniqueId);
  
  if (existingRow) {
    // Update existing
    const rowNum = existingRow;
    sheet.getRange(rowNum, 3).setValue(bonData.namaPelanggan);
    sheet.getRange(rowNum, 4).setValue(bonData.total);
    sheet.getRange(rowNum, 5).setValue(JSON.stringify(bonData.items));
    sheet.getRange(rowNum, 6).setValue(bonData.waktu);
    sheet.getRange(rowNum, 7).setValue(bonData.lastModified);
    sheet.getRange(rowNum, 8).setValue(true);
    sheet.getRange(rowNum, 9).setValue(VERSION);
  } else {
    // Insert new
    sheet.appendRow([
      bonData.uniqueId,
      username,
      bonData.namaPelanggan,
      bonData.total,
      JSON.stringify(bonData.items),
      bonData.waktu,
      bonData.lastModified,
      true,
      VERSION
    ]);
  }
  
  // Auto cleanup after sync
  cleanupLunasDataV372(username);
  
  addSyncLog(username, "syncBon", `Synced bon ${bonData.uniqueId}`, VERSION);
  
  return ContentService
    .createTextOutput(JSON.stringify({ success: true, version: VERSION }))
    .setMimeType(ContentService.MimeType.JSON);
}

function handleSyncPaymentV372(username, paymentData) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME_PAYMENTS);
  const existingRow = findRowByUniqueId(sheet, paymentData.uniqueId);
  
  if (existingRow) {
    const rowNum = existingRow;
    sheet.getRange(rowNum, 3).setValue(paymentData.namaPelanggan);
    sheet.getRange(rowNum, 4).setValue(paymentData.jumlah);
    sheet.getRange(rowNum, 5).setValue(paymentData.waktu);
    sheet.getRange(rowNum, 6).setValue(paymentData.lastModified);
    sheet.getRange(rowNum, 7).setValue(VERSION);
  } else {
    sheet.appendRow([
      paymentData.uniqueId,
      username,
      paymentData.namaPelanggan,
      paymentData.jumlah,
      paymentData.waktu,
      paymentData.lastModified,
      VERSION
    ]);
  }
  
  // Auto cleanup after payment sync
  cleanupLunasDataV372(username);
  
  addSyncLog(username, "syncPayment", `Synced payment ${paymentData.uniqueId}`, VERSION);
  
  return ContentService
    .createTextOutput(JSON.stringify({ success: true, version: VERSION }))
    .setMimeType(ContentService.MimeType.JSON);
}

function handleSyncUserAuth(authData) {
  const action = authData.subAction || authData.action;
  const username = authData.username.toLowerCase();
  const userData = authData.userData;
  
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME_USERS);
  const existingRow = findUserRow(sheet, username);
  
  if (action === "register") {
    if (existingRow) {
      return ContentService
        .createTextOutput(JSON.stringify({ success: false, error: "User already exists" }))
        .setMimeType(ContentService.MimeType.JSON);
    }
    
    sheet.appendRow([
      username,
      userData.password,
      userData.security.pet,
      userData.security.hobby,
      userData.security.food,
      userData.registeredAt || new Date().toISOString(),
      new Date().toISOString(),
      true,
      VERSION
    ]);
    
    addSyncLog(username, "register", "New user registered", VERSION);
    
  } else if (action === "resetPassword") {
    if (!existingRow) {
      return ContentService
        .createTextOutput(JSON.stringify({ success: false, error: "User not found" }))
        .setMimeType(ContentService.MimeType.JSON);
    }
    
    const rowNum = existingRow;
    sheet.getRange(rowNum, 2).setValue(userData.password);
    sheet.getRange(rowNum, 7).setValue(new Date().toISOString());
    sheet.getRange(rowNum, 9).setValue(VERSION);
    
    addSyncLog(username, "resetPassword", "Password reset", VERSION);
  }
  
  return ContentService
    .createTextOutput(JSON.stringify({ success: true, version: VERSION }))
    .setMimeType(ContentService.MimeType.JSON);
}

function handleGetUserAuth(username) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME_USERS);
  const data = sheet.getDataRange().getValues();
  
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === username.toLowerCase()) {
      const userData = {
        username: data[i][0],
        password: data[i][1],
        security: {
          pet: data[i][2],
          hobby: data[i][3],
          food: data[i][4]
        },
        registeredAt: data[i][5],
        lastLogin: data[i][6],
        isActive: data[i][7]
      };
      
      return ContentService
        .createTextOutput(JSON.stringify({ success: true, userData: userData, version: VERSION }))
        .setMimeType(ContentService.MimeType.JSON);
    }
  }
  
  return ContentService
    .createTextOutput(JSON.stringify({ success: false, message: "User not found" }))
    .setMimeType(ContentService.MimeType.JSON);
}

function handleCleanupLunasOtomatisV372(username) {
  const result = cleanupLunasDataV372(username);
  return ContentService
    .createTextOutput(JSON.stringify({ success: true, result: result, version: VERSION }))
    .setMimeType(ContentService.MimeType.JSON);
}

// ==================== HANDLE DELETE CUSTOMER DATA (INSTANT DELETE) ====================
function handleDeleteCustomerDataV372(username, customerName) {
  try {
    const result = deleteCustomerDataPermanenV372(username, customerName);
    addSyncLog(username, "deleteCustomer", `Deleted customer: ${customerName} (${result.deletedBons} bons, ${result.deletedPayments} payments)`, VERSION);
    
    return ContentService
      .createTextOutput(JSON.stringify({ 
        success: true, 
        result: result,
        message: `Customer ${customerName} deleted permanently`,
        version: VERSION
      }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (error) {
    return ContentService
      .createTextOutput(JSON.stringify({ success: false, error: error.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// ==================== DELETE CUSTOMER DATA PERMANENTLY ====================
function deleteCustomerDataPermanenV372(username, customerName) {
  const customerKey = customerName.toLowerCase();
  let deletedBons = 0;
  let deletedPayments = 0;
  
  // Hapus dari sheet BONS
  const sheetBons = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME_BONS);
  const bonsData = sheetBons.getDataRange().getValues();
  
  for (let i = bonsData.length - 1; i >= 1; i--) {
    const bonUsername = bonsData[i][1];
    const namaPelanggan = bonsData[i][2];
    if (bonUsername === username && namaPelanggan && namaPelanggan.toLowerCase() === customerKey) {
      sheetBons.deleteRow(i + 1);
      deletedBons++;
    }
  }
  
  // Hapus dari sheet PAYMENTS
  const sheetPayments = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME_PAYMENTS);
  const paymentsData = sheetPayments.getDataRange().getValues();
  
  for (let i = paymentsData.length - 1; i >= 1; i--) {
    const paymentUsername = paymentsData[i][1];
    const namaPelanggan = paymentsData[i][2];
    if (paymentUsername === username && namaPelanggan && namaPelanggan.toLowerCase() === customerKey) {
      sheetPayments.deleteRow(i + 1);
      deletedPayments++;
    }
  }
  
  return {
    deletedBons: deletedBons,
    deletedPayments: deletedPayments,
    customerName: customerName,
    timestamp: new Date().toISOString(),
    version: VERSION
  };
}

// ==================== CLEANUP LUNAS DATA (SISA <= 0) ====================
function cleanupLunasDataV372(username) {
  const bons = getAllBonsByUser(username);
  const payments = getAllPaymentsByUser(username);
  
  // Calculate per customer - PERHITUNGAN AKURAT
  const paymentSum = {};
  payments.forEach(p => {
    const key = p.namaPelanggan.toLowerCase();
    paymentSum[key] = (paymentSum[key] || 0) + p.jumlah;
  });
  
  const outstanding = {};
  bons.forEach(bon => {
    const key = bon.namaPelanggan.toLowerCase();
    outstanding[key] = (outstanding[key] || 0) + bon.total;
  });
  
  const customersToDelete = [];
  
  // Cari pelanggan dengan sisa utang <= 0 (termasuk kelebihan bayar)
  for (const key in outstanding) {
    const totalOwed = outstanding[key];
    const totalPaid = paymentSum[key] || 0;
    const sisa = totalOwed - totalPaid;
    
    // HAPUS jika sisa <= 0 (LUNAS total atau kelebihan bayar)
    if (sisa <= 0) {
      customersToDelete.push(key);
    }
  }
  
  // Hapus customers yang punya payment tapi tidak punya bon (data orphan)
  for (const key in paymentSum) {
    if (!outstanding[key] && paymentSum[key] > 0) {
      if (!customersToDelete.includes(key)) {
        customersToDelete.push(key);
      }
    }
  }
  
  let deletedBons = 0;
  let deletedPayments = 0;
  
  // Hapus BONS untuk customers yang ditandai
  const sheetBons = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME_BONS);
  const bonsData = sheetBons.getDataRange().getValues();
  
  for (let i = bonsData.length - 1; i >= 1; i--) {
    const bonUsername = bonsData[i][1];
    const namaPelanggan = bonsData[i][2];
    if (bonUsername === username && namaPelanggan && customersToDelete.includes(namaPelanggan.toLowerCase())) {
      sheetBons.deleteRow(i + 1);
      deletedBons++;
    }
  }
  
  // Hapus PAYMENTS untuk customers yang ditandai
  const sheetPayments = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME_PAYMENTS);
  const paymentsData = sheetPayments.getDataRange().getValues();
  
  for (let i = paymentsData.length - 1; i >= 1; i--) {
    const paymentUsername = paymentsData[i][1];
    const namaPelanggan = paymentsData[i][2];
    if (paymentUsername === username && namaPelanggan && customersToDelete.includes(namaPelanggan.toLowerCase())) {
      sheetPayments.deleteRow(i + 1);
      deletedPayments++;
    }
  }
  
  if (customersToDelete.length > 0) {
    addSyncLog(username, "cleanupLunas", `Deleted ${customersToDelete.length} customers (${deletedBons} bons, ${deletedPayments} payments) - LUNAS total`, VERSION);
  }
  
  return {
    deletedCustomers: customersToDelete,
    count: customersToDelete.length,
    deletedBons: deletedBons,
    deletedPayments: deletedPayments,
    timestamp: new Date().toISOString(),
    version: VERSION
  };
}

// ==================== HELPER FUNCTIONS ====================

function getAllBonsByUser(username) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME_BONS);
  const data = sheet.getDataRange().getValues();
  const result = [];
  
  for (let i = 1; i < data.length; i++) {
    if (data[i][1] === username && data[i][7] !== false) {
      let items = [];
      try {
        items = JSON.parse(data[i][4] || "[]");
      } catch(e) {
        items = [];
      }
      result.push({
        uniqueId: data[i][0],
        username: data[i][1],
        namaPelanggan: data[i][2],
        total: data[i][3],
        items: items,
        waktu: data[i][5],
        lastModified: data[i][6],
        isActive: data[i][7],
        version: data[i][8] || "37.0"
      });
    }
  }
  return result;
}

function getAllPaymentsByUser(username) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME_PAYMENTS);
  const data = sheet.getDataRange().getValues();
  const result = [];
  
  for (let i = 1; i < data.length; i++) {
    if (data[i][1] === username) {
      result.push({
        uniqueId: data[i][0],
        username: data[i][1],
        namaPelanggan: data[i][2],
        jumlah: data[i][3],
        waktu: data[i][4],
        lastModified: data[i][5],
        version: data[i][6] || "37.0"
      });
    }
  }
  return result;
}

function filterActiveBonsV372(bons, payments) {
  // Group payments by customer
  const paymentSum = {};
  payments.forEach(p => {
    const key = p.namaPelanggan.toLowerCase();
    paymentSum[key] = (paymentSum[key] || 0) + p.jumlah;
  });
  
  // Calculate outstanding per customer
  const outstanding = {};
  bons.forEach(bon => {
    const key = bon.namaPelanggan.toLowerCase();
    outstanding[key] = (outstanding[key] || 0) + bon.total;
  });
  
  const remaining = {};
  for (const key in outstanding) {
    const paid = paymentSum[key] || 0;
    const sisa = outstanding[key] - paid;
    remaining[key] = sisa;
  }
  
  // Filter bons: keep if customer still has outstanding debt (sisa > 0)
  // TIDAK MENGAMBIL DATA LUNAS
  const activeBons = [];
  bons.forEach(bon => {
    const key = bon.namaPelanggan.toLowerCase();
    const sisa = remaining[key];
    
    // Keep bon only if sisa > 0 (still owe money)
    if (sisa > 0) {
      activeBons.push(bon);
    }
  });
  
  return activeBons;
}

function saveAllBonsToSheet(username, bons) {
  // Clear existing bons for this user
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME_BONS);
  const data = sheet.getDataRange().getValues();
  
  // Mark all existing as inactive first
  for (let i = 1; i < data.length; i++) {
    if (data[i][1] === username) {
      sheet.getRange(i + 1, 8).setValue(false);
    }
  }
  
  // Write new/updated bons
  bons.forEach(bon => {
    const existingRow = findRowByUniqueId(sheet, bon.uniqueId);
    if (existingRow) {
      const rowNum = existingRow;
      sheet.getRange(rowNum, 3).setValue(bon.namaPelanggan);
      sheet.getRange(rowNum, 4).setValue(bon.total);
      sheet.getRange(rowNum, 5).setValue(JSON.stringify(bon.items));
      sheet.getRange(rowNum, 6).setValue(bon.waktu);
      sheet.getRange(rowNum, 7).setValue(bon.lastModified);
      sheet.getRange(rowNum, 8).setValue(true);
      sheet.getRange(rowNum, 9).setValue(VERSION);
    } else {
      sheet.appendRow([
        bon.uniqueId,
        username,
        bon.namaPelanggan,
        bon.total,
        JSON.stringify(bon.items),
        bon.waktu,
        bon.lastModified,
        true,
        VERSION
      ]);
    }
  });
}

function saveAllPaymentsToSheet(username, payments) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME_PAYMENTS);
  const data = sheet.getDataRange().getValues();
  
  // Delete all existing payments for this user
  for (let i = data.length - 1; i >= 1; i--) {
    if (data[i][1] === username) {
      sheet.deleteRow(i + 1);
    }
  }
  
  // Write all payments
  payments.forEach(payment => {
    sheet.appendRow([
      payment.uniqueId,
      username,
      payment.namaPelanggan,
      payment.jumlah,
      payment.waktu,
      payment.lastModified,
      VERSION
    ]);
  });
}

function findRowByUniqueId(sheet, uniqueId) {
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === uniqueId) {
      return i + 1;
    }
  }
  return null;
}

function findUserRow(sheet, username) {
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === username) {
      return i + 1;
    }
  }
  return null;
}

function mergeBonDataV372(existingBons, remoteBons) {
  const bonMap = new Map();
  
  existingBons.forEach(bon => {
    if (bon && bon.uniqueId) {
      bonMap.set(bon.uniqueId, bon);
    }
  });
  
  remoteBons.forEach(bon => {
    if (bon && bon.uniqueId) {
      const existing = bonMap.get(bon.uniqueId);
      if (!existing) {
        bonMap.set(bon.uniqueId, bon);
      } else {
        const existingTime = new Date(existing.lastModified || 0).getTime();
        const remoteTime = new Date(bon.lastModified || 0).getTime();
        if (remoteTime > existingTime) {
          bonMap.set(bon.uniqueId, bon);
        }
      }
    }
  });
  
  return Array.from(bonMap.values());
}

function mergePaymentDataV372(existingPayments, remotePayments) {
  const paymentMap = new Map();
  
  existingPayments.forEach(payment => {
    if (payment && payment.uniqueId) {
      paymentMap.set(payment.uniqueId, payment);
    }
  });
  
  remotePayments.forEach(payment => {
    if (payment && payment.uniqueId) {
      const existing = paymentMap.get(payment.uniqueId);
      if (!existing) {
        paymentMap.set(payment.uniqueId, payment);
      } else {
        const existingTime = new Date(existing.lastModified || 0).getTime();
        const remoteTime = new Date(payment.lastModified || 0).getTime();
        if (remoteTime > existingTime) {
          paymentMap.set(payment.uniqueId, payment);
        }
      }
    }
  });
  
  return Array.from(paymentMap.values());
}

function addSyncLog(username, action, details, version) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME_SYNC_LOG);
    sheet.appendRow([
      new Date().toISOString(),
      username || "system",
      action,
      details,
      version || VERSION
    ]);
  } catch(e) {
    console.error("Failed to add sync log:", e);
  }
}
