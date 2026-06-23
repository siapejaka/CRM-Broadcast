const SPREADSHEET_ID = '1xAlCqJuJevCxCk58XX0gDgaFyZvwoNs4hedcqOI8hcI';

function doGet(e) {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('CRM Contact Hub')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, maximum-scale=1, user-scalable=no')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// Recursively converts dates to ISO strings and returns plain, JSON-safe objects
function serializeForClient(value) {
  if (value instanceof Date) {
    return value.toISOString();
  }
  if (Array.isArray(value)) {
    return value.map(function(item) {
      return serializeForClient(item);
    });
  }
  if (value !== null && typeof value === 'object') {
    var obj = {};
    for (var key in value) {
      if (value.hasOwnProperty(key)) {
        obj[key] = serializeForClient(value[key]);
      }
    }
    return obj;
  }
  return value;
}

// Helper to generate SHA-256 Hash for Passwords
function hashPassword(password) {
  if (!password) return '';
  var digest = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, password, Utilities.Charset.UTF_8);
  var signature = [];
  for (var i = 0; i < digest.length; i++) {
    var byteVal = digest[i];
    if (byteVal < 0) byteVal += 256;
    var byteString = byteVal.toString(16);
    if (byteString.length == 1) byteString = '0' + byteString;
    signature.push(byteString);
  }
  return signature.join('');
}

// Ensures database structure exists and is healthy
function setupDatabase() {
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheets = {
      'Leads': ['ID', 'Name', 'WhatsApp', 'Tags', 'Owner WhatsApp', 'Created At', 'Last Broadcast At', 'Number Status', 'Last Number Check At', 'Inactive Reason', 'Email'],
      'Segments': ['ID', 'Name', 'Include Tags', 'Exclude Tags', 'Created At'],
      'Templates': ['ID', 'Name', 'Content', 'Product ID', 'Image URL', 'Created At'],
      'Broadcasts': ['ID', 'Name', 'Segment ID', 'Template ID', 'Status', 'Sent Count', 'Total Count', 'Created At', 'Scheduled At', 'Delay Contact Min', 'Delay Contact Max', 'Delay Bubble Min', 'Delay Bubble Max', 'Throttle Count', 'Throttle Minutes', 'Add Tags', 'Remove Tags', 'Sender CS', 'Target Lead IDs'],
      'BroadcastLogs': ['ID', 'Broadcast ID', 'Lead ID', 'Lead Name', 'WhatsApp', 'Status', 'Message ID', 'Error Message', 'Sent At', 'Delivered At', 'Read At', 'Scheduled At', 'Message Content', 'Sender CS', 'Retry Count', 'Max Retry', 'Processing Started At', 'Updated At', 'Image URL'],
      'TagRules': ['ID', 'Trigger Tag', 'Tags to Add', 'Tags to Remove', 'Assign CS', 'Created At'],
      'CSNumbers': ['ID', 'CS Name', 'WhatsApp Number', 'Session Name', 'Created At'],
      'Products': ['ID', 'Product Name', 'Price', 'Description', 'Created At'],
      'Tags': ['ID', 'Tag Name', 'Created At'],
      'Users': ['ID', 'Username', 'Password', 'Role', 'Linked CS Number', 'Google Contact Label', 'Created At']
    };
    
    // Auto-Migrate Broadcasts sheet if it has the old 17-column format
    var broadcastsSheet = ss.getSheetByName('Broadcasts');
    if (broadcastsSheet && broadcastsSheet.getLastColumn() > 0) {
      var currentHeaders = broadcastsSheet.getRange(1, 1, 1, broadcastsSheet.getLastColumn()).getValues()[0];
      if (currentHeaders.length === 17 || currentHeaders.indexOf('Delay Contact') !== -1) {
        var data = broadcastsSheet.getDataRange().getValues();
        var migratedData = [];
        migratedData.push(sheets['Broadcasts']); // Gunakan header baru
        
        for (var i = 1; i < data.length; i++) {
          var row = data[i];
          if (row[0]) { // Hanya migrasikan baris yang tidak kosong
            var id = row[0];
            var name = row[1];
            var segmentId = row[2];
            var templateId = row[3];
            var status = row[4];
            var sentCount = row[5];
            var totalCount = row[6];
            var createdAt = row[7];
            var scheduledAt = row[8];
            var delayContact = row[9] ? row[9].toString() : '15-30';
            var delayBubble = row[10] ? row[10].toString() : '0-2';
            var throttleCount = row[11] || 0;
            var throttleMinutes = row[12] || 0;
            var addTags = row[13] || '';
            var removeTags = row[14] || '';
            var senderCS = row[15] || '';
            var targetLeadIds = row[16] || '';
            
            // Konversi Delay Contact lama ke Min-Max
            var dcMin = 15, dcMax = 30;
            if (delayContact.indexOf('-') !== -1) {
              var dcParts = delayContact.split('-');
              dcMin = parseInt(dcParts[0]) || 15;
              dcMax = parseInt(dcParts[1]) || dcMin;
            } else {
              dcMin = parseInt(delayContact) || 15;
              dcMax = dcMin;
            }
            
            // Konversi Delay Bubble lama ke Min-Max
            var dbMin = 0, dbMax = 2;
            if (delayBubble.indexOf('-') !== -1) {
              var dbParts = delayBubble.split('-');
              dbMin = parseInt(dbParts[0]) || 0;
              dbMax = parseInt(dbParts[1]) || dbMin;
            } else {
              dbMin = parseInt(delayBubble) || 0;
              dbMax = dbMin;
            }
            
            migratedData.push([
              id, name, segmentId, templateId, status, sentCount, totalCount, createdAt, scheduledAt,
              dcMin, dcMax, dbMin, dbMax, throttleCount, throttleMinutes, addTags, removeTags, senderCS, targetLeadIds
            ]);
          }
        }
        
        broadcastsSheet.clearContents();
        broadcastsSheet.getRange(1, 1, migratedData.length, sheets['Broadcasts'].length).setValues(migratedData);
      }
    }
    
    for (var sheetName in sheets) {
      var sheet = ss.getSheetByName(sheetName);
      if (!sheet) {
        sheet = ss.insertSheet(sheetName);
      }
      sheet.getRange(1, 1, 1, sheets[sheetName].length).setValues([sheets[sheetName]]);
    }
    
    // Seed default administrator if Users sheet is empty
    var usersSheet = ss.getSheetByName('Users');
    var usersData = getSheetDataAsObjects(usersSheet);
    if (usersData.length === 0) {
      var defaultAdminId = 'U' + Utilities.getUuid().substring(0, 8).toUpperCase();
      var hashedDefaultPassword = hashPassword('admin');
      usersSheet.appendRow([defaultAdminId, 'admin', hashedDefaultPassword, 'Super Admin', '', 'CRM Leads', new Date().toISOString()]);
    }

    return { ok: true, message: 'Database schema successfully generated/verified.' };
  } catch (err) {
    return { ok: false, message: err.toString() };
  }
}

// Helper to normalize phone numbers to global 62... format
function normalizePhoneNumber(phone) {
  if (!phone) return '';
  var clean = phone.toString().replace(/\D/g, ''); // Ambil angka saja
  if (clean.indexOf('0') === 0) {
    clean = '62' + clean.substring(1);
  } else if (clean.indexOf('62') !== 0 && clean.length > 0) {
    clean = '62' + clean;
  }
  return clean;
}

// Helper to process Spintax {kata1|kata2|kata3} into a randomly selected word
function spinText(text) {
  if (!text) return '';
  var regex = /\{([^{}]+)\}/g;
  var match;
  // Loop terus menerus untuk mendukung spintax bersarang (nested spintax) jika ada
  while ((match = regex.exec(text)) !== null) {
    var fullMatch = match[0];
    var optionsStr = match[1];
    if (optionsStr.indexOf('|') !== -1) {
      var options = optionsStr.split('|');
      var randomOption = options[Math.floor(Math.random() * options.length)].trim();
      text = text.replace(fullMatch, randomOption);
      // Reset regex index agar pencarian diulang dari awal teks baru
      regex.lastIndex = 0;
    }
  }
  return text;
}

// Helper to locate a row by ID
function findRowIndexById(sheet, id) {
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (data[i][0].toString() === id.toString()) {
      return i + 1;
    }
  }
  return -1;
}

// Helper to convert CS Phone Number to clean WAHA Session Name (e.g. "Adit Coffiy" -> "Adit_Coffiy")
function getSessionNameByPhone(phone) {
  var cleanPhone = normalizePhoneNumber(phone);
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName('CSNumbers');
  if (sheet) {
    var data = getSheetDataAsObjects(sheet);
    for (var i = 0; i < data.length; i++) {
      if (normalizePhoneNumber(data[i].whatsapp_number) === cleanPhone) {
        if (data[i].session_name) {
          return data[i].session_name.toString().trim();
        }
        var formattedName = data[i].cs_name.toString().trim()
          .replace(/[^a-zA-Z0-9]/g, '_') // Ganti spasi & karakter spesial dengan underscore
          .replace(/_+/g, '_');          // Gabung underscore ganda jika ada
        return formattedName;
      }
    }
  }
  return cleanPhone; // Fallback ke nomor jika CS tidak ditemukan
}

// Helper to convert sheet rows into key-value JS objects
function getSheetDataAsObjects(sheet) {
  var range = sheet.getDataRange();
  var values = range.getValues();
  if (values.length <= 1) return [];
  
  var headers = values[0].map(function(h) {
    return h.toString().toLowerCase().replace(/\s+/g, '_');
  });
  
  var result = [];
  for (var i = 1; i < values.length; i++) {
    var obj = {};
    for (var j = 0; j < headers.length; j++) {
      obj[headers[j]] = values[i][j];
    }
    result.push(obj);
  }
  return result;
}

// User Authentication Handler
function loginUser(username, password) {
  try {
    setupDatabase();
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheetByName('Users');
    var users = getSheetDataAsObjects(sheet);
    
    var cleanUsername = username.toString().trim().toLowerCase();
    var hashedInput = hashPassword(password);
    
    for (var i = 0; i < users.length; i++) {
      if (users[i].username.toString().toLowerCase() === cleanUsername && users[i].password === hashedInput) {
        return {
          ok: true,
          user: {
            id: users[i].id,
            username: users[i].username,
            role: users[i].role,
            linked_cs_number: users[i].linked_cs_number || '',
            google_contact_label: users[i].google_contact_label || 'CRM Leads'
          }
        };
      }
    }
    return { ok: false, message: 'Username atau Password salah!' };
  } catch (err) {
    return { ok: false, message: err.toString() };
  }
}

// Google Login Token Verification & Session Handler
function loginWithGoogleToken(accessToken) {
  try {
    setupDatabase();
    var url = 'https://www.googleapis.com/oauth2/v3/userinfo';
    var response = UrlFetchApp.fetch(url, {
      headers: { 'Authorization': 'Authorization Bearer ' + accessToken },
      muteHttpExceptions: true
    });
    if (response.getResponseCode() !== 200) {
      return { ok: false, message: 'Gagal memverifikasi akun Google dengan server.' };
    }
    var userInfo = JSON.parse(response.getContentText());
    var email = userInfo.email.toLowerCase();
    
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheetByName('Users');
    var users = getSheetDataAsObjects(sheet);
    
    for (var i = 0; i < users.length; i++) {
      if (users[i].username.toString().toLowerCase() === email) {
        return {
          ok: true,
          user: {
            id: users[i].id,
            username: users[i].username,
            role: users[i].role,
            linked_cs_number: users[i].linked_cs_number || '',
            google_contact_label: users[i].google_contact_label || 'CRM Leads'
          }
        };
      }
    }
    return { ok: false, message: 'Email Google "' + email + '" belum didaftarkan di Manajemen User oleh Admin.' };
  } catch (err) {
    return { ok: false, message: err.toString() };
  }
}

// User Management CRUD
function saveUser(userObj) {
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheetByName('Users');
    var now = new Date().toISOString();
    
    var cleanUsername = userObj.username.trim().toLowerCase();
    var contactLabel = userObj.google_contact_label ? userObj.google_contact_label.trim() : 'CRM Leads';
    
    // Validasi duplikasi username
    var existing = getSheetDataAsObjects(sheet);
    for (var i = 0; i < existing.length; i++) {
      if (existing[i].username.toLowerCase() === cleanUsername) {
        if (!userObj.id || existing[i].id.toString() !== userObj.id.toString()) {
          throw new Error('Username/Email "' + userObj.username + '" sudah digunakan oleh user lain.');
        }
      }
    }
    
    if (!userObj.id) {
      var id = 'U' + Utilities.getUuid().substring(0, 8).toUpperCase();
      var hashedPass = hashPassword(userObj.password || '123456'); // Default password jika kosong
      sheet.appendRow([id, userObj.username, hashedPass, userObj.role, userObj.linked_cs_number || '', contactLabel, now]);
    } else {
      var rowIndex = findRowIndexById(sheet, userObj.id);
      if (rowIndex === -1) throw new Error('User tidak ditemukan');
      
      // Update data dasar
      sheet.getRange(rowIndex, 2).setValue(userObj.username);
      sheet.getRange(rowIndex, 4, 1, 3).setValues([[userObj.role, userObj.linked_cs_number || '', contactLabel]]);
      
      // Update password hanya jika diisi baru
      if (userObj.password && userObj.password.trim() !== '') {
        sheet.getRange(rowIndex, 3).setValue(hashPassword(userObj.password));
      }
    }
    return { ok: true };
  } catch (err) {
    return { ok: false, message: err.toString() };
  }
}

function deleteUser(id) {
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheetByName('Users');
    var rowIndex = findRowIndexById(sheet, id);
    if (rowIndex === -1) throw new Error('User tidak ditemukan');
    
    // Proteksi agar admin tidak menghapus akun dirinya sendiri secara tidak sengaja
    var usernameToDelete = sheet.getRange(rowIndex, 2).getValue().toString();
    if (usernameToDelete === 'admin') {
      throw new Error('User master "admin" tidak boleh dihapus demi keamanan sistem.');
    }
    
    sheet.deleteRow(rowIndex);
    return { ok: true };
  } catch (err) {
    return { ok: false, message: err.toString() };
  }
}

// Retrieves complete CRM context with Server-Side Role Filtering
function getDashboardData(activeUser) {
  try {
    setupDatabase(); // Selalu jalankan verifikasi & migrasi database saat dashboard dibuka
    syncAllBroadcastCounters(); // Sinkronisasi otomatis data log dari n8n ke tabel campaign
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    
    var leadsSheet = ss.getSheetByName('Leads');
    var segmentsSheet = ss.getSheetByName('Segments');
    var templatesSheet = ss.getSheetByName('Templates');
    var broadcastsSheet = ss.getSheetByName('Broadcasts');
    var csNumbersSheet = ss.getSheetByName('CSNumbers');
    var productsSheet = ss.getSheetByName('Products');
    var tagsSheet = ss.getSheetByName('Tags');
    var broadcastLogsSheet = ss.getSheetByName('BroadcastLogs');
    var tagRulesSheet = ss.getSheetByName('TagRules');
    var usersSheet = ss.getSheetByName('Users');
    
    var leads = getSheetDataAsObjects(leadsSheet);
    var segments = getSheetDataAsObjects(segmentsSheet);
    var templates = getSheetDataAsObjects(templatesSheet);
    var broadcasts = getSheetDataAsObjects(broadcastsSheet);
    var tagRules = getSheetDataAsObjects(tagRulesSheet);
    var csNumbers = getSheetDataAsObjects(csNumbersSheet);
    var products = getSheetDataAsObjects(productsSheet);
    var tags = getSheetDataAsObjects(tagsSheet);
    var broadcastLogs = getSheetDataAsObjects(broadcastLogsSheet);
    var users = getSheetDataAsObjects(usersSheet);
    var wahaSettings = getWahaSettings();
    
    // --- SERVER-SIDE DATA FILTERING BASED ON ROLE ---
    if (activeUser && activeUser.role === 'CS Agent') {
      var linkedCSStr = String(activeUser.linked_cs_number || '');
      var linkedCSNumbers = linkedCSStr.split(',').map(function(num) {
        return normalizePhoneNumber(num.trim());
      }).filter(Boolean);
      
      // 1. Saring Kontak (Leads) milik CS tersebut saja (jika nomor owner cocok dengan salah satu nomor yang dikaitkan)
      leads = leads.filter(function(lead) {
        return linkedCSNumbers.indexOf(normalizePhoneNumber(lead.owner_whatsapp)) !== -1;
      });
      
      // 2. Saring Broadcast Campaign yang dikirim oleh CS tersebut saja
      broadcasts = broadcasts.filter(function(bc) {
        return linkedCSNumbers.indexOf(normalizePhoneNumber(bc.sender_cs)) !== -1;
      });
      
      // 3. Saring Log Detail yang dikirim oleh CS tersebut saja
      broadcastLogs = broadcastLogs.filter(function(log) {
        return linkedCSNumbers.indexOf(normalizePhoneNumber(log.sender_cs)) !== -1;
      });
    }
    
    return {
      ok: true,
      data: serializeForClient({
        leads: leads,
        segments: segments,
        templates: templates,
        broadcasts: broadcasts,
        tag_rules: tagRules,
        cs_numbers: csNumbers,
        products: products,
        tags: tags,
        broadcast_logs: broadcastLogs,
        waha_settings: wahaSettings,
        users: users
      })
    };
  } catch (err) {
    return { ok: false, message: err.toString() };
  }
}

// Helper to process Tag Automation Rules
function processTagAutomations(tagsStr) {
  if (!tagsStr) return '';
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var ruleSheet = ss.getSheetByName('TagRules');
    if (!ruleSheet) return tagsStr;
    
    var rules = getSheetDataAsObjects(ruleSheet);
    var currentTags = tagsStr.split(',').map(function(t) { return t.trim(); }).filter(Boolean);
    var lowerTags = currentTags.map(function(t) { return t.toLowerCase(); });
    
    var changed = true;
    var iterations = 0;
    // Limit to 5 cascading loops to prevent infinite logic loops
    while (changed && iterations < 5) {
      changed = false;
      iterations++;
      
      for (var i = 0; i < rules.length; i++) {
        var rule = rules[i];
        var trigger = rule.trigger_tag.trim().toLowerCase();
        
        // If contact has the trigger tag
        if (lowerTags.indexOf(trigger) !== -1) {
          // Process Tags to Add
          if (rule.tags_to_add) {
            var toAdd = rule.tags_to_add.split(',').map(function(t) { return t.trim(); }).filter(Boolean);
            for (var j = 0; j < toAdd.length; j++) {
              var addTag = toAdd[j];
              var addTagLower = addTag.toLowerCase();
              if (lowerTags.indexOf(addTagLower) === -1) {
                currentTags.push(addTag);
                lowerTags.push(addTagLower);
                changed = true;
              }
            }
          }
          // Process Tags to Remove
          if (rule.tags_to_remove) {
            var toRemove = rule.tags_to_remove.split(',').map(function(t) { return t.trim().toLowerCase(); }).filter(Boolean);
            for (var j = 0; j < toRemove.length; j++) {
              var removeTagLower = toRemove[j];
              var index = lowerTags.indexOf(removeTagLower);
              if (index !== -1) {
                currentTags.splice(index, 1);
                lowerTags.splice(index, 1);
                changed = true;
              }
            }
          }
        }
      }
    }
    return currentTags.join(', ');
  } catch (e) {
    return tagsStr; // Fallback to raw tags on error
  }
}

// Helper to auto-assign CS based on tags
function processCSAutoAssign(tagsStr, currentOwner) {
  if (!tagsStr) return currentOwner;
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var ruleSheet = ss.getSheetByName('TagRules');
    if (!ruleSheet) return currentOwner;
    
    var rules = getSheetDataAsObjects(ruleSheet);
    var currentTags = tagsStr.split(',').map(function(t) { return t.trim().toLowerCase(); }).filter(Boolean);
    
    for (var i = 0; i < rules.length; i++) {
      var rule = rules[i];
      var trigger = rule.trigger_tag.trim().toLowerCase();
      if (currentTags.indexOf(trigger) !== -1 && rule.assign_cs) {
        return rule.assign_cs.trim(); // Returns the assigned CS WhatsApp number
      }
    }
  } catch (e) {
    // Fallback
  }
  return currentOwner;
}

// CRUD for Leads
function saveLead(lead) {
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheetByName('Leads');
    var now = new Date().toISOString();
    
    var processedTags = processTagAutomations(lead.tags);
    var finalOwner = processCSAutoAssign(processedTags, lead.owner_whatsapp);
    
    var formattedWhatsapp = normalizePhoneNumber(lead.whatsapp);
    var formattedOwner = finalOwner ? normalizePhoneNumber(finalOwner) : '';
    var lookupWA = formattedWhatsapp;
    
    if (!lead.id) {
      // Proteksi Duplikat Manual
      var data = sheet.getDataRange().getValues();
      for (var i = 1; i < data.length; i++) {
        if (normalizePhoneNumber(data[i][2]) === lookupWA) {
          throw new Error('Nomor WhatsApp ' + lookupWA + ' sudah terdaftar dengan nama ' + data[i][1]);
        }
      }
      var id = 'L' + Utilities.getUuid().substring(0, 8).toUpperCase();
      sheet.appendRow([id, lead.name, formattedWhatsapp, processedTags, formattedOwner, now, '', 'Unchecked', '', '']);
    } else {
      var rowIndex = findRowIndexById(sheet, lead.id);
      if (rowIndex === -1) throw new Error('Lead not found');
      // Pastikan edit tidak mengubah nomor menjadi nomor milik orang lain yang sudah terdaftar
      var data = sheet.getDataRange().getValues();
      for (var i = 1; i < data.length; i++) {
        if (i + 1 !== rowIndex && normalizePhoneNumber(data[i][2]) === lookupWA) {
          throw new Error('Gagal update: Nomor WhatsApp ' + lookupWA + ' sudah digunakan oleh kontak lain (' + data[i][1] + ')');
        }
      }
      sheet.getRange(rowIndex, 2, 1, 4).setValues([[lead.name, formattedWhatsapp, processedTags, formattedOwner]]);
    }
    return { ok: true };
  } catch (err) {
    return { ok: false, message: err.toString() };
  }
}

function bulkUpdateLeadTags(leadIds, actionType, tagsStr) {
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheetByName('Leads');
    var data = sheet.getDataRange().getValues();
    var tagsToApply = tagsStr.split(',').map(function(t) { return t.trim(); }).filter(Boolean);
    
    for (var i = 1; i < data.length; i++) {
      var rowId = data[i][0].toString();
      if (leadIds.indexOf(rowId) !== -1) {
        var currentTags = data[i][3].toString().split(',').map(function(t) { return t.trim(); }).filter(Boolean);
        var currentOwner = data[i][4].toString();
        
        if (actionType === 'ADD') {
          tagsToApply.forEach(function(t) {
            if (currentTags.map(function(ct) { return ct.toLowerCase(); }).indexOf(t.toLowerCase()) === -1) {
              currentTags.push(t);
            }
          });
        } else if (actionType === 'REMOVE') {
          var toRemoveLower = tagsToApply.map(function(t) { return t.toLowerCase(); });
          currentTags = currentTags.filter(function(t) {
            return toRemoveLower.indexOf(t.toLowerCase()) === -1;
          });
        }
        
        var finalizedTags = processTagAutomations(currentTags.join(', '));
        var finalizedOwner = processCSAutoAssign(finalizedTags, currentOwner);
        sheet.getRange(i + 1, 4, 1, 2).setValues([[finalizedTags, finalizedOwner]]);
      }
    }
    return { ok: true };
  } catch (err) {
    return { ok: false, message: err.toString() };
  }
}

function bulkDeleteLeads(leadIds) {
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheetByName('Leads');
    var data = sheet.getDataRange().getValues();
    
    // Delete from bottom to top to preserve correct row indices
    for (var i = data.length - 1; i >= 1; i--) {
      var rowId = data[i][0].toString();
      if (leadIds.indexOf(rowId) !== -1) {
        sheet.deleteRow(i + 1);
      }
    }
    return { ok: true };
  } catch (err) {
    return { ok: false, message: err.toString() };
  }
}

function deleteLead(id) {
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheetByName('Leads');
    var rowIndex = findRowIndexById(sheet, id);
    if (rowIndex === -1) throw new Error('Lead not found');
    sheet.deleteRow(rowIndex);
    return { ok: true };
  } catch (err) {
    return { ok: false, message: err.toString() };
  }
}

// Bulk CSV/Spreadsheet Import (Optimized Bulk Write)
function importLeads(leads) {
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheetByName('Leads');
    if (!sheet) {
      setupDatabase();
      sheet = ss.getSheetByName('Leads');
    }
    
    var data = sheet.getDataRange().getValues();
    var headers = data[0];
    
    // Build map of existing normalized WhatsApp numbers to row index (0-based in data array)
    var existingMap = {};
    for (var i = 1; i < data.length; i++) {
      var wa = normalizePhoneNumber(data[i][2]);
      if (wa) {
        existingMap[wa] = i;
      }
    }
    
    var now = new Date().toISOString();
    
    // LOAD RULES ONCE TO PREVENT TIMEOUTS/LAG (OPTIMIZATION)
    var rules = [];
    var ruleSheet = ss.getSheetByName('TagRules');
    if (ruleSheet) {
      rules = getSheetDataAsObjects(ruleSheet);
    }
    
    for (var j = 0; j < leads.length; j++) {
      var lead = leads[j];
      var rawWA = lead.whatsapp ? lead.whatsapp.toString().trim() : '';
      if (!rawWA) continue;
      
      var wa = normalizePhoneNumber(rawWA);
      
      // Local processing using cached rules
      var processedTags = processTagAutomationsWithRules(lead.tags || '', rules);
      var finalOwner = processCSAutoAssignWithRules(processedTags, lead.owner_whatsapp || '', rules);
      
      var formattedOwner = finalOwner ? normalizePhoneNumber(finalOwner) : '';
      
      var waForSheet = wa;
      var ownerForSheet = formattedOwner ? formattedOwner : '';
      
      if (existingMap[wa]) {
        var idx = existingMap[wa];
        data[idx][1] = lead.name || 'No Name';
        data[idx][2] = waForSheet;
        data[idx][3] = processedTags;
        data[idx][4] = ownerForSheet;
        // Keep other columns intact
      } else {
        var id = 'L' + Utilities.getUuid().substring(0, 8).toUpperCase();
        data.push([
          id,
          lead.name || 'No Name',
          waForSheet,
          processedTags,
          ownerForSheet,
          now,          // Created At
          '',           // Last Broadcast At
          'Unchecked',  // Number Status
          '',           // Last Number Check At
          ''            // Inactive Reason
        ]);
        existingMap[wa] = data.length - 1;
      }
    }
    
    sheet.clearContents();
    sheet.getRange(1, 1, data.length, headers.length).setValues(data);
    
    return { ok: true, message: 'Berhasil memproses ' + leads.length + ' kontak dalam sekejap!' };
  } catch (err) {
    return { ok: false, message: err.toString() };
  }
}

// Optimized Helper using cached rules
function processTagAutomationsWithRules(tagsStr, rules) {
  if (!tagsStr || rules.length === 0) return tagsStr;
  try {
    var currentTags = tagsStr.split(',').map(function(t) { return t.trim(); }).filter(Boolean);
    var lowerTags = currentTags.map(function(t) { return t.toLowerCase(); });
    
    var changed = true;
    var iterations = 0;
    while (changed && iterations < 5) {
      changed = false;
      iterations++;
      
      for (var i = 0; i < rules.length; i++) {
        var rule = rules[i];
        var trigger = rule.trigger_tag.trim().toLowerCase();
        
        if (lowerTags.indexOf(trigger) !== -1) {
          if (rule.tags_to_add) {
            var toAdd = rule.tags_to_add.split(',').map(function(t) { return t.trim(); }).filter(Boolean);
            for (var j = 0; j < toAdd.length; j++) {
              var addTag = toAdd[j];
              var addTagLower = addTag.toLowerCase();
              if (lowerTags.indexOf(addTagLower) === -1) {
                currentTags.push(addTag);
                lowerTags.push(addTagLower);
                changed = true;
              }
            }
          }
          if (rule.tags_to_remove) {
            var toRemove = rule.tags_to_remove.split(',').map(function(t) { return t.trim().toLowerCase(); }).filter(Boolean);
            for (var j = 0; j < toRemove.length; j++) {
              var removeTagLower = toRemove[j];
              var index = lowerTags.indexOf(removeTagLower);
              if (index !== -1) {
                currentTags.splice(index, 1);
                lowerTags.splice(index, 1);
                changed = true;
              }
            }
          }
        }
      }
    }
    return currentTags.join(', ');
  } catch (e) {
    return tagsStr;
  }
}

// Optimized Helper using cached rules
function processCSAutoAssignWithRules(tagsStr, currentOwner, rules) {
  if (!tagsStr || rules.length === 0) return currentOwner;
  try {
    var currentTags = tagsStr.split(',').map(function(t) { return t.trim().toLowerCase(); }).filter(Boolean);
    for (var i = 0; i < rules.length; i++) {
      var rule = rules[i];
      var trigger = rule.trigger_tag.trim().toLowerCase();
      if (currentTags.indexOf(trigger) !== -1 && rule.assign_cs) {
        return rule.assign_cs.trim();
      }
    }
  } catch (e) {}
  return currentOwner;
}

// CRUD for Segments
function saveSegment(segment) {
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheetByName('Segments');
    var now = new Date().toISOString();
    
    if (!segment.id) {
      var id = 'S' + Utilities.getUuid().substring(0, 8).toUpperCase();
      sheet.appendRow([id, segment.name, segment.include_tags, segment.exclude_tags, now]);
    } else {
      var rowIndex = findRowIndexById(sheet, segment.id);
      if (rowIndex === -1) throw new Error('Segment not found');
      sheet.getRange(rowIndex, 2, 1, 3).setValues([[segment.name, segment.include_tags, segment.exclude_tags]]);
    }
    return { ok: true };
  } catch (err) {
    return { ok: false, message: err.toString() };
  }
}

function deleteSegment(id) {
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheetByName('Segments');
    var rowIndex = findRowIndexById(sheet, id);
    if (rowIndex === -1) throw new Error('Segment not found');
    sheet.deleteRow(rowIndex);
    return { ok: true };
  } catch (err) {
    return { ok: false, message: err.toString() };
  }
}

// CRUD for Templates
function saveTemplate(template) {
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheetByName('Templates');
    var now = new Date().toISOString();
    
    var productId = template.product_id || '';
    var imageUrl = template.image_url || '';
    
    if (!template.id) {
      var id = 'T' + Utilities.getUuid().substring(0, 8).toUpperCase();
      sheet.appendRow([id, template.name, template.content, productId, imageUrl, now]);
    } else {
      var rowIndex = findRowIndexById(sheet, template.id);
      if (rowIndex === -1) throw new Error('Template not found');
      sheet.getRange(rowIndex, 2, 1, 4).setValues([[template.name, template.content, productId, imageUrl]]);
    }
    return { ok: true };
  } catch (err) {
    return { ok: false, message: err.toString() };
  }
}

function deleteTemplate(id) {
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheetByName('Templates');
    var rowIndex = findRowIndexById(sheet, id);
    if (rowIndex === -1) throw new Error('Template not found');
    sheet.deleteRow(rowIndex);
    return { ok: true };
  } catch (err) {
    return { ok: false, message: err.toString() };
  }
}

// CRUD for Broadcasts
function saveBroadcast(broadcast) {
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheetByName('Broadcasts');
    var now = new Date().toISOString();
    
    var senderCS = broadcast.sender_cs ? normalizePhoneNumber(broadcast.sender_cs) : '';
    var targetLeadIds = broadcast.target_lead_ids || '';

    if (!broadcast.id) {
      var id = 'B' + Utilities.getUuid().substring(0, 8).toUpperCase();

      // Ambil template dan produk untuk denormalisasi pesan langsung di Apps Script
      var templatesSheet = ss.getSheetByName('Templates');
      var templates = getSheetDataAsObjects(templatesSheet);
      var template = templates.find(function(t) { return t.id === broadcast.template_id; });
      var rawContent = template ? template.content : '';
      
      var productsSheet = ss.getSheetByName('Products');
      var products = getSheetDataAsObjects(productsSheet);
      var associatedProduct = template ? products.find(function(p) { return p.id === template.product_id; }) : null;
      var pName = associatedProduct ? associatedProduct.product_name : '';
      var pPrice = associatedProduct ? 'Rp ' + Number(associatedProduct.price).toLocaleString('id-ID') : '';

      // Hitung daftar target penerima
      var leadsSheet = ss.getSheetByName('Leads');
      var leadsData = getSheetDataAsObjects(leadsSheet);
      var targetIds = targetLeadIds.split(',').map(function(t) { return t.trim(); }).filter(Boolean);
      
      if (targetIds.length === 0 && broadcast.segment_id) {
        var segmentsSheet = ss.getSheetByName('Segments');
        var segments = getSheetDataAsObjects(segmentsSheet);
        var segment = segments.find(function(s) { return s.id === broadcast.segment_id; });
        if (segment) {
          var inc = segment.include_tags ? segment.include_tags.split(',').map(function(t) { return t.trim().toLowerCase(); }).filter(Boolean) : [];
          var exc = segment.exclude_tags ? segment.exclude_tags.split(',').map(function(t) { return t.trim().toLowerCase(); }).filter(Boolean) : [];
          targetIds = leadsData.filter(function(lead) {
            var leadTags = lead.tags ? lead.tags.split(',').map(function(t) { return t.trim().toLowerCase(); }).filter(Boolean) : [];
            if (leadTags.some(function(t) { return exc.indexOf(t) !== -1; })) return false;
            if (inc.length > 0 && !leadTags.some(function(t) { return inc.indexOf(t) !== -1; })) return false;
            return true;
          }).map(function(l) { return l.id; });
        }
      }

      // Pecah teks berdasarkan [BUBBLE] untuk menghitung total baris log sesungguhnya
      var bubbles = rawContent.split('[BUBBLE]').map(function(b) { return b.trim(); }).filter(Boolean);
      if (bubbles.length === 0) bubbles = [rawContent];
      
      var calculatedTotalCount = targetIds.length * bubbles.length;

              sheet.appendRow([
                id, 
                broadcast.name, 
                broadcast.segment_id, 
                broadcast.template_id, 
                'Pending', 
                0, 
                calculatedTotalCount, // Total baris log yang akan dibuat
                now,
                broadcast.scheduled_at || '',
                broadcast.delay_contact_min || 15, // Kolom J: Delay Contact Min
                broadcast.delay_contact_max || 30, // Kolom K: Delay Contact Max
                broadcast.delay_bubble_min || 0,   // Kolom L: Delay Bubble Min
                broadcast.delay_bubble_max || 0,   // Kolom M: Delay Bubble Max
                broadcast.throttle_count || 0,      // Kolom N: Throttle Count
                broadcast.throttle_minutes || 0,    // Kolom O: Throttle Minutes
                broadcast.add_tags || '',           // Kolom P: Add Tags
                broadcast.remove_tags || '',        // Kolom Q: Remove Tags
                senderCS,                           // Kolom R: Sender CS
                targetLeadIds                       // Kolom S: Target Lead IDs
              ]);

      // Eksekusi perubahan tag langsung ke database kontak saat submit broadcast (Instant)
      if (broadcast.add_tags) {
        bulkUpdateLeadTags(targetIds, 'ADD', broadcast.add_tags);
      }
      if (broadcast.remove_tags) {
        bulkUpdateLeadTags(targetIds, 'REMOVE', broadcast.remove_tags);
      }

      // Buat log antrean penerima detail (BroadcastLogs) secara otomatis per bubble
      var logSheet = ss.getSheetByName('BroadcastLogs');
      var baseTime = broadcast.scheduled_at ? new Date(broadcast.scheduled_at) : new Date();
      
      // Ambil parameter jeda kontak & jeda bubble secara terpisah
      var delayContactMin = parseInt(broadcast.delay_contact_min || 15);
      var delayContactMax = parseInt(broadcast.delay_contact_max || 30);
      var delayBubbleMin = parseInt(broadcast.delay_bubble_min || 2);
      var delayBubbleMax = parseInt(broadcast.delay_bubble_max || 5);
      
      var throttleCount = parseInt(broadcast.throttle_count || 0);
      var throttleMinutes = parseInt(broadcast.throttle_minutes || 0);
      
      var runningTime = baseTime.getTime();

      targetIds.forEach(function(leadId, leadIndex) {
        var lead = leadsData.find(function(l) { return l.id === leadId; });
        if (lead) {
          
          // 1. JEDA ANTAR KONTAK: Jika beralih ke kontak baru (bukan kontak pertama), tambahkan jeda kontak acak
          if (leadIndex > 0) {
            var contactDelay = delayContactMin;
            if (delayContactMax > delayContactMin) {
              contactDelay = Math.floor(Math.random() * (delayContactMax - delayContactMin + 1)) + delayContactMin;
            }
            runningTime += contactDelay * 1000;

            // Istirahat Sesi (Throttle): Jika batas sesi tercapai, tambahkan waktu istirahat panjang
            if (throttleCount > 0 && throttleMinutes > 0 && leadIndex % throttleCount === 0) {
              runningTime += throttleMinutes * 60 * 1000;
            }
          }

          // Proses setiap bubble untuk kontak saat ini
          bubbles.forEach(function(bubbleText, bubbleIndex) {
            
            // 2. JEDA ANTAR BUBBLE: Jika ini adalah bubble ke-2 atau seterusnya untuk kontak yang sama
            if (bubbleIndex > 0) {
              var bubbleDelay = delayBubbleMin;
              if (delayBubbleMax > delayBubbleMin) {
                bubbleDelay = Math.floor(Math.random() * (delayBubbleMax - delayBubbleMin + 1)) + delayBubbleMin;
              }
              runningTime += bubbleDelay * 1000;
            }

            var scheduledTimeStr = new Date(runningTime).toISOString();
            var logId = 'LOG' + Utilities.getUuid().substring(0, 8).toUpperCase();
            
            // Ekstrak tag gambar [IMAGE:url] khusus pada bubble ini jika ada
            var bubbleImage = '';
            var cleanBubbleText = bubbleText;
            if (bubbleText.indexOf('[IMAGE:') === 0) {
              var closeIndex = bubbleText.indexOf(']');
              if (closeIndex !== -1) {
                bubbleImage = bubbleText.substring(7, closeIndex);
                cleanBubbleText = bubbleText.substring(closeIndex + 1).trim();
              }
            }

            var personalizedMessage = cleanBubbleText
              .replace(/\{\{name\}\}/g, lead.name)
              .replace(/\{\{whatsapp\}\}/g, lead.whatsapp)
              .replace(/\{\{product_name\}\}/g, pName)
              .replace(/\{\{product_price\}\}/g, pPrice);
            
            // Acak kata menggunakan sistem Spintax sebelum masuk ke database log
            var spunMessage = spinText(personalizedMessage);
            
            logSheet.appendRow([
              logId,
              id, // Broadcast ID
              lead.id,
              lead.name,
              lead.whatsapp,
              'Pending',
              '', // Message ID
              '', // Error Message
              '', // Sent At
              '', // Delivered At
              '', // Read At
              scheduledTimeStr, // Scheduled At
              spunMessage, // Message Content
              senderCS, // Sender CS
              0, // Retry Count
              3, // Max Retry
              '', // Processing Started At
              now, // Updated At
              bubbleImage // Image URL (Kolom S) - Menggunakan gambar spesifik dari bubble ini
            ]);
          });
        }
      });
    } else {
      var rowIndex = findRowIndexById(sheet, broadcast.id);
      if (rowIndex === -1) throw new Error('Broadcast not found');
      sheet.getRange(rowIndex, 2, 1, 18).setValues([[
        broadcast.name, 
        broadcast.segment_id, 
        broadcast.template_id, 
        broadcast.status, 
        broadcast.sent_count, 
        broadcast.total_count,
        broadcast.created_at || now,
        broadcast.scheduled_at || '',
        broadcast.delay_contact_min || 0,
        broadcast.delay_contact_max || 0,
        broadcast.delay_bubble_min || 0,
        broadcast.delay_bubble_max || 0,
        broadcast.throttle_count || 0,
        broadcast.throttle_minutes || 0,
        broadcast.add_tags || '',
        broadcast.remove_tags || '',
        senderCS,
        targetLeadIds
      ]]);
    }
    return { ok: true };
  } catch (err) {
    return { ok: false, message: err.toString() };
  }
}

function executePostBroadcastAction(leadWhatsapp, addTagsStr, removeTagsStr) {
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheetByName('Leads');
    var data = sheet.getDataRange().getValues();
    var waCol = 2; // WhatsApp column index
    var tagsCol = 3; // Tags column index
    var lastBcCol = 6; // Last Broadcast At column index (0-based column 7)
    var now = new Date().toISOString();
    
    var lookupWA = normalizePhoneNumber(leadWhatsapp);
    for (var i = 1; i < data.length; i++) {
      if (normalizePhoneNumber(data[i][waCol]) === lookupWA) {
        var currentTags = data[i][tagsCol].toString().split(',').map(function(t) { return t.trim(); }).filter(Boolean);
        
        // Process remove tags
        if (removeTagsStr) {
          var toRemove = removeTagsStr.split(',').map(function(t) { return t.trim().toLowerCase(); }).filter(Boolean);
          currentTags = currentTags.filter(function(t) {
            return toRemove.indexOf(t.toLowerCase()) === -1;
          });
        }
        
        // Process add tags
        if (addTagsStr) {
          var toAdd = addTagsStr.split(',').map(function(t) { return t.trim(); }).filter(Boolean);
          toAdd.forEach(function(t) {
            var lowerTags = currentTags.map(function(ct) { return ct.toLowerCase(); });
            if (lowerTags.indexOf(t.toLowerCase()) === -1) {
              currentTags.push(t);
            }
          });
        }
        
        var finalizedTags = processTagAutomations(currentTags.join(', '));
        var currentOwner = data[i][tagsCol + 1].toString();
        var finalizedOwner = processCSAutoAssign(finalizedTags, currentOwner);
        
        // Update Tags, CS Owner, and Last Broadcast At timestamp
        sheet.getRange(i + 1, tagsCol + 1, 1, 2).setValues([[finalizedTags, finalizedOwner]]);
        sheet.getRange(i + 1, lastBcCol + 1).setValue(now);
        break;
      }
    }
    return { ok: true };
  } catch (err) {
    return { ok: false, message: err.toString() };
  }
}

function deleteBroadcast(id) {
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheetByName('Broadcasts');
    var rowIndex = findRowIndexById(sheet, id);
    if (rowIndex === -1) throw new Error('Broadcast not found');
    sheet.deleteRow(rowIndex);

    // Hapus juga log detail terkait di BroadcastLogs agar tidak memenuhi database
    var logSheet = ss.getSheetByName('BroadcastLogs');
    var logData = logSheet.getDataRange().getValues();
    for (var i = logData.length - 1; i >= 1; i--) {
      if (logData[i][1].toString() === id.toString()) {
        logSheet.deleteRow(i + 1);
      }
    }
    return { ok: true };
  } catch (err) {
    return { ok: false, message: err.toString() };
  }
}

// Fungsi Endpoint untuk n8n mengunci baris sebelum diproses (mencegah double send)
function lockBroadcastLogForProcessing(logId) {
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheetByName('BroadcastLogs');
    var rowIndex = findRowIndexById(sheet, logId);
    if (rowIndex === -1) throw new Error('Log row tidak ditemukan');
    
    var now = new Date().toISOString();
    // Set Status ke 'Processing' (Kolom F) dan Processing Started At (Kolom Q) serta Updated At (Kolom R)
    sheet.getRange(rowIndex, 6).setValue('Processing');
    sheet.getRange(rowIndex, 17).setValue(now);
    sheet.getRange(rowIndex, 18).setValue(now);
    
    return { ok: true, message: 'Log row locked successfully for processing.' };
  } catch (err) {
    return { ok: false, message: err.toString() };
  }
}

// Fungsi Endpoint untuk n8n memperbarui status log per nomor
function updateBroadcastLogStatus(logId, status, messageId, errorMessage) {
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheetByName('BroadcastLogs');
    var rowIndex = findRowIndexById(sheet, logId);
    if (rowIndex === -1) throw new Error('Log row tidak ditemukan');
    
    var now = new Date().toISOString();
    sheet.getRange(rowIndex, 6).setValue(status); // Kolom F: Status
    sheet.getRange(rowIndex, 18).setValue(now);   // Kolom R: Updated At
    
    if (messageId) {
      sheet.getRange(rowIndex, 7).setValue(messageId); // Kolom G: Message ID
    }
    if (errorMessage) {
      sheet.getRange(rowIndex, 8).setValue(errorMessage); // Kolom H: Error Message
    }
    
    // Ambil data Lead ID terkait untuk pembaruan otomatis status nomor kontak
    var leadId = sheet.getRange(rowIndex, 3).getValue().toString();
    var leadsSheet = ss.getSheetByName('Leads');
    var leadRowIndex = findRowIndexById(leadsSheet, leadId);

    if (status === 'Sent' || status === 'Success') {
      sheet.getRange(rowIndex, 9).setValue(now); // Kolom I: Sent At
      sheet.getRange(rowIndex, 17).setValue(''); // Reset Processing Started At karena sudah sukses
      
      // Otomatis memperbarui kolom 'Last Broadcast At' & 'Number Status' di tab Leads & Jalankan Aksi Tag Pasca Terkirim
      try {
        var leadWhatsapp = sheet.getRange(rowIndex, 5).getValue().toString();
        var broadcastId = sheet.getRange(rowIndex, 2).getValue().toString();
        
        if (leadRowIndex !== -1) {
          leadsSheet.getRange(leadRowIndex, 7).setValue(now); // Kolom G: Last Broadcast At
          leadsSheet.getRange(leadRowIndex, 8).setValue('Sent'); // Kolom H: Number Status
          leadsSheet.getRange(leadRowIndex, 9).setValue(now); // Kolom I: Last Number Check At
          leadsSheet.getRange(leadRowIndex, 10).setValue(''); // Kolom J: Inactive Reason (Reset karena aktif)
        }

        // Ambil aturan Tambah/Hapus tag dari tabel Broadcasts induk
        var broadcastsSheet = ss.getSheetByName('Broadcasts');
        var bcRowIndex = findRowIndexById(broadcastsSheet, broadcastId);
        if (bcRowIndex !== -1) {
          var addTagsStr = broadcastsSheet.getRange(bcRowIndex, 14).getValue().toString(); // Kolom N: Add Tags
          var removeTagsStr = broadcastsSheet.getRange(bcRowIndex, 15).getValue().toString(); // Kolom O: Remove Tags
          
          // Jalankan penambahan/penghapusan tag & alokasi CS otomatis
          if (addTagsStr || removeTagsStr) {
            executePostBroadcastAction(leadWhatsapp, addTagsStr, removeTagsStr);
          }
        }
      } catch (e) {
        // Abaikan jika terjadi error pencatatan agar tidak menghentikan proses utama
      }
    } else if (status === 'Invalid Number') {
      sheet.getRange(rowIndex, 17).setValue(''); // Reset Processing lock
      
      // Otomatis tandai nomor mati permanen di tabel Leads
      if (leadRowIndex !== -1) {
        leadsSheet.getRange(leadRowIndex, 8).setValue('Invalid Number'); // Kolom H: Number Status
        leadsSheet.getRange(leadRowIndex, 9).setValue(now); // Kolom I: Last Number Check At
        leadsSheet.getRange(leadRowIndex, 10).setValue(errorMessage || 'Nomor tidak terdaftar di WhatsApp'); // Kolom J: Inactive Reason
      }
    } else if (status === 'Delivered') {
      sheet.getRange(rowIndex, 10).setValue(now); // Kolom J: Delivered At
      
      // Otomatis tandai nomor aktif di tabel Leads karena sukses diterima
      if (leadRowIndex !== -1) {
        leadsSheet.getRange(leadRowIndex, 8).setValue('Delivered'); // Kolom H: Number Status
        leadsSheet.getRange(leadRowIndex, 9).setValue(now); // Kolom I: Last Number Check At
        leadsSheet.getRange(leadRowIndex, 10).setValue(''); // Kolom J: Inactive Reason (Reset)
      }
    } else if (status === 'Read') {
      sheet.getRange(rowIndex, 11).setValue(now); // Kolom K: Read At
      
      // Otomatis tandai nomor aktif di tabel Leads karena sukses dibaca
      if (leadRowIndex !== -1) {
        leadsSheet.getRange(leadRowIndex, 8).setValue('Read'); // Kolom H: Number Status
        leadsSheet.getRange(leadRowIndex, 9).setValue(now); // Kolom I: Last Number Check At
        leadsSheet.getRange(leadRowIndex, 10).setValue(''); // Kolom J: Inactive Reason (Reset)
      }
    } else if (status === 'Failed') {
      sheet.getRange(rowIndex, 17).setValue(''); // Reset Processing lock agar bisa dicoba lagi jika dibutuhkan
      
      // Aturan Retry Logic: Ambil Retry Count (Kolom O) dan Max Retry (Kolom P)
      var retryCountRange = sheet.getRange(rowIndex, 15);
      var maxRetry = parseInt(sheet.getRange(rowIndex, 16).getValue() || 3);
      var currentRetry = parseInt(retryCountRange.getValue() || 0);
      
      if (currentRetry < maxRetry) {
        var nextRetry = currentRetry + 1;
        retryCountRange.setValue(nextRetry);
        sheet.getRange(rowIndex, 6).setValue('Pending'); // Kembalikan ke Pending agar n8n bisa mengulanginya
        sheet.getRange(rowIndex, 8).setValue(errorMessage + ' (Retry #' + nextRetry + ')');
      } else {
        sheet.getRange(rowIndex, 6).setValue('Failed'); // Permanen gagal jika sudah melebihi Max Retry
        if (leadRowIndex !== -1) {
          leadsSheet.getRange(leadRowIndex, 8).setValue('Failed');
          leadsSheet.getRange(leadRowIndex, 9).setValue(now);
          leadsSheet.getRange(leadRowIndex, 10).setValue(errorMessage || 'Gagal terkirim setelah retry maksimal');
        }
      }
    }
    
    // Hitung ulang progress di tabel Broadcasts utama
    var broadcastId = sheet.getRange(rowIndex, 2).getValue().toString();
    recalculateBroadcastCounters(broadcastId);
    
    return { ok: true };
  } catch (err) {
    return { ok: false, message: err.toString() };
  }
}

// Fungsi Endpoint untuk n8n memperbarui status log menggunakan Message ID (Sangat pas untuk Webhook Read/Delivered WhatsApp)
function updateBroadcastLogByMessageId(messageId, status) {
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheetByName('BroadcastLogs');
    var data = sheet.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      if (data[i][6].toString() === messageId.toString()) {
        var logId = data[i][0].toString();
        return updateBroadcastLogStatus(logId, status, messageId, '');
      }
    }
    return { ok: false, message: 'Message ID tidak ditemukan di database' };
  } catch (err) {
    return { ok: false, message: err.toString() };
  }
}

// Menghitung ulang persentase & jumlah sukses terkirim pada tabel Broadcasts
function recalculateBroadcastCounters(broadcastId) {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var logSheet = ss.getSheetByName('BroadcastLogs');
  var logs = getSheetDataAsObjects(logSheet);
  
  var sentCount = logs.filter(function(l) {
    return l.broadcast_id === broadcastId && (l.status === 'Sent' || l.status === 'Delivered' || l.status === 'Read' || l.status === 'Success');
  }).length;
  
  var processedCount = logs.filter(function(l) {
    return l.broadcast_id === broadcastId && ['Sent', 'Delivered', 'Read', 'Success', 'Failed', 'Invalid Number', 'Cancelled'].indexOf(l.status) !== -1;
  }).length;
  
  var bcSheet = ss.getSheetByName('Broadcasts');
  var rowIndex = findRowIndexById(bcSheet, broadcastId);
  if (rowIndex !== -1) {
    bcSheet.getRange(rowIndex, 6).setValue(sentCount); // Kolom F: Sent Count
    
    var totalCount = parseInt(bcSheet.getRange(rowIndex, 7).getValue() || 0);
    if (processedCount >= totalCount && totalCount > 0) {
      bcSheet.getRange(rowIndex, 5).setValue('Completed'); // Kolom E: Status
    } else if (processedCount > 0) {
      bcSheet.getRange(rowIndex, 5).setValue('Running'); // Kolom E: Status
    } else {
      bcSheet.getRange(rowIndex, 5).setValue('Pending'); // Kolom E: Status
    }
  }
}

// Sinkronisasi otomatis seluruh campaign yang belum selesai berdasarkan data log asli dari n8n
function syncAllBroadcastCounters() {
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var bcSheet = ss.getSheetByName('Broadcasts');
    var logSheet = ss.getSheetByName('BroadcastLogs');
    if (!bcSheet || !logSheet) return;
    
    var broadcasts = getSheetDataAsObjects(bcSheet);
    var logs = getSheetDataAsObjects(logSheet);
    
    for (var i = 0; i < broadcasts.length; i++) {
      var bc = broadcasts[i];
      if (bc.status !== 'Completed') {
        var bcId = bc.id;
        var sentCount = logs.filter(function(l) {
          return l.broadcast_id === bcId && (l.status === 'Sent' || l.status === 'Delivered' || l.status === 'Read' || l.status === 'Success');
        }).length;
        
        var processedCount = logs.filter(function(l) {
          return l.broadcast_id === bcId && ['Sent', 'Delivered', 'Read', 'Success', 'Failed', 'Invalid Number', 'Cancelled'].indexOf(l.status) !== -1;
        }).length;
        
        var rowIndex = findRowIndexById(bcSheet, bcId);
        if (rowIndex !== -1) {
          bcSheet.getRange(rowIndex, 6).setValue(sentCount); // Kolom F: Sent Count
          
          var totalCount = parseInt(bc.total_count || 0);
          if (processedCount >= totalCount && totalCount > 0) {
            bcSheet.getRange(rowIndex, 5).setValue('Completed');
          } else if (processedCount > 0) {
            bcSheet.getRange(rowIndex, 5).setValue('Running');
          } else {
            bcSheet.getRange(rowIndex, 5).setValue('Pending');
          }
        }
      }
    }
  } catch (e) {
    // Diamkan error agar tidak mengganggu proses load utama dashboard
  }
}

// CRUD for Tag Rules (Otomatisasi Tag)
function saveTagRule(rule) {
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheetByName('TagRules');
    if (!sheet) {
      setupDatabase();
      sheet = ss.getSheetByName('TagRules');
    }
    var now = new Date().toISOString();
    
    var tagsToAdd = rule.tags_to_add || '';
    var tagsToRemove = rule.tags_to_remove || '';
    var assignCS = rule.assign_cs || '';
    
    if (!rule.id) {
      var id = 'R' + Utilities.getUuid().substring(0, 8).toUpperCase();
      sheet.appendRow([id, rule.trigger_tag, tagsToAdd, tagsToRemove, assignCS, now]);
    } else {
      var rowIndex = findRowIndexById(sheet, rule.id);
      if (rowIndex === -1) throw new Error('Rule not found');
      sheet.getRange(rowIndex, 2, 1, 4).setValues([[rule.trigger_tag, tagsToAdd, tagsToRemove, assignCS]]);
    }
    return { ok: true };
  } catch (err) {
    return { ok: false, message: err.toString() };
  }
}

function deleteTagRule(id) {
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheetByName('TagRules');
    var rowIndex = findRowIndexById(sheet, id);
    if (rowIndex === -1) throw new Error('Rule not found');
    sheet.deleteRow(rowIndex);
    return { ok: true };
  } catch (err) {
    return { ok: false, message: err.toString() };
  }
}

// CRUD for CS Numbers (Kelola CS)
function saveCSNumber(cs) {
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheetByName('CSNumbers');
    if (!sheet) {
      setupDatabase();
      sheet = ss.getSheetByName('CSNumbers');
    }
    var now = new Date().toISOString();
    
    var formattedWA = normalizePhoneNumber(cs.whatsapp_number);
    var sessionName = cs.session_name ? cs.session_name.trim() : '';
    
    // Validasi: 1 nomor CS hanya boleh dikelola oleh 1 entitas CS saja
    var data = getSheetDataAsObjects(sheet);
    var cleanInputWA = normalizePhoneNumber(cs.whatsapp_number);
    for (var i = 0; i < data.length; i++) {
      var existingWA = normalizePhoneNumber(data[i].whatsapp_number);
      if (existingWA === cleanInputWA) {
        if (!cs.id || data[i].id.toString() !== cs.id.toString()) {
          throw new Error('Nomor WhatsApp ini sudah digunakan oleh CS lain (' + data[i].cs_name + ')');
        }
      }
    }
    
    if (!cs.id) {
      var id = 'CS' + Utilities.getUuid().substring(0, 8).toUpperCase();
      sheet.appendRow([id, cs.cs_name, formattedWA, sessionName, now]);
    } else {
      var rowIndex = findRowIndexById(sheet, cs.id);
      if (rowIndex === -1) throw new Error('CS Number tidak ditemukan');
      sheet.getRange(rowIndex, 2, 1, 3).setValues([[cs.cs_name, formattedWA, sessionName]]);
    }
    return { ok: true };
  } catch (err) {
    return { ok: false, message: err.toString() };
  }
}

// CRUD for Products (Kelola Produk)
function saveProduct(product) {
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheetByName('Products');
    if (!sheet) {
      setupDatabase();
      sheet = ss.getSheetByName('Products');
    }
    var now = new Date().toISOString();
    
    if (!product.id) {
      var id = 'P' + Utilities.getUuid().substring(0, 8).toUpperCase();
      sheet.appendRow([id, product.product_name, product.price, product.description, now]);
    } else {
      var rowIndex = findRowIndexById(sheet, product.id);
      if (rowIndex === -1) throw new Error('Produk tidak ditemukan');
      sheet.getRange(rowIndex, 2, 1, 3).setValues([[product.product_name, product.price, product.description]]);
    }
    return { ok: true };
  } catch (err) {
    return { ok: false, message: err.toString() };
  }
}

function deleteProduct(id) {
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheetByName('Products');
    var rowIndex = findRowIndexById(sheet, id);
    if (rowIndex === -1) throw new Error('Produk tidak ditemukan');
    sheet.deleteRow(rowIndex);
    return { ok: true };
  } catch (err) {
    return { ok: false, message: err.toString() };
  }
}

function deleteCSNumber(id) {
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheetByName('CSNumbers');
    var rowIndex = findRowIndexById(sheet, id);
    if (rowIndex === -1) throw new Error('CS Number tidak ditemukan');
    sheet.deleteRow(rowIndex);
    return { ok: true };
  } catch (err) {
    return { ok: false, message: err.toString() };
  }
}

// CRUD for Tags (Kelola Tags)
function saveTag(tagObj) {
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheetByName('Tags');
    if (!sheet) {
      setupDatabase();
      sheet = ss.getSheetByName('Tags');
    }
    var now = new Date().toISOString();
    var newTagName = tagObj.tag_name.trim();

    if (!tagObj.id) {
      var existing = getSheetDataAsObjects(sheet);
      for (var i = 0; i < existing.length; i++) {
        if (existing[i].tag_name.toLowerCase() === newTagName.toLowerCase()) {
          throw new Error('Tag "' + newTagName + '" sudah terdaftar.');
        }
      }
      var id = 'TAG' + Utilities.getUuid().substring(0, 8).toUpperCase();
      sheet.appendRow([id, newTagName, now]);
    } else {
      var rowIndex = findRowIndexById(sheet, tagObj.id);
      if (rowIndex === -1) throw new Error('Tag tidak ditemukan');
      
      var oldTagName = sheet.getRange(rowIndex, 2).getValue().toString().trim();
      if (oldTagName.toLowerCase() !== newTagName.toLowerCase()) {
        sheet.getRange(rowIndex, 2).setValue(newTagName);
        propagateTagRename(oldTagName, newTagName);
      }
    }
    return { ok: true };
  } catch (err) {
    return { ok: false, message: err.toString() };
  }
}

function propagateTagRename(oldTag, newTag) {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var oldLower = oldTag.toLowerCase();
  
  // 1. Leads
  var leadsSheet = ss.getSheetByName('Leads');
  if (leadsSheet) {
    var leadsData = leadsSheet.getDataRange().getValues();
    for (var i = 1; i < leadsData.length; i++) {
      var tagsStr = leadsData[i][3].toString();
      if (tagsStr) {
        var tagsArr = tagsStr.split(',').map(function(t) { return t.trim(); });
        var updated = false;
        for (var j = 0; j < tagsArr.length; j++) {
          if (tagsArr[j].toLowerCase() === oldLower) {
            tagsArr[j] = newTag;
            updated = true;
          }
        }
        if (updated) {
          leadsSheet.getRange(i + 1, 4).setValue(tagsArr.join(', '));
        }
      }
    }
  }

  // 2. Segments
  var segmentsSheet = ss.getSheetByName('Segments');
  if (segmentsSheet) {
    var segData = segmentsSheet.getDataRange().getValues();
    for (var i = 1; i < segData.length; i++) {
      var incStr = segData[i][2].toString();
      var excStr = segData[i][3].toString();
      var updateInc = false, updateExc = false;

      if (incStr) {
        var incArr = incStr.split(',').map(function(t) { return t.trim(); });
        for (var j = 0; j < incArr.length; j++) {
          if (incArr[j].toLowerCase() === oldLower) { incArr[j] = newTag; updateInc = true; }
        }
      }
      if (excStr) {
        var excArr = excStr.split(',').map(function(t) { return t.trim(); });
        for (var j = 0; j < excArr.length; j++) {
          if (excArr[j].toLowerCase() === oldLower) { excArr[j] = newTag; updateExc = true; }
        }
      }
      if (updateInc) segmentsSheet.getRange(i + 1, 3).setValue(incArr.join(', '));
      if (updateExc) segmentsSheet.getRange(i + 1, 4).setValue(excArr.join(', '));
    }
  }

  // 3. TagRules
  var rulesSheet = ss.getSheetByName('TagRules');
  if (rulesSheet) {
    var rulesData = rulesSheet.getDataRange().getValues();
    for (var i = 1; i < rulesData.length; i++) {
      var trigger = rulesData[i][1].toString();
      var toAddStr = rulesData[i][2].toString();
      var toRemoveStr = rulesData[i][3].toString();
      
      var updateTrigger = (trigger.toLowerCase() === oldLower);
      var updateAdd = false;
      var updateRemove = false;

      if (toAddStr) {
        var addArr = toAddStr.split(',').map(function(t) { return t.trim(); });
        for (var j = 0; j < addArr.length; j++) {
          if (addArr[j].toLowerCase() === oldLower) { addArr[j] = newTag; updateAdd = true; }
        }
      }
      if (toRemoveStr) {
        var remArr = toRemoveStr.split(',').map(function(t) { return t.trim(); });
        for (var j = 0; j < remArr.length; j++) {
          if (remArr[j].toLowerCase() === oldLower) { remArr[j] = newTag; updateRemove = true; }
        }
      }

      if (updateTrigger) rulesSheet.getRange(i + 1, 2).setValue(newTag);
      if (updateAdd) rulesSheet.getRange(i + 1, 3).setValue(addArr.join(', '));
      if (updateRemove) rulesSheet.getRange(i + 1, 4).setValue(remArr.join(', '));
    }
  }
}

function deleteTag(id) {
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheetByName('Tags');
    var rowIndex = findRowIndexById(sheet, id);
    if (rowIndex === -1) throw new Error('Tag tidak ditemukan');
    
    var tagName = sheet.getRange(rowIndex, 2).getValue().toString().trim();
    sheet.deleteRow(rowIndex);
    
    propagateTagDeletion(tagName);
    return { ok: true };
  } catch (err) {
    return { ok: false, message: err.toString() };
  }
}

function propagateTagDeletion(tagName) {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var tagLower = tagName.toLowerCase();
  
  // 1. Leads
  var leadsSheet = ss.getSheetByName('Leads');
  if (leadsSheet) {
    var leadsData = leadsSheet.getDataRange().getValues();
    for (var i = 1; i < leadsData.length; i++) {
      var tagsStr = leadsData[i][3].toString();
      if (tagsStr) {
        var tagsArr = tagsStr.split(',').map(function(t) { return t.trim(); });
        var filtered = tagsArr.filter(function(t) { return t.toLowerCase() !== tagLower; });
        if (filtered.length !== tagsArr.length) {
          leadsSheet.getRange(i + 1, 4).setValue(filtered.join(', '));
        }
      }
    }
  }

  // 2. Segments
  var segmentsSheet = ss.getSheetByName('Segments');
  if (segmentsSheet) {
    var segData = segmentsSheet.getDataRange().getValues();
    for (var i = 1; i < segData.length; i++) {
      var incStr = segData[i][2].toString();
      var excStr = segData[i][3].toString();
      
      if (incStr) {
        var incArr = incStr.split(',').map(function(t) { return t.trim(); });
        var filteredInc = incArr.filter(function(t) { return t.toLowerCase() !== tagLower; });
        if (filteredInc.length !== incArr.length) segmentsSheet.getRange(i + 1, 3).setValue(filteredInc.join(', '));
      }
      if (excStr) {
        var excArr = excStr.split(',').map(function(t) { return t.trim(); });
        var filteredExc = excArr.filter(function(t) { return t.toLowerCase() !== tagLower; });
        if (filteredExc.length !== excArr.length) segmentsSheet.getRange(i + 1, 4).setValue(filteredExc.join(', '));
      }
    }
  }
}

// Sync single contact to Google Contacts (Direct Server-Side without OAuth Client ID)
function syncContactToGoogleDirect(contactName, contactPhone, labelName) {
  try {
    var formattedPhone = '+' + normalizePhoneNumber(contactPhone);
    if (!formattedPhone) throw new Error('Nomor telepon tidak valid');
    
    // Ambil atau buat group label secara otomatis
    var group = ContactsApp.getContactGroup(labelName);
    if (!group) {
      group = ContactsApp.createContactGroup(labelName);
    }
    
    // Cek duplikasi kontak berdasarkan nomor telepon
    var existingContacts = ContactsApp.getContactsByPhone(formattedPhone);
    var contact;
    if (existingContacts.length > 0) {
      contact = existingContacts[0];
      contact.setFullName(contactName);
    } else {
      contact = ContactsApp.createContact(contactName, '', '');
      contact.addPhone(ContactsApp.Field.MOBILE_PHONE, formattedPhone);
    }
    
    contact.addToGroup(group);
    return { ok: true };
  } catch (err) {
    return { ok: false, message: err.toString() };
  }
}

// Bulk sync contacts to Google Contacts (Direct Server-Side)
function bulkSyncContactsToGoogleDirect(contactsArray, labelName) {
  try {
    var successCount = 0;
    var errors = [];
    
    for (var i = 0; i < contactsArray.length; i++) {
      var c = contactsArray[i];
      var res = syncContactToGoogleDirect(c.name, c.whatsapp, labelName);
      if (res.ok) {
        successCount++;
      } else {
        errors.push(c.name + ': ' + res.message);
      }
    }
    
    return {
      ok: true,
      successCount: successCount,
      total: contactsArray.length,
      errors: errors
    };
  } catch (err) {
    return { ok: false, message: err.toString() };
  }
}

// Mengambil semua Group/Label Google Contacts milik user aktif
function getGoogleContactGroups() {
  try {
    var groups = ContactsApp.getContactGroups();
    var list = groups.map(function(g) {
      return { id: g.getId(), name: g.getName() };
    });
    return { ok: true, data: list };
  } catch (e) {
    return { ok: false, message: e.toString() };
  }
}

// Menarik semua kontak berdasarkan label terpilih dari Google Contacts user aktif
function getContactsFromGoogleGroup(groupId) {
  try {
    var contacts = [];
    if (!groupId || groupId === 'ALL_CONTACTS') {
      contacts = ContactsApp.getContacts();
    } else {
      var group = ContactsApp.getContactGroupById(groupId);
      if (group) {
        contacts = group.getContacts();
      } else {
        throw new Error('Grup kontak tidak ditemukan.');
      }
    }

    var result = [];
    for (var i = 0; i < contacts.length; i++) {
      var c = contacts[i];
      var phones = c.getPhones();
      var mainPhone = '';
      if (phones.length > 0) {
        mainPhone = phones[0].getPhoneNumber();
      }
      
      // Hanya masukkan kontak yang memiliki nomor telepon
      if (mainPhone) {
        result.push({
          name: c.getFullName() || 'No Name',
          whatsapp: mainPhone
        });
      }
    }
    return { ok: true, data: result };
  } catch (e) {
    return { ok: false, message: e.toString() };
  }
}

// Handler Pengiriman Webhook Server-Side untuk menghindari CORS Error
function sendToN8nWebhook(webhookUrl, payload) {
  try {
    var options = {
      method: 'post',
      contentType: 'application/json',
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    };
    var response = UrlFetchApp.fetch(webhookUrl, options);
    var code = response.getResponseCode();
    if (code >= 200 && code < 300) {
      return { ok: true, code: code, body: response.getContentText() };
    } else {
      return { ok: false, message: 'n8n merespon dengan status ' + code + ': ' + response.getContentText() };
    }
  } catch (err) {
    return { ok: false, message: err.toString() };
  }
}

// WAHA Session Controller & QR Code Fetcher
function getWahaSettings() {
  var props = PropertiesService.getScriptProperties();
  return {
    waha_url: props.getProperty('WAHA_URL') || '',
    waha_api_key: props.getProperty('WAHA_API_KEY') || ''
  };
}

function saveWahaSettings(url, apiKey) {
  try {
    var props = PropertiesService.getScriptProperties();
    var cleanUrl = url ? url.toString().trim() : '';
    // Hapus tanda garis miring di akhir URL jika ada
    if (cleanUrl.indexOf('/', cleanUrl.length - 1) !== -1) {
      cleanUrl = cleanUrl.substring(0, cleanUrl.length - 1);
    }
    var cleanApiKey = apiKey ? apiKey.toString().trim() : '';
    props.setProperty('WAHA_URL', cleanUrl);
    props.setProperty('WAHA_API_KEY', cleanApiKey);
    return { ok: true };
  } catch(e) {
    return { ok: false, message: e.toString() };
  }
}

function getWahaSessionsStatus() {
  try {
    var settings = getWahaSettings();
    if (!settings.waha_url) {
      return { ok: false, message: 'URL WAHA belum dikonfigurasi di Pengaturan.' };
    }
    var url = settings.waha_url + '/api/sessions?all=true';
    var headers = {};
    if (settings.waha_api_key) {
      headers['Authorization'] = 'Bearer ' + settings.waha_api_key;
      headers['X-Api-Key'] = settings.waha_api_key;
    }
    var response = UrlFetchApp.fetch(url, {
      method: 'get',
      headers: headers,
      muteHttpExceptions: true
    });
    var code = response.getResponseCode();
    if (code !== 200) {
      return { ok: false, message: 'WAHA merespon dengan status ' + code };
    }
    var sessions = JSON.parse(response.getContentText());
    return { ok: true, sessions: sessions };
  } catch(e) {
    return { ok: false, message: e.toString() };
  }
}

function controlWahaSession(whatsapp, action) {
  try {
    var settings = getWahaSettings();
    var sessionName = getSessionNameByPhone(whatsapp);
    if (!settings.waha_url) throw new Error('URL WAHA belum dikonfigurasi');
    
    // WAHA API: POST /api/sessions/start, stop, logout, restart dengan body { "name": "session_name" }
    var url = settings.waha_url + '/api/sessions/' + action;
    var headers = {
      'Content-Type': 'application/json'
    };
    if (settings.waha_api_key) {
      headers['Authorization'] = 'Bearer ' + settings.waha_api_key;
      headers['X-Api-Key'] = settings.waha_api_key;
    }
    
    var payload = {
      name: sessionName
    };
    
    var response = UrlFetchApp.fetch(url, {
      method: 'post',
      headers: headers,
      contentType: 'application/json',
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    });
    
    var responseCode = response.getResponseCode();
    var responseText = response.getContentText();
    
    if (responseCode === 200 || responseCode === 201) {
      return { ok: true, body: responseText };
    } else {
      var errorMsg = responseText;
      try {
        var json = JSON.parse(responseText);
        if (json.message) {
          errorMsg = json.message;
        }
      } catch (err) {
        // Bukan format JSON, gunakan text biasa
      }
      return { ok: false, message: 'Status ' + responseCode + ': ' + errorMsg };
    }
  } catch(e) {
    return { ok: false, message: e.toString() };
  }
}

function getWahaSessionQR(whatsapp) {
  try {
    var settings = getWahaSettings();
    var sessionName = getSessionNameByPhone(whatsapp);
    if (!settings.waha_url) throw new Error('URL WAHA belum dikonfigurasi');
    
    // Perbaikan Analisis Mendalam: Menggunakan rute standar WAHA untuk QR Code yaitu /api/qr?session={sessionName}
    var url = settings.waha_url + '/api/qr?session=' + sessionName;
    var headers = {};
    if (settings.waha_api_key) {
      headers['Authorization'] = 'Bearer ' + settings.waha_api_key;
      headers['X-Api-Key'] = settings.waha_api_key;
    }
    var response = UrlFetchApp.fetch(url, {
      method: 'get',
      headers: headers,
      muteHttpExceptions: true
    });
    
    var code = response.getResponseCode();
    if (code === 200) {
      var blob = response.getBlob();
      var base64 = Utilities.base64Encode(blob.getBytes());
      var contentType = blob.getContentType() || 'image/png';
      return { ok: true, qr: 'data:' + contentType + ';base64,' + base64 };
    } else {
      var text = response.getContentText();
      try {
        var json = JSON.parse(text);
        return { ok: false, message: json.message || 'Sesi tidak siap memberikan QR (mungkin sudah terhubung atau mati).' };
      } catch(err) {
        return { ok: false, message: 'Status ' + code + ': ' + text };
      }
    }
  } catch(e) {
    return { ok: false, message: e.toString() };
  }
}

// Helper to upload image directly to Google Drive, make it public, and return direct link
function uploadImageToDrive(base64Data, fileName, mimeType) {
  try {
    var splitData = base64Data.split(',');
    var actualData = splitData.length > 1 ? splitData[1] : splitData[0];
    var decodedBytes = Utilities.base64Decode(actualData);
    var blob = Utilities.newBlob(decodedBytes, mimeType, fileName);
    
    // Simpan file ke Google Drive root
    var file = DriveApp.createFile(blob);
    
    // Ubah izin akses menjadi publik (Anyone with link can view)
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    
    // Buat direct link yang kompatibel dengan WAHA & n8n
    var directLink = 'https://lh3.googleusercontent.com/d/' + file.getId();
    
    return { ok: true, url: directLink };
  } catch (err) {
    return { ok: false, message: err.toString() };
  }
}

// Fungsi pemicu otorisasi UrlFetchApp secara paksa di Google
function mintaIzinGoogle() {
  try {
    // Melakukan fetch buatan ke google untuk memicu dialog izin akses internet
    UrlFetchApp.fetch('https://www.google.com', { muteHttpExceptions: true });
    Logger.log('Izin internet berhasil diverifikasi dan aktif!');
    return true;
  } catch (e) {
    Logger.log('Error: ' + e.toString());
    return false;
  }
}

// Mengambil kontak dari Google Contacts dan menyimpannya ke Sheets secara aman (Anti-Duplikasi & JSON-Safe)
function importGoogleContacts() {
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheetByName('Leads');
    if (!sheet) {
      setupDatabase();
      sheet = ss.getSheetByName('Leads');
    }
    
    // Tarik kontak dari Google People API v1
    var peopleFeed;
    try {
      peopleFeed = People.People.Connections.list('people/me', {
        personFields: 'names,phoneNumbers,emailAddresses',
        pageSize: 1000
      });
    } catch (err) {
      throw new Error('Gagal mengakses Google Contacts. Pastikan Anda telah mengizinkan otorisasi kontak: ' + err.toString());
    }

    var connections = peopleFeed.connections || [];
    if (connections.length === 0) {
      return { ok: true, importedCount: 0, message: 'Tidak ada kontak yang ditemukan di Google Contacts Anda.' };
    }

    var data = sheet.getDataRange().getValues();
    var headers = data[0];
    var emailColIndex = headers.indexOf('Email');
    
    // Jika kolom Email belum ada karena database lama, buat kolom baru secara dinamis
    if (emailColIndex === -1) {
      sheet.getRange(1, headers.length + 1).setValue('Email');
      data = sheet.getDataRange().getValues();
      headers = data[0];
      emailColIndex = headers.indexOf('Email');
    }

    // Map email yang sudah terdaftar di Sheets untuk mencegah duplikasi
    var existingEmails = {};
    for (var i = 1; i < data.length; i++) {
      var emailVal = data[i][emailColIndex];
      if (emailVal) {
        existingEmails[emailVal.toString().toLowerCase().trim()] = true;
      }
    }

    var now = new Date().toISOString();
    var importedCount = 0;
    var rowsToAdd = [];

    for (var j = 0; j < connections.length; j++) {
      var person = connections[j];
      
      // Ambil Nama
      var name = 'No Name';
      if (person.names && person.names.length > 0) {
        name = person.names[0].displayName || 'No Name';
      }

      // Ambil Email utama
      var email = '';
      if (person.emailAddresses && person.emailAddresses.length > 0) {
        email = person.emailAddresses[0].value || '';
      }

      // Ambil Nomor Telepon utama
      var phone = '';
      if (person.phoneNumbers && person.phoneNumbers.length > 0) {
        phone = person.phoneNumbers[0].value || '';
      }

      var normalizedPhone = phone ? normalizePhoneNumber(phone) : '';
      var cleanEmail = email.toLowerCase().trim();

      // Lewati kontak jika tidak memiliki email (karena validasi duplikasi berbasis email)
      if (!cleanEmail) continue;

      // Validasi duplikasi email
      if (!existingEmails[cleanEmail]) {
        var id = 'L' + Utilities.getUuid().substring(0, 8).toUpperCase();
        
        // Buat baris baru sesuai urutan header Leads database
        var newRow = [
          id,
          name,
          normalizedPhone,
          'Google Contacts', // Tag Sumber
          '',                // Owner WhatsApp
          now,               // Created At
          '',                // Last Broadcast At
          'Unchecked',       // Number Status
          '',                // Last Number Check At
          '',                // Inactive Reason
          email              // Email
        ];
        
        rowsToAdd.push(newRow);
        existingEmails[cleanEmail] = true;
        importedCount++;
      }
    }

    // Lakukan bulk append ke Google Sheets jika ada data baru
    if (rowsToAdd.length > 0) {
      var startRow = sheet.getLastRow() + 1;
      sheet.getRange(startRow, 1, rowsToAdd.length, headers.length).setValues(rowsToAdd);
    }

    return { ok: true, importedCount: importedCount };
  } catch (err) {
    return { ok: false, message: err.toString() };
  }
}