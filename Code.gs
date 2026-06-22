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

// Ensures database structure exists and is healthy
function setupDatabase() {
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheets = {
      'Leads': ['ID', 'Name', 'WhatsApp', 'Tags', 'Owner WhatsApp', 'Created At', 'Last Broadcast At', 'Number Status', 'Last Number Check At', 'Inactive Reason'],
      'Segments': ['ID', 'Name', 'Include Tags', 'Exclude Tags', 'Created At'],
      'Templates': ['ID', 'Name', 'Content', 'Product ID', 'Image URL', 'Created At'],
      'Broadcasts': ['ID', 'Name', 'Segment ID', 'Template ID', 'Status', 'Sent Count', 'Total Count', 'Created At', 'Scheduled At', 'Delay Contact Min', 'Delay Contact Max', 'Delay Bubble Min', 'Delay Bubble Max', 'Throttle Count', 'Throttle Minutes', 'Add Tags', 'Remove Tags', 'Sender CS', 'Target Lead IDs'],
      'BroadcastLogs': ['ID', 'Broadcast ID', 'Lead ID', 'Lead Name', 'WhatsApp', 'Status', 'Message ID', 'Error Message', 'Sent At', 'Delivered At', 'Read At', 'Scheduled At', 'Message Content', 'Sender CS', 'Retry Count', 'Max Retry', 'Processing Started At', 'Updated At', 'Image URL'],
      'TagRules': ['ID', 'Trigger Tag', 'Tags to Add', 'Tags to Remove', 'Assign CS', 'Created At'],
      'CSNumbers': ['ID', 'CS Name', 'WhatsApp Number', 'Created At'],
      'Products': ['ID', 'Product Name', 'Price', 'Description', 'Created At'],
      'Tags': ['ID', 'Tag Name', 'Created At']
    };
    
    // 1. Buat sheet jika belum ada, dan paksa update Header Row (Baris 1) agar selalu sinkron
    for (var sheetName in sheets) {
      var sheet = ss.getSheetByName(sheetName);
      if (!sheet) {
        sheet = ss.insertSheet(sheetName);
      }
      // Selalu perbarui header row ke versi standar terbaru
      sheet.getRange(1, 1, 1, sheets[sheetName].length).setValues([sheets[sheetName]]);
    }

    // 2. Pembersihan & Sanitasi Data Rusak di tabel Broadcasts (Misal: "1-1" di kolom Jeda)
    var broadcastsSheet = ss.getSheetByName('Broadcasts');
    if (broadcastsSheet) {
      var lastRow = broadcastsSheet.getLastRow();
      if (lastRow > 1) {
        var range = broadcastsSheet.getRange(2, 10, lastRow - 1, 6); // Ambil kolom J (10) sampai O (15)
        var values = range.getValues();
        var changed = false;
        
        for (var r = 0; r < values.length; r++) {
          // Kolom J (Delay Contact Min)
          var dcMin = values[r][0].toString();
          if (dcMin.indexOf('-') !== -1 || isNaN(parseInt(dcMin))) {
            values[r][0] = parseInt(dcMin.split('-')[0]) || 15;
            changed = true;
          } else {
            values[r][0] = parseInt(dcMin);
          }
          
          // Kolom K (Delay Contact Max)
          var dcMax = values[r][1].toString();
          if (dcMax.indexOf('-') !== -1 || isNaN(parseInt(dcMax))) {
            values[r][1] = parseInt(dcMax.split('-')[1]) || 30;
            changed = true;
          } else {
            values[r][1] = parseInt(dcMax);
          }

          // Kolom L (Delay Bubble Min)
          var dbMin = values[r][2].toString();
          if (dbMin.indexOf('-') !== -1 || isNaN(parseInt(dbMin))) {
            values[r][2] = parseInt(dbMin.split('-')[0]) || 2;
            changed = true;
          } else {
            values[r][2] = parseInt(dbMin);
          }

          // Kolom M (Delay Bubble Max)
          var dbMax = values[r][3].toString();
          if (dbMax.indexOf('-') !== -1 || isNaN(parseInt(dbMax))) {
            values[r][3] = parseInt(dbMax.split('-')[1]) || 5;
            changed = true;
          } else {
            values[r][3] = parseInt(dbMax);
          }
        }
        
        if (changed) {
          range.setValues(values);
        }
      }
    }
    
    // Paksa kolom WhatsApp menjadi format Plain Text (@) agar tidak berubah jadi format ilmiah (6.28E+11)
    var leadsSheet = ss.getSheetByName('Leads');
    if (leadsSheet) {
      leadsSheet.getRange('C2:C').setNumberFormat('@'); // Kolom WhatsApp
      leadsSheet.getRange('E2:E').setNumberFormat('@'); // Kolom Owner WhatsApp
    }
    var csSheet = ss.getSheetByName('CSNumbers');
    if (csSheet) {
      csSheet.getRange('C2:C').setNumberFormat('@'); // Kolom WhatsApp CS
    }
    var logsSheet = ss.getSheetByName('BroadcastLogs');
    if (logsSheet) {
      logsSheet.getRange('E2:E').setNumberFormat('@'); // Kolom WhatsApp Penerima
      logsSheet.getRange('N2:N').setNumberFormat('@'); // Kolom Sender CS
    }
    
    // Jalankan pembersihan otomatis untuk data lama yang masih mengandung tanda petik
    cleanExistingQuotes(ss);

    return { ok: true, message: 'Database schema successfully generated/verified.' };
  } catch (err) {
    return { ok: false, message: err.toString() };
  }
}

// Fungsi pembantu untuk membersihkan tanda petik satu (') dari database lama Anda
function cleanExistingQuotes(ss) {
  var sheetsToClean = ['Leads', 'CSNumbers', 'BroadcastLogs'];
  sheetsToClean.forEach(function(sheetName) {
    var sheet = ss.getSheetByName(sheetName);
    if (sheet) {
      var lastRow = sheet.getLastRow();
      var lastCol = sheet.getLastColumn();
      if (lastRow > 1) {
        var range = sheet.getRange(2, 1, lastRow - 1, lastCol);
        var values = range.getValues();
        var changed = false;
        
        for (var r = 0; r < values.length; r++) {
          for (var c = 0; c < values[r].length; c++) {
            var val = values[r][c];
            if (val && typeof val === 'string' && val.indexOf("'") === 0) {
              values[r][c] = val.replace(/'/g, '');
              changed = true;
            }
          }
        }
        if (changed) {
          range.setValues(values);
        }
      }
    }
  });
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

// Retrieves complete CRM context in one single safe database query
function getDashboardData() {
  try {
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
    if (!tagRulesSheet) {
      setupDatabase();
      tagRulesSheet = ss.getSheetByName('TagRules');
    }
    // Selalu jalankan setupDatabase di background untuk memastikan struktur tabel 100% sehat & ter-sanitasi
    setupDatabase();

    var leads = getSheetDataAsObjects(leadsSheet);
    var segments = getSheetDataAsObjects(segmentsSheet);
    var templates = getSheetDataAsObjects(templatesSheet);
    var broadcasts = getSheetDataAsObjects(broadcastsSheet);
    var tagRules = getSheetDataAsObjects(tagRulesSheet);
    var csNumbers = getSheetDataAsObjects(csNumbersSheet);
    var products = getSheetDataAsObjects(productsSheet);
    var tags = getSheetDataAsObjects(tagsSheet);
    var broadcastLogs = getSheetDataAsObjects(broadcastLogsSheet);
    var wahaSettings = getWahaSettings();
    
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
        waha_settings: wahaSettings
      })
    };
  } catch (err) {
    return { ok: false, message: err.toString() };
  }
}

// Helper to automatically register new tags into Master Tags sheet if they don't exist
function ensureTagsExistInMaster(tagsStr) {
  if (!tagsStr) return;
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var tagsSheet = ss.getSheetByName('Tags');
    if (!tagsSheet) return;
    
    var existingTags = getSheetDataAsObjects(tagsSheet).map(function(t) {
      return t.tag_name.toLowerCase().trim();
    });
    
    var inputTags = tagsStr.split(',').map(function(t) { return t.trim(); }).filter(Boolean);
    var now = new Date().toISOString();
    
    inputTags.forEach(function(tag) {
      var tagLower = tag.toLowerCase();
      if (existingTags.indexOf(tagLower) === -1) {
        var id = 'TAG' + Utilities.getUuid().substring(0, 8).toUpperCase();
        tagsSheet.appendRow([id, tag, now]);
        existingTags.push(tagLower); // Tambahkan ke cache lokal agar tidak dobel dalam satu proses
      }
    });
  } catch (e) {
    console.error('Gagal mendaftarkan tag master otomatis: ' + e.toString());
  }
}

// Helper to process Tag Automation Rules
function processTagAutomations(tagsStr) {
  if (!tagsStr) return '';
  // Pastikan semua tag yang diproses terdaftar di master tags
  ensureTagsExistInMaster(tagsStr);
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
    
    for (var j = 0; j < leads.length; j++) {
      var lead = leads[j];
      var rawWA = lead.whatsapp ? lead.whatsapp.toString().trim() : '';
      if (!rawWA) continue;
      
      var wa = normalizePhoneNumber(rawWA);
      var processedTags = processTagAutomations(lead.tags || '');
      var finalOwner = processCSAutoAssign(processedTags, lead.owner_whatsapp || '');
      
      var formattedOwner = finalOwner ? normalizePhoneNumber(finalOwner) : '';
      
      var waForSheet = wa;
      var ownerForSheet = formattedOwner;
      
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
            // Proteksi: Lewati otomatis jika nomor berstatus Inactive (Mati)
            if (lead.number_status === 'Inactive') return false;
            
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
                parseInt(broadcast.delay_contact_min) || 15,  // Kolom J: Delay Contact Min (Pastikan Integer)
                parseInt(broadcast.delay_contact_max) || 30,  // Kolom K: Delay Contact Max (Pastikan Integer)
                parseInt(broadcast.delay_bubble_min) || 2,    // Kolom L: Delay Bubble Min (Pastikan Integer)
                parseInt(broadcast.delay_bubble_max) || 5,    // Kolom M: Delay Bubble Max (Pastikan Integer)
                parseInt(broadcast.throttle_count) || 0,       // Kolom N: Throttle Count (Pastikan Integer)
                parseInt(broadcast.throttle_minutes) || 0,     // Kolom O: Throttle Minutes (Pastikan Integer)
                broadcast.add_tags || '',                      // Kolom P: Add Tags
                broadcast.remove_tags || '',                   // Kolom Q: Remove Tags
                senderCS,                                      // Kolom R: Sender CS
                targetLeadIds                                  // Kolom S: Target Lead IDs
              ]);

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

function executePostBroadcastAction(leadId, addTagsStr, removeTagsStr) {
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheetByName('Leads');
    var rowIndex = findRowIndexById(sheet, leadId);
    if (rowIndex === -1) return { ok: false, message: 'Lead ID tidak ditemukan' };
    
    var tagsCol = 4; // Kolom D: Tags (1-based)
    var lastBcCol = 7; // Kolom G: Last Broadcast At (1-based)
    var now = new Date().toISOString();
    
    var currentTagsVal = sheet.getRange(rowIndex, tagsCol).getValue() ? sheet.getRange(rowIndex, tagsCol).getValue().toString() : '';
    var currentTags = currentTagsVal.split(',').map(function(t) { return t.trim(); }).filter(Boolean);
    
    // Proses hapus tag pasca broadcast
    if (removeTagsStr) {
      var toRemove = removeTagsStr.split(',').map(function(t) { return t.trim().toLowerCase(); }).filter(Boolean);
      currentTags = currentTags.filter(function(t) {
        return toRemove.indexOf(t.toLowerCase()) === -1;
      });
    }
    
    // Proses tambah tag pasca broadcast
    if (addTagsStr) {
      // Daftarkan otomatis ke Master Tags jika belum ada
      ensureTagsExistInMaster(addTagsStr);
      
      var toAdd = addTagsStr.split(',').map(function(t) { return t.trim(); }).filter(Boolean);
      toAdd.forEach(function(t) {
        var lowerTags = currentTags.map(function(ct) { return ct.toLowerCase(); });
        if (lowerTags.indexOf(t.toLowerCase()) === -1) {
          currentTags.push(t);
        }
      });
    }
    
    var finalizedTags = processTagAutomations(currentTags.join(', '));
    var currentOwner = sheet.getRange(rowIndex, tagsCol + 1).getValue() ? sheet.getRange(rowIndex, tagsCol + 1).getValue().toString() : '';
    var finalizedOwner = processCSAutoAssign(finalizedTags, currentOwner);
    
    // Update Tags, CS Owner, dan Last Broadcast At secara akurat
    sheet.getRange(rowIndex, tagsCol, 1, 2).setValues([[finalizedTags, finalizedOwner]]);
    sheet.getRange(rowIndex, lastBcCol).setValue(now);
    
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
// Helper untuk mendeteksi apakah status pengiriman bernilai SUKSES (Mendukung format teks biasa & JSON dari n8n/WAHA)
function isStatusSuccess(statusValue) {
  if (!statusValue) return false;
  var str = statusValue.toString().trim().toUpperCase();
  
  // Jika status berupa format JSON, cari kata kunci sukses di dalamnya
  if (str.indexOf('{') === 0) {
    try {
      var parsed = JSON.parse(statusValue);
      if (parsed) {
        var val = '';
        if (parsed.result) val = parsed.result.toString().toUpperCase();
        else if (parsed.status) val = parsed.status.toString().toUpperCase();
        else if (parsed.success !== undefined) val = parsed.success.toString().toUpperCase();
        
        if (val === 'OK' || val === 'TRUE' || val === 'SENT' || val === 'DELIVERED' || val === 'READ' || val === 'SUCCESS') {
          return true;
        }
      }
    } catch (e) {
      // Fallback ke pencarian teks manual jika gagal parse JSON
    }
  }
  
  // Pencarian kata kunci sukses pada string biasa
  var successKeywords = ['SENT', 'DELIVERED', 'READ', 'SUCCESS', 'OK', 'TRUE'];
  for (var i = 0; i < successKeywords.length; i++) {
    if (str.indexOf(successKeywords[i]) !== -1) {
      return true;
    }
  }
  return false;
}

function updateBroadcastLogStatus(logId, status, messageId, errorMessage) {
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheetByName('BroadcastLogs');
    var rowIndex = findRowIndexById(sheet, logId);
    if (rowIndex === -1) throw new Error('Log row tidak ditemukan');
    
    var now = new Date().toISOString();
    var oldStatus = sheet.getRange(rowIndex, 6).getValue() ? sheet.getRange(rowIndex, 6).getValue().toString().trim() : '';
    var newStatusUpper = status ? status.toString().trim().toUpperCase() : '';
    
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

    var isSuccess = isStatusSuccess(status);
    
    if (isSuccess) {
      // Set Sent At (Kolom I) jika belum terisi
      var currentSentAt = sheet.getRange(rowIndex, 9).getValue();
      if (!currentSentAt) {
        sheet.getRange(rowIndex, 9).setValue(now);
      }
      sheet.getRange(rowIndex, 17).setValue(''); // Reset Processing Started At karena sudah sukses
      
      if (newStatusUpper.indexOf('DELIVERED') !== -1) {
        sheet.getRange(rowIndex, 10).setValue(now); // Kolom J: Delivered At
      } else if (newStatusUpper.indexOf('READ') !== -1) {
        sheet.getRange(rowIndex, 11).setValue(now); // Kolom K: Read At
      }
      
      // Jalankan aksi update tag
      try {
        var broadcastId = sheet.getRange(rowIndex, 2).getValue() ? sheet.getRange(rowIndex, 2).getValue().toString() : '';
        
        if (leadRowIndex !== -1) {
          leadsSheet.getRange(leadRowIndex, 7).setValue(now); // Kolom G: Last Broadcast At
          leadsSheet.getRange(leadRowIndex, 8).setValue('Active'); // Kolom H: Number Status
          leadsSheet.getRange(leadRowIndex, 9).setValue(now); // Kolom I: Last Number Check At
          leadsSheet.getRange(leadRowIndex, 10).setValue(''); // Kolom J: Inactive Reason (Reset karena aktif)
        }

        // Ambil aturan Tambah/Hapus tag secara presisi dari tabel Broadcasts induk
        var broadcastsSheet = ss.getSheetByName('Broadcasts');
        var bcRowIndex = findRowIndexById(broadcastsSheet, broadcastId);
        if (bcRowIndex !== -1) {
          // Membaca Kolom 16 (P: Add Tags) dan Kolom 17 (Q: Remove Tags) secara aman
          var addTagsStr = broadcastsSheet.getRange(bcRowIndex, 16).getValue() ? broadcastsSheet.getRange(bcRowIndex, 16).getValue().toString() : ''; 
          var removeTagsStr = broadcastsSheet.getRange(bcRowIndex, 17).getValue() ? broadcastsSheet.getRange(bcRowIndex, 17).getValue().toString() : ''; 
          
          // Jalankan penambahan/penghapusan tag & alokasi CS otomatis menggunakan Lead ID yang 100% unik
          if (addTagsStr || removeTagsStr) {
            executePostBroadcastAction(leadId, addTagsStr, removeTagsStr);
          }
        }
      } catch (e) {
        console.error("Gagal memproses update tag pasca broadcast: " + e.toString());
      }
    } else if (newStatusUpper.indexOf('INVALID') !== -1) {
      sheet.getRange(rowIndex, 17).setValue(''); // Reset Processing lock
      
      // Otomatis tandai nomor mati permanen di tabel Leads
      if (leadRowIndex !== -1) {
        leadsSheet.getRange(leadRowIndex, 8).setValue('Inactive'); // Kolom H: Number Status
        leadsSheet.getRange(leadRowIndex, 9).setValue(now); // Kolom I: Last Number Check At
        leadsSheet.getRange(leadRowIndex, 10).setValue(errorMessage || 'Nomor tidak terdaftar di WhatsApp'); // Kolom J: Inactive Reason
      }
    } else if (newStatusUpper === 'FAILED') {
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
      var rowMessageId = data[i][6].toString();
      
      var match = false;
      if (rowMessageId === messageId.toString()) {
        match = true;
      } else if (rowMessageId.indexOf('{') === 0) {
        try {
          var parsed = JSON.parse(rowMessageId);
          if (parsed.id === messageId.toString() || parsed._serialized === messageId.toString()) {
            match = true;
          }
        } catch(e) {}
      }
      
      if (match) {
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
      sheet.appendRow([id, cs.cs_name, formattedWA, now]);
    } else {
      var rowIndex = findRowIndexById(sheet, cs.id);
      if (rowIndex === -1) throw new Error('CS Number tidak ditemukan');
      sheet.getRange(rowIndex, 2, 1, 2).setValues([[cs.cs_name, formattedWA]]);
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
}``