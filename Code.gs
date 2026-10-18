const SPREADSHEET_ID = '1yjfGwJ0UKC1E2953P3n-sVkPxyKsEixlmoolE2OzDfk';
const CACHE_KEY_STATE = 'OKEGAS_APP_INITIAL_STATE';
const CACHE_TTL = 300; // 5 minutes cache

/**
 * Serves the Web App Frontend & Extension REST Endpoint
 */
function doGet(e) {
  // REST Endpoint for OKEGAS Chrome Extension
  if (e && e.parameter && e.parameter.format === 'json') {
    try {
      var state = getInitialAppState({});
      return ContentService.createTextOutput(JSON.stringify({ ok: true, data: state }))
        .setMimeType(ContentService.MimeType.JSON);
    } catch (err) {
      return ContentService.createTextOutput(JSON.stringify({ ok: false, message: err.message }))
        .setMimeType(ContentService.MimeType.JSON);
    }
  }

  // Frontend SPA Web App
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('OKEGAS Central Template Hub & CS Broadcast Engine')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, maximum-scale=1, user-scalable=no')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/**
 * REST API Dispatcher via doPost
 */
function doPost(e) {
  try {
    if (!e || !e.postData || !e.postData.contents) {
      throw new Error('Payload kosong atau request tidak valid.');
    }
    var body = JSON.parse(e.postData.contents);
    var action = body.action;
    var payload = body.payload || {};
    var result;

    if (action === 'getInitialAppState') {
      result = getInitialAppState(payload);
    } else if (action === 'saveTemplate') {
      result = saveTemplate(payload);
    } else if (action === 'deleteTemplate') {
      result = deleteTemplate(payload);
    } else if (action === 'manageProduct') {
      result = manageProduct(payload);
    } else if (action === 'manageCategory') {
      result = manageCategory(payload);
    } else if (action === 'setupDatabase') {
      result = setupDatabase();
    } else {
      throw new Error('Aksi tidak dikenal: ' + action);
    }

    return ContentService.createTextOutput(JSON.stringify({ ok: true, data: serializeForClient(result) }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ ok: false, message: err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

/**
 * Shared serializer ensuring JSON-safe payloads
 */
function serializeForClient(val) {
  if (val === null || val === undefined) return null;
  if (val instanceof Date) return val.toISOString();
  if (Array.isArray(val)) {
    return val.map(function(item) { return serializeForClient(item); });
  }
  if (typeof val === 'object') {
    var copy = {};
    for (var k in val) {
      if (Object.prototype.hasOwnProperty.call(val, k)) {
        copy[k] = serializeForClient(val[k]);
      }
    }
    return copy;
  }
  return val;
}

/**
 * Helper to get or create sheet tab and sync header row safely
 */
function getOrCreateSheetWithHeaders(ss, sheetName, headers) {
  var sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
  }
  // Always synchronize headers on row 1 to guarantee schema consistency
  var currentHeaders = sheet.getLastColumn() > 0 ? sheet.getRange(1, 1, 1, headers.length).getValues()[0] : [];
  var mismatch = false;
  if (currentHeaders.length < headers.length) {
    mismatch = true;
  } else {
    for (var i = 0; i < headers.length; i++) {
      if (currentHeaders[i] !== headers[i]) {
        mismatch = true;
        break;
      }
    }
  }
  if (mismatch || sheet.getLastRow() === 0) {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  }
  return sheet;
}

/**
 * Initializes Google Sheets database tables and seed records. Safe to rerun.
 */
function setupDatabase() {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);

  var templateHeaders = ['id', 'name', 'productId', 'categoryId', 'engine', 'bubblesJson', 'createdAt', 'updatedAt'];
  var productHeaders = ['id', 'name', 'status'];
  var categoryHeaders = ['id', 'name', 'status'];

  var tplSheet = getOrCreateSheetWithHeaders(ss, 'Templates', templateHeaders);
  var prodSheet = getOrCreateSheetWithHeaders(ss, 'Products', productHeaders);
  var catSheet = getOrCreateSheetWithHeaders(ss, 'Categories', categoryHeaders);

  // Seed default Products if empty
  if (prodSheet.getLastRow() <= 1) {
    var defaultProducts = [
      ['prod_1', 'DVN', 'active'],
      ['prod_2', 'Stiju', 'active']
    ];
    prodSheet.getRange(2, 1, defaultProducts.length, 3).setValues(defaultProducts);
  }

  // Seed default Categories if empty
  if (catSheet.getLastRow() <= 1) {
    var defaultCategories = [
      ['cat_1', 'Cold Lead', 'active'],
      ['cat_2', 'Warm Lead', 'active'],
      ['cat_3', 'Repeat Order (RO)', 'active'],
      ['cat_4', 'Handling Inquiry', 'active'],
      ['cat_5', 'Payment Reminder', 'active']
    ];
    catSheet.getRange(2, 1, defaultCategories.length, 3).setValues(defaultCategories);
  }

  // Seed default Templates if empty
  if (tplSheet.getLastRow() <= 1) {
    var now = new Date().toISOString();
    var sampleBubbles1 = [
      {
        id: 'b1_1',
        text: 'Halo Kak *{name}*! 👋 Terima kasih sudah mampir di official store {product}.\n\nAda yang bisa kami bantu terkait konsultasi pemakaian hari ini? 😊',
        images: ['https://lh3.googleusercontent.com/d/1sampleImageIdDVNPromo']
      },
      {
        id: 'b1_2',
        text: 'Khusus pemesanan hari ini, dapatkan promo bundling *{var1}* + _Free Ongkir_ ke seluruh Indonesia! 🚀\n\nMau kami bantu amankan slot promonya sekarang Kak?',
        images: []
      }
    ];

    var sampleBubbles2 = [
      {
        id: 'b2_1',
        text: 'Hai Kak {name}, paket {product} kemarin sudah sampai ya? 🙏 Semoga cocok dan memberikan hasil maksimal!',
        images: []
      },
      {
        id: 'b2_2',
        text: 'Bagi customer setia seperti Kakak, kami sediakan voucher Repeat Order 25% dengan kode *ROSETIA25* 🎁.\n\nStok terbatas, mau dibantu proses kirim hari ini Kak?',
        images: ['https://lh3.googleusercontent.com/d/1sampleImageIdVoucher']
      }
    ];

    var defaultTemplates = [
      ['tpl_1', 'DVN - Followup Cold Lead Promo', 'prod_1', 'cat_1', 'manual', JSON.stringify(sampleBubbles1), now, now],
      ['tpl_2', 'Stiju - Special Repeat Order VIP', 'prod_2', 'cat_3', 'woowa', JSON.stringify(sampleBubbles2), now, now]
    ];
    tplSheet.getRange(2, 1, defaultTemplates.length, 8).setValues(defaultTemplates);
  }

  clearScriptCache();
  return { status: 'Database setup completed successfully' };
}

/**
 * Cache management helpers
 */
function clearScriptCache() {
  try {
    var cache = CacheService.getScriptCache();
    cache.remove(CACHE_KEY_STATE);
  } catch (e) {
    // Graceful fallback
  }
}

/**
 * Fetches all products, categories, and templates in a single batch read.
 */
function getInitialAppState(payload) {
  var cache = CacheService.getScriptCache();
  var cached = cache.get(CACHE_KEY_STATE);
  if (cached && (!payload || !payload.forceRefresh)) {
    try {
      return JSON.parse(cached);
    } catch (e) {
      // Continue to fresh read
    }
  }

  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var prodSheet = ss.getSheetByName('Products');
  var catSheet = ss.getSheetByName('Categories');
  var tplSheet = ss.getSheetByName('Templates');

  if (!prodSheet || !catSheet || !tplSheet) {
    setupDatabase();
    prodSheet = ss.getSheetByName('Products');
    catSheet = ss.getSheetByName('Categories');
    tplSheet = ss.getSheetByName('Templates');
  }

  // Batch read Products
  var products = [];
  if (prodSheet.getLastRow() > 1) {
    var pRows = prodSheet.getRange(2, 1, prodSheet.getLastRow() - 1, 3).getValues();
    for (var i = 0; i < pRows.length; i++) {
      if (pRows[i][0]) {
        products.push({ id: String(pRows[i][0]), name: String(pRows[i][1]), status: String(pRows[i][2]) });
      }
    }
  }

  // Batch read Categories
  var categories = [];
  if (catSheet.getLastRow() > 1) {
    var cRows = catSheet.getRange(2, 1, catSheet.getLastRow() - 1, 3).getValues();
    for (var j = 0; j < cRows.length; j++) {
      if (cRows[j][0]) {
        categories.push({ id: String(cRows[j][0]), name: String(cRows[j][1]), status: String(cRows[j][2]) });
      }
    }
  }

  // Batch read Templates
  var templates = [];
  if (tplSheet.getLastRow() > 1) {
    var tRows = tplSheet.getRange(2, 1, tplSheet.getLastRow() - 1, 8).getValues();
    for (var k = 0; k < tRows.length; k++) {
      if (tRows[k][0]) {
        templates.push({
          id: String(tRows[k][0]),
          name: String(tRows[k][1]),
          productId: String(tRows[k][2]),
          categoryId: String(tRows[k][3]),
          engine: String(tRows[k][4] || 'manual'),
          bubblesJson: String(tRows[k][5]),
          createdAt: tRows[k][6] instanceof Date ? tRows[k][6].toISOString() : String(tRows[k][6]),
          updatedAt: tRows[k][7] instanceof Date ? tRows[k][7].toISOString() : String(tRows[k][7])
        });
      }
    }
  }

  var responseData = {
    products: products,
    categories: categories,
    templates: templates
  };

  try {
    cache.put(CACHE_KEY_STATE, JSON.stringify(responseData), CACHE_TTL);
  } catch (e) {
    // Cache payload size threshold safe bypass
  }

  return responseData;
}

/**
 * Inserts or updates a template row.
 */
function saveTemplate(payload) {
  if (!payload || !payload.name) {
    throw new Error('Nama template wajib diisi.');
  }

  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName('Templates');
  if (!sheet) {
    setupDatabase();
    sheet = ss.getSheetByName('Templates');
  }

  var now = new Date().toISOString();
  var id = payload.id;
  var isNew = !id;

  if (isNew) {
    id = 'tpl_' + Utilities.getUuid().substring(0, 8);
  }

  var record = [
    id,
    payload.name,
    payload.productId || '',
    payload.categoryId || '',
    payload.engine || 'manual',
    payload.bubblesJson || '[]',
    isNew ? now : (payload.createdAt || now),
    now
  ];

  var existingRowIndex = -1;
  if (!isNew && sheet.getLastRow() > 1) {
    var ids = sheet.getRange(2, 1, sheet.getLastRow() - 1, 1).getValues();
    for (var i = 0; i < ids.length; i++) {
      if (String(ids[i][0]) === String(id)) {
        existingRowIndex = i + 2;
        break;
      }
    }
  }

  if (existingRowIndex > 1) {
    sheet.getRange(existingRowIndex, 1, 1, 8).setValues([record]);
  } else {
    sheet.appendRow(record);
  }

  clearScriptCache();

  return {
    id: id,
    name: payload.name,
    productId: payload.productId || '',
    categoryId: payload.categoryId || '',
    engine: payload.engine || 'manual',
    bubblesJson: payload.bubblesJson || '[]',
    createdAt: isNew ? now : (payload.createdAt || now),
    updatedAt: now
  };
}

/**
 * Deletes a template by ID
 */
function deleteTemplate(payload) {
  if (!payload || !payload.id) {
    throw new Error('ID Template tidak valid');
  }

  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName('Templates');
  if (!sheet || sheet.getLastRow() <= 1) return { deleted: false };

  var ids = sheet.getRange(2, 1, sheet.getLastRow() - 1, 1).getValues();
  for (var i = 0; i < ids.length; i++) {
    if (String(ids[i][0]) === String(payload.id)) {
      sheet.deleteRow(i + 2);
      clearScriptCache();
      return { deleted: true, id: payload.id };
    }
  }

  return { deleted: false };
}

/**
 * Product Master Management (Create / Delete)
 */
function manageProduct(payload) {
  var action = payload.action;
  var data = payload.data || {};
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName('Products');
  if (!sheet) {
    setupDatabase();
    sheet = ss.getSheetByName('Products');
  }

  if (action === 'create') {
    if (!data.name) throw new Error('Nama produk wajib diisi.');
    var id = 'prod_' + Utilities.getUuid().substring(0, 6);
    sheet.appendRow([id, data.name, 'active']);
    clearScriptCache();
    return { id: id, name: data.name, status: 'active' };
  } else if (action === 'delete') {
    if (!data.id) throw new Error('ID Produk wajib dicantumkan.');
    if (sheet.getLastRow() > 1) {
      var ids = sheet.getRange(2, 1, sheet.getLastRow() - 1, 1).getValues();
      for (var i = 0; i < ids.length; i++) {
        if (String(ids[i][0]) === String(data.id)) {
          sheet.deleteRow(i + 2);
          clearScriptCache();
          return { deleted: true };
        }
      }
    }
  }
  return { status: 'noop' };
}

/**
 * Category Master Management (Create / Delete)
 */
function manageCategory(payload) {
  var action = payload.action;
  var data = payload.data || {};
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName('Categories');
  if (!sheet) {
    setupDatabase();
    sheet = ss.getSheetByName('Categories');
  }

  if (action === 'create') {
    if (!data.name) throw new Error('Nama kategori wajib diisi.');
    var id = 'cat_' + Utilities.getUuid().substring(0, 6);
    sheet.appendRow([id, data.name, 'active']);
    clearScriptCache();
    return { id: id, name: data.name, status: 'active' };
  } else if (action === 'delete') {
    if (!data.id) throw new Error('ID Kategori wajib dicantumkan.');
    if (sheet.getLastRow() > 1) {
      var ids = sheet.getRange(2, 1, sheet.getLastRow() - 1, 1).getValues();
      for (var i = 0; i < ids.length; i++) {
        if (String(ids[i][0]) === String(data.id)) {
          sheet.deleteRow(i + 2);
          clearScriptCache();
          return { deleted: true };
        }
      }
    }
  }
  return { status: 'noop' };
}
