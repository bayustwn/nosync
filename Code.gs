// === NOTION → GOOGLE SHEETS SYNC ===

var PROPS = PropertiesService.getScriptProperties();
var NOTION_VERSION = '2022-06-28';

// === MENU ===

function onOpen() {
  try {
    SpreadsheetApp.getUi()
      .createMenu('Notion Sync')
      .addItem('Setup', 'showSetupDialog')
      .addSeparator()
      .addItem('Sync All', 'syncAll')
      .addSeparator()
      .addItem('Reset Config', 'resetConfig')
      .addToUi();
  } catch (e) {}
}

// === DIALOG ===

function showSetupDialog() {
  var html = HtmlService.createHtmlOutputFromFile('SetupDialog')
    .setWidth(600).setHeight(520);
  SpreadsheetApp.getUi().showModalDialog(html, 'Notion Sync');
}

// === NOTION API ===

function notion(method, path, body) {
  var key = PROPS.getProperty('NOTION_API_KEY');
  if (!key) throw new Error('API key belum di-set.');

  var opts = {
    method: method,
    headers: {
      'Authorization': 'Bearer ' + key,
      'Notion-Version': NOTION_VERSION,
      'Content-Type': 'application/json',
    },
    muteHttpExceptions: true,
  };
  if (body) opts.payload = JSON.stringify(body);

  var resp = UrlFetchApp.fetch('https://api.notion.com/v1' + path, opts);
  var data = JSON.parse(resp.getContentText());
  if (data.object === 'error') throw new Error('[' + data.status + '] ' + data.message);
  return data;
}

// === SETUP — CONNECTION ===

function testConnection(apiKey) {
  try {
    PROPS.setProperty('NOTION_API_KEY', apiKey);
    var data = notion('get', '/users/me');
    return { ok: true, name: data.name };
  } catch (e) {
    PROPS.deleteProperty('NOTION_API_KEY');
    return { ok: false, error: e.message };
  }
}

// === FETCH ALL DATABASES ===

function fetchAllDatabases() {
  var all = [];
  var cursor = null;

  while (true) {
    var body = {
      filter: { value: 'database', property: 'object' },
      page_size: 100,
    };
    if (cursor) body.start_cursor = cursor;

    var data = notion('post', '/search', body);
    (data.results || []).forEach(function(db) {
      var title = '';
      if (db.title && db.title.length > 0) {
        title = db.title.map(function(t) { return t.plain_text; }).join('');
      }
      all.push({ id: db.id, title: title || 'Untitled' });
    });

    if (!data.has_more) break;
    cursor = data.next_cursor;
  }

  all.sort(function(a, b) { return a.title.localeCompare(b.title); });
  return all;
}

// === LOOKUP SINGLE DATABASE (manual add) ===

function lookupDatabase(idOrUrl) {
  var id = idOrUrl;
  var m = idOrUrl.match(/([a-f0-9]{32})(?:\?|$|#)/i);
  if (m) id = m[1];

  try {
    var data = notion('get', '/databases/' + id);
    var title = '';
    if (data.title && data.title.length > 0) {
      title = data.title.map(function(t) { return t.plain_text; }).join('');
    }
    return { id: data.id, title: title || 'Untitled' };
  } catch (e) {
    return null;
  }
}

// === SAVE SELECTION ===

function getSavedSelection() {
  var raw = PROPS.getProperty('DB_SELECTION');
  if (!raw) return [];
  return JSON.parse(raw);
}

function saveSelection(selected) {
  PROPS.setProperty('DB_SELECTION', JSON.stringify(selected));
  return { ok: true, count: selected.length };
}

// === SYNC ===

function syncAll() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var selected = getSavedSelection();

  if (selected.length === 0) {
    ss.toast('Belum ada database dipilih. Buka Notion Sync → Setup.', 'Notion Sync', 5);
    return;
  }

  var ok = 0, fail = 0;
  selected.forEach(function(item) {
    try {
      var schema = notion('get', '/databases/' + item.id);
      var data = notion('post', '/databases/' + item.id + '/query', { page_size: 100 });

      var types = {};
      Object.keys(schema.properties || {}).forEach(function(k) {
        types[k] = schema.properties[k].type;
      });

      var result = parseResults(data.results, types);
      writeSheet(item.sheet || item.title, result.rows, result.types);
      ok++;
    } catch (e) {
      fail++;
      Logger.log('Sync gagal: ' + item.title + ' — ' + e.message);
    }
  });

  ss.toast('Sync selesai: ' + ok + ' berhasil' + (fail ? ', ' + fail + ' gagal' : ''), 'Notion Sync', 5);
}

function writeSheet(sheetName, rows, colTypes) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(sheetName);
  if (!sheet) sheet = ss.insertSheet(sheetName);
  else sheet.clearContents();

  if (rows.length === 0) return;

  var numCols = rows[0].length;
  var numRows = rows.length;

  // Write data
  sheet.getRange(1, 1, numRows, numCols).setValues(rows);

  // Header
  var header = sheet.getRange(1, 1, 1, numCols);
  header
    .setFontWeight('bold')
    .setFontSize(11)
    .setBackground('#6366f1')
    .setFontColor('#fff')
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle');
  sheet.setFrozenRows(1);

  // Alternating row colors
  if (numRows > 1) {
    var body = sheet.getRange(2, 1, numRows - 1, numCols);
    body.setVerticalAlignment('middle');
    var altRange = sheet.getRange(2, 1, numRows - 1, numCols);
    var bg = altRange.getBackgrounds();
    for (var r = 0; r < numRows - 1; r++) {
      for (var c = 0; c < numCols; c++) {
        bg[r][c] = r % 2 === 0 ? '#ffffff' : '#f8f8ff';
      }
    }
    altRange.setBackgrounds(bg);
  }

  // Column formatting
  if (colTypes) {
    var colKeys = Object.keys(colTypes);
    colKeys.forEach(function(key, i) {
      var col = i + 1;
      var type = colTypes[key];

      // Number
      if (type === 'number') {
        if (numRows > 1) {
          var range = sheet.getRange(2, col, numRows - 1, 1);
          range.setNumberFormat('#,##0.##');
        }
      }

      // Currency
      if (type === 'number' && key.toLowerCase().match(/price|cost|harga|biaya|amount|total/)) {
        if (numRows > 1) {
          var range = sheet.getRange(2, col, numRows - 1, 1);
          range.setNumberFormat('$#,##0.00');
        }
      }

      // Date
      if (type === 'date' || type === 'created_time' || type === 'last_edited_time') {
        if (numRows > 1) {
          var range = sheet.getRange(2, col, numRows - 1, 1);
          range.setNumberFormat('yyyy-mm-dd');
        }
      }

      // Checkbox
      if (type === 'checkbox') {
        if (numRows > 1) {
          var range = sheet.getRange(2, col, numRows - 1, 1);
          range.setHorizontalAlignment('center');
        }
      }
    });
  }

  // Auto-fit columns
  for (var c = 1; c <= numCols; c++) {
    sheet.autoResizeColumn(c);
  }

  // Add filter
  if (numRows > 1) {
    sheet.getRange(1, 1, numRows, numCols).createFilter();
  }
}

function parseResults(results, types) {
  if (!results || results.length === 0) return { rows: [['No data']], types: {} };

  var headers = Object.keys(results[0].properties);
  var rows = [headers];
  var colTypes = types || {};

  results.forEach(function(page) {
    var row = [];
    headers.forEach(function(h) { row.push(extract(page.properties[h])); });
    rows.push(row);
  });

  return { rows: rows, types: colTypes };
}

function extract(prop) {
  if (!prop) return '';
  switch (prop.type) {
    case 'title':
    case 'rich_text':
      return (prop[prop.type] || []).map(function(t) { return t.plain_text; }).join('');
    case 'select':
    case 'status':
      return prop[prop.type] ? prop[prop.type].name : '';
    case 'multi_select':
      return (prop.multi_select || []).map(function(s) { return s.name; }).join(', ');
    case 'date':
      return prop.date ? prop.date.start + (prop.date.end ? ' → ' + prop.date.end : '') : '';
    case 'number':     return prop.number !== null ? prop.number : '';
    case 'checkbox':   return prop.checkbox ? 'Yes' : 'No';
    case 'url':        return prop.url || '';
    case 'email':      return prop.email || '';
    case 'phone_number': return prop.phone_number || '';
    case 'people':     return (prop.people || []).map(function(p) { return p.name || p.id; }).join(', ');
    case 'created_time':
    case 'last_edited_time':
      return prop[prop.type] || '';
    case 'formula':
      if (prop.formula.type === 'string') return prop.formula.string || '';
      if (prop.formula.type === 'number') return '' + prop.formula.number;
      return '';
    default: return JSON.stringify(prop);
  }
}

function resetConfig() {
  PROPS.deleteAllProperties();
  SpreadsheetApp.getActiveSpreadsheet().toast('Config direset.', 'Notion Sync', 3);
}
