const SHEETS = {
  CURRENT_STOCK: 'Current Stock',
  TRANSACTIONS: 'Transactions',
  USERS: 'Users'
};

const STOCK_HEADERS = [
  'Stock ID',
  'Godown',
  'Item',
  'Quantity Cases',
  'Last Updated',
  'Last Updated By'
];

const TRANSACTION_HEADERS = [
  'Timestamp',
  'Action',
  'Godown',
  'Item',
  'Quantity Change',
  'Quantity Before',
  'Quantity After',
  'Note',
  'Telegram User ID',
  'Telegram Name',
  'Platform',
  'Request ID'
];

const USER_HEADERS = [
  'Telegram User ID',
  'Name',
  'Role',
  'Active'
];

function doGet(e) {
  const p = (e && e.parameter) || {};
  const action = p.action;
  const cb = p.callback; // JSONP support

  function respond(result) {
    const json = JSON.stringify(result);
    if (cb) {
      return ContentService.createTextOutput(cb + '(' + json + ')')
        .setMimeType(ContentService.MimeType.JAVASCRIPT);
    }
    return buildJsonResponse_(result);
  }

  if (action === 'load') {
    try {
      return respond(getInitialData((p.initData || '').trim()));
    } catch (err) {
      return respond({ok: false, error: err.message});
    }
  }

  if (action === 'batchSubmit') {
    try {
      const data = JSON.parse(p.data || '{}');
      return respond(batchSubmit_(
        (p.initData || '').trim(),
        data.movements || [],
        data.newItems  || [],
        data.note      || ''
      ));
    } catch (err) {
      return respond({ok: false, error: err.message});
    }
  }

  return HtmlService
    .createHtmlOutputFromFile('Index')
    .setTitle('NK Inventory')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, viewport-fit=cover')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function doPost(e) {
  try {
    const body = JSON.parse(e.postData.contents);
    const initData = (body.initData || '').trim();
    if (body.action === 'submit') {
      return buildJsonResponse_(submitMovement(initData, body.payload));
    }
    if (body.action === 'batchSubmit') {
      return buildJsonResponse_(batchSubmit_(initData, body.movements || [], body.newItems || [], body.note || ''));
    }
    throw new Error('Unknown action');
  } catch (err) {
    return buildJsonResponse_({ok: false, error: err.message});
  }
}

function batchSubmit_(initData, movements, newItems, note) {
  const userContext = requireAllowedTelegramUser_(initData);
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);

  try {
    const ss = getSpreadsheet_();
    const stockSheet = requireSheet_(ss, SHEETS.CURRENT_STOCK);
    const transactionSheet = requireSheet_(ss, SHEETS.TRANSACTIONS);
    const notifications = [];
    const now = new Date();

    // 1. Add new items
    if (newItems.length) {
      const existing = stockSheet.getDataRange().getValues();
      for (let i = 0; i < newItems.length; i++) {
        const ni = newItems[i];
        const godown = String(ni.godown || '').trim();
        const item = String(ni.item || '').trim();
        const qty = Math.max(0, Number(ni.quantity) || 0);
        if (!godown || !item) continue;

        const stockId = makeStockId_(godown, item);
        for (let r = 1; r < existing.length; r++) {
          if (String(existing[r][0]) === stockId) {
            throw new Error('"' + item + '" already exists in ' + godown + '.');
          }
        }

        stockSheet.appendRow([stockId, godown, item, qty, now, userContext.displayName]);
        transactionSheet.appendRow([
          now, 'IN', godown, item, qty, 0, qty,
          note || 'New item added',
          String(userContext.user.id), userContext.displayName, 'miniapp', Utilities.getUuid()
        ]);

        notifications.push(
          '<b>New Item</b>\n<b>Item:</b> ' + escapeHtml_(item) +
          '\n<b>Godown:</b> ' + escapeHtml_(godown) +
          '\n<b>Qty:</b> ' + qty + ' cases' +
          '\n<b>By:</b> ' + escapeHtml_(userContext.displayName)
        );
      }
    }

    // 2. Apply movements
    for (let i = 0; i < movements.length; i++) {
      const m = movements[i];
      const delta = Number(m.delta) || 0;
      if (delta === 0) continue;

      const godown = String(m.godown || '').trim();
      const item = String(m.item || '').trim();
      if (!godown || !item) continue;

      const payload = {
        action: delta > 0 ? 'IN' : 'OUT',
        godown: godown,
        item: item,
        quantity: Math.abs(delta),
        note: note || ''
      };
      const result = applyMovement_(payload, userContext);
      notifications.push(result.notification);
    }

    if (notifications.length > 0) {
      sendTelegramNotification_(notifications.join('\n\n─────────\n\n'));
    }

    return buildClientState_(userContext);
  } finally {
    lock.releaseLock();
  }
}

function buildJsonResponse_(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

function setupInventoryWorkbook() {
  const ss = getSpreadsheet_();
  const stockSheet = ensureSheet_(ss, SHEETS.CURRENT_STOCK, STOCK_HEADERS);
  const transactionSheet = ensureSheet_(ss, SHEETS.TRANSACTIONS, TRANSACTION_HEADERS);
  const usersSheet = ensureSheet_(ss, SHEETS.USERS, USER_HEADERS);

  if (usersSheet.getLastRow() === 1) {
    usersSheet.getRange(2, 1, 1, USER_HEADERS.length).setValues([
      ['PASTE_YOUR_TELEGRAM_USER_ID', 'Rahul', 'admin', true]
    ]);
  }

  formatSheet_(stockSheet, STOCK_HEADERS.length);
  formatSheet_(transactionSheet, TRANSACTION_HEADERS.length);
  formatSheet_(usersSheet, USER_HEADERS.length);
  stockSheet.getRange('A:A').setNumberFormat('@');
  transactionSheet.getRange('I:I').setNumberFormat('@');
  usersSheet.getRange('A:A').setNumberFormat('@');
  stockSheet.autoResizeColumns(1, STOCK_HEADERS.length);
  transactionSheet.autoResizeColumns(1, TRANSACTION_HEADERS.length);
  usersSheet.autoResizeColumns(1, USER_HEADERS.length);

  return {
    message: 'Inventory sheets checked. Add stock rows to Current Stock and update Users with allowed Telegram user IDs.'
  };
}

function getInitialData(initData) {
  const userContext = requireAllowedTelegramUser_(initData);
  return buildClientState_(userContext);
}

function submitMovement(initData, payload) {
  const userContext = requireAllowedTelegramUser_(initData);
  const clean = validatePayload_(payload);
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);

  let result;
  try {
    result = applyMovement_(clean, userContext);
  } finally {
    lock.releaseLock();
  }

  sendTelegramNotification_(result.notification);
  return buildClientState_(userContext, result);
}

function testTelegramNotification() {
  sendTelegramNotification_('Inventory bot is connected to Telegram.');
}

function applyMovement_(payload, userContext) {
  const ss = getSpreadsheet_();
  const stockSheet = requireSheet_(ss, SHEETS.CURRENT_STOCK);
  const transactionSheet = requireSheet_(ss, SHEETS.TRANSACTIONS);
  const stockValues = stockSheet.getDataRange().getValues();
  const stockId = makeStockId_(payload.godown, payload.item);
  const now = new Date();

  let rowIndex = -1;
  for (let i = 1; i < stockValues.length; i += 1) {
    if (String(stockValues[i][0]) === stockId) {
      rowIndex = i + 1;
      break;
    }
  }

  if (rowIndex < 0) {
    throw new Error('Item not found in selected godown.');
  }

  const before = Number(stockSheet.getRange(rowIndex, 4).getValue()) || 0;
  let after;
  let quantityChange;

  if (payload.action === 'IN') {
    quantityChange = payload.quantity;
    after = before + payload.quantity;
  } else if (payload.action === 'OUT') {
    quantityChange = -payload.quantity;
    after = before - payload.quantity;
    if (after < 0) {
      throw new Error('Stock cannot go negative. Current stock is ' + before + ' cases.');
    }
  } else if (payload.action === 'ADJUST') {
    after = payload.quantity;
    quantityChange = after - before;
  } else {
    throw new Error('Unsupported action.');
  }

  stockSheet.getRange(rowIndex, 4, 1, 3).setValues([[
    after,
    now,
    userContext.displayName
  ]]);

  const requestId = Utilities.getUuid();
  transactionSheet.appendRow([
    now,
    payload.action,
    payload.godown,
    payload.item,
    quantityChange,
    before,
    after,
    payload.note || '',
    String(userContext.user.id),
    userContext.displayName,
    userContext.platform || '',
    requestId
  ]);

  const signedChange = quantityChange > 0 ? '+' + quantityChange : String(quantityChange);
  const notification = [
    '<b>' + escapeHtml_(payload.action_LABEL || actionLabel_(payload.action)) + '</b>',
    '<b>Item:</b> ' + escapeHtml_(payload.item),
    '<b>Godown:</b> ' + escapeHtml_(payload.godown),
    '<b>Change:</b> ' + escapeHtml_(signedChange) + ' cases',
    '<b>Stock:</b> ' + escapeHtml_(before) + ' -> ' + escapeHtml_(after) + ' cases',
    '<b>By:</b> ' + escapeHtml_(userContext.displayName)
  ].join('\n');

  return {
    requestId: requestId,
    notification: notification,
    lastMovement: {
      action: payload.action,
      godown: payload.godown,
      item: payload.item,
      quantityChange: quantityChange,
      before: before,
      after: after
    }
  };
}

function buildClientState_(userContext, result) {
  const ss = getSpreadsheet_();
  const stockSheet = requireSheet_(ss, SHEETS.CURRENT_STOCK);
  const transactionSheet = requireSheet_(ss, SHEETS.TRANSACTIONS);
  const stockValues = stockSheet.getDataRange().getValues();
  const transactionValues = transactionSheet.getDataRange().getValues();
  const stock = [];
  const godownMap = {};

  for (let i = 1; i < stockValues.length; i += 1) {
    const row = stockValues[i];
    if (!row[1] || !row[2]) {
      continue;
    }

    const godown = String(row[1]).trim();
    const item = String(row[2]).trim();
    const quantity = Number(row[3]) || 0;
    godownMap[godown] = true;
    stock.push({
      godown: godown,
      item: item,
      quantity: quantity
    });
  }

  const recentTransactions = [];
  const start = Math.max(1, transactionValues.length - 10);
  for (let i = transactionValues.length - 1; i >= start; i -= 1) {
    const row = transactionValues[i];
    if (!row[0]) {
      continue;
    }
    recentTransactions.push({
      timestamp: row[0] instanceof Date ? row[0].toISOString() : String(row[0]),
      action: String(row[1] || ''),
      godown: String(row[2] || ''),
      item: String(row[3] || ''),
      quantityChange: Number(row[4]) || 0,
      after: Number(row[6]) || 0,
      user: String(row[9] || '')
    });
  }

  return {
    ok: true,
    user: {
      id: String(userContext.user.id),
      name: userContext.displayName
    },
    godowns: Object.keys(godownMap).sort(),
    stock: stock.sort(function (a, b) {
      return a.godown.localeCompare(b.godown) || a.item.localeCompare(b.item);
    }),
    recentTransactions: recentTransactions,
    result: result || null
  };
}

function validatePayload_(payload) {
  if (!payload || typeof payload !== 'object') {
    throw new Error('Missing stock update.');
  }

  const action = String(payload.action || '').trim().toUpperCase();
  const godown = String(payload.godown || '').trim();
  const item = String(payload.item || '').trim();
  const quantity = Number(payload.quantity);
  const note = String(payload.note || '').trim().slice(0, 250);

  if (['IN', 'OUT', 'ADJUST'].indexOf(action) < 0) {
    throw new Error('Choose Stock In, Stock Out, or Set Stock.');
  }
  if (!godown) {
    throw new Error('Choose a godown.');
  }
  if (!item) {
    throw new Error('Choose an item.');
  }
  if (!isFinite(quantity) || quantity < 0) {
    throw new Error('Enter a valid quantity.');
  }
  if ((action === 'IN' || action === 'OUT') && quantity <= 0) {
    throw new Error('Quantity must be greater than zero.');
  }

  return {
    action: action,
    godown: godown,
    item: item,
    quantity: quantity,
    note: note
  };
}

function requireAllowedTelegramUser_(initData) {
  const context = validateTelegramInitData_(initData);
  const ss = getSpreadsheet_();
  const usersSheet = requireSheet_(ss, SHEETS.USERS);
  const users = usersSheet.getDataRange().getValues();
  const targetId = normalizeId_(context.user.id);

  for (let i = 1; i < users.length; i += 1) {
    const row = users[i];
    const userId = normalizeId_(row[0]);
    const active = row[3] === true || String(row[3]).toUpperCase() === 'TRUE';
    if (userId === targetId && active) {
      context.displayName = String(row[1] || context.user.first_name || context.user.username || targetId);
      return context;
    }
  }

  throw new Error('Access denied. This Telegram user is not allowed to edit inventory.');
}

function validateTelegramInitData_(initData) {
  const token = getRequiredProperty_('BOT_TOKEN');
  if (!initData) {
    throw new Error('Open this inventory app inside Telegram.');
  }

  const parsed = parseQueryString_(initData);
  const receivedHash = parsed.hash;
  if (!receivedHash) {
    throw new Error('Telegram login data is missing its hash.');
  }

  const dataCheckPairs = [];
  Object.keys(parsed).sort().forEach(function (key) {
    if (key !== 'hash') {
      dataCheckPairs.push(key + '=' + parsed[key]);
    }
  });

  const dataCheckString = dataCheckPairs.join('\n');
  const secretKey = Utilities.computeHmacSha256Signature(token, 'WebAppData');
  const dataBytes = Utilities.newBlob(dataCheckString, 'UTF-8').getBytes();
  const computedHashBytes = Utilities.computeHmacSignature(
    Utilities.MacAlgorithm.HMAC_SHA_256,
    dataBytes,
    secretKey
  );
  const computedHash = bytesToHex_(computedHashBytes);

  if (computedHash !== receivedHash) {
    throw new Error('Telegram login validation failed.');
  }

  const authDate = Number(parsed.auth_date || 0);
  const nowSeconds = Math.floor(Date.now() / 1000);
  if (!authDate || nowSeconds - authDate > 172800) {
    throw new Error('Telegram session expired. Close and reopen the Mini App.');
  }

  const user = parsed.user ? JSON.parse(parsed.user) : null;
  if (!user || !user.id) {
    throw new Error('Telegram user data is missing.');
  }

  return {
    user: user,
    platform: parsed.platform || ''
  };
}

function ensureSheet_(ss, name, headers) {
  let sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  } else if (sheet.getLastRow() === 0) {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  }
  return sheet;
}

function requireSheet_(ss, name) {
  const sheet = ss.getSheetByName(name);
  if (!sheet) {
    throw new Error('Missing sheet: ' + name + '. Run setupInventoryWorkbook first.');
  }
  return sheet;
}

function formatSheet_(sheet, columnCount) {
  sheet.setFrozenRows(1);
  sheet.getRange(1, 1, 1, columnCount)
    .setFontWeight('bold')
    .setBackground('#111827')
    .setFontColor('#ffffff');
}

function sendTelegramNotification_(htmlMessage) {
  const token = getRequiredProperty_('BOT_TOKEN');
  const chatId = getRequiredProperty_('TELEGRAM_CHAT_ID');
  const url = 'https://api.telegram.org/bot' + token + '/sendMessage';
  const response = UrlFetchApp.fetch(url, {
    method: 'post',
    contentType: 'application/json',
    muteHttpExceptions: true,
    payload: JSON.stringify({
      chat_id: chatId,
      text: htmlMessage,
      parse_mode: 'HTML',
      disable_web_page_preview: true
    })
  });

  const code = response.getResponseCode();
  if (code < 200 || code >= 300) {
    throw new Error('Telegram notification failed: ' + response.getContentText());
  }
}

function getRequiredProperty_(key) {
  const value = PropertiesService.getScriptProperties().getProperty(key);
  if (!value) {
    throw new Error('Missing Script Property: ' + key);
  }
  return value;
}

function getSpreadsheet_() {
  const spreadsheetId = PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID');
  if (spreadsheetId) {
    return SpreadsheetApp.openById(spreadsheetId);
  }

  const active = SpreadsheetApp.getActive();
  if (!active) {
    throw new Error('Missing Script Property: SPREADSHEET_ID');
  }
  return active;
}

function makeStockId_(godown, item) {
  return String(godown).trim().toUpperCase() + '::' + String(item).trim().toUpperCase();
}

function actionLabel_(action) {
  if (action === 'IN') return 'Stock In';
  if (action === 'OUT') return 'Stock Out';
  if (action === 'ADJUST') return 'Set Stock';
  return action;
}

function normalizeId_(value) {
  if (typeof value === 'number') {
    return String(Math.trunc(value));
  }
  return String(value || '').trim();
}

function parseQueryString_(value) {
  const parsed = {};
  String(value).split('&').forEach(function (part) {
    const eq = part.indexOf('=');
    if (eq < 0) {
      return;
    }
    const key = decodeURIComponent(part.slice(0, eq).replace(/\+/g, ' '));
    const val = decodeURIComponent(part.slice(eq + 1).replace(/\+/g, ' '));
    parsed[key] = val;
  });
  return parsed;
}

function bytesToHex_(bytes) {
  return bytes.map(function (byte) {
    const normalized = byte < 0 ? byte + 256 : byte;
    return ('0' + normalized.toString(16)).slice(-2);
  }).join('');
}

function escapeHtml_(value) {
  return String(value)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;');
}
