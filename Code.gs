// ============================================================
// 記帳系統 — Google Apps Script Backend
// v5 — 帳戶管理 + 固定扣款 + 轉帳
// ============================================================

function getConfig() {
  var props = PropertiesService.getScriptProperties();
  return { LINE_ACCESS_TOKEN: props.getProperty('LINE_ACCESS_TOKEN') || '' };
}

var SHEET_NAMES = {
  TRANSACTIONS: 'transactions',
  BUDGETS: 'budgets',
  CATEGORIES: 'categories',
  ACCOUNTS: 'accounts',
  RECURRING: 'recurring',
};

var CATEGORY_KEYWORDS = {
  '餐飲': ['早餐','午餐','晚餐','消夜','飲料','咖啡','奶茶','便當','餐廳','火鍋','麥當勞','肯德基','摩斯','星巴克','路易莎','711','全家','萊爾富','吃飯','食物','茶'],
  '交通': ['捷運','公車','計程車','uber','油費','停車','高鐵','台鐵','火車','客運','機票','加油','gogoro','悠遊卡','youbike'],
  '購物': ['超市','全聯','好市多','costco','網購','蝦皮','pchome','衣服','鞋子','電器','百貨'],
  '居家': ['房租','水費','電費','瓦斯','網路費','電話費','修繕','家具','日用品','清潔'],
  '醫療': ['醫院','診所','藥局','藥','掛號','健保','看診','牙醫','眼科'],
  '娛樂': ['電影','遊戲','書','音樂','ktv','netflix','旅遊','旅行','住宿','門票','concert'],
  '教育': ['學費','補習','課程','書本','教材','訂閱','udemy'],
  '薪資': ['薪水','薪資','獎金','加班費','副業','稿費','租金收入','利息','股息','收款'],
};

// ============================================================
// 主要進入點
// ============================================================

function doGet(e) {
  var action = e.parameter.action || '';
  var p = e.parameter;
  try {
    var result;
    if (action === 'getTransactions') {
      result = getTransactions(p.year, p.month);
    } else if (action === 'getSummary') {
      result = getMonthlySummary(p.year, p.month);
    } else if (action === 'getBudget') {
      result = getBudget(p.year, p.month);
    } else if (action === 'getCategories') {
      result = getCategories();
    } else if (action === 'addTransaction') {
      result = addTransaction({ date: p.date, type: p.type, category: p.category, amount: parseFloat(p.amount), note: p.note || '', account: p.account || '' });
    } else if (action === 'deleteTransaction') {
      result = deleteTransaction(p.id);
    } else if (action === 'setBudget') {
      result = setBudget(p.year, p.month, parseFloat(p.amount));
    } else if (action === 'getAccounts') {
      result = getAccounts();
    } else if (action === 'getAccountBalances') {
      result = getAccountBalances();
    } else if (action === 'addAccount') {
      result = addAccount({ name: p.name, emoji: p.emoji, initialBalance: parseFloat(p.initialBalance) || 0, color: p.color });
    } else if (action === 'deleteAccount') {
      result = deleteAccount(p.id);
    } else if (action === 'updateAccount') {
      result = updateAccount(p.id, { initialBalance: parseFloat(p.initialBalance) });
    } else if (action === 'transfer') {
      result = transfer({ fromAccount: p.fromAccount, toAccount: p.toAccount, amount: parseFloat(p.amount), date: p.date, note: p.note || '' });
    } else if (action === 'getRecurring') {
      result = getRecurring();
    } else if (action === 'addRecurring') {
      result = addRecurring({ name: p.name, amount: parseFloat(p.amount), category: p.category, accountId: p.accountId, dayOfMonth: parseInt(p.dayOfMonth), type: p.type || '支出' });
    } else if (action === 'deleteRecurring') {
      result = deleteRecurring(p.id);
    } else if (action === 'toggleRecurring') {
      result = toggleRecurring(p.id);
    } else if (action === 'processRecurring') {
      result = processRecurring();
    } else if (action === 'debugTx') {
      var dSheet = getSheet(SHEET_NAMES.TRANSACTIONS);
      var dData = dSheet.getDataRange().getValues();
      var dPrefix = (p.year || '2026') + '-' + String(p.month || '7').padStart(2, '0');
      var dRows = [];
      for (var di = 1; di < Math.min(dData.length, 5); di++) {
        var dr = dData[di];
        var dRaw = dr[1];
        var dConverted = sheetDateStr(dRaw);
        dRows.push({ isDate: dRaw instanceof Date, raw: String(dRaw), converted: dConverted, matches: dConverted.startsWith(dPrefix) });
      }
      result = { prefix: dPrefix, rowCount: dData.length, rows: dRows };
    } else {
      result = { status: 'ok', message: 'API running v5' };
    }
    return jsonOk(result);
  } catch (err) {
    return jsonErr(err.message);
  }
}

function doPost(e) {
  var raw = e.postData.contents;
  var body;
  try { body = JSON.parse(raw); } catch (_) { body = {}; }
  if (body.events !== undefined) return handleLineWebhook(body);
  try { return jsonOk({ echoed: body }); } catch (err) { return jsonErr(err.message); }
}

// ============================================================
// Google Sheets CRUD
// ============================================================

function getSheet(name) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    initSheet(sheet, name);
  }
  return sheet;
}

function initSheet(sheet, name) {
  if (name === SHEET_NAMES.TRANSACTIONS) {
    sheet.getRange(1, 1, 1, 9).setValues([['ID', '日期', '類型', '分類', '金額', '備註', '時間戳', '帳戶ID', '轉入帳戶ID']]);
  } else if (name === SHEET_NAMES.BUDGETS) {
    sheet.getRange(1, 1, 1, 3).setValues([['年', '月', '預算']]);
  } else if (name === SHEET_NAMES.CATEGORIES) {
    sheet.getRange(1, 1, 1, 3).setValues([['名稱', 'Emoji', '顏色']]);
    sheet.getRange(2, 1, 9, 3).setValues([
      ['餐飲','🍱','#FF6B6B'], ['交通','🚌','#4ECDC4'], ['購物','🛒','#45B7D1'],
      ['居家','🏠','#96CEB4'], ['醫療','💊','#FF6B9D'], ['娛樂','🎮','#C7B9FF'],
      ['教育','📚','#FFD93D'], ['薪資','💼','#10B981'], ['其他','📦','#95A5A6'],
    ]);
  } else if (name === SHEET_NAMES.ACCOUNTS) {
    sheet.getRange(1, 1, 1, 5).setValues([['ID', '名稱', 'Emoji', '初始餘額', '顏色']]);
  } else if (name === SHEET_NAMES.RECURRING) {
    sheet.getRange(1, 1, 1, 8).setValues([['ID', '名稱', '金額', '分類', '帳戶ID', '每月日期', '啟用', '類型']]);
  }
}

function sheetDateStr(val) {
  if (val && typeof val.getTime === 'function') {
    return Utilities.formatDate(val, 'Asia/Taipei', 'yyyy-MM-dd');
  }
  return String(val);
}

// ============================================================
// Transactions
// ============================================================

function addTransaction(data) {
  var sheet = getSheet(SHEET_NAMES.TRANSACTIONS);
  var id = Utilities.getUuid();
  var ts = Utilities.formatDate(new Date(), 'Asia/Taipei', "yyyy-MM-dd'T'HH:mm:ss");
  var lastRow = sheet.getLastRow() + 1;
  sheet.appendRow([id, data.date, data.type, data.category, data.amount, data.note || '', ts, data.account || '', data.toAccount || '']);
  sheet.getRange(lastRow, 2).setNumberFormat('@STRING@');
  return { id: id, date: data.date, type: data.type, category: data.category, amount: data.amount, note: data.note, timestamp: ts, account: data.account || '', toAccount: data.toAccount || '' };
}

function getTransactions(year, month) {
  var sheet = getSheet(SHEET_NAMES.TRANSACTIONS);
  var data = sheet.getDataRange().getValues();
  if (data.length <= 1) return [];
  var prefix = year + '-' + String(month).padStart(2, '0');
  var results = [];
  for (var i = 1; i < data.length; i++) {
    var r = data[i];
    if (!r[0]) continue;
    var dateStr = sheetDateStr(r[1]);
    if (dateStr.startsWith(prefix)) {
      results.push({
        id: r[0], date: dateStr, type: r[2], category: r[3],
        amount: Number(r[4]), note: String(r[5] || ''), timestamp: r[6],
        account: String(r[7] || ''), toAccount: String(r[8] || ''),
      });
    }
  }
  return results.sort(function(a, b) { return b.date.localeCompare(a.date); });
}

function getMonthlySummary(year, month) {
  var transactions = getTransactions(year, month);
  var summary = { totalIncome: 0, totalExpense: 0, byCategory: {}, dailyExpense: {} };
  for (var i = 0; i < transactions.length; i++) {
    var t = transactions[i];
    if (t.type === '收入') {
      summary.totalIncome += t.amount;
    } else if (t.type === '支出') {
      summary.totalExpense += t.amount;
      summary.byCategory[t.category] = (summary.byCategory[t.category] || 0) + t.amount;
      summary.dailyExpense[t.date] = (summary.dailyExpense[t.date] || 0) + t.amount;
    }
  }
  return summary;
}

function deleteTransaction(id) {
  var sheet = getSheet(SHEET_NAMES.TRANSACTIONS);
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(id)) {
      sheet.deleteRow(i + 1);
      return { deleted: id };
    }
  }
  throw new Error('找不到此筆記錄: ' + id);
}

// ============================================================
// Budget
// ============================================================

function getBudget(year, month) {
  var sheet = getSheet(SHEET_NAMES.BUDGETS);
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (data[i][0] == year && data[i][1] == month) {
      return { year: Number(year), month: Number(month), amount: data[i][2] };
    }
  }
  return { year: Number(year), month: Number(month), amount: 0 };
}

function setBudget(year, month, amount) {
  var sheet = getSheet(SHEET_NAMES.BUDGETS);
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (data[i][0] == year && data[i][1] == month) {
      sheet.getRange(i + 1, 3).setValue(amount);
      return { year: Number(year), month: Number(month), amount: amount };
    }
  }
  sheet.appendRow([year, month, amount]);
  return { year: Number(year), month: Number(month), amount: amount };
}

function getCategories() {
  var sheet = getSheet(SHEET_NAMES.CATEGORIES);
  var data = sheet.getDataRange().getValues();
  if (data.length <= 1) return [];
  var result = [];
  for (var i = 1; i < data.length; i++) {
    result.push({ name: data[i][0], emoji: data[i][1], color: data[i][2] });
  }
  return result;
}

// ============================================================
// Accounts
// ============================================================

function getAccounts() {
  var sheet = getSheet(SHEET_NAMES.ACCOUNTS);
  var data = sheet.getDataRange().getValues();
  if (data.length <= 1) return [];
  var result = [];
  for (var i = 1; i < data.length; i++) {
    if (!data[i][0]) continue;
    result.push({ id: String(data[i][0]), name: String(data[i][1]), emoji: String(data[i][2]), initialBalance: Number(data[i][3]) || 0, color: String(data[i][4] || '#6366f1') });
  }
  return result;
}

function addAccount(data) {
  var sheet = getSheet(SHEET_NAMES.ACCOUNTS);
  var id = Utilities.getUuid();
  sheet.appendRow([id, data.name, data.emoji || '💳', Number(data.initialBalance) || 0, data.color || '#6366f1']);
  return { id: id, name: data.name, emoji: data.emoji || '💳', initialBalance: Number(data.initialBalance) || 0, color: data.color || '#6366f1' };
}

function updateAccount(id, data) {
  var sheet = getSheet(SHEET_NAMES.ACCOUNTS);
  var rows = sheet.getDataRange().getValues();
  for (var i = 1; i < rows.length; i++) {
    if (String(rows[i][0]) === String(id)) {
      if (data.initialBalance !== undefined && !isNaN(data.initialBalance)) {
        sheet.getRange(i + 1, 4).setValue(data.initialBalance);
      }
      return { updated: id };
    }
  }
  throw new Error('找不到帳戶: ' + id);
}

function deleteAccount(id) {
  var sheet = getSheet(SHEET_NAMES.ACCOUNTS);
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(id)) {
      sheet.deleteRow(i + 1);
      return { deleted: id };
    }
  }
  throw new Error('找不到帳戶: ' + id);
}

function getAccountBalances() {
  var accounts = getAccounts();
  if (!accounts.length) return [];
  var txSheet = getSheet(SHEET_NAMES.TRANSACTIONS);
  var txData = txSheet.getDataRange().getValues();
  var balances = {};
  for (var a = 0; a < accounts.length; a++) {
    balances[accounts[a].id] = accounts[a].initialBalance;
  }
  for (var i = 1; i < txData.length; i++) {
    var r = txData[i];
    if (!r[0]) continue;
    var type = String(r[2] || '');
    var amount = Number(r[4]) || 0;
    var acctId = String(r[7] || '');
    var toAcctId = String(r[8] || '');
    if (type === '收入' && balances.hasOwnProperty(acctId)) {
      balances[acctId] += amount;
    } else if (type === '支出' && balances.hasOwnProperty(acctId)) {
      balances[acctId] -= amount;
    } else if (type === '轉帳') {
      if (balances.hasOwnProperty(acctId)) balances[acctId] -= amount;
      if (balances.hasOwnProperty(toAcctId)) balances[toAcctId] += amount;
    }
  }
  return accounts.map(function(a) {
    return { id: a.id, name: a.name, emoji: a.emoji, color: a.color, balance: Math.round((balances[a.id] || 0) * 100) / 100 };
  });
}

// ============================================================
// Transfer
// ============================================================

function transfer(data) {
  if (!data.fromAccount || !data.toAccount) throw new Error('請選擇轉出與轉入帳戶');
  if (data.fromAccount === data.toAccount) throw new Error('轉出與轉入帳戶不能相同');
  var date = data.date || Utilities.formatDate(new Date(), 'Asia/Taipei', 'yyyy-MM-dd');
  var sheet = getSheet(SHEET_NAMES.TRANSACTIONS);
  var id = Utilities.getUuid();
  var ts = Utilities.formatDate(new Date(), 'Asia/Taipei', "yyyy-MM-dd'T'HH:mm:ss");
  var lastRow = sheet.getLastRow() + 1;
  sheet.appendRow([id, date, '轉帳', '轉帳', Number(data.amount), data.note || '', ts, data.fromAccount, data.toAccount]);
  sheet.getRange(lastRow, 2).setNumberFormat('@STRING@');
  return { id: id, date: date, type: '轉帳', amount: Number(data.amount), fromAccount: data.fromAccount, toAccount: data.toAccount };
}

// ============================================================
// Recurring
// ============================================================

function getRecurring() {
  var sheet = getSheet(SHEET_NAMES.RECURRING);
  var data = sheet.getDataRange().getValues();
  if (data.length <= 1) return [];
  var result = [];
  for (var i = 1; i < data.length; i++) {
    if (!data[i][0]) continue;
    result.push({
      id: String(data[i][0]), name: String(data[i][1]), amount: Number(data[i][2]),
      category: String(data[i][3]), accountId: String(data[i][4]),
      dayOfMonth: Number(data[i][5]), enabled: data[i][6] === true || String(data[i][6]) === 'TRUE',
      type: String(data[i][7] || '支出'),
    });
  }
  return result;
}

function addRecurring(data) {
  var sheet = getSheet(SHEET_NAMES.RECURRING);
  var id = Utilities.getUuid();
  sheet.appendRow([id, data.name, Number(data.amount), data.category, data.accountId || '', Number(data.dayOfMonth), true, data.type || '支出']);
  return { id: id, name: data.name, amount: Number(data.amount), category: data.category, accountId: data.accountId || '', dayOfMonth: Number(data.dayOfMonth), enabled: true, type: data.type || '支出' };
}

function deleteRecurring(id) {
  var sheet = getSheet(SHEET_NAMES.RECURRING);
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(id)) {
      sheet.deleteRow(i + 1);
      return { deleted: id };
    }
  }
  throw new Error('找不到固定扣款: ' + id);
}

function toggleRecurring(id) {
  var sheet = getSheet(SHEET_NAMES.RECURRING);
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(id)) {
      var current = data[i][6] === true || String(data[i][6]) === 'TRUE';
      sheet.getRange(i + 1, 7).setValue(!current);
      return { id: id, enabled: !current };
    }
  }
  throw new Error('找不到固定扣款: ' + id);
}

function processRecurring() {
  var now = new Date();
  var today = Utilities.formatDate(now, 'Asia/Taipei', 'yyyy-MM-dd');
  var dayOfMonth = Number(Utilities.formatDate(now, 'Asia/Taipei', 'd'));
  var items = getRecurring().filter(function(r) { return r.enabled && r.dayOfMonth === dayOfMonth; });
  if (!items.length) return { processed: 0, date: today };
  // Check already-processed today
  var txSheet = getSheet(SHEET_NAMES.TRANSACTIONS);
  var txData = txSheet.getDataRange().getValues();
  var todayNotes = {};
  for (var i = 1; i < txData.length; i++) {
    if (sheetDateStr(txData[i][1]) === today) {
      todayNotes[String(txData[i][5] || '')] = true;
    }
  }
  var processed = 0;
  for (var j = 0; j < items.length; j++) {
    var r = items[j];
    var marker = '[自動]' + r.id;
    if (!todayNotes[marker]) {
      addTransaction({ date: today, type: r.type, category: r.category, amount: r.amount, note: marker, account: r.accountId });
      processed++;
    }
  }
  return { processed: processed, date: today };
}

// ============================================================
// Time trigger (執行一次 setupTrigger() 即可設定每日自動執行)
// ============================================================

function setupTrigger() {
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === 'dailyProcessRecurring') {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }
  ScriptApp.newTrigger('dailyProcessRecurring').timeBased().atHour(8).everyDays(1).inTimezone('Asia/Taipei').create();
  Logger.log('✅ 每日觸發器設定完成（每天 08:00 台北時間）');
}

function dailyProcessRecurring() {
  var result = processRecurring();
  Logger.log('固定扣款處理: ' + JSON.stringify(result));
}

// ============================================================
// LINE Bot
// ============================================================

function autoCategory(text) {
  var t = text.toLowerCase();
  var cats = Object.keys(CATEGORY_KEYWORDS);
  for (var i = 0; i < cats.length; i++) {
    var kws = CATEGORY_KEYWORDS[cats[i]];
    for (var j = 0; j < kws.length; j++) {
      if (t.indexOf(kws[j].toLowerCase()) !== -1) return cats[i];
    }
  }
  return '其他';
}

function parseLineMessage(text) {
  text = text.trim();
  if (text === '本月' || text === '統計' || text === '月報') return { cmd: 'summary' };
  if (text === '今日' || text === '今天') return { cmd: 'today' };
  if (text === '幫助' || text === 'help' || text === '說明' || text === '?' || text === '？') return { cmd: 'help' };
  if (text === '預算') return { cmd: 'getBudget' };
  var budgetSet = text.match(/^預算\s+(\d+(?:\.\d+)?)$/);
  if (budgetSet) return { cmd: 'setBudget', amount: parseFloat(budgetSet[1]) };
  var incFull = text.match(/^收入\s+(.+)\s+(\d+(?:\.\d+)?)$/);
  if (incFull) return { cmd: 'add', type: '收入', note: incFull[1], amount: parseFloat(incFull[2]), category: autoCategory(incFull[1]) };
  var incShort = text.match(/^收入\s+(\d+(?:\.\d+)?)$/);
  if (incShort) return { cmd: 'add', type: '收入', note: '收入', amount: parseFloat(incShort[1]), category: '薪資' };
  var expMatch = text.match(/^(.+)\s+(\d+(?:\.\d+)?)$/);
  if (expMatch && parseFloat(expMatch[2]) > 0) return { cmd: 'add', type: '支出', note: expMatch[1], amount: parseFloat(expMatch[2]), category: autoCategory(expMatch[1]) };
  return { cmd: 'unknown' };
}

function handleLineWebhook(body) {
  var cfg = getConfig();
  var events = body.events || [];
  for (var i = 0; i < events.length; i++) {
    var event = events[i];
    if (event.type !== 'message' || event.message.type !== 'text') continue;
    var parsed = parseLineMessage(event.message.text);
    var reply = '';
    if (parsed.cmd === 'add') {
      var today = Utilities.formatDate(new Date(), 'Asia/Taipei', 'yyyy-MM-dd');
      addTransaction({ date: today, type: parsed.type, category: parsed.category, amount: parsed.amount, note: parsed.note, account: '' });
      var sign = parsed.type === '支出' ? '-' : '+';
      var icon = parsed.type === '支出' ? '💸' : '💰';
      reply = '✅ 已記錄\n📅 ' + today + '\n' + icon + ' ' + parsed.category + '・' + parsed.note + '\n' + sign + '$' + parsed.amount;
    } else if (parsed.cmd === 'summary') {
      var now = new Date(); var y = now.getFullYear(); var m = now.getMonth() + 1;
      var s = getMonthlySummary(y, m); var b = getBudget(y, m); var bLine = '';
      if (b.amount > 0) bLine = '\n💳 預算 $' + b.amount + '（已用 ' + Math.round(s.totalExpense / b.amount * 100) + '%）';
      var topEntries = Object.keys(s.byCategory).map(function(c) { return [c, s.byCategory[c]]; });
      topEntries.sort(function(a, b) { return b[1] - a[1]; });
      var top = topEntries.slice(0, 3).map(function(e) { return '  ' + e[0] + ' $' + e[1]; }).join('\n');
      reply = '📊 ' + y + '年' + m + '月\n💸 支出 $' + s.totalExpense + '\n💰 收入 $' + s.totalIncome + bLine + '\n\n前三大：\n' + (top || '  (尚無)');
    } else if (parsed.cmd === 'today') {
      var today2 = Utilities.formatDate(new Date(), 'Asia/Taipei', 'yyyy-MM-dd');
      var now2 = new Date(); var txs = getTransactions(now2.getFullYear(), now2.getMonth() + 1);
      var todayTxs = txs.filter(function(t) { return t.date === today2; });
      var total = todayTxs.filter(function(t) { return t.type === '支出'; }).reduce(function(s, t) { return s + t.amount; }, 0);
      reply = todayTxs.length === 0 ? '📅 今日尚無記錄' : '📅 今日消費 $' + total + '\n' + todayTxs.map(function(t) { return '  ' + t.category + ' ' + t.note + ' $' + t.amount; }).join('\n');
    } else if (parsed.cmd === 'getBudget') {
      var now3 = new Date(); var b2 = getBudget(now3.getFullYear(), now3.getMonth() + 1); var s2 = getMonthlySummary(now3.getFullYear(), now3.getMonth() + 1);
      reply = b2.amount > 0 ? '💳 本月預算 $' + b2.amount + '\n💸 已用 $' + s2.totalExpense + '（' + Math.round(s2.totalExpense / b2.amount * 100) + '%）\n✅ 剩餘 $' + (b2.amount - s2.totalExpense) : '💳 本月尚未設定預算\n\n發送「預算 10000」即可設定';
    } else if (parsed.cmd === 'setBudget') {
      var now4 = new Date(); setBudget(now4.getFullYear(), now4.getMonth() + 1, parsed.amount);
      reply = '✅ 本月預算已設為 $' + parsed.amount;
    } else if (parsed.cmd === 'help') {
      reply = '📖 記帳機器人\n\n【記帳】\n午餐 85\n星巴克咖啡 120\n捷運 30\n收入 薪水 50000\n\n【查詢】\n本月 → 月統計\n今日 → 今日明細\n預算 → 查詢預算\n預算 10000 → 設定預算\n\n更多功能請開啟 App ✨';
    } else {
      reply = '❓ 無法識別\n\n試試：午餐 85\n或輸入「說明」查看完整用法';
    }
    replyToLine(cfg.LINE_ACCESS_TOKEN, event.replyToken, reply);
  }
  return ContentService.createTextOutput('OK');
}

function replyToLine(token, replyToken, text) {
  if (!token) return;
  UrlFetchApp.fetch('https://api.line.me/v2/bot/message/reply', {
    method: 'post', contentType: 'application/json',
    headers: { Authorization: 'Bearer ' + token },
    payload: JSON.stringify({ replyToken: replyToken, messages: [{ type: 'text', text: text }] }),
    muteHttpExceptions: true,
  });
}

// ============================================================
// 工具函式
// ============================================================

function jsonOk(data) {
  return ContentService.createTextOutput(JSON.stringify({ success: true, data: data })).setMimeType(ContentService.MimeType.JSON);
}

function jsonErr(msg) {
  return ContentService.createTextOutput(JSON.stringify({ success: false, error: msg })).setMimeType(ContentService.MimeType.JSON);
}

// ============================================================
// 初始化（第一次使用前執行一次）
// ============================================================

function setup() {
  getSheet(SHEET_NAMES.TRANSACTIONS);
  getSheet(SHEET_NAMES.BUDGETS);
  getSheet(SHEET_NAMES.CATEGORIES);
  getSheet(SHEET_NAMES.ACCOUNTS);
  getSheet(SHEET_NAMES.RECURRING);
  Logger.log('✅ Setup 完成');
  Logger.log('📌 請在 Script Properties 設定 LINE_ACCESS_TOKEN');
  Logger.log('📌 執行 setupTrigger() 可設定固定扣款每日自動執行');
}
