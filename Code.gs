// ============================================================
// 記帳系統 — Google Apps Script Backend
// 支援 PWA API + LINE Bot Webhook
// 貼到 Google Apps Script 後，執行 setup() 一次即可初始化
// ============================================================

// ---- 設定 (透過 Script Properties 儲存，不要寫死在這裡) ----
function getConfig() {
  const props = PropertiesService.getScriptProperties();
  return {
    LINE_ACCESS_TOKEN: props.getProperty('LINE_ACCESS_TOKEN') || '',
  };
}

const SHEET_NAMES = {
  TRANSACTIONS: 'transactions',
  BUDGETS: 'budgets',
  CATEGORIES: 'categories',
};

// ============================================================
// 主要進入點
// ============================================================

function doGet(e) {
  const action = e.parameter.action || '';
  try {
    let result;
    const p = e.parameter;
    switch (action) {
      case 'getTransactions':
        result = getTransactions(p.year, p.month); break;
      case 'getSummary':
        result = getMonthlySummary(p.year, p.month); break;
      case 'getBudget':
        result = getBudget(p.year, p.month); break;
      case 'getCategories':
        result = getCategories(); break;
      case 'addTransaction':
        result = addTransaction({
          date: p.date, type: p.type, category: p.category,
          amount: parseFloat(p.amount), note: p.note || '',
        }); break;
      case 'deleteTransaction':
        result = deleteTransaction(p.id); break;
      case 'setBudget':
        result = setBudget(p.year, p.month, parseFloat(p.amount)); break;
      default:
        result = { status: 'ok', message: 'API running' };
    }
    return jsonOk(result);
  } catch (err) {
    return jsonErr(err.message);
  }
}

function doPost(e) {
  const raw = e.postData.contents;
  let body;
  try { body = JSON.parse(raw); } catch (_) { body = {}; }

  // LINE Webhook 特徵：有 events 陣列
  if (body.events !== undefined) {
    return handleLineWebhook(body);
  }

  // PWA API POST（備用，現在改用 GET）
  try {
    return jsonOk({ echoed: body });
  } catch (err) {
    return jsonErr(err.message);
  }
}

// ============================================================
// Google Sheets CRUD
// ============================================================

function getSheet(name) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    initSheet(sheet, name);
  }
  return sheet;
}

function initSheet(sheet, name) {
  if (name === SHEET_NAMES.TRANSACTIONS) {
    sheet.getRange(1, 1, 1, 7)
      .setValues([['ID', '日期', '類型', '分類', '金額', '備註', '時間戳']]);
  } else if (name === SHEET_NAMES.BUDGETS) {
    sheet.getRange(1, 1, 1, 3).setValues([['年', '月', '預算']]);
  } else if (name === SHEET_NAMES.CATEGORIES) {
    sheet.getRange(1, 1, 1, 3).setValues([['名稱', 'Emoji', '顏色']]);
    const defaults = [
      ['餐飲', '🍱', '#FF6B6B'], ['交通', '🚌', '#4ECDC4'],
      ['購物', '🛒', '#45B7D1'], ['居家', '🏠', '#96CEB4'],
      ['醫療', '💊', '#FF6B9D'], ['娛樂', '🎮', '#C7B9FF'],
      ['教育', '📚', '#FFD93D'], ['薪資', '💼', '#10B981'],
      ['其他', '📦', '#95A5A6'],
    ];
    sheet.getRange(2, 1, defaults.length, 3).setValues(defaults);
  }
}

function addTransaction(data) {
  const sheet = getSheet(SHEET_NAMES.TRANSACTIONS);
  const id = Utilities.getUuid();
  const ts = Utilities.formatDate(new Date(), 'Asia/Taipei', "yyyy-MM-dd'T'HH:mm:ss");
  sheet.appendRow([id, data.date, data.type, data.category, data.amount, data.note || '', ts]);
  return { id, ...data, timestamp: ts };
}

function getTransactions(year, month) {
  const sheet = getSheet(SHEET_NAMES.TRANSACTIONS);
  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return [];

  const prefix = year + '-' + String(month).padStart(2, '0');
  const results = [];
  for (let i = 1; i < data.length; i++) {
    const r = data[i];
    if (!r[0]) continue;
    if (String(r[1]).startsWith(prefix)) {
      results.push({ id: r[0], date: r[1], type: r[2], category: r[3],
                     amount: Number(r[4]), note: r[5], timestamp: r[6] });
    }
  }
  return results.sort((a, b) => b.date.localeCompare(a.date));
}

function getMonthlySummary(year, month) {
  const transactions = getTransactions(year, month);
  const summary = { totalIncome: 0, totalExpense: 0, byCategory: {}, dailyExpense: {} };
  transactions.forEach(t => {
    if (t.type === '收入') {
      summary.totalIncome += t.amount;
    } else {
      summary.totalExpense += t.amount;
      summary.byCategory[t.category] = (summary.byCategory[t.category] || 0) + t.amount;
      summary.dailyExpense[t.date] = (summary.dailyExpense[t.date] || 0) + t.amount;
    }
  });
  return summary;
}

function deleteTransaction(id) {
  const sheet = getSheet(SHEET_NAMES.TRANSACTIONS);
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === id) { sheet.deleteRow(i + 1); return { deleted: id }; }
  }
  throw new Error('找不到此筆記錄: ' + id);
}

function getBudget(year, month) {
  const sheet = getSheet(SHEET_NAMES.BUDGETS);
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] == year && data[i][1] == month) {
      return { year: Number(year), month: Number(month), amount: data[i][2] };
    }
  }
  return { year: Number(year), month: Number(month), amount: 0 };
}

function setBudget(year, month, amount) {
  const sheet = getSheet(SHEET_NAMES.BUDGETS);
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] == year && data[i][1] == month) {
      sheet.getRange(i + 1, 3).setValue(amount);
      return { year: Number(year), month: Number(month), amount };
    }
  }
  sheet.appendRow([year, month, amount]);
  return { year: Number(year), month: Number(month), amount };
}

function getCategories() {
  const sheet = getSheet(SHEET_NAMES.CATEGORIES);
  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return [];
  return data.slice(1).map(r => ({ name: r[0], emoji: r[1], color: r[2] }));
}

// ============================================================
// LINE Bot
// ============================================================

const CATEGORY_KEYWORDS = {
  '餐飲':  ['早餐','午餐','晚餐','消夜','飲料','咖啡','奶茶','便當','餐廳','火鍋','麥當勞',
             '肯德基','摩斯','星巴克','路易莎','711','全家','萊爾富','吃飯','食物','茶'],
  '交通':  ['捷運','公車','計程車','uber','油費','停車','高鐵','台鐵','火車','客運','機票',
             '加油','gogoro','悠遊卡','youbike'],
  '購物':  ['超市','全聯','好市多','costco','網購','蝦皮','pchome','衣服','鞋子','電器','百貨'],
  '居家':  ['房租','水費','電費','瓦斯','網路費','電話費','修繕','家具','日用品','清潔'],
  '醫療':  ['醫院','診所','藥局','藥','掛號','健保','看診','牙醫','眼科'],
  '娛樂':  ['電影','遊戲','書','音樂','ktv','netflix','旅遊','旅行','住宿','門票','concert'],
  '教育':  ['學費','補習','課程','書本','教材','訂閱','udemy'],
  '薪資':  ['薪水','薪資','獎金','加班費','副業','稿費','租金收入','利息','股息','收款'],
};

function autoCategory(text) {
  const t = text.toLowerCase();
  for (const [cat, kws] of Object.entries(CATEGORY_KEYWORDS)) {
    if (kws.some(kw => t.includes(kw.toLowerCase()))) return cat;
  }
  return '其他';
}

function parseLineMessage(text) {
  text = text.trim();

  // 純指令
  if (['本月','統計','月報'].includes(text)) return { cmd: 'summary' };
  if (['今日','今天'].includes(text))         return { cmd: 'today' };
  if (['幫助','help','說明','?','？'].includes(text.toLowerCase())) return { cmd: 'help' };
  if (text === '預算')                        return { cmd: 'getBudget' };

  // 設定預算：「預算 10000」
  const budgetSet = text.match(/^預算\s+(\d+(?:\.\d+)?)$/);
  if (budgetSet) return { cmd: 'setBudget', amount: parseFloat(budgetSet[1]) };

  // 收入：「收入 薪水 50000」或「收入 50000」
  const incFull = text.match(/^收入\s+(.+)\s+(\d+(?:\.\d+)?)$/);
  if (incFull) {
    return { cmd: 'add', type: '收入', note: incFull[1],
             amount: parseFloat(incFull[2]), category: autoCategory(incFull[1]) };
  }
  const incShort = text.match(/^收入\s+(\d+(?:\.\d+)?)$/);
  if (incShort) {
    return { cmd: 'add', type: '收入', note: '收入', amount: parseFloat(incShort[1]), category: '薪資' };
  }

  // 支出：「午餐 85」或「星巴克拿鐵 120」
  const expMatch = text.match(/^(.+)\s+(\d+(?:\.\d+)?)$/);
  if (expMatch) {
    const amt = parseFloat(expMatch[2]);
    if (amt > 0) {
      return { cmd: 'add', type: '支出', note: expMatch[1],
               amount: amt, category: autoCategory(expMatch[1]) };
    }
  }

  return { cmd: 'unknown' };
}

function handleLineWebhook(body) {
  const cfg = getConfig();
  (body.events || []).forEach(event => {
    if (event.type !== 'message' || event.message.type !== 'text') return;

    const parsed = parseLineMessage(event.message.text);
    let reply = '';

    if (parsed.cmd === 'add') {
      const today = Utilities.formatDate(new Date(), 'Asia/Taipei', 'yyyy-MM-dd');
      addTransaction({ date: today, type: parsed.type, category: parsed.category,
                       amount: parsed.amount, note: parsed.note });
      const sign = parsed.type === '支出' ? '-' : '+';
      const icon = parsed.type === '支出' ? '💸' : '💰';
      reply = `✅ 已記錄\n📅 ${today}\n${icon} ${parsed.category}・${parsed.note}\n${sign}$${parsed.amount}`;

    } else if (parsed.cmd === 'summary') {
      const now = new Date();
      const y = now.getFullYear(), m = now.getMonth() + 1;
      const s = getMonthlySummary(y, m);
      const b = getBudget(y, m);
      let bLine = '';
      if (b.amount > 0) {
        const pct = Math.round(s.totalExpense / b.amount * 100);
        bLine = `\n💳 預算 $${b.amount}（已用 ${pct}%）`;
      }
      const top = Object.entries(s.byCategory).sort((a,b)=>b[1]-a[1]).slice(0,3)
        .map(([c,a]) => `  ${c} $${a}`).join('\n');
      reply = `📊 ${y}年${m}月\n💸 支出 $${s.totalExpense}\n💰 收入 $${s.totalIncome}${bLine}\n\n前三大：\n${top || '  (尚無)'}`;

    } else if (parsed.cmd === 'today') {
      const today = Utilities.formatDate(new Date(), 'Asia/Taipei', 'yyyy-MM-dd');
      const now = new Date();
      const txs = getTransactions(now.getFullYear(), now.getMonth()+1).filter(t => t.date === today);
      const total = txs.filter(t=>t.type==='支出').reduce((s,t)=>s+t.amount, 0);
      if (txs.length === 0) {
        reply = '📅 今日尚無記錄';
      } else {
        const lines = txs.map(t => `  ${t.category} ${t.note} $${t.amount}`).join('\n');
        reply = `📅 今日消費 $${total}\n${lines}`;
      }

    } else if (parsed.cmd === 'getBudget') {
      const now = new Date();
      const b = getBudget(now.getFullYear(), now.getMonth()+1);
      const s = getMonthlySummary(now.getFullYear(), now.getMonth()+1);
      if (b.amount > 0) {
        const rem = b.amount - s.totalExpense;
        const pct = Math.round(s.totalExpense / b.amount * 100);
        reply = `💳 本月預算 $${b.amount}\n💸 已用 $${s.totalExpense}（${pct}%）\n✅ 剩餘 $${rem}`;
      } else {
        reply = '💳 本月尚未設定預算\n\n發送「預算 10000」即可設定';
      }

    } else if (parsed.cmd === 'setBudget') {
      const now = new Date();
      setBudget(now.getFullYear(), now.getMonth()+1, parsed.amount);
      reply = `✅ 本月預算已設為 $${parsed.amount}`;

    } else if (parsed.cmd === 'help') {
      reply = '📖 記帳機器人\n\n【記帳】\n午餐 85\n星巴克咖啡 120\n捷運 30\n收入 薪水 50000\n\n【查詢】\n本月 → 月統計\n今日 → 今日明細\n預算 → 查詢預算\n預算 10000 → 設定預算\n\n更多功能請開啟 App ✨';

    } else {
      reply = '❓ 無法識別\n\n試試：午餐 85\n或輸入「說明」查看完整用法';
    }

    replyToLine(cfg.LINE_ACCESS_TOKEN, event.replyToken, reply);
  });

  return ContentService.createTextOutput('OK');
}

function replyToLine(token, replyToken, text) {
  if (!token) return;
  UrlFetchApp.fetch('https://api.line.me/v2/bot/message/reply', {
    method: 'post',
    contentType: 'application/json',
    headers: { Authorization: 'Bearer ' + token },
    payload: JSON.stringify({ replyToken, messages: [{ type: 'text', text }] }),
    muteHttpExceptions: true,
  });
}

// ============================================================
// 工具函式
// ============================================================

function jsonOk(data) {
  return ContentService
    .createTextOutput(JSON.stringify({ success: true, data }))
    .setMimeType(ContentService.MimeType.JSON);
}

function jsonErr(msg) {
  return ContentService
    .createTextOutput(JSON.stringify({ success: false, error: msg }))
    .setMimeType(ContentService.MimeType.JSON);
}

// ============================================================
// 初始化（第一次使用前執行一次）
// ============================================================

function setup() {
  getSheet(SHEET_NAMES.TRANSACTIONS);
  getSheet(SHEET_NAMES.BUDGETS);
  getSheet(SHEET_NAMES.CATEGORIES);
  Logger.log('✅ Setup 完成！Spreadsheet ID: ' + SpreadsheetApp.getActiveSpreadsheet().getId());
  Logger.log('📌 請在 Script Properties 設定 LINE_ACCESS_TOKEN');
}
