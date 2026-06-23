// ============================================================
// 便當評分系統 - Google Apps Script
// 部署為 Web App: 執行身分「我」, 存取權「所有人」
// ============================================================

const VERSION = '1.1.3';

const SHEET_ID = PropertiesService.getScriptProperties().getProperty('SHEET_ID');
const ORDER_SHEET_ID = PropertiesService.getScriptProperties().getProperty('ORDER_SHEET_ID');
const RATINGS_SHEET = 'ratings';
const MENU_SHEET = 'menu';

function getSheet(name) {
  return SpreadsheetApp.openById(SHEET_ID).getSheetByName(name);
}

// ── GET：讀取資料 ──────────────────────────────────────────
function doGet(e) {
  const action = e.parameter.action;
  let result;

  try {
    if (action === 'syncTodayMenu') {
      result = syncTodayMenu(e.parameter.date);
    } else if (action === 'setupDailyTrigger') {
      result = setupDailyTrigger();
    } else if (action === 'listTriggers') {
      result = listTriggers();
    } else if (action === 'getMenu') {
      result = getMenu(e.parameter.date);
    } else if (action === 'getRatings') {
      result = getRatings(e.parameter.store);
    } else if (action === 'getStores') {
      result = getStores();
    } else if (action === 'getTodayStores') {
      result = getTodayStores(e.parameter.date);
    } else {
      result = { error: 'Unknown action' };
    }
  } catch (err) {
    result = { error: err.message };
  }

  return ContentService
    .createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}

// ── POST：新增評分 / 新增菜單 ──────────────────────────────
function doPost(e) {
  let result;
  try {
    const data = JSON.parse(e.postData.contents);
    if (data.action === 'addRating') {
      result = addRating(data);
    } else if (data.action === 'addMenu') {
      result = addMenu(data);
    } else if (data.action === 'deleteRating') {
      result = deleteRating(data);
    } else {
      result = { error: 'Unknown action' };
    }
  } catch (err) {
    result = { error: err.message };
  }

  return ContentService
    .createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}

// ── 取得今日供應店家（依餐別分組）──────────────────────────
function getTodayStores(dateStr) {
  const sheet = getSheet(MENU_SHEET);
  const data = sheet.getDataRange().getValues();
  const today = dateStr || formatDate(new Date());

  const groups = {};
  for (let i = 1; i < data.length; i++) {
    const [date, store, dish, mealType] = data[i];
    if (!store || !dish) continue;
    const d = date ? formatDate(new Date(date)) : '';
    if (d !== today) continue;
    const meal = String(mealType || '').trim();
    if (!groups[meal]) groups[meal] = {};
    if (!groups[meal][store]) groups[meal][store] = [];
    groups[meal][store].push(dish);
  }
  return { date: today, groups };
}

// ── 取得菜單（可篩選日期，每店去重）────────────────────────
function getMenu(dateFilter) {
  const sheet = getSheet(MENU_SHEET);
  const data = sheet.getDataRange().getValues();
  const result = {};
  const seen = {}; // store → Set，避免同店重複菜品

  for (let i = 1; i < data.length; i++) {
    const [date, store, dish] = data[i];
    if (!store || !dish) continue;
    if (dateFilter) {
      const d = date ? formatDate(new Date(date)) : '';
      if (d !== dateFilter) continue;
    }
    if (!result[store]) { result[store] = []; seen[store] = {}; }
    const key = String(dish).trim();
    if (seen[store][key]) continue;
    seen[store][key] = true;
    result[store].push(dish);
  }
  return result;
}

// ── 取得所有店家（含均分）──────────────────────────────────
function getStores() {
  const sheet = getSheet(RATINGS_SHEET);
  const data = sheet.getDataRange().getValues();
  const stores = {};

  for (let i = 1; i < data.length; i++) {
    const [date, rater, store, dish, sides, score, comment] = data[i];
    if (!store || score === '' || score === null) continue;
    if (!stores[store]) stores[store] = { total: 0, count: 0, reviews: [] };
    stores[store].total += Number(score);
    stores[store].count += 1;
    stores[store].reviews.push({ date: formatDate(new Date(date)), rater, dish, sides, score: Number(score), comment });
  }

  return Object.entries(stores).map(([name, d]) => ({
    name,
    avg: Math.round((d.total / d.count) * 10) / 10,
    count: d.count,
    reviews: d.reviews.sort((a, b) => b.date.localeCompare(a.date)).slice(0, 20)
  })).sort((a, b) => b.avg - a.avg);
}

// ── 取得特定店家的評分 ─────────────────────────────────────
function getRatings(store) {
  const sheet = getSheet(RATINGS_SHEET);
  const data = sheet.getDataRange().getValues();
  const reviews = [];

  for (let i = 1; i < data.length; i++) {
    const [date, rater, storeName, dish, sides, score, comment] = data[i];
    if (!storeName || storeName !== store) continue;
    reviews.push({
      date: date ? formatDate(new Date(date)) : '',
      rater: rater || '',
      dish: dish || '',
      sides: sides || '',
      score: score !== '' && score !== null ? Number(score) : null,
      comment: comment || ''
    });
  }
  return reviews.sort((a, b) => b.date.localeCompare(a.date));
}

// ── 新增評分 ───────────────────────────────────────────────
function addRating(data) {
  const sheet = getSheet(RATINGS_SHEET);
  const date = data.date || formatDate(new Date());
  sheet.appendRow([
    date,
    data.rater || '',
    data.store,
    data.dish,
    data.sides || '',
    data.score,
    data.comment || ''
  ]);
  return { success: true };
}

// ── 新增菜單（Google Form Webhook 或手動）──────────────────
function addMenu(data) {
  const sheet = getSheet(MENU_SHEET);
  sheet.appendRow([
    data.date || '',
    data.store,
    data.dish
  ]);
  return { success: true };
}

// ── 刪除評分 ───────────────────────────────────────────────
const ADMIN_PASSWORD = PropertiesService.getScriptProperties().getProperty('ADMIN_PASSWORD');

function deleteRating(data) {
  if (data.adminPassword !== ADMIN_PASSWORD) return { error: '密碼錯誤' };

  const sheet = getSheet(RATINGS_SHEET);
  const rows = sheet.getDataRange().getValues();

  for (let i = rows.length - 1; i >= 1; i--) {
    const [date, rater, store, dish, , score] = rows[i];
    if (
      formatDate(new Date(date)) === data.date &&
      store === data.store &&
      dish === data.dish &&
      String(rater) === String(data.rater || '') &&
      Number(score) === Number(data.score)
    ) {
      sheet.deleteRow(i + 1);
      return { success: true };
    }
  }
  return { error: '找不到該筆評價' };
}

// ── 從餐點確認 Sheet 同步今日菜單 ──────────────────────────
// 讀取 ORDER_SHEET_ID 中符合今日日期的分頁（如 0616午餐）
// 欄位標題格式：「午餐店家一：店名」，格子內容為餐點名稱
function syncTodayMenu(dateStr) {
  const today = dateStr || formatDate(new Date());
  const mmdd = today.slice(5).replace('/', ''); // "2026/06/16" → "0616"

  const orderWb = SpreadsheetApp.openById(ORDER_SHEET_ID);
  const allSheets = orderWb.getSheets();

  // 找出所有符合今日日期的分頁（0616午餐、0616晚餐 等）
  const todaySheets = allSheets.filter(s => s.getName().startsWith(mmdd));
  if (!todaySheets.length) return { success: false, message: `找不到 ${mmdd} 的分頁` };

  const menuSheet = getSheet(MENU_SHEET);
  let added = 0;

  // 取得已存在的菜單避免重複
  const existing = new Set(
    menuSheet.getDataRange().getValues().slice(1)
      .map(r => `${r[0]}|${r[1]}|${r[2]}`)
  );

  for (const ws of todaySheets) {
    const data = ws.getDataRange().getValues();
    if (data.length < 2) continue;

    const headers = data[0]; // 第一行是標題
    const mealType = ws.getName().slice(mmdd.length).trim(); // "午餐", "晚餐", "茶湯會" 等

    // 找出店家欄位：標題含「店家」且包含冒號
    for (let col = 0; col < headers.length; col++) {
      const header = String(headers[col]);
      if (!header.includes('店家') || !header.includes('：')) continue;

      // 解析店名：「午餐店家一：一米粒」→「一米粒」
      const storeName = header.split('：').pop().trim();
      if (!storeName) continue;

      // 收集該欄所有不重複的餐點（去除飯量備註）
      const dishes = new Set();
      for (let row = 1; row < data.length; row++) {
        const raw = String(data[row][col] || '').trim();
        if (!raw) continue;
        const cell = raw.replace(/[\s　]*(半飯|去飯|少飯|多飯|全飯|正常飯|不要飯|少量飯|半份飯|去飯量)/g, '').trim();
        if (cell) dishes.add(cell);
      }

      // 寫入 menu sheet
      for (const dish of dishes) {
        const key = `${today}|${storeName}|${dish}`;
        if (existing.has(key)) continue;
        menuSheet.appendRow([today, storeName, dish, mealType]);
        existing.add(key);
        added++;
      }
    }
  }

  return { success: true, date: today, added, sheets: todaySheets.map(s => s.getName()) };
}

// ── Google Form 提交觸發器 ─────────────────────────────────
// 在 Apps Script 設定觸發器: onFormSubmit → 此函式
function onFormSubmit(e) {
  const responses = e.namedValues;
  const date = responses['供應日期'] ? responses['供應日期'][0] : '';
  const store = responses['店家名稱'] ? responses['店家名稱'][0] : '';
  const dishKey = '餐點（每行一道，格式：餐點名稱 價格）';
  const dishes = responses[dishKey] ? responses[dishKey][0] : '';

  if (!store || !dishes) return;

  const sheet = getSheet(MENU_SHEET);
  const lines = dishes.split('\n').filter(l => l.trim());
  for (const dish of lines) {
    sheet.appendRow([date, store.trim(), dish.trim()]);
  }
}

// ── 排程：每天 17:00 自動同步隔天菜單 ────────────────────────
// 在 Apps Script 執行一次 setupDailyTrigger() 即可設定
function setupDailyTrigger() {
  // 刪除舊的同名觸發器避免重複
  ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction() === 'syncScheduled')
    .forEach(t => ScriptApp.deleteTrigger(t));

  ScriptApp.newTrigger('syncScheduled')
    .timeBased()
    .atHour(17)
    .everyDays(1)
    .create();

  return { success: true, message: '已設定每天 17:00 自動同步' };
}

// 列出目前所有觸發器（診斷用）
function listTriggers() {
  const triggers = ScriptApp.getProjectTriggers().map(t => ({
    handler: t.getHandlerFunction(),
    type: String(t.getEventType()),
    source: String(t.getTriggerSource())
  }));
  return { count: triggers.length, triggers };
}

// 自動同步：計算下一個工作日並同步
function syncScheduled() {
  const nextDate = getNextBusinessDay(new Date());
  const result = syncTodayMenu(formatDate(nextDate));
  Logger.log('syncScheduled: ' + JSON.stringify(result));
  return result;
}

// 取得下一個工作日（跳過週末與台灣國定假日）
function getNextBusinessDay(fromDate) {
  const holidays = getTaiwanHolidays();
  const d = new Date(fromDate);
  d.setDate(d.getDate() + 1);
  while (d.getDay() === 0 || d.getDay() === 6 || holidays.has(formatDate(d))) {
    d.setDate(d.getDate() + 1);
  }
  return d;
}

// 台灣 2026 國定假日（可逐年更新）
function getTaiwanHolidays() {
  return new Set([
    '2026/01/01', // 元旦
    '2026/01/28', '2026/01/29', '2026/01/30', '2026/01/31',
    '2026/02/01', '2026/02/02', // 春節
    '2026/02/28', // 和平紀念日
    '2026/04/03', '2026/04/04', // 兒童節/清明
    '2026/05/01', // 勞動節
    '2026/05/31', // 端午
    '2026/09/26', // 中秋
    '2026/10/10', // 國慶
  ]);
}

// ── 工具函式 ───────────────────────────────────────────────
function formatDate(d) {
  const y = d.getFullYear();
  const m = String(d.getMonth() + 1).padStart(2, '0');
  const day = String(d.getDate()).padStart(2, '0');
  return `${y}/${m}/${day}`;
}
