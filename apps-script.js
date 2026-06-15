// ============================================================
// 便當評分系統 - Google Apps Script
// 部署為 Web App: 執行身分「我」, 存取權「所有人」
// ============================================================

const SHEET_ID = '18DCXoe8iXKaEA21zyzqU0kC5rrUsCcGQE9Evepti8ok';
const ORDER_SHEET_ID = '1XffLpd0nUkgp5TElqAwiSVLchrmDxDarJa-jJAn1lOg'; // 餐點確認 Sheet（固定）
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

// ── 取得今日供應店家 ────────────────────────────────────────
function getTodayStores(dateStr) {
  const sheet = getSheet(MENU_SHEET);
  const data = sheet.getDataRange().getValues();
  const today = dateStr || formatDate(new Date());

  const stores = {};
  for (let i = 1; i < data.length; i++) {
    const [date, store, dish] = data[i];
    if (!store || !dish) continue;
    const d = date ? formatDate(new Date(date)) : '';
    if (d === today) {
      if (!stores[store]) stores[store] = [];
      stores[store].push(dish);
    }
  }
  return { date: today, stores };
}

// ── 取得菜單（可篩選日期）──────────────────────────────────
function getMenu(dateFilter) {
  const sheet = getSheet(MENU_SHEET);
  const data = sheet.getDataRange().getValues();
  const result = {};

  for (let i = 1; i < data.length; i++) {
    const [date, store, dish] = data[i];
    if (!store || !dish) continue;
    if (dateFilter) {
      const d = date ? formatDate(new Date(date)) : '';
      if (d !== dateFilter) continue;
    }
    if (!result[store]) result[store] = [];
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
  const today = formatDate(new Date());
  sheet.appendRow([
    today,
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

    // 找出店家欄位：標題含「店家」且包含冒號
    for (let col = 0; col < headers.length; col++) {
      const header = String(headers[col]);
      if (!header.includes('店家') || !header.includes('：')) continue;

      // 解析店名：「午餐店家一：一米粒」→「一米粒」
      const storeName = header.split('：').pop().trim();
      if (!storeName) continue;

      // 收集該欄所有不重複的餐點
      const dishes = new Set();
      for (let row = 1; row < data.length; row++) {
        const cell = String(data[row][col] || '').trim();
        if (!cell) continue;
        // 去掉飯量備註（如「牛逼菲力 半飯」→ 保留完整，或可自訂）
        dishes.add(cell);
      }

      // 寫入 menu sheet
      for (const dish of dishes) {
        const key = `${today}|${storeName}|${dish}`;
        if (existing.has(key)) continue;
        menuSheet.appendRow([today, storeName, dish]);
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
  const dishes = responses['餐點（每行一道，格式：餐點名稱 價格）']
    ? responses['餐點（每行一道，格式：餐點名稱 價格）'][0] : '';

  if (!store || !dishes) return;

  const sheet = getSheet(MENU_SHEET);
  const lines = dishes.split('\n').filter(l => l.trim());
  for (const dish of lines) {
    sheet.appendRow([date, store.trim(), dish.trim()]);
  }
}

// ── 工具函式 ───────────────────────────────────────────────
function formatDate(d) {
  const y = d.getFullYear();
  const m = String(d.getMonth() + 1).padStart(2, '0');
  const day = String(d.getDate()).padStart(2, '0');
  return `${y}/${m}/${day}`;
}
