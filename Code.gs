// ============================================================
// SmartExpense - Google Apps Script Backend (Code.gs)
// Deploy as Web App: Execute as Me, Anyone can access
// ============================================================

const SHEET_NAME = "Expenses";
const CATEGORIES = [
  "Milk","Mutton","Fish","Grocery","Gas","Petrol",
  "Internet","Mobile Bill","Electricity Bill","Eating Out",
  "Vegetables","Snacks"
];

function doGet(e) {
  return handleRequest(e);
}

function doPost(e) {
  return handleRequest(e);
}

function handleRequest(e) {
  const callback = e.parameter.callback || null;
  let result;
  try {
    const action = e.parameter.action;
    if (action === "addExpense") {
      result = addExpense(JSON.parse(e.parameter.data));
    } else if (action === "getExpenses") {
      result = getExpenses(e.parameter);
    } else if (action === "deleteExpense") {
      result = deleteExpense(e.parameter.id);
    } else if (action === "updateExpense") {
      result = updateExpense(JSON.parse(e.parameter.data));
    } else if (action === "getCategories") {
      result = { success: true, categories: getCustomCategories() };
    } else if (action === "addCategory") {
      result = addCategory(e.parameter.category);
    } else if (action === "getSummary") {
      result = getSummary(e.parameter);
    } else {
      result = { success: false, error: "Unknown action" };
    }
  } catch (err) {
    result = { success: false, error: err.toString() };
  }

  const json = JSON.stringify(result);
  const output = callback
    ? ContentService.createTextOutput(`${callback}(${json})`).setMimeType(ContentService.MimeType.JAVASCRIPT)
    : ContentService.createTextOutput(json).setMimeType(ContentService.MimeType.JSON);

  return output;
}

// ── Sheet helpers ──────────────────────────────────────────

function getOrCreateSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    setupHeaders(sheet);
  }
  return sheet;
}

function getCustomCategories() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let settingsSheet = ss.getSheetByName("Settings");
  if (!settingsSheet) {
    settingsSheet = ss.insertSheet("Settings");
    settingsSheet.getRange("A1").setValue("CustomCategories");
  }
  const data = settingsSheet.getRange("A2:A").getValues().flat().filter(v => v !== "");
  return [...CATEGORIES, ...data];
}

function addCategory(category) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let settingsSheet = ss.getSheetByName("Settings");
  if (!settingsSheet) {
    settingsSheet = ss.insertSheet("Settings");
    settingsSheet.getRange("A1").setValue("CustomCategories");
  }
  const existing = getCustomCategories();
  if (existing.map(c => c.toLowerCase()).includes(category.toLowerCase())) {
    return { success: false, error: "Category already exists" };
  }
  const lastRow = settingsSheet.getLastRow();
  settingsSheet.getRange(lastRow + 1, 1).setValue(category);
  // Add column to expenses sheet
  const expSheet = getOrCreateSheet();
  const headers = expSheet.getRange(1, 1, 1, expSheet.getLastColumn()).getValues()[0];
  expSheet.getRange(1, expSheet.getLastColumn() + 1).setValue(category);
  return { success: true };
}

function getAllCategories() {
  return getCustomCategories();
}

function setupHeaders(sheet) {
  const allCats = getAllCategories();
  const baseHeaders = ["ID","Date","Category","Amount","Comment"];
  const headers = [...baseHeaders, ...allCats];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  sheet.getRange(1, 1, 1, headers.length).setFontWeight("bold");
  sheet.setFrozenRows(1);
}

function getCategoryColumnIndex(sheet, category) {
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const idx = headers.findIndex(h => h.toString().toLowerCase() === category.toLowerCase());
  return idx >= 0 ? idx + 1 : -1;
}

// ── CRUD ───────────────────────────────────────────────────

function addExpense(data) {
  const sheet = getOrCreateSheet();
  const allCats = getAllCategories();
  // Ensure category columns exist
  const headerRow = sheet.getLastColumn() > 0
    ? sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0]
    : [];

  allCats.forEach(cat => {
    if (!headerRow.map(h => h.toString().toLowerCase()).includes(cat.toLowerCase())) {
      sheet.getRange(1, sheet.getLastColumn() + 1).setValue(cat);
    }
  });

  // If category is new custom one, add column
  const category = data.category;
  const refreshedHeaders = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  if (!refreshedHeaders.map(h => h.toString().toLowerCase()).includes(category.toLowerCase())) {
    sheet.getRange(1, sheet.getLastColumn() + 1).setValue(category);
  }

  // Get next ID
  const lastRow = sheet.getLastRow();
  const newId = lastRow <= 1 ? 1 : parseInt(sheet.getRange(lastRow, 1).getValue() || 0) + 1;

  // Build row
  const totalCols = sheet.getLastColumn();
  const row = new Array(totalCols).fill("");
  row[0] = newId;           // ID
  row[1] = data.date;       // Date
  row[2] = data.category;   // Category
  row[3] = data.amount;     // Amount
  row[4] = data.comment || ""; // Comment

  // Fill category column
  const catIdx = sheet.getRange(1, 1, 1, totalCols).getValues()[0]
    .findIndex(h => h.toString().toLowerCase() === category.toLowerCase());
  if (catIdx >= 0) row[catIdx] = data.amount;

  sheet.appendRow(row);
  return { success: true, id: newId };
}

function getExpenses(params) {
  const sheet = getOrCreateSheet();
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return { success: true, data: [], headers: [] };

  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const rows = sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).getValues();

  let data = rows.map(row => {
    const obj = {};
    headers.forEach((h, i) => { obj[h] = row[i]; });
    return obj;
  }).filter(r => r["ID"] !== "");

  // Filters
  if (params.category && params.category !== "all") {
    data = data.filter(r => r["Category"].toString().toLowerCase() === params.category.toLowerCase());
  }
  if (params.dateFrom) {
    data = data.filter(r => r["Date"] >= params.dateFrom);
  }
  if (params.dateTo) {
    data = data.filter(r => r["Date"] <= params.dateTo);
  }
  if (params.search) {
    const q = params.search.toLowerCase();
    data = data.filter(r =>
      r["Category"].toString().toLowerCase().includes(q) ||
      r["Comment"].toString().toLowerCase().includes(q)
    );
  }

  return { success: true, data, headers };
}

function deleteExpense(id) {
  const sheet = getOrCreateSheet();
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return { success: false, error: "No data" };
  const ids = sheet.getRange(2, 1, lastRow - 1, 1).getValues().flat();
  const rowIdx = ids.findIndex(r => r.toString() === id.toString());
  if (rowIdx < 0) return { success: false, error: "Not found" };
  sheet.deleteRow(rowIdx + 2);
  return { success: true };
}

function updateExpense(data) {
  const sheet = getOrCreateSheet();
  const lastRow = sheet.getLastRow();
  const ids = sheet.getRange(2, 1, lastRow - 1, 1).getValues().flat();
  const rowIdx = ids.findIndex(r => r.toString() === data.id.toString());
  if (rowIdx < 0) return { success: false, error: "Not found" };
  const sheetRow = rowIdx + 2;
  const totalCols = sheet.getLastColumn();
  const headers = sheet.getRange(1, 1, 1, totalCols).getValues()[0];

  const row = new Array(totalCols).fill("");
  row[0] = data.id;
  row[1] = data.date;
  row[2] = data.category;
  row[3] = data.amount;
  row[4] = data.comment || "";
  const catIdx = headers.findIndex(h => h.toString().toLowerCase() === data.category.toLowerCase());
  if (catIdx >= 0) row[catIdx] = data.amount;

  sheet.getRange(sheetRow, 1, 1, totalCols).setValues([row]);
  return { success: true };
}

function getSummary(params) {
  const result = getExpenses(params);
  if (!result.success) return result;

  const data = result.data;
  const allCats = getAllCategories();

  // Category totals
  const categoryTotals = {};
  allCats.forEach(c => { categoryTotals[c] = 0; });
  data.forEach(r => {
    const cat = r["Category"];
    categoryTotals[cat] = (categoryTotals[cat] || 0) + (parseFloat(r["Amount"]) || 0);
  });

  // Daily totals
  const dailyTotals = {};
  data.forEach(r => {
    const d = r["Date"];
    dailyTotals[d] = (dailyTotals[d] || 0) + (parseFloat(r["Amount"]) || 0);
  });

  // Monthly totals
  const monthlyTotals = {};
  data.forEach(r => {
    const month = r["Date"].toString().substring(0, 7);
    monthlyTotals[month] = (monthlyTotals[month] || 0) + (parseFloat(r["Amount"]) || 0);
  });

  const total = data.reduce((s, r) => s + (parseFloat(r["Amount"]) || 0), 0);

  return {
    success: true,
    total,
    categoryTotals,
    dailyTotals,
    monthlyTotals,
    count: data.length
  };
}
