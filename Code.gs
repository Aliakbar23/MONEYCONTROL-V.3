// ============================================================
//  DOMPETKU — Google Apps Script Backend
//  Salin kode ini ke Apps Script project kamu
// ============================================================

const SHEET_NAME = "Pengeluaran";
const CATEGORY_SHEET = "Kategori";

function doGet(e) {
  return handleRequest(e);
}

function doPost(e) {
  return handleRequest(e);
}

function handleRequest(e) {
  const params = e.parameter;
  const action = params.action;

  let result;

  try {
    switch (action) {
      case "getAll":
        result = getAllExpenses();
        break;
      case "add":
        result = addExpense(params);
        break;
      case "delete":
        result = deleteExpense(params.id);
        break;
      case "getCategories":
        result = getCategories();
        break;
      case "getSummary":
        result = getSummary(params.month, params.year);
        break;
      default:
        result = { error: "Action tidak dikenal" };
    }
  } catch (err) {
    result = { error: err.message };
  }

  return ContentService
    .createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}

// ── Setup sheet pertama kali ─────────────────────────────────
function setupSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // Sheet Pengeluaran
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    sheet.appendRow(["ID", "Tanggal", "Keterangan", "Kategori", "Jumlah", "Catatan"]);
    sheet.getRange(1, 1, 1, 6).setFontWeight("bold");
  }

  // Sheet Kategori
  let catSheet = ss.getSheetByName(CATEGORY_SHEET);
  if (!catSheet) {
    catSheet = ss.insertSheet(CATEGORY_SHEET);
    catSheet.appendRow(["Nama", "Ikon", "Warna"]);
    const defaultCats = [
      ["Makanan & Minuman", "🍜", "#FF6B6B"],
      ["Transportasi", "🚗", "#4ECDC4"],
      ["Belanja", "🛍️", "#45B7D1"],
      ["Hiburan", "🎮", "#96CEB4"],
      ["Kesehatan", "💊", "#FFEAA7"],
      ["Tagihan", "📱", "#DDA0DD"],
      ["Pendidikan", "📚", "#98D8C8"],
      ["Lainnya", "💰", "#B0B0B0"]
    ];
    defaultCats.forEach(row => catSheet.appendRow(row));
  }
}

// ── Ambil semua pengeluaran ──────────────────────────────────
function getAllExpenses() {
  const sheet = getOrCreateSheet();
  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return { data: [] };

  const headers = data[0];
  const rows = data.slice(1).map(row => {
    const obj = {};
    headers.forEach((h, i) => obj[h] = row[i]);
    return obj;
  }).filter(r => r.ID !== "");

  return { data: rows.reverse() }; // terbaru dulu
}

// ── Tambah pengeluaran ───────────────────────────────────────
function addExpense(params) {
  const sheet = getOrCreateSheet();
  const id = new Date().getTime().toString();
  const tanggal = params.tanggal || Utilities.formatDate(new Date(), "Asia/Jakarta", "yyyy-MM-dd");
  const jumlah = parseFloat(params.jumlah) || 0;

  sheet.appendRow([
    id,
    tanggal,
    params.keterangan || "",
    params.kategori || "Lainnya",
    jumlah,
    params.catatan || ""
  ]);

  return { success: true, id: id };
}

// ── Hapus pengeluaran ────────────────────────────────────────
function deleteExpense(id) {
  const sheet = getOrCreateSheet();
  const data = sheet.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    if (data[i][0].toString() === id.toString()) {
      sheet.deleteRow(i + 1);
      return { success: true };
    }
  }
  return { error: "Data tidak ditemukan" };
}

// ── Ambil kategori ───────────────────────────────────────────
function getCategories() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let catSheet = ss.getSheetByName(CATEGORY_SHEET);

  if (!catSheet) {
    setupSheets();
    catSheet = ss.getSheetByName(CATEGORY_SHEET);
  }

  const data = catSheet.getDataRange().getValues();
  if (data.length <= 1) return { data: [] };

  const cats = data.slice(1).map(row => ({
    nama: row[0],
    ikon: row[1],
    warna: row[2]
  }));

  return { data: cats };
}

// ── Ringkasan bulanan ────────────────────────────────────────
function getSummary(month, year) {
  const sheet = getOrCreateSheet();
  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return { total: 0, byCategory: {} };

  const targetMonth = parseInt(month);
  const targetYear = parseInt(year);

  let total = 0;
  const byCategory = {};

  data.slice(1).forEach(row => {
    const tanggal = new Date(row[1]);
    if (tanggal.getMonth() + 1 === targetMonth && tanggal.getFullYear() === targetYear) {
      const jumlah = parseFloat(row[4]) || 0;
      const kategori = row[3];
      total += jumlah;
      byCategory[kategori] = (byCategory[kategori] || 0) + jumlah;
    }
  });

  return { total, byCategory, month: targetMonth, year: targetYear };
}

// ── Helper ───────────────────────────────────────────────────
function getOrCreateSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    setupSheets();
    sheet = ss.getSheetByName(SHEET_NAME);
  }
  return sheet;
}
