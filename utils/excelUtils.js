const XLSX = require("xlsx");

// Chuyển Excel serial date sang JS Date
function excelDateToJSDate(value) {
  if (!value) return null;
  if (typeof value === "number") return new Date((value - 25569) * 86400 * 1000);
  if (typeof value === "string" && !isNaN(Date.parse(value))) return new Date(value);
  if (value instanceof Date) return value;
  return null;
}

// Đọc sheet Excel theo tên
function readExcelSheet(filePath, preferredSheetName, sheetIndex = 0) {
  const workbook = XLSX.readFile(filePath);
  const sheetName = workbook.SheetNames.find((s) => s === preferredSheetName) || workbook.SheetNames[sheetIndex];
  const sheet = workbook.Sheets[sheetName];
  const data = XLSX.utils.sheet_to_json(sheet, { defval: "", raw: true });
  return { data, sheetName };
}

function parseMonthYear(val) {
  if (!val) return null;
  const s = String(val).trim();

  // Hỗ trợ: "Aug-2025", "Aug 2025", "August-2025", "August 2025"
  const match = s.match(/([A-Za-z]+)[\s-]*(\d{4})/i);
  if (!match) return null;

  const monthName = match[1].toLowerCase();
  const year = parseInt(match[2], 10);

  const months = {
    jan: 1, january: 1,
    feb: 2, february: 2,
    mar: 3, march: 3,
    apr: 4, april: 4,
    may: 5,
    jun: 6, june: 6,
    jul: 7, july: 7,
    aug: 8, august: 8,
    sep: 9, sept: 9, september: 9,
    oct: 10, october: 10,
    nov: 11, november: 11,
    dec: 12, december: 12
  };

  // Tìm theo 3 chữ cái đầu hoặc tên đầy đủ
  const key = monthName.substring(0, 3);
  const month = months[key] || months[monthName];

  if (!month || year < 2000 || year > 2100) return null;

  return { month, year };
}

module.exports = { excelDateToJSDate, readExcelSheet, parseMonthYear };
