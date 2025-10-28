const { excelDateToJSDate, parseMonthYear } = require("../utils/excelUtils");
const XLSX = require('xlsx');

// Hàm validate row (sử dụng chung cho processEtsyStatement)
function validateRow(row) {
  const requiredFields = ["Date", "Type", "Order ID (sale, refund)"];
  const missingFields = requiredFields.filter((field) => !row[field] || String(row[field]).trim() === "");
  return missingFields.length === 0 ? null : `Thiếu cột: ${missingFields.join(", ")}`;
}

function processTargetKpi(data) {
  if (!Array.isArray(data) || data.length === 0) {
    throw new Error("Dữ liệu Excel rỗng hoặc không hợp lệ");
  }
  const result = data.map((row, index) => {
    const rawMonth = parseMonthYear(row["Month"]);
    return {
        Month:      rawMonth,
        PIC:        row["PIC"] ? String(row["PIC"]).trim() : "Unknown",
        Position:   row["Position"] ? String(row["Position"]).trim() : "Unknown",
        Target:     row["Target (100%)"] != null ? parseFloat(row["Target (100%)"]) || 0 : 0,
    };
  });

  return result;
}

module.exports = { processTargetKpi };