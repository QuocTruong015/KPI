const { excelDateToJSDate } = require("../utils/excelUtils");

function processScanLabel(data, month, year) {
  const filtered = data.filter((row) => {
    const date = excelDateToJSDate(row.Date);
    if (!date) return false;
    return date.getMonth() + 1 === month && date.getFullYear() === year;
  });

  const result = {};
  filtered.forEach((row) => {
    const seller = row.Seller?.trim() || "Unknown";
    const user = row.User?.trim() || "";
    const createdAt = row["Created At"] ? excelDateToJSDate(row["Created At"]) : null;
    const cost = parseFloat(row.Cost) || 0;

    let profit = 0;
    if (rev === 1.5) {
      profit = cost * 0.3;
    } else {
      profit = (rev - cost) + (cost * 0.3);
    }

    if (!result[seller]) result[seller] = { Seller: seller, TotalRev: 0, TotalProfit: 0 };
    result[seller].TotalRev += rev;
    result[seller].TotalProfit += profit;
  });

  return Object.values(result).map((s) => ({
    Seller: s.Seller,
    TotalRev: +s.TotalRev.toFixed(2),
    TotalProfit: +s.TotalProfit.toFixed(2),
  }));
}
module.exports = { processScanLabel };
