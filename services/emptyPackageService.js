const { excelDateToJSDate } = require("../utils/excelUtils");

function processEmptyPackage(data, month, year) {
  const filtered = data.filter((row) => {
    const date = excelDateToJSDate(row.Date);
    if (!date) return false;
    return date.getMonth() + 1 === month && date.getFullYear() === year;
  });

  const result = {};
  let grandTotalProfit = 0; // ⭐ tổng empty profit

  filtered.forEach((row) => {
    const seller = row.Seller?.trim() || "Unknown";
    const rev = parseFloat(row.Rev) || 0;
    const cost = parseFloat(row.Cost) || 0;

    let profit = 0;
    if (rev === 1.5) {
      profit = cost * 0.3;
    } else {
      profit = (rev - cost) + (cost * 0.3);
    }

    // ⭐ cộng dồn tổng profit tất cả seller
    grandTotalProfit += profit;

    if (!result[seller]) {
      result[seller] = { Seller: seller, TotalRev: 0, TotalProfit: 0 };
    }

    result[seller].TotalRev += rev;
    result[seller].TotalProfit += profit;
  });

  const sellerList = Object.values(result).map((s) => ({
    Seller: s.Seller,
    TotalRev: +s.TotalRev.toFixed(2),
    TotalProfit: +s.TotalProfit.toFixed(2),
  }));

  return {
    sellers: sellerList,
    emptyTotalProfit: +grandTotalProfit.toFixed(2), // ⭐ trả ra luôn
  };
}

module.exports = { processEmptyPackage };
