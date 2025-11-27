const { excelDateToJSDate } = require("../utils/excelUtils");

// Buying label for service staff 1
function processBuyingLabel(data, month, year) {
  const filtered = data.filter((row) => {
    const date = excelDateToJSDate(row.Date);
    if (!date) return false;
    return date.getMonth() + 1 === month && date.getFullYear() === year;
  });

  const result = {};
  let grandTotalProfit = 0; // ⭐ Tổng profit của tất cả seller

  filtered.forEach((row) => {
    const seller = row.Seller?.trim() || "Unknown";
    const rev = parseFloat(row.REV) || 0;
    const cost = parseFloat(row.Cost) || 0;

    const profit = rev - cost;

    // ⭐ cộng dồn vào tổng profit toàn bộ
    grandTotalProfit += profit;

    if (!result[seller]) {
      result[seller] = {
        Seller: seller,
        TotalRev: 0,
        TotalCost: 0,
        TotalProfit: 0
      };
    }

    result[seller].TotalRev += rev;
    result[seller].TotalCost += cost;
    result[seller].TotalProfit += profit;
  });

  const sellerList = Object.values(result).map((s) => ({
    Seller: s.Seller,
    TotalRev: +s.TotalRev.toFixed(2),
    TotalCost: +s.TotalCost.toFixed(2),
    TotalProfit: +s.TotalProfit.toFixed(2),
  }));

  return {
    sellers: sellerList,
    buyingTotalProfit: +grandTotalProfit.toFixed(2), // ⭐ trả về tổng Buying Label Profit
  };
}

module.exports = { processBuyingLabel };
