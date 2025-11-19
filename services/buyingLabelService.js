const { excelDateToJSDate } = require("../utils/excelUtils");

//Buying label for service staff 1
function processBuyingLabel(data, month, year) {
  const filtered = data.filter((row) => {
    const date = excelDateToJSDate(row.Date);
    if (!date) return false;
    return date.getMonth() + 1 === month && date.getFullYear() === year;
  });

  const result = {};
  filtered.forEach((row) => {
    const seller = row.Seller?.trim() || "Unknown";
    const rev = parseFloat(row.REV) || 0;
    const cost = parseFloat(row.Cost) || 0;

    let profit = rev - cost;

    if (!result[seller]) result[seller] = { Seller: seller, TotalRev: 0, TotalCost: 0, TotalProfit: 0 };
    result[seller].TotalRev += rev;
    result[seller].TotalCost += cost;
    result[seller].TotalProfit += profit;
  });

  return Object.values(result).map((s) => ({
    Seller: s.Seller,
    TotalRev: +s.TotalRev.toFixed(2),
    TotalCost: +s.TotalCost.toFixed(2),
    TotalProfit: +s.TotalProfit.toFixed(2),
  }));
}

module.exports = { processBuyingLabel };
