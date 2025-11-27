const { excelDateToJSDate } = require("../utils/excelUtils");

function processScanLabel(data, month, year) {
  const filtered = data.filter((row) => {
    const date = excelDateToJSDate(row['Date']);
    if (!date) 
      return false;
    return date.getMonth() + 1 === month && date.getFullYear() === year;
  });

  let totalRevenue = 0;
  const uspsCost = 26 * 40;
  let totalCostGA = 0;
  let totalCostTX = 0;
  filtered.forEach((row) => {
    const revenue = parseFloat(row['Total Revenue']) || 0;
    const costGA = parseFloat(row['Cost GA']) || 0;
    const costTX = parseFloat(row['Cost TX']) || 0;
    totalRevenue += revenue;
    totalCostGA += costGA;
    totalCostTX += costTX;
  });

  const profit = totalRevenue - (totalCostGA + totalCostTX + uspsCost);

  return { scanLabelTotalProfit: profit };
}
module.exports = { processScanLabel };
