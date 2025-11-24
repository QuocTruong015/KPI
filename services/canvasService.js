const { excelDateToJSDate } = require("../utils/excelUtils");

function processCanvasRev(data, month, year) {

  const filtered = data.filter((row) => {
    const date = excelDateToJSDate(row[" Month"]);
    if (!date) return false;
    return date.getMonth() + 1 === month && date.getFullYear() === year;
  });

  let totalRev = 0;
  filtered.forEach((row) => {
    // Giả sử cột Z tương ứng với __EMPTY_2
    const rev = parseFloat(row.__EMPTY_17) || 0;
    totalRev += rev;
  });

  return [
    {
      Month: month,
      Year: year,
      TotalRev: Number(totalRev.toFixed(2)),
    },
  ];
}

function processCanvasCost(data, month, year) {
  const filtered = data.filter((row) => {
    const date = excelDateToJSDate(row[" Month"]);
    if (!date) return false;
    return date.getMonth() + 1 === month && date.getFullYear() === year;
  });

  let totalCost = 0;
  filtered.forEach((row) => {
    // Giả sử cột Z tương ứng với __EMPTY_2
    const cost = parseFloat(row.__EMPTY_5) || 0;
    totalCost += cost;
  });
  return [
    {
      Month: month,
      Year: year,
      TotalCost: Number(totalCost.toFixed(2)),
    },
  ];
}

module.exports = { processCanvasRev, processCanvasCost };