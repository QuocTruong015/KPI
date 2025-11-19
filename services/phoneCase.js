const { da } = require("date-fns/locale");
const { excelDateToJSDate } = require("../utils/excelUtils");

function processPhoneCaseCost(data, month, year) {
  const filtered = data.filter((row) => {
    const date = excelDateToJSDate(row.created_at);
    if (!date) return false;
    return date.getMonth() + 1 === month && date.getFullYear() === year;
  });

  let totalCost = 0;

  filtered.forEach((row) => {
    const cost = parseFloat(row.grand_total) || 0;
    totalCost += cost;
  });

  // Trả về kết quả tổng chi phí tháng đó
  return [
    {
      Month: month,
      Year: year,
      TotalCost: Number(totalCost.toFixed(2)),
    },
  ];
}
function processPhoneCaseRev(data, month, year) {
  const filtered = data.filter((row) => {
    const date = excelDateToJSDate(row[" Month"]);
    if (!date) return false;
    return date.getMonth() + 1 === month && date.getFullYear() === year;
  });

  let totalRev = 0;
  filtered.forEach((row) => {
    // Giả sử cột Z tương ứng với __EMPTY_2
    const rev = parseFloat(row.__EMPTY_15) || 0;
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

function processPhoneCaseProfit(revData, costData, month, year) {
  // Gọi 2 hàm xử lý dữ liệu doanh thu và chi phí
  const revGroup = processPhoneCaseRev(revData, month, year);
  const costGroup = processPhoneCaseCost(costData, month, year);

  console.log("📊 Dữ liệu nhóm doanh thu:", revGroup);

  // Kiểm tra dữ liệu hợp lệ
  if (!revGroup || !revGroup.length || !costGroup || !costGroup.length) {
    throw new Error("Dữ liệu doanh thu hoặc chi phí rỗng!");
  }
  const totalRev = revGroup[0].TotalRev || 0;
  const totalCost = costGroup[0].TotalCost || 0;

  // Tính profit
  const totalProfit = totalRev - totalCost;

  return [
    {
      Month: month,
      Year: year,
      TotalRev: Number(totalRev.toFixed(2)),
      TotalCost: Number(totalCost.toFixed(2)),
      TotalProfit: Number(totalProfit.toFixed(2)),
    },
  ];
}

module.exports = { processPhoneCaseCost, processPhoneCaseRev, processPhoneCaseProfit };
