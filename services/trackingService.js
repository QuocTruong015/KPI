const { excelDateToJSDate } = require("../utils/excelUtils");

function processTracking(data, month, year) {
  // Bước 1: Chỉ giữ lại 5 cột cần thiết
  const selectedColumns = data.map((row) => ({
    Date: row["Date"],
    Rev: parseFloat(row["Rev_1"]) || 0,
    Cost: parseFloat(row["Cost_1"]) || 0,
    Profit: parseFloat(row["Profit"]) || 0,
    Type: (row["Type_1"] || "").trim(), // trim() để bỏ khoảng trắng thừa
  }));

  // Bước 2: Lọc theo tháng và năm
  const filtered = selectedColumns.filter((row) => {
    const date = excelDateToJSDate(row.Date);
    if (!date) return false;
    return date.getMonth() + 1 === month && date.getFullYear() === year;
  });

  console.log("✅ Dữ liệu sau khi lọc:", filtered);

  // Bước 3: Tính tổng profit của các dòng có Type = "Tracking Ảo"
  const totalTracking = filtered
    .filter((row) => row.Type === "Tracking Ảo")
    .reduce((sum, row) => sum + row.Profit, 0);

  return [
    {
      Month: month,
      Year: year,
      TotalTracking: totalTracking,
    },
  ];
}

module.exports = { processTracking };
