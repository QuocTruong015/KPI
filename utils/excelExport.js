// excelExport.js
const XLSX = require('xlsx');
const fs = require('fs');
const path = require('path');

function exportProfitToExcel(aggregated, outputDir = "./exports") {
  const { designerProfit, rdProfit, platformSummary, totalProfit, month, year } = aggregated;

  // Tạo thư mục
  if (!fs.existsSync(outputDir)) fs.mkdirSync(outputDir, { recursive: true });

  const filename = `Profit_Summary_${year}_${String(month).padStart(2, '0')}.xlsx`;
  const filePath = path.join(outputDir, filename);

  const wb = XLSX.utils.book_new();

  // Sheet 1: Designer Profit
  const designerData = Object.entries(designerProfit).map(([id, profit]) => ({
    DesignerID: id,
    Profit: profit
  }));
  const ws1 = XLSX.utils.json_to_sheet(designerData);
  XLSX.utils.book_append_sheet(wb, ws1, "Designer_Profit");

  // Sheet 2: R&D Profit
  const rdData = Object.entries(rdProfit).map(([id, profit]) => ({
    RAndDID: id,
    Profit: profit
  }));
  const ws2 = XLSX.utils.json_to_sheet(rdData);
  XLSX.utils.book_append_sheet(wb, ws2, "RD_Profit");

  // Sheet 3: Platform Summary
  const summaryData = [
    { Platform: "Amazon", Profit: platformSummary.Amazon },
    { Platform: "Merch", Profit: platformSummary.Merch },
    { Platform: "Web", Profit: platformSummary.Web },
    { Platform: "Etsy", Profit: platformSummary.Etsy },
    { Platform: "TOTAL", Profit: totalProfit }
  ];
  const ws3 = XLSX.utils.json_to_sheet(summaryData);
  XLSX.utils.book_append_sheet(wb, ws3, "Platform_Summary");

  // Ghi file
  XLSX.writeFile(wb, filePath);
  console.log(`Exported: ${filePath}`);
  return filePath;
}

module.exports = { exportProfitToExcel };