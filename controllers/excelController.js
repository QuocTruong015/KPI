const fs = require("fs");
const path = require("path");
const XLSX = require('xlsx');
const { readExcelSheet, excelDateToJSDate } = require("../utils/excelUtils");
const { processEmptyPackage } = require("../services/emptyPackageService");
const { processBuyingLabel } = require("../services/buyingLabelService");
const { processScanLabel } = require("../services/scanLabelService");
const { processEtsyFFCost, processEtsyOrder, calculateKPI, processEtsyStatement } = require("../services/etsyService");
const { processAmzTransaction, processAmzFFCost, processAmzOrder, calculateAmzKPI } = require("../services/amzService");
const { processWebIdAndRev, processWebFFCostAtWebCost, processWebFFCostAtFFCost, calculateTotalCost, assignProfitToDesignerAndRD } = require("../services/webService");
const { processMerchOrder, processMerchSku, assignProfitToDesignerAndRDMerch } = require("../services/merchService");
const { exportProfitToExcel } = require('../utils/excelExport');
const { getEtsyProfit, getAmazonProfit, getWebProfit, getMerchProfit, aggregateProfit } = require("../services/profitAggregatorService");

async function uploadFileCommon(req, res, sheetName, sheetIndex, processFunc, totalKey = "totalSellers") {
  try {
    const month = parseInt(req.query.month);
    const year = parseInt(req.query.year);

    if (!req.file)
      return res.status(400).json({ error: "Vui lòng upload 1 file Excel!" });
    if (!month || !year)
      return res.status(400).json({ error: "Vui lòng nhập ?month=...&year=..." });

    const filePath = path.join(__dirname, "..", req.file.path);
    const { data, sheetName: actualSheetName } = readExcelSheet(filePath, sheetName, sheetIndex);

    const finalData = processFunc(data, month, year);

    fs.unlinkSync(filePath);

    res.json({
      sheetName: actualSheetName,
      month,
      year,
      [totalKey]: finalData.length,
      data: finalData,
    });
  } catch (error) {
    console.error(error);
    res.status(500).json({ error: "Đọc file Excel thất bại!" });
  }
}

async function uploadEtsyProfit(req, res) {
  try {
    const month = parseInt(req.query.month);
    const year = parseInt(req.query.year);

    if (!month || !year) {
      return res.status(400).json({ error: "Vui lòng nhập ?month=...&year=..." });
    }

    if (!req.file) {
      return res.status(400).json({ error: "Vui lòng upload 1 file Excel chứa 3 sheet!" });
    }

    const filePath = path.join(__dirname, "..", req.file.path);

    // Đọc dữ liệu từ 3 sheet
    const statementData = readExcelSheet(filePath, "", 11).data;
    const ffCostData = readExcelSheet(filePath, "", 12).data;
    const orderData = readExcelSheet(filePath, "", 10).data;

    // Kiểm tra dữ liệu có rỗng không
    if (!statementData.length || !ffCostData.length || !orderData.length) {
      fs.unlinkSync(filePath);
      return res.status(400).json({ error: "Một hoặc nhiều sheet trong file Excel rỗng!" });
    }

    // Gọi hàm calculateKPI
    const finalData = await calculateKPI(statementData, ffCostData, orderData, month, year);

    // Xóa file tạm
    fs.unlinkSync(filePath);

    // Trả về kết quả chuẩn hóa
    res.json(finalData); // Trả về trực tiếp { designerProfit, rdProfit }
  } catch (error) {
    console.error("Lỗi xử lý file:", error.message, error.stack);
    res.status(500).json({ error: "Xử lý file Excel thất bại! Chi tiết: " + error.message });
  }
}

async function uploadEmptyPackage(req, res) {
  return uploadFileCommon(req, res, "Empty Package", 8, processEmptyPackage);
}

async function uploadBuyingLabel(req, res) {
  return uploadFileCommon(req, res, "Buying Label", 9, processBuyingLabel);
}

async function uploadScanLabel(req, res) {
  return uploadFileCommon(req, res, "Scan Label", 7, processScanLabel);
}

async function uploadEtsyStatement(req, res) {
  return uploadFileCommon(req, res, "", 11, processEtsyStatement);
}

async function uploadEtsyFFCost(req, res) {
  return uploadFileCommon(req, res, "", 12, processEtsyFFCost);
}

async function uploadEtsyOrder(req, res) {
  return uploadFileCommon(req, res, "", 10, processEtsyOrder);
}

//AMZ
async function uploadAmzTransaction(req, res) {
  return uploadFileCommon(req, res, "", 15, processAmzTransaction);
}

async function uploadAmzFFCost(req, res) {
  return uploadFileCommon(req, res, "", 16, processAmzFFCost);
}

async function uploadAmzOrder(req, res) {
  return uploadFileCommon(req, res, "", 14, processAmzOrder);
}

async function uploadAmzProfit(req, res) {
  try {
    console.log("Received request with query:", JSON.stringify(req.query, null, 2));

    const month = parseInt(req.query.month);
    const year = parseInt(req.query.year);

    console.log(`Parsed query params - month: ${month}, year: ${year}`);

    if (isNaN(month) || isNaN(year)) {
      console.error("Invalid or missing month/year query parameters", { month, year });
      return res.status(400).json({ error: "Vui lòng nhập ?month=...&year=... với giá trị hợp lệ" });
    }

    if (!req.file) {
      console.error("No Excel file uploaded");
      return res.status(400).json({ error: "Vui lòng upload 1 file Excel chứa 3 sheet!" });
    }

    const filePath = path.join(__dirname, "..", req.file.path);

    // Đọc dữ liệu từ 3 sheet
    const statementData = readExcelSheet(filePath, "", 15).data;
    const ffCostData = readExcelSheet(filePath, "", 16).data;
    const orderData = readExcelSheet(filePath, "", 14).data;

    console.log(`statementData sample (${statementData.length} rows):`, JSON.stringify(statementData.slice(0, 2), null, 2));
    console.log(`ffCostData sample (${ffCostData.length} rows):`, JSON.stringify(ffCostData.slice(0, 2), null, 2));
    console.log(`orderData sample (${orderData.length} rows):`, JSON.stringify(orderData.slice(0, 2), null, 2));

    // Kiểm tra dữ liệu có rỗng không
    if (!statementData.length || !ffCostData.length || !orderData.length) {
      fs.unlinkSync(filePath);
      console.error("One or more Excel sheets are empty");
      return res.status(400).json({ error: "Một hoặc nhiều sheet trong file Excel rỗng!" });
    }

    // Gọi hàm calculateAmzKPI với month và year
    const finalData = await calculateAmzKPI(statementData, ffCostData, orderData, month, year);

    // Xóa file tạm
    fs.unlinkSync(filePath);

    // Trả về kết quả
    res.json(finalData);
  } catch (error) {
    console.error("Lỗi xử lý file:", error.message, error.stack);
    res.status(500).json({ error: "Xử lý file Excel thất bại! Chi tiết: " + error.message });
  }
}

//Web
  async function uploadWebOrder(req, res) {
    return uploadFileCommon(req, res, "", 18, processWebIdAndRev);
  }

  async function uploadWebCost1(req, res) {
    return uploadFileCommon(req, res, "", 19, processWebFFCostAtWebCost);
  }

  async function uploadWebCost2(req, res) {
    return uploadFileCommon(req, res, "", 20, processWebFFCostAtFFCost);
  }

async function uploadWebTotalCost(req, res) {
  try {
    const month = parseInt(req.query.month);
    const year = parseInt(req.query.year);

    if (!month || !year) {
      return res.status(400).json({ error: "Vui lòng nhập ?month=...&year=..." });
    }

    if (!req.file) {
      return res.status(400).json({ error: "Vui lòng upload 1 file Excel chứa 2 sheet!" });
    }

    const filePath = path.join(__dirname, "..", req.file.path);

    // Đọc dữ liệu từ hai sheet
    const webCostData = readExcelSheet(filePath, "", 19).data;
    const ffCostData = readExcelSheet(filePath, "", 20).data;

    // Log dữ liệu từ hai sheet
    console.log("=== Dữ liệu từ sheet Web Cost (Sheet 19) ===");
    console.log("Số dòng:", webCostData.length);
    webCostData.forEach((row, index) => {
      console.log(`Dòng ${index + 1}:`, JSON.stringify(row, null, 2));
    });

    console.log("\n=== Dữ liệu từ sheet FF Cost (Sheet 20) ===");
    console.log("Số dòng:", ffCostData.length);
    ffCostData.forEach((row, index) => {
      console.log(`Dòng ${index + 1}:`, JSON.stringify(row, null, 2));
    });

    // Tính tổng chi phí
    const finalData = calculateTotalCost(webCostData, ffCostData);

    // Xóa file tạm
    fs.unlinkSync(filePath);

    // Trả về kết quả
    res.json({
      sheetName: "Total Cost",
      month,
      year,
      totalRecords: finalData.length,
      data: finalData,
    });
  } catch (error) {
    console.error("Lỗi xử lý file:", error.message, error.stack);
    res.status(500).json({ error: "Xử lý file Excel thất bại! Chi tiết: " + error.message });
  }
}

async function uploadWebProfit(req, res) {
  try {
    const month = parseInt(req.query.month);
    const year = parseInt(req.query.year);

    if (!month || !year) {
      return res.status(400).json({ error: "Vui lòng nhập ?month=...&year=..." });
    }

    if (!req.file) {
      return res.status(400).json({ error: "Vui lòng upload 1 file Excel chứa 3 sheet!" });
    }

    const filePath = path.join(__dirname, "..", req.file.path);

    // Đọc dữ liệu từ 3 sheet
    const orderData = readExcelSheet(filePath, "", 18).data;
    const webCostData = readExcelSheet(filePath, "", 19).data;
    const ffCostData = readExcelSheet(filePath, "", 20).data;

    // Kiểm tra dữ liệu có rỗng không
    if (!orderData.length || !webCostData.length || !ffCostData.length) {
      fs.unlinkSync(filePath);
      return res.status(400).json({ error: "Một hoặc nhiều sheet trong file Excel rỗng!" });
    }

    // Lọc orderData theo tháng và năm
    const filteredOrderData = orderData.filter(row => {
      const date = row["Date"] ? excelDateToJSDate(row["Date"]) : null;
      return date && date.getMonth() + 1 === month && date.getFullYear() === year;
    });

    // Log dữ liệu từ ba sheet
    console.log("=== Dữ liệu từ sheet Order (Sheet 18) ===");
    console.log("Số dòng (trước lọc):", orderData.length);
    console.log("Số dòng (sau lọc tháng " + month + "/" + year + "):", filteredOrderData.length);
    filteredOrderData.forEach((row, index) => {
      console.log(`Dòng ${index + 1}:`, JSON.stringify(row, null, 2));
    });

    console.log("\n=== Dữ liệu từ sheet Web Cost (Sheet 19) ===");
    console.log("Số dòng:", webCostData.length);
    webCostData.forEach((row, index) => {
      console.log(`Dòng ${index + 1}:`, JSON.stringify(row, null, 2));
    });

    console.log("\n=== Dữ liệu từ sheet FF Cost (Sheet 20) ===");
    console.log("Số dòng:", ffCostData.length);
    ffCostData.forEach((row, index) => {
      console.log(`Dòng ${index + 1}:`, JSON.stringify(row, null, 2));
    });

    // Tính profit cho DesignerID và RAndDID
    const { designerProfit, rdProfit } = assignProfitToDesignerAndRD(filteredOrderData, webCostData, ffCostData);

    // Xóa file tạm
    fs.unlinkSync(filePath);

    // Trả về kết quả
    res.json({
      sheetName: "Designer and R&D Profit",
      month,
      year,
      totalDesignerRecords: designerProfit.length,
      totalRDRecords: rdProfit.length,
      designerProfit,
      rdProfit,
    });
  } catch (error) {
    console.error("Lỗi xử lý file:", error.message, error.stack);
    res.status(500).json({ error: "Xử lý file Excel thất bại! Chi tiết: " + error.message });
  }
}

//Merch
async function uploadMerchOrder(req, res) {
  return uploadFileCommon(req, res, "", 8, processMerchOrder);
}

async function uploadMerchSku(req, res) {
  return uploadFileCommon(req, res, "", 9, processMerchSku);
}

async function uploadMerchProfitByDesignerAndRD(req, res) {
  try {
    const month = parseInt(req.query.month);
    const year = parseInt(req.query.year);

    if (!month || !year) {
      return res.status(400).json({ error: "Vui lòng nhập ?month=...&year=..." });
    }

    if (!req.file) {
      return res.status(400).json({ error: "Vui lòng upload 1 file Excel chứa 2 sheet!" });
    }

    const filePath = path.join(__dirname, "..", req.file.path);

    // Đọc dữ liệu từ 2 sheet
    const orderData = readExcelSheet(filePath, "", 8).data; // Merch Order
    const skuData = readExcelSheet(filePath, "", 9).data;   // Merch SKU

    // Kiểm tra dữ liệu có rỗng không
    if (!orderData.length || !skuData.length) {
      fs.unlinkSync(filePath);
      return res.status(400).json({ error: "Một hoặc cả hai sheet trong file Excel rỗng!" });
    }

    // Lọc orderData theo tháng và năm
    const filteredOrderData = orderData.filter(row => {
      const date = row["Date"] ? excelDateToJSDate(row["Date"]) : null;
      return date && date.getMonth() + 1 === month && date.getFullYear() === year;
    });

    // Log dữ liệu từ hai sheet
    console.log("=== Dữ liệu từ sheet Merch Order (Sheet 18) ===");
    console.log("Số dòng (trước lọc):", orderData.length);
    console.log("Số dòng (sau lọc tháng " + month + "/" + year + "):", filteredOrderData.length);
    filteredOrderData.forEach((row, index) => {
      console.log(`Dòng ${index + 1}:`, JSON.stringify(row, null, 2));
    });

    console.log("\n=== Dữ liệu từ sheet Merch SKU (Sheet 19) ===");
    console.log("Số dòng:", skuData.length);
    skuData.forEach((row, index) => {
      console.log(`Dòng ${index + 1}:`, JSON.stringify(row, null, 2));
    });

    // Tính profit cho DesignerID và RAndDID
    const { designerProfit, rdProfit } = assignProfitToDesignerAndRDMerch(filteredOrderData, skuData);

    // Xóa file tạm
    fs.unlinkSync(filePath);

    // Trả về kết quả
    res.json({
      sheetName: "Designer and R&D Profit",
      month,
      year,
      totalDesignerRecords: designerProfit.length,
      totalRDRecords: rdProfit.length,
      designerProfit,
      rdProfit,
    });
  } catch (error) {
    console.error("Lỗi xử lý file:", error.message, error.stack);
    res.status(500).json({ error: "Xử lý file Excel thất bại! Chi tiết: " + error.message });
  }
}
// excelController.js
async function exportAllProfit(req, res) {
  try {
    const month = parseInt(req.query.month);
    const year = parseInt(req.query.year);

    if (!month || !year) {
      return res.status(400).json({ error: "Thiếu ?month=&year=" });
    }

    if (!req.file) {
      return res.status(400).json({ error: "Vui lòng upload file Excel!" });
    }

    const filePath = path.join(__dirname, "..", req.file.path);

    const [amazon, etsy, web, merch] = await Promise.all([
      getAmazonProfit(filePath, month, year).catch(() => null),
      getEtsyProfit(filePath, month, year).catch(() => null),
      getWebProfit(filePath, month, year).catch(() => null),
      getMerchProfit(filePath, month, year).catch(() => null),
    ]);

    const missing = [];
    if (!amazon) missing.push("Amazon");
    if (!etsy) missing.push("Etsy");
    if (!web) missing.push("Web");
    if (!merch) missing.push("Merch");

    if (missing.length > 0) {
      fs.unlinkSync(filePath);
      return res.status(400).json({
        error: `Thiếu dữ liệu: ${missing.join(", ")}`,
        tip: "Kiểm tra file có đủ sheet không"
      });
    }

    const inputData = {
      amazon,
      etsy: [etsy],
      web,
      merch
    };

    const aggregated = aggregateProfit(inputData);
    const exportPath = exportProfitToExcel(aggregated);

    fs.unlinkSync(filePath);

    res.download(exportPath, `Profit_Summary_${year}_${month}.xlsx`, (err) => {
      if (err) console.error("Download error:", err);
    });

  } catch (error) {
    console.error("Export error:", error);
    if (req.file?.path) fs.unlinkSync(path.join(__dirname, "..", req.file.path));
    res.status(500).json({ error: error.message });
  }
}

module.exports = {
  uploadEmptyPackage,
  uploadBuyingLabel,
  uploadScanLabel,
  uploadEtsyStatement,
  uploadEtsyFFCost,
  uploadEtsyOrder,
  uploadEtsyProfit,
  uploadAmzTransaction,
  uploadAmzFFCost,
  uploadAmzOrder,
  uploadAmzProfit,
  uploadWebOrder,
  uploadWebCost1,
  uploadWebCost2,
  uploadWebTotalCost,
  uploadWebProfit,
  uploadMerchOrder,
  uploadMerchSku,
  uploadMerchProfitByDesignerAndRD,
  exportAllProfit
};
