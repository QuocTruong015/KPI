const fs = require("fs");
const path = require("path");
const { readExcelSheet } = require("../utils/excelUtils");
const { processEtsyStatement } = require("../services/etsyService");
const { processEmptyPackage } = require("../services/emptyPackageService");
const { processBuyingLabel } = require("../services/buyingLabelService");
const { processScanLabel } = require("../services/scanLabelService");
const { processEtsyFFCost, processEtsyOrder } = require("../services/etsyService");

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



module.exports = {
  uploadEmptyPackage,
  uploadBuyingLabel,
  uploadScanLabel,
  uploadEtsyStatement,
  uploadEtsyFFCost,
  uploadEtsyOrder,
};
