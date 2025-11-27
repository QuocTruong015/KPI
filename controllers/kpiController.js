const fs = require("fs");
const path = require("path");
const XLSX = require('xlsx');
const { readExcelSheet } = require("../utils/excelUtils");
const { processTargetKpi } = require("../services/kpiService");
const { getEtsyProfit, getAmazonProfit, getWebProfit, getMerchProfit } = require("../services/profitAggregatorService");
const { aggregateProfit } = require("../services/profitAggregatorService");
const { excelDateToJSDate } = require("../utils/excelUtils");
const { processFulfillmentPosterCost } = require("../services/fulfillmentPoster");
const { processEmptyPackage } = require("../services/emptyPackageService");
const { processBuyingLabel } = require("../services/buyingLabelService");
const { processScanLabel } = require("../services/scanLabelService");
const { uploadFulfillmentPosterCost } = require("./excelController");
const { processServiceStaff2 } = require("../services/serviceStaff_2Service");

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
      sheetName: actualSheetName,finalData,});
  } catch (error) {
    console.error(error);
    res.status(500).json({ error: "Đọc file Excel thất bại!" });
  }
}

async function uploadKpiTargetFile(req, res) {
  return uploadFileCommon(req, res, "KPI", 0, processTargetKpi, "totalKpiTargets");
}

async function calculateCombinedKPI(req, res) {
  let profitPath, targetPath, exportPath;

  try {
    const month = parseInt(req.query.month);
    const year = parseInt(req.query.year);

    if (!month || !year || month < 1 || month > 12) {
      return res.status(400).json({ error: "Month (1-12) và year là bắt buộc" });
    }

    // === KIỂM TRA FILE ===
    const profitFile = req.files?.profit_file?.[0] || req.files?.profit_file;
    const targetFile = req.files?.target_file?.[0] || req.files?.target_file;

    const ucs2025File = path.join(__dirname, "..", req.files?.ucs2025_file?.[0]?.path || req.files?.ucs2025_file);
    const daisyFile = path.join(__dirname, "..", req.files?.daisy_file?.[0]?.path || req.files?.daisy_file);
    const ucsPosterFile = path.join(__dirname, "..", req.files?.ucs_poster_file?.[0]?.path || req.files?.ucs_poster_file);
    const ucsSellerManagementFile = path.join(__dirname, "..", req.files?.ucs_seller_management_file?.[0]?.path || req.files?.ucs_seller_management_file);

    if (!profitFile || !targetFile) {
      return res.status(400).json({
        error: "Cần upload 2 file: profit_file và target_file",
        received: Object.keys(req.files || {})
      });
    }

    profitPath = path.join(__dirname, "..", profitFile.path);
    targetPath = path.join(__dirname, "..", targetFile.path);

    // === BƯỚC 1: TÍNH PROFIT ===
    // 1. Lấy profit từng nền tảng
    const [amazon, etsy, web, merch] = await Promise.all([
      getAmazonProfit(profitPath, month, year).catch(err => { console.log("Amazon:", err); return null; }),
      getEtsyProfit(profitPath, month, year).catch(err => { console.log("Etsy:", err); return null; }),
      getWebProfit(profitPath, month, year).catch(err => { console.log("Web:", err); return null; }),
      getMerchProfit(profitPath, month, year).catch(err => { console.log("Merch:", err); return null; }),
    ]);

    if (!amazon || !etsy || !web || !merch) {
      return res.status(400).json({ error: "File Profit thiếu dữ liệu từ một hoặc nhiều nền tảng" });
    }

    // 2. Tính profit SS1
    const data1 = readExcelSheet(ucs2025File, "FF Cost - Ship by Tiktok", 19).data; //UCS_2025
    const data2 = readExcelSheet(ucs2025File, "FF Cost - Ship by seller", 18).data; //UCS_2025
    const data3 = readExcelSheet(ucs2025File, "FF Refund - Sellers", 14).data; //UCS_2025
    const data4 = readExcelSheet(daisyFile, "Poster US", 0).data; //Daisy
    const data5 = readExcelSheet(ucs2025File, "UCS - Buying label", 10).data; //UCS_2025
    const data6 = readExcelSheet(ucs2025File, "FF Order", 11).data; //UCS_2025
    const data7 = readExcelSheet(ucs2025File, "FF Phone Case", 12).data; //UCS_2025
    const data8 = readExcelSheet(ucs2025File, "FF Revenue - Sellers", 13).data; //UCS_2025
    const data9 = readExcelSheet(ucsPosterFile, "Fulfillment", 3).data; //UCS Seller Management
    const data10 = readExcelSheet(ucs2025File, "OTHERS PROJECT", 4).data; //UCS_2025

    const serviceStaff1Data = processFulfillmentPosterCost (data1, data2, data3, data4, data5, data6, data6, data6, data7, data8, data9, data10, month, year);

    // const data11 = readExcelSheet(ucs2025File, "Empty Package", 8).data; //UCS_2025
    // const data12 = readExcelSheet(ucs2025File, "Buying Label", 9).data; //UCS_2025
    const data13 = readExcelSheet(ucs2025File, "SCAN LABEL", 3).data; //UCS_2025

    // const emptyPackage = processEmptyPackage(data11, month, year);
    // const buyingLabel = processBuyingLabel(data12, month, year);
    const scanLabel = processScanLabel(data13, month, year);

    // const emptyProfit = emptyPackage.emptyTotalProfit;
    // const buyingLabelProfit = buyingLabel.buyingTotalProfit;
    const scanLabelProfit = scanLabel.scanLabelTotalProfit;

    const totalSS1Profit = serviceStaff1Data.TotalProfitSS1 + scanLabelProfit;

    // 3. Tính profit SS2
    const ss2data1 = readExcelSheet(ucsSellerManagementFile, "Buying Labels", 9).data; //UCS_seller management
    const ss2data2 = readExcelSheet(ucs2025File, "OTHERS PROJECT", 4).data; //UCS_2025
    const ss2data3 = readExcelSheet(ucs2025File, "SCAN LABEL", 3).data; //UCS_2025
    const ss2data4 = readExcelSheet(ucs2025File, "KPI Detail", 6).data; //UCS_2025

    const serviceStaff2ProfitMap = processServiceStaff2(ss2data1, ss2data2, ss2data3, ss2data4, month, year);

    // === QUAN TRỌNG: TÁI TẠO LẠI CẤU TRÚC GIỐNG HÀM EXPORT GỐC ===
    const inputData = {
      amazon,
      etsy: [etsy],   // GIỮ ĐÚNG FORMAT CỦA HÀM GỐC
      web,
      merch
    };

    // === AGGREGATE PROFIT GIỐNG HÀM GỐC ===
    const aggregated = aggregateProfit(inputData);

    const csmProfit = aggregated.mainPlatformProfit;
    const designerProfit = aggregated.designerProfit;
    const rdProfit = aggregated.rdProfit;

    // === BƯỚC 2: ĐỌC TARGET ===
    let targetData = readExcelSheet(targetPath, "KPI", 0).data;

    const filtered = targetData.filter((row, index) => {
      const date = excelDateToJSDate(row.Month);
      const isValidDate = date && !isNaN(date.getTime());
    
      if (!isValidDate) {
        console.warn(`Row ${index + 2}: Ngày không hợp lệ (${row.Month})`);
        return false;
      }

      return date.getMonth() + 1 === month && date.getFullYear() === year;
    });

    // === BƯỚC 3: KẾT HỢP & TÍNH KPI ===
    const result = filtered.map(t => {

      // Lấy mã pic
      const picKey =
        t.PIC?.match(/\(([^)]+)\)/)?.[1]?.trim() ||
        t.PIC?.trim();

      let profit = 0;
      const kpiDesc = t['KPI Desciption'];

      // --- GÁN PROFIT THEO POSITION ---
      if (t.Position === "R&D") {
        profit = rdProfit[picKey] || 0;
      } 
      else if (t.Position === "Designer") {
        profit = designerProfit[picKey] || 0;
      } 
      else if (t.Position === "CSM - Bán hàng") {
        profit = csmProfit;
      } 
      else if (t.Position === "Service Staff" && kpiDesc.includes("Tổng lợi nhuận gộp của toàn bộ mảng dịch vụ")) {
        profit = totalSS1Profit;  
      }
      else if (t.Position === "Service Staff" && kpiDesc.includes("Tổng lợi nhuận gộp từ khách hàng do người thực hiện KPI chốt được trong 3")) {
        // Tra cứu lợi nhuận cá nhân (dạng {PIC: Profit})
        profit = serviceStaff2ProfitMap[picKey] || 0;  
        console.log(`SS2 Profit for ${picKey}: ${profit}`);
      }

      // --- TÍNH KPI ---
      const kpi = t["Target (100%)"] > 0
        ? (profit / t["Target (100%)"]) * 100
        : 0;

      return {
        PIC: t.PIC,
        PIC_Key: picKey,
        Description: t['KPI Desciption'],
        Position: t.Position,
        Profit: profit,
        Target: t["Target (100%)"],
        KPI: kpi.toFixed(2) + '%'
      };
    });


    // === BƯỚC 4: XUẤT FILE ===
    const exportDir = path.join(__dirname, '..', 'exports');
    if (!fs.existsSync(exportDir)) {
      fs.mkdirSync(exportDir, { recursive: true });
    }

    exportPath = path.join(exportDir, `KPI_Result_${year}_${month}.xlsx`);

    const ws = XLSX.utils.json_to_sheet(result);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'KPI');
    XLSX.writeFile(wb, exportPath);

    // === GỬI FILE ===
    return res.json({
        message: "Xuất file thành công",
        file: `/exports/KPI_Result_${year}_${month}.xlsx`
    });
  } catch (error) {
    console.error("Lỗi trong calculateCombinedKPI:", error);
    return res.status(500).json({ error: "Lỗi server nội bộ" });
  }
}

module.exports = { uploadKpiTargetFile, calculateCombinedKPI };