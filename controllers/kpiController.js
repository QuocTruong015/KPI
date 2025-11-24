const fs = require("fs");
const path = require("path");
const XLSX = require('xlsx');
const { readExcelSheet } = require("../utils/excelUtils");
const { processTargetKpi } = require("../services/kpiService");
const { getEtsyProfit, getAmazonProfit, getWebProfit, getMerchProfit } = require("../services/profitAggregatorService");
const { aggregateProfit } = require("../services/profitAggregatorService");
const { excelDateToJSDate } = require("../utils/excelUtils");
const { ta, fi } = require("date-fns/locale");

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

    if (!profitFile || !targetFile) {
      return res.status(400).json({
        error: "Cần upload 2 file: profit_file và target_file",
        received: Object.keys(req.files || {})
      });
    }

    // ==== GIỐNG HÀM GỐC: SỬ DỤNG PATH JOIN ====
    profitPath = path.join(__dirname, "..", profitFile.path);
    targetPath = path.join(__dirname, "..", targetFile.path);

    // === BƯỚC 1: TÍNH PROFIT ===
    const [amazon, etsy, web, merch] = await Promise.all([
      getAmazonProfit(profitPath, month, year).catch(err => { console.log("Amazon:", err); return null; }),
      getEtsyProfit(profitPath, month, year).catch(err => { console.log("Etsy:", err); return null; }),
      getWebProfit(profitPath, month, year).catch(err => { console.log("Web:", err); return null; }),
      getMerchProfit(profitPath, month, year).catch(err => { console.log("Merch:", err); return null; }),
    ]);

    if (!amazon || !etsy || !web || !merch) {
      return res.status(400).json({ error: "File Profit thiếu dữ liệu từ một hoặc nhiều nền tảng" });
    }

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

    console.log("=== Designer Profit & R&D Profit (Final Aggregated) ===");
    console.log(designerProfit);
    console.log(rdProfit);
    console.log("=== CSM Profit (Final Aggregated) ===");
    console.log(csmProfit);

    // === BƯỚC 2: ĐỌC TARGET ===
    let targetData = readExcelSheet(targetPath, "KPI", 0).data;

    const filtered = targetData.filter((row, index) => {
      const date = excelDateToJSDate(row.Month);
      const isValidDate = date && !isNaN(date.getTime());
      if (row.Position == null || row.Position.toString().trim() === "") {
        return false;
      }
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

      // --- GÁN PROFIT THEO POSITION ---
      if (t.Position === "R&D") {
        profit = rdProfit[picKey] || 0;
      } 
      else if (t.Position === "Designer") {
        profit = designerProfit[picKey] || 0;
      } 
      else if (t.Position === "CSM - Bán hàng") {
        profit = csmProfit;   // 👈 TẤT CẢ CSM LẤY CHUNG SỐ NÀY
      }

      // --- TÍNH KPI ---
      const kpi = t["Target (100%)"] > 0
        ? (profit / t["Target (100%)"]) * 100
        : 0;

      return {
        PIC: t.PIC,
        PIC_Key: picKey,
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