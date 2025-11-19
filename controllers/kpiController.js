const fs = require("fs");
const path = require("path");
const XLSX = require('xlsx');
const { readExcelSheet } = require("../utils/excelUtils");
const { processTargetKpi } = require("../services/kpiService");
const { getEtsyProfit, getAmazonProfit, getWebProfit, getMerchProfit } = require("../services/profitAggregatorService");

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

    profitPath = profitFile.path;
    targetPath = targetFile.path;

    // === BƯỚC 1: TÍNH PROFIT ===
    const [amazon, etsy, web, merch] = await Promise.all([
      getAmazonProfit(profitPath, month, year).catch(() => null),
      getEtsyProfit(profitPath, month, year).catch(() => null),
      getWebProfit(profitPath, month, year).catch(() => null),
      getMerchProfit(profitPath, month, year).catch(() => null),
    ]);

    if (!amazon || !etsy || !web || !merch) {
      return res.status(400).json({ error: "File Profit thiếu dữ liệu từ một hoặc nhiều nền tảng" });
    }

    const designerProfit = {
      ...amazon.designerProfit,
      ...etsy.designerProfit,
      ...web.designerProfit,
      ...merch.designerProfit
    };

    const rdProfit = {
      ...amazon.rdProfit,
      ...etsy.rdProfit,
      ...web.rdProfit,
      ...merch.rdProfit
    };

    // === BƯỚC 2: ĐỌC TARGET ===
    const workbook = XLSX.readFile(targetPath);
    const sheet = workbook.Sheets[workbook.SheetNames[0]];
    const rawData = XLSX.utils.sheet_to_json(sheet, { defval: null });
    const targetList = processTargetKpi(rawData);

    const validTargets = targetList
      .filter(t => t.Month?.month === month && t.Month?.year === year)
      .filter(t => ['Designer', 'R&D', 'designer', 'r&d'].includes(t.Position));

    if (validTargets.length === 0) {
      return res.status(400).json({ error: `Không tìm thấy Target cho tháng ${month}/${year}` });
    }

    // === BƯỚC 3: KẾT HỢP & TÍNH KPI ===
    const result = validTargets.map(t => {
      const isRD = /r&d/i.test(t.Position);
      const profitMap = isRD ? rdProfit : designerProfit;

      // 🔹 Lấy phần mã trong ngoặc, ví dụ "huy (TH)" -> "TH"
      const picKey = t.PIC?.match(/\(([^)]+)\)/)?.[1]?.trim() || t.PIC?.trim();

      const profit = profitMap[picKey] || 0;
      const kpi = t.Target > 0 ? (profit / t.Target) * 100 : 0;

      return {
        PIC: t.PIC,
        PIC_Key: picKey, // để debug nếu cần
        Position: t.Position,
        Profit: profit,
        Target: t.Target,
        KPI: kpi.toFixed(2) + '%'
      };
    });

    // === BƯỚC 4: XUẤT FILE EXCEL ===
    const exportDir = path.join(__dirname, '..', 'exports');
    if (!fs.existsSync(exportDir)) {
      fs.mkdirSync(exportDir, { recursive: true });
    }

    exportPath = path.join(exportDir, `KPI_Result_${year}_${month}.xlsx`);
    console.log("Đường dẫn xuất:", exportPath);

    const ws = XLSX.utils.json_to_sheet(result);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'KPI');
    XLSX.writeFile(wb, exportPath);

    console.log("File đã tạo:", fs.existsSync(exportPath));

    // === GỬI FILE VỀ CLIENT ===
    res.download(exportPath, `KPI_Result_${year}_${month}.xlsx`, (err) => {
      if (err) {
        console.error("Lỗi tải file:", err);
        if (!res.headersSent) res.status(500).json({ error: "Không thể tải file" });
      } else {
        console.log("File đã gửi về client");
      }
    });

  } catch (error) {
    console.error("Combined KPI Error:", error);
    [profitPath, targetPath, exportPath].forEach(p => {
      if (p && fs.existsSync(p)) {
        try { fs.unlinkSync(p); } catch {}
      }
    });
    if (!res.headersSent) {
      res.status(500).json({ error: error.message });
    }
  }
}

module.exports = { uploadKpiTargetFile, calculateCombinedKPI };