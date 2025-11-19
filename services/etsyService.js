const { excelDateToJSDate } = require("../utils/excelUtils");
const XLSX = require("xlsx");
const { Parser: FormulaParser } = require('hot-formula-parser');
const { is } = require("date-fns/locale");

const globalParser = new FormulaParser();
// Helper: Chuẩn hóa ID (loại bỏ dấu cách, chuẩn hóa định dạng)
function normalizeId(id) {
  if (!id || id === "Unknown" || id === "") return null;
  return id.toString().trim().toUpperCase(); // Loại bỏ dấu cách, chuyển về chữ hoa
}

// Helper: sanitize and parse amount strings with various formats/currency symbols
function sanitizeAmount(raw) {
  if (raw == null) return { amount: 0, currency: "UNKNOWN", raw: "" };
  let s = String(raw).trim();

  // Normalize common unicode spaces
  s = s.replace(/\u00A0/g, "").trim(); // non-breaking space

  // Detect negative via parentheses (e.g. (CA$1,234.56)) or leading '-'
  let isNegative = false;
  if (/^\(.*\)$/.test(s)) {
    isNegative = true;
    s = s.replace(/^\(|\)$/g, "");
  }
  if (/^-+/.test(s)) {
    isNegative = true;
    s = s.replace(/^-+/, "");
  }

  // Detect currency symbols anywhere and remove them
  let currency = "UNKNOWN";
  const currencyPatterns = [
    ["CAD", /CA\$|CAD/i],
    ["VND", /₫|VND/i],
    ["USD", /\$/i],
  ];
  for (const [cur, pat] of currencyPatterns) {
    if (pat.test(s)) {
      currency = cur;
      s = s.replace(pat, "");
      break;
    }
  }

  // Remove any spaces now
  s = s.replace(/\s+/g, "");

  // Heuristics for decimal/thousand separators
  const lastDot = s.lastIndexOf('.');
  const lastComma = s.lastIndexOf(',');
  if (lastDot !== -1 && lastComma !== -1) {
    if (lastComma > lastDot) {
      // comma is decimal, dot is thousands
      s = s.replace(/\./g, '');
      s = s.replace(/,/g, '.');
    } else {
      // dot is decimal, comma is thousands
      s = s.replace(/,/g, '');
    }
  } else if (lastComma !== -1 && lastDot === -1) {
    const partAfter = s.slice(lastComma + 1);
    if (/^\d{1,2}$/.test(partAfter)) {
      // comma used as decimal
      s = s.replace(/,/g, '.');
    } else {
      // comma used as thousands separator
      s = s.replace(/,/g, '');
    }
  } else {
    // remove any grouping commas
    s = s.replace(/,/g, '');
  }

  // Remove any remaining non-digit/dot characters
  s = s.replace(/[^0-9.\-]/g, '');

  let num = parseFloat(s);
  if (isNaN(num)) num = 0;
  if (isNegative) num = -Math.abs(num);

  return { amount: num, currency, raw: String(raw) };
}

// Hàm validate row (sử dụng chung cho processEtsyStatement)
function validateRow(row) {
  const requiredFields = ["Date", "Type", "Order ID (sale, refund)"];
  const missingFields = requiredFields.filter((field) => !row[field] || String(row[field]).trim() === "");
  return missingFields.length === 0 ? null : `Thiếu cột: ${missingFields.join(", ")}`;
}

function parseNetValue(value) {
  if (!value) return 0;
  // Chuyển sang chuỗi, loại bỏ ký tự $, CA, khoảng trắng, dấu phẩy
  let cleaned = String(value).replace(/CA\$|\$|,/g, "").trim();

  // parseFloat sẽ tự nhận dấu âm nếu có
  const num = parseFloat(cleaned);
  return isNaN(num) ? 0 : num;
}

// === HÀM XỬ LÝ ETSY STATEMENT ===
function processEtsyStatement(data, month, year) {
  if (!Array.isArray(data) || data.length === 0) {
    throw new Error("Dữ liệu Excel rỗng hoặc không hợp lệ");
  }

  // ===== B1: Lọc dữ liệu theo tháng & năm =====
  const filtered = data.filter((row, index) => {
    const date = excelDateToJSDate(row.Date);
    const isValidDate = date && !isNaN(date.getTime());
    if (!isValidDate) {
      console.warn(`Row ${index + 2}: Ngày không hợp lệ (${row.Date})`);
      return false;
    }
    return date.getMonth() + 1 === month && date.getFullYear() === year;
  });

  // ===== B2: Trích xuất Order ID =====
  const extractOrderId = (info) => {
    if (!info || info === "--") return "unknown";
    const match = info.match(/(?<=[:#])\s*(\d+)/);
    return match ? match[1] : "unknown";
  };

  const dataWithOrderId = filtered.map((row) => ({
    ...row,
    ExtractedOrderID: extractOrderId(row["Info"] || row["Title"]),
  }));

  // ===== B3: Hàm parse tiền =====
  const parseCurrency = (str) => {
    if (str === null || str === undefined || str === "--") return 0;
    const s = String(str).trim();
    if (s === "" || s.toLowerCase() === "nan") return 0;
    const cleaned = parseFloat(s.replace(/CA\$|,/g, "").replace("-", ""));
    if (isNaN(cleaned)) return 0;
    return s.includes("-") ? -cleaned : cleaned;
  };

  // ===== B4: Gộp theo (StoreID + OrderID) =====
  const groupedByStoreAndOrder = {};

  dataWithOrderId.forEach((row) => {
    const storeId = row["Store ID"];
    const orderId = row["ExtractedOrderID"];
    const key = `${storeId}_${orderId}`; // gộp theo cặp khóa

    if (!groupedByStoreAndOrder[key]) {
      groupedByStoreAndOrder[key] = {
        StoreID: storeId,
        OrderID: orderId,
        Month: row["Month"],
        Currency: row["Currency"],
        FeesAndTaxes: 0,
        SaleRev: 0,
        TaxAds: 0,
        MarketingFee: 0,
        ListingFee: 0.2,
      };
    }

    let rev = 0,
      marketingFee = 0,
      feesAndTaxes = 0,
      totalTaxAds = 0;

    if (row["Type"] === "Refund" || row["Type"] === "Sale") {
      rev = parseCurrency(row["Net"]) / 1.37;
    } else if (row["Type"].includes("Marketing")) {
      marketingFee = parseCurrency(row["Net"]) / 1.37;
    } else if (row["Title"].includes("Tax: Etsy Ads")) {
      totalTaxAds = parseCurrency(row["Net"]) / 1.37;
    } else {
      feesAndTaxes = parseCurrency(row["Net"]) / 1.37;
    }

    const g = groupedByStoreAndOrder[key];
    g.FeesAndTaxes += Math.abs(feesAndTaxes);
    g.TaxAds += Math.abs(totalTaxAds);
    g.MarketingFee += Math.abs(marketingFee);
    g.SaleRev += Math.abs(rev);
  });

  const groupedData = Object.values(groupedByStoreAndOrder);

  // ===== B5: Chia đều TaxAds & MarketingFee của “unknown” cho các order cùng store =====
  const storesMap = {};
  groupedData.forEach((o) => {
    if (!storesMap[o.StoreID]) storesMap[o.StoreID] = [];
    storesMap[o.StoreID].push(o);
  });

  Object.values(storesMap).forEach((orders) => {
    const unknownOrders = orders.filter((o) => o.OrderID === "unknown");
    const validOrders = orders.filter((o) => o.OrderID !== "unknown");
    if (unknownOrders.length === 0 || validOrders.length === 0) return;

    const totalTaxAdsUnknown = unknownOrders.reduce((sum, o) => sum + o.TaxAds, 0);
    const totalMarketingUnknown = unknownOrders.reduce((sum, o) => sum + o.MarketingFee, 0);

    const shareTax = totalTaxAdsUnknown / validOrders.length;
    const shareMarketing = totalMarketingUnknown / validOrders.length;

    validOrders.forEach((o) => {
      o.TaxAds += shareTax;
      o.MarketingFee += shareMarketing;
    });

    // Xóa order "unknown"
    orders.splice(0, orders.length, ...validOrders);
  });

  // ===== B6: Gộp lại tất cả orders sau chia đều =====
  const finalResult = Object.values(storesMap).flat();

  // ===== B7: Làm tròn số =====
  return finalResult.map((g) => ({
    ...g,
    FeesAndTaxes: g.FeesAndTaxes.toFixed(2),
    SaleRev: g.SaleRev.toFixed(2),
    TaxAds: g.TaxAds.toFixed(2),
    MarketingFee: g.MarketingFee.toFixed(2),
  }));
}

// Hàm xử lý dữ liệu Etsy FFCost
function processEtsyFFCost(data) {
  if (!Array.isArray(data) || data.length === 0) {
    throw new Error("Dữ liệu Excel rỗng hoặc không hợp lệ");
  }

  const result = data.map((row, index) => {
    // Lấy và xử lý OrderName (chuỗi số)
    let orderName = row["OrderName"];
    if (orderName == null || orderName === undefined) {
      orderName = "Unknown";
    } else {
      orderName = String(orderName); // Đảm bảo là chuỗi, giữ nguyên số
    }

    // Lấy NetPrice (đã xử lý, chỉ cần parse thành số)
    const netPrice = row["NetPrice"] != null ? parseFloat(row["NetPrice"]) || 0 : 0;

    return {
      OrderName: orderName,
      StoreID: row["Store ID"]?.trim() || "Unknown",
      Cost: netPrice,
      Supplier: row["Supplier"]?.trim() || "Unknown",
    };
  });

  console.log(`Processed ${result.length} rows for Etsy FFCost`);
  return result;
}

// Hàm xử lý dữ liệu Etsy Order
function processEtsyOrder(data, month, year) {
  if (!Array.isArray(data) || data.length === 0) {
    throw new Error("Dữ liệu Excel rỗng hoặc không hợp lệ");
  }

  const filtered = data.filter((row, index) => {
    const saleDate = excelDateToJSDate(row["Sale Date"]);
    const isValidDate = saleDate && !isNaN(saleDate.getTime());

    if (!isValidDate) {
      console.warn(`Row ${index + 2}: Sale Date không hợp lệ (${row["Sale Date"]})`);
      return false;
    }

    const isValidPeriod = saleDate.getMonth() + 1 === month && saleDate.getFullYear() === year;
    if (!isValidPeriod) return false;

    return true;
  });

  const result = filtered.map((row, index) => {
    // Xử lý Sale Date
    const saleDate = excelDateToJSDate(row["Sale Date"]);
    if (!saleDate || isNaN(saleDate.getTime())) {
      console.warn(`Row ${index + 2}: Sale Date không hợp lệ (${row["Sale Date"]})`);
    }

    // Trích xuất Designer ID và R&D ID từ SKU
    let designerId = "Unknown";
    let rAndDId = "Unknown";
    const sku = row["SKU"]?.trim() || "";
    if (sku) {
      const parts = sku.split("-");
      if (parts.length >= 2) {
        designerId = parts[0] || "Unknown";
        rAndDId = parts[1] || "Unknown";
      } else {
        console.warn(`Row ${index + 2}: SKU không đúng định dạng (${sku})`);
      }
    }

    return {
      SaleDate: saleDate || null,
      OrderID: String(row["Order ID"] || "").trim(),
      SKU: sku,
      StoreID: String(row["Store ID "] || "").trim() || "Unknown",
      DesignerID: designerId,
      RAndDID: rAndDId,
    };
  }).filter(row => row.SaleDate !== null);
  
  console.log(`Processed ${result.length} rows for Etsy Order`);
  return result;
}
function calculateEtsyProfit(statementData, ffCostData, orderData, month, year) {
  // Xử lý 3 nguồn dữ liệu
  const statementProcessed = processEtsyStatement(statementData, month, year);
  const ffCostProcessed = processEtsyFFCost(ffCostData);
  const orderProcessed = processEtsyOrder(orderData, month, year);

  // Chuẩn hóa key OrderID + StoreID
  const normalizeKey = (orderId, storeId) => {
    const cleanOrderId = String(orderId || "").trim().replace(/-/g, "");
    const cleanStoreId = String(storeId || "").trim();
    return `${cleanOrderId}|${cleanStoreId}`;
  };

  // Map dữ liệu để tra nhanh
  const ffCostMap = new Map(ffCostProcessed.map(r => [normalizeKey(r.OrderName, r.StoreID), r]));
  const orderMap = new Map(orderProcessed.map(r => [normalizeKey(r.OrderID, r.StoreID), r]));

  // Hàm chuyển chuỗi tiền về số an toàn
  const toNum = (v) => {
    if (v === null || v === undefined) return 0;
    if (typeof v === "number") return v;
    const s = String(v).replace(/CA\$|,|\$/g, "").trim();
    const n = parseFloat(s);
    return isNaN(n) ? 0 : n;
  };

  // Bắt đầu tính profit
  const result = [];
  statementProcessed.forEach((stRow) => {
    const key = normalizeKey(stRow.OrderID, stRow.StoreID);
    const ffCostRow = ffCostMap.get(key);
    const orderRow = orderMap.get(key);

    const revenue = toNum(stRow.SaleRev);
    const totalCost =
      toNum(ffCostRow?.Cost) +
      toNum(stRow.FeesAndTaxes) +
      toNum(stRow.MarketingFee) +
      toNum(stRow.ListingFee) +
      toNum(stRow.TaxAds);

    const profit = revenue - totalCost;

    result.push({
      OrderID: stRow.OrderID,
      StoreID: stRow.StoreID,
      Revenue: Number(revenue.toFixed(2)),
      Cost: Number(totalCost.toFixed(2)),
      Profit: Number(profit.toFixed(2)),
      DesignerID: orderRow ? normalizeId(orderRow.DesignerID) : null,
      RAndDID: orderRow ? normalizeId(orderRow.RAndDID) : null,
      SKU: orderRow?.SKU || "",
      _matchedFFCost: !!ffCostRow,
      _matchedOrder: !!orderRow,
    });
  });

  return result;
}


// Hàm tính KPI cho Etsy
function calculateKPI(statementData, ffCostData, orderData, customData, month, year) {
  // Gọi calculateEtsyProfit để lấy dữ liệu gộp
  const profitData = calculateEtsyProfit(statementData, ffCostData, orderData, month, year);
  
  // Đọc dữ liệu Custom Order
  const customOrderData = readCustomOrder(customData, month, year);

  // Tạo object để tổng hợp Profit cho DesignerID và R&DID
  const designerProfitTotal = {};
  const rdProfitTotal = {};

  // Duyệt qua profitData
  profitData.forEach(row => {
    const { DesignerID, RAndDID, Profit, OrderID } = row;
    const roundedProfit = Number(Profit.toFixed(2));

    // Kiểm tra trùng với CustomOrderData
    const isCustomMatch = customOrderData.some(custom =>
      custom.OrderID === OrderID && custom.DesignerID === DesignerID
    );

    let designerProfitToAdd = roundedProfit;
    if (isCustomMatch) {
      designerProfitToAdd = roundedProfit * 2; // nhân đôi nếu trùng
      console.log(`✅ Custom match found! OrderID=${OrderID}, Designer=${DesignerID}, Profit x2`);
    }

    // Gán cho Designer
    if (DesignerID) {
      designerProfitTotal[DesignerID] = Number(
        ((designerProfitTotal[DesignerID] || 0) + designerProfitToAdd).toFixed(2)
      );
    }

    // Gán cho R&D (giữ nguyên)
    if (RAndDID) {
      rdProfitTotal[RAndDID] = Number(
        ((rdProfitTotal[RAndDID] || 0) + roundedProfit).toFixed(2)
      );
    }
  });

  // Tổng hợp kết quả
  const totalDesignerProfit = Object.values(designerProfitTotal).reduce((sum, p) => sum + p, 0);
  const totalRDProfit = Object.values(rdProfitTotal).reduce((sum, p) => sum + p, 0);
  const totalOrderProfit = profitData.reduce((sum, r) => sum + Number(r.Profit.toFixed(2)), 0);

  console.log(`Calculated KPI for month: ${month}, year: ${year}`);
  console.log("Designer Profit Total:", designerProfitTotal);
  console.log("R&D Profit Total:", rdProfitTotal);
  console.log("Total Designer Profit:", totalDesignerProfit);
  console.log("Total R&D Profit:", totalRDProfit);
  console.log("Total Order Profit:", totalOrderProfit);

  return {
    designerProfit: designerProfitTotal,
    rdProfit: rdProfitTotal
  };
}

function calculateProfitByStoreID(statementData, ffCostData, orderData, month, year) {
  const profitDetails = calculateEtsyProfit(statementData, ffCostData, orderData, month, year);

  if (!profitDetails || profitDetails.length === 0) {
    console.warn("Không có dữ liệu profit để tổng hợp theo StoreID");
    return [];
  }

  const profitMap = new Map();

  profitDetails.forEach(row => {
    let storeId = String(row.StoreID || "").trim();
    if (!storeId) storeId = "UNKNOWN";

    const profit = Number(row.Profit) || 0;

    if (profitMap.has(storeId)) {
      const curr = profitMap.get(storeId);
      profitMap.set(storeId, {
        TotalProfit: curr.TotalProfit + profit,
        OrderCount: curr.OrderCount + 1
      });
    } else {
      profitMap.set(storeId, { TotalProfit: profit, OrderCount: 1 });
    }
  });

  const result = Array.from(profitMap, ([StoreID, data]) => ({
    StoreID,
    TotalProfit: Number(data.TotalProfit.toFixed(2)),
    OrderCount: data.OrderCount
  }));

  result.sort((a, b) => b.TotalProfit - a.TotalProfit);
  console.log(`Tổng hợp ${result.length} StoreID`);
  return result;
}

function readCustomOrder(data, month, year) {
  if (!Array.isArray(data) || data.length === 0) {
    throw new Error("Dữ liệu Excel rỗng hoặc không hợp lệ");
  }

  const result = data
    .map((row, index) => {
      const keys = Object.keys(row);
      const designerColIndex = keys.indexOf("Assignee");

      // Tạo đối tượng row
      const rowData = {
        Date: excelDateToJSDate(row["Last Modified Date"]),
        Task_Name: String(row["Task name"] || "").trim(),
        DesignerID: String(row[keys[designerColIndex + 1]] || "").trim(),
        OrderID: String(row["Order ID"] || "").trim(),
      };

      if (
        rowData.DesignerID &&
        rowData.OrderID &&
        rowData.Date instanceof Date &&
        !isNaN(rowData.Date) &&
        rowData.Date.getMonth() + 1 === month && // getMonth() trả về 0-11, nên +1 để khớp với month (1-12)
        rowData.Date.getFullYear() === year
      ) {
        return rowData;
      }
      return null;
    })
    .filter(row => row !== null); // Loại bỏ các row null

  console.log(`Processed ${result.length} rows for Custom Order in ${month}/${year}`);
  return result;
}

module.exports = { processEtsyStatement, processEtsyFFCost, processEtsyOrder, calculateEtsyProfit, calculateKPI, calculateProfitByStoreID, readCustomOrder };