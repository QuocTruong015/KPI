const { excelDateToJSDate } = require("../utils/excelUtils");

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

// Hàm xử lý dữ liệu Etsy Statement
function processEtsyStatement(data, month, year) {
  if (!Array.isArray(data) || data.length === 0) {
    throw new Error("Dữ liệu Excel rỗng hoặc không hợp lệ");
  }

    const filtered = data.filter((row, index) => {
    const date = excelDateToJSDate(row.Date);
    const isValidDate = date && !isNaN(date.getTime());
    
    if (!isValidDate) {
      console.warn(`Row ${index + 2}: Ngày không hợp lệ (${row.Date})`);
      return false;
    }

    const isValidPeriod = date.getMonth() + 1 === month && date.getFullYear() === year;
    if (!isValidPeriod) return false;

    const validationError = validateRow(row);
    if (validationError) {
      console.warn(`Row ${index + 2}: ${validationError}`);
      return false;
    }

    return true;
  });

  const result = filtered.map((row, index) => {
    const amountKey = Object.keys(row).find((k) => k.toLowerCase().includes("amount")) || "Amount";
    let rawAmount = amountKey && row[amountKey] != null ? String(row[amountKey]).trim() : "0";
    let currency = "UNKNOWN";
    let cleanAmount = rawAmount;

    // Nhận diện và loại bỏ ký hiệu tiền tệ
    if (rawAmount.startsWith("-CA$")) {
      cleanAmount = rawAmount.replace("-CA$", "").trim();
      currency = "CAD";
    } else if (rawAmount.startsWith("-₫")) {
      cleanAmount = rawAmount.replace("-₫", "").trim();
      currency = "VND";
    } else if (rawAmount.startsWith("-$")) {
      cleanAmount = rawAmount.replace("-$", "").trim();
      currency = "USD";
    } else if (rawAmount.startsWith("CA$")) {
      cleanAmount = rawAmount.replace("CA$", "").trim();
      currency = "CAD";
    } else if (rawAmount.startsWith("₫")) {
      cleanAmount = rawAmount.replace("₫", "").trim();
      currency = "VND";
    } else if (rawAmount.startsWith("$")) {
      cleanAmount = rawAmount.replace("$", "").trim();
      currency = "USD";
    } else {
      console.warn(`Row ${index + 2}: Ký hiệu tiền tệ không nhận diện được: ${rawAmount}`);
    }

    // Xử lý dấu phẩy (ngàn)
    cleanAmount = cleanAmount.replace(/,/g, "");
    const amount = parseFloat(cleanAmount) || 0;
    const isNegative = rawAmount.startsWith("-");

    // Tính Revenue (rev) dựa trên currency
    let rev = 0;
    if (currency === "CAD") {
      rev = amount / 1.37;
    } else if (currency === "USD") {
      rev = amount;
    } else if (currency === "VND") {
      rev = amount / 26000;
    } else {
      rev = amount; // Mặc định nếu không nhận diện được currency
    }
    if (isNegative) {
      rev = -rev;
    }

    return {
      Date: excelDateToJSDate(row.Date),
      Type: String(row["Type"] || "").trim(),
      Title: String(row["Title"] || "").trim(),
      Currency: currency,
      Amount: isNaN(amount) ? 0 : (isNegative ? -amount : amount),
      StoreID: String(row["Store ID"] || "").trim(),
      OrderID: String(row["Order ID (sale, refund)"] || "").trim(),
      Revenue: isNaN(rev) ? 0 : Number(rev.toFixed(2)), // Làm tròn 2 chữ số thập phân
    };
  });

  console.log(`Processed ${result.length}/${data.length} rows for Etsy Statement (month: ${month}, year: ${year})`);
  return result;
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
function processEtsyOrder(data) {
  if (!Array.isArray(data) || data.length === 0) {
    throw new Error("Dữ liệu Excel rỗng hoặc không hợp lệ");
  }

  const result = data.map((row, index) => {
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
  });

  console.log(`Processed ${result.length} rows for Etsy Order`);
  return result;
}

// Hàm tính profit cho Etsy
function calculateEtsyProfit(statementData, ffCostData, orderData, month, year) {
  // Xử lý dữ liệu từ các hàm hiện có
  const statementProcessed = processEtsyStatement(statementData, month, year);
  const ffCostProcessed = processEtsyFFCost(ffCostData);
  const orderProcessed = processEtsyOrder(orderData);

  // Tạo map để dễ tra cứu theo OrderID và StoreID
  const statementMap = new Map();
  statementProcessed.forEach(row => {
    const key = `${row.OrderID}|${row.StoreID}`;
    statementMap.set(key, row);
  });

  const ffCostMap = new Map();
  ffCostProcessed.forEach(row => {
    const key = `${row.OrderName}|${row.StoreID}`;
    ffCostMap.set(key, row);
  });

  const orderMap = new Map();
  orderProcessed.forEach(row => {
    const key = `${row.OrderID}|${row.StoreID}`;
    orderMap.set(key, row);
  });

  // Gộp dữ liệu và tính profit
  const result = [];
  statementMap.forEach((statementRow, key) => {
    const [orderId, storeId] = key.split("|");
    const ffCostRow = ffCostMap.get(key);
    const orderRow = orderMap.get(key);

    // Nếu không có ffCostRow hoặc orderRow, bỏ qua
    if (!ffCostRow || !orderRow) {
      console.warn(`Không tìm thấy dữ liệu khớp cho OrderID: ${orderId}, StoreID: ${storeId}`);
      return;
    }

    const profit = statementRow.Revenue - ffCostRow.Cost;

    result.push({
      OrderID: orderId,
      StoreID: storeId,
      Date: statementRow.Date,
      Revenue: statementRow.Revenue,
      Cost: ffCostRow.Cost,
      Profit: Number(profit.toFixed(2)),
      DesignerID: normalizeId(orderRow.DesignerID), // Chuẩn hóa ID
      RAndDID: normalizeId(orderRow.RAndDID),       // Chuẩn hóa ID
      Type: statementRow.Type,
      SKU: orderRow.SKU
    });
  });

  console.log(`Processed ${result.length} rows with profit calculation for month: ${month}, year: ${year}`);
  return result;
}

// Hàm tính KPI cho Etsy
function calculateKPI(statementData, ffCostData, orderData, month, year) {
  // Gọi calculateEtsyProfit để lấy dữ liệu gộp
  const profitData = calculateEtsyProfit(statementData, ffCostData, orderData, month, year);

  // Log dữ liệu đầu vào để kiểm tra
  console.log("profitData:", JSON.stringify(profitData, null, 2));

  // Tạo object để tổng hợp Profit cho DesignerID và RAndDID
  const designerProfitTotal = {};
  const rdProfitTotal = {};

  // Gán Profit từ profitData cho DesignerID và RAndDID
  profitData.forEach(row => {
    const { DesignerID, RAndDID, Profit } = row;

    // Làm tròn profit
    const roundedProfit = Number(Profit.toFixed(2));

    console.log(`Processing Order: OrderID=${row.OrderID}, DesignerID=${DesignerID}, RAndDID=${RAndDID}, Profit=${roundedProfit}`);

    // Xử lý DesignerID
    if (DesignerID) {
      designerProfitTotal[DesignerID] = Number(
        ((designerProfitTotal[DesignerID] || 0) + roundedProfit).toFixed(2)
      );
    } else {
      console.log(`Skipped DesignerID: ${DesignerID} (invalid)`);
    }

    // Xử lý RAndDID
    if (RAndDID) {
      rdProfitTotal[RAndDID] = Number(
        ((rdProfitTotal[RAndDID] || 0) + roundedProfit).toFixed(2)
      );
    } else {
      console.log(`Skipped RAndDID: ${RAndDID} (invalid)`);
    }
  });

  // Tính tổng profit cho log kiểm tra
  const totalDesignerProfit = Object.values(designerProfitTotal).reduce(
    (sum, profit) => sum + profit,
    0
  );
  const totalRDProfit = Object.values(rdProfitTotal).reduce(
    (sum, profit) => sum + profit,
    0
  );
  const totalOrderProfit = profitData.reduce(
    (sum, row) => sum + Number(row.Profit.toFixed(2)),
    0
  );

  console.log(`Calculated KPI for month: ${month}, year: ${year}`);
  console.log("Designer Profit Total:", JSON.stringify(designerProfitTotal, null, 2));
  console.log("R&D Profit Total:", JSON.stringify(rdProfitTotal, null, 2));
  console.log("Total Designer Profit:", Number(totalDesignerProfit.toFixed(2)));
  console.log("Total R&D Profit:", Number(totalRDProfit.toFixed(2)));
  console.log("Total Order Profit:", Number(totalOrderProfit.toFixed(2)));

  return {
    designerProfit: designerProfitTotal, // { "XT": 1000.00, "YZ": 2000.00 }
    rdProfit: rdProfitTotal             // { "MK": 1500.00, "AB": 2500.00 }
  };
}

module.exports = { processEtsyStatement, processEtsyFFCost, processEtsyOrder, calculateEtsyProfit, calculateKPI };