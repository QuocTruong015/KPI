const { excelDateToJSDate } = require("../utils/excelUtils");

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
      Supplier: row["Supplier"]?.trim() || "Unknown", // Thêm cột Supplier từ Excel
    };
  });

  console.log(`Processed ${result.length} rows for Etsy FFCost`);
  return result;
}

function processEtsyOrder(data) {
  if (!Array.isArray(data) || data.length === 0) {
    throw new Error("Dữ liệu Excel rỗng hoặc không hợp lệ");
  }
  const result = data.map((row, index) => {
    return {
      OrderID: String(row["Order ID"] || "").trim(),
      BuyerEmail: String(row["Buyer Email"] || "").trim(),
      OrderDate: excelDateToJSDate(row["Order Date"]),
      TotalAmount: parseFloat(row["Total Amount"]) || 0,
      Currency: String(row["Currency"] || "").trim(),
      Status: String(row["Status"] || "").trim(),
      StoreID: String(row["Store ID"] || "").trim(),
    };
  });

  console.log(`Processed ${result.length} rows for Etsy Order`);
  return result;
}

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
        designerId = parts[0] || "Unknown"; // XT
        rAndDId = parts[1] || "Unknown";    // MK
      } else {
        console.warn(`Row ${index + 2}: SKU không đúng định dạng (${sku})`);
      }
    }

    return {
      SaleDate: saleDate || null, // Giữ null nếu không hợp lệ
      OrderID: row["Order ID"],
      SKU: sku,
      StoreID: row["Store ID"]?.trim() || "Unknown",
      DesignerID: designerId,
      RAndDID: rAndDId,
    };
  });

  console.log(`Processed ${result.length} rows for Sale Data`);
  return result;
}

module.exports = { processEtsyStatement, processEtsyFFCost, processEtsyOrder };