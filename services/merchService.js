const { excelDateToJSDate } = require("../utils/excelUtils");
const XLSX = require('xlsx');

// Hàm validate row (sử dụng chung cho processEtsyStatement)
function validateRow(row) {
  const requiredFields = ["Date", "Type", "Order ID (sale, refund)"];
  const missingFields = requiredFields.filter((field) => !row[field] || String(row[field]).trim() === "");
  return missingFields.length === 0 ? null : `Thiếu cột: ${missingFields.join(", ")}`;
}

function processMerchOrder(data, month, year) {
  if (!Array.isArray(data) || data.length === 0) {
    throw new Error("Dữ liệu Excel rỗng hoặc không hợp lệ");
  }

  if (!month || !year) {
    throw new Error("Cần truyền vào month và year để lọc dữ liệu");
  }

  // Map để tổng hợp Profit theo cặp OrderID và StoreID
  const profitMap = new Map();

  // Lọc và tổng hợp dữ liệu
  data.forEach((row, index) => {
    const orderID = row["ASIN"] ? String(row["ASIN"]).trim() : "Unknown";
    const storeID = row["Store ID"] ? String(row["Store ID"]).trim() : "Unknown";
    const profit = row["Royalties"] != null ? parseFloat(row["Royalties"]) || 0 : 0;

    // Chuyển đổi ngày
    const date = row["Date"] ? excelDateToJSDate(row["Date"]) : null;

    // 🔎 Lọc theo tháng và năm (nếu có cột Date hợp lệ)
    if (date instanceof Date && !isNaN(date)) {
      const dataMonth = date.getMonth() + 1;
      const dataYear = date.getFullYear();

      if (dataMonth !== month || dataYear !== year) {
        // Bỏ qua dòng không nằm trong tháng-năm được chọn
        return;
      }
    } else {
      console.warn(`Row ${index + 2}: Ngày không hợp lệ (${row["Date"]})`);
      return;
    }

    // Tạo key duy nhất
    const key = `${orderID}|${storeID}`;

    // Gộp profit theo OrderID + StoreID
    if (orderID !== "Unknown" && storeID !== "Unknown") {
      const currentEntry = profitMap.get(key) || {
        Date: date,
        OrderID: orderID,
        StoreID: storeID,
        Profit: 0,
      };
      currentEntry.Profit += profit;
      profitMap.set(key, currentEntry);
    }
  });

  // Kết quả cuối cùng
  const result = Array.from(profitMap.values());

  console.log(`Processed ${result.length} unique OrderID-StoreID pairs for ${month}/${year}`);
  return result;
}

function processMerchSku(data, month, year) {
  if (!Array.isArray(data) || data.length === 0) {
    throw new Error("❌ Dữ liệu Excel rỗng hoặc không hợp lệ");
  }

  if (!month || !year) {
    throw new Error("❌ Cần truyền vào 'month' và 'year' để lọc dữ liệu");
  }

  const result = [];

  data.forEach((row, index) => {
    // === Lấy SKU và tách mã Designer / R&D ===
    const rawSku = row["SKU"] ? String(row["SKU"]).trim() : "";
    const sku = rawSku || "Unknown";

    let designerId = "Unknown";
    let rAndDId = "Unknown";

    if (rawSku) {
      const parts = rawSku.split("-");
      if (parts.length >= 2) {
        designerId = parts[0].trim() || "Unknown";
        rAndDId = parts[1].trim() || "Unknown";
      } else {
        console.warn(`⚠️ Row ${index + 2}: SKU không đúng định dạng (${rawSku})`);
      }
    }

    // === Xử lý ngày tạo ===
    const dateValue = row["Last Modified Date"];
    const date = dateValue ? excelDateToJSDate(dateValue) : null;

    if (!(date instanceof Date) || isNaN(date)) {
      console.warn(`⚠️ Row ${index + 2}: Ngày không hợp lệ (${row["Last Modified Date"]})`);
      return; // bỏ qua dòng này
    }

    // === Lọc theo tháng / năm ===
    const dataMonth = date.getMonth() + 1;
    const dataYear = date.getFullYear();

    if (dataMonth !== month || dataYear !== year) {
      return; // không thuộc tháng-năm cần lọc
    }

    // === Thêm dòng hợp lệ vào kết quả ===
    result.push({
      Date: date,
      SKU: sku,
      OrderID: row["ASIN"] ? String(row["ASIN"]).trim() : "Unknown",
      StoreID: row["Store ID"] ? String(row["Store ID"]).trim() : "Unknown",
      ProductStatus: row["Product Status"] ? String(row["Product Status"]).trim() : "Unknown",
      DesignerID: designerId,
      RAndDID: rAndDId,
    });
  });

  console.log(`✅ Đã xử lý ${result.length} dòng SKU hợp lệ cho tháng ${month}/${year}`);
  return result;
}
function assignProfitToDesignerAndRDMerch(orderData, skuData, month, year) {
  // Kiểm tra đầu vào
  if (!Array.isArray(orderData) || !orderData) {
    throw new Error("❌ Dữ liệu order rỗng hoặc không hợp lệ");
  }
  if (!Array.isArray(skuData) || !skuData) {
    throw new Error("❌ Dữ liệu SKU rỗng hoặc không hợp lệ");
  }
  if (!month || !year) {
    throw new Error("❌ Cần truyền vào 'month' và 'year' để lọc dữ liệu");
  }

  // Xử lý dữ liệu từ processMerchOrder và processMerchSku
  const orders = processMerchOrder(orderData, month, year);
  const skus = processMerchSku(skuData, month, year);

  // Map để nhóm SKU theo OrderID
  const skuMap = new Map();
  skus.forEach((sku) => {
    const key = sku.OrderID;
    if (!skuMap.has(key)) {
      skuMap.set(key, []);
    }
    skuMap.get(key).push(sku);
  });

  // Object để tổng hợp profit theo DesignerID và RAndDID
  const designerProfit = {};
  const rdProfit = {};

  // Duyệt qua các đơn hàng
  orders.forEach((order, index) => {
    const key = order.OrderID;
    const matchingSkus = skuMap.get(key) || [];

    if (matchingSkus.length === 0) {
      console.warn(`⚠️ Order ${index + 1}: Không tìm thấy SKU cho OrderID=${order.OrderID}`);
      return;
    }
    const profitPerSku = order.Profit / matchingSkus.length;

    matchingSkus.forEach((sku) => {
      if (sku.DesignerID === "Unknown" || sku.RAndDID === "Unknown") {
        console.warn(
          `⚠️ SKU ${sku.SKU}: DesignerID=${sku.DesignerID}, RAndDID=${sku.RAndDID} không hợp lệ, bỏ qua`
        );
        return;
      }

      // Gán profit cho DesignerID
      designerProfit[sku.DesignerID] = (designerProfit[sku.DesignerID] || 0) + profitPerSku;

      // Gán profit cho RAndDID
      rdProfit[sku.RAndDID] = (rdProfit[sku.RAndDID] || 0) + profitPerSku;
    });

    console.log(
      `Skipped ${matchingSkus.filter((sku) => sku.DesignerID === "Unknown" || sku.RAndDID === "Unknown").length} SKUs due to invalid DesignerID or RAndDID`
    );
  });

  // Làm tròn profit đến 2 chữ số thập phân
  Object.keys(designerProfit).forEach((key) => {
    designerProfit[key] = Number(designerProfit[key].toFixed(2));
  });
  Object.keys(rdProfit).forEach((key) => {
    rdProfit[key] = Number(rdProfit[key].toFixed(2));
  });

  return { designerProfit, rdProfit };
}
module.exports = { processMerchOrder, processMerchSku, assignProfitToDesignerAndRDMerch };