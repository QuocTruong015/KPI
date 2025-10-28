const { excelDateToJSDate } = require("../utils/excelUtils");
const XLSX = require('xlsx');

// Hàm validate row (sử dụng chung cho processEtsyStatement)
function validateRow(row) {
  const requiredFields = ["Date", "Type", "Order ID (sale, refund)"];
  const missingFields = requiredFields.filter((field) => !row[field] || String(row[field]).trim() === "");
  return missingFields.length === 0 ? null : `Thiếu cột: ${missingFields.join(", ")}`;
}

function processMerchOrder(data) {
  if (!Array.isArray(data) || data.length === 0) {
    throw new Error("Dữ liệu Excel rỗng hoặc không hợp lệ");
  }

  // Tạo một Map để tổng hợp Profit theo cặp OrderID và StoreID
  const profitMap = new Map();

  // Xử lý dữ liệu và tổng hợp Profit
  data.forEach(row => {
    const orderID = row["ASIN"] ? String(row["ASIN"]).trim() : "Unknown";
    const storeID = row["Store ID"] ? String(row["Store ID"]).trim() : "Unknown";
    const profit = row["Royalties"] != null ? parseFloat(row["Royalties"]) || 0 : 0;
    const date = row["Date"] ? excelDateToJSDate(row["Date"]) : null;

    // Tạo key duy nhất cho cặp OrderID và StoreID
    const key = `${orderID}|${storeID}`;

    if (orderID !== "Unknown" && storeID !== "Unknown") {
      const currentEntry = profitMap.get(key) || { Date: date, OrderID: orderID, StoreID: storeID, Profit: 0 };
      currentEntry.Profit += profit;
      profitMap.set(key, currentEntry);
    }
  });

  // Chuyển Map thành mảng kết quả
  const result = Array.from(profitMap.values());

  // Log số bản ghi đã xử lý
  console.log(`Processed ${result.length} unique OrderID-StoreID pairs with aggregated profit`);

  return result;
}

function processMerchSku(data) {
  if (!Array.isArray(data) || data.length === 0) {
    throw new Error("Dữ liệu Excel rỗng hoặc không hợp lệ");
  }
  const result = data.map((row, index) => {
    let sku = row["SKU"]?.trim() || "Unknown";
    let designerId = "Unknown";
    let rAndDId = "Unknown";
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
      Date: row["Created Date"] ? excelDateToJSDate(row["Created Date"]) : null,
      SKU: sku,
      OrderID: row["ASIN"]?.trim() || "Unknown",
      ProductStatus: row["Product Status"]?.trim() || "Unknown",
      DesignerID: designerId,
      RAndDID: rAndDId,
    };
  });

  // console.log(`Processed ${result.length} rows for Etsy Order`);
  return result;
}

function assignProfitToDesignerAndRDMerch(orderData, skuData) {
  // Gọi hai hàm xử lý dữ liệu
  const orderResult = processMerchOrder(orderData);
  const skuResult = processMerchSku(skuData);

  // Log dữ liệu đầu vào để kiểm tra
  console.log("orderResult:", JSON.stringify(orderResult, null, 2));
  console.log("skuResult:", JSON.stringify(skuResult, null, 2));

  // Tạo object để lưu trữ Profit theo OrderID và StoreID (hoặc chỉ OrderID nếu StoreID không có)
  const profitMap = {};
  orderResult.forEach(row => {
    const key = row.StoreID ? `${row.OrderID}|${row.StoreID}` : row.OrderID;
    profitMap[key] = Number(row.Profit.toFixed(2));
    console.log(`ProfitMap[${key}]:`, profitMap[key]); // Log profitMap
  });

  // Tạo object để tổng hợp Profit cho DesignerID và RAndDID
  const designerProfitDetails = {}; // { DesignerID: [{ OrderID, profit }] }
  const rdProfitDetails = {}; // { RAndDID: [{ OrderID, profit }] }
  const designerProfitTotal = {}; // { DesignerID: totalProfit }
  const rdProfitTotal = {}; // { RAndDID: totalProfit }
  const processedOrdersByDesigner = {};
  const processedOrdersByRD = {};

  // Gán Profit từ orderResult cho DesignerID và RAndDID từ skuResult
  skuResult.forEach(sku => {
    const orderKey = sku.StoreID ? `${sku.OrderID}|${sku.StoreID}` : sku.OrderID;
    const profit = profitMap[orderKey] || 0;

    const { DesignerID, RAndDID, OrderID } = sku;

    console.log(`Processing SKU: OrderID=${OrderID}, StoreID=${sku.StoreID || 'N/A'}, DesignerID=${DesignerID}, RAndDID=${RAndDID}, Profit=${profit}`);

    // Xử lý DesignerID
    if (DesignerID && DesignerID !== "Unknown") {
      const designerOrderKey = `${DesignerID}|${orderKey}`;
      if (!processedOrdersByDesigner[designerOrderKey]) {
        if (!designerProfitDetails[DesignerID]) {
          designerProfitDetails[DesignerID] = [];
        }
        designerProfitDetails[DesignerID].push({
          OrderID,
          profit: Number(profit.toFixed(2))
        });

        designerProfitTotal[DesignerID] = Number(
          ((designerProfitTotal[DesignerID] || 0) + profit).toFixed(2)
        );

        processedOrdersByDesigner[designerOrderKey] = true;
      }
    } else {
      console.log(`Skipped DesignerID: ${DesignerID} (invalid or Unknown)`);
    }

    // Xử lý RAndDID
    if (RAndDID && RAndDID !== "Unknown") {
      const rdOrderKey = `${RAndDID}|${orderKey}`;
      if (!processedOrdersByRD[rdOrderKey]) {
        if (!rdProfitDetails[RAndDID]) {
          rdProfitDetails[RAndDID] = [];
        }
        rdProfitDetails[RAndDID].push({
          OrderID,
          profit: Number(profit.toFixed(2))
        });

        rdProfitTotal[RAndDID] = Number(
          ((rdProfitTotal[RAndDID] || 0) + profit).toFixed(2)
        );

        processedOrdersByRD[rdOrderKey] = true;
      }
    } else {
      console.log(`Skipped RAndDID: ${RAndDID} (invalid or Unknown)`);
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
  const totalOrderProfit = Object.values(profitMap).reduce(
    (sum, profit) => sum + profit,
    0
  );

  console.log("Designer Profit Details:", JSON.stringify(designerProfitDetails, null, 2));
  console.log("R&D Profit Details:", JSON.stringify(rdProfitDetails, null, 2));
  console.log("Designer Profit Total:", JSON.stringify(designerProfitTotal, null, 2));
  console.log("R&D Profit Total:", JSON.stringify(rdProfitTotal, null, 2));
  console.log("Total Designer Profit:", Number(totalDesignerProfit.toFixed(2)));
  console.log("Total R&D Profit:", Number(totalRDProfit.toFixed(2)));
  console.log("Total Order Profit:", Number(totalOrderProfit.toFixed(2)));

  return {
    totalRecords: orderResult.length,
    designerProfit: designerProfitTotal, // { "XT": 20.00, "YZ": 30.00 }
    rdProfit: rdProfitTotal, // { "R1": 15.00, "R2": 25.00 }
    profitDetails: {
      designer: designerProfitDetails, // { "XT": [{ OrderID, profit }, ...] }
      rd: rdProfitDetails // { "R1": [{ OrderID, profit }, ...] }
    },
    profitData: orderResult // Lưu dữ liệu gốc để kiểm tra
  };
}

module.exports = { processMerchOrder, processMerchSku, assignProfitToDesignerAndRDMerch };