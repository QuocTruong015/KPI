const { excelDateToJSDate } = require("../utils/excelUtils");
const XLSX = require('xlsx');

// Hàm validate row
function validateRow(row) {
  const requiredFields = ["Date", "Transaction type", "Order ID", "Total product charges"];
  const missingFields = requiredFields.filter((field) => !row[field] || String(row[field]).trim() === "");
  return missingFields.length === 0 ? null : `Thiếu cột: ${missingFields.join(", ")}`;
}

function processAmzTransaction(data) {
  if (!Array.isArray(data) || data.length === 0) {
    throw new Error("Dữ liệu Excel rỗng hoặc không hợp lệ");
  }

  const result = data.map((row, index) => {
    const date = excelDateToJSDate(row["Date"]);
    const orderId = row["Order ID"] ? String(row["Order ID"]).trim() : "Unknown";
    const storeId = row["Store ID "] ? String(row["Store ID "]).trim() : "Unknown";
    const total = row["Total (USD)"] ? String(row["Total (USD)"]).trim() : "0";


    return {
      Date: date,
      StoreID: storeId,
      OrderID: orderId,
      Rev: parseFloat(total) || 0,
    };
  });

  console.log(`Processed ${result.length} rows for AMZ Transaction`);
  return result;
}

function processAmzFFCost(data) {
  if (!Array.isArray(data) || data.length === 0) {
    throw new Error("Dữ liệu Excel rỗng hoặc không hợp lệ");
  }

  const result = data.map((row, index) => {
    const date = excelDateToJSDate(row["Date created"]);
    const orderId = row["Printify ID"] ? String(row["Printify ID"]).trim() : "Unknown";
    const storeId = row["Store ID"] ? String(row["Store ID"]).trim() : "Unknown";
    const cost = row["Total cost"] ? String(row["Total cost"]).trim() : "0";


    return {
      Date: date,
      StoreID: storeId,
      OrderID: orderId,
      Cost: parseFloat(cost) || 0,
    };
  });

  console.log(`Processed ${result.length} rows for AMZ Transaction`);
  return result;
}

function processAmzOrder(data) {
  if (!Array.isArray(data) || data.length === 0) {
    throw new Error("Dữ liệu Excel rỗng hoặc không hợp lệ");
  }

  const result = data.map((row, index) => {
    // Xử lý Sale Date
    const saleDate = excelDateToJSDate(row["payments-date"]);
    if (!saleDate || isNaN(saleDate.getTime())) {
      console.warn(`Row ${index + 2}: Date không hợp lệ (${row["payments-date"]})`);
    }

    // Trích xuất Designer ID và R&D ID từ SKU
    let designerId = "Unknown";
    let rAndDId = "Unknown";

    const columns = Object.keys(row);
    const skuIndex = columns.indexOf("sku");
    const sku = row[columns[skuIndex + 1]]?.trim() || "";
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
      Date: saleDate || null, // Giữ null nếu không hợp lệ
      OrderID: row["order-id"],
      SKU: sku,
      DesignerID: designerId,
      RAndDID: rAndDId,
    };
  });

  // console.log(`Processed ${result.length} rows for Sale Data`);
  return result;
}

function calculateAmzProfit(statementData, ffCostData, orderData) {
  // Xử lý dữ liệu từ các hàm hiện có
  const statementProcessed = processAmzTransaction(statementData);
  const ffCostProcessed = processAmzFFCost(ffCostData);
  const orderProcessed = processAmzOrder(orderData);

  // Tạo map để tra cứu Rev và Cost theo OrderID
  const statementMap = new Map();
  statementProcessed.forEach(row => {
    const key = `${row.OrderID}`;
    statementMap.set(key, row.Rev);
  });

  const ffCostMap = new Map();
  ffCostProcessed.forEach(row => {
    const key = `${row.OrderID}`;
    ffCostMap.set(key, row.Cost);
  });

  // Gộp dữ liệu và tính profit, lấy OrderID từ orderProcessed làm chuẩn
  const result = [];
  orderProcessed.forEach(orderRow => {
    const key = `${orderRow.OrderID}`;
    const rev = statementMap.get(key) || 0;
    const cost = ffCostMap.get(key) || 0;

    // Chỉ cảnh báo nếu cả Rev và Cost đều không có, nhưng vẫn xử lý
    if (rev === 0 && cost === 0) {
      console.warn(`Không tìm thấy Rev hoặc Cost khớp cho OrderID: ${orderRow.OrderID}`);
    }

    const profit = rev - cost;

    result.push({
      OrderID: orderRow.OrderID,
      StoreID: orderRow.StoreID || "Unknown",
      Date: orderRow.Date,
      Revenue: rev,
      Cost: cost,
      Profit: Number(profit.toFixed(2)),
      DesignerID: orderRow.DesignerID,
      RAndDID: orderRow.RAndDID,
      SKU: orderRow.SKU,
    });
  });

  console.log(`Processed ${result.length} rows with profit calculation`);
  return result;
}

// Hàm tính KPI AMZ
function calculateAmzKPI(statementData, ffCostData, orderData) {
  // Gọi calculateAmzProfit để lấy dữ liệu gộp, lưu tạm vào profitData
  const profitData = calculateAmzProfit(statementData, ffCostData, orderData);

  // Tính tổng profit cho từng DesignerID
  const designerProfit = profitData.reduce((acc, row) => {
    const id = row.DesignerID;
    acc[id] = (acc[id] || 0) + row.Profit;
    acc[id] = Number(acc[id].toFixed(2));
    return acc;
  }, {});

  // Tính tổng profit cho từng RAndDID
  const randProfit = profitData.reduce((acc, row) => {
    const id = row.RAndDID;
    acc[id] = (acc[id] || 0) + row.Profit;
    acc[id] = Number(acc[id].toFixed(2));
    return acc;
  }, {});

  console.log(`Calculated KPI`);
  console.log(`Designer Profit:`, designerProfit);
  console.log(`R&D Profit:`, randProfit);

  return {
    totalRecords: profitData.length,
    designerProfit,
    randProfit,
    profitData,
  };
}

module.exports = { processAmzTransaction, processAmzFFCost, processAmzOrder, calculateAmzProfit, calculateAmzKPI };
