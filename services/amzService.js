const { excelDateToJSDate } = require("../utils/excelUtils");
const XLSX = require('xlsx');

// Helper: Chuẩn hóa ID
function normalizeId(id) {
  if (!id || id === "Unknown" || id === "") return null;
  return id.toString().trim().toUpperCase();
}

// Hàm validate row
function validateRow(row) {
  const requiredFields = ["Date", "Transaction type", "Order ID", "Total product charges"];
  const missingFields = requiredFields.filter((field) => !row[field] || String(row[field]).trim() === "");
  return missingFields.length === 0 ? null : `Thiếu cột: ${missingFields.join(", ")}`;
}

// Hàm xử lý Amazon Transaction
function processAmzTransaction(data, month, year) {
  if (!Array.isArray(data) || data.length === 0) {
    throw new Error("Dữ liệu Excel rỗng hoặc không hợp lệ");
  }

  const result = data
    .filter((row, index) => {
      const rawDate = row["Date"];
      if (rawDate == null || rawDate === "" || rawDate === "Unknown" || rawDate === "last-updated-date") {
        console.warn(`Row ${index + 2}: Bỏ qua do Date không hợp lệ (raw value: "${rawDate}")`);
        return false;
      }

      const date = excelDateToJSDate(rawDate);
      if (!date || isNaN(date.getTime())) {
        console.warn(`Row ${index + 2}: Bỏ qua do không chuyển đổi được ngày (raw value: "${rawDate}")`);
        return false;
      }

      const isValidPeriod = date.getMonth() + 1 === month && date.getFullYear() === year;
      if (!isValidPeriod) {
        console.warn(`Row ${index + 2}: Bỏ qua do ngoài khoảng thời gian (raw: "${rawDate}", parsed: ${date.toISOString()}, month: ${month}, year: ${year})`);
        return false;
      }

      const validationError = validateRow(row);
      if (validationError) {
        console.warn(`Row ${index + 2}: ${validationError}`);
        return false;
      }
      return true;
    })
    .map((row, index) => {
      const orderId = row["Order ID"] ? String(row["Order ID"]).trim() : "Unknown";
      const storeId = row["Store ID "] ? String(row["Store ID "]).trim() : "Unknown";
      const total = row["Total (USD)"] ? String(row["Total (USD)"]).trim() : "0";

      return {
        Date: excelDateToJSDate(row["Date"]),
        StoreID: storeId,
        OrderID: orderId,
        Rev: parseFloat(total) || 0,
      };
    });

  console.log(`Processed ${result.length}/${data.length} rows for AMZ Transaction (month: ${month}, year: ${year})`);
  console.log(`Sample statementProcessed: ${JSON.stringify(result.slice(0, 2), null, 2)}`);
  return result;
}

// Hàm xử lý Amazon FFCost
function processAmzFFCost(data, month, year) {
  if (!Array.isArray(data) || data.length === 0) {
    throw new Error("Dữ liệu Excel rỗng hoặc không hợp lệ");
  }

  const result = data
    .filter((row, index) => {
      const rawDate = row["Date created"];
      if (rawDate == null || rawDate === "" || rawDate === "Unknown" || rawDate === "last-updated-date") {
        console.warn(`Row ${index + 2}: Bỏ qua do Date created không hợp lệ (raw value: "${rawDate}")`);
        return false;
      }

      const date = excelDateToJSDate(rawDate);
      if (!date || isNaN(date.getTime())) {
        console.warn(`Row ${index + 2}: Bỏ qua do không chuyển đổi được ngày (raw value: "${rawDate}")`);
        return false;
      }

      const isValidPeriod = date.getMonth() + 1 === month && date.getFullYear() === year;
      if (!isValidPeriod) {
        console.warn(`Row ${index + 2}: Bỏ qua do ngoài khoảng thời gian (raw: "${rawDate}", parsed: ${date.toISOString()}, month: ${month}, year: ${year})`);
        return false;
      }

      return true;
    })
    .map((row, index) => {
      const orderId = row["Printify ID"] ? String(row["Printify ID"]).trim() : "Unknown";
      const storeId = row["Store ID"] ? String(row["Store ID"]).trim() : "Unknown";
      const cost = row["Total cost"] ? String(row["Total cost"]).trim() : "0";

      return {
        Date: excelDateToJSDate(row["Date created"]),
        StoreID: storeId,
        OrderID: orderId,
        Cost: parseFloat(cost) || 0,
      };
    });

  console.log(`Processed ${result.length}/${data.length} rows for AMZ FFCost (month: ${month}, year: ${year})`);
  console.log(`Sample ffCostProcessed: ${JSON.stringify(result.slice(0, 2), null, 2)}`);
  return result;
}

// Hàm xử lý Amazon Order
function processAmzOrder(data, month, year) {
  if (!Array.isArray(data) || data.length === 0) {
    throw new Error("Dữ liệu Excel rỗng hoặc không hợp lệ");
  }

  if (!month || !year) {
    console.error(`Invalid month (${month}) or year (${year}) in processAmzOrder`);
    throw new Error("Month và year phải được cung cấp để lọc dữ liệu");
  }

  const result = data
    .filter((row, index) => {
      const rawDate = row["payments-date"];
      const orderId = row["order-id"];
      const sku = row["sku"]?.trim() || "";

      // Bỏ qua hàng tiêu đề hoặc dữ liệu không hợp lệ
      if (
        rawDate == null ||
        rawDate === "" ||
        rawDate === "Unknown" ||
        rawDate === "last-updated-date" ||
        orderId === "amazon-order-id" ||
        sku === "url" ||
        sku === "sku"
      ) {
        console.warn(`Row ${index + 2}: Bỏ qua do dữ liệu không hợp lệ (payments-date: "${rawDate}", order-id: "${orderId}", sku: "${sku}")`);
        return false;
      }

      const saleDate = excelDateToJSDate(rawDate);
      if (!saleDate || isNaN(saleDate.getTime())) {
        console.warn(`Row ${index + 2}: Bỏ qua do không chuyển đổi được ngày (raw value: "${rawDate}")`);
        return false;
      }

      // Lọc theo tháng và năm
      const isValidPeriod = saleDate.getMonth() + 1 === month && saleDate.getFullYear() === year;
      if (!isValidPeriod) {
        console.warn(`Row ${index + 2}: Bỏ qua do ngoài khoảng thời gian (raw: "${rawDate}", parsed: ${saleDate.toISOString()}, month: ${month}, year: ${year})`);
        return false;
      }

      return true;
    })
    .map((row, index) => {
      let designerId = "Unknown";
      let rAndDId = "Unknown";

      const columns = Object.keys(row);
      const skuIndex = columns.indexOf("sku");
      const sku = row[columns[skuIndex + 1]]?.trim() || "";
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
        Date: excelDateToJSDate(row["payments-date"]),
        OrderID: row["order-id"],
        SKU: sku,
        DesignerID: normalizeId(designerId),
        RAndDID: normalizeId(rAndDId),
      };
    });

  console.log(`Processed ${result.length}/${data.length} rows for AMZ Order (month: ${month}, year: ${year})`);
  console.log(`Sample orderProcessed: ${JSON.stringify(result.slice(0, 2), null, 2)}`);
  return result;
}

// Hàm tính profit cho Amazon
function calculateAmzProfit(statementData, ffCostData, orderData, month, year) {
  if (!month || !year) {
    console.error(`Invalid month (${month}) or year (${year}) in calculateAmzProfit`);
    throw new Error("Month và year phải được cung cấp để tính profit");
  }

  const statementProcessed = processAmzTransaction(statementData, month, year);
  const ffCostProcessed = processAmzFFCost(ffCostData, month, year);
  const orderProcessed = processAmzOrder(orderData, month, year);

  console.log(`statementProcessed (${statementProcessed.length} rows): ${JSON.stringify(statementProcessed.slice(0, 2), null, 2)}`);
  console.log(`ffCostProcessed (${ffCostProcessed.length} rows): ${JSON.stringify(ffCostProcessed.slice(0, 2), null, 2)}`);
  console.log(`orderProcessed (${orderProcessed.length} rows): ${JSON.stringify(orderProcessed.slice(0, 2), null, 2)}`);

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

  const result = [];
  orderProcessed.forEach(orderRow => {
    const key = `${orderRow.OrderID}`;
    const rev = statementMap.get(key) || 0;
    const cost = ffCostMap.get(key) || 0;

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

  console.log(`Processed ${result.length} rows with profit calculation (month: ${month}, year: ${year})`);
  console.log(`Sample profitData: ${JSON.stringify(result.slice(0, 2), null, 2)}`);
  return result;
}

// Hàm tính KPI cho Amazon
function calculateAmzKPI(statementData, ffCostData, orderData, month, year) {
  if (!month || !year) {
    console.error(`Invalid month (${month}) or year (${year}) in calculateAmzKPI`);
    throw new Error("Month và year phải được cung cấp để tính KPI");
  }

  const profitData = calculateAmzProfit(statementData, ffCostData, orderData, month, year);

  if (profitData.length === 0) {
    console.warn("No profit data generated. Check input data or OrderID matching.");
  }

  const designerProfit = profitData.reduce((acc, row) => {
    const id = row.DesignerID;
    acc[id] = (acc[id] || 0) + row.Profit;
    acc[id] = Number(acc[id].toFixed(2));
    return acc;
  }, {});

  const randProfit = profitData.reduce((acc, row) => {
    const id = row.RAndDID;
    acc[id] = (acc[id] || 0) + row.Profit;
    acc[id] = Number(acc[id].toFixed(2));
    return acc;
  }, {});

  console.log(`Calculated KPI for AMZ (month: ${month}, year: ${year})`);
  console.log(`Designer Profit:`, JSON.stringify(designerProfit, null, 2));
  console.log(`R&D Profit:`, JSON.stringify(randProfit, null, 2));
  console.log(`Total Records: ${profitData.length}`);

  return {
    totalRecords: profitData.length,
    designerProfit,
    randProfit,
    profitData,
  };
}

module.exports = { processAmzTransaction, processAmzFFCost, processAmzOrder, calculateAmzProfit, calculateAmzKPI };