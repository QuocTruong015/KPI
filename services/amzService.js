const { excelDateToJSDate } = require("../utils/excelUtils");
const XLSX = require('xlsx');

// Helper: Chuẩn hóa ID
function normalizeId(id) {
  if (!id || id === "Unknown" || id === "") return null;
  return id.toString().trim().toUpperCase();
}

// Hàm validate row
function validateRow(row) {
  const requiredFields = ["Date", "Transaction type"];
  const missingFields = requiredFields.filter((field) => !row[field] || String(row[field]).trim() === "");
  return missingFields.length === 0 ? null : `Thiếu cột: ${missingFields.join(", ")}`;
}

function processAmzTransaction(data, month, year) {
  if (!Array.isArray(data) || data.length === 0) {
    throw new Error("Dữ liệu Excel rỗng hoặc không hợp lệ");
  }

  // 1️⃣ Lọc dữ liệu hợp lệ
  const filtered = data.filter((row, index) => {
    const rawDate = row["Date"];
    if (!rawDate || rawDate === "Unknown" || rawDate === "last-updated-date") return false;

    const date = excelDateToJSDate(rawDate);
    if (!date || isNaN(date.getTime())) return false;

    const isValidPeriod = date.getMonth() + 1 === month && date.getFullYear() === year;
    if (!isValidPeriod) return false;

    const validationError = validateRow(row);
    if (validationError) return false;

    return true;
  });

  // 2️⃣ Đếm tổng số Order Payment của tất cả Store
  let totalQuantity = 0;
  filtered.forEach(row => {
    if (row["Transaction type"] === "Order Payment") {
      totalQuantity += 1;
    }
  });

  // 3️⃣ Map ra kết quả
  const result = filtered.map(row => {
    const orderId = row["Order ID"] ? String(row["Order ID"]).trim() : "Unknown";
    const storeId = row["Store ID "] ? String(row["Store ID "]).trim() : "Unknown";
    const total = row["Total (USD)"] ? String(row["Total (USD)"]).trim() : "0";

    return {
      Date: excelDateToJSDate(row["Date"]),
      StoreID: storeId,
      OrderID: orderId,
      TransactionType: row["Transaction type"],
      Rev: parseFloat(total) || 0,
      ServiceFee: ["Service Fees"].includes(row["Transaction type"])
        ? parseFloat(row["Total (USD)"]) || 0
        : 0,
      Quantity: totalQuantity, // ✅ tổng quantity của tất cả store
    };
  });

  console.log(`Tổng Quantity (Order Payment): ${totalQuantity}`);
  console.log(`Processed ${result.length}/${data.length} rows for AMZ Transaction (month: ${month}, year: ${year})`);
  console.log(`Sample:`, JSON.stringify(result.slice(0, 2), null, 2));
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

function calculateAmzProfit(statementData, ffCostData, orderData, month, year) {
  if (!month || !year) throw new Error("Month và year là bắt buộc");

  const statementProcessed = processAmzTransaction(statementData, month, year);
  const ffCostProcessed = processAmzFFCost(ffCostData, month, year);
  const orderProcessed = processAmzOrder(orderData, month, year);

  // === 1️⃣ Tính tổng ServiceFee và tổng Quantity toàn bộ ===
  let totalServiceFee = 0;
  let totalQuantity = 0;

  statementProcessed.forEach(row => {
    totalServiceFee += row.ServiceFee || 0;
    totalQuantity = row.Quantity || 0;
  });

  const feePerOrder = totalQuantity > 0 ? totalServiceFee / totalQuantity : 0;
  console.log(`🔹 Fee trung bình mỗi đơn = ${feePerOrder.toFixed(2)} USD`);

  // === 2️⃣ Tạo Map phục vụ join dữ liệu ===
  const statementMap = new Map(); // key: OrderID → { Rev, StoreID }
  statementProcessed.forEach(row => {
    const key = String(row.OrderID).trim();
    if (key && key !== "Unknown") {
      statementMap.set(key, {
        Rev: row.Rev,
        StoreID: row.StoreID,
      });
    }
  });

  const ffCostMap = new Map(); // key: OrderID → Cost
  ffCostProcessed.forEach(row => {
    const key = String(row.OrderID).trim();
    if (key && key !== "Unknown") {
      ffCostMap.set(key, row.Cost);
    }
  });

  // === 3️⃣ Ghép dữ liệu tính Profit ===
  const result = [];

  orderProcessed.forEach(orderRow => {
    const orderId = String(orderRow.OrderID).trim();
    if (!orderId || orderId === "Unknown") return;

    const stmt = statementMap.get(orderId) || { Rev: 0, StoreID: "Unknown" };
    const cost = ffCostMap.get(orderId) || 0;

    const profit = stmt.Rev - cost;

    result.push({
      OrderID: orderId,
      StoreID: stmt.StoreID,
      Date: orderRow.Date,
      Revenue: stmt.Rev,
      Cost: cost,
      Profit: Number(profit.toFixed(2)) + Number(feePerOrder.toFixed(2)),
      DesignerID: orderRow.DesignerID,
      RAndDID: orderRow.RAndDID,
      SKU: orderRow.SKU,
      Fee: Number(feePerOrder.toFixed(2)), // ✅ Gắn cùng 1 giá trị cho mọi đơn
      Quantity: totalQuantity
    });
  });

  return result;
}

// Hàm tính KPI cho Amazon
function calculateAmzKPI(statementData, ffCostData, orderData, customData, month, year) {
  if (!month || !year) {
    console.error(`Invalid month (${month}) or year (${year}) in calculateAmzKPI`);
    throw new Error("Month và year phải được cung cấp để tính KPI");
  }

  const profitData = calculateAmzProfit(statementData, ffCostData, orderData, month, year);
  const customOrderData = readCustomOrder(customData, month, year); // dùng sheet chứa custom order

  if (profitData.length === 0) {
    console.warn("No profit data generated. Check input data or OrderID matching.");
  }

  const designerProfit = {};
  const randProfit = {};

  profitData.forEach(row => {
    const { OrderID, DesignerID, RAndDID, Profit } = row;
    const roundedProfit = Number(Profit.toFixed(2));

    // Kiểm tra xem có trùng với CustomOrderData không
    const isCustomMatch = customOrderData.some(custom =>
      custom.OrderID === OrderID && custom.DesignerID === DesignerID
    );

    let designerProfitToAdd = roundedProfit;
    if (isCustomMatch) {
      designerProfitToAdd = roundedProfit * 2; // nhân đôi profit nếu trùng
      console.log(`✅ Custom match found! OrderID=${OrderID}, Designer=${DesignerID}, Profit x2`);
    }

    // === Gán cho Designer ===
    if (DesignerID) {
      designerProfit[DesignerID] = Number(
        ((designerProfit[DesignerID] || 0) + designerProfitToAdd).toFixed(2)
      );
    }

    // === Gán cho R&D (giữ nguyên profit gốc) ===
    if (RAndDID) {
      randProfit[RAndDID] = Number(
        ((randProfit[RAndDID] || 0) + roundedProfit).toFixed(2)
      );
    }
  });

  return {
    totalRecords: profitData.length,
    designerProfit,
    randProfit,
  };
}

function calculateProfitByStoreID_AMZ(statementData, ffCostData, orderData, month, year) {
  // Bước 1: Tính profit chi tiết từng đơn (đã xử lý lệch dữ liệu)
  const profitData = calculateAmzProfit(statementData, ffCostData, orderData, month, year);

  if (!Array.isArray(profitData) || profitData.length === 0) {
    console.warn("Không có dữ liệu profit để tổng hợp theo StoreID (Amazon)");
    return [];
  }

  // Bước 2: Gom nhóm theo StoreID
  const storeMap = new Map(); // StoreID → { TotalProfit, OrderCount }

  profitData.forEach(row => {
    // Chuẩn hóa StoreID
    let storeId = String(row.StoreID || "").trim();
    if (!storeId || storeId === "Unknown" || storeId === "null") {
      storeId = "UNKNOWN"; // Gom tất cả lỗi vào 1 nhóm
    }

    const profit = Number(row.Profit) || 0;

    if (storeMap.has(storeId)) {
      const curr = storeMap.get(storeId);
      storeMap.set(storeId, {
        TotalProfit: curr.TotalProfit + profit,
        OrderCount: curr.OrderCount + 1
      });
    } else {
      storeMap.set(storeId, {
        TotalProfit: profit,
        OrderCount: 1
      });
    }
  });

  // Bước 3: Chuyển sang mảng + làm tròn + sắp xếp
  const result = Array.from(storeMap, ([StoreID, data]) => ({
    StoreID,
    TotalProfit: Number(data.TotalProfit.toFixed(2)),
    OrderCount: data.OrderCount
  }));

  // Sắp xếp giảm dần theo Profit
  result.sort((a, b) => b.TotalProfit - a.TotalProfit);

  console.log(`Amazon: Tổng hợp thành công ${result.length} StoreID (tháng ${month}/${year})`);
  return result;
}

function readCustomOrder(data, month, year) {
  // const profitData = calculateEtsyProfit(statementData, ffCostData, orderData, month, year);
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

module.exports = { processAmzTransaction, processAmzFFCost, processAmzOrder, calculateAmzProfit, calculateAmzKPI, calculateProfitByStoreID_AMZ };