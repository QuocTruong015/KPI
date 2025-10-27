const { excelDateToJSDate } = require("../utils/excelUtils");
const XLSX = require('xlsx');

// Hàm validate row (sử dụng chung cho processEtsyStatement)
function validateRow(row) {
  const requiredFields = ["Date", "Type", "Order ID (sale, refund)"];
  const missingFields = requiredFields.filter((field) => !row[field] || String(row[field]).trim() === "");
  return missingFields.length === 0 ? null : `Thiếu cột: ${missingFields.join(", ")}`;
}

function processWebIdAndRev(data) {
  if (!Array.isArray(data) || data.length === 0) {
    throw new Error("Dữ liệu Excel rỗng hoặc không hợp lệ");
  }
  const result = data.map((row, index) => {
    let sku = row["Item ID"]?.trim() || "Unknown";
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
      Date: row["Date"] ? excelDateToJSDate(row["Date"]) : null,
      Rev: row["Net"] != null ? parseFloat(row["Net"]) || 0 : 0,
      SKU: sku,
      AddressStatus: row["Address Status"]?.trim() || "Unknown",
      OrderID: row["Custom Number"] ? String(row["Custom Number"]).trim() : "Unknown",
      Status: row["Status"]?.trim() || "Unknown",
      DesignerID: designerId,
      RAndDID: rAndDId,
    };
  }).filter(row => row.AddressStatus == "Confirmed");

  // console.log(`Processed ${result.length} rows for Etsy Order`);
  return result;
}

function processWebFFCostAtWebCost(data) {
  if (!Array.isArray(data) || data.length === 0) {
    throw new Error("Dữ liệu Excel rỗng hoặc không hợp lệ");
  }
  const result = data.map((row, index) => {
    return {
      Date: row["Date created"] ? excelDateToJSDate(row["Date created"]) : null,
      OrderStatus: row["Order Status"]?.trim() || "Unknown",
      OrderID: row["Sales channel Number"] ? String(row["Sales channel Number"]).trim() : "Unknown",
      StoreID: row["Store ID"] ? String(row["Store ID"]).trim() : "Unknown",
      Cost1: row["Total cost"] != null ? parseFloat(row["Total cost"]) || 0 : 0,
    };
  });

  // console.log(`Processed ${result.length} rows for Etsy Order`);
  return result;
}

function processWebFFCostAtFFCost(data) {
  if (!Array.isArray(data) || data.length === 0) {
    throw new Error("Dữ liệu Excel rỗng hoặc không hợp lệ");
  }
  const result = data.map((row, index) => {
    return {
      OrderID: row["Single Order ID"] ? String(row["Single Order ID"]).trim() : "Unknown",
      SellerID: row["Seller"] ? String(row["Seller"]).trim() : "Unknown",
      Cost2: row["Basecost"] + row["Poly Mailer"] + row["Cost Buying Label"],
    };
  })
  .filter(row => row.SellerID === "MER");

  // console.log(`Processed ${result.length} rows for Etsy Order`);
  return result;
}

function calculateTotalCost(webCostData, ffCostData) {
  // Gọi hai hàm xử lý dữ liệu
  const webData = processWebFFCostAtWebCost(webCostData);
  const ffData = processWebFFCostAtFFCost(ffCostData);

  // Tạo một Map để lưu trữ chi phí theo OrderID
  const costMap = new Map();

  // Xử lý dữ liệu từ webCostData
  webData.forEach(row => {
    costMap.set(row.OrderID, {
      Date: row.Date,
      OrderStatus: row.OrderStatus,
      OrderID: row.OrderID,
      StoreID: row.StoreID,
      Cost1: row.Cost1,
      Cost2: 0, // Mặc định Cost2 là 0
      totalCost: row.Cost1, // Ban đầu totalCost = Cost1
    });
  });

  // Xử lý dữ liệu từ ffCostData và cập nhật Cost2
  ffData.forEach(row => {
    if (costMap.has(row.OrderID)) {
      // Nếu OrderID đã tồn tại, cập nhật Cost2 và totalCost
      const existing = costMap.get(row.OrderID);
      existing.Cost2 = row.Cost2;
      existing.totalCost = existing.Cost1 + row.Cost2;
    } else {
      // Nếu OrderID chỉ có trong ffData
      costMap.set(row.OrderID, {
        Date: null,
        OrderStatus: "Unknown",
        OrderID: row.OrderID,
        StoreID: "Unknown",
        Cost1: 0,
        Cost2: row.Cost2,
        totalCost: row.Cost2,
      });
    }
  });

  // Chuyển Map thành mảng kết quả và lọc bỏ các bản ghi có totalCost = 0
  const result = Array.from(costMap.values()).filter(row => row.totalCost !== 0);

  // Log số bản ghi đã xử lý
  console.log(`Processed ${result.length} rows with total cost`);

  return result;
}

function calculateWebProfit(orderData, webCostData, ffCostData) {
  // Gọi hàm xử lý dữ liệu từ processWebIdAndRev
  const orderResult = processWebIdAndRev(orderData);

  // Gọi hàm calculateTotalCost để tính tổng chi phí
  const costResult = calculateTotalCost(webCostData, ffCostData);

  // Tạo một Map để lưu trữ chi phí theo OrderID
  const costMap = new Map();
  costResult.forEach(row => {
    costMap.set(row.OrderID, {
      Cost1: row.Cost1,
      Cost2: row.Cost2,
      totalCost: row.totalCost,
    });
  });

  // Tính profit dựa trên OrderID từ orderResult
  const result = orderResult.map(order => {
    const costData = costMap.get(order.OrderID) || {
      Cost1: 0,
      Cost2: 0,
      totalCost: 0,
    };

    return {
      Date: order.Date,
      OrderID: order.OrderID,
      Status: order.Status,
      AddressStatus: order.AddressStatus,
      SKU: order.SKU,
      DesignerID: order.DesignerID,
      RAndDID: order.RAndDID,
      Rev: order.Rev,
      Cost1: costData.Cost1,
      Cost2: costData.Cost2,
      totalCost: costData.totalCost,
      profit: order.Rev - costData.totalCost,
    };
  });

  // Log số bản ghi đã xử lý
  console.log(`Processed ${result.length} rows with profit calculated`);

  return result;
}

function assignProfitToDesignerAndRD(orderData, webCostData, ffCostData) {
  const profitData = calculateWebProfit(orderData, webCostData, ffCostData);
  const designerProfitMap = new Map();
  const rdProfitMap = new Map();
  profitData.forEach(order => {
    const { DesignerID, RAndDID, profit } = order;

    if (DesignerID && DesignerID !== "Unknown") {
      const currentDesignerProfit = designerProfitMap.get(DesignerID) || 0;
      designerProfitMap.set(DesignerID, currentDesignerProfit + profit);
    }

    if (RAndDID && RAndDID !== "Unknown") {
      const currentRDProfit = rdProfitMap.get(RAndDID) || 0;
      rdProfitMap.set(RAndDID, currentRDProfit + profit);
    }
  });

  const designerProfit = Array.from(designerProfitMap, ([DesignerID, profit]) => ({
    DesignerID,
    profit,
  }));

  const rdProfit = Array.from(rdProfitMap, ([RAndDID, profit]) => ({
    RAndDID,
    profit,
  }));
  return {
    designerProfit,
    rdProfit,
  };
}

module.exports = { processWebIdAndRev, processWebFFCostAtWebCost, processWebFFCostAtFFCost, calculateTotalCost, calculateWebProfit, assignProfitToDesignerAndRD };