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
  const orderResult = processWebIdAndRev(orderData);
  const costResult = calculateTotalCost(webCostData, ffCostData); 

  // Tạo một Map để lưu trữ chi phí theo OrderID
  const costMap = new Map();
  costResult.forEach(row => {
    costMap.set(row.OrderID, {
      Cost1: row.Cost1,
      Cost2: row.Cost2,
      totalCost: row.totalCost,
      StoreID: row.StoreID,
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
      StoreID: costData.StoreID || "Unknown",
      profit: order.Rev - costData.totalCost,
    };
  });

  // Log số bản ghi đã xử lý
  console.log(`Processed ${result.length} rows with profit calculated`);

  return result;
}

function assignProfitToDesignerAndRD(orderData, webCostData, ffCostData, month, year) {
  // Gọi hàm xử lý dữ liệu
  const profitData = calculateWebProfit(orderData, webCostData, ffCostData);

  // Đọc dữ liệu Custom Order
  const customData = readCustomOrder(orderData, month, year);

  console.log("profitData:", JSON.stringify(profitData, null, 2));

  // Tạo object để tổng hợp Profit cho DesignerID và RAndDID
  const designerProfitDetails = {}; // { DesignerID: [{ OrderID, profit }] }
  const rdProfitDetails = {};       // { RAndDID: [{ OrderID, profit }] }
  const designerProfitTotal = {};   // { DesignerID: totalProfit }
  const rdProfitTotal = {};         // { RAndDID: totalProfit }

  // Gán Profit từ profitData cho DesignerID và RAndDID
  profitData.forEach(order => {
    const { DesignerID, RAndDID, profit, OrderID } = order;
    const roundedProfit = Number(profit.toFixed(2));

    // ✅ Kiểm tra trùng với Custom Order
    const isCustomMatch = customData.some(custom =>
      custom.OrderID === OrderID && custom.DesignerID === DesignerID
    );

    let designerProfitToAdd = roundedProfit;
    if (isCustomMatch) {
      designerProfitToAdd = roundedProfit * 2;
      console.log(`✅ Custom match found! OrderID=${OrderID}, Designer=${DesignerID}, Profit x2`);
    }

    console.log(`Processing Order: OrderID=${OrderID}, DesignerID=${DesignerID}, RAndDID=${RAndDID}, Profit=${roundedProfit}`);

    // === Designer xử lý ===
    if (DesignerID && DesignerID !== "Unknown") {
      if (!designerProfitDetails[DesignerID]) designerProfitDetails[DesignerID] = [];

      designerProfitDetails[DesignerID].push({
        OrderID,
        profit: designerProfitToAdd
      });

      designerProfitTotal[DesignerID] = Number(
        ((designerProfitTotal[DesignerID] || 0) + designerProfitToAdd).toFixed(2)
      );
    } else {
      console.log(`Skipped DesignerID: ${DesignerID} (invalid or Unknown)`);
    }

    // === R&D xử lý (giữ nguyên profit) ===
    if (RAndDID && RAndDID !== "Unknown") {
      if (!rdProfitDetails[RAndDID]) rdProfitDetails[RAndDID] = [];

      rdProfitDetails[RAndDID].push({
        OrderID,
        profit: roundedProfit
      });

      rdProfitTotal[RAndDID] = Number(
        ((rdProfitTotal[RAndDID] || 0) + roundedProfit).toFixed(2)
      );
    } else {
      console.log(`Skipped RAndDID: ${RAndDID} (invalid or Unknown)`);
    }
  });

  // === Tổng hợp kiểm tra ===
  const totalDesignerProfit = Object.values(designerProfitTotal).reduce((sum, profit) => sum + profit, 0);
  const totalRDProfit = Object.values(rdProfitTotal).reduce((sum, profit) => sum + profit, 0);
  const totalOrderProfit = profitData.reduce((sum, order) => sum + Number(order.profit.toFixed(2)), 0);

  console.log("Designer Profit Details:", JSON.stringify(designerProfitDetails, null, 2));
  console.log("R&D Profit Details:", JSON.stringify(rdProfitDetails, null, 2));
  console.log("Designer Profit Total:", JSON.stringify(designerProfitTotal, null, 2));
  console.log("R&D Profit Total:", JSON.stringify(rdProfitTotal, null, 2));
  console.log("Total Designer Profit:", Number(totalDesignerProfit.toFixed(2)));
  console.log("Total R&D Profit:", Number(totalRDProfit.toFixed(2)));
  console.log("Total Order Profit:", Number(totalOrderProfit.toFixed(2)));

  return {
    totalRecords: profitData.length,
    designerProfit: designerProfitTotal,
    rdProfit: rdProfitTotal,
    profitDetails: {
      designer: designerProfitDetails,
      rd: rdProfitDetails
    },
    profitData
  };
}

function calculateProfitByStoreID_WEB(orderData, webCostData, ffCostData) {
  const profitData = calculateWebProfit(orderData, webCostData, ffCostData);

  if (!Array.isArray(profitData) || profitData.length === 0) {
    console.warn("Không có dữ liệu profit để tổng hợp (Web)");
    return [];
  }

  const storeMap = new Map();

  profitData.forEach(row => {
    let storeId = String(row.StoreID || "").trim();
    if (!storeId || storeId === "Unknown" || storeId === "null") {
      storeId = "UNKNOWN";
    }

    const profit = row.profit || 0;

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

  const result = Array.from(storeMap, ([StoreID, data]) => ({
    StoreID,
    TotalProfit: Number(data.TotalProfit.toFixed(2)),
    OrderCount: data.OrderCount,
    profitData: data.profitData
  }));

  result.sort((a, b) => b.TotalProfit - a.TotalProfit);

  console.log(`Website: Tổng hợp ${result.length} StoreID`);
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

module.exports = { processWebIdAndRev, processWebFFCostAtWebCost, processWebFFCostAtFFCost, calculateTotalCost, calculateWebProfit, assignProfitToDesignerAndRD, calculateProfitByStoreID_WEB };