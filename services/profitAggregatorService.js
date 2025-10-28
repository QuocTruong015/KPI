// src/services/profitAggregatorService.js
const { readExcelSheet, excelDateToJSDate } = require("../utils/excelUtils");
const { calculateKPI } = require("./etsyService");
const { calculateAmzKPI } = require("./amzService");
const { assignProfitToDesignerAndRD } = require("./webService");
const { assignProfitToDesignerAndRDMerch } = require("./merchService");

async function getEtsyProfit(filePath, month, year) {
  const statementData = readExcelSheet(filePath, "", 11).data;
  const ffCostData = readExcelSheet(filePath, "", 12).data;
  const orderData = readExcelSheet(filePath, "", 10).data;

  if (!statementData.length || !ffCostData.length || !orderData.length) {
    throw new Error("Etsy: Thiếu sheet 10, 11, 12");
  }

  return await calculateKPI(statementData, ffCostData, orderData, month, year);
}

async function getAmazonProfit(filePath, month, year) {
  const statementData = readExcelSheet(filePath, "", 15).data;
  const ffCostData = readExcelSheet(filePath, "", 16).data;
  const orderData = readExcelSheet(filePath, "", 14).data;

  if (!statementData.length || !ffCostData.length || !orderData.length) {
    throw new Error("Amazon: Thiếu sheet 14, 15, 16");
  }

  return await calculateAmzKPI(statementData, ffCostData, orderData, month, year);
}

async function getWebProfit(filePath, month, year) {
  const orderData = readExcelSheet(filePath, "", 18).data;
  const webCostData = readExcelSheet(filePath, "", 19).data;
  const ffCostData = readExcelSheet(filePath, "", 20).data;

  if (!orderData.length || !webCostData.length || !ffCostData.length) {
    throw new Error("Web: Thiếu sheet 18, 19, 20");
  }

  const filteredOrder = orderData.filter(row => {
    const d = row["Date"] ? excelDateToJSDate(row["Date"]) : null;
    return d && d.getMonth() + 1 === month && d.getFullYear() === year;
  });

  const { designerProfit, rdProfit } = assignProfitToDesignerAndRD(filteredOrder, webCostData, ffCostData);
  return { designerProfit, rdProfit };
}

async function getMerchProfit(filePath, month, year) {
  const orderData = readExcelSheet(filePath, "", 8).data;
  const skuData = readExcelSheet(filePath, "", 9).data;

  if (!orderData.length || !skuData.length) {
    throw new Error("Merch: Thiếu sheet 8, 9");
  }

  const filteredOrder = orderData.filter(row => {
    const d = row["Date"] ? excelDateToJSDate(row["Date"]) : null;
    return d && d.getMonth() + 1 === month && d.getFullYear() === year;
  });

  const { designerProfit, rdProfit } = assignProfitToDesignerAndRDMerch(filteredOrder, skuData);
  return { designerProfit, rdProfit };
}

// aggregateService.js
function aggregateProfit(inputData) {
  const { amazon = {}, merch = {}, web = {}, etsy = {} } = inputData;

  const totalDesignerProfit = {};
  const totalRDProfit = {};
  const platformSummary = { Amazon: 0, Merch: 0, Web: 0, Etsy: 0 };

  // Helper: Cộng profit từ object { id: value }
  const addProfit = (target, source) => {
    if (!source) return;
    Object.entries(source).forEach(([id, profit]) => {
      const p = parseFloat(profit) || 0;
      target[id] = (target[id] || 0) + p;
      target[id] = Number(target[id].toFixed(2));
    });
  };

  // Helper: Tính tổng profit từng nền tảng
  const sumPlatform = (obj) => Object.values(obj).reduce((a, b) => a + (parseFloat(b) || 0), 0);

  // 1. Amazon
  if (amazon.designerProfit) {
    addProfit(totalDesignerProfit, amazon.designerProfit);
    addProfit(totalRDProfit, amazon.rdProfit);
    platformSummary.Amazon = sumPlatform(amazon.designerProfit);
  }

  // 2. Merch
  if (merch.designerProfit) {
    addProfit(totalDesignerProfit, merch.designerProfit);
    addProfit(totalRDProfit, merch.rdProfit);
    platformSummary.Merch = sumPlatform(merch.designerProfit);
  }

  // 3. Web
  if (web.designerProfit) {
    addProfit(totalDesignerProfit, web.designerProfit);
    addProfit(totalRDProfit, web.rdProfit);
    platformSummary.Web = sumPlatform(web.designerProfit);
  }

  // 4. Etsy (có thể nhiều shop → gộp trước)
  let etsyDesigner = {}, etsyRD = {};
  if (Array.isArray(etsy) && etsy.length > 0) {
    etsy.forEach(shop => {
      addProfit(etsyDesigner, shop.designerProfit);
      addProfit(etsyRD, shop.rdProfit);
    });
  } else if (etsy.designerProfit) {
    addProfit(etsyDesigner, etsy.designerProfit);
    addProfit(etsyRD, etsy.rdProfit);
  }
  addProfit(totalDesignerProfit, etsyDesigner);
  addProfit(totalRDProfit, etsyRD);
  platformSummary.Etsy = sumPlatform(etsyDesigner);

  const totalProfit = Object.values(platformSummary).reduce((a, b) => a + b, 0);

  return {
    designerProfit: totalDesignerProfit,
    rdProfit: totalRDProfit,
    platformSummary,
    totalProfit: Number(totalProfit.toFixed(2)),
    month: inputData.amazon?.month || inputData.web?.month || inputData.etsy?.[0]?.month,
    year: inputData.amazon?.year || inputData.web?.year || inputData.etsy?.[0]?.year,
  };
}

module.exports = {
  getEtsyProfit,
  getAmazonProfit,
  getWebProfit,
  getMerchProfit,
  aggregateProfit
};