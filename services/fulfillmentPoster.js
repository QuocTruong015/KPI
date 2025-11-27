const { excelDateToJSDate } = require("../utils/excelUtils");

function filterByMonthYear(data, column, month, year) {
    return data.filter(row => {
        const date = excelDateToJSDate(row[column]);
        if (!date) return false;
        return date.getMonth() + 1 === month && date.getFullYear() === year;
    });
}

function processFulfillmentPosterCost  (data1, data2, data3, data4, data5, data6, data7, data8, data9, data10, data11, data12, month, year) {
    const filtered1 = data1.filter((row) => {
        const date = excelDateToJSDate(row["Date"]);
        if (!date) return false;
        return date.getMonth() + 1 === month && date.getFullYear() === year;
    });
    let totalCostSBTT = 0;
    let totalCostPolyTT = 0;
    filtered1.forEach((row) => {
        const costTT = parseFloat(row["Cost"]) || 0;
        totalCostSBTT += costTT;
        const costPolyTT = parseFloat(row["Poly Mailer"]) || 0;
        totalCostPolyTT += costPolyTT;
    });

    //Sheet2
    const filtered2 = data2.filter((row) => {
        const date = excelDateToJSDate(row["Date"]);
        if (!date) return false;
        return date.getMonth() + 1 === month && date.getFullYear() === year;
    });
    let totalCostBySeller = 0;
    let totalCostPolyBySeller = 0;
    filtered2.forEach((row) => {
        const costBySeller = parseFloat(row["Cost"]) || 0;
        totalCostBySeller += costBySeller;
        const costPolyBySeller = parseFloat(row["Poly Mailer"]) || 0;
        totalCostPolyBySeller += costPolyBySeller;
    });

    //Sheet3
    const filtered3 = data3.filter((row) => {
        const type = row["Type"];
        if (type === "UCS" || type === "Ship by TikTok") {
            return true;
        }
    });
    let refundPosterSeller = 0;
    filtered3.forEach((row) => {
        const refundBySeller = parseFloat(row["Rev"]) || 0;
        refundPosterSeller += refundBySeller;
    });


    //DAISY 
    const filtered4 = data4.filter((row) => {
        const date = excelDateToJSDate(row["Date Created"]);
        if (!date) return false;
        return date.getMonth() + 1 === month && date.getFullYear() === year;
    });

    let costPosterUSNC = 0;
    filtered4.forEach((row) => {
        const keys = Object.keys(row);
        const trackingIndex = keys.indexOf("Tracking");
        const prevColumn = keys[trackingIndex - 1];

        const costBySeller = parseFloat(row[prevColumn]) || 0;
        row["NOTE"] === "Ship by Seller" || row["NOTE"] === "Ship by Tiktok" ? costPosterUSNC += costBySeller : 0;
    });

    //UCS - Buying label
    const filtered5 = data5.filter((row) => {
        const date = excelDateToJSDate(row["Date"]);
        if (!date) return false;
        return date.getMonth() + 1 === month && date.getFullYear() === year;
    });
    let totalBuyingLabelCost = 0;
    filtered5.forEach((row) => {
        const buyingLabelCost = parseFloat(row["Số lượng đơn hàng"]) * parseFloat(row["Đơn giá"]);
        totalBuyingLabelCost += buyingLabelCost;
    });

    //FF-order
    let costPosterUKSeller = 0;
    let costPosterUSTiktok = 0;
    let filteredData6 = data6.filter((row) => {
      const date = excelDateToJSDate(row["Bulk Order ID"]);
      if (!date) return false;
      return date.getMonth() + 1 === month && date.getFullYear() === year;
    });

    filteredData6.forEach((row) => {
        const type = row["Type"];
        if (type === "Ship By Seller _ UK") {
            costPosterUKSeller += parseFloat(row["Buying Label"]) || 0;
        } else if (type === "Ship By Tiktok _ UK") {
            costPosterUSTiktok += parseFloat(row["Buying Label"]) || 0;
        }
    });

    let costGlonluxPoster = 0;
    let costCanvas = 0;

    let filteredData7 = data7.filter((row) => {
      const date = excelDateToJSDate(row["Single Order ID"]);
      if (!date) return false;
      return date.getMonth() + 1 === month && date.getFullYear() === year;
    });

    filteredData7.forEach((row) => {
        const type = row["Type"];
        if (type === "Gonlux Poster") {
            costGlonluxPoster += parseFloat(row["Base Cost"]) || 0;
        } else if (type === "Canvas") {
            costCanvas += parseFloat(row["Customer"]) || 0;
        }
    });

    let costMangoPoster = 0;

    let filteredData8 = data8.filter((row) => {
      const date = excelDateToJSDate(row["Size"]);
      if (!date) return false;
      return date.getMonth() + 1 === month && date.getFullYear() === year;
    });

    filteredData8.forEach((row) => {
        const type = row["Type"];
        if (type === "Mango Poster") {
            costMangoPoster += parseFloat(row["Date Created"]) || 0;
        }
    });

    let costPhonecase = 0;
    let filteredData9 = data9.filter((row) => {
      const date = excelDateToJSDate(row["created_at"]);
      if (!date) 
        return false;
      return date.getMonth() + 1 === month && date.getFullYear() === year;
    });

    filteredData9.forEach((row) => {
        costPhonecase += parseFloat(row["grand_total"]) || 0;
    });

    let revPosterTiktok = 0;
    let revPosterSeller = 0;

    const filteredData10 = data10.filter((row) => {
      const date = excelDateToJSDate(row["Month"]);
      if (!date) return false;
      return date.getMonth() + 1 === month && date.getFullYear() === year;
    });

    filteredData10.forEach((row) => {
        const type = row["Type"];
        if (type === "Ship by TikTok") {
            revPosterTiktok += parseFloat(row["Amount"]) || 0;
        } else {
            revPosterSeller += parseFloat(row["Amount"]) || 0;
        }  
    });

    let revPosterUK = 0;
    let revCanvas = 0;
    let revPhonecase = 0;

    const filtered11 = data11.filter((row) => {
        const date = excelDateToJSDate(row["Month"]);
        if (!date) return false;
        return date.getMonth() + 1 === month && date.getFullYear() === year;
    });

    filtered11.forEach((row) => {
        revPosterUK += parseFloat(row["__EMPTY_19"]) || 0;
        revCanvas += parseFloat(row["__EMPTY_17"]) || 0;
        revPhonecase += parseFloat(row["__EMPTY_15"]) || 0;
    });

    let profitBuyingLabel = 0;
    let profitEmptyPackage = 0;
    let profitTrackingAo = 0;

    const filteredData12 = data12.filter((row) => {
        const date = excelDateToJSDate(row["Date"]);
        if (!date) return false;
        return date.getMonth() + 1 === month && date.getFullYear() === year;
    });

    filteredData12.forEach((row) => {
        const type = row["Type_1"];
        if (type === "Buying Label ") {
            profitBuyingLabel += parseFloat(row["Profit"]) || 0;
        } else if (type === "Empty Package ") {
            profitEmptyPackage += parseFloat(row["Profit"]) || 0;
        } else if (type === "Tracking Ảo ") {
            profitTrackingAo += parseFloat(row["Profit"]) || 0;
        }
    });

    const totalCost = totalCostSBTT + totalCostPolyTT + totalCostBySeller + totalCostPolyBySeller + refundPosterSeller + costPosterUSNC + totalBuyingLabelCost + costPosterUKSeller + costPosterUSTiktok + costGlonluxPoster + costCanvas + costMangoPoster + costPhonecase;
    const totalRev = revPosterTiktok + revPosterSeller + revPosterUK + revCanvas + revPhonecase;
    const totalSS1ProfitNoScan = (totalRev - totalCost) + profitBuyingLabel + profitEmptyPackage + profitTrackingAo;

    return { 
        Month: month, 
        Year: year, 
        TotalCostSBTT: totalCostSBTT, 
        TotalCostPolyTT: totalCostPolyTT,
        TotalCostBySeller: totalCostBySeller,
        TotalCostPolyBySeller: totalCostPolyBySeller,
        RefundPosterSeller: refundPosterSeller,
        CostPosterUSNC: costPosterUSNC,
        TotalBuyingLabelCost: totalBuyingLabelCost,
        CostPosterUKSeller: costPosterUKSeller,
        CostPosterUSTiktok: costPosterUSTiktok,
        CostGlonluxPoster: costGlonluxPoster,
        CostCanvas: costCanvas,
        CostMangoPoster: costMangoPoster,
        CostPhonecase: costPhonecase,
        RevPosterTiktok: revPosterTiktok,
        RevPosterSeller: revPosterSeller,
        RevPosterUK: revPosterUK,
        RevCanvas: revCanvas,
        RevPhonecase: revPhonecase,
        ProfitBuyingLabel: profitBuyingLabel,
        ProfitEmptyPackage: profitEmptyPackage,
        ProfitTrackingAo: profitTrackingAo,
        TotalProfitSS1: totalSS1ProfitNoScan
    };
}

module.exports = { processFulfillmentPosterCost };