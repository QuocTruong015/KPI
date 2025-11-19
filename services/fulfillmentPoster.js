const { excelDateToJSDate } = require("../utils/excelUtils");

function processFulfillmentPosterCost  (data1, data2, data3, data4, month, year) {
    //Sheet1
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
    console.log(totalCostSBTT, totalCostPolyTT);

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
    console.log(totalCostBySeller, totalCostPolyBySeller);

    //Sheet3
    const filtered3 = data3.filter((row) => {
        const date = excelDateToJSDate(row["Month"]);
        if (!date) return false;
        return date.getMonth() + 1 === month && date.getFullYear() === year;
    });
    let refundPosterSeller = 0;
    filtered3.forEach((row) => {
        const refundBySeller = parseFloat(row["Rev"]) || 0;
        refundPosterSeller += refundBySeller;
    });

    //DAISY 
    console.log("Data4:", data4);
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

    return { 
        Month: month, 
        Year: year, 
        TotalCostSBTT: totalCostSBTT, 
        TotalCostPolyTT: totalCostPolyTT,
        TotalCostBySeller: totalCostBySeller,
        TotalCostPolyBySeller: totalCostPolyBySeller,
        RefundPosterSeller: refundPosterSeller,
        CostPosterUSNC: costPosterUSNC,
    };
}

module.exports = { processFulfillmentPosterCost };