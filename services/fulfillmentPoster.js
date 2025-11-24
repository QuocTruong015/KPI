const { excelDateToJSDate } = require("../utils/excelUtils");

function processFulfillmentPosterCost  (data1, data2, data3, data4, data5, data6, data7, data8, data9, data10, data11, month, year) {
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

    const filtered6 = data6.filter((row) => {
        const type = row["Type"];
        if (type === "Ship By Seller _ UK") {
            costPosterUKSeller += parseFloat(row["Buying Label"]) || 0;
        } else if (type === "Ship By Tiktok _ UK") {
            costPosterUSTiktok += parseFloat(row["Buying Label"]) || 0;
        } 
    });

    let costGlonluxPoster = 0;
    let costCanvas = 0;

    const filtered7 = data7.filter((row) => {
        const type = row["Type"];
        if (type === "Gonlux Poster") {
            costGlonluxPoster += parseFloat(row["Base Cost"]) || 0;
        } else if (type === "Canvas") {
            costCanvas += parseFloat(row["Customer"]) || 0;
        }
    });

    let costMangoPoster = 0;
    const filtered8 = data8.filter((row) => {
        const type = row["Type"];
        if (type === "Mango Poster") {
            costMangoPoster += parseFloat(row["Date Created"]) || 0;
        }
    });

    let costPhonecase = 0;
    const filtered9 = data9.filter((row) => {
        costPhonecase += parseFloat(row["grand_total"]) || 0;
    });

    console.log("Filtered Data 10:", data10);

    let revPosterTiktok = 0;
    let revPosterSeller = 0;
    const filtered10 = data10.filter((row) => {
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

    console.log("revPosterUK:", revPosterUK);
    console.log("revCanvas:", revCanvas);

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
        RevPhonecase: revPhonecase
    };
}

module.exports = { processFulfillmentPosterCost };