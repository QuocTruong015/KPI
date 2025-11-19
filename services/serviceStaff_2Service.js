const { fi } = require("date-fns/locale");
const {excelDateToJSDate} = require("../utils/excelUtils");

function processServiceStaff2(data1, data2, data3, data4, month, year) {
    const filtered1 = data1.filter((row) => {
        const date = excelDateToJSDate(row['Date']);
        if (!date) return false;
        return date.getMonth() + 1 === month && date.getFullYear() === year;
    });

    const filtered2 = data2.filter((row) => {
        const date = excelDateToJSDate(row['Date']);
        if (!date) return false;
        return date.getMonth() + 1 === month && date.getFullYear() === year;
    });

    const filtered3 = data3.filter((row) => {
        const date = excelDateToJSDate(row['Date']);
        if (!date) return false;
        return date.getMonth() + 1 === month && date.getFullYear() === year;
    });

    const filtered4 = data4.filter((row) => {
        const date = excelDateToJSDate(row['__EMPTY_22']);
        if (!date) return false;
        return date.getMonth() + 1 === month && date.getFullYear() === year;
    });
    const saleTotalTest = {};
    filtered4.forEach((row) => {
        const salesList = row["__EMPTY_27"].split(" ").filter(s => s.trim() !== "");
        const profit = parseFloat(row["__EMPTY_25"]) || 0;

        salesList.forEach((sale) => {
            if (!saleTotalTest[sale]) {
                saleTotalTest[sale] = {
                    month: month,
                    year: year,
                    sales: sale,
                    profit: 0
                };
            }
            saleTotalTest[sale].profit += profit;
        });
    });

    console.log("Sale Total Test:", saleTotalTest);

    const salesTotals = {};
    filtered1.forEach((row) => {
        const salesList = row.Sales.split(" ").filter(s => s.trim() !== "");
        
        const Rev = parseFloat(row.Rev) || 0;
        const Cost = parseFloat(row.Cost) || 0;
        const Profit = Rev - Cost;

        salesList.forEach((sale) => {
            if (!salesTotals[sale]) {
                salesTotals[sale] = {
                    month: month,
                    year: year,
                    sales: sale,    
                    profit: 0 
                };
            }
            salesTotals[sale].profit += Profit;
        });
    });

    filtered2.forEach((row) => {
        const type = String(row.Type_1).trim();
        if (type !== "Empty Package") return;

        const salesList = row.Sales.split(" ").filter(s => s.trim() !== "");

        salesList.forEach((sale) => {
            if (!salesTotals[sale]) {
                salesTotals[sale] = {
                    month: month,
                    year: year,
                    sales: sale,
                    profit: 0
                };
            }
            let profit = parseFloat(row.Profit) || 0;
            salesTotals[sale].profit += profit;
        });
    });

    filtered3.forEach((row) => {
        const salesList = row.Sales.split(" ").filter(s => s.trim() !== "");

        const profit = parseFloat(row.Profit_1) || 0;
        salesList.forEach((sale) => {
            if (!salesTotals[sale]) {
                salesTotals[sale] = {
                    month: month,
                    year: year,
                    sales: sale,
                    profit: 0
                };
            }
            salesTotals[sale].profit += profit;
        });
    });

    return Object.values(salesTotals);
}

module.exports = { processServiceStaff2 };