const {excelDateToJSDate} = require("../utils/excelUtils");

function processServiceStaff2(data1, data2, data3, data4, month, year) {
    const filtered1 = data1.filter(row => {
      const date = excelDateToJSDate(row['Date']);
      if (!date) 
        return false;
      const dMonths = date.getFullYear() * 12 + date.getMonth();
      const target = year * 12 + (month - 1);

      return dMonths >= target - 2 && dMonths <= target;
    });

    const filtered2 = data2.filter((row) => {
        const date = excelDateToJSDate(row['Date']);
      if (!date) 
        return false;
      const dMonths = date.getFullYear() * 12 + date.getMonth();
      const target = year * 12 + (month - 1);

      return dMonths >= target - 2 && dMonths <= target;
    });

    const filtered3 = data3.filter((row) => {
        const date = excelDateToJSDate(row['Date']);
      if (!date) return false;
      const dMonths = date.getFullYear() * 12 + date.getMonth();
      const target = year * 12 + (month - 1);

      return dMonths >= target - 2 && dMonths <= target;
    });

    const filtered4 = data4.filter((row) => {
        const date = excelDateToJSDate(row['Month']);
      if (!date) return false;
      const dMonths = date.getFullYear() * 12 + date.getMonth();
      const target = year * 12 + (month - 1);

      return dMonths >= target - 2 && dMonths <= target;
    });

    const filtered5 = data2.filter((row) => {
        const date = excelDateToJSDate(row['Date']);
      if (!date) 
        return false;
      const dMonths = date.getFullYear() * 12 + date.getMonth();
      const target = year * 12 + (month - 1);

      return dMonths >= target - 2 && dMonths <= target;
    });

    const filteredPCRev = data4.filter((row) => {
        const date = excelDateToJSDate(row['Month']);
        if (!date) return false;
        const dMonths = date.getFullYear() * 12 + date.getMonth();
        const target = year * 12 + (month - 1);

        return dMonths >= target - 2 && dMonths <= target;
    });

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

    filtered4.forEach((row) => {
        const salesList = row["__EMPTY_27"].split(" ").filter(s => s.trim() !== "");
        const profit = parseFloat(row["__EMPTY_25"]) || 0;

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

    filtered5.forEach((row) => {
        const type = String(row.Type_1).trim();
        if (type !== "Tracking Ảo") return;

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

    filteredPCRev.forEach((row) => {
    const salesList = row["__EMPTY_41"].split(" ").filter(s => s.trim() !== "");
        const profit = parseFloat(row["__EMPTY_39"]) || 0;

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