const Excel = require('exceljs')

var workbook = new Excel.Workbook();
workbook.xlsx.readFile("g.xlsx")
    .then(function () {
        var worksheet = workbook.getWorksheet("Table 1");
        worksheet.eachRow({ includeEmpty: true }, function (row, rowNumber) {
            if (typeof e === typeof [0, 0])
                for (let i = 0; i < length(row); i++) {
                    console.log(row[i])
                }
            //console.log("Row " + rowNumber + " = " + JSON.stringify(row.values));

        });
    });