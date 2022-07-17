const Excel = require('exceljs');
const fs = require('fs')

const cho3ab = { by: "hazem-elyes" }
var workbook = new Excel.Workbook();
var code = 0
var fac = ""
var spec = ""
var section = ""
var changed = false
var predCode =""
var predSpec= ""
var predFac=""
var predSection=""
workbook.xlsx.readFile("g.xlsx")
    .then(function () {
        var worksheet = workbook.getWorksheet("Table 1");
        worksheet.eachRow({ includeEmpty: false }, function (row, rowNumber) {

            const rowValues = [];
            for (let value of row.values) {
                if (typeof value === "object") {
                    if (value !== null && typeof value.richText !== "undefined") {
                        rowValues.push(value.richText[0].text)
                    }
                }
                else if (typeof value !== "undefined" && value !== null && value != "undefined") {
                    rowValues.push(value);
                }

            }

            values = rowValues.filter(function (item, pos, self) {
                return self.indexOf(item) == pos;
            })
            console.log(values)
            
            /*if(values.indexOf('ايرولاكبلا عون') >= 0){
                console.log(values)
            }*/



            if (values[0] === "-" || typeof values[0] === "number") {
                if (typeof values[3] === "number") {
                    code = values[3]
                }

                if(code === predCode){
                    spec = ""
                    fac = ""
                }
                if(values.length === 7){
                    if(predFac !== values[5]){
                        fac += values[5]
                        predFac = values[5]
                    }
                    if(predSpec !== values[4]){
                        spec += values[4]
                        predSpec = values[4]
                    }
                    if(predSection !== values[6]){
                        section += values[6]
                        predSection = values[6]
                    }
                }

                cho3ab[`${code}`]["fac"] = fac


                if (typeof values[3] !== "number") {
                    if (!cho3ab[`${code}`])
                        cho3ab[`${code}`] = {}
                    //console.log(" . " + values[2]+" " + `${code}`)
                    cho3ab[`${code}`][`${values[2]}`] = { score: values[0], formule: values[1] }
                }
                else {
                    if (!cho3ab[`${code}`])
                        cho3ab[`${code}`] = {}
                    cho3ab[`${code}`][`${values[2].split("").reverse().join("")}`] = { score: values[0], formule: values[1] }
                }

                //console.log(values+"\n\n")

            }




        });
        fs.writeFile('logs.txt', "----------\n", err => {
            if (err) {
                console.error(err);
                return;
            }
        });
        fs.appendFile('logs.txt', "" + JSON.stringify(cho3ab) + "\n", err => {
            if (err) {
                console.error(err);
                return;
            }
        });
        console.log(max);

    });


/*
for (let value of row.values) {
                if (typeof value === "object") {
                    fs.writeFile('logs.txt', "----------\n", err => {
                        if (err) {
                            console.error(err);
                            return;
                        }
                    });
                    fs.appendFile('logs.txt', "" + JSON.stringify(value) + "\n", err => {
                        if (err) {
                            console.error(err);
                            return;
                        }
                    });
                }

            }
*/