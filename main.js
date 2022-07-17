const Excel = require('exceljs');
const fs = require('fs')

const cho3ab = {by:"hazem-elyes"}
var workbook = new Excel.Workbook();
var code = 0
workbook.xlsx.readFile("g.xlsx")
    .then(function () {
        var worksheet = workbook.getWorksheet("Table 1");
        worksheet.eachRow({ includeEmpty: false }, function (row, rowNumber) {
            

            const rowValues = [];
            for (let value of row.values) {
                if (typeof value === "object") {
                    if (value !== null && typeof value.richText !== "undefined"){
                        rowValues.push(value.richText[0].text)
                    }
                }
                else if (typeof value === "number"){
                    rowValues.push(value);
                }

            }

            values = rowValues.filter(function(item, pos, self) {
                return self.indexOf(item) == pos;
            })

            /*if(values.indexOf('ايرولاكبلا عون') >= 0){
                console.log(values)
            }*/

            if(values[0] === "-" || typeof values[0] === "number"){
                if(typeof values[3] === "number"){
                    code = values[3]
                }

                if(typeof values[3] !=="number"){
                    if(!cho3ab[`${code}`])
                        cho3ab[`${code}`]={}
                    console.log(" . " + values[2]+" " + `${code}`)
                    cho3ab[`${code}`][`${values[2]}`] = {score : values[0], formule : values[1]}
                }
                else {
                    if(!cho3ab[`${values[3]}`])
                        cho3ab[`${values[3]}`]={}
                    cho3ab[`${values[3]}`][`${values[2].split("").reverse().join("")}`]={score : values[0], formule : values[1]}
                }

                console.log(values+"\n\n")
                
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
        //console.log(cho3ab);

    });
function uniq(a) {
    var prims = { "boolean": {}, "number": {}, "string": {} }, objs = [];

    return a.filter(function (item) {
        var type = typeof item;
        if (type in prims)
            return prims[type].hasOwnProperty(item) ? false : (prims[type][item] = true);
        else
            return objs.indexOf(item) >= 0 ? false : objs.push(item);
    });
}



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