// var readXlsxFile = require('read-excel-file/node');

// var fs = require('fs');

 var f = "/Users/shahab/lighthouse/scriv/render3/render0.3.xlsx";
// readXlsxFile(f).then((rows) => {
// console.log('read');
// })
// // readXlsxFile(fs.createReadStream('/Users/shahab/lighthouse/scriv/render3/render0.3.xlsx')).then((rows) => {

// // })


// var text = fs.readFileSync(f).toString('utf-8');
// console.log(text)


var xlsx = require('node-xlsx');

var obj = xlsx.parse(f); // parses a file
console.log(obj[0].data[0]);