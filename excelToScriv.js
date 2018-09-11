
var xlsx = require('node-xlsx');
var fs = require('fs');

var f = "/Users/shahab/lighthouse/scriv/render3/render0.3.xlsx";



var obj = xlsx.parse(f); // parses a file
buildCustomMetaDataSettings(obj[0].data[0]);
var text = fs.readFileSync('./template.scriv/template.scrivx').toString('utf-8');

function buildCustomMetaDataSettings(excel){

}