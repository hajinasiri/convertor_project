var readline = require('readline-sync');

var scrivModule = require('./mainModules.js')
// var f = "/Users/shahab/lighthouse/scriv/render3/render0.3.scriv";
var scrivPath = readline.question("Enter the scriv file address:");
// var excelPath = readline.question("Enter the excel file address:");
console.log(scrivPath);
var result = scrivModule.main(scrivPath,'No');
console.log(result);
