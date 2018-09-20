var parseString = require('xml2js').parseString;
var fs = require("fs");
require('log-timestamp');

var modules = require('./scrivModules');
var mainModules = require('./mainModules.js');

// var f = "/Users/shahab/lighthouse/scriv/render3/GenderFinance4.7test.scriv";

var f = process.argv[2];//reads the file address from user input in terminal

mainModules.main(f,'yes');

console.log(`Watching for file changes on ${f}`);

fs.watchFile(f, (curr, prev) => {
  console.log(`${f} file Changed`);
  mainModules.main(f,'yes');
});
