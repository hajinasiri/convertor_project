var parseString = require('xml2js').parseString;
var fs = require("fs");







// var text = fs.readFileSync(".../scriv/Archive/workflow5.scriv/workflow5.scrivx").toString('utf-8');
var f = "/Users/shahab/lighthouse/scriv/render2/render0.2.scriv";
var n = f.lastIndexOf('/');
var res = f.substr(n, f.length);
f = f + '/' +res+'x';
var text = fs.readFileSync(f).toString('utf-8');


parseString(text, function (err, result) {
    console.dir(result.ScrivenerProject.Binder[0]);
});