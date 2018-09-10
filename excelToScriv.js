
var xlsx = require('node-xlsx');
var fs = require('fs');

var f = "/Users/shahab/lighthouse/scriv/render3/render0.3.xlsx";



var obj = xlsx.parse(f); // parses a file
// console.log(obj[0].data[0]);
var scrivx = '<?xml version="1.0" encoding="UTF-8"?>\n<ScrivenerProject Identifier="8C9AF9D5-A46A-49D3-9484-B4BB7956232E" Version="2.0" Creator="SCRMAC-3.0.3-3032" Device="MacBook-Air-24791" Author="JS code" Modified="2018-09-05 15:41:51 -0700" ModID="8E331F26-C3A2-482B-A60D-4D14E796B45C">\n<Binder>' + '\n<BinderItem UUID="BFE9B371-C29C-4443-9DA1-A3F2CFA8EF19" Type="Folder" Created="2018-08-16 07:14:16 -0700" Modified="2018-08-16 07:15:06 -0700">\n<BinderItem>\n</ScrivenerProject>';


var path = f.substr(0,f.lastIndexOf('.xlsx'))+ 'haha.scriv';//creates the path for directory
fs.mkdir(path,function(){});//creates the directory
var scrivxPath = path + f.substr(f.lastIndexOf('/'),f.length);//This line and the next creates path for scrivx file
scrivxPath = scrivxPath.substr(0,scrivxPath.lastIndexOf('.xlsx')) + '.scrivx';
fs.writeFile(scrivxPath, scrivx, function(err) {//writes the animate.json file
  if(err) {
    return console.log(err);
  }
  console.log('scrivx ',"file was saved!");
});