
var xlsx = require('node-xlsx');
var fs = require('fs');
var ncp = require('ncp').ncp; //Module to copy folders

var f = "/Users/shahab/lighthouse/scriv/excelToScriv/render0.3.xlsx";

const scriv = f.substr(0,f.indexOf('xlsx')) + 'scriv'; //builds the destination path for scriv file
const fileName = scriv.substr(scriv.lastIndexOf('/'), scriv.length) + 'x';
const scrivx =scriv +fileName; //builds the destination path for scrivx file

ncp.limit = 16;//part of the module
ncp('./template.scriv', scriv, function (err) { //copies the template scriv file to the destination of scriv with the final name
 if (err) {
   return console.error(err);
 }
 console.log('done!');
});

fs.rename(scriv + 'template.scrivx', scrivx, function(err) {//changes the name of template.scrivx file to the final name
    if ( err ) console.log('ERROR: ' + err);
});







var excel = xlsx.parse(f); // parses the excel file
var text = fs.readFileSync('./template.scriv/template.scrivx').toString('utf-8'); //reads the text from scrivx file
buildKeywords(excel);
text = text.replace('<CustomMetaDataSettings></CustomMetaDataSettings>' , buildCustomMetaDataSettings(excel)); //builds CustomMetaDataSettings part and puts it in the text

fs.writeFile(scrivx, text, function(err) {//writes the text to scrivx file
    if(err) {
        return console.log(err);
    }

    console.log("The file was saved!");
});


function buildCustomMetaDataSettings(excel){
  var metaSetting = excel[0].data[0];
  metaSetting.splice(0,2);//removes manually added metadata
  metaSetting.splice(1,6);//removes manually added metadata
  var str = '';
  metaSetting.forEach(function(meta){
    str += '      <MetaDataField ID="' + meta + '" Type="Text" Wraps="No" Align="Left">\n       <Title>' + meta + '</Title>\n      </MetaDataField>\n';
  })
  str = '<CustomMetaDataSettings>' + str + '    </CustomMetaDataSettings>';
  return str
}

function buildKeywords (excel){
  console.log(excel[0].data[0]);

}