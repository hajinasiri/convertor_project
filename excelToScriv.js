
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
 }else{

  main();
 }
 console.log('done!');
});

function main(){
  var excel = xlsx.parse(f); // parses the excel file
  var text = fs.readFileSync('./template.scriv/template.scrivx').toString('utf-8'); //reads the template.scrivx file and puts it in text
  fs.unlinkSync(scriv + '/template.scrivx');//deletes the template.scrivx file from the destination
  var keywords = getKeywords(excel);
  text = text.replace('<Keywords></Keywords>', buildKeywords(excel,keywords));
  buildMap(excel,keywords);
  text = text.replace('<CustomMetaDataSettings></CustomMetaDataSettings>' , buildCustomMetaDataSettings(excel)); //builds CustomMetaDataSettings part and puts it in the text
  fs.writeFile(scrivx, text, function(err) {//writes the text to scrivx file
      if(err) {
          return console.log(err);
      }
      console.log("The file was saved!");
  });

}



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

function getKeywords (excel){ //makes an object of all the classes and returns the object
  var keywords = {};
  var classes;
  var rows = excel[0].data;//This lines puts all the rows in "rows" variable
  rows.forEach(function(row,index){//goes through each row
    if(index){//skips the row 0 which is the row that contains column titles
      classes = row[8];
      classes = classes.split(' ');
      classes.forEach(function(element){
        if(element && !(element in keywords)){
          keywords[element] = index;
        }
      })
    }
  });
  return keywords
}

function buildKeywords(excel,keywords){
  var str = '';
  for (var key in keywords) {
    str += '     <Keyword ID="' + keywords[key] + '">\n        <Title>' + key + '</Title>\n       <Color>0.614707 0.655123 1.0</Color></Keyword>\n'
  }
  str = '<Keywords>\n' + str + '    </Keywords>'
  return str
}

function buildMap(excel,keywords){
  var mapStr = '';
  var rows = excel[0].data;
  rows.forEach(function(row,index){
    buildBinderItem(row,rows,index);
  })
}

function buildBinderItem(row,rows,index){
  var binderItem = '<BinderItem UUID="' + index;
  // fs.mkdirSync(scriv+'/files/data/' + index);
}








