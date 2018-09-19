
var xlsx = require('node-xlsx');
var fs = require('fs');
var ncp = require('ncp').ncp; //Module to copy folders
var dateTime = require('node-datetime');//Module for getting dateTime

// var f = "/Users/shahab/lighthouse/scriv/render3/GenderFinance4.9test.xlsx";
var f = process.argv[2];//reads the file address from user input in terminal

const scriv = f.substr(0,f.indexOf('xlsx')) + 'scriv'; //builds the destination path for scriv file
const fileName = scriv.substr(scriv.lastIndexOf('/'), scriv.length) + 'x';
const scrivx =scriv +fileName; //builds the destination path for scrivx file
ncp.limit = 16;//part of the module
ncp('./template.scriv', scriv, function (err) { //copies the template scriv file to the destination of scriv with the final name
  if (err) {
    return console.error(err);
  }else{
    var text = main();
    writeFile(scrivx,text);
  }
  console.log('done!');
});

function main(){
  var excel = xlsx.parse(f); // parses the excel file
  var text = fs.readFileSync('./template.xml').toString('utf-8'); //reads the template.scrivx file and puts it in text
  var keywords = getKeywords(excel);
  text = text.replace('<Map></Map>', buildMap(excel,keywords))
  text = text.replace('<Keywords></Keywords>', buildKeywords(excel,keywords));
  text = text.replace('<CustomMetaDataSettings></CustomMetaDataSettings>' , buildCustomMetaDataSettings(excel)); //builds CustomMetaDataSettings part and puts it in the text
  text = text.replace('&','&amp;');
  return text
}

function writeFile(path,text){//This function writes the text to the path of the file
      fs.writeFile(path, text, function(err) {//writes the text to scrivx file
      if(err) {
          return console.log(err);
      }
    });
}

function buildCustomMetaDataSettings(excel){
  var metaSetting = excel[0].data[0];
  metaSetting.splice(0,2);//removes manually added metadata
  metaSetting.splice(1,6);//removes manually added metadata
  var str = '';
  metaSetting.forEach(function(meta){
    str += '\n      <MetaDataField ID="' + meta + '" Type="Text" Wraps="No" Align="Left">\n       <Title>' + meta + '</Title>\n      </MetaDataField>\n';
  })
  str = '<CustomMetaDataSettings>' + str + '    </CustomMetaDataSettings>';
  return str
}

function getKeywords (excel){ //makes an object of all the classes and returns the object
  var keywords = {};
  var classes;
  var rows = excel[0].data;//This lines puts all the rows in "rows" variable
  rows.forEach(function(row,index){//goes through each row
    if(row[8] && index ){//skips the row 0 which is the row that contains column titles
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
    if(index > 0){
      mapStr += '\n' +buildBinderItem(row,rows,index);//adding the binderItem string to it
      mapStr += '\n  <Title>' + row[1]+ '</Title>';
      mapStr += buildMetaData(row,rows);
      mapStr += buildUnoKeywords(row,keywords);
      mapStr += buildClose(row,rows,index);
    }

  })
  return mapStr
}

function buildBinderItem(row,rows,index){
  var binderItem = '';

  if(rows[index - 1][5] !=='outline' && row[5] > rows[index - 1 ][5]){
    binderItem += '\n<Children>';
  }
  for(i=0;i<row[5] + 2;i++){
    binderItem += ' '
  }
  binderItem += '\n<BinderItem UUID="' + index + '" ';
  var content = scriv+'/files/data/' + index ;//puts the address for content.rtf's folder for this row in variable content
  if (!fs.existsSync(content)) {//if the content folder does not exist
    fs.mkdirSync(content ); //creates the folder with row index as UUID for putting content.rtf and synopsis
  }
  if(row[6]){
    writeFile(content+'/synopsis.txt', row[6]);//writes the synopsis file
  }
  if(row[7]){
    writeFile(content+'/content.rtf',row[7]);//writes the content.rtf file
  }
  if(rows[index + 1] && row[5]<rows[index + 1][5]){//if outline number of this row is smaller than the next one's
    binderItem += 'Type="Folder"';
  }else{
    binderItem += 'Type="Text"';
  }
  var dt = dateTime.create();//gets date
  var formatted = dt.format('Y-m-d H:M:S');//formatting date
  var date = formatted + ' -0700';//adds needed string to it
  binderItem += ' Created="' + date + '"' + ' Modified="' + date + '">';//adds it to the binderItem
  return binderItem;
}

function buildMetaData(row,rows){
  var metaStr = '\n<MetaData>\n   <IncludeInCompile>Yes</IncludeInCompile>\n  <CustomMetaData>\n  <MetaDataItem> \n      <FieldID>id</FieldID>\n  <Value>' + row[2] + '</Value>\n   </MetaDataItem>';
  for(i=9; i<row.length;i++){//going through metadata columns
    if(row[i]){
      metaStr += '\n  <MetaDataItem> \n      <FieldID>' + rows[0][i] +'</FieldID>\n      <Value>' + row[i] + '</Value>\n   </MetaDataItem>';
    }
  }
  metaStr += '\n     </CustomMetaData>\n</MetaData>';
  metaStr +=  '\n<TextSettings>\n<TextSelection>0,0</TextSelection>\n</TextSettings>';
  return metaStr
}

function buildUnoKeywords(row,keywords){
  var keyStr = '';

  if(row[8]){
    var classes = row[8];
    classes = classes.split(' ');
    keyStr = '\n<Keywords>';
    classes.forEach(function(element){
      if(keywords[element]){
        keyStr += '\n<KeywordID>' + keywords[element] + '</KeywordID>';
      }
    });
  keyStr += '\n</Keywords>';
  }

  return keyStr
}


function buildClose(row,rows,index){
  var closeStr = '';
  var difference = 0;

  if(rows[index + 1]){
    difference = row[5]-rows[index + 1][5];

  }else{
    difference = row[5];
  }

  if(difference >-1){
    closeStr += '\n</BinderItem>';
  }

  for (i=0; i<difference; i++){
    closeStr += '    \n</Children>\n</BinderItem>';
  }
  return closeStr
}









