var excel4node = require('excel4node'); //initiates excel file

function createExcel(files,XML,name){//fetches data from XML, Uses addElement function to add the data to excel file. addElement function itself
  //uses propagate function to make child inherit MetaData from their parent
  var result =[[],[]];
  var Binder;
  XML.Binder.forEach(function(element){//to find Map element and put it in Binder varialble
    if(element.Title === 'Map'){
      Binder = [element];
    }
  })

  var workbook = new excel4node.Workbook();
  var excel = workbook.addWorksheet('Structure');//Setting the sheet name
  const uno = initialize(excel,XML);//uses initialize function to add all uno elements to the excel file. in rturn gets uno array and puts it in uno variable
  var row = 2;
  Binder.forEach(function(element,index){ //this forEach is here to go through all the elements in Binder if needed. But currently there is only map in Binder variable
    addElement(XML,element,files,[-1],row,excel,uno,{Title:'Binder'},result); //This line adds all the data of Map to excel file
    row += 1; //Then sets row number to the next row number

    if(element.Children){ //if there is children in Map, uses singleElement function to go through all the children inside map and gets the last row number from that function
      // row = singleElement(XML,element,index,excel,row,files,uno,result);
    }

  })

  setTimeout(function(){
    workbook.write('Excel.xlsx');; //generates the excel file. Uses setTime to let async readSingleFile function inside getText function read the rtf files and add them to the excel.
  }, 300);

}

function initialize(excel,XML){ //initializes the MetaData columns inside excel file
  excel.cell(1,3).string('Title');

  var hardCoded = ['id','parent','outlineNumber','outlineLevel','label', 'shortDescription', 'longDescription','classes'];

  var uno =[];

  XML.CustomMetaDataSettings.forEach(function(element){//This loop reads uno titles and makes uno array. Because id is in hardCoded array above, it wouldn't be added to uno array here. These two arrays will be merged below
    if(element.Title !== 'id'){
      uno.push(element.Title);
    }
  });
  uno = hardCoded.concat(uno);
  uno.forEach(function(element,index){//This loop puts all the uno metaData Titles from uno array into the excel file
    excel.cell(1,4+index).string(element);
  });
  uno = uno.map(a => a.toLowerCase());//Makes all uno titles lowercase to be able to search them
  return uno;
}

function addElement(XML,target,files,counter,row,excel,uno,parent,result){
  // getShort(files,excel,target,row,uno.indexOf('shortdescription') + 4,result);
  // getText(files,excel,target,row,uno.indexOf('longdescription') + 4,result);
  var outline = counter.map(a => a+1).map(String ).reduce((a, b) => a + '-' + b); //calculates outline number from counter variable
  excel.cell(row,3).string(target.Title);//sets the title column in excel

  excel.cell(row,uno.indexOf('id') + 4).string(target.Title.replace(/ /g,''));//sets the id column in excel
  excel.cell(row,uno.indexOf('label') + 4).string(target.Title);//sets the label column in excel
  excel.cell(row,uno.indexOf('outlinenumber') + 4).string(outline);// sets the outline number in excel
  var outlineLevel;
  if(counter[0] === -1){//sets the outlinelevel column
    excel.cell(row,uno.indexOf('outlinelevel') + 4).string('0');//sets 0 for map
    outlineLevel = 0;
  }else{
    excel.cell(row,uno.indexOf('outlinelevel') + 4).string(counter.length);//calculates and sets outlinelevel for things other than map
    outlineLevel = counter.length;
  }

  excel.cell(row,uno.indexOf('parent') + 4).string(parent.Title);//sets the parent column in excel
  //filling the classes column
  var Keys = XML.Keywords;
  var Keywords = target.Keywords;
  if(Keywords && !Array.isArray(Keywords)){ //If the keywords is just one propertiy not an array, puts it in an array
    Keywords = [Keywords];
  }else if(Keywords === undefined){//If Keywords is undefined sets its value as an empty array
    Keywords =[];
  }
  var classes = '';
  Keywords.forEach(function(element){//Gets the Keywords of each uno, finds the value of that key in Keywords of the Binder and adds all the keys for the uno and puts the value in classes variable
    Keys.forEach(function(key){
      if(key.ID === element){
        classes += ' ' + key.Title;
      }
    })
  });
  excel.cell(row,uno.indexOf('classes') + 4).string(classes); //sets the value of classes column in excel as classes variable value
  result[0][row - 2 ]= {title:target.Title, id:target.Title.replace(/ /g,''), label:target.Title, outlineNumber:outline, outlineLevel:outlineLevel, parent:parent.Title,classes:classes }; //putting the calculated metadata as the object in result array
  result[1][row - 2 ] = {title:target.Title, id:target.Title.replace(/ /g,''), label:target.Title, outlineNumber:outline, outlineLevel:outlineLevel, parent:parent.Title,classes:classes }; //putting the calculated metadata as the object in result array
  // propagate(XML,excel,row, target,parent,uno,counter,result);
}

module.exports =  {createExcel}