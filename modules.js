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
      row = singleElement(XML,element,index,excel,row,files,uno,result);
    }

  })

  setTimeout(function(){
    workbook.write('Excel.xlsx');; //generates the excel file. Uses setTime to let async readSingleFile function inside getText function read the rtf files and add them to the excel.
  }, 300);

}



function singleElement(XML,element,index,excel,row,files,uno,result){
  var counter =[0];//counter is the address to the current target that the code looks at. The code updates counter and then uses it to get the child
  //each element position in the array reperesent the number of generation of the target, and the value reperesents the number of the child in its generation.
  var finish = false; //finish variable changes to true when all the childs have been looked at and their data has been extracted. When finish is true the loop stops
  var validation;
  var target = element; //target is the current child
  var parent = element; //the parent of the current child is stored in parent. This varible will be used to get the MetaData and pass it to the target if the target does not have the value for MetaData

  while  (finish === false ) { //This loop starts from Map. If map has a child the loop updates variable counter to the address of that child. It goes on until the loop reaches the last genereation.
    //Then goes one generation back and checks if there is any other child in that generation. If yes, updates the counter to the new address. If no, goes one more generation back.
    //The loop continues searching like this until it goes through all the generation and children. Every child the loop goes through, uses addElement function to exctract the data from the child and add them to the excel file
    validation = true;

    for(i=0; i<counter.length; i++){ //this loop uses counter ass the address to get to the child and store it in the target variable
      if(target.Children[counter[i]]){//if there is a child in that address, sets the target
        target =target.Children[counter[i]];
      }else{ // if there is not a child in that address sets validation false
        validation = false;
      }
      if(i === (counter.length - 2)){
        parent = target;
      }
    }


    if(validation){//If the validation is true and there is a child in the address, addElement is used to add the data to the excel file
      addElement(XML,target,files,counter,row,excel,uno,parent,result);
       row += 1; //goes to the next row in excel
      if(target.Children){ //if the target has children sets the counter to the first of them
        counter.push(0);
      }else{//if the target does not have children, sets counter to the next sibling of the target
        counter[counter.length-1] +=1
      }

    }else{ //if there is not a child at the addressed obtained from the counter, sets counter to one generation back and the next sibling in that generation
      counter.splice(-1,1);
      counter[counter.length-1] +=1
    }
    target = element;//resets target to the first generation ancestor to make it ready to build the next address from

    if(counter[0] == undefined){ // when the loop goes through all the children and then comes back to the first generation and there is no more sibling there, it goes one generation back and first element in counter becomes undefined. That's how the loop realizes it should end.
      finish = true;
    }
  }
  return row;
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
    excel.cell(row,uno.indexOf('outlinelevel') + 4).number(0);//sets 0 for map
    outlineLevel = 0;
  }else{
    excel.cell(row,uno.indexOf('outlinelevel') + 4).number(counter.length);//calculates and sets outlinelevel for things other than map
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