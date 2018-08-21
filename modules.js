var excel4node = require('excel4node'); //initiates excel file
var fs = require("fs");

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

function addElement(XML,target,files,counter,row,excel,uno,parent,result){
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
    getShort(files,excel,target,row,uno.indexOf('shortdescription') + 4,result);
  // getText(files,excel,target,row,uno.indexOf('longdescription') + 4,result);

  propagate(XML,excel,row, target,parent,uno,counter,result);
}


function propagate(XML,excel,row, target,parent,uno,counter,result){

  if(target.MetaData && target.MetaData.CustomMetaData){
    var CustomMetaData = target.MetaData.CustomMetaData; //get the CustomMetaData from the child
  }else{
    var CustomMetaData = [];
  }

  if(!Array.isArray(CustomMetaData)){ // if the CustomMetaData is just one object put that object in an array to make
    //all CustomMetaDatas of type of array
    CustomMetaData = [CustomMetaData];
  }

  if(parent.MetaData && parent.MetaData.CustomMetaData){ //checks if the parent has CustomeMetaData. if yes, passes its value to parentMetaData
    var parentMetaData = parent.MetaData.CustomMetaData;
  }else{//if the parent does not have any MetaData sets parentMetaData as [] to make the code work
    var parentMetaData = [];
  }

  if(!Array.isArray(parentMetaData)){ // if the CustomMetaData is just one object put that object in an array to make
    //all CustomMetaDatas of type of array
    parentMetaData = [parentMetaData];
  }

  var unoto = ''; // to initiate the value for unoto. This variable is used to check if the child has unoto value and if yes store the value in it
  var unofrom = '';//To initiate the value for unfrom. This variable is used to chek if the child has unofrom value and if yes store the value in it
  var id = target.Title.replace(/ /g,''); //assumming that the target has no id, and making id from title. If the target has its own id, it will update id in the loop below

  uno.forEach(function(element,index){//goes through all the MetaData and checks if the child has that value or the parent and puts that vlue in excel file
    var found = false;
    CustomMetaData.forEach(function(childData){//Checks if the child has a value for it
      if( childData.FieldID === element){
        result[0][row - 2][element] = childData.Value;//putting all the properties of the uno inside the result arrays insid the row-2 element which is an array itself and inside it's first elemant
        result[1][row - 2][element] = childData.Value;
        excel.cell(row,index+ 4).string(childData.Value);
        if(element === 'unoto'){//To check if the child has a value for unoto
          unoto = childData.Value; //Then that value is stored in unoto variable
        }else if(element === 'unofrom'){//To check if the child has a vlue for unofrom
          unofrom = childData.Value; // Then that value is stored in unofrom variable
        }else if(element === 'id'){ //Stroing the value for id in id varaiable
          id = childData.Value;
        }
        found = true;
      }
    });
    if(unoto && !unofrom && id){

      excel.cell(row,uno.indexOf('unofrom') + 4).string(id); // if there is value for unoto, but no value for unofrom then the excel column value for unofrom is set as the value of id
      result[0][row - 2].id = id; //and then put that value for id in result array
      result[1][row - 2].id = id;
    }

    //making the object of inheritable properties
    const inheritable = {'classes':'',hoveraction:'',hoverfunction:'',clickaction:'',clickfunction:'',ondoubleclick:'',tooltip:'',infopane:'',onfunction:'',
      offfunction:'',openfunction:'',closefunction:'',ttstyle:'',render:'',symbol:'',location:'',xpos:'',ypos:'',xscale:'',yscale:'',xoffset:'',
      yoffset:'',xsize:'', ysize: ''};
    if(!found && element in inheritable){//if the child didn't have any value for the MetaData and the property is among the inheritables

      parentMetaData.forEach(function(parentData){//checks if the parent has the data for it

        if( parentData.FieldID === element){
          // excel.set({row:row,column:4+index,value:parentData.Value});
          excel.cell(row,index+ 4).string(parentData.Value);
          result[1][row - 2][element] = parentData.Value;
          var str = 'XML.Binder[0]';

          for(i=0; i<counter.length; i++){//builds the XML endpoint that should change. at the endpoint the value for the data will be added from the parent
            str += '.Children[' + String(counter[i]) + ']';
          }


          if(typeof(target.MetaData.CustomMetaData) !== 'object'){//checks if CustomMetaData is not an array makes it an array
            eval(str + '.MetaData = {}')//makes MetaData inside XML an object
            eval(str + '.MetaData.CustomMetaData=[]'); // makes the CustomMetaData key to MetaData and puts [] as its value
          }else if(!Array.isArray(target.MetaData.CustomMetaData)){
            var text = str + '.MetaData.CustomMetaData=['+JSON.stringify(target.MetaData.CustomMetaData)+']';
            eval(text);
          }

          str += '.MetaData.CustomMetaData.push({FieldID:'+'"'+ element+'"' +',Value:'+'"'+parentData.Value+'"'+'})'; //builds the string to
          // add the value to the child from the parent
          eval(str); //executes adding the MetaData to the child
        }
      });
    }
  })
}

function getShort(files,excel,target,row,column,result){
  if(target.UUID){//if the target has UUID this part will play
    const UUID = target.UUID;
    var path = files.substr(0,files.lastIndexOf('/'))+ 'Files/Data/' + UUID +'/synopsis.txt';//building address of the synopsis.txt
    if (fs.existsSync(path)) {//if the synopsis.txt exests
      var text = fs.readFileSync(path).toString('utf-8');//This line reads the synopsis.txt
      excel.cell(row,column).string(text);//This line puts text inside excel file in shortDescription column
      result[0][row - 2].shortdescription = text;//This puts text inside result element 0 as shortdescription
      result[1][row - 2].shortdescription = text;//This puts text inside result element 1 as shortdescription
    }
  }
}

function getText(files,excel,target,row,column,result){
  if(target.UUID){
    const UUID = target.UUID;
    var path = files.substr(0,files.lastIndexOf('/'))+ 'Files/Data/' + UUID +'/content.rtf';//building address of the content.rtf
    if (fs.existsSync(path)) {//if the content.rtf exests
      var text = fs.readFileSync(path).toString('utf-8');//This line reads the content.rtf
      const begin = text.indexOf('fs20') + 'fs20'.length;
      const end = text.indexOf('fs24 <') - 1;
      text = text.slice(begin, end);
      text = text.replace('cf0', '');
      text = text.replace(/'91/g, "'");
      text = text.replace(/'92/g, "'");
      text = text.replace(/a0/g, ' ');

      var first,sub,last,portion;
      first = text.indexOf("fldinst{HYPERLINK");
      while(first > -1){ //until there is no more hyperlink
        //to get rid of the whole hyperlink
        first = first -11; // -11 is to compensate for the "{\field{\*\" which comes before this string

        sub =  text.substr(first,text.length);
        last = sub.indexOf('}}');
        portion = text.slice(first, last+first+2);
        text = text.replace(portion,'');
        first = text.indexOf("fldinst{HYPERLINK");
      }

      text = text.replace(/}}/g,"");
      text = text.replace(/{/g,"");
      text = text.replace(/\\/g, '');
      text = text.replace(/ldrslt/g,'');
      excel.cell(row,column).string(text);//This line puts text inside excel file in longDescription column
      result[0][row - 2].longdescription = text;
      result[1][row - 2].longdescription = text;

    }
  }

}


module.exports =  {createExcel}