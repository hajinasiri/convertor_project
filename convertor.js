function addElement(XML,target,files,counter,row,excel,uno,parent){
  getShort(files,excel,target,row,uno.indexOf('shortdescription') + 4);
  getText(files,excel,target,row,uno.indexOf('longdescription') + 4)
  var outline = counter.map(a => a+1).map(String ).reduce((a, b) => a + '-' + b); //calculates outline number from counter variable
  excel.set({row:row,column:uno.indexOf('outlinenumber') + 4,value:outline});// sets the outline number in excel
  excel.set({row:row,column:3,value:target.Title});//sets the title column in excel
  excel.set({row:row,column:uno.indexOf('label') + 4,value:target.Title});//sets the label column in excel

  if(counter[0] === -1){//sets the outlinelevel column
    excel.set({row:row,column:uno.indexOf('outlinelevel') + 4,value:0});//sets 0 for map
  }else{
    excel.set({row:row,column:uno.indexOf('outlinelevel') + 4,value:counter.length});//calculates and sets outlinelevel for things other than map
  }

  excel.set({row:row,column:uno.indexOf('id') + 4,value:target.Title.replace(/ /g,'')});//sets the id column in excel
  excel.set({row:row,column:uno.indexOf('parent') + 4,value:parent.Title});//sets the parent column in excel
  propagate(XML,excel,row, target,parent,uno,counter);



}


function propagate(XML,excel,row, target,parent,uno,counter){

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

  uno.forEach(function(element,index){//goes through all the MetaData and checks if the child has that value or the parent and puts that vlue in excel file
     var found = false;
    CustomMetaData.forEach(function(childData){//Checks if the child has a value for it
      if( childData.FieldID === element){
        excel.set({row:row,column:4+index,value:childData.Value});
        found = true;
      }
    });

    if(!found){//if the child didn't have any value for the MetaData
      parentMetaData.forEach(function(parentData){//checks if the parent has the data for it

        if( parentData.FieldID === element){
          excel.set({row:row,column:4+index,value:parentData.Value});
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


function singleElement(XML,element,index,excel,row,files,uno){
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
      addElement(XML,target,files,counter,row,excel,uno,parent);
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

function initialize(excel){ //initializes the MetaData columns inside excel file
  excel.set({row:1,column:3,value:'Title'});

  var uno = ["id", "label", "outlineNumber", "outlineLevel", "parent", "classes", "unoFrom", "unoTo", "param1", "param2",
  "param3", "param4", "shortDescription", "longDescription", "hoverAction", "hoverFunction", "clickAction", "clickFunction",
  "onDoubleClick", "url", "urlText", "tooltip", "infoPane", "onURL", "offURL", "openURL", "closeURL", "onFunction", "offFunction",
  "openFunction", "closeFunction", "ttStyle", "render", "symbol", "location", "xpos","panzoom", "ypos", "xsize", "ysize", "xoffset",
  "yoffset","slideurl",'subtitle'];

  uno.forEach(function(element,index){
    excel.set({row:1,column:4+index,value:element});
  });
  uno = uno.map(a => a.toLowerCase());//Makes all uno titles lowercase to be able to search them
  return uno;

}

function createExcel(files,XML){//fetches data from XML, Uses addElement function to add the data to excel file. addElement function itself
  //uses propagate function to make child inherit MetaData from their parent

  var Binder = [XML.Binder[0]]; //puts Map in Binder varialble
  var excel = $JExcel.new(); //initiates excel file
  const uno = initialize(excel);//uses initialize function to add all uno elements to the excel file. in rturn gets uno array and puts it in uno variable
  var row = 2;
  Binder.forEach(function(element,index){ //this forEach is here to go through all the elements in Binder if needed. But currently there is only map in Binder variable
    addElement(XML,element,files,[-1],row,excel,uno,{Title:'Binder'}); //This line adds all the data of Map to excel file
    row += 1; //Then sets row number to the next row number

    if(element.Children){ //if there is children in Map, uses singleElement function to go through all the children inside map and gets the last row number from that function
      row = singleElement(XML,element,index,excel,row,files,uno);
    }

  })

  // setTimeout(function(){
  //   excel.generate('converted.xlsx'); //generates the excel file. Uses setTime to let async readSingleFile function inside getText function read the rtf files and add them to the excel.
  // }, 300);

}




function main(evt) { //This is the main function. Gets triggered when the button on the browser is clicked.
  const files = evt.target.files; //puts the read files in variable files
  const f =findFile(files,'name','.scrivx',''); //finds the scrivx file and puts it in f

  readSingleFile(f,function(text1){//reads the scrivx file and as call back puts it in xml, parses it and creates the excel file from it
    const xml = new DOMParser().parseFromString(text1, "text/xml"); //Parses the text into a DOM and puts it in xml variable
    const XML = parse(xml);//uses parse function to create an object from the DOM and puts it in XML variable
    createExcel(files,XML);//uses createExcel function to create the excel file from XML. files variable is passed to the function to read more files from it if needed
    console.log(XML);
  });

}

document.addEventListener('DOMContentLoaded', function () {
  document.getElementById('fileinput').addEventListener('change', main, false);
});