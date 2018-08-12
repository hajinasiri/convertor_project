function addElement(XML,target,files,counter,row,excel,uno,parent){
  // getText(files,excel,target,index,column+1);//if the grandChild has UUID this function will add the text to the excel file
  uno = uno.map(a => a.toLowerCase());//Makes all uno titles lowercase to be able to search them
  var outline = counter.map(a => a+1).map(String ).reduce((a, b) => a + '-' + b); //calculates outline number
  excel.set({row:row,column:uno.indexOf('outlinenumber')+ 4,value:outline});// sets the outline number in excel
  excel.set({row:row,column:3,value:target.Title});//sets the title column in excel
  excel.set({row:row,column:uno.indexOf('label')+ 4,value:target.Title});//sets the label column in excel

  if(counter[0] === -1){//sets the outlinelevel column
    excel.set({row:row,column:uno.indexOf('outlinelevel')+ 4,value:0});//sets 0 for map
  }else{
    excel.set({row:row,column:uno.indexOf('outlinelevel')+ 4,value:counter.length});//calculates and sets outlinelevel for things other than map
  }

  excel.set({row:row,column:uno.indexOf('id')+ 4,value:target.Title.replace(/ /g,'')});//sets the id column in excel
  excel.set({row:row,column:uno.indexOf('parent')+ 4,value:parent.Title});//sets the parent column in excel
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



   if(target.Title === 'Map'){
    console.log(target);
    console.log(CustomMetaData);
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
  var counter =[0];
  var finish = false;
  var c =0;

  var target = element;
  var parent = element;

  while  (finish === false && c<100 ) {
    validation = true;


    for(i=0; i<counter.length; i++){
      if(target.Children[counter[i]]){
        target =target.Children[counter[i]];
      }else{
        validation = false;
      }
      if(i === (counter.length - 2)){
        parent = target;
      }
    }


    if(validation){
      addElement(XML,target,files,counter,row,excel,uno,parent);

       row += 1;
      if(target.Children){
        counter.push(0);
      }else{
        counter[counter.length-1] +=1
      }

    }else{
      counter.splice(-1,1);
      counter[counter.length-1] +=1
    }
    target = element;
    c += 1;

    if(counter[0] == undefined){ // to stop the while loop when it has gone through all elements
      finish = true;
    }

  }

  return row;
}

function initialize(excel){ //initializes the MetaData columns inside excel file
  excel.set({row:1,column:3,value:'Title'});

  const uno = ["id", "label", "outlineNumber", "outlineLevel", "parent", "classes", "unoFrom", "unoTo", "param1", "param2",
  "param3", "param4", "shortDescription", "longDescription", "hoverAction", "hoverFunction", "clickAction", "clickFunction",
  "onDoubleClick", "url", "urlText", "tooltip", "infoPane", "onURL", "offURL", "openURL", "closeURL", "onFunction", "offFunction",
  "openFunction", "closeFunction", "ttStyle", "render", "symbol", "location", "xpos","panzoom", "ypos", "xsize", "ysize", "xoffset",
  "yoffset","slideurl",'subtitle'];

  uno.forEach(function(element,index){
    excel.set({row:1,column:4+index,value:element});
  });
  return uno;

}

function createExcel(files,XML){//fetches data from XML, Uses addElement function to add the data to excel file. addElement function itself
  //uses propagate function to make child inherit MetaData from their parent

  var Binder = [XML.Binder[0]];
  var excel = $JExcel.new();
  const uno = initialize(excel);
  var row = 2;
  Binder.forEach(function(element,index){
    addElement(XML,element,files,[-1],row,excel,uno,XML.Binder);
    // excel.set({row:row,column:3,value:element.Title});

    row += 1;

    if(element.Children){
      row = singleElement(XML,element,index,excel,row,files,uno);
    }

  })
  excel.generate('converted.xlsx');

}




function main(evt) {
  const files = evt.target.files;
  const f =findFile(files,'name','.scrivx','');

  readSingleFile(f,function(text1){
    const xml = new DOMParser().parseFromString(text1, "text/xml");
    const XML = parse(xml);
    createExcel(files,XML);
    console.log(XML);
  });

}

document.addEventListener('DOMContentLoaded', function () {
  document.getElementById('fileinput').addEventListener('change', main, false);
});