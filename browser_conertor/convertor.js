function addElement(XML,target,files,counter,row,excel,uno,parent,result){
  getShort(files,excel,target,row,uno.indexOf('shortdescription') + 4,result);
  getText(files,excel,target,row,uno.indexOf('longdescription') + 4,result);
  var outline = counter.map(a => a+1).map(String ).reduce((a, b) => a + '-' + b); //calculates outline number from counter variable
  excel.set({row:row,column:3,value:target.Title});//sets the title column in excel
  excel.set({row:row,column:uno.indexOf('id') + 4,value:target.Title.replace(/ /g,'')});//sets the id column in excel
  excel.set({row:row,column:uno.indexOf('label') + 4,value:target.Title});//sets the label column in excel
  excel.set({row:row,column:uno.indexOf('outlinenumber') + 4,value:outline});// sets the outline number in excel
  var outlineLevel;
  if(counter[0] === -1){//sets the outlinelevel column
    excel.set({row:row,column:uno.indexOf('outlinelevel') + 4,value:0});//sets 0 for map
    outlineLevel = 0;
  }else{
    excel.set({row:row,column:uno.indexOf('outlinelevel') + 4,value:counter.length});//calculates and sets outlinelevel for things other than map
    outlineLevel = counter.length;
  }


  excel.set({row:row,column:uno.indexOf('parent') + 4,value:parent.Title});//sets the parent column in excel

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
  excel.set({row:row,column:uno.indexOf('classes') + 4,value:classes}); //sets the value of classes column in excel as classes variable value
  result[0][row - 2 ]= {title:target.Title, id:target.Title.replace(/ /g,''), label:target.Title, outlineNumber:outline, outlineLevel:outlineLevel, parent:parent.Title,classes:classes }; //putting the calculated metadata as the object in result array
  result[1][row - 2 ] = {title:target.Title, id:target.Title.replace(/ /g,''), label:target.Title, outlineNumber:outline, outlineLevel:outlineLevel, parent:parent.Title,classes:classes }; //putting the calculated metadata as the object in result array
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
        excel.set({row:row,column:4+index,value:childData.Value});
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
      excel.set({row:row,column:4+uno.indexOf('unofrom'),value:id}); // if there is value for unoto, but no value for unofrom then the excel column value for unofrom is set as the value of id
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
          excel.set({row:row,column:4+index,value:parentData.Value});
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


function createExcel(files,XML,name){//fetches data from XML, Uses addElement function to add the data to excel file. addElement function itself
  //uses propagate function to make child inherit MetaData from their parent
  var result =[[],[]];
  var Binder;
  XML.Binder.forEach(function(element){//to find Map element and put it in Binder varialble
    if(element.Title === 'Map'){
      Binder = [element];
    }
  })
  var excel = $JExcel.new(); //initiates excel file
  excel.set( {sheet:0,value:"Structure" } ); //Setting the sheet name
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
    excel.generate(name+'.xlsx'); //generates the excel file. Uses setTime to let async readSingleFile function inside getText function read the rtf files and add them to the excel.
  }, 300);

}




function main(evt) { //This is the main function. Gets triggered when the button on the browser is clicked.
  const files = evt.target.files; //puts the read files in variable files
  const f =findFile(files,'name','.scrivx',''); //finds the scrivx file and puts it in f
  const name = f.name.replace('.scrivx','');

  readSingleFile(f,function(text1){//reads the scrivx file and as call back puts it in xml, parses it and creates the excel file from it
    const xml = new DOMParser().parseFromString(text1, "text/xml"); //Parses the text into a DOM and puts it in xml variable
    const XML = parse(xml);//uses parse function to create an object from the DOM and puts it in XML variable
    createExcel(files,XML,name);//uses createExcel function to create the excel file from XML. files variable is passed to the function to read more files from it if needed
    console.log(XML);
  });

}

document.addEventListener('DOMContentLoaded', function () {
  document.getElementById('fileinput').addEventListener('change', main, false);
});
