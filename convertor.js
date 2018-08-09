function addElement(target,files,counter,row,excel,uno){
  // getText(files,excel,target,index,column+1);//if the grandChild has UUID this function will add the text to the excel file
  uno = uno.map(a => a.toLowerCase());
  const outline = counter.map(a => a+1).map(String).reduce((a, b) => a + '.'+ b); //calculates outline number
  excel.set({row:row,column:uno.indexOf('outlinenumber')+ 4,value:outline});
  excel.set({row:row,column:3,value:target.Title});
  var CustomMetaData = target.MetaData.CustomMetaData;
  if(!Array.isArray(CustomMetaData)){
    CustomMetaData = [CustomMetaData];
  }
  var FieldID = '';
  var index = 0;

  CustomMetaData.forEach(function(element){
    if(element){
      FieldID = element.FieldID;
      index = uno.indexOf(FieldID);
      excel.set({row:row,column:4+index,value:element.Value});
    }
  })
}

function singleElement(element,index,excel,row,files,uno){
  var counter =[0];
  var finish = false;
  var c =0;

  var target = element;

  while  (finish === false && c<100 ) {
    validation = true;
    // console.log(counter);

    for(i=0; i<counter.length; i++){
      if(target.Children[counter[i]]){
        target =target.Children[counter[i]];
      }else{
        validation = false;
      }
    }


    if(validation){
       addElement(target,files,counter,row,excel,uno);
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

function initialize(excel){
  excel.set({row:1,column:3,value:'Title'});

  const uno = ["id", "label", "outlineNumber", "outlineLevel", "parent", "classes", "unoFrom", "unoTo", "param1", "param2",
  "param3", "param4", "shortDescription", "longDescription", "hoverAction", "hoverFunction", "clickAction", "clickFunction",
  "onDoubleClick", "url", "urlText", "tooltip", "infoPane", "onURL", "offURL", "openURL", "closeURL", "onFunction", "offFunction",
  "openFunction", "closeFunction", "ttStyle", "render", "symbol", "location", "xpos","panzoom", "ypos", "xsize", "ysize", "xoffset", "yoffset"];
  uno.forEach(function(element,index){
    excel.set({row:1,column:4+index,value:element});
  });
  return uno;

}

function createExcel(files,XML){

  var Binder = [XML.Binder[0]];
  var excel = $JExcel.new();
  const uno = initialize(excel);
  var row = 2;
  Binder.forEach(function(element,index){
    addElement(element,files,[-1],row,excel,uno);
    // excel.set({row:row,column:3,value:element.Title});
    console.log(row);
    row += 1;
    if(element.Children){
      row = singleElement(element,index,excel,row,files,uno);
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