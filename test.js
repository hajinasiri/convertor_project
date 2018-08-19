var parseString = require('xml2js').parseString;
var fs = require("fs");
var modules = require('./modules');


var f = "/Users/shahab/lighthouse/scriv/render2/render0.2.scriv";
var n = f.lastIndexOf('/');
var res = f.substr(n, f.length);
f = f + '/' +res+'x';
var text = fs.readFileSync(f).toString('utf-8');



if (fs.existsSync(f)) {
    var exist = true;
}

parseString(text, function (err, result) {
  var settings = result.ScrivenerProject.CustomMetaDataSettings[0].MetaDataField;

  settings = settings.map(a => {
    var obj =a.$;
    obj.Title = a.Title[0];
    return obj
  });
  var output = {};
  output.CustomMetaDataSettings = settings;


  var Keywords = result.ScrivenerProject.Keywords[0].Keyword;
  Keywords = Keywords.map(a => {
    return {ID:a.$.ID,Title:a.Title[0],Color:a.Color[0]}
  })
  output.Keywords = Keywords;


  var Binder = result.ScrivenerProject.Binder[0].BinderItem;//putting the Binder inside Binder Variable
  var MapArray;
  var config;
  Binder.forEach(function(element){//Finding the map inside Binder
    if(element.Title[0] === 'Map'){
      MapArray = element;
    }else if(element.Title[0]= 'Recovered Files (Aug 18, 2018 at 8:17 PM)'){
      config = element;
    }
  });

  var XML = {}
  XML = addToXML(MapArray,[],XML);
  buildXML(MapArray,XML);
  output.Binder = [XML];
  modules.createExcel(f,output,'name');


});


function buildXML (BinderMap,XML){
  var counter = [0];
  var target = BinderMap;
  var finish = false;
  var validation;
  var c = 0;


  while(finish === false){


    // console.log(counter);
    validation = true;

     for(i=0; i<counter.length; i++){ //this loop uses counter as the address to get to the child and store it in the target variable

      if(target.Children && target.Children[0].BinderItem[[counter[i]]]){//if there is a child in that address, sets the target
        target =target.Children[0].BinderItem[[counter[i]]];

      }else{ // if there is not a child in that address sets validation false
        validation = false;
      }
    }

    if(validation){//If the validation is true and there is a child in the address, addElement is used to add the data to the excel file
      // console.log(target.Title);
      // console.log(counter);
      addToXML(target,counter,XML);

      if(target.Children && target.Children[0].BinderItem){ //if the target has children sets the counter to the first of them
        counter.push(0);
      }else{//if the target does not have children, sets counter to the next sibling of the target
        counter[counter.length-1] +=1
      }

    }else{ //if there is not a child at the addressed obtained from the counter, sets counter to one generation back and the next sibling in that generation
      // console.log('validation flase')
      counter.splice(-1,1);
      counter[counter.length-1] +=1
    }
    target = BinderMap;//resets target to the first generation ancestor to make it ready to build the next address from
    if(counter[0] == undefined){ // when the loop goes through all the children and then comes back to the first generation and there is no more sibling there, it goes one generation back and first element in counter becomes undefined. That's how the loop realizes it should end.
      finish = true;
    }

  }

}

function addToXML(target,counter,XML){
  var str = 'XML';
  if(counter[0] !== undefined){//counter = [] is reserved for the Map object

    counter.forEach(function(element){//making the string for going to the right address in XML

      str += '.Children[' + element+']';
    })
    var n = str.lastIndexOf('Children');
    var res = str.substr(0,n) + 'Children';

   if(!eval(res)){//cheking if the children does not exist in XML, making the Children empty array
    eval(res+'=[]');

   }else if(!eval(str)){//Chekcing if the object in Children array does not exist, putting an empty object as its value
    eval(str+'={}');
   }
 }

 var MetaData = "yes";
  if(target.MetaData && target.MetaData[0] && target.MetaData[0].CustomMetaData && target.MetaData[0].CustomMetaData[0]){//extracting MetaData from rawXML target
    if(target.MetaData[0].CustomMetaData[0].MetaDataItem){
      var MetaData = target.MetaData[0].CustomMetaData[0].MetaDataItem;
      MetaData = MetaData.map(a => {
        var obj = {};
        obj[a.FieldID] = a.FieldID[0];
        obj.Value = a.Value[0]
        return obj
      })
    }
  }
  MetaData = {CustomMetaData:MetaData};
  var keyStr = str;
  str += ' ={Title:"'+target.Title[0]+'",UUID:"'+target.$.UUID+'"}';//Setting Title and UUID for each children

  eval(str);

  eval(keyStr+'.MetaData=MetaData');//Setting the data for each children

  if(target.Keywords && target.Keywords[0]){
    eval(keyStr+'.Keywords=target.Keywords[0].KeywordID')//setting the keywords for each children

  }
  return XML
}









