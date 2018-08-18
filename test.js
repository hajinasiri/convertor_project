var parseString = require('xml2js').parseString;
var fs = require("fs");







// var text = fs.readFileSync(".../scriv/Archive/workflow5.scriv/workflow5.scrivx").toString('utf-8');
//To build the address to the scrivx file
var f = "/Users/shahab/lighthouse/scriv/render2/render0.2.scriv";
var n = f.lastIndexOf('/');
var res = f.substr(n, f.length);
f = f + '/' +res+'x';
var text = fs.readFileSync(f).toString('utf-8');







if (fs.existsSync(f)) {
    var exist = true;
}

parseString(text, function (err, result) {

  try {
    var BinderMap = result.ScrivenerProject.Binder[0].BinderItem[0];
    var XML = {}
    // console.log(XML.ScrivenerProject.Binder[0].BinderItem[0].Children[0].BinderItem[0].Children);

    buildXML(BinderMap,XML);
    // console.log(XML);

  }
  catch(err) {
      console.log("There is not such fie");
  }

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
 var MetaData = "yes";
  if(target.MetaData && target.MetaData[0] && target.MetaData[0].CustomMetaData && target.MetaData[0].CustomMetaData[0]){
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




  str += ' ={Title:"'+target.Title[0]+'",UUID:"'+target.$.UUID+'",MetaData:"'+MetaData+'"}';

  eval(str);
   console.log(XML);



}









