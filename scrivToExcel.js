var parseString = require('xml2js').parseString;
var fs = require("fs");
var modules = require('./scrivModules');

require('log-timestamp');
// var f = "/Users/shahab/lighthouse/scriv/render3/GenderFinance4.7test.scriv";

var f = process.argv[2];//reads the file address from user input in terminal



// const buttonPressesLogFile = '/Users/shahab/lighthouse/scriv/render3/test.json';

console.log(`Watching for file changes on ${f}`);

fs.watchFile(f, (curr, prev) => {
  console.log(`${f} file Changed`);
  main(f);
});

main(f);



function main(f) {
  var n = f.lastIndexOf('/');
  var res = f.substr(n, f.length);
  f = f + '/' +res+'x';
  var text = fs.readFileSync(f).toString('utf-8');

    // var resultt =[[],[]];

  parseString(text, function (err, result) {//this line parses the text. the output is text. Then the desirable XML format will be built from this output
    var settings = result.ScrivenerProject.CustomMetaDataSettings[0].MetaDataField;//getting the metaDataSettings from the parsed text

    settings = settings.map(a => {//This creates a desirable format of settings and puts it inside the settings
      var obj =a.$;
      obj.Title = a.Title[0];
      return obj
    });
    var output = {};
    output.CustomMetaDataSettings = settings;//output is the variable holding the desirable XML format. This line adds the settings to it


    var Keywords = result.ScrivenerProject.Keywords[0].Keyword;//reads the Kewords from text
    Keywords = Keywords.map(a => {//makes the format desirable
      return {ID:a.$.ID,Title:a.Title[0],Color:a.Color[0]}
    })
    output.Keywords = Keywords;//adds the desirable format of keywords to the output


    var Binder = result.ScrivenerProject.Binder[0].BinderItem;//putts the Binder inside Binder Variable

    var MapArray;
    var config;
    Binder.forEach(function(element){//Finds the map inside Binder
      if(element.Title[0] === 'Map'){
        MapArray = element;
      }else if(element.Title[0] === 'config.json'){//finds config and pust it in config variable
        config = element;
      }
    });
    var configObject = modules.createConfig(f,config);//creates the confing.json file
    var XML = {}
    XML = addToXML(MapArray,[],XML);//adds map object to the final xml
    buildXML(MapArray,XML);//builds the desirable xml format from MapArray
    output.Binder = [XML];//sets the output.Binder to an array with XML inside
    var finalResult = modules.createExcel(f,output,'name');//creates the excel file from the final desirable xml format that is stored in output
    modules.findDuplicates(finalResult);//check if there are duplicate ids
    createStory(finalResult[1],f);

  });
}

function createStory(finalResult,f){
  // story,html code and animation

  var storyData = "";
  var voaData = "{\r\r";
  var voaIndex = 0;
  var indexString;
  var storyLink = "";
  var hasChildren;
  var div;
  // Find the story and/or voa links: go through all the UNO's checking for a classes of 'story' or 'voa'

  // Note that voa elements may also have a story link
  finalResult.forEach(function(element,index){
    // console.log(element.shortdescription);
    // Find any voa's
    if(element.classes.includes("voa") ){
      if(voaIndex == 0){ // add the starting element
        if( voaData.charAt(voaData.length - 1) == "}"){ // add the closing comma to the previous voa
            voaData += ",\r\r";
        }
        voaData += "\"" + element.id + "\": {  \"elements\":  [\r";
        voaIndex ++;
      }
      else {  // add a regular animation element
        voaData += "{ \"type\": \"synchronous\", \"elements\": [ { \"type\":\"url\", \"content\":\"";
        if(element.slideurl){//Checks if the element has slideurl, then adds it to voaData
          voaData += element.slideurl;
        }
        indexString = voaIndex;
        if(indexString.toString().length === 1){ //checks if the indexString is just one digit
          indexString = voaIndex.toString().padStart(2, '0');  // the MP3 file's index number must have 2 digits
        }

        voaData += "\" } , { \"type\":\"audio\", \"content\":\"audio/" + element.id + " " + indexString + ".mp3\" } ] }";

        if(element.outlinelevel > finalResult[index + 1].outlinelevel){  // we've hit the last item of this animation
          voaData += "\r] }";
          voaIndex = 0;
        }
        else {  // more elements yet to add
          voaData += ",\r";
          voaIndex ++;
        }
      }
    }

    // Adds the closing element to the json data


    // Finds any story links
    if(element.classes.includes("story") ){

      hasChildren = false;
      finalResult.forEach(function(children){//checks if any of the unos are children of the current element, then sets hasChildren as true
        if(children.parent === element.id){
          hasChildren = true;
        }
      });

      if(hasChildren){
        storyData += "<h"+element.outlinelevel+" class='storyHead storyHead"+element.outlinelevel.toString().padStart(2, '0')+
        "'> <button aria-expanded='false'><i class='icon-right-dir'></i><i class='icon-down-dir'></i>"+element.title+"</row></button></h"+
        element.outlinelevel+"><div hidden>"
      }else{
        storyData += '<a href='+"'"+"#/?"+(element.slideurl? element.slideurl:"");
        if(element.classes.includes('fileInfo') || element.classes.includes('formInfo')){
          storyData += element.classes;
        }else if (!element.slidurl){
          storyData += "+++&unoInfo=" + element.id;
        }else{
          storyData += "&unoInfo=" + element.id;
        }
        storyData += "' id='storyLink" + element.id + "'   class='slide storyItem"+
        (hasChildren? finalResult[index + 1].outlinelevel : element.outlinelevel).toString().padStart(2, '0') +
        "' >" + element.label + "</a><br>";
      }
      //Below creates an string of divs that are needed after each uno's html

      if(finalResult[index+1]){
        div = '';
        for(i=0; i < element.outlinelevel - finalResult[index+1].outlinelevel; i++){
          div = div + '</div>'
        }
      }else{
        for(i=1; i < element.outlinelevel; i++){
          div = div + '</div>'
        }
      }
      storyData += div;

      // adding the CR between links
      storyLink += "\r";
      storyData += storyLink;
      storyLink = "";

    }
  });

  var storyPath = f.substr(0,f.lastIndexOf('/') - 1);
  storyPath = storyPath.substr(0,storyPath.lastIndexOf('/')) + '/story.html';//Builds the path that animate.json will get written to
  fs.writeFile(storyPath, storyData, function(err) {//writes the animate.json file
    if(err) {
      return console.log(err);
    }
    console.log("story.html file was saved!");
  });


  voaData += "\r\r}";
  // Save the voaData to the animate.json file
  var animatePath = f.substr(0,f.lastIndexOf('/') - 1);
      animatePath = animatePath.substr(0,animatePath.lastIndexOf('/')) + '/animate.json';//Builds the path that animate.json will get written to
  fs.writeFile(animatePath, voaData, function(err) {//writes the animate.json file
    if(err) {
      return console.log(err);
    }
    console.log("animate.json file was saved!");
  });

}

function buildXML (BinderMap,XML){
  var counter = [0];
  var target = BinderMap;
  var finish = false;
  var validation;
  var c = 0;


  while(finish === false){

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
        obj.FieldID = a.FieldID[0];
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