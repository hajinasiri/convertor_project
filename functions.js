
function readSingleFile(f,cb) {
  //Retrieve the first (and only!) File from the FileList object

  if (f) {
    var r = new FileReader();
    r.readAsText(f);
    r.onload = function(e) {
      var contents = e.target.result;
      alert( "Got the file.n"
            +"name: " + f.name + "n"
            +"type: " + f.type + "n"
            +"size: " + f.size + " bytesn"
            + "starts with: " + contents.substr(1, contents.indexOf("n"))
      );
    }
    r.onload = function(e) {
      const text = e.target.result;
      cb(text);


    };

  } else {
    // alert("Failed to load file");
  }
}

function flatten(object) {
  var check = _.isPlainObject(object) && _.size(object) === 1;
  return check ? flatten(_.values(object)[0]) : object;
}

//to parse xml to an object
function parse(xml) {
  var data = {};

  var isText = xml.nodeType === 3,
      isElement = xml.nodeType === 1,
      body = xml.textContent && xml.textContent.trim(),
      hasChildren = xml.children && xml.children.length,
      hasAttributes = xml.attributes && xml.attributes.length;

  // if it's text just return it
  if (isText) { return xml.nodeValue.trim(); }

  // if it doesn't have any children or attributes, just return the contents
  if (!hasChildren && !hasAttributes) { return body; }

  // if it doesn't have children but _does_ have body content, we'll use that
  if (!hasChildren && body.length) { data.text = body; }

  // if it's an element with attributes, add them to data.attributes
  if (isElement && hasAttributes) {
    data.attributes = _.reduce(xml.attributes, function(obj, name, id) {
      var attr = xml.attributes.item(id);
      obj[attr.name] = attr.value;
      return obj;
    }, {});
  }

  // recursively call #parse over children, adding results to data
  _.each(xml.children, function(child) {
    var name = child.nodeName;

    // if we've not come across a child with this nodeType, add it as an object
    // and return here
    if (!_.has(data, name)) {
      data[name] = parse(child);
      return;
    }

    // if we've encountered a second instance of the same nodeType, make our
    // representation of it an array
    if (!_.isArray(data[name])) { data[name] = [data[name]]; }

    // and finally, append the new child
    data[name].push(parse(child));
  });

  // if we can, let's fold some attributes into the body
  _.each(data.attributes, function(value, key) {
    if (data[key] != null) { return; }
    data[key] = value;
    delete data.attributes[key];
  });

  // if data.attributes is now empty, get rid of it
  if (_.isEmpty(data.attributes)) { delete data.attributes; }

  // simplify to reduce number of final leaf nodes and return
  return flatten(data);
}


function findFile (files,key,str1,str2) {
//to finds the file by key and returns the corresponding array element from files
  var result='';
  for (i = 0; i < files.length; i++ ){
      if (files[i][key].indexOf(str1) > -1 && files[i][key].indexOf(str2) > -1 ){
        result = (files[i]);
      }
  }
  return result;
}

function getText(files,excel,target,row,column){
  if(target.UUID){
    const UUID = target.UUID;
    const textPath = findFile(files,'webkitRelativePath',UUID,'.rtf');
    readSingleFile(textPath,function(rawText){
      var text = rawText;
      const begin = text.indexOf('fs20') + 'fs20'.length;
      const end = text.indexOf('fs24 <') - 1;
      text = text.slice(begin, end);
      text = text.replace('\cf0', '');
      text = text.replace("'91", "'");
      text = text.replace("'92", "'");
      text = text.replace("'a0", ' ');

      var first,sub,last,portion;
      first = 1;
      while(first > -1){ //until there is no more hyperlink
        //to get rid of the whole hyperlink
        first = text.indexOf("fldinst{HYPERLINK");
        first = first -11; // -11 is to compensate for the "{\field{\*\" which comes before this string
         if(row == 9){
          console.log(first);
        }
        sub =  text.substr(first,text.length);
        last = sub.indexOf('}}');
        portion = text.slice(first, last+first+2);
        text = text.replace(portion,'');
      }

      //To get rid of {\fldrslt  and }} around the links
      text = text.replace('{\fldrslt ', '');
      text = text.replace('}}', '');
      if(row === 9){
        console.log(text);
      }

      excel.set({row:row,column:column,value:text});
    });
  }
}
