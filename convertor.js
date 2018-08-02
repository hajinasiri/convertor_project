// var fs = require("fs");
// var DOMParser = require('xmldom').DOMParser;
// var _ = require('lodash');



// const filename = "../../Downloads/Archive/workflow5.scriv/workflow5.scrivx";
// var text= fs.readFileSync(filename, "utf8");


// var parser = new DOMParser()
// var el = parser.parseFromString(text, "text/xml");
// const Binder = el.getElementsByTagName("Binder");
// // console.log(el.documentElement == "parsererror" ? "error while parsing" : el.documentElement);






function flatten(object) {
  var check = _.isPlainObject(object) && _.size(object) === 1;
  return check ? flatten(_.values(object)[0]) : object;
}


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



function download(filename, text) {
  var element = document.createElement('a');
  element.setAttribute('href', 'data:text/plain;charset=utf-8,' + encodeURIComponent(text));
  element.setAttribute('download', filename);

  element.style.display = 'none';
  document.body.appendChild(element);

  element.click();

  document.body.removeChild(element);
}




function readSingleFile(evt) {
  //Retrieve the first (and only!) File from the FileList object
  var f = evt.target.files[0];
  console.log(f);

  if (f) {
    var r = new FileReader();
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
      const text =e.target.result;
      const xml = new DOMParser().parseFromString(text, "text/xml");
      const XML = parse(xml);
      console.log(XML);
      download('file', text);
    };
    r.readAsText(f);

  } else {
    alert("Failed to load file");
  }
}

document.addEventListener('DOMContentLoaded', function () {
  document.getElementById('fileinput').addEventListener('change', readSingleFile, false);
});


