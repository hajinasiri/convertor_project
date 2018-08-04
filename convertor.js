
// //to download a file to client's computer
// function download(filename, text) {
//   var element = document.createElement('a');
//   element.setAttribute('href', 'data:text/plain;charset=utf-8,' + encodeURIComponent(text));
//   element.setAttribute('download', filename);

//   element.style.display = 'none';
//   document.body.appendChild(element);

//   element.click();

//   document.body.removeChild(element);
// }



function readSingleFile(f,cb) {
  //Retrieve the first (and only!) File from the FileList object

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
      cb(XML);
    };
    r.readAsText(f);
  } else {
    alert("Failed to load file");
  }
}



function main(evt) {
  const files = evt.target.files;
  const f =findFile(files,'name','.scrivx');
  const data = readSingleFile(f,function(XML){
    const UUID = XML.Binder[0].Children[0].UUID;
    const textPath = findFile(files,'webkitRelativePath',UUID);
    // console.log(textPath);
    // console.log(files);
    myAsyncFunction().then(val => {
      console.log(val);
    })
  });

}

document.addEventListener('DOMContentLoaded', function () {
  document.getElementById('fileinput').addEventListener('change', main, false);
});

