
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

// function myAsyncFunction(path) {
//   return new Promise((resolve, reject) => {


function getText(text){
  const begin = text.indexOf('fs20') + 'fs20'.length;
  const end = text.indexOf('fs24 <') - 1;
  const mainText = text.slice(begin, end);
  console.log(text);
  console.log(mainText);
}





function main(evt) {
  const files = evt.target.files;
  const f =findFile(files,'name','.scrivx','');
  readSingleFile(f,function(text1){
    const xml = new DOMParser().parseFromString(text1, "text/xml");
    const XML = parse(xml);
    const UUID = XML.Binder[0].Children[0].UUID;
    const textPath = findFile(files,'webkitRelativePath',UUID,'.rtf');
    readSingleFile(textPath,function(text2){
      getText(text2);
    })
  });

}

document.addEventListener('DOMContentLoaded', function () {
  document.getElementById('fileinput').addEventListener('change', main, false);
});

