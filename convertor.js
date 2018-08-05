function createExcel(files,XML){
  const Binder = XML.Binder;
  var excel = $JExcel.new();
  excel.set({row:1,column:1,value:'Title'})
  var titleRowIndex = 2;
  Binder.forEach(function(child){

    excel.set({row:titleRowIndex,column:1,value:child.Title});


    if(child.Children && Array.isArray(child.Children)){
      grandChildRowIndex = titleRowIndex + 1;
       child.Children.forEach(function(grandChild){
        if(grandChild.UUID){
          getText(files,excel,grandChild,grandChildRowIndex,3);
          // excel.set({row:titleRowIndex,column:3,value:getText(files,grandChild,grandChildRowIndex)});
        }
        excel.set({row:grandChildRowIndex,column:2,value:grandChild.Title});
        if(grandChild.Children){
          grandChildRowIndex += 1+grandChild.Children.length;
          titleRowIndex += 1+grandChild.Children.length;
        }else{
          grandChildRowIndex += 1;
          titleRowIndex += 1;
        }
      // //   const UUID = XML.Binder[index1].Children[index2].UUID;
      // //   const textPath = findFile(files,'webkitRelativePath',UUID,'.rtf');

      });
      titleRowIndex += 1+child.Children.length;
    }else{
      titleRowIndex += 1;
    }

  });
  setTimeout(function(){excel.generate("SampleData.xlsx"); }, 600);
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

