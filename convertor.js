function createExcel(files,XML){
  const Binder = XML.Binder;
  var excel = $JExcel.new();
  excel.set({row:1,column:1,value:'Title'})
  var titleRowIndex = 2;
  var childSize = 0;
  Binder.forEach(function(child,index1){

    excel.set({row:(titleRowIndex),column:1,value:child.Title});
    // const childTitleRowIndex = 3;
    // child.forEach(function(grandChild,index2){
    //   excel.set({row:childTitleRowIndex,column:2,value:'Title'});
    //   const UUID = XML.Binder[index1].Children[index2].UUID;
    //   const textPath = findFile(files,'webkitRelativePath',UUID,'.rtf');
    //   readSingleFile(textPath,function(text2){
    //     getText(text2);
    //     excel.set({row:childTitleRowIndex+index2,column:3,value:'Title'});
    //   });
    // });
    if(child.Children){
      titleRowIndex += 1+child.Children.length;
    }else{
      titleRowIndex += 1;
    }

  });
  excel.generate("SampleData.xlsx");
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

