function addElement(object,files,excel,index,column){
  getText(files,excel,object,index,column+1);//if the grandChild has UUID this function will add the text to the excel file
  excel.set({row:index,column:column,value:object.Title});
  var rowAdditionNumber = 0;
  if(object.Children && Array.isArray(object.Children)){
    rowAdditionNumber = 1+object.Children.length;
  }else{
  rowAdditionNumber = 1;
  }
  return rowAdditionNumber;
}



function createExcel(files,XML){
  const Binder = XML.Binder;
  var excel = $JExcel.new();
  excel.set({row:1,column:1,value:'Title'})
  var titleRowIndex = 2;
  Binder.forEach(function(child){
    excel.set({row:titleRowIndex,column:1,value:child.Title});
    if(child.Children && Array.isArray(child.Children)){//checking if the child has children and children is an array
      grandChildRowIndex = titleRowIndex + 1;//setting the starting row for the children
      var addition = 0;
       child.Children.forEach(function(grandChild){
        addition = addElement(grandChild,files,excel,grandChildRowIndex,2);
        grandChildRowIndex += addition;
        titleRowIndex += addition;
        if(grandChild.Children && Array.isArray(grandChild.Children)){
          // const forthGenerationRowIndex = grandChildRowIndex + 1;
          // grandChild.Children.forEach(function(forthGeneration){
          //   excel.set({row:forthGenerationRowIndex,column:3,value:forthGeneration.Title});

          // })

        }
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

