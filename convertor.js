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
  var counter =[0];
  var finish = false;
  var c =0;
  var d = 0;

  var target = XML.Binder[0];

  while  (finish === false && c<100 ) {
    validation = true;
    console.log(counter);

    for(i=0; i<counter.length; i++){
      if(target.Children[counter[i]]){
        target =target.Children[counter[i]];
      }else{
        validation = false;
      }
    }


    if(validation){
      if(target.Children){
        counter.push(0);
      }else{
        counter[counter.length-1] +=1
      }
      console.log(target.Title);
    }else{
      counter.splice(-1,1);
      counter[counter.length-1] +=1
    }
    target = XML.Binder[0];
    c += 1;

    if(counter[0] == undefined){ // to stop the while loop when it has gone through all elements
      finish = true;
    }

  }
  console.log(c);
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

