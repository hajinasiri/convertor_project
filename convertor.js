function addElement(object,files,counter,row,excel){
  // getText(files,excel,object,index,column+1);//if the grandChild has UUID this function will add the text to the excel file
  const id = counter.map(a => a+1).map(String).reduce((a, b) => a + '.'+ b);
  excel.set({row:row,column:1,value:id});
  excel.set({row:row,column:2,value:object.Title});
}

function singleElement(element,index,excel,row,files){
  var counter =[0];
  var finish = false;
  var c =0;

  var target = element;

  while  (finish === false && c<100 ) {
    validation = true;
    // console.log(counter);

    for(i=0; i<counter.length; i++){
      if(target.Children[counter[i]]){
        target =target.Children[counter[i]];
      }else{
        validation = false;
      }
    }


    if(validation){
       console.log(target.Title);
       console.log(counter);
       console.log(row);
       addElement(target,files,counter,row,excel);
       row += 1;
      if(target.Children){
        counter.push(0);
      }else{
        counter[counter.length-1] +=1
      }

    }else{
      counter.splice(-1,1);
      counter[counter.length-1] +=1
    }
    target = element;
    c += 1;

    if(counter[0] == undefined){ // to stop the while loop when it has gone through all elements
      finish = true;
    }

  }
  return row;
}


function createExcel(files,XML){
  var Binder = [XML.Binder[0]];
  var excel = $JExcel.new();
  excel.set({row:1,column:1,value:'id'},{row:1,column:2,value:'Title'});
  var row = 2;
  Binder.forEach(function(element,index){
    excel.set({row:row,column:2,value:element.Title});
    console.log(row);
    row += 1;
    if(element.Children){
      row = singleElement(element,index,excel,row,files);
    }

  })
  excel.generate('converted.xlsx');

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