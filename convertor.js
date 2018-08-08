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

function singleElement(element,index,excel,row){
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
       console.log(row);
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
  var Binder = XML.Binder;
  var excel = $JExcel.new();
  excel.set({row:1,column:1,value:'Title'});
  var row = 2;
  Binder.forEach(function(element,index){
    console.log(element.Title);
    console.log(row);
    row += 1;
    if(element.Children){
      singleElement(element,index,excel,row);
    }

  })

}




// function createExcel(files,XML){
//   const Binder = XML.Binder;
//   var excel = $JExcel.new();
//   excel.set({row:1,column:1,value:'Title'})
//   var childRowIndex = 2;
//   var addition1 = 0;
//   Binder.forEach(function(child){
//     addition1 = addElement(child,files,excel,childRowIndex,1);
//     childRowIndex += addition1;
//     if(child.Children && Array.isArray(child.Children)){//checking if the child has children and children is an array
//       grandChildRowIndex = childRowIndex  + 1;//setting the starting row for the children
//       var addition2 = 0;
//        child.Children.forEach(function(grandChild){
//         addition2 = addElement(grandChild,files,excel,grandChildRowIndex,2);
//         grandChildRowIndex += addition2;
//         // childRowIndex  += addition2;
//         if(grandChild.Children && Array.isArray(grandChild.Children)){

//           var forthGenerationRowIndex = grandChildRowIndex + 1;
//           var addition3 = 0;
//           grandChild.Children.forEach(function(forthGeneration){
//             // addition3 = addElement(forthGeneration,files,excel,forthGenerationRowIndex,3);
//             forthGenerationRowIndex += addition3;
//             // grandChildRowIndex += addition3;
//             // childRowIndex  += addition3;

//           })

//         }
//       });

//       childRowIndex  += 1+child.Children.length;
//     }else{
//       childRowIndex += 1;
//     }

//   });
//   setTimeout(function(){excel.generate("SampleData.xlsx"); }, 600);
// }



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