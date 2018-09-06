var fs = require("fs");





// var f = "/Users/shahab/lighthouse/scriv/render3/render0.3.scriv"
var f = "/Users/shahab/lighthouse/scriv/render3/render0.3.scriv"

 f = f + '/Files/Data/575ECF9D-5023-4920-BC65-A48CF0285D37/content.rtf'
  var text = fs.readFileSync(f).toString('utf-8');





// const begin = text.indexOf('fs20') + 'fs20'.length;
// const end = text.indexOf('fs24 <') - 1;


var first,sub,last,portion;
// first = text.indexOf("fldinst{HYPERLINK");
// while(first > -1){ //until there is no more hyperlink
//   //to get rid of the whole hyperlink
//   first = first -11; // -11 is to compensate for the "{\field{\*\" which comes before this string

//   sub =  text.substr(first,text.length);
//   last = sub.indexOf('}}');
//   portion = text.slice(first, last+first+2);
//   text = text.replace(portion,'');
//   first = text.indexOf("fldinst{HYPERLINK");
// }


function remove(text,begin, end){//finds all portion of text that begins with 'begin' and ends with 'end' and replaces them with ''.
  var last;
  var str = text;
  var first = str.indexOf(begin); //getting rid of the image part
  while(first > -1){
    sub =  str.substr(first,str.length);
    last = sub.indexOf(end) + first + 1;
    portion = str.slice(first, last);
    console.log(portion)
    str = str.replace(portion,"");
    first = str.indexOf(begin);
    console.log(str.length)
  }

  return str
}






function stripRtf(str){
    var basicRtfPattern = /\{\*?\\[^{}]+;}|[{}]|\\[A-Za-z]+\n?(?:-?\d+)?[ ]?/g;
    var newLineSlashesPattern = /\\\n/g;
    var ctrlCharPattern = /\n\\f[0-9]\s/g;

    //Remove RTF Formatting, replace RTF new lines with real line breaks, and remove whitespace
    return str
        .replace(ctrlCharPattern, "")
        .replace(basicRtfPattern, "")
        .replace(newLineSlashesPattern, "\n")
        .trim();
}


var first = text.indexOf("nisusfilename"); //getting rid of the image part



while(first > -1){ //until there is no more hyperlink
  //to get rid of the whole hyperlink
  first = first - 22;
  sub =  text.substr(first,text.length);
  last = text.lastIndexOf('}')


  portion = text.slice(first, last+first+3);
  text = text.replace(portion,'');
  first = text.indexOf("nisusfilename");

}
text = stripRtf(text);


text = text.replace('cf0', '');
text = text.replace(/'91/g, "'");
text = text.replace(/'92/g, "'");
text = text.replace(/a0/g, ' ')
  .replace(/\\/g, '')

text = remove(text,"<!$Scr_", ">");
text = remove(text,"<$Scr_", ">");

console.log(text)

// var char = '<$Scr_Ps::0>';

// str = str.replace(char,"")

// console.log(text);



