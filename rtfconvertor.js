// const rtfParse = require( 'rtf-parse' );
//rtfParse.parseString( text )
//     .then( doc => {
//         console.log(doc.children[0].children[0])
//     } );

var fs = require("fs");

var f = "/Users/shahab/lighthouse/scriv/render3/render0.3.scriv"

 f = f + '/Files/Data/0B289284-0623-4EAD-93D2-82698847D120/content.rtf'
  var text = fs.readFileSync(f).toString('utf-8');
  console.log(text)

var convert = require('convert-rich-text');
var html = convert(delta);
console.log