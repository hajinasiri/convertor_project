var fs = require("fs");
const rtfParse = require( 'rtf-parse' );




// var f = "/Users/shahab/lighthouse/scriv/render3/render0.3.scriv"
var f = "/Users/shahab/lighthouse/scriv/render2.0/render0.2.scriv"

 f = f + '/Files/Data/FA5DE601-3097-42D2-A087-F8522006F426/content.rtf'
  var text = fs.readFileSync(f).toString('utf-8');

// rtfParse.parseString( text )
//     .then( doc => {
//         var keys = Object.keys(doc.children[0]);
//         console.log(doc.children[0].children[doc.children[0].children.length - 1].value);
//     } );



var rtf2html = require('rtf2html')
console.log(rtf2html(text))

