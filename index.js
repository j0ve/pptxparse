var fs = require('fs');
var unzip = require('unzip');
var path = require('path');
var convert = require('xml-js');
var stream = require('stream');
var util = require('util');
function consolelog(data) {
    console.log(util.inspect(data, {showHidden: false, depth: null}));
}
// test comment
function getrangedata(rangeobj) {
    var range = {};
    if(rangeobj["a:rPr"]["a:latin"]) range.font = rangeobj["a:rPr"]["a:latin"]._attributes.typeface;
    if(Object.keys(rangeobj["a:t"]).length === 0) return false;
    else range.text = rangeobj["a:t"]._text;
    return range;
}

function getrangesdata(rangesobj) {
    var ranges = [];
    var range = {};
    // if the text ranges in the paragraph are also an aray, iterate em
    if(Object.prototype.toString.call(rangesobj) === '[object Array]') {
        ranges = [];
        for(var k = 0; k < rangesobj.length; k++) {                                     
            range = getrangedata(rangesobj[k]);
            if(range) ranges.push(range);
        }
    } else {
            range = getrangedata(rangesobj);
            if(range) ranges.push(range);
    }    
    return ranges;
}

function getparagraphdata(paragraphobj) {
    var paragraph = {};
    if(paragraphobj["a:r"]) {
        paragraph.ranges = getrangesdata(paragraphobj["a:r"]);
        return paragraph;
    }
    else {
        return false;
    }

}

function getparagraphsdata(paragraphsobj) {
    var paragraphs = [];
    var oaragraph = {};
    // if its an array, iterate em
    if(Object.prototype.toString.call(paragraphsobj) === '[object Array]') {
        for(var j = 0; j < paragraphsobj.length; j++) {
            paragraph = getparagraphdata(paragraphsobj[j]);
            if(paragraph) paragraphs.push(paragraph);
        }
    }
    else {
        paragraph = getparagraphdata(paragraphsobj);
        if(paragraph) paragraphs.push(paragraph);                    
    }    
    if(paragraphs.length === 0) return false;
    else return paragraphs;
}

function gettextbodydata(txbodyobj) {
    var textbodydata = {};
    var paragraphs = getparagraphsdata(txbodyobj["a:p"]);
    if(paragraphs) textbodydata.paragraphs = paragraphs;
    return textbodydata;
}

function getshapedata(shapeobj) {
    // empty shape object
    var shape = {};
    // get some attributes from the non visible shape properties object
    var shapeattributes = shapeobj["p:nvSpPr"]["p:cNvPr"]._attributes;
    // add some more attributes
    if(shapeattributes,shapeobj["p:nvSpPr"]["p:nvPr"]["p:ph"]) {
        Object.assign(shapeattributes,shapeobj["p:nvSpPr"]["p:nvPr"]["p:ph"]._attributes);
    }
    if(shapeattributes,shapeobj["p:spPr"]["a:xFrm"]) {
        Object.assign(shapeattributes,shapeobj["p:spPr"]["a:xfrm"]["a:off"]._attributes);
        Object.assign(shapeattributes,shapeobj["p:spPr"]["a:xfrm"]["a:ext"]._attributes);
    }    
    shape.attributes = shapeattributes;        
    // see if there is a text body in this shape
    if(shapeobj["p:txBody"]) {
        shape.textbodydata = gettextbodydata(shapeobj["p:txBody"]);
    }
    return shape;
}
function getpicdata(picobj) {
    // empty pic object
    var pic = {};
    // get some attributes from the non visible pic properties object
    var picattributes = picobj["p:nvPicPr"]["p:cNvPr"]._attributes;
    picattributes = Object.assign(picattributes, picobj["p:blipFill"]["a:blip"]._attributes);
    picattributes = Object.assign(picattributes, picobj["p:spPr"]["a:xfrm"]["a:off"]._attributes);
    picattributes = Object.assign(picattributes, picobj["p:spPr"]["a:xfrm"]["a:ext"]._attributes); 
    pic.attributes = picattributes;
    return pic;       
}
function getslidedata(slideobj) {
    var slide = {};
    // the pptx slide.xml format starts with a xml declaration
    // then a root sld element with some attributes
    // as its child a csld element with common slide data
    // this has a shape tree as a child which we're gonna access for our data (p:spTree)
    // the sptree has non visible and visible group shape properties, and a shapes array (p:sp)
    // that MIGHT be "all" we need to get the content from the slide file
    slide.shapetree = getshapegroupdata(slideobj['p:sld']['p:cSld']['p:spTree']);
    return slide;
}
function getpresentationdata(presentationobj) {
    var presentationdata = {};
    presentationdata.slidesizeX = presentationobj["p:presentation"]["p:sldSz"]._attributes.cx;
    presentationdata.slidesizeY = presentationobj["p:presentation"]["p:sldSz"]._attributes.cy;
    return presentationdata;
    //console.log(util.inspect(presentationobj, {showHidden: false, depth: null}));
}
// dit was eerst de functie om door een shapetree heen te lopen, maar aangezien een shapegroup precies hetzelfde opgebouwd is, en bovendien zelf ook weer shapegroups kan bevatten
// hiervan een recursieve algemene functie gemaakt om door shapegroups (waarvan de root shapetree is) te lopen.
function getshapegroupdata(shapegroupobj) {
    var shapegroupdata = {};
    var shapegroups = [];
    var shapes = [];  
    var pics = [];
    if(shapegroupobj["p:grpSp"]) {
        if(Object.prototype.toString.call(shapegroupobj["p:grpSp"]) === '[object Array]') {
            for(var i = 0; i < shapegroupobj["p:grpSp"].length; i ++) {
                shapegroups.push(getshapegroupdata(shapegroupobj["p:grpSp"][i]));
            }
        }
        else {
            shapegroups.push(getshapegroupdata(shapegroupobj["p:grpSp"]));
        }
    }
    if(shapegroupobj["p:sp"]) {
        if(Object.prototype.toString.call(shapegroupobj["p:sp"]) === '[object Array]') {
            for(var j = 0; j < shapegroupobj["p:sp"].length; j ++) {
                shapes.push(getshapedata(shapegroupobj["p:sp"][j]));
            }
        }
        else {
            shapes.push(getshapedata(shapegroupobj["p:sp"]));
        }        
    }
    if(shapegroupobj['p:pic']) {
        if(Object.prototype.toString.call(shapegroupobj["p:pic"]) === '[object Array]') {
            for(var k = 0; k < shapegroupobj["p:pic"].length; k ++) {
                pics.push(getpicdata(shapegroupobj["p:pic"][k]));
            }
        }
        else {
            pics.push(getpicdata(shapegroupobj["p:pic"]));
        }         
    }
    shapegroupdata.shapegroups = shapegroups;
    shapegroupdata.shapes = shapes;   
    shapegroupdata.pics = pics;
    return shapegroupdata; 
}
function getslidelayoutdata(slidelayoutobj) {
    var slidelayoutdata = {};
    slidelayoutdata.shapetree = getshapegroupdata(slidelayoutobj["p:sldLayout"]["p:cSld"]["p:spTree"]);

    return slidelayoutdata;
}
function getslidereldata(sliderelobj) {

}
var slides = [];
var presentationdata = {};
var slidelayouts = [];
var sliderels = [];
//fs.createReadStream('SlimmerIQuiz_voorronde_en_antwoorden.pptx')
var readstream = fs.createReadStream('Quiz.pptx')
    .on('end', function () {
        consolelog(slidelayouts);
    })
    .pipe(unzip.Parse())
    .on('entry', function (entry) {
    var fileName = entry.path;
    var type = entry.type; // 'Directory' or 'File' 
    var size = entry.size;
    var slidenr = 0;
    var slidelayoutnr = 0;
    var sliderelnr = 0;
    var filetype = '';
    // prepare var for storing slide content
    var contentbuffer = "";
    // make a new stream for writing the reading stream to
    var ws = new stream.Writable();
    // whenever a chunk is read, add it to the variable
    ws.write = function(chunk, encoding, callback) {
        contentbuffer += chunk;
    };    
    // once the end of the file is reached..
    ws.end = function(chunk, encoding, callback) {
        // write the remainder of the file to the content holder var
        if(typeof chunk !== 'undefined') contentbuffer += chunk;
        if(filetype === "slide") {
            // and then convert all of it to a json object
            var slidejs = convert.xml2js(contentbuffer, {compact: true, spaces: 4});              
            slides.push({"id": slidenr, "data" :getslidedata(slidejs)});   
        }
        if(filetype === "presentationdata") {
            var presentationjs = convert.xml2js(contentbuffer, {compact: true, spaces: 4});
            presentationdata = getpresentationdata(presentationjs);
        }
        if(filetype === "slidelayout") {
            var slidelayoutjs = convert.xml2js(contentbuffer, {compact: true, spaces: 4});
            slidelayouts.push({"id": slidelayoutnr, "data": getslidelayoutdata(slidelayoutjs)});
        }
        if(filetype === "sliderel") {
            var slidereljs = convert.xml2js(contentbuffer, {compact: true, spaces: 4});
            sliderels.push({"id": sliderelnr, "data": getslidereldata(slidereljs)});
        }        
    };
    if(path.dirname(fileName) === 'ppt/slides') {
        // probably got a slide here. maybe some more checks?
        // also, gotta get the slide number from the filename probable, or maybe slide order is in the main files
        // but for now, gotta parse the slide at hand
        // get the number from the filename
        filetype = 'slide';
        slidenr = parseInt(path.basename(fileName).match(/\d+/)[0]);       
        entry.pipe(ws);
        
    } else if(path.basename(fileName)=="presentation.xml") {
        filetype = 'presentationdata';
        entry.pipe(ws);
    } else if(path.dirname(fileName) === 'ppt/slideLayouts') {
        filetype = 'slidelayout';
        slidelayoutnr = parseInt(path.basename(fileName).match(/\d+/)[0]);    
        entry.pipe(ws);
    } else if(path.dirname(fileName) === 'ppt/slides/_rels') {
        filetype = 'sliderel';
        sliderelnr = parseInt(path.basename(fileName).match(/\d+/)[0]);    
        entry.pipe(ws);
    } else {
        entry.autodrain();
    }
});

