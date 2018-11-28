/* eslint-env browser */


function getFormulaXml(jsonStr){

    var xml = document.implementation.createDocument("", "", null);
    
    var doc = xml.createElement('w:document');
    
    ['xmlns:wpc="http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas"',
        'xmlns:cx="http://schemas.microsoft.com/office/drawing/2014/chartex"',
        'xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"',
        'xmlns:o="urn:schemas-microsoft-com:office:office"',
        'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"',
        'xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math"',
        'xmlns:v="urn:schemas-microsoft-com:vml"',
        'xmlns:wp14="http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing"',
        'xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"',
        'xmlns:w10="urn:schemas-microsoft-com:office:word"',
        'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"',
        'xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"',
        'xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml"',
        'xmlns:w16se="http://schemas.microsoft.com/office/word/2015/wordml/symex"',
        'xmlns:wpg="http://schemas.microsoft.com/office/word/2010/wordprocessingGroup"',
        'xmlns:wpi="http://schemas.microsoft.com/office/word/2010/wordprocessingInk"',
        'xmlns:wne="http://schemas.microsoft.com/office/word/2006/wordml"',
        'xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"',
        'mc:Ignorable="w14 w15 w16se wp14"'
    ].forEach(function(line){
        var attrs = line.split('=');
        var key = attrs[0];
        var value = attrs[1].replace(/\"/g, '');
        doc.setAttribute(key, value);
    });
    
    function json2xml(json, parent) {
        var thisEle = xml.createElement(json.name);
        if (json.attributes) {
            Object.entries(json.attributes).forEach(function(ele) {
                thisEle.setAttribute(ele[0], ele[1]);
            });
        }
    
        if (json.type === 'text') {
            parent.textContent = json.value;
            return;
        } else if (json.children) {
            json.children.forEach(function(child) {
                json2xml(child, thisEle);
            });
        }
    
        parent.appendChild(thisEle);
    }
    
    json2xml(JSON.parse(jsonStr), doc);
    xml.appendChild(doc);

    return xml;
}


function loadXMLDoc(filename) {

    var xhttp = new XMLHttpRequest();

    xhttp.open("GET", filename, false);
    xhttp.send("");
    return xhttp.responseXML;
}


var xsl = loadXMLDoc("OMML2MML.XSL");

function getMathMl(Ommlxml){
    var xml = getFormulaXml(Ommlxml);
    if (document.implementation && document.implementation.createDocument) {
        var xsltProcessor = new XSLTProcessor();
        xsltProcessor.importStylesheet(xsl);
    
        var resultDocument = xsltProcessor.transformToFragment(xml, document);
        return resultDocument;
    }
}

