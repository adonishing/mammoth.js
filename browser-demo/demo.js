/* eslint-env browser */
/* global mammoth :true */
/* global MathJax :true */
/* global getMathMl :true */

var file;
document.getElementById("document")
    .addEventListener("change", handleFileSelect, false);

document.getElementById("convert")
    .addEventListener("click", reconvert, false);

function reconvert() {
    var reader = new FileReader();

    reader.onload = function(loadEvent) {
        var arrayBuffer = loadEvent.target.result;
        convert(arrayBuffer);
    };

    reader.readAsArrayBuffer(file);
}
function convert(arrayBuffer) {
    var options = {};
    var optStr = document.getElementById("options").value.trim();
    if (optStr) {
        options = JSON.parse(optStr);
    }
    
    mammoth.convertToHtml({arrayBuffer: arrayBuffer}, options)
        .then(displayResult)
        .done();
}
function handleFileSelect(event) {
    readFileInputEventAsArrayBuffer(event, convert);
}

function displayResult(result) {
    document.getElementById("output").innerHTML = result.value;

    var messageHtml = result.messages.map(function(message) {
        return '<li class="' + message.type + '">' + escapeHtml(message.message) + "</li>";
    }).join("");

    document.getElementById("messages").innerHTML = "<ul>" + messageHtml + "</ul>";
}

function readFileInputEventAsArrayBuffer(event, callback) {
    file = event.target.files[0];

    var reader = new FileReader();

    reader.onload = function(loadEvent) {
        var arrayBuffer = loadEvent.target.result;
        callback(arrayBuffer);
    };

    reader.readAsArrayBuffer(file);
}

function escapeHtml(value) {
    return value
        .replace(/&/g, '&amp;')
        .replace(/"/g, '&quot;')
        .replace(/</g, '&lt;')
        .replace(/>/g, '&gt;');
}

/**
 *  未处理的omml公式放在一个img元素里，元素的onload会调用该函数。
 *  @param ommlImg 存放omml的img元素
 */
function showOneFormula(ommlImg) {
    var ommlXml = ommlImg.getAttribute('omath');
    var mathMlEle = getMathMl(ommlXml);
    var parent = ommlImg.parentElement;
    parent.replaceChild(mathMlEle, ommlImg);
    MathJax.Hub.Queue(["Typeset", MathJax.Hub, parent]);
}
