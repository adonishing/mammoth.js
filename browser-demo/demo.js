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
    options.convertImage = mammoth.images.imgElement(function(image) {
        /*if (image.extOptions.onload === "converOle(this)"){
            return image.read("base64").then(function(imageBuffer) {
                return {
                    src: "data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAABUAAAAVCAYAAACpF6WWAAAAAXNSR0IArs4c6QAAAARnQU1BAACxjwv8YQUAAAAJcEhZcwAADsMAAA7DAcdvqGQAAAElSURBVDhP7ZN9DYQwDMWxgAYs4AEJaMACDnCAAxSgAAMYwAEedvktvF03tgvJ5f67lzSDrn3t+lG5H+BPmsdxHG4YBlfXtauqyp/8oxemafJ3iEWWdFmWQIbjtm1Bh4i46zpvw2lxI933PUSHyEJ3fd+78zyDHUEtbqSK3rbtpYmh+3me/ZlmCSJSmyVOOYhMsq7rdfNGRGod0icJ6GXTNM2ljRGRjuMYHMi6BNnQxBwiUtUL+QTZlF5TJKW7OTBOssnVE0SkdphLWdi6P3q+dShlQXNkw7zmEJHakWIlUxCUEulFj7oPbF3tBJA5K4qOTZONysTkaH1vpDhp7xGCsF18a21pomw4uUfU3BspgJh6iZjvtHH8q76Uyk5LlvRb/IDUuReTzbnJqDLZsAAAAABJRU5ErkJggg==",
                    oleb64: imageBuffer,
                    id: image.extOptions.id,
                    onload: "converOle(this)"
                };
            });
        }*/
        
        return image.read("base64").then(function(imageBuffer) {
            var imgAttr = {
                src: "data:" + image.contentType + ";base64," + imageBuffer
            };
            console.log(image.extOptions)
            Object.assign(imgAttr, image.extOptions);

            return imgAttr;
        });
    });
    mammoth.convertToHtml({arrayBuffer: arrayBuffer}, options, getMathMl)
        .then(displayResult)
        .done();

}
function handleFileSelect(event) {
    readFileInputEventAsArrayBuffer(event, convert);
}

function displayResult(result) {
    console.log(result.value)
    document.getElementById("output").innerHTML = result.value;
    MathJax.Hub.Typeset();
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
function convertOMath(ommlImg) {
    var ommlXml = ommlImg.getAttribute('omath');
    var mathMlEle = getMathMl(ommlXml);
    var parent = ommlImg.parentElement;
    parent.replaceChild(mathMlEle, ommlImg);
    MathJax.Hub.Queue(["Typeset", MathJax.Hub, parent]);
}
