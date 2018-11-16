var file;
document.getElementById("document")
    .addEventListener("change", handleFileSelect, false);

document.getElementById("convert")
    .addEventListener("click", reconvert, false);

function reconvert() {
    var reader = new FileReader();

    reader.onload = function (loadEvent) {
        var arrayBuffer = loadEvent.target.result;
        convert(arrayBuffer);
    };

    reader.readAsArrayBuffer(file);
}
function convert(arrayBuffer) {
    var options = {};
    var optStr = document.getElementById("options").value.trim();
    console.log('optStr', optStr);
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

    var messageHtml = result.messages.map(function (message) {
        return '<li class="' + message.type + '">' + escapeHtml(message.message) + "</li>";
    }).join("");

    document.getElementById("messages").innerHTML = "<ul>" + messageHtml + "</ul>";
}

function readFileInputEventAsArrayBuffer(event, callback) {
    file = event.target.files[0];

    var reader = new FileReader();

    reader.onload = function (loadEvent) {
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
