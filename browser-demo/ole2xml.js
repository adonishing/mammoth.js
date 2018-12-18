/* eslint-env browser */


// let oleXsl = loadXMLDoc("transform.xsl");


function converOle(oleElement){
    let oleB64 = oleElement.getAttribute('oleb64').replace('data:image/png;base64,', '');
    let oleId = oleElement.getAttribute('id');
    let xhr = new XMLHttpRequest();
    xhr.open("post", "ole2mml?id=" + oleId, true);
    xhr.setRequestHeader("Content-Type", "text/plain");
    xhr.onreadystatechange = function() {
        if (xhr.readyState == 4 && xhr.status == 200) {
            let mathMlEle = xhr.responseXML.childNodes[0];
            console.log(oleId);

            if(mathMlEle.getAttribute('display') === 'block'){
                mathMlEle.setAttribute('display', 'inline')
            }else if(mathMlEle.getAttribute('display') === 'inline'){
                mathMlEle.setAttribute('display', 'block')
            }
            let parent = oleElement.parentElement;
            parent.replaceChild(mathMlEle, oleElement);
            MathJax.Hub.Queue(["Typeset", MathJax.Hub, parent]);
        }
    };
    xhr.send(oleB64);
}

