var _ = require("underscore");

var promises = require("./promises");
var Html = require("./html");

exports.imgElement = imgElement;

function imgElement(func) {
    return function(element, messages) {
        return promises.when(func(element)).then(function(result) {
            var attributes = _.clone(result);
            if (element.altText) {
                attributes.alt = element.altText;
            }
            return [Html.freshElement("img", attributes)];
        });
    };
}

// Undocumented, but retained for backwards-compatibility with 0.3.x
exports.inline = exports.imgElement;

exports.dataUri = imgElement(function(element) {
    return element.read("base64").then(function(imageBuffer) {
        var imgAttr = {
            src: "data:" + element.contentType + ";base64," + imageBuffer
        };
        _.extendOwn(imgAttr, element.extOptions);
        if (element.extOptions.hasOwnProperty('oleB64')){
            imgAttr.src = "data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAABUAAAAVCAYAAACpF6WWAAAAAXNSR0IArs4c6QAAAARnQU1BAACxjwv8YQUAAAAJcEhZcwAADsMAAA7DAcdvqGQAAAElSURBVDhP7ZN9DYQwDMWxgAYs4AEJaMACDnCAAxSgAAMYwAEedvktvF03tgvJ5f67lzSDrn3t+lG5H+BPmsdxHG4YBlfXtauqyp/8oxemafJ3iEWWdFmWQIbjtm1Bh4i46zpvw2lxI933PUSHyEJ3fd+78zyDHUEtbqSK3rbtpYmh+3me/ZlmCSJSmyVOOYhMsq7rdfNGRGod0icJ6GXTNM2ljRGRjuMYHMi6BNnQxBwiUtUL+QTZlF5TJKW7OTBOssnVE0SkdphLWdi6P3q+dShlQXNkw7zmEJHakWIlUxCUEulFj7oPbF3tBJA5K4qOTZONysTkaH1vpDhp7xGCsF18a21pomw4uUfU3BspgJh6iZjvtHH8q76Uyk5LlvRb/IDUuReTzbnJqDLZsAAAAABJRU5ErkJggg==";
        }
        return imgAttr;
    });
});
