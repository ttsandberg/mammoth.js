var documents = require("../documents");
var Result = require("../results").Result;

function createCommentsExtendedReader(bodyReader) {
    function readCommentsExtendedXml(element) {
        return Result.combine(element.getElementsByTagName("{http://schemas.microsoft.com/office/word/2012/wordml}commentEx")
            .map(readCommentExtendedElement));
    }

    function readCommentExtendedElement(element) {
        var paraId = element.attributes["{http://schemas.microsoft.com/office/word/2012/wordml}paraId"];

        function readOptionalAttribute(name) {
            return (element.attributes[name] || "").trim() || null;
        }

        return {value: [
            documents.commentExtended({
                paraId: paraId,
                done: readOptionalAttribute("{http://schemas.microsoft.com/office/word/2012/wordml}done")
            })
        ], messages: []};
    }

    return readCommentsExtendedXml;
}

exports.createCommentsExtendedReader = createCommentsExtendedReader;
