exports.DocumentXmlReader = DocumentXmlReader;

var documents = require("../documents");
var Result = require("../results").Result;


function DocumentXmlReader(options) {
    var bodyReader = options.bodyReader;

    function convertXmlToDocument(element) {
        var body = element.first("w:body");

        if (body == null) {
            throw new Error("Could not find the body element: are you sure this is a docx file?");
        }

        var parsedComments = options.comments ? options.comments.map(function(comment) {
            // TODO how does it work if there is multiple paragraphs
            if (comment.body && comment.body[0].paraId && options.commentsExtended) {
                var exComment = options.commentsExtended.find(function(ex) {
                    return ex.paraId === comment.body[0].paraId;
                });
                if (exComment) {
                    comment.done = exComment.done;
                }
            }
            return comment;
        }) : undefined;

        var result = bodyReader.readXmlElements(body.children)
            .map(function(children) {
                return new documents.Document(children, {
                    notes: options.notes,
                    comments: parsedComments
                });
            });
        return new Result(result.value, result.messages);
    }

    return {
        convertXmlToDocument: convertXmlToDocument
    };
}
