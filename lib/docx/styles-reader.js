exports.readStylesXml = readStylesXml;
exports.Styles = Styles;
exports.defaultStyles = new Styles({}, {});

function Styles(paragraphStyles, characterStyles, tableStyles, numberingStyles) {
    return {
        findParagraphStyleById: function(styleId) {
            return paragraphStyles[styleId];
        },
        findCharacterStyleById: function(styleId) {
            return characterStyles[styleId];
        },
        findTableStyleById: function(styleId) {
            return tableStyles[styleId];
        },
        findNumberingStyleById: function(styleId) {
            return numberingStyles[styleId];
        }
    };
}

Styles.EMPTY = new Styles({}, {}, {}, {});

function readStylesXml(root) {
    var paragraphStyles = {};
    var characterStyles = {};
    var tableStyles = {};
    var numberingStyles = {};

    var styles = {
        "paragraph": paragraphStyles,
        "character": characterStyles,
        "table": tableStyles
    };

    root.getElementsByTagName("w:style").forEach(function(styleElement) {
        var style = readStyleElement(styleElement);
        if (style.type === "numbering") {
            numberingStyles[style.styleId] = readNumberingStyleElement(styleElement);
        } else {
            var styleSet = styles[style.type];
            if (styleSet) {
                if (styleSet[style.styleId]) {
                    if (style.font) {
                        styleSet[style.styleId].font = style.font;
                    }
                } else {
                    styleSet[style.styleId] = style;
                }
            }
            styleSet[style.styleId] = style;
        }
    });

    return new Styles(paragraphStyles, characterStyles, tableStyles, numberingStyles);
}

function readStyleElement(styleElement) {
    var type = styleElement.attributes["w:type"];
    var styleId = styleElement.attributes["w:styleId"];
    var name = styleName(styleElement);
    var font;
    var bold;
    var underline;
    var italic;
    styleElement.children.forEach(function(element) {
        if (element.name && element.name === 'w:rPr' && element.children) {
            element.children.forEach(function(runProp) {
                if (runProp.name && runProp.name === 'w:rFonts' && runProp.attributes['w:ascii']) {
                    font = runProp.attributes['w:ascii'];
                }
                if (runProp.name && runProp.name === 'w:b') {
                    bold = runProp;
                }
                if (runProp.name && runProp.name === 'w:i') {
                    italic = runProp;
                }
                if (runProp.name && runProp.name === 'w:u') {
                    underline = runProp;
                }
            });
        }
    });
    return {
        type: type,
        styleId: styleId,
        name: name,
        font: font,
        bold: bold,
        underline: underline,
        italic: italic
    };
}

function styleName(styleElement) {
    var nameElement = styleElement.first("w:name");
    return nameElement ? nameElement.attributes["w:val"] : null;
}

function readNumberingStyleElement(styleElement) {
    var numId = styleElement
        .firstOrEmpty("w:pPr")
        .firstOrEmpty("w:numPr")
        .firstOrEmpty("w:numId")
        .attributes["w:val"];
    return {numId: numId};
}
