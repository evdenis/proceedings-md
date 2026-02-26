"use strict";
Object.defineProperty(exports, "__esModule", { value: true });
exports.xmlBuilder = exports.xmlParser = exports.xmlAttributes = exports.xmlText = exports.xmlComment = void 0;
exports.getChildTag = getChildTag;
exports.getChildTagRequired = getChildTagRequired;
exports.getTagName = getTagName;
exports.getXmlTextTag = getXmlTextTag;
exports.getAttributesXml = getAttributesXml;
exports.getParagraphText = getParagraphText;
const fast_xml_parser_1 = require("fast-xml-parser");
exports.xmlComment = "__comment__";
exports.xmlText = "__text__";
exports.xmlAttributes = ":@";
exports.xmlParser = new fast_xml_parser_1.XMLParser({
    ignoreAttributes: false,
    alwaysCreateTextNode: true,
    attributeNamePrefix: "",
    preserveOrder: true,
    trimValues: false,
    commentPropName: exports.xmlComment,
    textNodeName: exports.xmlText
});
exports.xmlBuilder = new fast_xml_parser_1.XMLBuilder({
    ignoreAttributes: false,
    attributeNamePrefix: "",
    preserveOrder: true,
    commentPropName: exports.xmlComment,
    textNodeName: exports.xmlText
});
function getChildTag(styles, name) {
    for (let child of styles) {
        if (child[name]) {
            return child;
        }
    }
    return undefined;
}
function getChildTagRequired(styles, name) {
    let result = getChildTag(styles, name);
    if (result === undefined) {
        throw new Error(`Required child tag '${name}' not found`);
    }
    return result;
}
function getTagName(tag) {
    for (let key of Object.getOwnPropertyNames(tag)) {
        if (key === exports.xmlAttributes)
            continue;
        return key;
    }
    return undefined;
}
function getXmlTextTag(text) {
    let result = {};
    result[exports.xmlText] = text;
    return result;
}
function getAttributesXml(attributes) {
    let result = {};
    result[exports.xmlAttributes] = attributes;
    return result;
}
function getRawText(tag) {
    let result = "";
    let tagName = getTagName(tag);
    if (tagName === exports.xmlText) {
        result += tag[exports.xmlText];
    }
    if (Array.isArray(tag[tagName])) {
        for (let child of tag[tagName]) {
            result += getRawText(child);
        }
    }
    return result;
}
function getParagraphText(paragraph) {
    let result = "";
    if (paragraph["w:t"]) {
        result += getRawText(paragraph);
    }
    for (let name of Object.getOwnPropertyNames(paragraph)) {
        if (name === exports.xmlAttributes) {
            continue;
        }
        if (Array.isArray(paragraph[name])) {
            for (let child of paragraph[name]) {
                result += getParagraphText(child);
            }
        }
    }
    return result;
}
