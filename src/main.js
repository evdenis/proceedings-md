"use strict";
var __createBinding = (this && this.__createBinding) || (Object.create ? (function(o, m, k, k2) {
    if (k2 === undefined) k2 = k;
    var desc = Object.getOwnPropertyDescriptor(m, k);
    if (!desc || ("get" in desc ? !m.__esModule : desc.writable || desc.configurable)) {
      desc = { enumerable: true, get: function() { return m[k]; } };
    }
    Object.defineProperty(o, k2, desc);
}) : (function(o, m, k, k2) {
    if (k2 === undefined) k2 = k;
    o[k2] = m[k];
}));
var __setModuleDefault = (this && this.__setModuleDefault) || (Object.create ? (function(o, v) {
    Object.defineProperty(o, "default", { enumerable: true, value: v });
}) : function(o, v) {
    o["default"] = v;
});
var __importStar = (this && this.__importStar) || (function () {
    var ownKeys = function(o) {
        ownKeys = Object.getOwnPropertyNames || function (o) {
            var ar = [];
            for (var k in o) if (Object.prototype.hasOwnProperty.call(o, k)) ar[ar.length] = k;
            return ar;
        };
        return ownKeys(o);
    };
    return function (mod) {
        if (mod && mod.__esModule) return mod;
        var result = {};
        if (mod != null) for (var k = ownKeys(mod), i = 0; i < k.length; i++) if (k[i] !== "default") __createBinding(result, mod, k[i]);
        __setModuleDefault(result, mod);
        return result;
    };
})();
Object.defineProperty(exports, "__esModule", { value: true });
const path = __importStar(require("path"));
const fs = __importStar(require("fs"));
const JSZip = __importStar(require("jszip"));
const xml_helpers_1 = require("./xml-helpers");
const pandoc_helpers_1 = require("./pandoc-helpers");
const references_1 = require("./references");
const properDocXmlns = new Map([
    ["xmlns:w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main"],
    ["xmlns:m", "http://schemas.openxmlformats.org/officeDocument/2006/math"],
    ["xmlns:r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships"],
    ["xmlns:o", "urn:schemas-microsoft-com:office:office"],
    ["xmlns:v", "urn:schemas-microsoft-com:vml"],
    ["xmlns:w10", "urn:schemas-microsoft-com:office:word"],
    ["xmlns:a", "http://schemas.openxmlformats.org/drawingml/2006/main"],
    ["xmlns:pic", "http://schemas.openxmlformats.org/drawingml/2006/picture"],
    ["xmlns:wp", "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"],
]);
const languages = ["ru", "en"];
// Numbering definition IDs from the ISP RAS template
const NUM_ID_ORDERED = "33";
const NUM_ID_BULLET = "43";
const NUM_ID_BIBLIOGRAPHY = "80";
// Starting counter for dynamically assigned list numIds
const DYNAMIC_NUM_ID_START = 10000;
// Spacing values in twentieths of a point
const SPACING_BEFORE_FIRST_AUTHOR = "480";
const SPACING_BEFORE_FIRST_ORG = "60";
function getStyleCrossReferences(styles) {
    let result = [];
    for (let style of (0, xml_helpers_1.getChildTagRequired)(styles, "w:styles")["w:styles"]) {
        if (!style["w:style"])
            continue;
        result.push(style[xml_helpers_1.xmlAttributes]);
        let basedOnTag = (0, xml_helpers_1.getChildTag)(style["w:style"], "w:basedOn");
        if (basedOnTag)
            result.push(basedOnTag[xml_helpers_1.xmlAttributes]);
        let linkTag = (0, xml_helpers_1.getChildTag)(style["w:style"], "w:link");
        if (linkTag)
            result.push(linkTag[xml_helpers_1.xmlAttributes]);
        let nextTag = (0, xml_helpers_1.getChildTag)(style["w:style"], "w:next");
        if (nextTag)
            result.push(nextTag[xml_helpers_1.xmlAttributes]);
    }
    return result;
}
function getDocStyleUseReferences(doc, result = [], met = new Set()) {
    if (!doc || typeof doc !== "object" || met.has(doc)) {
        return result;
    }
    met.add(doc);
    if (Array.isArray(doc)) {
        for (let child of doc) {
            result = getDocStyleUseReferences(child, result, met);
        }
    }
    let tagName = (0, xml_helpers_1.getTagName)(doc);
    if (tagName === "w:pStyle" || tagName === "w:rStyle") {
        result.push(doc[xml_helpers_1.xmlAttributes]);
    }
    result = getDocStyleUseReferences(doc[tagName], result, met);
    return result;
}
function extractStyleDefs(styles, usedStyles) {
    let result = [];
    for (let style of (0, xml_helpers_1.getChildTagRequired)(styles, "w:styles")["w:styles"]) {
        if (!style["w:style"])
            continue;
        if (usedStyles.has(style[xml_helpers_1.xmlAttributes]["w:styleId"])) {
            let copy = JSON.parse(JSON.stringify(style));
            result.push(copy);
        }
    }
    return result;
}
function patchStyleUseReferences(doc, styles, map) {
    let docReferences = getDocStyleUseReferences(doc);
    let crossReferences = getStyleCrossReferences(styles);
    for (let ref of docReferences.concat(crossReferences)) {
        if (ref["w:val"] && map.has(ref["w:val"])) {
            ref["w:val"] = map.get(ref["w:val"]);
        }
    }
}
function getUsedStyles(doc) {
    let references = getDocStyleUseReferences(doc);
    let set = new Set();
    for (let ref of references) {
        set.add(ref["w:val"]);
    }
    return set;
}
function populateStyles(styles, table) {
    for (let styleId of styles) {
        let style = table.get(styleId);
        if (!style) {
            throw new Error("Style id " + styleId + " not found");
        }
        let basedOnTag = (0, xml_helpers_1.getChildTag)(style["w:style"], "w:basedOn");
        if (basedOnTag)
            styles.add(basedOnTag[xml_helpers_1.xmlAttributes]["w:val"]);
        let linkTag = (0, xml_helpers_1.getChildTag)(style["w:style"], "w:link");
        if (linkTag)
            styles.add(linkTag[xml_helpers_1.xmlAttributes]["w:val"]);
        let nextTag = (0, xml_helpers_1.getChildTag)(style["w:style"], "w:next");
        if (nextTag)
            styles.add(nextTag[xml_helpers_1.xmlAttributes]["w:val"]);
    }
}
function getUsedStylesDeep(doc, styleTable, requiredStyles = []) {
    let usedStyles = getUsedStyles(doc);
    for (let requiredStyle of requiredStyles) {
        usedStyles.add(requiredStyle);
    }
    let prevSize;
    do {
        prevSize = usedStyles.size;
        populateStyles(usedStyles, styleTable);
    } while (usedStyles.size > prevSize);
    return usedStyles;
}
function getStyleTable(styles) {
    let table = new Map();
    for (let style of (0, xml_helpers_1.getChildTagRequired)(styles, "w:styles")["w:styles"]) {
        if (!style["w:style"])
            continue;
        table.set(style[xml_helpers_1.xmlAttributes]["w:styleId"], style);
    }
    return table;
}
function getStyleIdsByNameFromDefs(styles) {
    let table = new Map();
    for (let style of styles) {
        if (!style["w:style"])
            continue;
        let nameNode = (0, xml_helpers_1.getChildTag)(style["w:style"], "w:name");
        if (nameNode) {
            table.set(nameNode[xml_helpers_1.xmlAttributes]["w:val"], style[xml_helpers_1.xmlAttributes]["w:styleId"]);
        }
    }
    return table;
}
function appendStyles(target, defs) {
    let styles = (0, xml_helpers_1.getChildTagRequired)(target, "w:styles")["w:styles"];
    for (let def of defs) {
        styles.push(def);
    }
}
function applyListStyles(doc, styles) {
    let stack = [];
    let currentState = undefined;
    let met = new Set();
    let newStyles = new Map();
    let lastId = DYNAMIC_NUM_ID_START;
    const walk = (node) => {
        if (!node || typeof node !== "object" || met.has(node)) {
            return;
        }
        met.add(node);
        for (let key of Object.getOwnPropertyNames(node)) {
            walk(node[key]);
            if (key === "w:pPr" && currentState) {
                // Remove any old pStyle and add our own
                for (let i = 0; i < node[key].length; i++) {
                    if (node[key][i]["w:pStyle"]) {
                        node[key].splice(i, 1);
                        i--;
                    }
                }
                node[key].unshift({
                    "w:pStyle": {},
                    ...(0, xml_helpers_1.getAttributesXml)({ "w:val": styles[currentState.listStyle].styleName })
                });
                // Set explicit spacing on bullet list items to match
                // the official template (before="60" after="60")
                if (currentState.listStyle === "BulletList") {
                    for (let i = 0; i < node[key].length; i++) {
                        if (node[key][i]["w:spacing"]) {
                            node[key].splice(i, 1);
                            i--;
                        }
                    }
                    node[key].push({
                        "w:spacing": [],
                        ...(0, xml_helpers_1.getAttributesXml)({ "w:before": "60", "w:after": "60" })
                    });
                }
            }
            if (key === "w:numId" && currentState) {
                node[xml_helpers_1.xmlAttributes]["w:val"] = String(currentState.numId);
            }
            if (key === xml_helpers_1.xmlComment) {
                let commentValue = node[key][0][xml_helpers_1.xmlText];
                for (let mode of ["OrderedList", "BulletList"]) {
                    if (commentValue.includes(`ListMode ${mode}`)) {
                        stack.push(currentState);
                        currentState = { numId: lastId++, listStyle: mode };
                        newStyles.set(String(currentState.numId), styles[currentState.listStyle].numId);
                    }
                }
                if (commentValue.includes("ListMode None")) {
                    currentState = stack.pop();
                }
            }
        }
    };
    walk(doc);
    return newStyles;
}
function removeCollidedStyles(styles, collisions) {
    let newContents = [];
    for (let style of (0, xml_helpers_1.getChildTagRequired)(styles, "w:styles")["w:styles"]) {
        if (!style["w:style"] || !collisions.has(style[xml_helpers_1.xmlAttributes]["w:styleId"])) {
            newContents.push(style);
        }
    }
    (0, xml_helpers_1.getChildTagRequired)(styles, "w:styles")["w:styles"] = newContents;
}
function copyStyleSection(source, target, tagName) {
    let sourceStyles = (0, xml_helpers_1.getChildTagRequired)(source, "w:styles")["w:styles"];
    let targetStyles = (0, xml_helpers_1.getChildTagRequired)(target, "w:styles")["w:styles"];
    let sourceSection = (0, xml_helpers_1.getChildTagRequired)(sourceStyles, tagName);
    let targetSection = (0, xml_helpers_1.getChildTagRequired)(targetStyles, tagName);
    targetSection[tagName] = JSON.parse(JSON.stringify(sourceSection[tagName]));
    if (sourceSection[xml_helpers_1.xmlAttributes]) {
        targetSection[xml_helpers_1.xmlAttributes] = JSON.parse(JSON.stringify(sourceSection[xml_helpers_1.xmlAttributes]));
    }
}
async function copyFile(source, target, filePath) {
    target.file(filePath, await source.file(filePath).async("arraybuffer"));
}
function addNewNumberings(targetNumberingParsed, newListStyles) {
    let numberingTag = (0, xml_helpers_1.getChildTagRequired)(targetNumberingParsed, "w:numbering")["w:numbering"];
    // Build numId → abstractNumId lookup from existing entries
    let numIdToAbstractNumId = new Map();
    for (let entry of numberingTag) {
        if (entry["w:num"]) {
            let numId = entry[xml_helpers_1.xmlAttributes]["w:numId"];
            for (let child of entry["w:num"]) {
                if (child["w:abstractNumId"]) {
                    numIdToAbstractNumId.set(numId, child[xml_helpers_1.xmlAttributes]["w:val"]);
                }
            }
        }
    }
    // <w:num w:numId="newNum">
    //   <w:abstractNumId w:val="abstractNumId"/>
    // </w:num>
    for (let [newNum, oldNum] of newListStyles) {
        let abstractNumId = numIdToAbstractNumId.get(oldNum) || oldNum;
        let overrides = [];
        for (let i = 0; i < 9; i++) {
            overrides.push({
                "w:lvlOverride": [{
                        "w:startOverride": [],
                        ...(0, xml_helpers_1.getAttributesXml)({ "w:val": "1" })
                    }],
                ...(0, xml_helpers_1.getAttributesXml)({ "w:ilvl": String(i) })
            });
        }
        numberingTag.push({
            "w:num": [{
                    "w:abstractNumId": [],
                    ...(0, xml_helpers_1.getAttributesXml)({ "w:val": abstractNumId })
                }, ...overrides],
            ...(0, xml_helpers_1.getAttributesXml)({ "w:numId": newNum })
        });
    }
}
function addContentType(contentTypes, partName, contentType) {
    let typesTag = (0, xml_helpers_1.getChildTagRequired)(contentTypes, "Types")["Types"];
    typesTag.push({
        "Override": [],
        ...(0, xml_helpers_1.getAttributesXml)({
            "PartName": partName,
            "ContentType": contentType
        })
    });
}
function transferRels(source, target) {
    let sourceRels = (0, xml_helpers_1.getChildTagRequired)(source, "Relationships")["Relationships"];
    let targetRels = (0, xml_helpers_1.getChildTagRequired)(target, "Relationships")["Relationships"];
    let presentIds = new Map();
    let idMap = new Map();
    for (let rel of targetRels) {
        presentIds.set(rel[xml_helpers_1.xmlAttributes]["Target"], rel[xml_helpers_1.xmlAttributes]["Id"]);
    }
    let newIdCounter = 0;
    for (let rel of sourceRels) {
        if (presentIds.has(rel[xml_helpers_1.xmlAttributes]["Target"])) {
            idMap.set(rel[xml_helpers_1.xmlAttributes]["Id"], presentIds.get(rel[xml_helpers_1.xmlAttributes]["Target"]));
        }
        else {
            let newId = "template-id-" + (newIdCounter++);
            let relCopy = JSON.parse(JSON.stringify(rel));
            relCopy[xml_helpers_1.xmlAttributes]["Id"] = newId;
            targetRels.push(relCopy);
            idMap.set(rel[xml_helpers_1.xmlAttributes]["Id"], newId);
        }
    }
    return idMap;
}
function replaceInlineTemplate(body, template, value) {
    if (value === "@none") {
        let i = findParagraphWithPattern(body, template, 0);
        while (i !== null) {
            body.splice(i, 1);
            i = findParagraphWithPattern(body, template, i);
        }
    }
    else {
        replaceStringTemplate(body, template, value);
    }
}
function replaceStringTemplate(tag, template, value) {
    if (Array.isArray(tag)) {
        for (let child of tag) {
            replaceStringTemplate(child, template, value);
        }
        return;
    }
    let tagName = (0, xml_helpers_1.getTagName)(tag);
    if (tagName === xml_helpers_1.xmlText) {
        tag[xml_helpers_1.xmlText] = String(tag[xml_helpers_1.xmlText]).replace(template, value);
    }
    else if (typeof tag[tagName] === "object") {
        replaceStringTemplate(tag[tagName], template, value);
    }
}
function findParagraphWithPattern(body, pattern, startIndex = 0) {
    for (let i = startIndex; i < body.length; i++) {
        let text = (0, xml_helpers_1.getParagraphText)(body[i]);
        if (!text.includes(pattern)) {
            continue;
        }
        return i;
    }
    return null;
}
function findParagraphWithPatternStrict(body, pattern, startIndex = 0) {
    let paragraphIndex = findParagraphWithPattern(body, pattern, startIndex);
    if (paragraphIndex === null) {
        throw new Error(`The template document should have pattern ${pattern}`);
    }
    let text = (0, xml_helpers_1.getParagraphText)(body[paragraphIndex]);
    if (text !== pattern) {
        throw new Error(`The ${pattern} pattern should be the only text of the paragraph`);
    }
    return paragraphIndex;
}
function templateReplaceBodyContents(templateBody, body) {
    let paragraphIndex = findParagraphWithPatternStrict(templateBody, "{{{body}}}");
    templateBody.splice(paragraphIndex, 1, ...body);
}
function clearParagraphContents(paragraph) {
    let contents = paragraph["w:p"];
    for (let i = 0; i < contents.length; i++) {
        let tagName = (0, xml_helpers_1.getTagName)(contents[i]);
        if (tagName === "w:r") {
            contents.splice(i, 1);
            i--;
        }
    }
}
function getSuperscriptTextStyle() {
    return [
        { "w:i": [], ...(0, xml_helpers_1.getAttributesXml)({ "w:val": "false" }) },
        { "w:vertAlign": [], ...(0, xml_helpers_1.getAttributesXml)({ "w:val": "superscript" }) }
    ];
}
function getParagraphTextTag(text, styles) {
    let result = {
        "w:r": [
            {
                "w:t": [(0, xml_helpers_1.getXmlTextTag)(text)],
                ...(0, xml_helpers_1.getAttributesXml)({ "xml:space": "preserve" })
            }
        ]
    };
    if (styles) {
        result["w:r"].unshift({
            "w:rPr": styles
        });
    }
    return result;
}
function getLanguageStyles(language) {
    let superStyles = getSuperscriptTextStyle();
    let textStyles = undefined;
    if (language === "en") {
        let langTag = { "w:lang": [], ...(0, xml_helpers_1.getAttributesXml)({ "w:val": "en-US" }) };
        superStyles = [...superStyles, langTag];
        textStyles = [langTag];
    }
    return { superStyles, textStyles };
}
function templateAuthorList(templateBody, meta) {
    let authors = meta["ispras_templates"].authors;
    let organizations = meta["ispras_templates"].organizations;
    // Build org ID → 1-based index map
    let orgIdToIndex = new Map();
    if (organizations) {
        for (let i = 0; i < organizations.length; i++) {
            let org = organizations[i];
            if (!org.id) {
                throw new Error(`Organization at index ${i} is missing required 'id' field`);
            }
            if (!org.name_ru || !org.name_en) {
                throw new Error(`Organization '${org.id}' is missing required 'name_ru' or 'name_en' field`);
            }
            orgIdToIndex.set(org.id, i + 1);
        }
    }
    for (let language of languages) {
        let paragraphIndex = findParagraphWithPatternStrict(templateBody, `{{{authors_${language}}}}`);
        let newParagraphs = [];
        for (let author of authors) {
            let newParagraph = JSON.parse(JSON.stringify(templateBody[paragraphIndex]));
            clearParagraphContents(newParagraph);
            // Build superscript index from author's organizations
            let indexLine;
            if (author.organizations && organizations) {
                let indices = author.organizations.map((orgId) => {
                    let idx = orgIdToIndex.get(orgId);
                    if (idx === undefined) {
                        throw new Error(`Author '${author["name_" + language]}' references unknown organization '${orgId}'`);
                    }
                    return String(idx);
                });
                indexLine = indices.join(",");
            }
            else {
                // Fallback: sequential numbering (legacy format)
                indexLine = String(authors.indexOf(author) + 1);
            }
            let authorLine = author["name_" + language] + ", ORCID: " + author.orcid + " <" + author.email + ">";
            let { superStyles, textStyles } = getLanguageStyles(language);
            let indexTag = getParagraphTextTag(indexLine, superStyles);
            let authorTag = getParagraphTextTag(authorLine, textStyles);
            let spaceTag = getParagraphTextTag(" ", textStyles);
            newParagraph["w:p"].push(indexTag, spaceTag, authorTag);
            newParagraphs.push(newParagraph);
        }
        // Add spacing override to first author paragraph
        if (newParagraphs.length > 0 && language === "ru") {
            addParagraphSpacing(newParagraphs[0], { "w:before": SPACING_BEFORE_FIRST_AUTHOR, "w:after": "0" });
        }
        templateBody.splice(paragraphIndex, 1, ...newParagraphs);
    }
    for (let language of languages) {
        let paragraphIndex = findParagraphWithPatternStrict(templateBody, `{{{organizations_${language}}}}`);
        let newParagraphs = [];
        let orgNames;
        if (organizations) {
            orgNames = organizations.map(org => org["name_" + language]);
        }
        else {
            let orgList = meta["ispras_templates"]["organizations_" + language];
            if (!orgList) {
                throw new Error(`Missing organizations data: provide either 'organizations' or 'organizations_${language}'`);
            }
            orgNames = orgList;
        }
        for (let i = 0; i < orgNames.length; i++) {
            let orgName = orgNames[i];
            let lines = Array.isArray(orgName) ? orgName : [orgName];
            let orgFirstParagraphIndex = newParagraphs.length;
            for (let j = 0; j < lines.length; j++) {
                let newParagraph = JSON.parse(JSON.stringify(templateBody[paragraphIndex]));
                clearParagraphContents(newParagraph);
                let { superStyles, textStyles } = getLanguageStyles(language);
                if (j === 0) {
                    let indexTag = getParagraphTextTag(String(i + 1), superStyles);
                    let organizationTag = getParagraphTextTag(lines[j], textStyles);
                    newParagraph["w:p"].push(indexTag, organizationTag);
                }
                else {
                    let organizationTag = getParagraphTextTag(lines[j], textStyles);
                    newParagraph["w:p"].push(organizationTag);
                }
                newParagraphs.push(newParagraph);
            }
            // Add spacing before=60 to the first paragraph of each org
            addParagraphSpacing(newParagraphs[orgFirstParagraphIndex], { "w:before": SPACING_BEFORE_FIRST_ORG, "w:after": "0" });
        }
        templateBody.splice(paragraphIndex, 1, ...newParagraphs);
    }
}
function getParagraphWithStyle(style) {
    return {
        "w:p": [{
                "w:pPr": [{
                        "w:pStyle": [],
                        ...(0, xml_helpers_1.getAttributesXml)({ "w:val": style })
                    }]
            }]
    };
}
function getNumPr(ilvl, numId) {
    // <w:numPr>
    //    <w:ilvl w:val="<ilvl>"/>
    //    <w:numId w:val="<numId>"/>
    // </w:numPr>
    return {
        "w:numPr": [{
                "w:ilvl": [],
                ...(0, xml_helpers_1.getAttributesXml)({ "w:val": ilvl })
            }, {
                "w:numId": [],
                ...(0, xml_helpers_1.getAttributesXml)({ "w:val": numId })
            }]
    };
}
function templateReplaceLinks(templateBody, meta, listRules) {
    let litListRule = listRules["LitList"];
    let paragraphIndex = findParagraphWithPatternStrict(templateBody, "{{{links}}}");
    let links = meta["ispras_templates"].links;
    let newParagraphs = [];
    for (let link of links) {
        let newParagraph = getParagraphWithStyle(litListRule.styleName);
        let style = (0, xml_helpers_1.getChildTagRequired)(newParagraph["w:p"], "w:pPr");
        style["w:pPr"].push(getNumPr("0", litListRule.numId));
        newParagraph["w:p"].push(getParagraphTextTag(link));
        newParagraphs.push(newParagraph);
    }
    templateBody.splice(paragraphIndex, 1, ...newParagraphs);
}
function templateReplaceAuthorsDetail(templateBody, meta) {
    let paragraphIndex = findParagraphWithPatternStrict(templateBody, "{{{authors_detail}}}");
    let authors = meta["ispras_templates"].authors;
    let newParagraphs = [];
    for (let author of authors) {
        for (let language of languages) {
            let newParagraph = JSON.parse(JSON.stringify(templateBody[paragraphIndex]));
            let line = author["details_" + language];
            clearParagraphContents(newParagraph);
            newParagraph["w:p"].push(getParagraphTextTag(line));
            newParagraphs.push(newParagraph);
        }
    }
    templateBody.splice(paragraphIndex, 1, ...newParagraphs);
}
/**
 * Reverse an author name from "И.И. Иванов" to "Иванов И.И."
 * Splits at the last space and swaps the two parts.
 */
function reverseAuthorName(name) {
    let lastSpace = name.lastIndexOf(" ");
    if (lastSpace < 0)
        return name;
    return name.substring(lastSpace + 1) + " " + name.substring(0, lastSpace);
}
/**
 * Auto-generate the page header prefix from authors and title metadata.
 * Example: "Иванов И.И., Петров П.П. Заголовок статьи. "
 */
function generatePageHeaderPrefix(templates, lang) {
    let authors = templates.authors;
    let names = authors.map((a) => reverseAuthorName(a["name_" + lang]));
    let title = templates["header_" + lang];
    // Strip trailing period from title to avoid ".."
    title = title.replace(/\.\s*$/, "");
    return names.join(", ") + " " + title + ". ";
}
function replacePageHeaders(headers, meta) {
    let templates = meta["ispras_templates"];
    let header_ru = templates.page_header_ru;
    let header_en = templates.page_header_en;
    if (!header_ru) {
        header_ru = generatePageHeaderPrefix(templates, "ru");
    }
    if (!header_en) {
        header_en = generatePageHeaderPrefix(templates, "en");
    }
    for (let header of headers) {
        replacePageHeaderTemplate(header, `{{{page_header_ru}}}`, header_ru);
        replacePageHeaderTemplate(header, `{{{page_header_en}}}`, header_en);
    }
}
/**
 * Replace a page header placeholder with text.
 * Italic formatting now comes from the reference template itself,
 * so the value is treated as plain text.
 */
function replacePageHeaderTemplate(headerXml, template, value) {
    if (value === "@none") {
        replaceInlineTemplate(headerXml, template, value);
        return;
    }
    replaceStringTemplate(headerXml, template, value);
}
function addParagraphSpacing(paragraph, spacingAttrs) {
    let contents = paragraph["w:p"];
    if (!contents)
        return;
    let pPr = (0, xml_helpers_1.getChildTag)(contents, "w:pPr");
    if (!pPr) {
        pPr = { "w:pPr": [] };
        contents.unshift(pPr);
    }
    // Remove existing spacing if any
    for (let i = 0; i < pPr["w:pPr"].length; i++) {
        if (pPr["w:pPr"][i]["w:spacing"] !== undefined) {
            pPr["w:pPr"].splice(i, 1);
            i--;
        }
    }
    pPr["w:pPr"].push({
        "w:spacing": [],
        ...(0, xml_helpers_1.getAttributesXml)(spacingAttrs)
    });
}
function ensureParagraphStyle(paragraph, styleId) {
    let contents = paragraph["w:p"];
    if (!contents)
        return;
    let pPr = (0, xml_helpers_1.getChildTag)(contents, "w:pPr");
    if (!pPr) {
        pPr = { "w:pPr": [] };
        contents.unshift(pPr);
    }
    // Skip if pStyle already present
    let existingPStyle = (0, xml_helpers_1.getChildTag)(pPr["w:pPr"], "w:pStyle");
    if (existingPStyle)
        return;
    pPr["w:pPr"].unshift({
        "w:pStyle": [],
        ...(0, xml_helpers_1.getAttributesXml)({ "w:val": styleId })
    });
}
function patchMetadataParagraphs(templateBody, normalStyleId, headerStyleId) {
    for (let i = 0; i < templateBody.length; i++) {
        let para = templateBody[i];
        if (!para["w:p"])
            continue;
        let contents = para["w:p"];
        let pPr = (0, xml_helpers_1.getChildTag)(contents, "w:pPr");
        // Check if paragraph already has a pStyle
        if (pPr) {
            let existingPStyle = (0, xml_helpers_1.getChildTag)(pPr["w:pPr"], "w:pStyle");
            if (existingPStyle)
                continue;
        }
        // Detect title paragraphs: center-justified with font size 32 (16pt)
        let isTitle = false;
        if (pPr) {
            let jc = (0, xml_helpers_1.getChildTag)(pPr["w:pPr"], "w:jc");
            if (jc && jc[xml_helpers_1.xmlAttributes] && jc[xml_helpers_1.xmlAttributes]["w:val"] === "center") {
                // Check if any run has w:sz val="32"
                for (let child of contents) {
                    if (child["w:r"]) {
                        let rPr = (0, xml_helpers_1.getChildTag)(child["w:r"], "w:rPr");
                        if (rPr) {
                            let sz = (0, xml_helpers_1.getChildTag)(rPr["w:rPr"], "w:sz");
                            if (sz && sz[xml_helpers_1.xmlAttributes] && sz[xml_helpers_1.xmlAttributes]["w:val"] === "32") {
                                isTitle = true;
                                break;
                            }
                        }
                    }
                }
            }
        }
        if (isTitle && headerStyleId) {
            ensureParagraphStyle(para, headerStyleId);
        }
        else if (normalStyleId) {
            ensureParagraphStyle(para, normalStyleId);
        }
    }
}
function replaceTemplates(template, body, meta) {
    let templateCopy = JSON.parse(JSON.stringify(template));
    let templateBody = (0, xml_helpers_1.getDocumentBody)(templateCopy);
    templateReplaceBodyContents(templateBody, body);
    templateAuthorList(templateBody, meta);
    // Auto-generate for_citation prefix from authors + title if not explicitly set
    for (let language of languages) {
        let key = "for_citation_" + language;
        if (!meta["ispras_templates"][key]) {
            meta["ispras_templates"][key] = generatePageHeaderPrefix(meta["ispras_templates"], language);
        }
    }
    let templates = ["header", "abstract", "keywords", "for_citation", "acknowledgements"];
    for (let templateName of templates) {
        for (let language of languages) {
            let template_lang = templateName + "_" + language;
            let value = meta["ispras_templates"][template_lang];
            replaceInlineTemplate(templateBody, `{{{${template_lang}}}}`, value);
        }
    }
    templateReplaceAuthorsDetail(templateBody, meta);
    return templateCopy;
}
function setXmlns(xml, xmlns) {
    let documentTag = (0, xml_helpers_1.getChildTagRequired)(xml, "w:document");
    for (let [key, value] of xmlns) {
        documentTag[xml_helpers_1.xmlAttributes][key] = value;
    }
}
function patchRelIds(doc, map) {
    if (Array.isArray(doc)) {
        for (let child of doc) {
            patchRelIds(child, map);
        }
        return;
    }
    if (typeof doc !== "object")
        return;
    let tagName = (0, xml_helpers_1.getTagName)(doc);
    let attrs = doc[xml_helpers_1.xmlAttributes];
    if (attrs) {
        for (let attr of ["r:id", "r:embed"]) {
            let relId = attrs[attr];
            if (relId && map.has(relId)) {
                attrs[attr] = map.get(relId);
            }
        }
    }
    patchRelIds(doc[tagName], map);
}
async function fixDocxStyles(sourcePath, targetPath, meta) {
    let resourcesDir = path.join(__dirname, "..", "resources");
    // Load the document (Pandoc output) and template (institutional reference)
    let document = await JSZip.loadAsync(fs.readFileSync(sourcePath));
    let template = await JSZip.loadAsync(fs.readFileSync(resourcesDir + '/isp-reference.docx'));
    let templateStylesXML = await template.file("word/styles.xml").async("string");
    let documentStylesXML = await document.file("word/styles.xml").async("string");
    let templateDocXML = await template.file("word/document.xml").async("string");
    let documentDocXML = await document.file("word/document.xml").async("string");
    let documentContentTypesXML = await document.file("[Content_Types].xml").async("string");
    let documentRelsXML = await document.file("word/_rels/document.xml.rels").async("string");
    let templateRelsXML = await template.file("word/_rels/document.xml.rels").async("string");
    let templateNumberingXML = await template.file("word/numbering.xml").async("string");
    let templateHeader1 = await template.file("word/header1.xml").async("string");
    let templateHeader2 = await template.file("word/header2.xml").async("string");
    let templateHeader3 = await template.file("word/header3.xml").async("string");
    let documentContentTypesParsed = xml_helpers_1.xmlParser.parse(documentContentTypesXML);
    let documentRelsParsed = xml_helpers_1.xmlParser.parse(documentRelsXML);
    let templateRelsParsed = xml_helpers_1.xmlParser.parse(templateRelsXML);
    let templateStylesParsed = xml_helpers_1.xmlParser.parse(templateStylesXML);
    let documentStylesParsed = xml_helpers_1.xmlParser.parse(documentStylesXML);
    let templateDocParsed = xml_helpers_1.xmlParser.parse(templateDocXML);
    let documentDocParsed = xml_helpers_1.xmlParser.parse(documentDocXML);
    let numberingParsed = xml_helpers_1.xmlParser.parse(templateNumberingXML);
    let templateHeader1Parsed = xml_helpers_1.xmlParser.parse(templateHeader1);
    let templateHeader2Parsed = xml_helpers_1.xmlParser.parse(templateHeader2);
    let templateHeader3Parsed = xml_helpers_1.xmlParser.parse(templateHeader3);
    copyStyleSection(templateStylesParsed, documentStylesParsed, "w:latentStyles");
    copyStyleSection(templateStylesParsed, documentStylesParsed, "w:docDefaults");
    let documentStylesNamesToId = getStyleIdsByNameFromDefs((0, xml_helpers_1.getChildTagRequired)(documentStylesParsed, "w:styles")["w:styles"]);
    let templateStylesNamesToId = getStyleIdsByNameFromDefs((0, xml_helpers_1.getChildTagRequired)(templateStylesParsed, "w:styles")["w:styles"]);
    let templateStyleTable = getStyleTable(templateStylesParsed);
    let usedStyles = getUsedStylesDeep(templateDocParsed, templateStyleTable, [
        "ispSubHeader-1 level",
        "ispSubHeader-2 level",
        "ispSubHeader-3 level",
        "ispAuthor",
        "ispAnotation",
        "ispText_main",
        "ispList1",
        "ispListing",
        "ispListing Знак",
        "ispLitList",
        "ispPicture_sign",
        "ispNumList",
        "Normal",
        "ispHeader",
        "header",
        "footer"
    ].map(name => templateStylesNamesToId.get(name)).filter(id => id !== undefined));
    let extractedDefs = extractStyleDefs(templateStylesParsed, usedStyles);
    let extractedStyleIdsByName = getStyleIdsByNameFromDefs(extractedDefs);
    let stylePatch = new Map([
        ["Heading1", extractedStyleIdsByName.get("ispSubHeader-1 level")],
        ["Heading2", extractedStyleIdsByName.get("ispSubHeader-2 level")],
        ["Heading3", extractedStyleIdsByName.get("ispSubHeader-3 level")],
        ["Heading4", extractedStyleIdsByName.get("ispSubHeader-3 level")],
        ["Author", extractedStyleIdsByName.get("ispAuthor")],
        ["AbstractTitle", extractedStyleIdsByName.get("ispAnotation")],
        ["Abstract", extractedStyleIdsByName.get("ispAnotation")],
        ["BlockText", extractedStyleIdsByName.get("ispText_main")],
        ["BodyText", extractedStyleIdsByName.get("ispText_main")],
        ["FirstParagraph", extractedStyleIdsByName.get("ispText_main")],
        ["Normal", extractedStyleIdsByName.get("Normal")],
        ["SourceCode", extractedStyleIdsByName.get("ispListing")],
        ["VerbatimChar", extractedStyleIdsByName.get("ispListing Знак")],
        ["ImageCaption", extractedStyleIdsByName.get("ispPicture_sign")],
        ["Compact", extractedStyleIdsByName.get("Normal")],
    ]);
    let stylesToRemove = new Set([
        "Heading5",
        "Heading6",
        "Heading7",
        "Heading8",
        "Heading9",
    ]);
    for (let possibleCollision of extractedStyleIdsByName) {
        let templateStyleName = possibleCollision[0];
        let templateStyleId = possibleCollision[1];
        if (documentStylesNamesToId.has(templateStyleName)) {
            let documentStyleId = documentStylesNamesToId.get(templateStyleName);
            if (!stylePatch.has(documentStyleId)) {
                stylePatch.set(documentStyleId, templateStyleId);
            }
            stylesToRemove.add(documentStyleId);
        }
    }
    removeCollidedStyles(documentStylesParsed, stylesToRemove);
    patchStyleUseReferences(documentDocParsed, documentStylesParsed, stylePatch);
    appendStyles(documentStylesParsed, extractedDefs);
    let patchRules = {
        "OrderedList": { styleName: extractedStyleIdsByName.get("ispNumList"), numId: NUM_ID_ORDERED },
        "BulletList": { styleName: extractedStyleIdsByName.get("ispList1"), numId: NUM_ID_BULLET },
        "LitList": { styleName: extractedStyleIdsByName.get("ispLitList"), numId: NUM_ID_BIBLIOGRAPHY },
    };
    let newListStyles = applyListStyles(documentDocParsed, patchRules);
    setXmlns(templateDocParsed, properDocXmlns);
    let relMap = transferRels(templateRelsParsed, documentRelsParsed);
    patchRelIds(templateDocParsed, relMap);
    let documentBody = (0, xml_helpers_1.getDocumentBody)(documentDocParsed);
    // Strip Pandoc's sectPr — the template already has the correct one
    for (let i = documentBody.length - 1; i >= 0; i--) {
        if ((0, xml_helpers_1.getTagName)(documentBody[i]) === "w:sectPr") {
            documentBody.splice(i, 1);
        }
    }
    documentDocParsed = replaceTemplates(templateDocParsed, documentBody, meta);
    let finalBody = (0, xml_helpers_1.getDocumentBody)(documentDocParsed);
    patchMetadataParagraphs(finalBody, extractedStyleIdsByName.get("Normal"), extractedStyleIdsByName.get("ispHeader"));
    templateReplaceLinks(finalBody, meta, patchRules);
    addNewNumberings(numberingParsed, newListStyles);
    replacePageHeaders([templateHeader1Parsed, templateHeader2Parsed, templateHeader3Parsed], meta);
    let footerContentType = "application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml";
    let headerContentType = "application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml";
    for (let i = 1; i <= 3; i++) {
        addContentType(documentContentTypesParsed, `/word/footer${i}.xml`, footerContentType);
        addContentType(documentContentTypesParsed, `/word/header${i}.xml`, headerContentType);
    }
    let filesToCopy = [
        ...["header", "footer"].flatMap(t => [1, 2, 3].map(i => `word/_rels/${t}${i}.xml.rels`)),
        ...[1, 2, 3].map(i => `word/footer${i}.xml`),
        "word/footnotes.xml", "word/theme/theme1.xml", "word/fontTable.xml",
        "word/settings.xml", "word/webSettings.xml", "word/media/image1.png",
    ];
    for (let file of filesToCopy) {
        await copyFile(template, document, file);
    }
    document.file("word/header1.xml", xml_helpers_1.xmlBuilder.build(templateHeader1Parsed));
    document.file("word/header2.xml", xml_helpers_1.xmlBuilder.build(templateHeader2Parsed));
    document.file("word/header3.xml", xml_helpers_1.xmlBuilder.build(templateHeader3Parsed));
    document.file("word/_rels/document.xml.rels", xml_helpers_1.xmlBuilder.build(documentRelsParsed));
    document.file("[Content_Types].xml", xml_helpers_1.xmlBuilder.build(documentContentTypesParsed));
    document.file("word/numbering.xml", xml_helpers_1.xmlBuilder.build(numberingParsed));
    document.file("word/styles.xml", xml_helpers_1.xmlBuilder.build(documentStylesParsed));
    document.file("word/document.xml", xml_helpers_1.xmlBuilder.build(documentDocParsed));
    fs.writeFileSync(targetPath, await document.generateAsync({ type: "uint8array" }));
}
function fixCompactLists(list) {
    // For compact list, 'para' is replaced with 'plain'.
    // Compact lists were not mentioned in the
    // guidelines, so get rid of them
    for (let i = 0; i < list.c.length; i++) {
        let element = list.c[i];
        if (typeof element[0] === "object" && element[0].t === "Plain") {
            element[0].t = "Para";
        }
        list.c[i] = getPatchedMetaElement(list.c[i]);
    }
    return [
        {
            t: "RawBlock",
            c: ["openxml", `<!-- ListMode ${list.t} -->`]
        },
        list,
        {
            t: "RawBlock",
            c: ["openxml", `<!-- ListMode None -->`]
        }
    ];
}
function getImageCaption(content) {
    let elements = [
        {
            "w:pPr": [
                {
                    "w:pStyle": [],
                    ...(0, xml_helpers_1.getAttributesXml)({ "w:val": "ImageCaption" })
                }, {
                    "w:contextualSpacing": [],
                    ...(0, xml_helpers_1.getAttributesXml)({ "w:val": "true" })
                }
            ]
        },
        getParagraphTextTag((0, pandoc_helpers_1.getMetaString)(content))
    ];
    return {
        t: "RawBlock",
        c: ["openxml", `<w:p>${xml_helpers_1.xmlBuilder.build(elements)}</w:p>`]
    };
}
function getListingCaption(content) {
    let elements = [
        {
            "w:pPr": [
                {
                    "w:pStyle": [],
                    ...(0, xml_helpers_1.getAttributesXml)({ "w:val": "BodyText" })
                }, {
                    "w:jc": [],
                    ...(0, xml_helpers_1.getAttributesXml)({ "w:val": "left" })
                },
            ]
        },
        getParagraphTextTag((0, pandoc_helpers_1.getMetaString)(content), [
            { "w:i": [] },
            { "w:iCs": [] },
            { "w:sz": [], ...(0, xml_helpers_1.getAttributesXml)({ "w:val": "18" }) },
            { "w:szCs": [], ...(0, xml_helpers_1.getAttributesXml)({ "w:val": "18" }) },
        ])
    ];
    return {
        t: "RawBlock",
        c: ["openxml", `<w:p>${xml_helpers_1.xmlBuilder.build(elements)}</w:p>`]
    };
}
function getPatchedMetaElement(element) {
    if (Array.isArray(element)) {
        let newArray = [];
        for (let i = 0; i < element.length; i++) {
            let patched = getPatchedMetaElement(element[i]);
            if (Array.isArray(patched) && !Array.isArray(element[i])) {
                newArray.push(...patched);
            }
            else {
                newArray.push(patched);
            }
        }
        return newArray;
    }
    if (typeof element !== "object" || !element) {
        return element;
    }
    let type = element.t;
    let value = element.c;
    if (type === 'Div') {
        let content = value[1];
        let classes = value[0][1];
        if (classes) {
            if (classes.includes("img-caption")) {
                return getImageCaption(content);
            }
            if (classes.includes("table-caption") || classes.includes("listing-caption")) {
                return getListingCaption(content);
            }
        }
    }
    else if (type === 'BulletList' || type === 'OrderedList') {
        return fixCompactLists(element);
    }
    for (let key of Object.getOwnPropertyNames(element)) {
        element[key] = getPatchedMetaElement(element[key]);
    }
    return element;
}
async function generatePandocDocx(source, target) {
    let markdown = await fs.promises.readFile(source, "utf-8");
    let metaParsed = await (0, pandoc_helpers_1.markdownToJson)(markdown);
    // Resolve references BEFORE caption processing, so captions get resolved numbers
    (0, references_1.resolveReferences)(metaParsed);
    metaParsed.blocks = getPatchedMetaElement(metaParsed.blocks);
    await (0, pandoc_helpers_1.jsonToDocx)(metaParsed, target);
    return (0, pandoc_helpers_1.convertMetaToObject)(metaParsed.meta);
}
async function main() {
    let argv = process.argv;
    if (argv.length < 4) {
        console.log("Usage: main.js <source> <target>");
        process.exit(1);
    }
    let source = argv[2];
    let target = argv[3];
    let tmpFile = target + ".tmp";
    let meta = await generatePandocDocx(source, tmpFile);
    await fixDocxStyles(tmpFile, target, meta);
    fs.unlinkSync(tmpFile);
}
main().catch(console.error);
