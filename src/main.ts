import * as path from 'path';
import * as fs from 'fs';
import * as JSZip from 'jszip';

import {
    xmlComment, xmlText, xmlAttributes,
    xmlParser, xmlBuilder,
    getChildTag, getChildTagRequired, getTagName,
    getXmlTextTag, getAttributesXml, getParagraphText,
    getDocumentBody
} from './xml-helpers';

import {
    getMetaString, convertMetaToObject,
    markdownToJson, jsonToDocx
} from './pandoc-helpers';

import {resolveReferences} from './references';
import {parseBibFile} from 'bibtex';
import {resolveCitations, formatBibliography} from './bibliography';

const properDocXmlns = new Map<string, string>([
    ["xmlns:w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main"],
    ["xmlns:m", "http://schemas.openxmlformats.org/officeDocument/2006/math"],
    ["xmlns:r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships"],
    ["xmlns:o", "urn:schemas-microsoft-com:office:office"],
    ["xmlns:v", "urn:schemas-microsoft-com:vml"],
    ["xmlns:w10", "urn:schemas-microsoft-com:office:word"],
    ["xmlns:a", "http://schemas.openxmlformats.org/drawingml/2006/main"],
    ["xmlns:pic", "http://schemas.openxmlformats.org/drawingml/2006/picture"],
    ["xmlns:wp", "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"],
])

const languages = ["ru", "en"]

// Numbering definition IDs from the ISP RAS template
const NUM_ID_ORDERED = "33"
const NUM_ID_BULLET = "43"
const NUM_ID_BIBLIOGRAPHY = "80"

// Starting counter for dynamically assigned list numIds
const DYNAMIC_NUM_ID_START = 10000

// Spacing values in twentieths of a point
const SPACING_BEFORE_FIRST_AUTHOR = "480"
const SPACING_BEFORE_FIRST_ORG = "60"
const SPACING_AUTHOR_DETAIL = "120"  // 6pt before/after each author bio paragraph

function getStyleCrossReferences(styles: any): any[] {
    let result = []
    for (let style of getChildTagRequired(styles, "w:styles")["w:styles"]) {
        if (!style["w:style"]) continue
        result.push(style[xmlAttributes])

        let basedOnTag = getChildTag(style["w:style"], "w:basedOn")
        if (basedOnTag) result.push(basedOnTag[xmlAttributes])

        let linkTag = getChildTag(style["w:style"], "w:link")
        if (linkTag) result.push(linkTag[xmlAttributes])

        let nextTag = getChildTag(style["w:style"], "w:next")
        if (nextTag) result.push(nextTag[xmlAttributes])
    }
    return result
}

function getDocStyleUseReferences(doc: any, result: any[] = [], met = new Set()): any[] {
    if (!doc || typeof doc !== "object" || met.has(doc)) {
        return result
    }
    met.add(doc)

    if (Array.isArray(doc)) {
        for (let child of doc) {
            result = getDocStyleUseReferences(child, result, met)
        }
        return result
    }

    let tagName = getTagName(doc)
    if (tagName === "w:pStyle" || tagName === "w:rStyle") {
        result.push(doc[xmlAttributes])
    }
    result = getDocStyleUseReferences(doc[tagName], result, met)

    return result
}

function extractStyleDefs(styles: any, usedStyles: Set<string>): any[] {
    let result = []
    for (let style of getChildTagRequired(styles, "w:styles")["w:styles"]) {
        if (!style["w:style"]) continue
        if (usedStyles.has(style[xmlAttributes]["w:styleId"])) {
            let copy = JSON.parse(JSON.stringify(style))
            result.push(copy)
        }
    }
    return result
}

function setFontSizeForStyle(doc: any, styleId: string, szVal: string) {
    let body = getDocumentBody(doc)
    for (let element of body) {
        if (!element["w:p"]) continue
        let contents = element["w:p"]
        let pPr = getChildTag(contents, "w:pPr")
        if (!pPr) continue
        let pStyle = getChildTag(pPr["w:pPr"], "w:pStyle")
        if (!pStyle || pStyle[xmlAttributes]["w:val"] !== styleId) continue

        // Add w:sz/w:szCs to paragraph-level w:rPr
        let pRPr = getChildTag(pPr["w:pPr"], "w:rPr")
        if (!pRPr) {
            pRPr = { "w:rPr": [] }
            pPr["w:pPr"].push(pRPr)
        }
        pRPr["w:rPr"].push(
            { "w:sz": [], ...getAttributesXml({"w:val": szVal}) },
            { "w:szCs": [], ...getAttributesXml({"w:val": szVal}) },
        )

        // Add w:sz/w:szCs to each run's w:rPr
        for (let child of contents) {
            if (!child["w:r"]) continue
            let rPr = getChildTag(child["w:r"], "w:rPr")
            if (!rPr) {
                rPr = { "w:rPr": [] }
                child["w:r"].unshift(rPr)
            }
            rPr["w:rPr"].push(
                { "w:sz": [], ...getAttributesXml({"w:val": szVal}) },
                { "w:szCs": [], ...getAttributesXml({"w:val": szVal}) },
            )
        }
    }
}

function patchStyleUseReferences(doc: any, styles: any, map: Map<string, string>) {
    let docReferences = getDocStyleUseReferences(doc)
    let crossReferences = getStyleCrossReferences(styles)

    for (let ref of docReferences.concat(crossReferences)) {
        if (ref["w:val"] && map.has(ref["w:val"])) {
            ref["w:val"] = map.get(ref["w:val"])
        }
    }
}

function getUsedStyles(doc: any): Set<string> {
    let references = getDocStyleUseReferences(doc)
    let set = new Set<string>()

    for (let ref of references) {
        set.add(ref["w:val"])
    }

    return set
}

function populateStyles(styles: Set<string>, table: Map<string, any>) {
    for (let styleId of styles) {
        let style = table.get(styleId)

        if (!style) {
            throw new Error("Style id " + styleId + " not found")
        }

        let basedOnTag = getChildTag(style["w:style"], "w:basedOn")
        if (basedOnTag) styles.add(basedOnTag[xmlAttributes]["w:val"])

        let linkTag = getChildTag(style["w:style"], "w:link")
        if (linkTag) styles.add(linkTag[xmlAttributes]["w:val"])

        let nextTag = getChildTag(style["w:style"], "w:next")
        if (nextTag) styles.add(nextTag[xmlAttributes]["w:val"])
    }
}

function getUsedStylesDeep(doc: any, styleTable: Map<string, any>, requiredStyles: string[] = []): Set<string> {
    let usedStyles = getUsedStyles(doc)

    for (let requiredStyle of requiredStyles) {
        usedStyles.add(requiredStyle)
    }

    let prevSize: number
    do {
        prevSize = usedStyles.size
        populateStyles(usedStyles, styleTable)
    } while (usedStyles.size > prevSize)

    return usedStyles
}

function getStyleTable(styles: any): Map<string, any> {
    let table = new Map<string, any>()

    for (let style of getChildTagRequired(styles, "w:styles")["w:styles"]) {
        if (!style["w:style"]) continue
        table.set(style[xmlAttributes]["w:styleId"], style)
    }

    return table
}

function getStyleIdsByNameFromDefs(styles: any): Map<string, any> {
    let table = new Map<string, any>()

    for (let style of styles) {
        if (!style["w:style"]) continue
        let nameNode = getChildTag(style["w:style"], "w:name")

        if (nameNode) {
            table.set(nameNode[xmlAttributes]["w:val"], style[xmlAttributes]["w:styleId"])
        }
    }

    return table
}

function appendStyles(target: any, defs: any[]) {
    let styles = getChildTagRequired(target, "w:styles")["w:styles"]
    for (let def of defs) {
        styles.push(def)
    }
}

interface ListStyles {
    BulletList: NumIdPatchEntry
    OrderedList: NumIdPatchEntry

    [key: string]: NumIdPatchEntry | undefined
}

interface NumIdPatchEntry {
    styleName: string
    numId: string
}

function applyListStyles(doc: any, styles: ListStyles): Map<string, string> {

    let stack: any[] = []
    let currentState: any = undefined

    let met = new Set()
    let newStyles = new Map<string, string>()
    let lastId = DYNAMIC_NUM_ID_START

    const walk = (node: any) => {

        if (!node || typeof node !== "object" || met.has(node)) {
            return
        }
        met.add(node)

        for (let key of Object.getOwnPropertyNames(node)) {
            walk(node[key])

            if (key === "w:pPr" && currentState) {
                // Remove any old pStyle and add our own

                for (let i = 0; i < node[key].length; i++) {
                    if (node[key][i]["w:pStyle"]) {
                        node[key].splice(i, 1)
                        i--
                    }
                }

                node[key].unshift({
                    "w:pStyle": {},
                    ...getAttributesXml({"w:val": styles[currentState.listStyle].styleName})
                })

                // Set explicit paragraph properties on bullet list items
                // to match the official template
                if (currentState.listStyle === "BulletList") {
                    for (let i = 0; i < node[key].length; i++) {
                        if (node[key][i]["w:spacing"] || node[key][i]["w:tabs"] ||
                            node[key][i]["w:ind"] || node[key][i]["w:jc"]) {
                            node[key].splice(i, 1)
                            i--
                        }
                    }
                    node[key].push(
                        {
                            "w:tabs": [
                                { "w:tab": [], ...getAttributesXml({"w:val": "clear", "w:pos": "284"}) },
                                { "w:tab": [], ...getAttributesXml({"w:val": "left", "w:pos": "707", "w:leader": "none"}) },
                            ]
                        },
                        {
                            "w:spacing": [],
                            ...getAttributesXml({"w:before": "60", "w:after": "60"})
                        },
                        {
                            "w:ind": [],
                            ...getAttributesXml({"w:hanging": "284", "w:left": "709"})
                        },
                        {
                            "w:jc": [],
                            ...getAttributesXml({"w:val": "both"})
                        },
                    )
                }

                // Set explicit paragraph properties on ordered list items
                // to match the official template
                if (currentState.listStyle === "OrderedList") {
                    for (let i = 0; i < node[key].length; i++) {
                        if (node[key][i]["w:spacing"] || node[key][i]["w:ind"]) {
                            node[key].splice(i, 1)
                            i--
                        }
                    }
                    node[key].push(
                        {
                            "w:spacing": [],
                            ...getAttributesXml({"w:before": "60", "w:after": "60"})
                        },
                        {
                            "w:ind": [],
                            ...getAttributesXml({"w:hanging": "284", "w:left": "709"})
                        },
                    )
                }
            }

            if (key === "w:numId" && currentState) {
                node[xmlAttributes]["w:val"] = String(currentState.numId)
            }

            if (key === xmlComment) {
                let commentValue = node[key][0][xmlText]

                for (let mode of ["OrderedList", "BulletList"]) {
                    if (commentValue.includes(`ListMode ${mode}`)) {
                        stack.push(currentState)
                        currentState = { numId: lastId++, listStyle: mode }
                        newStyles.set(String(currentState.numId), styles[currentState.listStyle].numId)
                    }
                }

                if (commentValue.includes("ListMode None")) {
                    currentState = stack.pop()
                }
            }
        }
    }

    walk(doc)

    return newStyles
}

function removeCollidedStyles(styles: any, collisions: Set<string>) {
    let newContents = []

    for (let style of getChildTagRequired(styles, "w:styles")["w:styles"]) {
        if (!style["w:style"] || !collisions.has(style[xmlAttributes]["w:styleId"])) {
            newContents.push(style)
        }
    }

    getChildTagRequired(styles, "w:styles")["w:styles"] = newContents
}

function copyStyleSection(source: any, target: any, tagName: string) {
    let sourceStyles = getChildTagRequired(source, "w:styles")["w:styles"]
    let targetStyles = getChildTagRequired(target, "w:styles")["w:styles"]

    let sourceSection = getChildTagRequired(sourceStyles, tagName)
    let targetSection = getChildTagRequired(targetStyles, tagName)

    targetSection[tagName] = JSON.parse(JSON.stringify(sourceSection[tagName]))
    if (sourceSection[xmlAttributes]) {
        targetSection[xmlAttributes] = JSON.parse(JSON.stringify(sourceSection[xmlAttributes]))
    }
}

async function copyFile(source: any, target: any, filePath: string) {
    target.file(filePath, await source.file(filePath).async("arraybuffer"))
}

function addNewNumberings(targetNumberingParsed: any, newListStyles: Map<string, string>) {
    let numberingTag = getChildTagRequired(targetNumberingParsed, "w:numbering")["w:numbering"]

    // Build numId → abstractNumId lookup from existing entries
    let numIdToAbstractNumId = new Map<string, string>()
    for (let entry of numberingTag) {
        if (entry["w:num"]) {
            let numId = entry[xmlAttributes]["w:numId"]
            for (let child of entry["w:num"]) {
                if (child["w:abstractNumId"]) {
                    numIdToAbstractNumId.set(numId, child[xmlAttributes]["w:val"])
                }
            }
        }
    }

    // <w:num w:numId="newNum">
    //   <w:abstractNumId w:val="abstractNumId"/>
    // </w:num>

    for (let [newNum, oldNum] of newListStyles) {
        let abstractNumId = numIdToAbstractNumId.get(oldNum) || oldNum

        let overrides = []
        for (let i = 0; i < 9; i++) {
            overrides.push({
                "w:lvlOverride": [{
                    "w:startOverride": [],
                    ...getAttributesXml({"w:val": "1"})
                }],
                ...getAttributesXml({"w:ilvl": String(i)})
            })
        }

        numberingTag.push({
            "w:num": [{
                "w:abstractNumId": [],
                ...getAttributesXml({"w:val": abstractNumId})
            }, ...overrides],
            ...getAttributesXml({"w:numId": newNum})
        })
    }
}

function addContentType(contentTypes: any, partName: string, contentType: string) {
    let typesTag = getChildTagRequired(contentTypes, "Types")["Types"]

    typesTag.push({
        "Override": [],
        ...getAttributesXml({
            "PartName": partName,
            "ContentType": contentType
        })
    })
}

function transferRels(source, target): Map<string, string> {
    let sourceRels = getChildTagRequired(source, "Relationships")["Relationships"]
    let targetRels = getChildTagRequired(target, "Relationships")["Relationships"]

    let presentIds = new Map<string, string>()
    let idMap = new Map<string, string>()

    for (let rel of targetRels) {
        presentIds.set(rel[xmlAttributes]["Target"], rel[xmlAttributes]["Id"])
    }

    let newIdCounter = 0

    for (let rel of sourceRels) {
        if (presentIds.has(rel[xmlAttributes]["Target"])) {
            idMap.set(rel[xmlAttributes]["Id"], presentIds.get(rel[xmlAttributes]["Target"]))
        } else {
            let newId = "template-id-" + (newIdCounter++)
            let relCopy = JSON.parse(JSON.stringify(rel))
            relCopy[xmlAttributes]["Id"] = newId
            targetRels.push(relCopy)
            idMap.set(rel[xmlAttributes]["Id"], newId)
        }
    }

    return idMap
}

function replaceInlineTemplate(body: any[], template: string, value: string) {
    if (!value || value === "@none") {
        let i = findParagraphWithPattern(body, template, 0);
        while (i !== null) {
            body.splice(i, 1)
            i = findParagraphWithPattern(body, template, i)
        }
    } else {
        replaceStringTemplate(body, template, value)
    }
}

function replaceStringTemplate(tag: any, template: string, value: string) {
    if (Array.isArray(tag)) {
        for (let child of tag) {
            replaceStringTemplate(child, template, value)
        }
        return
    }

    let tagName = getTagName(tag)

    if (tagName === xmlText) {
        tag[xmlText] = String(tag[xmlText]).replace(template, value)
    } else if (typeof tag[tagName] === "object") {
        replaceStringTemplate(tag[tagName], template, value)
    }
}

function findParagraphWithPattern(body: any, pattern: string, startIndex: number = 0): number | null {
    for (let i = startIndex; i < body.length; i++) {
        let text = getParagraphText(body[i])
        if (!text.includes(pattern)) {
            continue
        }
        return i
    }

    return null
}

function findParagraphWithPatternStrict(body: any, pattern: string, startIndex: number = 0): number {
    let paragraphIndex = findParagraphWithPattern(body, pattern, startIndex)
    if (paragraphIndex === null) {
        throw new Error(`The template document should have pattern ${pattern}`)
    }

    let text = getParagraphText(body[paragraphIndex])
    if (text !== pattern) {
        throw new Error(`The ${pattern} pattern should be the only text of the paragraph`)
    }

    return paragraphIndex
}

function templateReplaceBodyContents(templateBody: any, body: any) {
    let paragraphIndex = findParagraphWithPatternStrict(templateBody, "{{{body}}}")

    templateBody.splice(paragraphIndex, 1, ...body)
}

function clearParagraphContents(paragraph: any): void {
    let contents = paragraph["w:p"]

    for (let i = 0; i < contents.length; i++) {
        let tagName = getTagName(contents[i])
        if (tagName === "w:r") {
            contents.splice(i, 1)
            i--
        }
    }
}

function getSuperscriptTextStyle(): any[] {
    return [
        { "w:i": [], ...getAttributesXml({"w:val": "false"}) },
        { "w:vertAlign": [], ...getAttributesXml({"w:val": "superscript"}) }
    ]
}

function getParagraphTextTag(text: string, styles?: any[]): any {
    let result = {
        "w:r": [
            {
                "w:t": [getXmlTextTag(text)],
                ...getAttributesXml({"xml:space": "preserve"})
            }
        ]
    }

    if(styles) {
        result["w:r"].unshift({
            "w:rPr": styles
        })
    }

    return result;
}

function getLanguageStyles(language: string): { superStyles: any[], textStyles: any[] | undefined } {
    let superStyles = getSuperscriptTextStyle()
    let textStyles: any[] | undefined = undefined
    if (language === "en") {
        let langTag = { "w:lang": [], ...getAttributesXml({"w:val": "en-US"}) }
        superStyles = [...superStyles, langTag]
        textStyles = [langTag]
    }
    return { superStyles, textStyles }
}

function templateAuthorList(templateBody: any, meta: any) {

    let authors = meta["ispras_templates"].authors
    let organizations = meta["ispras_templates"].organizations

    // Build org ID → 1-based index map
    let orgIdToIndex = new Map<string, number>()
    if (organizations) {
        for (let i = 0; i < organizations.length; i++) {
            let org = organizations[i]
            if (!org.id) {
                throw new Error(`Organization at index ${i} is missing required 'id' field`)
            }
            if (!org.name_ru || !org.name_en) {
                throw new Error(`Organization '${org.id}' is missing required 'name_ru' or 'name_en' field`)
            }
            orgIdToIndex.set(org.id, i + 1)
        }
    }

    for (let language of languages) {
        let paragraphIndex = findParagraphWithPatternStrict(templateBody, `{{{authors_${language}}}}`)

        let newParagraphs = []

        for (let author of authors) {
            let newParagraph = JSON.parse(JSON.stringify(templateBody[paragraphIndex]))
            clearParagraphContents(newParagraph)

            // Build superscript index from author's organizations
            let indexLine: string
            if (author.organizations && organizations) {
                let indices = author.organizations.map((orgId: string) => {
                    let idx = orgIdToIndex.get(orgId)
                    if (idx === undefined) {
                        throw new Error(`Author '${author["name_" + language]}' references unknown organization '${orgId}'`)
                    }
                    return String(idx)
                })
                indexLine = indices.join(",")
            } else {
                // Fallback: sequential numbering (legacy format)
                indexLine = String(authors.indexOf(author) + 1)
            }

            let authorLine = author["name_" + language] + ", ORCID: " + author.orcid + " <" + author.email + ">"

            let { superStyles, textStyles } = getLanguageStyles(language)

            let indexTag = getParagraphTextTag(indexLine, superStyles)
            let authorTag = getParagraphTextTag(authorLine, textStyles)

            let spaceTag = getParagraphTextTag(" ", textStyles)
            newParagraph["w:p"].push(indexTag, spaceTag, authorTag)
            newParagraphs.push(newParagraph)
        }

        // Add spacing override to first author paragraph
        if (newParagraphs.length > 0 && language === "ru") {
            addParagraphSpacing(newParagraphs[0], {"w:before": SPACING_BEFORE_FIRST_AUTHOR, "w:after": "0"})
        }

        templateBody.splice(paragraphIndex, 1, ...newParagraphs)
    }

    for (let language of languages) {
        let paragraphIndex = findParagraphWithPatternStrict(templateBody, `{{{organizations_${language}}}}`)

        let newParagraphs = []

        let orgNames: (string | string[])[]
        if (organizations) {
            orgNames = organizations.map(org => org["name_" + language])
        } else {
            let orgList = meta["ispras_templates"]["organizations_" + language]
            if (!orgList) {
                throw new Error(`Missing organizations data: provide either 'organizations' or 'organizations_${language}'`)
            }
            orgNames = orgList
        }

        for (let i = 0; i < orgNames.length; i++) {
            let orgName = orgNames[i]
            let lines = Array.isArray(orgName) ? orgName : [orgName]
            let orgFirstParagraphIndex = newParagraphs.length

            for (let j = 0; j < lines.length; j++) {
                let newParagraph = JSON.parse(JSON.stringify(templateBody[paragraphIndex]))
                clearParagraphContents(newParagraph)

                let { superStyles, textStyles } = getLanguageStyles(language)

                if (j === 0) {
                    let indexTag = getParagraphTextTag(String(i + 1), superStyles)
                    let organizationTag = getParagraphTextTag(lines[j], textStyles)
                    newParagraph["w:p"].push(indexTag, organizationTag)
                } else {
                    let organizationTag = getParagraphTextTag(lines[j], textStyles)
                    newParagraph["w:p"].push(organizationTag)
                }

                newParagraphs.push(newParagraph)
            }

            // Add spacing before=60 to the first paragraph of each org
            addParagraphSpacing(newParagraphs[orgFirstParagraphIndex], {"w:before": SPACING_BEFORE_FIRST_ORG, "w:after": "0"})
        }

        templateBody.splice(paragraphIndex, 1, ...newParagraphs)
    }
}

function getParagraphWithStyle(style: string): any {
    return {
        "w:p": [{
            "w:pPr": [{
                "w:pStyle": [],
                ...getAttributesXml({"w:val": style})
            }]
        }]
    };
}

function getNumPr(ilvl: string, numId: string): any {
    // <w:numPr>
    //    <w:ilvl w:val="<ilvl>"/>
    //    <w:numId w:val="<numId>"/>
    // </w:numPr>

    return {
        "w:numPr": [{
            "w:ilvl": [],
            ...getAttributesXml({"w:val": ilvl})
        }, {
            "w:numId": [],
            ...getAttributesXml({"w:val": numId})
        }]
    }
}

function templateReplaceLinks(templateBody: any, meta: any, listRules: any) {
    let litListRule = listRules["LitList"]
    let paragraphIndex = findParagraphWithPatternStrict(templateBody, "{{{links}}}")
    let links = meta["ispras_templates"].links

    let newParagraphs = []

    for (let link of links) {
        let newParagraph = getParagraphWithStyle(litListRule.styleName)
        let style = getChildTagRequired(newParagraph["w:p"], "w:pPr")
        style["w:pPr"].push(getNumPr("0", litListRule.numId))

        newParagraph["w:p"].push(getParagraphTextTag(link))
        newParagraphs.push(newParagraph)
    }

    templateBody.splice(paragraphIndex, 1, ...newParagraphs)
}

function templateReplaceAuthorsDetail(templateBody: any, meta: any) {
    let paragraphIndex = findParagraphWithPatternStrict(templateBody, "{{{authors_detail}}}")
    let authors = meta["ispras_templates"].authors

    let newParagraphs = []

    for (let author of authors) {
        for (let language of languages) {
            let newParagraph = JSON.parse(JSON.stringify(templateBody[paragraphIndex]))

            let line = author["details_" + language]

            clearParagraphContents(newParagraph)
            newParagraph["w:p"].push(getParagraphTextTag(line))
            addParagraphSpacing(newParagraph, {"w:before": SPACING_AUTHOR_DETAIL, "w:after": SPACING_AUTHOR_DETAIL})
            newParagraphs.push(newParagraph)
        }
    }

    templateBody.splice(paragraphIndex, 1, ...newParagraphs)
}

/**
 * Reverse an author name from "И.И. Иванов" to "Иванов И.И."
 * Splits at the last space and swaps the two parts.
 */
function reverseAuthorName(name: string): string {
    let lastSpace = name.lastIndexOf(" ")
    if (lastSpace < 0) return name
    return name.substring(lastSpace + 1) + " " + name.substring(0, lastSpace)
}

/**
 * Auto-generate the page header prefix from authors and title metadata.
 * Example: "Иванов И.И., Петров П.П. Заголовок статьи. "
 */
function generatePageHeaderPrefix(templates: any, lang: string): string {
    let authors = templates.authors
    let names = authors.map((a: any) => reverseAuthorName(a["name_" + lang]))
    let title = templates["header_" + lang]
    // Strip trailing period from title to avoid ".."
    title = title.replace(/\.\s*$/, "")
    return names.join(", ") + " " + title + ". "
}

function replacePageHeaders(headers: any[], meta: any): void {
    let templates = meta["ispras_templates"]
    let header_ru = templates.page_header_ru
    let header_en = templates.page_header_en

    if (!header_ru) {
        header_ru = generatePageHeaderPrefix(templates, "ru")
    }
    if (!header_en) {
        header_en = generatePageHeaderPrefix(templates, "en")
    }

    for (let header of headers) {
        replacePageHeaderTemplate(header, `{{{page_header_ru}}}`, header_ru)
        replacePageHeaderTemplate(header, `{{{page_header_en}}}`, header_en)
    }
}

/**
 * Replace a page header placeholder with text.
 * Italic formatting now comes from the reference template itself,
 * so the value is treated as plain text.
 */
function replacePageHeaderTemplate(headerXml: any[], template: string, value: string): void {
    if (value === "@none") {
        replaceInlineTemplate(headerXml, template, value)
        return
    }
    replaceStringTemplate(headerXml, template, value)
}

function addParagraphSpacing(paragraph: any, spacingAttrs: Record<string, string>): void {
    let contents = paragraph["w:p"]
    if (!contents) return

    let pPr = getChildTag(contents, "w:pPr")
    if (!pPr) {
        pPr = { "w:pPr": [] }
        contents.unshift(pPr)
    }

    // Remove existing spacing if any
    for (let i = 0; i < pPr["w:pPr"].length; i++) {
        if (pPr["w:pPr"][i]["w:spacing"] !== undefined) {
            pPr["w:pPr"].splice(i, 1)
            i--
        }
    }

    pPr["w:pPr"].push({
        "w:spacing": [],
        ...getAttributesXml(spacingAttrs)
    })
}

function ensureParagraphStyle(paragraph: any, styleId: string): void {
    let contents = paragraph["w:p"]
    if (!contents) return

    let pPr = getChildTag(contents, "w:pPr")
    if (!pPr) {
        pPr = { "w:pPr": [] }
        contents.unshift(pPr)
    }

    // Skip if pStyle already present
    let existingPStyle = getChildTag(pPr["w:pPr"], "w:pStyle")
    if (existingPStyle) return

    pPr["w:pPr"].unshift({
        "w:pStyle": [],
        ...getAttributesXml({"w:val": styleId})
    })
}

function patchMetadataParagraphs(templateBody: any, normalStyleId: string, headerStyleId: string): void {
    for (let i = 0; i < templateBody.length; i++) {
        let para = templateBody[i]
        if (!para["w:p"]) continue

        let contents = para["w:p"]
        let pPr = getChildTag(contents, "w:pPr")

        // Check if paragraph already has a pStyle
        if (pPr) {
            let existingPStyle = getChildTag(pPr["w:pPr"], "w:pStyle")
            if (existingPStyle) continue
        }

        // Detect title paragraphs: center-justified with font size 32 (16pt)
        let isTitle = false
        if (pPr) {
            let jc = getChildTag(pPr["w:pPr"], "w:jc")
            if (jc && jc[xmlAttributes] && jc[xmlAttributes]["w:val"] === "center") {
                // Check if any run has w:sz val="32"
                for (let child of contents) {
                    if (child["w:r"]) {
                        let rPr = getChildTag(child["w:r"], "w:rPr")
                        if (rPr) {
                            let sz = getChildTag(rPr["w:rPr"], "w:sz")
                            if (sz && sz[xmlAttributes] && sz[xmlAttributes]["w:val"] === "32") {
                                isTitle = true
                                break
                            }
                        }
                    }
                }
            }
        }

        if (isTitle && headerStyleId) {
            ensureParagraphStyle(para, headerStyleId)
        } else if (normalStyleId) {
            ensureParagraphStyle(para, normalStyleId)
        }
    }
}

function replaceTemplates(template: any, body: any, meta: any): any {
    let templateCopy = JSON.parse(JSON.stringify(template))

    let templateBody = getDocumentBody(templateCopy)

    templateReplaceBodyContents(templateBody, body)
    templateAuthorList(templateBody, meta)

    // Auto-generate for_citation prefix from authors + title if not explicitly set
    for (let language of languages) {
        let key = "for_citation_" + language
        if (!meta["ispras_templates"][key]) {
            meta["ispras_templates"][key] = generatePageHeaderPrefix(meta["ispras_templates"], language)
        }
    }

    let templates = ["header", "abstract", "keywords", "for_citation", "acknowledgements"]

    for (let templateName of templates) {
        for (let language of languages) {
            let template_lang = templateName + "_" + language
            let value = meta["ispras_templates"][template_lang]
            replaceInlineTemplate(templateBody, `{{{${template_lang}}}}`, value)
        }
    }

    templateReplaceAuthorsDetail(templateBody, meta)

    return templateCopy
}

function setXmlns(xml: any, xmlns: Map<string, string>) {
    let documentTag = getChildTagRequired(xml, "w:document")

    for (let [key, value] of xmlns) {
        documentTag[xmlAttributes][key] = value
    }
}

function patchRelIds(doc: any, map: Map<string, string>) {
    if (Array.isArray(doc)) {
        for (let child of doc) {
            patchRelIds(child, map)
        }
        return
    }

    if (typeof doc !== "object") return

    let tagName = getTagName(doc)

    let attrs = doc[xmlAttributes]
    if (attrs) {
        for (let attr of ["r:id", "r:embed"]) {
            let relId = attrs[attr]
            if (relId && map.has(relId)) {
                attrs[attr] = map.get(relId)
            }
        }
    }

    patchRelIds(doc[tagName], map)
}

async function fixDocxStyles(sourcePath: string, targetPath: string, meta: any): Promise<void> {
    let resourcesDir = path.join(__dirname, "..", "resources")

    // Load the document (Pandoc output) and template (institutional reference)
    let document = await JSZip.loadAsync(fs.readFileSync(sourcePath))
    let template = await JSZip.loadAsync(fs.readFileSync(resourcesDir + '/isp-reference.docx'))

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

    let documentContentTypesParsed = xmlParser.parse(documentContentTypesXML);
    let documentRelsParsed = xmlParser.parse(documentRelsXML);
    let templateRelsParsed = xmlParser.parse(templateRelsXML);
    let templateStylesParsed = xmlParser.parse(templateStylesXML);
    let documentStylesParsed = xmlParser.parse(documentStylesXML);
    let templateDocParsed = xmlParser.parse(templateDocXML);
    let documentDocParsed = xmlParser.parse(documentDocXML);
    let numberingParsed = xmlParser.parse(templateNumberingXML);
    let templateHeader1Parsed = xmlParser.parse(templateHeader1)
    let templateHeader2Parsed = xmlParser.parse(templateHeader2)
    let templateHeader3Parsed = xmlParser.parse(templateHeader3)

    copyStyleSection(templateStylesParsed, documentStylesParsed, "w:latentStyles")
    copyStyleSection(templateStylesParsed, documentStylesParsed, "w:docDefaults")

    let documentStylesNamesToId = getStyleIdsByNameFromDefs(getChildTagRequired(documentStylesParsed, "w:styles")["w:styles"]);
    let templateStylesNamesToId = getStyleIdsByNameFromDefs(getChildTagRequired(templateStylesParsed, "w:styles")["w:styles"]);

    let templateStyleTable = getStyleTable(templateStylesParsed)

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
    ].map(name => templateStylesNamesToId.get(name)).filter(id => id !== undefined))
    let extractedDefs = extractStyleDefs(templateStylesParsed, usedStyles)
    let extractedStyleIdsByName = getStyleIdsByNameFromDefs(extractedDefs)

    let stylePatch = new Map<string, string>([
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
    ])

    let stylesToRemove = new Set<string>([
        "Heading5",
        "Heading6",
        "Heading7",
        "Heading8",
        "Heading9",
    ])

    for (let possibleCollision of extractedStyleIdsByName) {
        let templateStyleName = possibleCollision[0]
        let templateStyleId = possibleCollision[1]

        if (documentStylesNamesToId.has(templateStyleName)) {
            let documentStyleId = documentStylesNamesToId.get(templateStyleName)

            if (!stylePatch.has(documentStyleId)) {
                stylePatch.set(documentStyleId, templateStyleId)
            }
            stylesToRemove.add(documentStyleId)
        }
    }

    removeCollidedStyles(documentStylesParsed, stylesToRemove)

    // Set explicit 10pt font size on Heading4 paragraphs before style remapping,
    // since Heading4 maps to ispSubHeader-3level which inherits 11pt from its parent
    setFontSizeForStyle(documentDocParsed, "Heading4", "20")

    patchStyleUseReferences(documentDocParsed, documentStylesParsed, stylePatch)

    appendStyles(documentStylesParsed, extractedDefs)

    let patchRules = {
        "OrderedList": {styleName: extractedStyleIdsByName.get("ispNumList"), numId: NUM_ID_ORDERED},
        "BulletList": {styleName: extractedStyleIdsByName.get("ispList1"), numId: NUM_ID_BULLET},
        "LitList": {styleName: extractedStyleIdsByName.get("ispLitList"), numId: NUM_ID_BIBLIOGRAPHY},
    };

    let newListStyles = applyListStyles(documentDocParsed, patchRules)

    setXmlns(templateDocParsed, properDocXmlns)

    let relMap = transferRels(templateRelsParsed, documentRelsParsed)
    patchRelIds(templateDocParsed, relMap)

    let documentBody = getDocumentBody(documentDocParsed)
    // Strip Pandoc's sectPr — the template already has the correct one
    for (let i = documentBody.length - 1; i >= 0; i--) {
        if (getTagName(documentBody[i]) === "w:sectPr") {
            documentBody.splice(i, 1)
        }
    }
    documentDocParsed = replaceTemplates(templateDocParsed, documentBody, meta)

    let finalBody = getDocumentBody(documentDocParsed)
    patchMetadataParagraphs(finalBody, extractedStyleIdsByName.get("Normal"), extractedStyleIdsByName.get("ispHeader"))
    templateReplaceLinks(finalBody, meta, patchRules)

    addNewNumberings(numberingParsed, newListStyles)

    replacePageHeaders([templateHeader1Parsed, templateHeader2Parsed, templateHeader3Parsed], meta)

    let footerContentType = "application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml"
    let headerContentType = "application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"
    for (let i = 1; i <= 3; i++) {
        addContentType(documentContentTypesParsed, `/word/footer${i}.xml`, footerContentType)
        addContentType(documentContentTypesParsed, `/word/header${i}.xml`, headerContentType)
    }

    let filesToCopy = [
        ...["header", "footer"].flatMap(t => [1, 2, 3].map(i =>
            `word/_rels/${t}${i}.xml.rels`
        )),
        ...[1, 2, 3].map(i => `word/footer${i}.xml`),
        "word/footnotes.xml", "word/theme/theme1.xml", "word/fontTable.xml",
        "word/settings.xml", "word/webSettings.xml", "word/media/image1.png",
    ]
    for (let file of filesToCopy) {
        await copyFile(template, document, file)
    }

    document.file("word/header1.xml", xmlBuilder.build(templateHeader1Parsed))
    document.file("word/header2.xml", xmlBuilder.build(templateHeader2Parsed))
    document.file("word/header3.xml", xmlBuilder.build(templateHeader3Parsed))

    document.file("word/_rels/document.xml.rels", xmlBuilder.build(documentRelsParsed))
    document.file("[Content_Types].xml", xmlBuilder.build(documentContentTypesParsed))
    document.file("word/numbering.xml", xmlBuilder.build(numberingParsed))
    document.file("word/styles.xml", xmlBuilder.build(documentStylesParsed))
    document.file("word/document.xml", xmlBuilder.build(documentDocParsed))

    fs.writeFileSync(targetPath, await document.generateAsync({type: "uint8array"}));
}

function fixCompactLists(list): any[] {
    // For compact list, 'para' is replaced with 'plain'.
    // Compact lists were not mentioned in the
    // guidelines, so get rid of them

    for (let i = 0; i < list.c.length; i++) {
        let element = list.c[i]
        if (typeof element[0] === "object" && element[0].t === "Plain") {
            element[0].t = "Para"
        }
        list.c[i] = getPatchedMetaElement(list.c[i])
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
    ]
}

function getImageCaption(content): any {
    let elements = [
        {
            "w:pPr": [
                {
                    "w:pStyle": [],
                    ...getAttributesXml({"w:val": "ImageCaption"})
                }, {
                    "w:contextualSpacing": [],
                    ...getAttributesXml({"w:val": "true"})
                }]
        },
        getParagraphTextTag(getMetaString(content))
    ];

    return {
        t: "RawBlock",
        c: ["openxml", `<w:p>${xmlBuilder.build(elements)}</w:p>`]
    };
}

function getListingCaption(content): any {
    let elements = [
        {
            "w:pPr": [
                {
                    "w:pStyle": [],
                    ...getAttributesXml({"w:val": "BodyText"})
                }, {
                    "w:jc": [],
                    ...getAttributesXml({"w:val": "left"})
                },
            ]
        },
        getParagraphTextTag(getMetaString(content), [
            {"w:i": []},
            {"w:iCs": []},
            {"w:sz": [], ...getAttributesXml({"w:val": "18"})},
            {"w:szCs": [], ...getAttributesXml({"w:val": "18"})},
        ])
    ];

    return {
        t: "RawBlock",
        c: ["openxml", `<w:p>${xmlBuilder.build(elements)}</w:p>`]
    };
}

function getPatchedMetaElement(element): any {
    if (Array.isArray(element)) {
        let newArray = []

        for (let i = 0; i < element.length; i++) {
            let patched = getPatchedMetaElement(element[i])
            if (Array.isArray(patched) && !Array.isArray(element[i])) {
                newArray.push(...patched)
            } else {
                newArray.push(patched)
            }
        }

        return newArray
    }

    if (typeof element !== "object" || !element) {
        return element
    }

    let type = element.t
    let value = element.c

    if (type === 'Div') {
        let content = value[1];
        let classes = value[0][1];
        if (classes) {
            if (classes.includes("img-caption")) {
                return getImageCaption(content)
            }

            if (classes.includes("table-caption") || classes.includes("listing-caption")) {
                return getListingCaption(content)
            }
        }
    } else if (type === 'BulletList' || type === 'OrderedList') {
        return fixCompactLists(element)
    }

    for (let key of Object.getOwnPropertyNames(element)) {
        element[key] = getPatchedMetaElement(element[key]);
    }

    return element
}

async function generatePandocDocx(source: string, target: string): Promise<any> {
    let markdown = await fs.promises.readFile(source, "utf-8")

    let metaParsed = await markdownToJson(markdown)

    // Check for bibliography field and resolve citations before other processing
    let meta = convertMetaToObject(metaParsed.meta)
    let bibField = meta["ispras_templates"]?.bibliography

    if (bibField) {
        let bibPath = path.resolve(path.dirname(source), bibField)
        let bibContent = await fs.promises.readFile(bibPath, "utf-8")
        let bibFile = parseBibFile(bibContent)

        let citationResult = resolveCitations(metaParsed, bibFile)
        let formattedLinks = formatBibliography(citationResult.citedKeys, bibFile)

        // Inject into meta for templateReplaceLinks() to consume
        meta["ispras_templates"].links = formattedLinks
    }

    // Resolve references BEFORE caption processing, so captions get resolved numbers
    resolveReferences(metaParsed)

    metaParsed.blocks = getPatchedMetaElement(metaParsed.blocks)

    await jsonToDocx(metaParsed, target)

    // When bibliography was used, meta already has the injected links — reuse it.
    // Otherwise, re-derive meta from the (now possibly mutated) AST.
    if (!bibField) {
        meta = convertMetaToObject(metaParsed.meta)
    }
    return meta
}

async function main(): Promise<void> {
    let argv = process.argv
    if (argv.length < 4) {
        console.log("Usage: main.js <source> <target>")
        process.exit(1)
    }

    let source = argv[2]
    let target = argv[3]

    let tmpFile = target + ".tmp"
    let meta = await generatePandocDocx(source, tmpFile)
    await fixDocxStyles(tmpFile, target, meta)
    fs.unlinkSync(tmpFile)
}

main().catch(console.error)
