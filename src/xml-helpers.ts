import {XMLParser, XMLBuilder} from 'fast-xml-parser';

export const xmlComment = "__comment__"
export const xmlText = "__text__"
export const xmlAttributes = ":@"

export const xmlParser = new XMLParser({
    ignoreAttributes: false,
    alwaysCreateTextNode: true,
    attributeNamePrefix: "",
    preserveOrder: true,
    trimValues: false,
    commentPropName: xmlComment,
    textNodeName: xmlText
})

export const xmlBuilder = new XMLBuilder({
    ignoreAttributes: false,
    attributeNamePrefix: "",
    preserveOrder: true,
    commentPropName: xmlComment,
    textNodeName: xmlText
})

export function getChildTag(styles: any, name: string): any {
    for (let child of styles) {
        if (child[name]) {
            return child
        }
    }
    return undefined
}

export function getChildTagRequired(styles: any, name: string): any {
    let result = getChildTag(styles, name)
    if (result === undefined) {
        throw new Error(`Required child tag '${name}' not found`)
    }
    return result
}

export function getTagName(tag: any): string | undefined {
    for (let key of Object.getOwnPropertyNames(tag)) {
        if (key === xmlAttributes) continue
        return key
    }
    return undefined
}

export function getXmlTextTag(text: string): any {
    let result = {};
    result[xmlText] = text
    return result
}

export function getAttributesXml(attributes: any): any {
    let result = {}
    result[xmlAttributes] = attributes
    return result
}

export function getRawText(tag): string {
    let result = ""
    let tagName = getTagName(tag)

    if (tagName === xmlText) {
        result += tag[xmlText]
    }
    if (Array.isArray(tag[tagName])) {
        for (let child of tag[tagName]) {
            result += getRawText(child)
        }
    }

    return result
}

export function getParagraphText(paragraph: any): string {
    let result = ""

    if (paragraph["w:t"]) {
        result += getRawText(paragraph)
    }

    for (let name of Object.getOwnPropertyNames(paragraph)) {
        if (name === xmlAttributes) {
            continue
        }
        if (Array.isArray(paragraph[name])) {
            for (let child of paragraph[name]) {
                result += getParagraphText(child)
            }
        }
    }

    return result
}
