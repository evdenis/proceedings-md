/**
 * generate-reference.ts
 *
 * Generates resources/isp-reference.docx from test/official.docx.
 *
 * The isp-reference.docx is a skeleton of the official ISP RAS proceedings
 * template with {{{placeholder}}} tokens where the converter injects content.
 * This script automates its creation so it stays in sync with the official
 * template.
 *
 * Usage:
 *   node scripts/generate-reference.js test/official.docx resources/isp-reference.docx
 */

import * as fs from "fs"
import JSZip from "jszip"
import {xmlParser, xmlBuilder, xmlAttributes, getChildTag, getChildTagRequired, getTagName, getParagraphText, getXmlTextTag, getAttributesXml, getDocumentBody} from "../src/xml-helpers"

// ── Helpers ──────────────────────────────────────────────────────────────────

function getStyle(paragraph: any): string | null {
    let contents = paragraph["w:p"]
    if (!contents) return null
    let pPr = getChildTag(contents, "w:pPr")
    if (!pPr) return null
    let pStyle = getChildTag(pPr["w:pPr"], "w:pStyle")
    if (!pStyle || !pStyle[xmlAttributes]) return null
    return pStyle[xmlAttributes]["w:val"]
}

function getSpacingBefore(paragraph: any): string | null {
    let contents = paragraph["w:p"]
    if (!contents) return null
    let pPr = getChildTag(contents, "w:pPr")
    if (!pPr) return null
    let spacing = getChildTag(pPr["w:pPr"], "w:spacing")
    if (!spacing || !spacing[xmlAttributes]) return null
    return spacing[xmlAttributes]["w:before"] || null
}

/** Create a paragraph containing only a text run with given content, no style. */
function makePlainParagraph(text: string): any {
    return {
        "w:p": [{
            "w:r": [{
                "w:t": [getXmlTextTag(text)],
                ...getAttributesXml({"xml:space": "preserve"})
            }]
        }]
    }
}

/**
 * Strip w:spacing from a paragraph's pPr.
 * Used for author/org placeholder paragraphs — the converter adds correct
 * spacing at runtime, and baked-in spacing would be cloned to all instances.
 */
function stripParagraphSpacing(paragraph: any): void {
    let contents = paragraph["w:p"]
    if (!contents) return
    let pPr = getChildTag(contents, "w:pPr")
    if (!pPr) return
    for (let i = pPr["w:pPr"].length - 1; i >= 0; i--) {
        if (pPr["w:pPr"][i]["w:spacing"] !== undefined) {
            pPr["w:pPr"].splice(i, 1)
        }
    }
}

/**
 * Replace text content of a paragraph's runs with a placeholder string.
 * Keeps the first run's formatting (w:rPr) but replaces all text.
 */
function replaceAllRunsWithPlaceholder(paragraph: any, placeholder: string): void {
    let contents = paragraph["w:p"]
    if (!contents) return

    // Find first run's rPr
    let firstRPr: any = null
    for (let child of contents) {
        if (child["w:r"]) {
            let rPrTag = getChildTag(child["w:r"], "w:rPr")
            if (rPrTag) {
                firstRPr = JSON.parse(JSON.stringify(rPrTag))
                // Strip highlight
                firstRPr["w:rPr"] = firstRPr["w:rPr"].filter(
                    (item: any) => getTagName(item) !== "w:highlight"
                )
            }
            break
        }
    }

    // Remove all w:r, w:hyperlink, w:bookmarkStart, w:bookmarkEnd
    for (let i = contents.length - 1; i >= 0; i--) {
        let tagName = getTagName(contents[i])
        if (tagName === "w:r" || tagName === "w:hyperlink" || tagName === "w:bookmarkStart" || tagName === "w:bookmarkEnd") {
            contents.splice(i, 1)
        }
    }

    // Build new run
    let newRun: any = {
        "w:r": [{
            "w:t": [getXmlTextTag(placeholder)],
            ...getAttributesXml({"xml:space": "preserve"})
        }]
    }
    if (firstRPr) {
        newRun["w:r"].unshift(firstRPr)
    }

    contents.push(newRun)
}

/**
 * For annotation paragraphs (Аннотация., Keywords:, etc.):
 * Keep the bold prefix, replace everything after with a non-bold placeholder run.
 */
function replaceAnnotationValue(paragraph: any, prefixText: string, placeholder: string): void {
    let contents = paragraph["w:p"]
    if (!contents) return

    // Remove all runs, hyperlinks, bookmarks — we'll reconstruct
    for (let i = contents.length - 1; i >= 0; i--) {
        let tagName = getTagName(contents[i])
        if (tagName === "w:r" || tagName === "w:hyperlink" || tagName === "w:bookmarkStart" || tagName === "w:bookmarkEnd") {
            contents.splice(i, 1)
        }
    }

    // Bold prefix run
    let prefixRun: any = {
        "w:r": [{
            "w:rPr": [{"w:b": []}]
        }, {
            "w:t": [getXmlTextTag(prefixText)],
            ...getAttributesXml({"xml:space": "preserve"})
        }]
    }

    // Non-bold placeholder run
    let placeholderRun: any = {
        "w:r": [{
            "w:rPr": [{
                "w:b": [],
                ...getAttributesXml({"w:val": "false"})
            }]
        }, {
            "w:t": [getXmlTextTag(placeholder)],
            ...getAttributesXml({"xml:space": "preserve"})
        }]
    }

    contents.push(prefixRun, placeholderRun)
}

/**
 * Replace all text content in a header XML with a placeholder string.
 */
function replaceHeaderContent(headerParsed: any, placeholder: string): void {
    let hdr = headerParsed.find((x: any) => x["w:hdr"])
    if (!hdr) return

    for (let para of hdr["w:hdr"]) {
        if (!para["w:p"]) continue

        let contents = para["w:p"]

        // Get first run's rPr
        let firstRPr: any = null
        for (let child of contents) {
            if (child["w:r"]) {
                let rPr = getChildTag(child["w:r"], "w:rPr")
                if (rPr) {
                    firstRPr = JSON.parse(JSON.stringify(rPr))
                    // Strip highlight and italic
                    firstRPr["w:rPr"] = firstRPr["w:rPr"].filter(
                        (item: any) => {
                            let name = getTagName(item)
                            return name !== "w:highlight" && name !== "w:i" && name !== "w:iCs" && name !== "w:spacing"
                        }
                    )
                }
                break
            }
        }

        // Remove all runs
        for (let i = contents.length - 1; i >= 0; i--) {
            if (getTagName(contents[i]) === "w:r") {
                contents.splice(i, 1)
            }
        }

        // Add single run with placeholder
        let newRun: any = {
            "w:r": [{
                "w:t": [getXmlTextTag(placeholder)],
                ...getAttributesXml({"xml:space": "preserve"})
            }]
        }
        if (firstRPr) {
            newRun["w:r"].unshift(firstRPr)
        }

        contents.push(newRun)
        contents.push({"w:r": []})
    }
}

// ── Main transformation ──────────────────────────────────────────────────────

type ParagraphRole =
    | { action: "keep" }
    | { action: "delete" }
    | { action: "replace_full", placeholder: string }
    | { action: "replace_annotation", prefix: string, placeholder: string }
    | { action: "body_placeholder" }
    | { action: "links_placeholder" }
    | { action: "sectPr" }

async function generateReference(inputPath: string, outputPath: string): Promise<void> {
    console.log(`Reading ${inputPath}...`)
    let zip = await JSZip.loadAsync(fs.readFileSync(inputPath))

    // Parse document.xml
    let docXml = await zip.file("word/document.xml")!.async("string")
    let docParsed = xmlParser.parse(docXml)
    let body = getDocumentBody(docParsed)

    // Build style name→ID mapping from styles.xml
    let stylesXml = await zip.file("word/styles.xml")!.async("string")
    let stylesParsed = xmlParser.parse(stylesXml)
    let stylesTag = stylesParsed.find((x: any) => x["w:styles"])
    let styleNameToId = new Map<string, string>()
    for (let child of stylesTag["w:styles"]) {
        if (child["w:style"]) {
            let attrs = child[xmlAttributes]
            let styleId = attrs ? attrs["w:styleId"] : null
            let nameTag = getChildTag(child["w:style"], "w:name")
            if (nameTag && nameTag[xmlAttributes] && styleId) {
                styleNameToId.set(nameTag[xmlAttributes]["w:val"], styleId)
            }
        }
    }

    let ispHeaderId = styleNameToId.get("ispHeader")
    let ispAuthorId = styleNameToId.get("ispAuthor")
    let ispAnotationId = styleNameToId.get("ispAnotation")
    let ispSubHeader2Id = styleNameToId.get("ispSubHeader-2 level")
    let ispLitListId = styleNameToId.get("ispLitList")
    let ispTextmainId = styleNameToId.get("ispText_main")

    console.log("Identified ISP style IDs:", [ispHeaderId, ispAuthorId, ispAnotationId, ispSubHeader2Id, ispLitListId, ispTextmainId].filter(Boolean).length, "of 6")

    // ── Inject missing styles that the official template lacks ──

    // ispListing (paragraph) and ispListing Знак (character) are needed for code blocks
    // but don't exist in the official template. Add them.
    if (!styleNameToId.has("ispListing")) {
        let listingParaId = "ispListing"
        let listingCharId = "ispListingChar"

        stylesTag["w:styles"].push({
            "w:style": [
                {"w:name": [], ...getAttributesXml({"w:val": "ispListing"})},
                {"w:basedOn": [], ...getAttributesXml({"w:val": "Normal"})},
                {"w:link": [], ...getAttributesXml({"w:val": listingCharId})},
                {"w:qFormat": []},
                {"w:rPr": [
                    {"w:rFonts": [], ...getAttributesXml({"w:ascii": "Courier New", "w:hAnsi": "Courier New", "w:cs": "Courier New"})},
                    {"w:color": [], ...getAttributesXml({"w:val": "000000"})},
                    {"w:sz": [], ...getAttributesXml({"w:val": "18"})},
                    {"w:szCs": [], ...getAttributesXml({"w:val": "18"})},
                    {"w:lang": [], ...getAttributesXml({"w:val": "en-US"})}
                ]}
            ],
            ...getAttributesXml({"w:type": "paragraph", "w:styleId": listingParaId, "w:customStyle": "1"})
        })

        stylesTag["w:styles"].push({
            "w:style": [
                {"w:name": [], ...getAttributesXml({"w:val": "ispListing Знак"})},
                {"w:basedOn": [], ...getAttributesXml({"w:val": "DefaultParagraphFont"})},
                {"w:link": [], ...getAttributesXml({"w:val": listingParaId})},
                {"w:rPr": [
                    {"w:rFonts": [], ...getAttributesXml({"w:ascii": "Courier New", "w:hAnsi": "Courier New", "w:cs": "Courier New"})},
                    {"w:color": [], ...getAttributesXml({"w:val": "000000"})},
                    {"w:sz": [], ...getAttributesXml({"w:val": "18"})},
                    {"w:szCs": [], ...getAttributesXml({"w:val": "18"})},
                    {"w:lang": [], ...getAttributesXml({"w:val": "en-US"})}
                ]}
            ],
            ...getAttributesXml({"w:type": "character", "w:styleId": listingCharId, "w:customStyle": "1"})
        })

        styleNameToId.set("ispListing", listingParaId)
        styleNameToId.set("ispListing Знак", listingCharId)
        console.log("  Injected ispListing + ispListing Знак styles")
    }

    // ── Fix style defects in the official template ──

    // ispAnotation: replace autospacing with explicit spacing values
    // (The official template has w:beforeAutospacing="1" / w:afterAutospacing="1"
    // which produces unpredictable spacing. Replace with explicit 120/120.)
    for (let child of stylesTag["w:styles"]) {
        if (!child["w:style"]) continue
        let nameTag = getChildTag(child["w:style"], "w:name")
        if (!nameTag || !nameTag[xmlAttributes] || nameTag[xmlAttributes]["w:val"] !== "ispAnotation") continue

        let pPr = getChildTag(child["w:style"], "w:pPr")
        if (!pPr) break
        let spacing = getChildTag(pPr["w:pPr"], "w:spacing")
        if (spacing && spacing[xmlAttributes]) {
            delete spacing[xmlAttributes]["w:beforeAutospacing"]
            delete spacing[xmlAttributes]["w:afterAutospacing"]
            spacing[xmlAttributes]["w:before"] = "120"
            spacing[xmlAttributes]["w:after"] = "120"
            console.log("  Fixed ispAnotation: autospacing → explicit 120/120")
        }
        break
    }

    // ispHeader: add w:pageBreakBefore val="false" to override inherited value
    // (ispHeader is basedOn Heading1 which has w:pageBreakBefore, but titles
    // should not force a page break.)
    for (let child of stylesTag["w:styles"]) {
        if (!child["w:style"]) continue
        let nameTag = getChildTag(child["w:style"], "w:name")
        if (!nameTag || !nameTag[xmlAttributes] || nameTag[xmlAttributes]["w:val"] !== "ispHeader") continue

        let pPr = getChildTag(child["w:style"], "w:pPr")
        if (!pPr) {
            // Create pPr if missing
            child["w:style"].push({"w:pPr": [
                {"w:pageBreakBefore": [], ...getAttributesXml({"w:val": "false"})}
            ]})
        } else {
            pPr["w:pPr"] = pPr["w:pPr"].filter(
                (item: any) => getTagName(item) !== "w:pageBreakBefore"
            )
            pPr["w:pPr"].push(
                {"w:pageBreakBefore": [], ...getAttributesXml({"w:val": "false"})}
            )
        }
        console.log("  Fixed ispHeader: added pageBreakBefore=false")
        break
    }

    // Write modified styles back
    zip.file("word/styles.xml", xmlBuilder.build(stylesParsed))

    // ── Inject numId 80 for bibliography if missing ──

    let numXml = await zip.file("word/numbering.xml")!.async("string")
    let numParsed = xmlParser.parse(numXml)
    let numTag = numParsed.find((x: any) => x["w:numbering"])
    let numEntries = numTag["w:numbering"]

    // Check if numId 80 already exists
    let hasNumId80 = numEntries.some((e: any) =>
        e["w:num"] && e[xmlAttributes] && e[xmlAttributes]["w:numId"] === "80"
    )

    if (!hasNumId80) {
        // Find max abstractNumId to avoid conflicts
        let maxAbsId = 0
        for (let entry of numEntries) {
            if (entry["w:abstractNum"] && entry[xmlAttributes]) {
                let id = parseInt(entry[xmlAttributes]["w:abstractNumId"])
                if (id > maxAbsId) maxAbsId = id
            }
        }

        let newAbsId = String(maxAbsId + 1)

        // Add abstractNum for bibliography list: "[%1]." format, right-justified
        let absNum: any = {
            "w:abstractNum": [
                {"w:multiLevelType": [], ...getAttributesXml({"w:val": "hybridMultilevel"})}
            ],
            ...getAttributesXml({"w:abstractNumId": newAbsId})
        }

        // Level 0: bibliography format "[1]."
        let lvlFormats = [
            {ilvl: "0", text: "[%1].", jc: "right", left: "360", hanging: "72"},
        ]
        // Levels 1-8: standard decimal sub-levels
        for (let i = 1; i <= 8; i++) {
            let text = Array.from({length: i + 1}, (_, j) => `%${j + 1}`).join(".") + "."
            lvlFormats.push({
                ilvl: String(i), text, jc: "left",
                left: String(360 + i * 504), hanging: String(Math.min(72 + i * 144, 1440))
            })
        }

        for (let fmt of lvlFormats) {
            absNum["w:abstractNum"].push({
                "w:lvl": [
                    {"w:start": [], ...getAttributesXml({"w:val": "1"})},
                    {"w:numFmt": [], ...getAttributesXml({"w:val": "decimal"})},
                    {"w:suff": [], ...getAttributesXml({"w:val": "tab"})},
                    {"w:lvlText": [], ...getAttributesXml({"w:val": fmt.text})},
                    {"w:lvlJc": [], ...getAttributesXml({"w:val": fmt.jc})},
                    {"w:pPr": [
                        {"w:ind": [], ...getAttributesXml({"w:left": fmt.left, "w:hanging": fmt.hanging})},
                        {"w:tabs": [
                            {"w:tab": [], ...getAttributesXml({"w:val": "left", "w:pos": fmt.left, "w:leader": "none"})}
                        ]}
                    ]},
                    {"w:rPr": [
                        {"w:rFonts": [], ...getAttributesXml({"w:hint": "default"})}
                    ]}
                ],
                ...getAttributesXml({"w:ilvl": fmt.ilvl})
            })
        }

        // Insert abstractNum before the first w:num entry
        let firstNumIdx = numEntries.findIndex((e: any) => e["w:num"])
        if (firstNumIdx >= 0) {
            numEntries.splice(firstNumIdx, 0, absNum)
        } else {
            numEntries.push(absNum)
        }

        // Add w:num for numId 80 pointing to the new abstractNum
        let numOverrides: any[] = []
        for (let i = 0; i <= 8; i++) {
            numOverrides.push({
                "w:lvlOverride": [
                    {"w:startOverride": [], ...getAttributesXml({"w:val": "1"})}
                ],
                ...getAttributesXml({"w:ilvl": String(i)})
            })
        }

        numEntries.push({
            "w:num": [
                {"w:abstractNumId": [], ...getAttributesXml({"w:val": newAbsId})},
                ...numOverrides
            ],
            ...getAttributesXml({"w:numId": "80"})
        })

        console.log(`  Injected abstractNum ${newAbsId} + numId 80 for bibliography`)
    }

    zip.file("word/numbering.xml", xmlBuilder.build(numParsed))

    // ── Classify paragraphs using a sequential state machine ──

    let roles: ParagraphRole[] = new Array(body.length)

    // Phase 1: Find key landmark indices
    let ruTitleIdx = -1
    let firstRuAuthorIdx = -1
    let firstRuOrgIdx = -1
    let lastRuAnnotationIdx = -1
    let enTitleIdx = -1
    let firstEnAuthorIdx = -1
    let firstEnOrgIdx = -1
    let lastEnAnnotationIdx = -1
    let firstBodyIdx = -1
    let bibHeadingIdx = -1
    let firstLitListIdx = -1
    let lastLitListIdx = -1
    let authorInfoHeadingIdx = -1
    let firstAuthorDetailIdx = -1
    let lastAuthorDetailIdx = -1
    let sectPrIdx = -1

    // Find RU title
    for (let i = 0; i < body.length; i++) {
        if (getStyle(body[i]) === ispHeaderId) {
            ruTitleIdx = i
            break
        }
    }

    // Find RU authors and orgs (ispAuthor paragraphs after RU title)
    let ruAuthorSection: number[] = []
    for (let i = ruTitleIdx + 1; i < body.length; i++) {
        if (getStyle(body[i]) === ispAuthorId) {
            ruAuthorSection.push(i)
        } else {
            break
        }
    }

    if (ruAuthorSection.length > 0) {
        firstRuAuthorIdx = ruAuthorSection[0]
        // Find first org (spacing before="60")
        for (let idx of ruAuthorSection) {
            if (getSpacingBefore(body[idx]) === "60") {
                firstRuOrgIdx = idx
                break
            }
        }
    }

    // Find RU annotations (ispAnotation after author section)
    let ruAnnotationStart = ruAuthorSection.length > 0 ? ruAuthorSection[ruAuthorSection.length - 1] + 1 : ruTitleIdx + 1
    for (let i = ruAnnotationStart; i < body.length; i++) {
        let style = getStyle(body[i])
        if (style === ispAnotationId) {
            lastRuAnnotationIdx = i
        } else if (lastRuAnnotationIdx > 0) {
            break
        }
    }

    // Find EN title: first non-empty, non-ispAuthor paragraph after spacers
    // Spacers are empty paragraphs after last RU annotation
    let searchStart = lastRuAnnotationIdx + 1
    for (let i = searchStart; i < body.length; i++) {
        let text = getParagraphText(body[i]).trim()
        if (text === "") continue // skip spacers
        // This should be the EN title
        enTitleIdx = i
        break
    }

    // Find EN authors/orgs (ispAuthor paragraphs after EN title)
    let enAuthorSection: number[] = []
    for (let i = enTitleIdx + 1; i < body.length; i++) {
        if (getStyle(body[i]) === ispAuthorId) {
            enAuthorSection.push(i)
        } else {
            break
        }
    }

    if (enAuthorSection.length > 0) {
        firstEnAuthorIdx = enAuthorSection[0]
        for (let idx of enAuthorSection) {
            if (getSpacingBefore(body[idx]) === "60") {
                firstEnOrgIdx = idx
                break
            }
        }
    }

    // Find EN annotations: paragraphs after EN authors with annotation text prefixes
    let enAnnotationStart = enAuthorSection.length > 0 ? enAuthorSection[enAuthorSection.length - 1] + 1 : enTitleIdx + 1
    let enAnnotationPrefixes = ["Abstract.", "Keywords:", "For citation:", "Acknowledgements."]
    for (let i = enAnnotationStart; i < body.length; i++) {
        let text = getParagraphText(body[i]).trim()
        if (enAnnotationPrefixes.some(prefix => text.startsWith(prefix))) {
            lastEnAnnotationIdx = i
        } else if (lastEnAnnotationIdx > 0) {
            break
        }
    }

    // Body starts after EN annotations
    firstBodyIdx = lastEnAnnotationIdx + 1

    // Find bibliography: ispLitList paragraphs, and the heading before them
    for (let i = firstBodyIdx; i < body.length; i++) {
        if (getStyle(body[i]) === ispLitListId) {
            if (firstLitListIdx === -1) firstLitListIdx = i
            lastLitListIdx = i
        }
    }

    // Bibliography heading: look for "Список литературы" or ispSubHeader-1 level before first litlist
    if (firstLitListIdx >= 0) {
        for (let i = firstLitListIdx - 1; i >= firstBodyIdx; i--) {
            let text = getParagraphText(body[i]).trim()
            if (text.includes("Список литературы") || text === "References") {
                bibHeadingIdx = i
                break
            }
        }
    }

    // Find "Информация об авторах" heading and author details after bibliography
    let afterBibStart = lastLitListIdx >= 0 ? lastLitListIdx + 1 : firstBodyIdx
    for (let i = afterBibStart; i < body.length; i++) {
        let text = getParagraphText(body[i]).trim()
        if (text.includes("Информация об авторах") || text.includes("Information about authors")) {
            authorInfoHeadingIdx = i
            break
        }
    }

    // Author detail paragraphs: all paragraphs after the heading until sectPr
    if (authorInfoHeadingIdx >= 0) {
        for (let i = authorInfoHeadingIdx + 1; i < body.length; i++) {
            let tagName = getTagName(body[i])
            if (tagName === "w:sectPr") break
            if (tagName === "w:p") {
                if (firstAuthorDetailIdx === -1) firstAuthorDetailIdx = i
                lastAuthorDetailIdx = i
            }
        }
    }

    // Find sectPr
    for (let i = body.length - 1; i >= 0; i--) {
        if (getTagName(body[i]) === "w:sectPr") {
            sectPrIdx = i
            break
        }
    }

    if (ruTitleIdx < 0 || enTitleIdx < 0 || firstRuAuthorIdx < 0 || firstEnAuthorIdx < 0 ||
        lastRuAnnotationIdx < 0 || lastEnAnnotationIdx < 0 || firstLitListIdx < 0 ||
        authorInfoHeadingIdx < 0 || firstAuthorDetailIdx < 0 || sectPrIdx < 0) {
        throw new Error("Failed to find all landmark paragraphs. Check that the input is a valid ISP RAS proceedings template.")
    }

    // ── Phase 2: Assign roles ──

    // Helper sets for quick lookup
    let ruAuthorSet = new Set(ruAuthorSection)
    let enAuthorSet = new Set(enAuthorSection)

    for (let i = 0; i < body.length; i++) {
        let p = body[i]
        let text = getParagraphText(p).trim()

        // sectPr
        if (i === sectPrIdx) {
            roles[i] = { action: "sectPr" }
            continue
        }

        // DOI paragraph
        if (i === 0) {
            roles[i] = { action: "keep" }
            continue
        }

        // RU title
        if (i === ruTitleIdx) {
            roles[i] = { action: "replace_full", placeholder: "{{{header_ru}}}" }
            continue
        }

        // RU authors
        if (ruAuthorSet.has(i)) {
            if (i === firstRuAuthorIdx) {
                roles[i] = { action: "replace_full", placeholder: "{{{authors_ru}}}" }
            } else if (i === firstRuOrgIdx) {
                roles[i] = { action: "replace_full", placeholder: "{{{organizations_ru}}}" }
            } else {
                roles[i] = { action: "delete" }
            }
            continue
        }

        // RU annotations
        if (i > (ruAuthorSection.length > 0 ? ruAuthorSection[ruAuthorSection.length - 1] : ruTitleIdx) && i <= lastRuAnnotationIdx) {
            if (getStyle(p) === ispAnotationId) {
                if (text.startsWith("Аннотация.")) {
                    roles[i] = { action: "replace_annotation", prefix: "Аннотация. ", placeholder: "{{{abstract_ru}}}" }
                } else if (text.startsWith("Ключевые слова:")) {
                    roles[i] = { action: "replace_annotation", prefix: "Ключевые слова: ", placeholder: "{{{keywords_ru}}}" }
                } else if (text.startsWith("Для цитирования:")) {
                    roles[i] = { action: "replace_annotation", prefix: "Для цитирования: ", placeholder: "{{{for_citation_ru}}}" }
                } else if (text.startsWith("Благодарности:")) {
                    roles[i] = { action: "replace_annotation", prefix: "Благодарности: ", placeholder: "{{{acknowledgements_ru}}}" }
                } else {
                    roles[i] = { action: "keep" }
                }
            } else {
                roles[i] = { action: "keep" }
            }
            continue
        }

        // Spacers between RU and EN
        if (i > lastRuAnnotationIdx && i < enTitleIdx) {
            roles[i] = { action: "keep" }
            continue
        }

        // EN title
        if (i === enTitleIdx) {
            roles[i] = { action: "replace_full", placeholder: "{{{header_en}}}" }
            continue
        }

        // EN authors
        if (enAuthorSet.has(i)) {
            if (i === firstEnAuthorIdx) {
                roles[i] = { action: "replace_full", placeholder: "{{{authors_en}}}" }
            } else if (i === firstEnOrgIdx) {
                roles[i] = { action: "replace_full", placeholder: "{{{organizations_en}}}" }
            } else {
                roles[i] = { action: "delete" }
            }
            continue
        }

        // EN annotations
        if (i > (enAuthorSection.length > 0 ? enAuthorSection[enAuthorSection.length - 1] : enTitleIdx) && i <= lastEnAnnotationIdx) {
            if (text.startsWith("Abstract.")) {
                roles[i] = { action: "replace_annotation", prefix: "Abstract. ", placeholder: "{{{abstract_en}}}" }
            } else if (text.startsWith("Keywords:")) {
                roles[i] = { action: "replace_annotation", prefix: "Keywords: ", placeholder: "{{{keywords_en}}}" }
            } else if (text.startsWith("For citation:")) {
                roles[i] = { action: "replace_annotation", prefix: "For citation: ", placeholder: "{{{for_citation_en}}}" }
            } else if (text.startsWith("Acknowledgements.")) {
                roles[i] = { action: "replace_annotation", prefix: "Acknowledgements. ", placeholder: "{{{acknowledgements_en}}}" }
            } else {
                roles[i] = { action: "keep" }
            }
            continue
        }

        // Body content (firstBodyIdx to bibHeading or firstLitList)
        let bodyEnd = bibHeadingIdx >= 0 ? bibHeadingIdx : (firstLitListIdx >= 0 ? firstLitListIdx : (authorInfoHeadingIdx >= 0 ? authorInfoHeadingIdx : sectPrIdx))
        if (i >= firstBodyIdx && i < bodyEnd) {
            if (i === firstBodyIdx) {
                roles[i] = { action: "body_placeholder" }
            } else {
                roles[i] = { action: "delete" }
            }
            continue
        }

        // Bibliography heading + intro text
        if (bibHeadingIdx >= 0 && i >= bibHeadingIdx && i < firstLitListIdx) {
            roles[i] = { action: "delete" }
            continue
        }

        // Bibliography entries
        if (firstLitListIdx >= 0 && i >= firstLitListIdx && i <= lastLitListIdx) {
            if (i === firstLitListIdx) {
                roles[i] = { action: "links_placeholder" }
            } else {
                roles[i] = { action: "delete" }
            }
            continue
        }

        // Content between bibliography and author info (if any)
        if (lastLitListIdx >= 0 && authorInfoHeadingIdx >= 0 && i > lastLitListIdx && i < authorInfoHeadingIdx) {
            roles[i] = { action: "delete" }
            continue
        }

        // Author info heading
        if (i === authorInfoHeadingIdx) {
            roles[i] = { action: "keep" }
            continue
        }

        // Author detail paragraphs
        if (firstAuthorDetailIdx >= 0 && i >= firstAuthorDetailIdx && i <= lastAuthorDetailIdx) {
            if (i === firstAuthorDetailIdx) {
                roles[i] = { action: "replace_full", placeholder: "{{{authors_detail}}}" }
            } else {
                roles[i] = { action: "delete" }
            }
            continue
        }

        // Fallback
        roles[i] = { action: "keep" }
    }

    // Count actions
    let counts = new Map<string, number>()
    for (let role of roles) {
        counts.set(role.action, (counts.get(role.action) || 0) + 1)
    }
    console.log(`  Paragraphs: ${roles.length} total, ${counts.get("keep") || 0} kept, ${counts.get("delete") || 0} deleted, ` +
        `${(counts.get("replace_full") || 0) + (counts.get("replace_annotation") || 0)} replaced, ` +
        `${(counts.get("body_placeholder") || 0) + (counts.get("links_placeholder") || 0)} placeholders`)

    // ── Build new body ──

    let newBody: any[] = []

    for (let i = 0; i < body.length; i++) {
        let role = roles[i]
        let p = body[i]

        switch (role.action) {
            case "keep":
            case "sectPr":
                newBody.push(JSON.parse(JSON.stringify(p)))
                break

            case "delete":
                break

            case "replace_full": {
                let clone = JSON.parse(JSON.stringify(p))
                replaceAllRunsWithPlaceholder(clone, role.placeholder)
                // Strip spacing from author/org paragraphs — the converter
                // adds correct spacing at runtime via addParagraphSpacing().
                // Baked-in spacing would be cloned to every author instance.
                if (role.placeholder.includes("authors_") || role.placeholder.includes("organizations_")) {
                    stripParagraphSpacing(clone)
                }
                newBody.push(clone)
                break
            }

            case "replace_annotation": {
                let clone = JSON.parse(JSON.stringify(p))
                replaceAnnotationValue(clone, role.prefix, role.placeholder)
                newBody.push(clone)
                break
            }

            case "body_placeholder":
                newBody.push(makePlainParagraph("{{{body}}}"))
                break

            case "links_placeholder":
                newBody.push(makePlainParagraph("{{{links}}}"))
                break
        }
    }

    // Verify placeholders
    let bodyText = newBody.map(p => getParagraphText(p)).join("\n")
    let required = [
        "{{{header_ru}}}", "{{{authors_ru}}}", "{{{organizations_ru}}}",
        "{{{abstract_ru}}}", "{{{keywords_ru}}}", "{{{for_citation_ru}}}", "{{{acknowledgements_ru}}}",
        "{{{header_en}}}", "{{{authors_en}}}", "{{{organizations_en}}}",
        "{{{abstract_en}}}", "{{{keywords_en}}}", "{{{for_citation_en}}}", "{{{acknowledgements_en}}}",
        "{{{body}}}", "{{{links}}}", "{{{authors_detail}}}"
    ]
    let allOk = true
    for (let placeholder of required) {
        if (!bodyText.includes(placeholder)) {
            console.error(`MISSING: ${placeholder}`)
            allOk = false
        }
    }
    if (allOk) {
        console.log("All 17 placeholders present.")
    }

    // Replace body in parsed document
    let docTag = docParsed.find((x: any) => x["w:document"])
    let bodyTag = docTag["w:document"].find((x: any) => x["w:body"])
    bodyTag["w:body"] = newBody

    // ── Patch headers ──

    let header1Xml = await zip.file("word/header1.xml")!.async("string")
    let header2Xml = await zip.file("word/header2.xml")!.async("string")
    let header3Xml = await zip.file("word/header3.xml")!.async("string")

    let header1Parsed = xmlParser.parse(header1Xml)
    let header2Parsed = xmlParser.parse(header2Xml)
    let header3Parsed = xmlParser.parse(header3Xml)

    // header1 = English page header, header2 = Russian page header
    replaceHeaderContent(header1Parsed, "{{{page_header_en}}}")
    replaceHeaderContent(header2Parsed, "{{{page_header_ru}}}")
    // header3 = first-page header — keep as-is

    zip.file("word/header1.xml", xmlBuilder.build(header1Parsed))
    zip.file("word/header2.xml", xmlBuilder.build(header2Parsed))
    zip.file("word/header3.xml", xmlBuilder.build(header3Parsed))

    // Write modified document.xml
    zip.file("word/document.xml", xmlBuilder.build(docParsed))

    // ── Clean up relationships ──
    // Remove hyperlinks and extra images from document.xml.rels
    // (body content was deleted, so those rIds are no longer referenced)

    let relsXml = await zip.file("word/_rels/document.xml.rels")!.async("string")
    let relsParsed = xmlParser.parse(relsXml)
    let relsTag = relsParsed.find((x: any) => x["Relationships"])
    let relsEntries = relsTag["Relationships"]

    // Filter: keep only structural rels (not hyperlinks or extra images)
    let keepTypes = new Set([
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles",
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/header",
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer",
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering",
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/fontTable",
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings",
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme",
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/webSettings",
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes",
    ])

    let imageRel = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image"
    let oldToNewId = new Map<string, string>()
    let filteredRels: any[] = []
    let nextId = 1

    for (let rel of relsEntries) {
        let attrs = rel[xmlAttributes]
        if (!attrs) continue // skip text nodes

        let type = attrs["Type"]
        let target = attrs["Target"]

        // Keep structural rels
        if (keepTypes.has(type)) {
            let newId = `rId${nextId++}`
            oldToNewId.set(attrs["Id"], newId)
            filteredRels.push({
                "Relationship": [],
                ...getAttributesXml({
                    "Id": newId,
                    "Type": type,
                    "Target": target
                })
            })
            continue
        }

        // Keep only image1.png (CC-BY logo used in DOI paragraph)
        if (type === imageRel && target === "media/image1.png") {
            let newId = `rId${nextId++}`
            oldToNewId.set(attrs["Id"], newId)
            filteredRels.push({
                "Relationship": [],
                ...getAttributesXml({
                    "Id": newId,
                    "Type": type,
                    "Target": target
                })
            })
            continue
        }

        // Skip hyperlinks, extra images, etc.
    }

    // Add webSettings and footnotes rels if not already present
    let hasWebSettings = filteredRels.some(r => r[xmlAttributes]?.["Target"] === "webSettings.xml")
    let hasFootnotes = filteredRels.some(r => r[xmlAttributes]?.["Target"] === "footnotes.xml")

    if (!hasWebSettings) {
        filteredRels.push({
            "Relationship": [],
            ...getAttributesXml({
                "Id": `rId${nextId++}`,
                "Type": "http://schemas.openxmlformats.org/officeDocument/2006/relationships/webSettings",
                "Target": "webSettings.xml"
            })
        })
    }
    if (!hasFootnotes) {
        filteredRels.push({
            "Relationship": [],
            ...getAttributesXml({
                "Id": `rId${nextId++}`,
                "Type": "http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes",
                "Target": "footnotes.xml"
            })
        })
    }

    relsTag["Relationships"] = filteredRels
    zip.file("word/_rels/document.xml.rels", xmlBuilder.build(relsParsed))

    // Update rId references in document.xml (DOI paragraph has image ref, sectPr has header/footer refs)
    let updatedDocXml = xmlBuilder.build(docParsed)
    for (let [oldId, newId] of oldToNewId) {
        if (oldId !== newId) {
            updatedDocXml = updatedDocXml.split(`"${oldId}"`).join(`"${newId}"`)
        }
    }
    zip.file("word/document.xml", updatedDocXml)

    // ── Add missing files ──

    // Minimal webSettings.xml
    if (!zip.file("word/webSettings.xml")) {
        zip.file("word/webSettings.xml",
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
            '<w:webSettings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">' +
            '<w:optimizeForBrowser/></w:webSettings>')
    }

    // Minimal footnotes.xml
    if (!zip.file("word/footnotes.xml")) {
        zip.file("word/footnotes.xml",
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
            '<w:footnotes xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" ' +
            'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" ' +
            'xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" ' +
            'xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" ' +
            'xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml" ' +
            'mc:Ignorable="w14 w15">' +
            '<w:footnote w:type="separator" w:id="-1"><w:p><w:r><w:separator/></w:r></w:p></w:footnote>' +
            '<w:footnote w:type="continuationSeparator" w:id="0"><w:p><w:r><w:continuationSeparator/></w:r></w:p></w:footnote>' +
            '</w:footnotes>')
    }

    // Empty .rels for headers and footers
    let emptyRels = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>'

    for (let prefix of ["header", "footer"]) {
        for (let i = 1; i <= 3; i++) {
            let relsPath = `word/_rels/${prefix}${i}.xml.rels`
            if (!zip.file(relsPath)) {
                zip.file(relsPath, emptyRels)
            }
        }
    }
    // Also add footnotes.xml.rels
    if (!zip.file("word/_rels/footnotes.xml.rels")) {
        zip.file("word/_rels/footnotes.xml.rels", emptyRels)
    }

    // Remove extra files from the body content that was deleted
    zip.remove("word/media/image2.png")

    // ── Fix [Content_Types].xml ──
    // Ensure webSettings and footnotes overrides exist

    let contentTypesXml = await zip.file("[Content_Types].xml")!.async("string")
    let contentTypesParsed = xmlParser.parse(contentTypesXml)
    let typesTag = contentTypesParsed.find((x: any) => x["Types"])
    let typesEntries = typesTag["Types"]

    // Remove text nodes
    typesEntries = typesEntries.filter((t: any) => t[xmlAttributes] !== undefined || t["Default"] !== undefined || t["Override"] !== undefined)

    let overrideParts = new Set<string>()
    for (let entry of typesEntries) {
        if (entry["Override"] !== undefined && entry[xmlAttributes]) {
            overrideParts.add(entry[xmlAttributes]["PartName"])
        }
    }

    let neededOverrides = [
        {partName: "/word/webSettings.xml", contentType: "application/vnd.openxmlformats-officedocument.wordprocessingml.webSettings+xml"},
        {partName: "/word/footnotes.xml", contentType: "application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml"},
    ]

    for (let {partName, contentType} of neededOverrides) {
        if (!overrideParts.has(partName)) {
            typesEntries.push({
                "Override": [],
                ...getAttributesXml({
                    "PartName": partName,
                    "ContentType": contentType
                })
            })
        }
    }

    typesTag["Types"] = typesEntries
    zip.file("[Content_Types].xml", xmlBuilder.build(contentTypesParsed))

    // ── Save output ──

    console.log(`Writing ${outputPath}...`)
    let output = await zip.generateAsync({type: "uint8array"})
    fs.writeFileSync(outputPath, output)
    console.log("Done!")
}

// ── CLI ──

let args = process.argv.slice(2)
if (args.length !== 2) {
    console.error("Usage: node scripts/generate-reference.js <input.docx> <output.docx>")
    process.exit(1)
}

generateReference(args[0], args[1]).catch(err => {
    console.error("Error:", err)
    process.exit(1)
})
