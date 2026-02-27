import * as fs from 'fs';
import * as path from 'path';
import JSZip from 'jszip';
import {xmlParser, xmlAttributes, getChildTag, getTagName} from '../src/xml-helpers';

const docxPath = process.argv[2];
const refDocxPath = process.argv[3];
if (!docxPath || !refDocxPath) {
    console.error('Usage: node test/validate-docx.js <path-to-docx> <reference-docx>');
    process.exit(1);
}

interface CheckResult {
    name: string;
    passed: boolean;
    details: string;
}

const results: CheckResult[] = [];

function check(name: string, passed: boolean, details: string) {
    results.push({name, passed, details});
}

async function main() {
    const buf = fs.readFileSync(path.resolve(docxPath));
    const zip = await JSZip.loadAsync(buf);

    // ── Check 1: ZIP integrity ──────────────────────────────────────────
    const requiredFiles = [
        '[Content_Types].xml',
        'word/document.xml',
        'word/styles.xml',
        'word/numbering.xml',
        'word/settings.xml',
        'word/_rels/document.xml.rels',
    ];
    const fileNames = Object.keys(zip.files);
    const missingFiles = requiredFiles.filter(f => !fileNames.includes(f));
    check('ZIP integrity', missingFiles.length === 0,
        missingFiles.length === 0
            ? `All ${requiredFiles.length} required XML files present`
            : `Missing: ${missingFiles.join(', ')}`);

    // ── Check 2: XML well-formedness ────────────────────────────────────
    const xmlFiles = fileNames.filter(f => f.endsWith('.xml') || f.endsWith('.rels'));
    const parseErrors: string[] = [];
    const parsedXml: Record<string, any> = {};
    for (const f of xmlFiles) {
        try {
            const content = await zip.file(f)!.async('string');
            parsedXml[f] = xmlParser.parse(content);
        } catch (e: any) {
            parseErrors.push(`${f}: ${e.message}`);
        }
    }
    check('XML well-formedness', parseErrors.length === 0,
        parseErrors.length === 0
            ? `All ${xmlFiles.length} XML files parsed successfully`
            : `Parse errors:\n  ${parseErrors.join('\n  ')}`);

    // ── Check 3: Required ISP styles ────────────────────────────────────
    const requiredStyles = [
        'ispSubHeader-1 level',
        'ispSubHeader-2 level',
        'ispSubHeader-3 level',
        'ispAuthor',
        'ispAnotation',
        'ispText_main',
        'ispList1',
        'ispListing',
        'ispListing Знак',
        'ispLitList',
        'ispPicture_sign',
        'ispNumList',
        'Normal',
        'ispHeader',
    ];
    const stylesXml = parsedXml['word/styles.xml'];
    const styleIds = new Set<string>();
    const styleNames = new Set<string>();
    const basedOnRefs = new Set<string>();
    const linkRefs = new Set<string>();
    const nextRefs = new Set<string>();

    if (stylesXml) {
        const stylesRoot = getChildTag(stylesXml, 'w:styles');
        if (stylesRoot) {
            for (const child of stylesRoot['w:styles']) {
                if (child['w:style']) {
                    const attrs = child[xmlAttributes];
                    if (attrs && attrs['w:styleId']) {
                        styleIds.add(attrs['w:styleId']);
                    }
                    for (const prop of child['w:style']) {
                        if (prop['w:name']) {
                            const nameAttrs = prop[xmlAttributes];
                            if (nameAttrs && nameAttrs['w:val']) {
                                styleNames.add(nameAttrs['w:val']);
                            }
                        }
                        if (prop['w:basedOn']) {
                            const val = prop[xmlAttributes]?.['w:val'];
                            if (val) basedOnRefs.add(val);
                        }
                        if (prop['w:link']) {
                            const val = prop[xmlAttributes]?.['w:val'];
                            if (val) linkRefs.add(val);
                        }
                        if (prop['w:next']) {
                            const val = prop[xmlAttributes]?.['w:val'];
                            if (val) nextRefs.add(val);
                        }
                    }
                }
            }
        }
    }
    const missingStyles = requiredStyles.filter(s => !styleNames.has(s) && !styleIds.has(s));
    check('Required ISP styles', missingStyles.length === 0,
        missingStyles.length === 0
            ? `All ${requiredStyles.length} required styles found`
            : `Missing styles: ${missingStyles.join(', ')}`);

    // ── Check 4: Style hierarchy ────────────────────────────────────────
    // Built-in OOXML styles and converter-removed headings are allowed to be unresolved
    const knownMissing = new Set([
        'TableNormal',
        'Heading4', 'Heading5', 'Heading6', 'Heading7', 'Heading8', 'Heading9',
    ]);
    const unresolvedRefs: string[] = [];
    for (const ref of basedOnRefs) {
        if (!styleIds.has(ref) && !knownMissing.has(ref)) unresolvedRefs.push(`basedOn: ${ref}`);
    }
    for (const ref of linkRefs) {
        if (!styleIds.has(ref) && !knownMissing.has(ref)) unresolvedRefs.push(`link: ${ref}`);
    }
    for (const ref of nextRefs) {
        if (!styleIds.has(ref) && !knownMissing.has(ref)) unresolvedRefs.push(`next: ${ref}`);
    }
    check('Style hierarchy', unresolvedRefs.length === 0,
        unresolvedRefs.length === 0
            ? 'All basedOn/link/next references resolve'
            : `Unresolved: ${unresolvedRefs.join(', ')}`);

    // ── Check 5: Numbering definitions ──────────────────────────────────
    const requiredNumIds = [33, 43, 80];
    const numXml = parsedXml['word/numbering.xml'];
    const foundNumIds = new Set<number>();
    if (numXml) {
        const numRoot = getChildTag(numXml, 'w:numbering');
        if (numRoot) {
            for (const child of numRoot['w:numbering']) {
                if (child['w:num']) {
                    const attrs = child[xmlAttributes];
                    if (attrs && attrs['w:numId']) {
                        foundNumIds.add(parseInt(attrs['w:numId'], 10));
                    }
                }
            }
        }
    }
    const missingNums = requiredNumIds.filter(id => !foundNumIds.has(id));
    check('Numbering definitions', missingNums.length === 0,
        missingNums.length === 0
            ? `numId ${requiredNumIds.join(', ')} all present`
            : `Missing numIds: ${missingNums.join(', ')}`);

    // ── Check 6: No leftover placeholders ───────────────────────────────
    const placeholderPattern = /\{\{\{[^}]+\}\}\}/g;
    const filesWithPlaceholders: string[] = [];
    for (const f of xmlFiles) {
        const content = await zip.file(f)!.async('string');
        const matches = content.match(placeholderPattern);
        if (matches) {
            filesWithPlaceholders.push(`${f}: ${matches.join(', ')}`);
        }
    }
    check('No leftover placeholders', filesWithPlaceholders.length === 0,
        filesWithPlaceholders.length === 0
            ? 'No {{{...}}} placeholders found'
            : `Leftover placeholders:\n  ${filesWithPlaceholders.join('\n  ')}`);

    // ── Check 7: Document style references ──────────────────────────────
    const docXml = parsedXml['word/document.xml'];
    const usedDocStyles = new Set<string>();
    function collectStyleRefs(node: any) {
        if (!node || typeof node !== 'object') return;
        if (Array.isArray(node)) {
            for (const item of node) collectStyleRefs(item);
            return;
        }
        const attrs = node[xmlAttributes];
        if (attrs) {
            if (attrs['w:val'] !== undefined) {
                const tagName = getTagName(node);
                if (tagName === 'w:pStyle' || tagName === 'w:rStyle') {
                    usedDocStyles.add(attrs['w:val']);
                }
            }
        }
        for (const key of Object.keys(node)) {
            if (key === xmlAttributes) continue;
            collectStyleRefs(node[key]);
        }
    }
    collectStyleRefs(docXml);
    // Heading4+ are stripped by the converter but may be referenced by Pandoc output
    const unresolvedDocStyles = [...usedDocStyles].filter(s => !styleIds.has(s) && !knownMissing.has(s));
    check('Document style references', unresolvedDocStyles.length === 0,
        unresolvedDocStyles.length === 0
            ? `All ${usedDocStyles.size} style references in document resolve`
            : `Unresolved styles in document: ${unresolvedDocStyles.join(', ')}`);

    // ── Check 8: Relationship IDs ───────────────────────────────────────
    const relsContent = parsedXml['word/_rels/document.xml.rels'];
    const relsIds = new Set<string>();
    if (relsContent) {
        const relsRoot = getChildTag(relsContent, 'Relationships');
        if (relsRoot) {
            for (const child of relsRoot['Relationships']) {
                if (child['Relationship']) {
                    const attrs = child[xmlAttributes];
                    if (attrs && attrs['Id']) {
                        relsIds.add(attrs['Id']);
                    }
                }
            }
        }
    }
    const usedRelIds = new Set<string>();
    function collectRelIds(node: any) {
        if (!node || typeof node !== 'object') return;
        if (Array.isArray(node)) {
            for (const item of node) collectRelIds(item);
            return;
        }
        const attrs = node[xmlAttributes];
        if (attrs) {
            for (const key of ['r:id', 'r:embed']) {
                if (attrs[key]) usedRelIds.add(attrs[key]);
            }
        }
        for (const key of Object.keys(node)) {
            if (key === xmlAttributes) continue;
            collectRelIds(node[key]);
        }
    }
    collectRelIds(docXml);
    const unresolvedRels = [...usedRelIds].filter(id => !relsIds.has(id));
    check('Relationship IDs', unresolvedRels.length === 0,
        unresolvedRels.length === 0
            ? `All ${usedRelIds.size} relationship references resolve`
            : `Unresolved relationship IDs: ${unresolvedRels.join(', ')}`);

    // ── Check 9: Headers/footers ────────────────────────────────────────
    const headerFiles = fileNames.filter(f => /^word\/header\d+\.xml$/.test(f));
    const footerFiles = fileNames.filter(f => /^word\/footer\d+\.xml$/.test(f));
    const contentTypesXml = parsedXml['[Content_Types].xml'];
    let headerCTCount = 0;
    let footerCTCount = 0;
    if (contentTypesXml) {
        const typesRoot = getChildTag(contentTypesXml, 'Types');
        if (typesRoot) {
            for (const child of typesRoot['Types']) {
                if (child['Override']) {
                    const attrs = child[xmlAttributes];
                    if (attrs?.['ContentType']?.includes('header+xml')) headerCTCount++;
                    if (attrs?.['ContentType']?.includes('footer+xml')) footerCTCount++;
                }
            }
        }
    }
    const hfOk = headerFiles.length >= 3 && footerFiles.length >= 3 &&
                 headerCTCount >= 3 && footerCTCount >= 3;
    check('Headers/footers', hfOk,
        `${headerFiles.length} headers, ${footerFiles.length} footers in ZIP; ` +
        `${headerCTCount} header, ${footerCTCount} footer Content_Type entries`);

    // ── Check 10: Page layout ───────────────────────────────────────────
    let pgSzOk = false;
    let pgSzDetails = 'sectPr not found';
    function findSectPr(node: any): any {
        if (!node || typeof node !== 'object') return null;
        if (Array.isArray(node)) {
            for (const item of node) {
                const r = findSectPr(item);
                if (r) return r;
            }
            return null;
        }
        if (node['w:sectPr']) return node;
        for (const key of Object.keys(node)) {
            if (key === xmlAttributes) continue;
            const r = findSectPr(node[key]);
            if (r) return r;
        }
        return null;
    }
    const sectPrNode = findSectPr(docXml);
    if (sectPrNode) {
        const sectPr = sectPrNode['w:sectPr'];
        const pgSzTag = getChildTag(sectPr, 'w:pgSz');
        if (pgSzTag) {
            const attrs = pgSzTag[xmlAttributes];
            const w = attrs?.['w:w'];
            const h = attrs?.['w:h'];
            pgSzOk = w === '9360' && h === '13608';
            pgSzDetails = `w:w=${w}, w:h=${h} (expected 9360x13608)`;
        } else {
            pgSzDetails = 'w:pgSz not found in sectPr';
        }
    }
    check('Page layout', pgSzOk, pgSzDetails);

    // ── Check 11: sectPr header/footer refs ─────────────────────────────
    let sectPrRefsOk = false;
    let sectPrRefsDetails = 'sectPr not found';
    if (sectPrNode) {
        const sectPr = sectPrNode['w:sectPr'];
        const sectPrRelIds: string[] = [];
        for (const child of sectPr) {
            const tagName = getTagName(child);
            if (tagName === 'w:headerReference' || tagName === 'w:footerReference') {
                const attrs = child[xmlAttributes];
                if (attrs?.['r:id']) sectPrRelIds.push(attrs['r:id']);
            }
        }
        const unresolvedSectPrRels = sectPrRelIds.filter(id => !relsIds.has(id));
        sectPrRefsOk = sectPrRelIds.length > 0 && unresolvedSectPrRels.length === 0;
        sectPrRefsDetails = unresolvedSectPrRels.length === 0
            ? `All ${sectPrRelIds.length} header/footer refs in sectPr resolve`
            : `Unresolved sectPr refs: ${unresolvedSectPrRels.join(', ')}`;
    }
    check('sectPr header/footer refs', sectPrRefsOk, sectPrRefsDetails);

    // ── Check 12: ISP logo ──────────────────────────────────────────────
    const logoFile = zip.file('word/media/image1.png');
    let logoOk = false;
    let logoDetails = 'word/media/image1.png not found';
    if (logoFile) {
        const logoBytes = await logoFile.async('uint8array');
        const pngMagic = [0x89, 0x50, 0x4e, 0x47]; // \x89PNG
        const hasMagic = pngMagic.every((b, i) => logoBytes[i] === b);
        logoOk = hasMagic;
        logoDetails = hasMagic
            ? `image1.png present (${logoBytes.length} bytes, valid PNG header)`
            : `image1.png present but invalid PNG header`;
    }
    check('ISP logo', logoOk, logoDetails);

    // ── Load reference DOCX ─────────────────────────────────────────────
    const refBuf = fs.readFileSync(path.resolve(refDocxPath));
    const refZip = await JSZip.loadAsync(refBuf);
    const refParsedXml: Record<string, any> = {};
    const refXmlFiles = Object.keys(refZip.files).filter(f => f.endsWith('.xml') || f.endsWith('.rels'));
    for (const f of refXmlFiles) {
        try {
            const content = await refZip.file(f)!.async('string');
            refParsedXml[f] = xmlParser.parse(content);
        } catch {}
    }

    // Build reference style name→id and id→name maps
    const refStyleNames = new Set<string>();
    const refStyleIdToName = new Map<string, string>();
    const refStylesXml = refParsedXml['word/styles.xml'];
    if (refStylesXml) {
        const refStylesRoot = getChildTag(refStylesXml, 'w:styles');
        if (refStylesRoot) {
            for (const child of refStylesRoot['w:styles']) {
                if (child['w:style']) {
                    const attrs = child[xmlAttributes];
                    let sid = attrs?.['w:styleId'];
                    for (const prop of child['w:style']) {
                        if (prop['w:name']) {
                            const nameVal = prop[xmlAttributes]?.['w:val'];
                            if (nameVal) {
                                refStyleNames.add(nameVal);
                                if (sid) refStyleIdToName.set(sid, nameVal);
                            }
                        }
                    }
                }
            }
        }
    }

    // Build generated style id→name map
    const genStyleIdToName = new Map<string, string>();
    if (stylesXml) {
        const stylesRoot = getChildTag(stylesXml, 'w:styles');
        if (stylesRoot) {
            for (const child of stylesRoot['w:styles']) {
                if (child['w:style']) {
                    const attrs = child[xmlAttributes];
                    let sid = attrs?.['w:styleId'];
                    for (const prop of child['w:style']) {
                        if (prop['w:name']) {
                            const nameVal = prop[xmlAttributes]?.['w:val'];
                            if (nameVal && sid) genStyleIdToName.set(sid, nameVal);
                        }
                    }
                }
            }
        }
    }

    // ── Check 13: Style coverage ─────────────────────────────────────────
    // Collect style names used in official document.xml
    const refDocXml = refParsedXml['word/document.xml'];
    const refUsedStyleIds = new Set<string>();
    function collectStyleIds(node: any, target: Set<string>) {
        if (!node || typeof node !== 'object') return;
        if (Array.isArray(node)) {
            for (const item of node) collectStyleIds(item, target);
            return;
        }
        const a = node[xmlAttributes];
        if (a && a['w:val'] !== undefined) {
            const tagName = getTagName(node);
            if (tagName === 'w:pStyle' || tagName === 'w:rStyle') {
                target.add(a['w:val']);
            }
        }
        for (const key of Object.keys(node)) {
            if (key === xmlAttributes) continue;
            collectStyleIds(node[key], target);
        }
    }
    collectStyleIds(refDocXml, refUsedStyleIds);

    // Map reference style IDs to names
    const refUsedNames = new Set<string>();
    for (const id of refUsedStyleIds) {
        const name = refStyleIdToName.get(id);
        if (name) refUsedNames.add(name);
        else refUsedNames.add(id); // fallback to ID
    }

    // Map generated style IDs to names
    const genUsedNames = new Set<string>();
    for (const id of usedDocStyles) {
        const name = genStyleIdToName.get(id);
        if (name) genUsedNames.add(name);
        else genUsedNames.add(id);
    }

    // Styles the official document may use that the converter intentionally handles differently
    const knownMissingStyles = new Set([
        'Emphasis',            // converter uses direct italic formatting
        'Hyperlink',           // converter applies hyperlink formatting differently
        'ispAnotation2 Знак',  // character style variant not used by converter
        'текст1',              // text style variant in official only
        'ispUList',            // converter uses different list styling
        'Style',               // generic placeholder style
        'st',                  // abbreviated style in official
    ]);
    const missingStyleNames = [...refUsedNames].filter(n => !genUsedNames.has(n) && !knownMissingStyles.has(n));
    check('Style coverage', missingStyleNames.length === 0,
        missingStyleNames.length === 0
            ? `All ${refUsedNames.size} reference style names covered (${knownMissingStyles.size} known exclusions)`
            : `Missing style names: ${missingStyleNames.join(', ')}`);

    // ── Check 14: Paragraph count ────────────────────────────────────────
    function countParagraphs(node: any): number {
        if (!node || typeof node !== 'object') return 0;
        if (Array.isArray(node)) {
            let sum = 0;
            for (const item of node) sum += countParagraphs(item);
            return sum;
        }
        let count = 0;
        if (node['w:p']) count++;
        for (const key of Object.keys(node)) {
            if (key === xmlAttributes) continue;
            count += countParagraphs(node[key]);
        }
        return count;
    }

    const genBody = getChildTag(docXml, 'w:body');
    const refBody = getChildTag(refDocXml, 'w:body');
    const genPCount = genBody ? countParagraphs(genBody) : 0;
    const refPCount = refBody ? countParagraphs(refBody) : 0;
    const pCountRatio = refPCount > 0 ? Math.abs(genPCount - refPCount) / refPCount : 0;
    check('Paragraph count', pCountRatio <= 0.2,
        `Generated: ${genPCount}, Reference: ${refPCount} (${(pCountRatio * 100).toFixed(1)}% difference, max 20%)`);

    // ── Check 15: Style spacing values ───────────────────────────────────
    const spacingStyles = ['ispText_main', 'ispAnotation', 'ispAuthor'];

    function getStyleSpacing(stylesData: any, idToName: Map<string, string>, styleName: string): Record<string, string> | null {
        if (!stylesData) return null;
        const root = getChildTag(stylesData, 'w:styles');
        if (!root) return null;
        for (const child of root['w:styles']) {
            if (!child['w:style']) continue;
            let name: string | undefined;
            for (const prop of child['w:style']) {
                if (prop['w:name']) {
                    name = prop[xmlAttributes]?.['w:val'];
                }
            }
            if (name !== styleName) continue;
            for (const prop of child['w:style']) {
                if (prop['w:pPr']) {
                    for (const pPrChild of prop['w:pPr']) {
                        if (pPrChild['w:spacing']) {
                            const a = pPrChild[xmlAttributes] || {};
                            return {
                                'w:before': a['w:before'] || '',
                                'w:after': a['w:after'] || '',
                                'w:line': a['w:line'] || '',
                                'w:beforeAutospacing': a['w:beforeAutospacing'] || '',
                                'w:afterAutospacing': a['w:afterAutospacing'] || '',
                            };
                        }
                    }
                }
            }
            return {};
        }
        return null;
    }

    const spacingMismatches: string[] = [];
    for (const sn of spacingStyles) {
        const genSpacing = getStyleSpacing(stylesXml, genStyleIdToName, sn);
        const refSpacing = getStyleSpacing(refStylesXml, refStyleIdToName, sn);
        if (!genSpacing && !refSpacing) continue;
        if (!genSpacing) {
            spacingMismatches.push(`${sn}: missing in generated`);
            continue;
        }
        if (!refSpacing) continue; // style only in generated, OK
        for (const attr of ['w:before', 'w:after', 'w:line'] as const) {
            const gv = genSpacing[attr] || '';
            const rv = refSpacing[attr] || '';
            if (gv !== rv) {
                // Allow explicit spacing to replace autospacing (intentional fix)
                const autoAttr = attr === 'w:before' ? 'w:beforeAutospacing' : attr === 'w:after' ? 'w:afterAutospacing' : '';
                const validAutospacingReplacements = new Set(['120', '240']);
                if (autoAttr && (refSpacing[autoAttr] === '1') && gv && !rv
                    && validAutospacingReplacements.has(gv)) {
                    continue; // explicit value replaces autospacing — expected
                }
                spacingMismatches.push(`${sn} ${attr}: gen=${gv || '(unset)'} ref=${rv || '(unset)'}`);
            }
        }
    }
    check('Style spacing values', spacingMismatches.length === 0,
        spacingMismatches.length === 0
            ? `Spacing matches for ${spacingStyles.join(', ')}`
            : `Mismatches:\n  ${spacingMismatches.join('\n  ')}`);

    // ── Summary ─────────────────────────────────────────────────────────
    console.log('\n  DOCX Structural Validation\n');
    let passed = 0;
    let failed = 0;
    for (const r of results) {
        const icon = r.passed ? '\x1b[32mPASS\x1b[0m' : '\x1b[31mFAIL\x1b[0m';
        console.log(`  ${icon}  ${r.name}`);
        if (!r.passed) {
            console.log(`         ${r.details}`);
        }
        if (r.passed) passed++;
        else failed++;
    }
    console.log(`\n  ${passed} passed, ${failed} failed, ${results.length} total\n`);
    process.exit(failed > 0 ? 1 : 0);
}

main().catch(err => {
    console.error('Fatal error:', err);
    process.exit(1);
});
