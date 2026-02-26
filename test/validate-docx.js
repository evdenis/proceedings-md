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
var __importDefault = (this && this.__importDefault) || function (mod) {
    return (mod && mod.__esModule) ? mod : { "default": mod };
};
Object.defineProperty(exports, "__esModule", { value: true });
const fs = __importStar(require("fs"));
const path = __importStar(require("path"));
const jszip_1 = __importDefault(require("jszip"));
const xml_helpers_1 = require("../src/xml-helpers");
const docxPath = process.argv[2];
if (!docxPath) {
    console.error('Usage: node test/validate-docx.js <path-to-docx>');
    process.exit(1);
}
const results = [];
function check(name, passed, details) {
    results.push({ name, passed, details });
}
async function main() {
    const buf = fs.readFileSync(path.resolve(docxPath));
    const zip = await jszip_1.default.loadAsync(buf);
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
    check('ZIP integrity', missingFiles.length === 0, missingFiles.length === 0
        ? `All ${requiredFiles.length} required XML files present`
        : `Missing: ${missingFiles.join(', ')}`);
    // ── Check 2: XML well-formedness ────────────────────────────────────
    const xmlFiles = fileNames.filter(f => f.endsWith('.xml') || f.endsWith('.rels'));
    const parseErrors = [];
    const parsedXml = {};
    for (const f of xmlFiles) {
        try {
            const content = await zip.file(f).async('string');
            parsedXml[f] = xml_helpers_1.xmlParser.parse(content);
        }
        catch (e) {
            parseErrors.push(`${f}: ${e.message}`);
        }
    }
    check('XML well-formedness', parseErrors.length === 0, parseErrors.length === 0
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
    ];
    const stylesXml = parsedXml['word/styles.xml'];
    const styleIds = new Set();
    const styleNames = new Set();
    const basedOnRefs = new Set();
    const linkRefs = new Set();
    const nextRefs = new Set();
    if (stylesXml) {
        const stylesRoot = (0, xml_helpers_1.getChildTag)(stylesXml, 'w:styles');
        if (stylesRoot) {
            for (const child of stylesRoot['w:styles']) {
                if (child['w:style']) {
                    const attrs = child[xml_helpers_1.xmlAttributes];
                    if (attrs && attrs['w:styleId']) {
                        styleIds.add(attrs['w:styleId']);
                    }
                    for (const prop of child['w:style']) {
                        if (prop['w:name']) {
                            const nameAttrs = prop[xml_helpers_1.xmlAttributes];
                            if (nameAttrs && nameAttrs['w:val']) {
                                styleNames.add(nameAttrs['w:val']);
                            }
                        }
                        if (prop['w:basedOn']) {
                            const val = prop[xml_helpers_1.xmlAttributes]?.['w:val'];
                            if (val)
                                basedOnRefs.add(val);
                        }
                        if (prop['w:link']) {
                            const val = prop[xml_helpers_1.xmlAttributes]?.['w:val'];
                            if (val)
                                linkRefs.add(val);
                        }
                        if (prop['w:next']) {
                            const val = prop[xml_helpers_1.xmlAttributes]?.['w:val'];
                            if (val)
                                nextRefs.add(val);
                        }
                    }
                }
            }
        }
    }
    const missingStyles = requiredStyles.filter(s => !styleNames.has(s) && !styleIds.has(s));
    check('Required ISP styles', missingStyles.length === 0, missingStyles.length === 0
        ? `All ${requiredStyles.length} required styles found`
        : `Missing styles: ${missingStyles.join(', ')}`);
    // ── Check 4: Style hierarchy ────────────────────────────────────────
    // Built-in OOXML styles and converter-removed headings are allowed to be unresolved
    const knownMissing = new Set([
        'TableNormal',
        'Heading4', 'Heading5', 'Heading6', 'Heading7', 'Heading8', 'Heading9',
    ]);
    const unresolvedRefs = [];
    for (const ref of basedOnRefs) {
        if (!styleIds.has(ref) && !knownMissing.has(ref))
            unresolvedRefs.push(`basedOn: ${ref}`);
    }
    for (const ref of linkRefs) {
        if (!styleIds.has(ref) && !knownMissing.has(ref))
            unresolvedRefs.push(`link: ${ref}`);
    }
    for (const ref of nextRefs) {
        if (!styleIds.has(ref) && !knownMissing.has(ref))
            unresolvedRefs.push(`next: ${ref}`);
    }
    check('Style hierarchy', unresolvedRefs.length === 0, unresolvedRefs.length === 0
        ? 'All basedOn/link/next references resolve'
        : `Unresolved: ${unresolvedRefs.join(', ')}`);
    // ── Check 5: Numbering definitions ──────────────────────────────────
    const requiredNumIds = [33, 43, 80];
    const numXml = parsedXml['word/numbering.xml'];
    const foundNumIds = new Set();
    if (numXml) {
        const numRoot = (0, xml_helpers_1.getChildTag)(numXml, 'w:numbering');
        if (numRoot) {
            for (const child of numRoot['w:numbering']) {
                if (child['w:num']) {
                    const attrs = child[xml_helpers_1.xmlAttributes];
                    if (attrs && attrs['w:numId']) {
                        foundNumIds.add(parseInt(attrs['w:numId'], 10));
                    }
                }
            }
        }
    }
    const missingNums = requiredNumIds.filter(id => !foundNumIds.has(id));
    check('Numbering definitions', missingNums.length === 0, missingNums.length === 0
        ? `numId ${requiredNumIds.join(', ')} all present`
        : `Missing numIds: ${missingNums.join(', ')}`);
    // ── Check 6: No leftover placeholders ───────────────────────────────
    const placeholderPattern = /\{\{\{[^}]+\}\}\}/g;
    const filesWithPlaceholders = [];
    for (const f of xmlFiles) {
        const content = await zip.file(f).async('string');
        const matches = content.match(placeholderPattern);
        if (matches) {
            filesWithPlaceholders.push(`${f}: ${matches.join(', ')}`);
        }
    }
    check('No leftover placeholders', filesWithPlaceholders.length === 0, filesWithPlaceholders.length === 0
        ? 'No {{{...}}} placeholders found'
        : `Leftover placeholders:\n  ${filesWithPlaceholders.join('\n  ')}`);
    // ── Check 7: Document style references ──────────────────────────────
    const docXml = parsedXml['word/document.xml'];
    const usedDocStyles = new Set();
    function collectStyleRefs(node) {
        if (!node || typeof node !== 'object')
            return;
        if (Array.isArray(node)) {
            for (const item of node)
                collectStyleRefs(item);
            return;
        }
        const attrs = node[xml_helpers_1.xmlAttributes];
        if (attrs) {
            if (attrs['w:val'] !== undefined) {
                const tagName = (0, xml_helpers_1.getTagName)(node);
                if (tagName === 'w:pStyle' || tagName === 'w:rStyle') {
                    usedDocStyles.add(attrs['w:val']);
                }
            }
        }
        for (const key of Object.keys(node)) {
            if (key === xml_helpers_1.xmlAttributes)
                continue;
            collectStyleRefs(node[key]);
        }
    }
    collectStyleRefs(docXml);
    // Heading4+ are stripped by the converter but may be referenced by Pandoc output
    const unresolvedDocStyles = [...usedDocStyles].filter(s => !styleIds.has(s) && !knownMissing.has(s));
    check('Document style references', unresolvedDocStyles.length === 0, unresolvedDocStyles.length === 0
        ? `All ${usedDocStyles.size} style references in document resolve`
        : `Unresolved styles in document: ${unresolvedDocStyles.join(', ')}`);
    // ── Check 8: Relationship IDs ───────────────────────────────────────
    const relsContent = parsedXml['word/_rels/document.xml.rels'];
    const relsIds = new Set();
    if (relsContent) {
        const relsRoot = (0, xml_helpers_1.getChildTag)(relsContent, 'Relationships');
        if (relsRoot) {
            for (const child of relsRoot['Relationships']) {
                if (child['Relationship']) {
                    const attrs = child[xml_helpers_1.xmlAttributes];
                    if (attrs && attrs['Id']) {
                        relsIds.add(attrs['Id']);
                    }
                }
            }
        }
    }
    const usedRelIds = new Set();
    function collectRelIds(node) {
        if (!node || typeof node !== 'object')
            return;
        if (Array.isArray(node)) {
            for (const item of node)
                collectRelIds(item);
            return;
        }
        const attrs = node[xml_helpers_1.xmlAttributes];
        if (attrs) {
            for (const key of ['r:id', 'r:embed']) {
                if (attrs[key])
                    usedRelIds.add(attrs[key]);
            }
        }
        for (const key of Object.keys(node)) {
            if (key === xml_helpers_1.xmlAttributes)
                continue;
            collectRelIds(node[key]);
        }
    }
    collectRelIds(docXml);
    const unresolvedRels = [...usedRelIds].filter(id => !relsIds.has(id));
    check('Relationship IDs', unresolvedRels.length === 0, unresolvedRels.length === 0
        ? `All ${usedRelIds.size} relationship references resolve`
        : `Unresolved relationship IDs: ${unresolvedRels.join(', ')}`);
    // ── Check 9: Headers/footers ────────────────────────────────────────
    const headerFiles = fileNames.filter(f => /^word\/header\d+\.xml$/.test(f));
    const footerFiles = fileNames.filter(f => /^word\/footer\d+\.xml$/.test(f));
    const contentTypesXml = parsedXml['[Content_Types].xml'];
    let headerCTCount = 0;
    let footerCTCount = 0;
    if (contentTypesXml) {
        const typesRoot = (0, xml_helpers_1.getChildTag)(contentTypesXml, 'Types');
        if (typesRoot) {
            for (const child of typesRoot['Types']) {
                if (child['Override']) {
                    const attrs = child[xml_helpers_1.xmlAttributes];
                    if (attrs?.['ContentType']?.includes('header+xml'))
                        headerCTCount++;
                    if (attrs?.['ContentType']?.includes('footer+xml'))
                        footerCTCount++;
                }
            }
        }
    }
    const hfOk = headerFiles.length >= 3 && footerFiles.length >= 3 &&
        headerCTCount >= 3 && footerCTCount >= 3;
    check('Headers/footers', hfOk, `${headerFiles.length} headers, ${footerFiles.length} footers in ZIP; ` +
        `${headerCTCount} header, ${footerCTCount} footer Content_Type entries`);
    // ── Check 10: Page layout ───────────────────────────────────────────
    let pgSzOk = false;
    let pgSzDetails = 'sectPr not found';
    function findSectPr(node) {
        if (!node || typeof node !== 'object')
            return null;
        if (Array.isArray(node)) {
            for (const item of node) {
                const r = findSectPr(item);
                if (r)
                    return r;
            }
            return null;
        }
        if (node['w:sectPr'])
            return node;
        for (const key of Object.keys(node)) {
            if (key === xml_helpers_1.xmlAttributes)
                continue;
            const r = findSectPr(node[key]);
            if (r)
                return r;
        }
        return null;
    }
    const sectPrNode = findSectPr(docXml);
    if (sectPrNode) {
        const sectPr = sectPrNode['w:sectPr'];
        const pgSzTag = (0, xml_helpers_1.getChildTag)(sectPr, 'w:pgSz');
        if (pgSzTag) {
            const attrs = pgSzTag[xml_helpers_1.xmlAttributes];
            const w = attrs?.['w:w'];
            const h = attrs?.['w:h'];
            pgSzOk = w === '9360' && h === '13608';
            pgSzDetails = `w:w=${w}, w:h=${h} (expected 9360x13608)`;
        }
        else {
            pgSzDetails = 'w:pgSz not found in sectPr';
        }
    }
    check('Page layout', pgSzOk, pgSzDetails);
    // ── Check 11: sectPr header/footer refs ─────────────────────────────
    let sectPrRefsOk = false;
    let sectPrRefsDetails = 'sectPr not found';
    if (sectPrNode) {
        const sectPr = sectPrNode['w:sectPr'];
        const sectPrRelIds = [];
        for (const child of sectPr) {
            const tagName = (0, xml_helpers_1.getTagName)(child);
            if (tagName === 'w:headerReference' || tagName === 'w:footerReference') {
                const attrs = child[xml_helpers_1.xmlAttributes];
                if (attrs?.['r:id'])
                    sectPrRelIds.push(attrs['r:id']);
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
        if (r.passed)
            passed++;
        else
            failed++;
    }
    console.log(`\n  ${passed} passed, ${failed} failed, ${results.length} total\n`);
    process.exit(failed > 0 ? 1 : 0);
}
main().catch(err => {
    console.error('Fatal error:', err);
    process.exit(1);
});
