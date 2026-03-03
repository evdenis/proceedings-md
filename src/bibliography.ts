import { BibFilePresenter, BibEntry } from 'bibtex'

export interface CitationResult {
    citedKeys: string[]                    // keys in first-citation order (original case)
    keyToNumber: Map<string, number>       // lowercase key → [N]
}

// Map of common LaTeX accent commands to Unicode characters
const latexAccents: Record<string, Record<string, string>> = {
    "'": { "o": "ó", "e": "é", "a": "á", "i": "í", "u": "ú", "O": "Ó", "E": "É", "A": "Á", "I": "Í", "U": "Ú", "c": "ć", "n": "ń", "s": "ś", "z": "ź", "y": "ý" },
    "`": { "a": "à", "e": "è", "i": "ì", "o": "ò", "u": "ù", "A": "À", "E": "È", "I": "Ì", "O": "Ò", "U": "Ù" },
    "^": { "a": "â", "e": "ê", "i": "î", "o": "ô", "u": "û", "A": "Â", "E": "Ê", "I": "Î", "O": "Ô", "U": "Û" },
    "\"": { "a": "ä", "e": "ë", "i": "ï", "o": "ö", "u": "ü", "A": "Ä", "E": "Ë", "I": "Ï", "O": "Ö", "U": "Ü" },
    "~": { "a": "ã", "n": "ñ", "o": "õ", "A": "Ã", "N": "Ñ", "O": "Õ" },
    "c": { "c": "ç", "C": "Ç", "s": "ş", "S": "Ş" },
    "v": { "c": "č", "s": "š", "z": "ž", "C": "Č", "S": "Š", "Z": "Ž", "r": "ř", "R": "Ř" },
}

/** Convert common LaTeX accent commands to Unicode */
function delatex(str: string): string {
    // Handle {\'o}, {\`e}, {\"u} etc.
    return str.replace(/\{?\\(['`^"~cv])\{?([a-zA-Z])\}?\}?/g, (match, cmd, letter) => {
        let map = latexAccents[cmd]
        if (map && map[letter]) return map[letter]
        return letter
    })
}

/** Normalize pages: convert BibTeX `--` to single `-` */
function normalizePages(pages: string): string {
    return pages.replace(/--/g, '-')
}

/** Reformat date from YYYY-MM-DD to DD.MM.YYYY */
function formatDate(dateStr: string): string {
    let parts = dateStr.split('-')
    if (parts.length === 3) {
        return parts[2] + '.' + parts[1] + '.' + parts[0]
    }
    return dateStr
}

/** Get a field value as a string, or undefined if missing */
function getField(entry: BibEntry, fieldName: string): string | undefined {
    let val = entry.getFieldAsString(fieldName)
    if (val === undefined || val === null) return undefined
    return delatex(String(val))
}

/** Format an author name as "Surname I. I." or just the corporate name */
function formatAuthorName(author: { firstNames: string[], lastNames: string[], vons: string[], jrs: string[] }): string {
    let lastParts = author.lastNames.map(n => delatex(n))
    let vonParts = author.vons.map(n => delatex(n))

    let surname = ""
    if (vonParts.length > 0) {
        surname = vonParts.join(' ') + ' '
    }
    surname += lastParts.join(' ')

    if (author.jrs.length > 0) {
        surname += ' ' + author.jrs.join(' ')
    }

    // Corporate author: no first names
    if (author.firstNames.length === 0) {
        return surname
    }

    let initials = author.firstNames.map(name => {
        let clean = delatex(name)
        // Already an initial (e.g. "М.", "M.", "Ya.")
        if (clean.length <= 3 && clean.endsWith('.')) {
            return clean
        }
        // Single letter
        if (clean.length === 1) {
            return clean + '.'
        }
        // Full name - take first letter
        return clean.charAt(0) + '.'
    })

    return surname + ' ' + initials.join(' ')
}

/** Format all authors of an entry as "Surname I. I., Surname I. I." */
function formatAuthors(entry: BibEntry): string {
    let authors = entry.getAuthors()
    if (!authors || !authors.authors$ || authors.authors$.length === 0) {
        return ''
    }
    return authors.authors$.map(formatAuthorName).join(', ')
}

/** Join authors and title, handling the case where authors may be empty */
function authorsTitle(authors: string, title: string): string {
    if (authors) return authors + ' ' + title + '.'
    return title + '.'
}

/** Format a Russian article entry */
function formatArticleRu(entry: BibEntry): string {
    let authors = formatAuthors(entry)
    let title = getField(entry, 'title') || ''
    let journal = getField(entry, 'journaltitle') || ''
    let volume = getField(entry, 'volume')
    let number = getField(entry, 'number')
    let year = getField(entry, 'year') || ''
    let pages = getField(entry, 'pages')
    let doi = getField(entry, 'doi')

    let parts: string[] = []
    parts.push(authorsTitle(authors, title))

    let journalParts = [journal]
    if (volume) journalParts.push('том ' + volume)
    if (number) journalParts.push('вып. ' + number)
    journalParts.push(year + ' г.')
    if (pages) journalParts.push('стр. ' + normalizePages(pages))

    parts.push(journalParts.join(', ') + '.')
    if (doi) parts.push('DOI: ' + doi + '.')

    return parts.join(' ')
}

/** Format an English article entry */
function formatArticleEn(entry: BibEntry): string {
    let authors = formatAuthors(entry)
    let title = getField(entry, 'title') || ''
    let journal = getField(entry, 'journaltitle') || ''
    let volume = getField(entry, 'volume')
    let number = getField(entry, 'number')
    let year = getField(entry, 'year') || ''
    let pages = getField(entry, 'pages')
    let doi = getField(entry, 'doi')
    let addendum = getField(entry, 'addendum')
    let url = getField(entry, 'url')
    let urldate = getField(entry, 'urldate')

    let parts: string[] = []
    parts.push(authorsTitle(authors, title))

    let journalParts = [journal]
    if (volume) journalParts.push('vol. ' + volume)
    if (number) journalParts.push('issue ' + number)
    journalParts.push(year)
    if (pages) journalParts.push('pp. ' + normalizePages(pages))

    let journalStr = journalParts.join(', ')
    if (addendum) {
        journalStr += ' ' + addendum
    }
    journalStr += '.'

    parts.push(journalStr)
    if (doi) parts.push('DOI: ' + doi + '.')

    if (url && urldate) {
        parts.push('Available at: ' + url + ', accessed ' + formatDate(urldate) + '.')
    } else if (url) {
        parts.push('Available at: ' + url + '.')
    }

    return parts.join(' ')
}

/** Format an inproceedings entry */
function formatInproceedings(entry: BibEntry): string {
    let authors = formatAuthors(entry)
    let title = getField(entry, 'title') || ''
    let booktitle = getField(entry, 'booktitle') || ''
    let year = getField(entry, 'year') || ''
    let pages = getField(entry, 'pages')
    let doi = getField(entry, 'doi')
    let addendum = getField(entry, 'addendum')

    // Strip "Proceedings of the " prefix to avoid "In Proc. of the Proceedings of the ..."
    let procName = booktitle.replace(/^Proceedings of the /i, '')

    let parts: string[] = []
    parts.push(authorsTitle(authors, title))
    let procPart = 'In Proc. of the ' + procName + ', ' + year + '.'
    if (pages) procPart += ' pp. ' + normalizePages(pages) + '.'
    parts.push(procPart)

    if (addendum) parts[parts.length - 1] = parts[parts.length - 1].replace(/\.$/, ' ' + addendum + '.')
    if (doi) parts.push('DOI: ' + doi + '.')

    return parts.join(' ')
}

/** Format a Russian book entry */
function formatBookRu(entry: BibEntry): string {
    let authors = formatAuthors(entry)
    let title = getField(entry, 'title') || ''
    let publisher = getField(entry, 'publisher') || ''
    let location = getField(entry, 'location')
    let year = getField(entry, 'year') || ''
    let pagetotal = getField(entry, 'pagetotal')

    let parts: string[] = []
    parts.push(authorsTitle(authors, title))

    let pubParts: string[] = []
    if (location) pubParts.push(location)
    pubParts.push(publisher)
    pubParts.push(year)
    if (pagetotal) pubParts.push(pagetotal + ' c.')

    let pubStr = pubParts.join(', ')
    if (!pubStr.endsWith('.')) pubStr += '.'
    parts.push(pubStr)

    return parts.join(' ')
}

/** Format an English book entry */
function formatBookEn(entry: BibEntry): string {
    let authors = formatAuthors(entry)
    let title = getField(entry, 'title') || ''
    let publisher = getField(entry, 'publisher') || ''
    let location = getField(entry, 'location')
    let year = getField(entry, 'year') || ''
    let pagetotal = getField(entry, 'pagetotal')
    let addendum = getField(entry, 'addendum')

    let parts: string[] = []
    parts.push(authorsTitle(authors, title))

    let pubParts: string[] = []
    if (location) pubParts.push(location)
    pubParts.push(publisher)
    pubParts.push(year)
    let pubStr = pubParts.join(', ') + '.'
    if (pagetotal) pubStr += ' ' + pagetotal + ' p.'
    if (addendum) pubStr += ' ' + addendum + '.'

    parts.push(pubStr)

    return parts.join(' ')
}

/** Format an online entry */
function formatOnline(entry: BibEntry): string {
    let authors = formatAuthors(entry)
    let title = getField(entry, 'title') || ''
    let url = getField(entry, 'url') || ''
    let urldate = getField(entry, 'urldate')
    let year = getField(entry, 'year')

    let parts: string[] = []
    parts.push(authorsTitle(authors, title))
    if (year) parts.push(year + '.')
    if (url && urldate) {
        parts.push('Available at: ' + url + ', accessed ' + formatDate(urldate) + '.')
    } else if (url) {
        parts.push('Available at: ' + url + '.')
    }

    return parts.join(' ')
}

/** Format a single entry using its language to pick the right formatter */
function formatEntry(entry: BibEntry): string {
    let langid = getField(entry, 'langid') || 'english'
    let isRussian = langid === 'russian'
    let type = entry.type

    if (type === 'article') {
        return isRussian ? formatArticleRu(entry) : formatArticleEn(entry)
    } else if (type === 'inproceedings') {
        return formatInproceedings(entry)
    } else if (type === 'book') {
        return isRussian ? formatBookRu(entry) : formatBookEn(entry)
    } else if (type === 'online') {
        return formatOnline(entry)
    }

    // Fallback: generic format
    let authors = formatAuthors(entry)
    let title = getField(entry, 'title') || ''
    return authorsTitle(authors, title)
}

/**
 * Resolve the primary key for a citation:
 * - If the entry has relatedtype=translationof, return its related (the Russian primary)
 * - Otherwise return the key itself
 */
function resolvePrimaryKey(key: string, bibFile: BibFilePresenter): string {
    let entry = bibFile.getEntry(key)
    if (!entry) return key
    let relatedType = entry.getFieldAsString('relatedtype')
    if (relatedType === 'translationof') {
        let related = entry.getFieldAsString('related')
        if (related) return String(related)
    }
    return key
}

/** Walk Pandoc AST, replace Cite nodes with [N] text, return key→number mapping */
export function resolveCitations(ast: any, bibFile: BibFilePresenter): CitationResult {
    let citedKeys: string[] = []          // original-case keys in first-citation order
    let keyToNumber = new Map<string, number>()  // lowercase key → number

    function assignNumber(citationId: string): number {
        let primaryKey = resolvePrimaryKey(citationId, bibFile)
        let lowerKey = primaryKey.toLowerCase()

        if (keyToNumber.has(lowerKey)) {
            return keyToNumber.get(lowerKey)!
        }

        let num = citedKeys.length + 1
        citedKeys.push(primaryKey)
        keyToNumber.set(lowerKey, num)
        return num
    }

    function walk(element: any): any {
        if (Array.isArray(element)) {
            return element.map(walk)
        }

        if (typeof element !== 'object' || !element) {
            return element
        }

        if (element.t === 'Cite') {
            let citations = element.c[0]
            let numbers = citations.map((cite: any) => assignNumber(cite.citationId))

            if (numbers.length === 1) {
                return { t: 'Str', c: '[' + numbers[0] + ']' }
            } else {
                return { t: 'Str', c: '[' + numbers.join(', ') + ']' }
            }
        }

        for (let key of Object.getOwnPropertyNames(element)) {
            element[key] = walk(element[key])
        }

        return element
    }

    ast.blocks = walk(ast.blocks)
    return { citedKeys, keyToNumber }
}

/** Format cited entries in order into ISPRAS house-style text strings */
export function formatBibliography(citedKeys: string[], bibFile: BibFilePresenter): string[] {
    return citedKeys.map(key => {
        let entry = bibFile.getEntry(key)
        if (!entry) {
            console.warn(`Warning: bibliography key "${key}" not found in .bib file`)
            return key
        }

        let relatedType = entry.getFieldAsString('relatedtype')
        let relatedKey = entry.getFieldAsString('related')

        // Bilingual pair: this is the Russian primary, related is the English translation
        if (relatedType === 'translationas' && relatedKey) {
            let enEntry = bibFile.getEntry(String(relatedKey))
            if (enEntry) {
                return formatEntry(entry) + ' / ' + formatEntry(enEntry)
            }
        }

        return formatEntry(entry)
    })
}
