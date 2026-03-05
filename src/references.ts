export class ReferenceResolver {
    private counters: Map<string, number> = new Map()
    private labels: Map<string, number> = new Map()

    resolve(label: string): string {
        if (this.labels.has(label)) {
            return String(this.labels.get(label))
        }

        let colonIndex = label.indexOf(":")
        let prefix = colonIndex !== -1 ? label.substring(0, colonIndex) : label

        let count = (this.counters.get(prefix) || 0) + 1
        this.counters.set(prefix, count)
        this.labels.set(label, count)

        return String(count)
    }
}

export function resolveReferences(ast: any): any {
    let resolver = new ReferenceResolver()

    function walk(element: any): any {
        if (Array.isArray(element)) {
            return element.map(walk)
        }

        if (typeof element !== "object" || !element) {
            return element
        }

        // Pandoc Cite: @ref:prefix:label → {t: "Cite", c: [[{citationId: "ref:prefix:label", ...}], [...]]}
        if (element.t === "Cite") {
            let citations = element.c[0]
            if (citations.length === 1 && citations[0].citationId.startsWith("ref:")) {
                let label = citations[0].citationId.substring(4) // strip "ref:"
                let number = resolver.resolve(label)
                return {t: "Str", c: number}
            }
        }

        for (let key of Object.getOwnPropertyNames(element)) {
            element[key] = walk(element[key])
        }

        return element
    }

    ast.blocks = walk(ast.blocks)
    return ast
}
