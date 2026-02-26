"use strict";
Object.defineProperty(exports, "__esModule", { value: true });
exports.ReferenceResolver = void 0;
exports.resolveReferences = resolveReferences;
class ReferenceResolver {
    counters = new Map();
    labels = new Map();
    resolve(label) {
        if (this.labels.has(label)) {
            return String(this.labels.get(label));
        }
        let colonIndex = label.indexOf(":");
        let prefix = colonIndex !== -1 ? label.substring(0, colonIndex) : label;
        let count = (this.counters.get(prefix) || 0) + 1;
        this.counters.set(prefix, count);
        this.labels.set(label, count);
        return String(count);
    }
}
exports.ReferenceResolver = ReferenceResolver;
function resolveReferences(ast) {
    let resolver = new ReferenceResolver();
    function walk(element) {
        if (Array.isArray(element)) {
            return element.map(walk);
        }
        if (typeof element !== "object" || !element) {
            return element;
        }
        // Pandoc Span: {t: "Span", c: [[id, classes, attrs], inlines]}
        // We look for class "ref"
        if (element.t === "Span") {
            let attrs = element.c[0];
            let classes = attrs[1];
            if (classes && classes.includes("ref")) {
                let inlines = element.c[1];
                let label = inlinesToString(inlines);
                let number = resolver.resolve(label);
                return { t: "Str", c: number };
            }
        }
        for (let key of Object.getOwnPropertyNames(element)) {
            element[key] = walk(element[key]);
        }
        return element;
    }
    ast.blocks = walk(ast.blocks);
    return ast;
}
function inlinesToString(inlines) {
    let result = "";
    for (let inline of inlines) {
        if (inline.t === "Str") {
            result += inline.c;
        }
    }
    return result;
}
