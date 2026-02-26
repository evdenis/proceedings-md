"use strict";
Object.defineProperty(exports, "__esModule", { value: true });
exports.getMetaString = getMetaString;
exports.convertMetaToObject = convertMetaToObject;
exports.markdownToJson = markdownToJson;
exports.jsonToDocx = jsonToDocx;
const child_process_1 = require("child_process");
const pandocFlags = ["--tab-stop=8"];
function pandoc(src, args) {
    return new Promise((resolve, reject) => {
        let stdout = "";
        let stderr = "";
        let pandocProcess = (0, child_process_1.spawn)('pandoc', args);
        pandocProcess.on('error', (err) => {
            reject(new Error(`Failed to start pandoc: ${err.message}. Is pandoc installed?`));
        });
        pandocProcess.stdin.on('error', () => {
            // Ignore stdin errors — the process 'error' or 'exit' handler will report the failure
        });
        pandocProcess.stdout.on('data', (data) => {
            stdout += data;
        });
        pandocProcess.stderr.on('data', (data) => {
            stderr += data;
        });
        pandocProcess.on('exit', function (code) {
            if (stderr.length) {
                console.error("There was some pandoc warnings along the way:");
                console.error(stderr);
            }
            if (code === 0) {
                resolve(stdout);
            }
            else {
                reject(new Error(`Pandoc returned non-zero exit code: ${code}`));
            }
        });
        pandocProcess.stdin.end(src, 'utf-8');
    });
}
function getMetaString(value) {
    if (Array.isArray(value)) {
        let result = "";
        for (let component of value) {
            result += getMetaString(component);
        }
        return result;
    }
    if (typeof value !== "object" || !value.t) {
        return "";
    }
    if (value.t === "Str") {
        return value.c;
    }
    if (value.t === "Strong") {
        return "__" + getMetaString(value.c) + "__";
    }
    if (value.t === "Emph") {
        return "_" + getMetaString(value.c) + "_";
    }
    if (value.t === "Cite") {
        return getMetaString(value.c[1]);
    }
    if (value.t === "Space") {
        return " ";
    }
    if (value.t === "Link") {
        return getMetaString(value.c[1]);
    }
    return getMetaString(value.c);
}
function convertMetaToJsonRecursive(meta) {
    if (meta.t === "MetaList") {
        return meta.c.map((element) => {
            return convertMetaToJsonRecursive(element);
        });
    }
    if (meta.t === "MetaMap") {
        let result = {};
        for (let key of Object.getOwnPropertyNames(meta.c)) {
            result[key] = convertMetaToJsonRecursive(meta.c[key]);
        }
        return result;
    }
    if (meta.t === "MetaInlines") {
        return getMetaString(meta.c);
    }
    return undefined;
}
function convertMetaToObject(meta) {
    let result = {};
    for (let key of Object.getOwnPropertyNames(meta)) {
        result[key] = convertMetaToJsonRecursive(meta[key]);
    }
    return result;
}
async function markdownToJson(markdown) {
    let json = await pandoc(markdown, ["-f", "markdown", "-t", "json", ...pandocFlags]);
    return JSON.parse(json);
}
async function jsonToDocx(ast, target) {
    await pandoc(JSON.stringify(ast), ["-f", "json", "-t", "docx", "-o", target]);
}
