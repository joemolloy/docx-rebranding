// Build step: inline rebrand-core.mjs (the source of truth for document logic)
// into index.html between the CORE markers, stripping ES `export` keywords so it
// runs inside the page's IIFE, then mirror the result to rebrand-prototype.html.
// Run after editing rebrand-core.mjs:  node build.mjs
import { readFileSync, writeFileSync } from "node:fs";
import { fileURLToPath } from "node:url";
import { dirname, join } from "node:path";

const DIR = dirname(fileURLToPath(import.meta.url));
const core = readFileSync(join(DIR, "rebrand-core.mjs"), "utf8");

// Drop the leading module-doc comment lines, keep the code; strip `export`.
const codeLines = core.split("\n");
while (codeLines.length && /^\s*(\/\/|$)/.test(codeLines[0])) codeLines.shift();
const body = codeLines.join("\n")
  .replace(/^(\s*)export (async function |function |const )/gm, "$1$2")
  .replace(/\s+$/, "");

const START = "/*CORE_START*/";
const END = "/*CORE_END*/";
const inlined = `${START}\n${body}\n    ${END}`;

const htmlPath = join(DIR, "index.html");
const html = readFileSync(htmlPath, "utf8");
const s = html.indexOf(START);
const e = html.indexOf(END);
if (s === -1 || e === -1) throw new Error("CORE markers not found in index.html");
const before = html.slice(0, s);
const after = html.slice(e + END.length);
const out = before + inlined + after;

writeFileSync(htmlPath, out);
writeFileSync(join(DIR, "rebrand-prototype.html"), out);
console.log("build: index.html + rebrand-prototype.html regenerated from rebrand-core.mjs");
