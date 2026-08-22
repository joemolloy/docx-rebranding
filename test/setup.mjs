// Test harness bootstrap: wire the browser globals the core relies on to Node
// equivalents (xmldom for DOM, a node:zlib shim for fflate), then re-export the
// real rebrand-core.mjs so tests exercise the exact shipped logic.
import { DOMParser, XMLSerializer } from "@xmldom/xmldom";
import fflate from "./zip-shim.mjs";

globalThis.DOMParser = DOMParser;
globalThis.XMLSerializer = XMLSerializer;
globalThis.fflate = fflate;

// xmldom lacks the CSS-selector DOM methods the core uses in two spots:
//  - parse() probes doc.querySelector("parsererror"); xmldom throws on malformed
//    XML instead, so a null-returning shim is the correct equivalent.
//  - cloneChangelog() calls tr.querySelectorAll("*") to walk descendants when
//    regenerating w14 ids; back it with getElementsByTagName("*").
const sampleDoc = new DOMParser().parseFromString("<a><b/></a>", "text/xml");
const Document = sampleDoc.constructor;
const Element = sampleDoc.documentElement.constructor;
if (!Document.prototype.querySelector) Document.prototype.querySelector = () => null;
if (!Element.prototype.querySelectorAll)
  Element.prototype.querySelectorAll = function (sel) {
    return sel === "*" ? [...this.getElementsByTagName("*")] : [];
  };

// xmldom predates the modern ChildNode/ParentNode convenience methods the core
// relies on (present natively in the browser): back them with the classic API.
if (!Element.prototype.append)
  Element.prototype.append = function (...nodes) {
    for (const n of nodes)
      this.appendChild(typeof n === "string" ? this.ownerDocument.createTextNode(n) : n);
  };
if (!Element.prototype.remove)
  Element.prototype.remove = function () {
    if (this.parentNode) this.parentNode.removeChild(this);
  };

export * from "../rebrand-core.mjs";
