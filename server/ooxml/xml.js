// Small, exact-offset XML helpers for surgical edits of OOXML parts.
// We never re-serialize a whole part: callers locate elements by offset and splice text,
// so everything outside the edited elements stays byte-for-byte identical.

import { SaxesParser } from "saxes";

export class XmlError extends Error {}

/** Index of the '>' ending the tag that starts at `lt`, honouring quoted attribute values. */
function tagEnd(xml, lt) {
  let q = null;
  for (let i = lt + 1; i < xml.length; i++) {
    const ch = xml[i];
    if (q) {
      if (ch === q) q = null;
    } else if (ch === '"' || ch === "'") q = ch;
    else if (ch === ">") return i;
  }
  throw new XmlError("Unterminated tag.");
}

function skipSpecial(xml, lt) {
  if (xml.startsWith("<!--", lt)) return end(xml.indexOf("-->", lt + 4), 3);
  if (xml.startsWith("<![CDATA[", lt)) return end(xml.indexOf("]]>", lt + 9), 3);
  if (xml.startsWith("<?", lt)) return end(xml.indexOf("?>", lt + 2), 2);
  if (xml.startsWith("<!", lt)) return tagEnd(xml, lt) + 1;
  return -1;
  function end(i, n) {
    if (i < 0) throw new XmlError("Unterminated comment/CDATA/processing instruction.");
    return i + n;
  }
}

/**
 * Direct child elements of the region [from, to).
 * Each: { name, local, start, openEnd, selfClosing, closeStart, end }.
 */
export function scanChildren(xml, from, to) {
  const out = [];
  let i = from;
  let depth = 0;
  let cur = null;
  while (i < to) {
    const lt = xml.indexOf("<", i);
    if (lt < 0 || lt >= to) break;
    const skipped = skipSpecial(xml, lt);
    if (skipped >= 0) {
      i = skipped;
      continue;
    }
    const gt = tagEnd(xml, lt);
    const closing = xml[lt + 1] === "/";
    const self = !closing && xml[gt - 1] === "/";
    if (closing) {
      depth--;
      if (depth === 0 && cur) {
        cur.closeStart = lt;
        cur.end = gt + 1;
        out.push(cur);
        cur = null;
      }
      if (depth < 0) throw new XmlError("Unbalanced closing tag.");
    } else {
      if (depth === 0) {
        const m = /^<([^\s/>]+)/.exec(xml.slice(lt, Math.min(gt + 1, lt + 256)));
        const name = m[1];
        const el = { name, local: name.includes(":") ? name.slice(name.indexOf(":") + 1) : name, start: lt, openEnd: gt + 1 };
        if (self) {
          el.selfClosing = true;
          el.closeStart = gt + 1;
          el.end = gt + 1;
          out.push(el);
        } else {
          cur = el;
          depth = 1;
        }
      } else if (!self) depth++;
    }
    i = gt + 1;
  }
  if (depth !== 0) throw new XmlError("Unclosed element.");
  return out;
}

/** The document's root element. */
export function rootElement(xml) {
  const els = scanChildren(xml, 0, xml.length);
  if (els.length !== 1) throw new XmlError("Expected exactly one root element.");
  return els[0];
}

/** Ordered attributes of a start tag: [{ name, value (raw, still escaped), quote }]. */
export function parseAttrs(xml, el) {
  const tag = xml.slice(el.start, el.openEnd);
  const body = tag.replace(/^<[^\s/>]+/, "").replace(/\/?>$/, "");
  const attrs = [];
  const re = /([^\s=]+)\s*=\s*(["'])([\s\S]*?)\2/g;
  let m;
  while ((m = re.exec(body))) attrs.push({ name: m[1], value: m[3], quote: m[2] });
  return attrs;
}

export const getAttr = (attrs, name) => attrs.find((a) => a.name === name)?.value;

export function setAttr(attrs, name, value, after) {
  const i = attrs.findIndex((a) => a.name === name);
  if (i >= 0) attrs[i] = { ...attrs[i], value };
  else {
    const j = after ? attrs.findIndex((a) => a.name === after) : -1;
    attrs.splice(j >= 0 ? j + 1 : attrs.length, 0, { name, value, quote: '"' });
  }
  return attrs;
}

export const removeAttrs = (attrs, names) => attrs.filter((a) => !names.includes(a.name));

export function startTag(name, attrs, selfClosing = false) {
  const a = attrs.map((x) => ` ${x.name}=${x.quote || '"'}${x.value}${x.quote || '"'}`).join("");
  return `<${name}${a}${selfClosing ? "/>" : ">"}`;
}

/** Local name → prefixed name, using the same prefix as `likeName`. */
export const withPrefix = (likeName, local) => (likeName.includes(":") ? likeName.slice(0, likeName.indexOf(":") + 1) + local : local);

/** Prefix bound to namespace `ns` on an element's start tag ("" for default, null if absent). */
export function prefixFor(xml, el, ns) {
  for (const a of parseAttrs(xml, el)) {
    if (a.value !== ns) continue;
    if (a.name === "xmlns") return "";
    if (a.name.startsWith("xmlns:")) return a.name.slice(6);
  }
  return null;
}

// ---------- escaping ----------

export const escapeText = (s) => s.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");
export const escapeAttr = (s) =>
  s.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;").replace(/\n/g, "&#10;").replace(/\r/g, "&#13;").replace(/\t/g, "&#9;");

export function decodeEntities(s) {
  return s.replace(/&(#x[0-9a-fA-F]+|#\d+|amp|lt|gt|quot|apos);/g, (m, e) => {
    if (e === "amp") return "&";
    if (e === "lt") return "<";
    if (e === "gt") return ">";
    if (e === "quot") return '"';
    if (e === "apos") return "'";
    return String.fromCodePoint(e[1] === "x" ? parseInt(e.slice(2), 16) : parseInt(e.slice(1), 10));
  });
}

/**
 * OOXML ST_Xstring encoding: characters not allowed in XML become _xHHHH_, and literal
 * "_xHHHH_" sequences are protected as _x005F_xHHHH_ so Excel reads them back verbatim.
 */
export function encodeXstring(s) {
  if (/[\uD800-\uDBFF](?![\uDC00-\uDFFF])|(?<![\uD800-\uDBFF])[\uDC00-\uDFFF]/.test(s)) throw new XmlError("Text contains an unpaired surrogate.");
  return s
    .replace(/_x[0-9A-Fa-f]{4}_/g, (m) => "_x005F" + m)
    .replace(/[\x00-\x08\x0B\x0C\x0E-\x1F￾￿]/g, (c) => `_x${c.charCodeAt(0).toString(16).toUpperCase().padStart(4, "0")}_`);
}

// ---------- validation ----------

/** Throws if `xml` is not well-formed (strict parser, namespaces checked). */
export function assertWellFormed(xml, partName) {
  const parser = new SaxesParser({ xmlns: true, position: true });
  let error = null;
  parser.on("error", (e) => {
    if (!error) error = e;
  });
  parser.write(xml.charCodeAt(0) === 0xfeff ? xml.slice(1) : xml).close();
  if (error) throw new XmlError(`Generated XML for ${partName} is not well-formed: ${error.message}`);
}
