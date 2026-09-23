/**
 * The preflight counts what the parser will build, or it bounds nothing.
 *
 * The bound is enforced by a lexical scan and paid by an object tree built
 * later by a different implementation. Nothing in either forces the two to
 * agree, and a scan that undercounts is a bound that does not hold: markup
 * counted as one element and parsed as a thousand passes the budget and
 * allocates past it. So the scan's count is compared against the tree's over
 * arbitrary parts, rather than being trusted to stay in step.
 *
 * The generator covers what the two could disagree about: self-closing and
 * paired elements, attributes whose values contain `=`, `<`, `>` and quotes,
 * text that looks like markup, comments, CDATA, processing instructions, and
 * nesting.
 */

import { describe, expect, setDefaultTimeout, test } from "bun:test";
import fc from "fast-check";

import { propertyTestTimeout } from "../../../../test/property-testing";

import { escapeXmlAttribute, escapeXmlText } from "@stll/docx-core";

import { parseXml, type XmlElement } from "./xmlParser";
import {
  assertXmlResourceLimits,
  exceedsUtf8ByteLimit,
  FOLIO_XML_RESOURCE_LIMITS,
} from "./xmlResourceLimits";

setDefaultTimeout(propertyTestTimeout(30_000));

const name = fc.constantFrom("a", "b", "w:p", "w:r", "w:t", "ns:x");
const attributeValue = fc.constantFrom("", "1", "x=y", "a<b", "a>b", "it's", 'say "hi"', "a=b=c");
const text = fc.constantFrom("", " ", "text", "a < b", "a > b", "5 = 5", "&amp;");

// The fixtures are escaped by the owner rather than by a local pair: a
// generator that escapes differently from the writer under test would compare
// the scan against markup folio never produces.
const escapeAttribute = escapeXmlAttribute;
const escapeText = escapeXmlText;

type Node =
  | { kind: "element"; name: string; attributes: [string, string][]; children: Node[] }
  | { kind: "text"; value: string }
  | { kind: "comment" }
  | { kind: "cdata" }
  | { kind: "instruction" };

const node: fc.Arbitrary<Node> = fc.letrec<{ node: Node }>((tie) => ({
  node: fc.oneof(
    // The leaves come first so shrinking has a base case to fall back to, and
    // the depth is capped explicitly: an unbounded recursion exhausts the
    // generator's own stack before it ever reaches the scan.
    { maxDepth: 4, depthIdentifier: "node", withCrossShrink: true },
    fc.record({ kind: fc.constant("text" as const), value: text }),
    fc.record({ kind: fc.constant("comment" as const) }),
    fc.record({ kind: fc.constant("cdata" as const) }),
    fc.record({ kind: fc.constant("instruction" as const) }),
    fc.record({
      kind: fc.constant("element" as const),
      name,
      attributes: fc.array(fc.tuple(name, attributeValue), { maxLength: 4 }),
      children: fc.array(tie("node"), { maxLength: 4 }),
    }),
  ),
})).node;

const render = (value: Node): string => {
  switch (value.kind) {
    case "text":
      return escapeText(value.value);
    case "comment":
      return "<!-- <w:r/> a = b -->";
    case "cdata":
      return "<![CDATA[<w:r/> a=b]]>";
    case "instruction":
      return `<?pi a="1"?>`;
    case "element": {
      // Duplicate attribute names are not well-formed XML and the two
      // implementations are entitled to differ on them, so the last wins here
      // exactly as it does in the tree.
      const attributes = [...new Map(value.attributes)]
        .map(([key, raw]) => ` ${key}="${escapeAttribute(raw)}"`)
        .join("");
      if (value.children.length === 0) {
        return `<${value.name}${attributes}/>`;
      }
      return `<${value.name}${attributes}>${value.children.map(render).join("")}</${value.name}>`;
    }
    default:
      throw new Error("unreachable");
  }
};

/**
 * Count the elements under the document node `parseXml` returns. That node is
 * synthetic: no tag produced it, so the scan never sees it and it is not part
 * of what the two implementations have to agree on.
 */
const treeCounts = (documentNode: XmlElement): { elements: number; attributes: number } => {
  let elements = 0;
  let attributes = 0;
  const pending = [...(documentNode.elements ?? [])];
  while (pending.length > 0) {
    const current = pending.pop();
    if (!current || current.type !== "element") {
      continue;
    }
    elements += 1;
    attributes += Object.keys(current.attributes ?? {}).length;
    for (const child of current.elements ?? []) {
      pending.push(child);
    }
  }
  return { elements, attributes };
};

describe("preflight counts and parsed tree", () => {
  test("agree on elements and attributes for a part under the bound", () => {
    fc.assert(
      fc.property(node, (generated) => {
        const xml = `<root>${render(generated)}</root>`;
        const scanned = assertXmlResourceLimits({ xml, limits: FOLIO_XML_RESOURCE_LIMITS });
        const counted = treeCounts(parseXml(xml));

        expect(scanned.elements).toBe(counted.elements);
        expect(scanned.attributes).toBe(counted.attributes);
      }),
      { numRuns: 500 },
    );
  });
});

// The byte bound skips its count for a string too short to reach the limit, so
// the shortcut is checked against the encoder at and around that threshold,
// including lone surrogates, which encode as a three-byte replacement.
describe("byte bound", () => {
  test("agrees with the UTF-8 encoder at every limit around the string's size", () => {
    const encoder = new TextEncoder();
    fc.assert(
      fc.property(
        fc.string({ unit: "binary", maxLength: 64 }),
        fc.integer({ min: -2, max: 2 }),
        fc.constantFrom(1, 2, 3, 4),
        (value, offset, perUnit) => {
          const maxBytes = Math.max(0, value.length * perUnit + offset);
          const bytes = encoder.encode(value).length;
          expect(exceedsUtf8ByteLimit(value, maxBytes)).toBe(bytes > maxBytes);
        },
      ),
      { numRuns: 2_000 },
    );
  });
});
