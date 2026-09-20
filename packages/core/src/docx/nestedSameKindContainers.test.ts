/**
 * A container nested in one of its own kind comes back inside the outer one.
 *
 * `w:hyperlink` inside a `w:hyperlink`, `w:fldSimple` inside a `w:fldSimple`,
 * `w:r` inside the `w:rt` and `w:rubyBase` of a `w:ruby`, and `w:r` inside a
 * run-level `w:customXml` are all legal, and none of them is modelled: folio
 * keeps each as a capture in the owner's own content union, at the position it
 * was read at.
 *
 * The failure this pins is not "the markup is gone" but "the markup is
 * somebody else's": a parser that splices the inner container's runs into the
 * outer one writes a part that still contains every element name the source
 * had, with the inner container's content attributed to the outer container.
 * So every assertion here is about the inner element's *parent*, the way the
 * survival law's probe asks its question, and none of them is about the part
 * merely containing a name.
 */

import { describe, expect, test } from "bun:test";
import fc from "fast-check";

import { propertyConfig } from "../../../../test/property-testing";
import { fromProseDoc } from "../prosemirror/conversion/fromProseDoc";
import { toProseDoc } from "../prosemirror/conversion/toProseDoc";
import type { Document, Paragraph } from "../types/document";

import { parseParagraph } from "./paragraphParser";
import { serializeParagraph } from "./serializer/paragraphSerializer";
import { parseXmlDocument, type XmlElement } from "./xmlParser";

const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

const RUBY =
  "<w:ruby><w:rubyPr/>" +
  "<w:rt><w:r><w:t>reading</w:t></w:r></w:rt>" +
  "<w:rubyBase><w:r><w:t>word</w:t></w:r></w:rubyBase>" +
  "</w:ruby>";

/**
 * One nesting: the paragraph that carries it, and the (parent, child) pair the
 * save owes back. `expected` counts the occurrences under that parent, so a
 * reader that keeps one of the two runs a `w:ruby` holds is a failure rather
 * than a pass.
 */
const NESTINGS = [
  {
    name: "w:hyperlink in a w:hyperlink",
    body:
      '<w:hyperlink w:anchor="outer"><w:hyperlink w:anchor="inner">' +
      "<w:r><w:t>linked</w:t></w:r></w:hyperlink></w:hyperlink>",
    parent: "w:hyperlink",
    child: "w:hyperlink",
    expected: 1,
    text: "linked",
  },
  {
    name: "w:fldSimple in a w:fldSimple",
    body:
      '<w:fldSimple w:instr="IF"><w:fldSimple w:instr="PAGE">' +
      "<w:r><w:t>7</w:t></w:r></w:fldSimple></w:fldSimple>",
    parent: "w:fldSimple",
    child: "w:fldSimple",
    expected: 1,
    text: "7",
  },
  {
    name: "w:r in a w:rt",
    body: `<w:r>${RUBY}</w:r>`,
    parent: "w:rt",
    child: "w:r",
    expected: 1,
    // The reading printed above the base is not the sentence's text.
    text: "word",
  },
  {
    name: "w:r in a w:rubyBase",
    body: `<w:r>${RUBY}</w:r>`,
    parent: "w:rubyBase",
    child: "w:r",
    expected: 1,
    text: "word",
  },
  {
    name: "w:r in a run-level w:customXml",
    body: '<w:customXml w:element="party"><w:r><w:t>Acme</w:t></w:r></w:customXml>',
    parent: "w:customXml",
    child: "w:r",
    expected: 1,
    text: "Acme",
  },
] as const;

const TAG = /<(\/?)([^\s/>!?]+)((?:"[^"]*"|[^>"])*)>/gu;
const QUOTED = /"[^"]*"/gu;

const closesItself = (attributes: string): boolean =>
  attributes.replaceAll(QUOTED, "").trimEnd().endsWith("/");

/** Occurrences of `child` whose own parent element is `parent`. */
const childrenOf = (xml: string, parent: string, child: string): number => {
  const stack: string[] = [];
  let found = 0;
  for (const [, closing, name, attributes = ""] of xml.matchAll(TAG)) {
    if (name === undefined) {
      continue;
    }
    if (closing === "/") {
      stack.pop();
      continue;
    }
    const empty = closesItself(attributes);
    stack.push(name);
    if (name === child && stack.at(-2) === parent) {
      found += 1;
    }
    if (empty) {
      stack.pop();
    }
  }
  return found;
};

const parseParagraphXml = (xml: string): Paragraph => {
  const root = parseXmlDocument(xml) as XmlElement | null;
  if (!root) {
    throw new Error("the paragraph fixture did not parse");
  }
  return parseParagraph(root, null, null, null, null, null);
};

const paragraphXml = (body: string): string => `<w:p xmlns:w="${W}">${body}</w:p>`;

/** Every `text` the paragraph's captures carry, concatenated. */
const capturedText = (value: unknown): string => {
  if (Array.isArray(value)) {
    return value.map((item) => capturedText(item)).join("");
  }
  if (typeof value !== "object" || value === null) {
    return "";
  }
  const record = value as Record<string, unknown>;
  const type = record["type"];
  if (
    (type === "preservedInline" || type === "preservedXml") &&
    typeof record["text"] === "string"
  ) {
    return record["text"];
  }
  return Object.values(record)
    .map((item) => capturedText(item))
    .join("");
};

/** The paragraph as it comes back from the editor, with nothing edited. */
const throughTheEditor = (paragraph: Paragraph): Paragraph => {
  const input: Document = { package: { document: { content: [paragraph] } } };
  const first = fromProseDoc(toProseDoc(input), input).package.document.content.at(0);
  if (first?.type !== "paragraph") {
    throw new Error("the round trip did not give a paragraph back");
  }
  return first;
};

describe("a container nested in one of its own kind", () => {
  test("comes back under the container it was written into", () => {
    fc.assert(
      fc.property(fc.constantFrom(...NESTINGS), ({ body, parent, child, expected }) => {
        const saved = serializeParagraph(parseParagraphXml(paragraphXml(body)));
        expect(childrenOf(saved, parent, child)).toBe(expected);
      }),
      propertyConfig({ numRuns: 50 }),
    );
  });

  test("is a fixed point: the second save writes what the first did", () => {
    fc.assert(
      fc.property(fc.constantFrom(...NESTINGS), ({ body }) => {
        const first = serializeParagraph(parseParagraphXml(paragraphXml(body)));
        const second = serializeParagraph(
          parseParagraphXml(`<w:p xmlns:w="${W}">${first.slice(first.indexOf(">") + 1)}`),
        );
        expect(second).toBe(first);
      }),
      propertyConfig({ numRuns: 50 }),
    );
  });

  // A capture keeps the markup; the text beside it is what keeps the words on
  // the line. Every one of these wrappers is transparent, so the runs it holds
  // print like any other and a capture that carried none would take a linked
  // clause or a field's result off the page.
  test("carries the words the wrapper shows", () => {
    fc.assert(
      fc.property(fc.constantFrom(...NESTINGS), ({ body, text }) => {
        expect(capturedText(parseParagraphXml(paragraphXml(body)))).toContain(text);
      }),
      propertyConfig({ numRuns: 50 }),
    );
  });

  test("survives the editor under the same container", () => {
    fc.assert(
      fc.property(fc.constantFrom(...NESTINGS), ({ body, parent, child, expected }) => {
        const reopened = throughTheEditor(parseParagraphXml(paragraphXml(body)));
        expect(childrenOf(serializeParagraph(reopened), parent, child)).toBe(expected);
      }),
      propertyConfig({ numRuns: 50 }),
    );
  });
});
