/**
 * `elementToXml` stopped going through fast-xml-parser's builder, and the
 * bytes it produces are replayed verbatim into saved packages. A difference of
 * one character is a corrupted document, so the old serializer lives on here,
 * reconstructed from the builder and the options it was configured with, and
 * the two are compared over arbitrary trees.
 *
 * The generator covers what the old pair could disagree about: the five
 * escaped characters and the carriage return, in text and in attribute values;
 * empty and whitespace-only text; elements with no children, whose serialized
 * form self-closes, and elements whose children are all empty text, whose form
 * does too; attribute order; nesting deep enough that a per-node copy would
 * show up as a different answer rather than only as a slower one.
 *
 * The serializer is deliberately not byte-identical in three places, and only
 * those three: a tab or a line feed between attribute quotes is written as a
 * character reference (§3.3.3), an attribute the model dropped is absent
 * rather than the word "undefined", and a character §2.2 admits no spelling
 * for is dropped. The first is applied to the reference below so the property
 * stays an equality; the other two are pinned by their own tests, since this
 * generator produces neither.
 */

import { describe, expect, test } from "bun:test";
import { XMLBuilder } from "fast-xml-parser";
import fc from "fast-check";

import { elementToXml, type XmlElement } from "./xmlParser";

/** The builder, configured exactly as `xmlParser` configured it. */
const fxpBuilder = new XMLBuilder({
  preserveOrder: true,
  ignoreAttributes: false,
  attributeNamePrefix: "",
  textNodeName: "#text",
  suppressEmptyNode: true,
});

/** The converter `elementToXml` used to build the builder's input with. */
const elementToFxpNode = (el: XmlElement): Record<string, unknown> => {
  if (el.type === "text") {
    return { "#text": el.text ?? "" };
  }
  const name = el.name ?? "";
  const children = el.elements ? el.elements.map(elementToFxpNode) : [];
  const node: Record<string, unknown> = { [name]: children };
  if (el.attributes && Object.keys(el.attributes).length > 0) {
    node[":@"] = el.attributes;
  }
  return node;
};

/**
 * A tab or a line feed between attribute quotes is the one place the two are
 * meant to disagree: XML 1.0 §3.3.3 attribute-value normalisation flattens a
 * literal one to a space in every conformant reader, so `escapeXmlAttribute`
 * writes the character reference the builder wrote literally. Substituting a
 * private-use placeholder before the build and the reference after it keeps
 * the comparison exact, and keeps the difference in one readable place rather
 * than loosening the assertion. The generators below draw printable ASCII plus
 * the characters named here, so neither placeholder can arrive in the tree.
 */
const ATTRIBUTE_TAB_PLACEHOLDER = "";
const ATTRIBUTE_LINE_FEED_PLACEHOLDER = "";

const withPlaceholderAttributes = (element: XmlElement): XmlElement => {
  if (element.type === "text") {
    return element;
  }
  const next: XmlElement = { ...element };
  if (element.attributes) {
    next.attributes = Object.fromEntries(
      Object.entries(element.attributes).map(([name, value]) => [
        name,
        String(value)
          .replaceAll("\t", ATTRIBUTE_TAB_PLACEHOLDER)
          .replaceAll("\n", ATTRIBUTE_LINE_FEED_PLACEHOLDER),
      ]),
    );
  }
  if (element.elements) {
    next.elements = element.elements.map(withPlaceholderAttributes);
  }
  return next;
};

/** The previous implementation, in full, with that one difference applied. */
const previousElementToXml = (element: XmlElement): string =>
  (fxpBuilder.build([elementToFxpNode(withPlaceholderAttributes(element))]) as string)
    .replaceAll("\r", "&#13;")
    .replaceAll(ATTRIBUTE_TAB_PLACEHOLDER, "&#9;")
    .replaceAll(ATTRIBUTE_LINE_FEED_PLACEHOLDER, "&#10;");

/**
 * Names are drawn from the XML Name production rather than from arbitrary
 * strings: a name containing a delimiter does not round-trip through either
 * serializer, and one containing a carriage return cannot be parsed at all.
 */
const xmlName = fc
  .tuple(
    fc.constantFrom("w", "a", "wp", "mc", "v", "m", ""),
    fc.stringMatching(/^[A-Za-z_][A-Za-z0-9_.-]{0,12}$/u),
  )
  .map(([prefix, local]) => (prefix === "" ? local : `${prefix}:${local}`));

/** Content that exercises every escape, plus the values that self-close. */
const textContent = fc.oneof(
  { weight: 3, arbitrary: fc.string({ maxLength: 20 }) },
  { weight: 2, arbitrary: fc.constantFrom("", " ", "   ", "\r", "\n", "\t", "\r\n") },
  {
    weight: 3,
    arbitrary: fc
      .array(fc.constantFrom("&", "<", ">", '"', "'", "\r", "\n", "a", " "), { maxLength: 16 })
      .map((characters) => characters.join("")),
  },
);

const attributes = fc.dictionary(xmlName, textContent, { maxKeys: 4 });

const textNode = textContent.map((text): XmlElement => ({ type: "text", text }));

const elementNode = fc.letrec<{ node: XmlElement }>((tie) => ({
  node: fc
    .tuple(xmlName, attributes, fc.array(fc.oneof(textNode, tie("node")), { maxLength: 4 }))
    .map(([name, attrs, children]): XmlElement => {
      const element: XmlElement = { type: "element", name };
      if (Object.keys(attrs).length > 0) {
        element.attributes = attrs;
      }
      if (children.length > 0) {
        element.elements = children;
      }
      return element;
    }),
})).node;

describe("element serialization", () => {
  test("matches the builder it replaced, on arbitrary trees", () => {
    fc.assert(
      fc.property(elementNode, (element) => {
        expect(elementToXml(element)).toBe(previousElementToXml(element));
      }),
      { numRuns: 2000 },
    );
  });

  test("matches it on a text node at the root", () => {
    fc.assert(
      fc.property(textNode, (node) => {
        expect(elementToXml(node)).toBe(previousElementToXml(node));
      }),
      { numRuns: 300 },
    );
  });

  test("matches it on trees nested deeply enough to copy", () => {
    const nest = (depth: number): XmlElement =>
      depth === 0
        ? { type: "element", name: "w:t", elements: [{ type: "text", text: "leaf & <tail>" }] }
        : {
            type: "element",
            name: `w:level${String(depth)}`,
            attributes: { "w:val": `d"${String(depth)}` },
            elements: [nest(depth - 1), { type: "text", text: "\r" }],
          };

    // 100 levels is where the builder refused ("Maximum nested tags
    // exceeded"), and it is also `xmlResourceLimits.maxDepth`, so it is the
    // deepest tree a capture can be handed in the first place. The new
    // serializer carries no limit of its own because that one already bounds
    // it; the differential runs up to the boundary the old one accepted.
    for (const depth of [1, 8, 64, 98]) {
      const tree = nest(depth);
      expect(elementToXml(tree)).toBe(previousElementToXml(tree));
    }
  });

  test("an element whose children are all empty text self-closes", () => {
    const element: XmlElement = {
      type: "element",
      name: "w:t",
      attributes: { "xml:space": "preserve" },
      elements: [
        { type: "text", text: "" },
        { type: "text", text: "" },
      ],
    };
    expect(elementToXml(element)).toBe('<w:t xml:space="preserve"/>');
    expect(elementToXml(element)).toBe(previousElementToXml(element));
  });

  test("writes attribute whitespace as references and content whitespace as itself", () => {
    const element: XmlElement = {
      type: "element",
      name: "wp:docPr",
      attributes: { descr: "two\nlines\twide" },
      elements: [{ type: "text", text: "two\nlines\twide" }],
    };

    expect(elementToXml(element)).toBe(
      '<wp:docPr descr="two&#10;lines&#9;wide">two\nlines\twide</wp:docPr>',
    );
    // The builder wrote both literally, and §3.3.3 flattens the attribute's
    // to spaces on the way back in. That is the difference, and it is the
    // only one this tree exposes.
    expect(fxpBuilder.build([elementToFxpNode(element)])).toBe(
      '<wp:docPr descr="two\nlines\twide">two\nlines\twide</wp:docPr>',
    );
  });

  test("omits an attribute the model dropped", () => {
    const element = {
      type: "element",
      name: "w:pgSz",
      attributes: { "w:w": "11906", "w:h": undefined },
    } as unknown as XmlElement;

    expect(elementToXml(element)).toBe('<w:pgSz w:w="11906"/>');
  });

  test("drops a character XML admits no spelling for", () => {
    const element: XmlElement = {
      type: "element",
      name: "w:t",
      attributes: { "w:val": "a\u0000b" },
      elements: [{ type: "text", text: "c\u0007d" }],
    };

    expect(elementToXml(element)).toBe('<w:t w:val="ab">cd</w:t>');
  });
});

/**
 * The reason the serializer was rewritten is that a subtree's bytes used to be
 * copied into a parallel node for every ancestor on the way out, so nesting
 * multiplied the work. Depth is the variable that exposed it, and the output
 * is a function of the tree rather than of how it was produced, so identity
 * across depths is what a rewrite has to keep: the property above drives it
 * with random shapes, this one with the shape that made the copying expensive.
 */
describe("deeply nested serialization", () => {
  test("output is identical at every depth, and grows only with content", () => {
    const leaf: XmlElement = {
      type: "element",
      name: "w:t",
      elements: [{ type: "text", text: "cell & <content>" }],
    };
    const wrap = (depth: number): XmlElement =>
      depth === 0
        ? { type: "element", name: "w:tbl", elements: [leaf, leaf, leaf, leaf] }
        : { type: "element", name: `w:n${String(depth)}`, elements: [wrap(depth - 1)] };

    let previousLength = 0;
    for (const depth of [1, 2, 4, 8, 16, 32]) {
      const tree = wrap(depth);
      const xml = elementToXml(tree);
      expect(xml).toBe(previousElementToXml(tree));
      // Each level adds its own open and close tag and nothing else: the
      // content below it is written once, not once per level.
      if (previousLength > 0) {
        expect(xml.length - previousLength).toBeLessThan(depth * 20);
      }
      previousLength = xml.length;
    }
  });
});
