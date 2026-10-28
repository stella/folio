/**
 * `elementToXml` writes the bytes replayed verbatim into saved packages, so a
 * difference of one character is a corrupted document.
 *
 * Two references hold it in place. Over every element of the repository's
 * fixtures, its output is pinned to a digest of what the general-purpose
 * builder folio serialized with before produced for the same trees (the one
 * deliberate difference below applied). Over arbitrary trees, reading the
 * output back with folio's well-formed reader returns the tree it was written
 * from: the five escaped characters and the carriage return, in text and in
 * attribute values; empty and whitespace-only text; elements with no
 * children, whose serialized form self-closes; attribute order; nesting.
 *
 * The serializer differs from that builder in three places, and only those
 * three: a tab or a line feed between attribute quotes is written as a
 * character reference (§3.3.3), an attribute the model dropped is absent
 * rather than the word "undefined", and a character §2.2 admits no spelling
 * for is dropped. Each is pinned by its own test below.
 */

import { describe, expect, setDefaultTimeout, test } from "bun:test";
import fc from "fast-check";
import { readFileSync } from "node:fs";
import { resolve } from "node:path";

import { readXml } from "@stll/docx-core/xml";

import {
  propertyConfig,
  propertyTestTimeout,
} from "../../../../test/property-testing";
import {
  answerDigest,
  readFixturePackages,
} from "../../../../tests/support/xmlFixtureParts";
import { elementToXml, parseXml, type XmlElement } from "./xmlParser";

setDefaultTimeout(propertyTestTimeout(30_000));

type SerializerGoldens = {
  elementToXml: { fixtures: Record<string, string> };
};

const GOLDENS: SerializerGoldens = JSON.parse(
  readFileSync(
    resolve(import.meta.dir, "__fixtures__/xml-goldens.json"),
    "utf8",
  ),
);

/**
 * The tree a reader returns for `element`: adjacent text merged, empty text
 * gone, and an element left with no children written without the array.
 */
const asRead = (element: XmlElement): XmlElement => {
  if (element.type === "text") {
    return { type: "text", text: String(element.text ?? "") };
  }
  const read: XmlElement = { type: "element", name: element.name ?? "" };
  if (element.attributes && Object.keys(element.attributes).length > 0) {
    read.attributes = Object.fromEntries(
      Object.entries(element.attributes).map(([name, value]) => [
        name,
        String(value),
      ]),
    );
  }
  const children: XmlElement[] = [];
  for (const child of element.elements ?? []) {
    const next = asRead(child);
    const previous = children.at(-1);
    if (next.type === "text") {
      if (next.text === "") continue;
      if (previous?.type === "text") {
        previous.text = `${String(previous.text)}${String(next.text)}`;
        continue;
      }
    }
    children.push(next);
  }
  if (children.length > 0) {
    read.elements = children;
  }
  return read;
};

const readBack = (xml: string): XmlElement | undefined => {
  const read = readXml(xml, { conformance: "well-formed" });
  expect(read.status).toBe("parsed");
  return read.status === "parsed" ? read.value.elements?.at(0) : undefined;
};

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
  {
    weight: 2,
    arbitrary: fc.constantFrom("", " ", "   ", "\r", "\n", "\t", "\r\n"),
  },
  {
    weight: 3,
    arbitrary: fc
      .array(fc.constantFrom("&", "<", ">", '"', "'", "\r", "\n", "a", " "), {
        maxLength: 16,
      })
      .map((characters) => characters.join("")),
  },
);

const attributes = fc.dictionary(xmlName, textContent, { maxKeys: 4 });

const textNode = textContent.map((text): XmlElement => ({
  type: "text",
  text,
}));

const elementNode = fc.letrec<{ node: XmlElement }>((tie) => ({
  node: fc
    .tuple(
      xmlName,
      attributes,
      fc.array(fc.oneof(textNode, tie("node")), { maxLength: 4 }),
    )
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
  test("writes every fixture element as the previous builder did", async () => {
    const packages = await readFixturePackages();
    expect(Object.keys(GOLDENS.elementToXml.fixtures)).toEqual(
      packages.map(({ path }) => path),
    );
    for (const { path, parts } of packages) {
      const answers: [string, string][] = [];
      for (const [part, xml] of parts) {
        for (const root of parseXml(xml).elements ?? []) {
          if (root.type === "element") answers.push([part, elementToXml(root)]);
        }
      }
      expect(answerDigest(answers), path).toBe(
        GOLDENS.elementToXml.fixtures[path],
      );
    }
  });

  test("reads back as the tree it was written from, on arbitrary trees", () => {
    fc.assert(
      fc.property(elementNode, (element) => {
        expect(readBack(elementToXml(element))).toEqual(asRead(element));
      }),
      propertyConfig({ numRuns: 2000 }),
    );
  });

  test("reads back on trees nested up to the depth limit", () => {
    const nest = (depth: number): XmlElement =>
      depth === 0
        ? {
            type: "element",
            name: "w:t",
            elements: [{ type: "text", text: "leaf & <tail>" }],
          }
        : {
            type: "element",
            name: `w:level${String(depth)}`,
            attributes: { "w:val": `d"${String(depth)}` },
            elements: [nest(depth - 1), { type: "text", text: "\r" }],
          };

    // `xmlResourceLimits.maxDepth` bounds every tree a capture can be handed,
    // so the serializer carries no limit of its own.
    for (const depth of [1, 8, 64, 98]) {
      const tree = nest(depth);
      expect(readBack(elementToXml(tree))).toEqual(asRead(tree));
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
    // Written literally, §3.3.3 would flatten the attribute's to spaces on
    // the way back in; the references survive it.
    expect(readBack(elementToXml(element))).toEqual(asRead(element));
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
        : {
            type: "element",
            name: `w:n${String(depth)}`,
            elements: [wrap(depth - 1)],
          };

    let previousLength = 0;
    for (const depth of [1, 2, 4, 8, 16, 32]) {
      const tree = wrap(depth);
      const xml = elementToXml(tree);
      expect(readBack(xml)).toEqual(asRead(tree));
      // Each level adds its own open and close tag and nothing else: the
      // content below it is written once, not once per level.
      if (previousLength > 0) {
        expect(xml.length - previousLength).toBeLessThan(depth * 20);
      }
      previousLength = xml.length;
    }
  });
});
