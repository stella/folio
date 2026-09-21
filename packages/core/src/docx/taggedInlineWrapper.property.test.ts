/**
 * `w:smartTag` and the run-level `w:customXml` survive as inline wrappers.
 *
 * Before they were wrapper kinds, folio spliced a smart tag's children into
 * the paragraph and kept no wrapper at all, and captured a run-level
 * `w:customXml` whole so everything inside it was opaque bytes. Both are the
 * same transparent shape as `w:bdo`/`w:dir`: a name around content that is
 * ordinary paragraph content.
 *
 * The property is over the nesting rather than over one example, because the
 * nesting is what the two legs disagree about when they drift. The save leg
 * rebuilds it from a stack key, so a field left out of the key folds two
 * wrappers into one, and a kind left out of a `switch` drops one.
 */

import { describe, expect, test } from "bun:test";
import fc from "fast-check";

import { propertyConfig, propertyTestTimeout } from "../../../../test/property-testing";

import { fromProseDoc } from "../prosemirror/conversion/fromProseDoc";
import { toProseDoc } from "../prosemirror/conversion/toProseDoc";
import type { Document, Paragraph, ParagraphContent } from "../types/document";
import { parseParagraph } from "./paragraphParser";
import { serializeParagraph } from "./serializer/paragraphSerializer";
import { parseXmlDocument, type XmlElement } from "./xmlParser";

const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

const parseParagraphXml = (body: string): Paragraph => {
  const root = parseXmlDocument(`<w:p xmlns:w="${W}" xmlns:r="${R}">${body}</w:p>`);
  if (!root) {
    throw new Error("Failed to parse the paragraph fixture");
  }
  return parseParagraph(root as XmlElement, null, null, null, null, null);
};

/** The serializer writes a bare `<w:p>`, so re-reading it needs the bindings back. */
const reparse = (saved: string): Paragraph =>
  parseParagraphXml(saved.slice(saved.indexOf(">") + 1, saved.lastIndexOf("</w:p>")));

const wrapParagraph = (paragraph: Paragraph): Document => ({
  package: { document: { content: [paragraph] } },
});

const throughTheEditor = (paragraph: Paragraph): Paragraph => {
  const source = wrapParagraph(paragraph);
  const block = fromProseDoc(toProseDoc(source), source).package.document.content.at(0);
  if (block?.type !== "paragraph") {
    throw new Error("The round trip lost its paragraph");
  }
  return block;
};

/**
 * The shape a comparison is about: the wrappers, their fields and the text.
 *
 * A run comes back carrying whatever formatting the paragraph resolved onto
 * it, and a revision comes back with a re-minted id, so comparing the models
 * whole would fail on facts neither leg is being asked about.
 */
type Shape =
  | { of: "text"; text: string }
  | { of: "wrapper"; wrapper: string; content: Shape[] }
  | { of: "revision"; type: string; content: Shape[] }
  | { of: "hyperlink"; anchor: string | undefined; content: Shape[] }
  | { of: "other"; type: string };

const shapeOf = (item: ParagraphContent): Shape => {
  switch (item.type) {
    case "run":
      return {
        of: "text",
        text: item.content.map((part) => (part.type === "text" ? part.text : "")).join(""),
      };
    case "inlineWrapper":
      return {
        of: "wrapper",
        wrapper:
          item.kind === "bidi"
            ? `bidi:${item.control}:${item.direction ?? ""}`
            : `${item.kind}:${item.element}:${item.uri ?? ""}:${item.propertiesXml ?? ""}`,
        content: item.content.map(shapeOf),
      };
    case "insertion":
    case "deletion":
    case "moveFrom":
    case "moveTo":
      return { of: "revision", type: item.type, content: item.content.map(shapeOf) };
    case "hyperlink":
      return { of: "hyperlink", anchor: item.anchor, content: item.children.map(shapeOf) };
    default:
      return { of: "other", type: item.type };
  }
};

const shapesOf = (paragraph: Paragraph): Shape[] => paragraph.content.map(shapeOf);

type Layer = { xml: (inner: string) => string; shape: string };

const bidiLayer = (control: "override" | "embedding", direction?: "ltr" | "rtl"): Layer => {
  const tag = control === "override" ? "bdo" : "dir";
  const value = direction === undefined ? "" : ` w:val="${direction}"`;
  return {
    xml: (inner) => `<w:${tag}${value}>${inner}</w:${tag}>`,
    shape: `bidi:${control}:${direction ?? ""}`,
  };
};

const taggedLayer = (
  kind: "smartTag" | "customXml",
  element: string,
  uri: string | undefined,
  properties: string | undefined,
): Layer => {
  const propertiesXml =
    properties === undefined
      ? undefined
      : `<w:${kind}Pr><w:attr w:name="k" w:val="${properties}"/></w:${kind}Pr>`;
  return {
    xml: (inner) =>
      `<w:${kind}${uri === undefined ? "" : ` w:uri="${uri}"`} w:element="${element}">` +
      `${propertiesXml ?? ""}${inner}</w:${kind}>`,
    shape: `${kind}:${element}:${uri ?? ""}:${propertiesXml ?? ""}`,
  };
};

const bidiArbitrary: fc.Arbitrary<Layer> = fc
  .record({
    control: fc.constantFrom("override" as const, "embedding" as const),
    direction: fc.constantFrom("ltr" as const, "rtl" as const, undefined),
  })
  .map(({ control, direction }) => bidiLayer(control, direction));

const taggedArbitrary: fc.Arbitrary<Layer> = fc
  .record({
    kind: fc.constantFrom("smartTag" as const, "customXml" as const),
    element: fc.constantFrom("City", "party"),
    uri: fc.constantFrom("urn:example:tags", undefined),
    properties: fc.constantFrom("v", undefined),
  })
  .map(({ kind, element, uri, properties }) => taggedLayer(kind, element, uri, properties));

/** One to three layers, at least one of which is the kind under test. */
const layersArbitrary: fc.Arbitrary<Layer[]> = fc
  .tuple(
    fc.array(fc.oneof(bidiArbitrary, taggedArbitrary), { maxLength: 2 }),
    taggedArbitrary,
    fc.nat(),
  )
  .map(([others, tagged, at]) => {
    const layers: Layer[] = [];
    layers.push(...others);
    layers.splice(at % (layers.length + 1), 0, tagged);
    return layers;
  });

const RUN = "<w:r><w:t>x</w:t></w:r>";
const HYPERLINK = `<w:hyperlink w:anchor="top">${RUN}</w:hyperlink>`;
const INSERTED = `<w:ins w:id="1" w:author="Reviewer" w:date="2026-01-01T00:00:00Z">${RUN}</w:ins>`;

const contentArbitrary = fc.constantFrom(RUN, HYPERLINK, INSERTED);

const nest = (layers: readonly Layer[], inner: string): string => {
  let xml = inner;
  for (const layer of layers.toReversed()) {
    xml = layer.xml(xml);
  }
  return xml;
};

/**
 * The nesting the save leg writes: the revision outermost.
 *
 * The editor carries a revision as a mark and a wrapper as a mark, and neither
 * says which is inside which, so the save leg fixes one order. This restates
 * the rule the wrapper design records rather than the value the code produced.
 */
const canonicalShape = (layers: readonly Layer[], inner: string): Shape => {
  let wrapped: Shape =
    inner === HYPERLINK
      ? { of: "hyperlink", anchor: "top", content: [{ of: "text", text: "x" }] }
      : { of: "text", text: "x" };
  for (const layer of layers.toReversed()) {
    wrapped = { of: "wrapper", wrapper: layer.shape, content: [wrapped] };
  }
  return inner === INSERTED ? { of: "revision", type: "insertion", content: [wrapped] } : wrapped;
};

describe("a smart tag or a run-level custom-XML wrapper through parse and save", () => {
  test(
    "comes back as the nesting the source wrote",
    () => {
      fc.assert(
        fc.property(layersArbitrary, contentArbitrary, (layers, inner) => {
          const authored = parseParagraphXml(nest(layers, inner));
          expect(shapesOf(reparse(serializeParagraph(authored)))).toEqual(shapesOf(authored));
        }),
        propertyConfig({ numRuns: 200 }),
      );
    },
    propertyTestTimeout(),
  );

  test(
    "is a fixed point of a second save",
    () => {
      fc.assert(
        fc.property(layersArbitrary, contentArbitrary, (layers, inner) => {
          const once = serializeParagraph(parseParagraphXml(nest(layers, inner)));
          expect(serializeParagraph(reparse(once))).toBe(once);
        }),
        propertyConfig({ numRuns: 200 }),
      );
    },
    propertyTestTimeout(),
  );

  test(
    "keeps every layer the source wrote",
    () => {
      fc.assert(
        fc.property(layersArbitrary, contentArbitrary, (layers, inner) => {
          const saved = serializeParagraph(parseParagraphXml(nest(layers, inner)));
          for (const layer of layers) {
            expect(saved).toContain(layer.xml("").slice(0, layer.xml("").indexOf(">") + 1));
          }
        }),
        propertyConfig({ numRuns: 200 }),
      );
    },
    propertyTestTimeout(),
  );
});

describe("a smart tag or a run-level custom-XML wrapper through the editor", () => {
  test(
    "comes back as the canonical nesting",
    () => {
      fc.assert(
        fc.property(layersArbitrary, contentArbitrary, (layers, inner) => {
          const reopened = throughTheEditor(parseParagraphXml(nest(layers, inner)));
          expect(shapesOf(reopened)).toEqual([canonicalShape(layers, inner)]);
        }),
        propertyConfig({ numRuns: 200 }),
      );
    },
    propertyTestTimeout(),
  );

  test(
    "is a fixed point of a second round trip",
    () => {
      fc.assert(
        fc.property(layersArbitrary, contentArbitrary, (layers, inner) => {
          const once = throughTheEditor(parseParagraphXml(nest(layers, inner)));
          expect(shapesOf(throughTheEditor(once))).toEqual(shapesOf(once));
        }),
        propertyConfig({ numRuns: 200 }),
      );
    },
    propertyTestTimeout(),
  );
});

describe("the properties a tagged wrapper carries", () => {
  for (const kind of ["smartTag", "customXml"] as const) {
    test(`w:${kind}Pr survives the editor verbatim`, () => {
      const properties =
        `<w:${kind}Pr><w:attr w:uri="urn:example:tags" w:name="year" w:val="2026"/>` +
        `<w:attr w:name="month" w:val="1"/></w:${kind}Pr>`;
      const authored = parseParagraphXml(
        `<w:${kind} w:uri="urn:example:tags" w:element="Date">${properties}${RUN}</w:${kind}>`,
      );
      const wrapper = authored.content.at(0);
      expect(wrapper?.type).toBe("inlineWrapper");
      if (wrapper?.type !== "inlineWrapper" || wrapper.kind === "bidi") {
        throw new Error(`Expected a ${kind} wrapper`);
      }
      expect(wrapper.propertiesXml).toBe(properties);
      expect(serializeParagraph(throughTheEditor(authored))).toContain(properties);
    });

    test(`two adjacent w:${kind} that differ only in their properties stay two`, () => {
      const one = `<w:${kind} w:element="Date"><w:${kind}Pr><w:attr w:name="k" w:val="1"/></w:${kind}Pr><w:r><w:t>a</w:t></w:r></w:${kind}>`;
      const two = `<w:${kind} w:element="Date"><w:${kind}Pr><w:attr w:name="k" w:val="2"/></w:${kind}Pr><w:r><w:t>b</w:t></w:r></w:${kind}>`;
      const reopened = throughTheEditor(parseParagraphXml(`${one}${two}`));
      expect(reopened.content).toHaveLength(2);
      expect(shapesOf(reopened)).toEqual(shapesOf(parseParagraphXml(`${one}${two}`)));
    });
  }

  test("properties that are not the element they claim to be are not replayed", () => {
    const authored = parseParagraphXml(`<w:smartTag w:element="Date">${RUN}</w:smartTag>`);
    const wrapper = authored.content.at(0);
    if (wrapper?.type !== "inlineWrapper" || wrapper.kind !== "smartTag") {
      throw new Error("Expected a smart tag wrapper");
    }
    // What a paste from outside the editor can put on the mark.
    wrapper.propertiesXml = "</w:smartTag><w:r><w:t>injected</w:t></w:r>";
    const saved = serializeParagraph(authored);
    expect(saved).not.toContain("injected");
    expect(saved).toBe(
      '<w:p><w:smartTag w:element="Date"><w:r><w:t>x</w:t></w:r></w:smartTag></w:p>',
    );
  });

  test.each([
    ["<!DOCTYPE x><w:smartTagPr/>", ""],
    ["<?producer hidden?><w:smartTagPr/>", "<w:smartTagPr/>"],
    ['<x:smartTagPr xmlns:x="urn:evil"/>', ""],
    ['<w:smartTagPr xmlns:w="urn:evil"/>', ""],
  ] as const)("unsafe smart-tag properties are sanitized: %s", (propertiesXml, sanitized) => {
    const authored = parseParagraphXml(`<w:smartTag w:element="Date">${RUN}</w:smartTag>`);
    const wrapper = authored.content.at(0);
    if (wrapper?.type !== "inlineWrapper" || wrapper.kind !== "smartTag") {
      throw new Error("Expected a smart tag wrapper");
    }
    wrapper.propertiesXml = propertiesXml;

    expect(serializeParagraph(authored)).toBe(
      `<w:p><w:smartTag w:element="Date">${sanitized}<w:r><w:t>x</w:t></w:r></w:smartTag></w:p>`,
    );
  });
});
