/**
 * A transparent wrapper authored *inside* a link or a simple field.
 *
 * `CT_Hyperlink` and `CT_SimpleField` are both `EG_PContent`, which declares
 * `w:bdo`, `w:dir`, `w:smartTag` and the run-level `w:customXml`, so
 * `w:hyperlink > w:bdo > w:r` is markup a producer may write. folio used to
 * keep the smart tag and the custom-XML wrapper as opaque bytes and capture
 * the two bidirectional wrappers whole, which kept the markup and cost the
 * runs inside it: linked text inside a `w:dir` was not text the editor held.
 *
 * Two legs, two different promises, and the pair is the point:
 *
 * - **Parse and save keep the authored nesting.** A document folio opens and
 *   writes back is the document it opened, wrapper inside link, with nothing
 *   captured.
 * - **The editor canonicalises.** A wrapper is a mark on the leaves it held
 *   and a link is a mark on the same leaves, and neither mark says which is
 *   inside which, so the save leg writes the one order the wrapper design
 *   fixed: revision, then wrapper, then hyperlink, then run. `w:bdo` moves
 *   from inside the link to around it. That is a documented canonicalisation
 *   and not a loss — the link, the wrapper and the text all survive, and a
 *   paragraph nobody edited keeps its authored order because selective save
 *   replays its bytes.
 *
 * A field is not canonicalised the same way, because the field node holds its
 * own inline content: a wrapper inside `w:fldSimple` comes back inside it.
 */

import { describe, expect, test } from "bun:test";
import fc from "fast-check";
import { EditorState, TextSelection } from "prosemirror-state";

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
 * The shape a comparison is about: the containers, what identifies each, and
 * the text. A run comes back carrying whatever formatting the paragraph
 * resolved onto it, so comparing the models whole would fail on facts neither
 * leg is being asked about.
 */
type Shape =
  | { of: "text"; text: string }
  | { of: "wrapper"; wrapper: string; content: Shape[] }
  | { of: "hyperlink"; anchor: string | undefined; content: Shape[] }
  | { of: "field"; instruction: string; content: Shape[] }
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
    case "hyperlink":
      return { of: "hyperlink", anchor: item.anchor, content: item.children.map(shapeOf) };
    case "simpleField":
      return { of: "field", instruction: item.instruction, content: item.content.map(shapeOf) };
    default:
      return { of: "other", type: item.type };
  }
};

const shapesOf = (paragraph: Paragraph): Shape[] => paragraph.content.map(shapeOf);

/** Whether anything in the paragraph came back as opaque bytes. */
const holdsCapture = (items: readonly ParagraphContent[]): boolean =>
  items.some((item) => {
    switch (item.type) {
      case "preservedInline":
        return true;
      case "inlineWrapper":
        return holdsCapture(item.content);
      case "hyperlink":
        return holdsCapture(item.children);
      case "simpleField":
        return holdsCapture(item.content);
      default:
        return false;
    }
  });

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

const layerArbitrary: fc.Arbitrary<Layer> = fc.oneof(
  fc
    .record({
      control: fc.constantFrom("override" as const, "embedding" as const),
      direction: fc.constantFrom("ltr" as const, "rtl" as const, undefined),
    })
    .map(({ control, direction }) => bidiLayer(control, direction)),
  fc
    .record({
      kind: fc.constantFrom("smartTag" as const, "customXml" as const),
      element: fc.constantFrom("City", "party"),
      uri: fc.constantFrom("urn:example:tags", undefined),
      properties: fc.constantFrom("v", undefined),
    })
    .map(({ kind, element, uri, properties }) => taggedLayer(kind, element, uri, properties)),
);

/** One or two wrappers, outermost first. */
const layersArbitrary: fc.Arbitrary<Layer[]> = fc.array(layerArbitrary, {
  minLength: 1,
  maxLength: 2,
});

const RUN = "<w:r><w:t>abc</w:t></w:r>";
const ANCHOR = "top";
const INSTRUCTION = "PAGE";

const nest = (layers: readonly Layer[], inner: string): string => {
  let xml = inner;
  for (const layer of layers.toReversed()) {
    xml = layer.xml(xml);
  }
  return xml;
};

/** `<w:hyperlink>` with the wrappers inside it, as an author may write them. */
const linkOverWrappers = (layers: readonly Layer[]): string =>
  `<w:hyperlink w:anchor="${ANCHOR}">${nest(layers, RUN)}</w:hyperlink>`;

/** `<w:fldSimple>` with the wrappers around its cached result. */
const fieldOverWrappers = (layers: readonly Layer[]): string =>
  `<w:fldSimple w:instr="${INSTRUCTION}">${nest(layers, RUN)}</w:fldSimple>`;

/** A link in a field whose cached result is itself inside transparent wrappers. */
const fieldOverLinkOverWrappers = (layers: readonly Layer[]): string =>
  `<w:fldSimple w:instr="${INSTRUCTION}">${linkOverWrappers(layers)}</w:fldSimple>`;

const textShape: Shape = { of: "text", text: "abc" };

const wrappersAround = (layers: readonly Layer[], inner: Shape): Shape => {
  let shape = inner;
  for (const layer of layers.toReversed()) {
    shape = { of: "wrapper", wrapper: layer.shape, content: [shape] };
  }
  return shape;
};

const linkShape = (content: Shape[]): Shape => ({
  of: "hyperlink",
  anchor: ANCHOR,
  content,
});

describe("a transparent wrapper inside a link, through parse and save", () => {
  test(
    "comes back as the nesting the source wrote, with nothing captured",
    () => {
      fc.assert(
        fc.property(layersArbitrary, (layers) => {
          const authored = parseParagraphXml(linkOverWrappers(layers));
          expect(holdsCapture(authored.content)).toBe(false);
          expect(shapesOf(authored)).toEqual([linkShape([wrappersAround(layers, textShape)])]);
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
        fc.property(layersArbitrary, (layers) => {
          const once = serializeParagraph(parseParagraphXml(linkOverWrappers(layers)));
          expect(serializeParagraph(reparse(once))).toBe(once);
        }),
        propertyConfig({ numRuns: 200 }),
      );
    },
    propertyTestTimeout(),
  );
});

describe("a transparent wrapper inside a simple field, through parse and save", () => {
  test(
    "comes back as the nesting the source wrote, with nothing captured",
    () => {
      fc.assert(
        fc.property(layersArbitrary, (layers) => {
          const authored = parseParagraphXml(fieldOverWrappers(layers));
          expect(holdsCapture(authored.content)).toBe(false);
          expect(shapesOf(authored)).toEqual([
            { of: "field", instruction: INSTRUCTION, content: [wrappersAround(layers, textShape)] },
          ]);
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
        fc.property(layersArbitrary, (layers) => {
          const once = serializeParagraph(parseParagraphXml(fieldOverWrappers(layers)));
          expect(serializeParagraph(reparse(once))).toBe(once);
        }),
        propertyConfig({ numRuns: 200 }),
      );
    },
    propertyTestTimeout(),
  );
});

describe("a transparent wrapper inside a link, through the editor", () => {
  test(
    "comes back with the wrapper around the link, the canonical order",
    () => {
      fc.assert(
        fc.property(layersArbitrary, (layers) => {
          const reopened = throughTheEditor(parseParagraphXml(linkOverWrappers(layers)));
          expect(shapesOf(reopened)).toEqual([wrappersAround(layers, linkShape([textShape]))]);
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
        fc.property(layersArbitrary, (layers) => {
          const once = throughTheEditor(parseParagraphXml(linkOverWrappers(layers)));
          expect(shapesOf(throughTheEditor(once))).toEqual(shapesOf(once));
        }),
        propertyConfig({ numRuns: 200 }),
      );
    },
    propertyTestTimeout(),
  );
});

describe("a transparent wrapper inside a simple field, through the editor", () => {
  test(
    "keeps the wrapper inside the field, which is where the field node holds it",
    () => {
      fc.assert(
        fc.property(layersArbitrary, (layers) => {
          const reopened = throughTheEditor(parseParagraphXml(fieldOverWrappers(layers)));
          expect(shapesOf(reopened)).toEqual([
            { of: "field", instruction: INSTRUCTION, content: [wrappersAround(layers, textShape)] },
          ]);
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
        fc.property(layersArbitrary, (layers) => {
          const once = throughTheEditor(parseParagraphXml(fieldOverWrappers(layers)));
          expect(shapesOf(throughTheEditor(once))).toEqual(shapesOf(once));
        }),
        propertyConfig({ numRuns: 200 }),
      );
    },
    propertyTestTimeout(),
  );
});

describe("a wrapped hyperlink inside a simple field, through the editor", () => {
  test(
    "keeps the field display text and structured result",
    () => {
      fc.assert(
        fc.property(layersArbitrary, (layers) => {
          const reopened = throughTheEditor(parseParagraphXml(fieldOverLinkOverWrappers(layers)));

          expect(shapesOf(reopened)).toEqual([
            {
              of: "field",
              instruction: INSTRUCTION,
              content: [wrappersAround(layers, linkShape([textShape]))],
            },
          ]);
        }),
        propertyConfig({ numRuns: 200 }),
      );
    },
    propertyTestTimeout(),
  );
});

/** The first text node holding `text`, as an editor range. */
const rangeOfText = (state: EditorState, text: string): { from: number; to: number } => {
  let range: { from: number; to: number } | undefined;
  state.doc.descendants((node, position) => {
    const at = node.isText ? (node.text?.indexOf(text) ?? -1) : -1;
    if (range === undefined && at >= 0) {
      range = { from: position + at, to: position + at + text.length };
    }
  });
  if (range === undefined) {
    throw new Error(`The editor holds no text ${text}`);
  }
  return range;
};

const editedThroughTheEditor = (paragraph: Paragraph): Paragraph => {
  const source = wrapParagraph(paragraph);
  const state = EditorState.create({ doc: toProseDoc(source) });
  const { from, to } = rangeOfText(state, "b");
  const edited = state.apply(
    state.tr.setSelection(TextSelection.create(state.doc, from, to)).insertText("EDIT", from, to),
  );
  const block = fromProseDoc(edited.doc, source).package.document.content.at(0);
  if (block?.type !== "paragraph") {
    throw new Error("The edit lost its paragraph");
  }
  return block;
};

describe("an edit inside a wrapper inside a link", () => {
  test(
    "keeps the link, the wrapper and the new text",
    () => {
      fc.assert(
        fc.property(layersArbitrary, (layers) => {
          const edited = editedThroughTheEditor(parseParagraphXml(linkOverWrappers(layers)));
          expect(shapesOf(edited)).toEqual([
            wrappersAround(layers, linkShape([{ of: "text", text: "aEDITc" }])),
          ]);
        }),
        propertyConfig({ numRuns: 100 }),
      );
    },
    propertyTestTimeout(),
  );
});
