/**
 * The mark a projection builds and the mark a paste rebuilds from its own DOM
 * must be the same value.
 *
 * `fromProseDoc` groups adjacent inline leaves into one `w:r` by
 * `JSON.stringify` over their mark attrs. Object key order is insertion order,
 * so a layer written `{kind, control, direction}` and the same layer written
 * `{control, kind, direction}` are equal marks with different keys: two runs
 * where the author wrote one, and on the save leg two wrappers where the
 * author wrote one. The factory is what makes that impossible, and this holds
 * it to it across the DOM round trip.
 */

import { describe, expect, test } from "bun:test";
import { DOMParser, DOMSerializer, type Mark } from "prosemirror-model";
import { Window } from "happy-dom";

import {
  INLINE_WRAPPER_STACK_ATTRIBUTE,
  inlineWrapperLayer,
  parseInlineWrapperStack,
  serializeInlineWrapperStack,
} from "./inlineWrapperStack";
import { schema } from "./schema";
import type { InlineWrapperLayer } from "./schema/marks";

const markType = schema.marks["inlineWrapper"]!;

/** The key `fromProseDoc` groups runs by. */
const marksKey = (marks: readonly Mark[]): string =>
  marks
    .map((mark) => `${mark.type.name}:${JSON.stringify(mark.attrs)}`)
    .toSorted()
    .join("|");

const STACKS: readonly (readonly InlineWrapperLayer[])[] = [
  [{ kind: "bidi", control: "override", direction: "rtl" }],
  [{ kind: "bidi", control: "embedding" }],
  [
    { kind: "bidi", control: "embedding", direction: "ltr" },
    { kind: "bidi", control: "override", direction: "rtl" },
  ],
  [{ kind: "smartTag", element: "City" }],
  [
    {
      kind: "smartTag",
      element: "City",
      uri: "urn:schemas-microsoft-com:office:smarttags",
      propertiesXml: '<w:smartTagPr><w:attr w:name="kind" w:val="example"/></w:smartTagPr>',
    },
  ],
  [{ kind: "customXml", element: "party", uri: "urn:folio:test" }],
  [
    { kind: "bidi", control: "embedding", direction: "rtl" },
    { kind: "smartTag", element: "City" },
    { kind: "customXml", element: "party" },
  ],
];

const roundTripThroughDom = (
  stack: readonly InlineWrapperLayer[],
  origin?: { _docxHyperlinkIndex: number; _docxInsideHyperlinkStackStart: number },
): Mark[] => {
  const window = new Window();
  const document = window.document as unknown as globalThis.Document;
  const fragment = DOMSerializer.fromSchema(schema).serializeFragment(
    schema.node("paragraph", null, [
      schema.text("x", [markType.create({ stack, ...(origin ?? {}) })]),
    ]).content,
    { document },
  );
  const host = document.createElement("div");
  host.append(fragment);
  const parsed = DOMParser.fromSchema(schema).parse(host);
  const marks: Mark[] = [];
  parsed.descendants((node) => {
    if (node.isText) {
      marks.push(...node.marks);
    }
  });
  return marks.filter((mark) => mark.type.name === "inlineWrapper");
};

describe("an inline wrapper mark through the DOM", () => {
  for (const stack of STACKS) {
    test(`keeps its grouping key: ${serializeInlineWrapperStack(stack)}`, () => {
      const authored = markType.create({ stack: stack.map(inlineWrapperLayer) });
      const [reparsed] = roundTripThroughDom(stack);
      expect(reparsed).toBeDefined();
      expect(marksKey([reparsed!])).toBe(marksKey([authored]));
    });
  }

  test("a bidi override is spelled as <bdo dir>", () => {
    const window = new Window();
    const document = window.document as unknown as globalThis.Document;
    const host = document.createElement("div");
    host.append(
      DOMSerializer.fromSchema(schema).serializeFragment(
        schema.node("paragraph", null, [
          schema.text("x", [
            markType.create({ stack: [{ kind: "bidi", control: "override", direction: "rtl" }] }),
          ]),
        ]).content,
        { document },
      ),
    );
    const bdo = host.querySelector("bdo");
    expect(bdo).not.toBeNull();
    expect(bdo?.getAttribute("dir")).toBe("rtl");
  });

  test("keeps a wrapper's imported hyperlink boundary", () => {
    const stack = STACKS.at(-1);
    if (stack === undefined) {
      throw new Error("Expected a wrapper stack fixture");
    }
    const origin = { _docxHyperlinkIndex: 4, _docxInsideHyperlinkStackStart: 1 };
    const authored = markType.create({ stack, ...origin });
    const [reparsed] = roundTripThroughDom(stack, origin);
    expect(reparsed).toBeDefined();
    expect(marksKey([reparsed!])).toBe(marksKey([authored]));
  });

  test.each([
    ["smartTag", "data-smart-tag-element"],
    ["customXml", "data-custom-xml-element"],
  ] as const)("a %s is a span that names its element", (kind, attribute) => {
    const window = new Window();
    const document = window.document as unknown as globalThis.Document;
    const host = document.createElement("div");
    host.append(
      DOMSerializer.fromSchema(schema).serializeFragment(
        schema.node("paragraph", null, [
          schema.text("x", [markType.create({ stack: [{ kind, element: "City" }] })]),
        ]).content,
        { document },
      ),
    );
    expect(host.querySelector(`span[${attribute}]`)?.getAttribute(attribute)).toBe("City");
  });
});

describe("a stack read from markup the editor did not write", () => {
  test("paste from outside carries none", () => {
    expect(parseInlineWrapperStack(null)).toBeNull();
    expect(parseInlineWrapperStack(undefined)).toBeNull();
    expect(parseInlineWrapperStack("")).toBeNull();
  });

  test("markup that is not a stack carries none", () => {
    expect(parseInlineWrapperStack("not json")).toBeNull();
    expect(parseInlineWrapperStack("[]")).toBeNull();
    expect(parseInlineWrapperStack('[{"kind":"ruby","element":"date"}]')).toBeNull();
    expect(parseInlineWrapperStack('[{"kind":"bidi"}]')).toBeNull();
    expect(parseInlineWrapperStack('[{"kind":"bidi","control":"sideways"}]')).toBeNull();
    expect(
      parseInlineWrapperStack('[{"kind":"bidi","control":"override","element":"date"}]'),
    ).toBeNull();
  });

  test("a layer that states a field its kind does not carry is refused", () => {
    expect(parseInlineWrapperStack('[{"kind":"smartTag"}]')).toBeNull();
    expect(
      parseInlineWrapperStack('[{"kind":"smartTag","element":"City","control":"override"}]'),
    ).toBeNull();
    expect(parseInlineWrapperStack('[{"kind":"customXml","element":"party","uri":7}]')).toBeNull();
  });

  test("a stack spelled with its keys in another order is canonicalised", () => {
    expect(
      parseInlineWrapperStack('[{"direction":"rtl","control":"override","kind":"bidi"}]')?.map(
        (layer) => Object.keys(layer),
      ),
    ).toEqual([["kind", "control", "direction"]]);
    expect(
      parseInlineWrapperStack(
        '[{"propertiesXml":"<w:smartTagPr/>","uri":"urn:x","element":"City","kind":"smartTag"}]',
      )?.map((layer) => Object.keys(layer)),
    ).toEqual([["kind", "element", "uri", "propertiesXml"]]);
  });

  test("the attribute name is what the mark writes", () => {
    expect(INLINE_WRAPPER_STACK_ATTRIBUTE).toBe("data-inline-wrapper");
  });
});
