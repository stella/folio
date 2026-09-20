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
];

const roundTripThroughDom = (stack: readonly InlineWrapperLayer[]): Mark[] => {
  const window = new Window();
  const document = window.document as unknown as globalThis.Document;
  const fragment = DOMSerializer.fromSchema(schema).serializeFragment(
    schema.node("paragraph", null, [schema.text("x", [markType.create({ stack })])]).content,
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
    expect(parseInlineWrapperStack('[{"kind":"smartTag","element":"date"}]')).toBeNull();
    expect(parseInlineWrapperStack('[{"kind":"bidi"}]')).toBeNull();
    expect(parseInlineWrapperStack('[{"kind":"bidi","control":"sideways"}]')).toBeNull();
    expect(
      parseInlineWrapperStack('[{"kind":"bidi","control":"override","element":"date"}]'),
    ).toBeNull();
  });

  test("a stack spelled with its keys in another order is canonicalised", () => {
    expect(
      parseInlineWrapperStack('[{"direction":"rtl","control":"override","kind":"bidi"}]')?.map(
        (layer) => Object.keys(layer),
      ),
    ).toEqual([["kind", "control", "direction"]]);
  });

  test("the attribute name is what the mark writes", () => {
    expect(INLINE_WRAPPER_STACK_ATTRIBUTE).toBe("data-inline-wrapper");
  });
});
