/**
 * What a transparent inline wrapper carries into the editor.
 *
 * The projection lifts `w:bdo`/`w:dir` out of the paragraph's content tree and
 * records what it lifted on the `inlineWrapper` mark of the leaves the wrapper
 * held. The property is the whole contract: for any nesting of wrappers the
 * parser can produce, the stack on a leaf is the authored nesting, outermost
 * first — so the save leg that rebuilds the wrappers has a faithful record and
 * not a lossy summary of one.
 *
 * Before the mark existed the leaves carried no stack at all, and every case
 * below failed on an empty list.
 */

import { describe, expect, test } from "bun:test";
import fc from "fast-check";
import type { Node as PMNode } from "prosemirror-model";

import { propertyConfig, propertyTestTimeout } from "../../../../../test/property-testing";

import type { Document, InlineWrapper, Paragraph, ParagraphContent } from "../../types/document";
import { createEmptyDocument } from "../../utils/createDocument";
import type { InlineWrapperLayer } from "../schema/marks";
import { toProseDoc } from "./toProseDoc";

const TEXT = "x";
const RUN = { type: "run", content: [{ type: "text", text: TEXT }] } as const;

const documentWith = (content: Paragraph["content"]): Document => {
  const template = createEmptyDocument();
  return {
    ...template,
    package: {
      ...template.package,
      document: {
        ...template.package.document,
        content: [{ type: "paragraph", paraId: "C0000001", content }],
      },
    },
  };
};

const leafOf = (content: Paragraph["content"]): PMNode => {
  let leaf: PMNode | undefined;
  toProseDoc(documentWith(content)).descendants((node) => {
    if (node.isText && node.text === TEXT) {
      leaf = node;
    }
  });
  if (!leaf) {
    throw new Error("The paragraph's own text did not reach the editor");
  }
  return leaf;
};

const markNamesOf = (leaf: PMNode): string[] => leaf.marks.map((mark) => mark.type.name);

const stackOf = (leaf: PMNode): unknown =>
  leaf.marks.find((mark) => mark.type.name === "inlineWrapper")?.attrs["stack"] ?? [];

/** An authored wrapper and the layer it must reach the editor as. */
type AuthoredWrapper = {
  wrapper: Omit<InlineWrapper, "content">;
  layer: InlineWrapperLayer;
};

const wrapperArbitrary: fc.Arbitrary<AuthoredWrapper> = fc
  .record({
    control: fc.constantFrom("override" as const, "embedding" as const),
    direction: fc.constantFrom("ltr" as const, "rtl" as const, undefined),
  })
  .map(({ control, direction }) =>
    direction === undefined
      ? {
          wrapper: { type: "inlineWrapper", kind: "bidi", control },
          layer: { kind: "bidi", control },
        }
      : {
          wrapper: { type: "inlineWrapper", kind: "bidi", control, direction },
          layer: { kind: "bidi", control, direction },
        },
  );

/** `inner` wrapped by `authored`, outermost first. */
const nest = (authored: readonly AuthoredWrapper[], inner: ParagraphContent): ParagraphContent => {
  let nested = inner;
  for (const { wrapper } of [...authored].reverse()) {
    nested = { ...wrapper, content: [nested] };
  }
  return nested;
};

describe("the inline wrapper mark a projected leaf carries", () => {
  test(
    "is the authored nesting, outermost first",
    () => {
      fc.assert(
        fc.property(fc.array(wrapperArbitrary, { minLength: 1, maxLength: 3 }), (authored) => {
          expect(stackOf(leafOf([nest(authored, RUN)]))).toEqual(
            authored.map(({ layer }) => layer),
          );
        }),
        propertyConfig({ numRuns: 200 }),
      );
    },
    propertyTestTimeout(),
  );

  test("is absent when no wrapper was authored", () => {
    expect(markNamesOf(leafOf([RUN]))).not.toContain("inlineWrapper");
  });

  test("survives inside a revision, beside the revision's own mark", () => {
    const leaf = leafOf([
      {
        type: "insertion",
        info: { id: 1, author: "Reviewer", date: "2026-01-01T00:00:00Z" },
        content: [
          {
            type: "inlineWrapper",
            kind: "bidi",
            control: "override",
            direction: "rtl",
            content: [RUN],
          },
        ],
      },
    ]);
    expect(markNamesOf(leaf)).toEqual(expect.arrayContaining(["insertion", "inlineWrapper"]));
    expect(stackOf(leaf)).toEqual([{ kind: "bidi", control: "override", direction: "rtl" }]);
  });

  test("survives inside a content control", () => {
    expect(
      stackOf(
        leafOf([
          {
            type: "inlineSdt",
            properties: { sdtType: "richText", tag: "bound" },
            content: [
              {
                type: "inlineWrapper",
                kind: "bidi",
                control: "embedding",
                direction: "ltr",
                content: [RUN],
              },
            ],
          },
        ]),
      ),
    ).toEqual([{ kind: "bidi", control: "embedding", direction: "ltr" }]);
  });
});
