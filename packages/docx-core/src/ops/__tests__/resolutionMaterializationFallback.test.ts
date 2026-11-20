import { expect, test } from "bun:test";

import type { Document, Paragraph } from "../../model/document";
import { storyParagraphs } from "../blocks";
import { tryMaterializeResolution } from "../resolutionMaterialization";
import { DOCUMENT_OP_TYPES, OP_STORIES } from "../types";

const paragraph = {
  type: "paragraph",
  paraId: "10000001",
  content: [{ type: "run", content: [{ type: "text", text: "text" }] }],
} satisfies Paragraph;

const makeDocument = (content: Paragraph[]): Document => ({
  package: { document: { content } },
});

const materialize = (
  document: Document,
  overrides: Partial<Parameters<typeof tryMaterializeResolution>[0]> = {},
) =>
  tryMaterializeResolution({
    document,
    story: OP_STORIES.MAIN,
    locations: storyParagraphs(document.package.document),
    inline: [],
    join: { type: "none" },
    noteReferences: [],
    receiverEffects: [],
    ...overrides,
  });

test("non-main stories and reference effects use native sequential checks", () => {
  const document = makeDocument([paragraph]);
  expect(
    materialize(document, { story: { kind: "header", rId: "rId1" } }),
  ).toEqual({ type: "fallback", reason: "story" });
  expect(
    materialize(document, {
      receiverEffects: [{ type: "paragraph-join", receiver: paragraph }],
    }),
  ).toEqual({ type: "fallback", reason: "referenceEffects" });
});

test("a cold trailing section must describe the actual body blocks", () => {
  const content = [paragraph];
  const properties = {};
  const document = {
    package: {
      document: {
        content,
        finalSectionProperties: properties,
        sections: [
          {
            properties,
            content: [{ ...paragraph, content: [] }],
          },
        ],
      },
    },
  } satisfies Document;
  expect(materialize(document)).toEqual({
    type: "fallback",
    reason: "sectionView",
  });
});

test("an unproven physical list or a duplicate ID stays on the general path", () => {
  const document = makeDocument([paragraph]);
  expect(materialize(document, { locations: [] })).toEqual({
    type: "fallback",
    reason: "paragraphList",
  });
  expect(materialize(makeDocument([paragraph, { ...paragraph }]))).toEqual({
    type: "fallback",
    reason: "inputContract",
  });
});

test.each([
  [16_384, "materialized"],
  [16_385, "fallback"],
] as const)(
  "inline resolution bounds a native splice of %i blocks",
  (count, outcome) => {
    const content = Array.from({ length: count }, (_, index) => ({
      ...paragraph,
      paraId: (0x10000000 + index).toString(16),
    }));
    const document = makeDocument(content);
    const first = content.at(0);
    const last = content.at(-1);
    if (first === undefined || last === undefined)
      throw new Error("Expected two changed paragraph endpoints");
    const result = materialize(document, {
      inline: [first, last].map(({ paraId, content: expected }) => ({
        type: DOCUMENT_OP_TYPES.REPLACE_INLINE,
        story: OP_STORIES.MAIN,
        blockId: paraId,
        expected,
        content: [{ type: "run", content: [{ type: "text", text: "edited" }] }],
      })),
    });
    expect(result.type).toBe(outcome);
    if (result.type === "fallback") expect(result.reason).toBe("wideSpan");
    else {
      expect(result.edit.document.package.document.content).toHaveLength(count);
      expect(result.edit.touched.modified).toEqual([first.paraId, last.paraId]);
    }
  },
);
