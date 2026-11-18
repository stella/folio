import { panic } from "better-result";
import { expect, test } from "bun:test";

import type {
  Document,
  Paragraph,
  ParagraphFormatting,
} from "../../model/document";
import { applyDocumentOp, applyDocumentOps } from "../apply";
import { normalizeForOps } from "../contract";
import {
  DOCUMENT_OP_TYPES,
  type DocumentOp,
  OP_STORIES,
  SPLIT_HALVES,
  type SplitParagraphFields,
} from "../types";

const sourceXml =
  '<w:pPr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"/>';
const fields = [
  { name: "empty", given: { formatting: {} }, expected: undefined },
  {
    name: "source-only",
    given: { formatting: { sourceXml } },
    expected: undefined,
  },
  {
    name: "modeled with source",
    given: { formatting: { alignment: "center", sourceXml } },
    expected: { alignment: "center", sourceXml },
  },
  {
    name: "explicit mark properties",
    given: { formatting: { runProperties: {} } },
    expected: { runProperties: {} },
  },
  {
    name: "other own fields",
    given: { textId: "10000002", renderedPageBreakBefore: true },
    expected: undefined,
  },
] as const satisfies readonly {
  name: string;
  given: SplitParagraphFields;
  expected: ParagraphFormatting | undefined;
}[];

/** Every authored state remains exact through wire undo, redo and later property edits. */
const apply = (document: Document, operation: DocumentOp) => {
  const original = structuredClone(document);
  const input = structuredClone(operation);
  const applied = applyDocumentOp(document, operation).unwrap();
  expect(document).toStrictEqual(original);
  expect(operation).toStrictEqual(input);
  const undone = applyDocumentOps(
    applied.document,
    JSON.parse(JSON.stringify(applied.inverse)),
  ).unwrap();
  expect(undone.document).toStrictEqual(document);
  expect(
    applyDocumentOps(
      undone.document,
      JSON.parse(JSON.stringify(undone.inverse)),
    ).unwrap().document,
  ).toStrictEqual(applied.document);
  return applied.document;
};

for (const half of [SPLIT_HALVES.FIRST, SPLIT_HALVES.SECOND])
  for (const sourceFormatting of [
    undefined,
    { alignment: "left", sourceXml },
  ] as const)
    test(`split ${half} finalizes new formatting and preserves ${sourceFormatting === undefined ? "absent" : "authored"} source properties`, () => {
      const original = {
        type: "paragraph",
        paraId: "00000001",
        textId: "10000001",
        preservedAttributes: [
          {
            name: "rsidR",
            namespace:
              "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
            value: "00112233",
          },
        ],
        ...(sourceFormatting === undefined
          ? {}
          : { formatting: sourceFormatting }),
        content: [{ type: "run", content: [{ type: "text", text: "abc" }] }],
      } satisfies Paragraph;
      const source = normalizeForOps({
        package: { document: { content: [original] } },
      } satisfies Document);
      for (const { given, expected } of fields) {
        const split = apply(source, {
          type: DOCUMENT_OP_TYPES.SPLIT_BLOCK,
          at: { story: OP_STORIES.MAIN, blockId: "00000001", offset: 1 },
          newBlockId: "00000002",
          newHalf: half,
          newParagraph: given,
        });
        const made = split.package.document.content.find(
          (block) => block.type === "paragraph" && block.paraId === "00000002",
        );
        const kept = split.package.document.content.find(
          (block) =>
            block.type === "paragraph" && block.paraId === original.paraId,
        );
        expect(made?.type).toBe("paragraph");
        expect(kept?.type).toBe("paragraph");
        if (made?.type !== "paragraph" || kept?.type !== "paragraph")
          panic("A split must produce both paragraphs");
        for (const [key, value] of Object.entries(given)) {
          if (key === "formatting") continue;
          expect(Reflect.get(made, key)).toStrictEqual(value);
        }
        expect(kept.formatting).toStrictEqual(sourceFormatting);
        expect(kept.textId).toBe(original.textId);
        expect(kept.preservedAttributes).toStrictEqual(
          original.preservedAttributes,
        );
        apply(split, {
          type: DOCUMENT_OP_TYPES.SET_PARAGRAPH_PROPS,
          story: OP_STORIES.MAIN,
          blockId: "00000002",
          patch: { alignment: "right" },
        });
        expect(made.formatting).toStrictEqual(expected);
        if (expected === undefined)
          expect(made).not.toHaveProperty("formatting");
      }
      const inherited = apply(source, {
        type: DOCUMENT_OP_TYPES.SPLIT_BLOCK,
        at: { story: OP_STORIES.MAIN, blockId: "00000001", offset: 1 },
        newBlockId: "00000002",
        newHalf: half,
      });
      for (const paragraph of inherited.package.document.content) {
        expect(paragraph.type).toBe("paragraph");
        if (paragraph.type !== "paragraph")
          panic("A split must produce paragraphs");
        expect(paragraph.formatting).toStrictEqual(sourceFormatting);
      }
    });
