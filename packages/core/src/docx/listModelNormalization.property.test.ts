/**
 * List normalization keeps the list renderings the parser gave a package.
 *
 * The parser counts a story's lists through every paragraph a block content
 * control or custom XML wrapper holds, while a table's cells count their own.
 * Normalization recomputes the renderings of a story's top-level paragraphs
 * from counters of its own, so on a package nothing has edited it must arrive
 * at the parser's renderings: the first time, and again when a cache reuses
 * an earlier run's counters past a changed block.
 */

import { describe, expect, test } from "bun:test";
import fc from "fast-check";

import {
  applyDocumentOps,
  DOCUMENT_OP_TYPES,
  OP_STORIES,
  reviewFieldsOf,
} from "@stll/docx-core/ops";

import {
  assertProperty,
  propertyTestTimeout,
} from "../../../../test/property-testing";
import type {
  BlockContent,
  Document,
  ListRendering,
  Paragraph,
} from "../types/document";
import { prepareCanonicalDocxInput } from "./canonicalSessionInput";
import {
  ListNormalizationCache,
  normalizeListModelRendering,
} from "./listModelNormalization";
import { parseDocx } from "./parser";
import { createDocx } from "./rezip";

type Item = {
  kind: "reference" | "style" | "plain";
  numId: number;
  ilvl: number;
};

type GeneratedBlock =
  | { type: "item"; item: Item }
  | { type: "control" | "customXml"; blocks: GeneratedBlock[] }
  | { type: "table"; items: Item[] };

const itemArbitrary: fc.Arbitrary<Item> = fc.record({
  kind: fc.constantFrom("reference", "reference", "style", "plain"),
  numId: fc.integer({ min: 1, max: 3 }),
  ilvl: fc.nat({ max: 2 }),
});

const { block: blockArbitrary } = fc.letrec<{ block: GeneratedBlock }>(
  (tie) => ({
    block: fc.oneof(
      { depthSize: "small", withCrossShrink: true },
      {
        weight: 4,
        arbitrary: itemArbitrary.map((item): GeneratedBlock => ({
          type: "item",
          item,
        })),
      },
      fc.record({
        type: fc.constantFrom("control" as const, "customXml" as const),
        blocks: fc.array(tie("block"), { minLength: 1, maxLength: 3 }),
      }),
      fc.record({
        type: fc.constant("table" as const),
        items: fc.array(itemArbitrary, { minLength: 1, maxLength: 3 }),
      }),
    ),
  }),
);

const level = (ilvl: number) => ({
  ilvl,
  start: 1,
  numFmt: "decimal" as const,
  lvlText: `%${String(ilvl + 1)}.`,
});

const paragraphOf = ({ kind, numId, ilvl }: Item, text: string): Paragraph => ({
  type: "paragraph",
  ...(kind === "reference" && {
    formatting: { numPr: { kind: "reference", numId, ilvl } },
  }),
  ...(kind === "style" && { formatting: { styleId: "ListStyle" } }),
  content: [{ type: "run", content: [{ type: "text", text }] }],
});

const W_NAMESPACE =
  "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

/**
 * A story of numbered items, some inside content controls and custom XML
 * wrappers nested in each other or in a table, over three instances of one
 * abstract list (one restarting it) and a list style.
 */
const generatedDocument = (blocks: readonly GeneratedBlock[]): Document => {
  let next = 0;
  const text = () => `Item ${String((next += 1))}`;
  const content = (generated: readonly GeneratedBlock[]): BlockContent[] =>
    generated.map((block): BlockContent => {
      switch (block.type) {
        case "item":
          return paragraphOf(block.item, text());
        case "control":
          return {
            type: "blockSdt",
            properties: { sdtType: "richText", id: (next += 1) },
            content: content(block.blocks),
          };
        case "customXml":
          return {
            type: "blockCustomXml",
            openingXml: `<w:customXml xmlns:w="${W_NAMESPACE}" w:element="clause">`,
            closingXml: "</w:customXml>",
            content: content(block.blocks),
          };
        case "table":
          return {
            type: "table",
            columnWidths: [4800],
            rows: block.items.map((item) => ({
              type: "tableRow",
              cells: [
                { type: "tableCell", content: [paragraphOf(item, text())] },
              ],
            })),
          };
        default: {
          const unreachable: never = block;
          return unreachable;
        }
      }
    });
  return {
    package: {
      document: { content: content(blocks) },
      styles: {
        styles: [
          {
            styleId: "ListStyle",
            type: "paragraph",
            pPr: { numPr: { kind: "reference", numId: 3, ilvl: 0 } },
          },
        ],
      },
      numbering: {
        abstractNums: [
          { abstractNumId: 0, levels: [level(0), level(1), level(2)] },
        ],
        nums: [
          { numId: 1, abstractNumId: 0 },
          {
            numId: 2,
            abstractNumId: 0,
            levelOverrides: [{ ilvl: 0, startOverride: 3 }],
          },
          { numId: 3, abstractNumId: 0 },
        ],
      },
    },
  };
};

/** Every paragraph's rendering in document order, nested ones included. */
const renderings = (
  blocks: readonly BlockContent[],
): (ListRendering | null)[] =>
  blocks.flatMap((block) => {
    switch (block.type) {
      case "paragraph":
        return [block.listRendering ?? null];
      case "blockSdt":
      case "blockCustomXml":
        return renderings(block.content);
      case "table":
        return block.rows.flatMap(({ cells }) =>
          cells.flatMap((cell) => renderings(cell.content)),
        );
      default:
        return [];
    }
  });

const bodyRenderings = (document: Document) =>
  renderings(document.package.document.content);

/**
 * `document` with its last top-level paragraph a new, equal object: a cache
 * normalizes it again from the counters it kept for the blocks before.
 */
const withLastParagraphReplaced = (document: Document): Document => {
  const content = [...document.package.document.content];
  const index = content.findLastIndex(({ type }) => type === "paragraph");
  const last = content[index];
  if (last?.type !== "paragraph") return document;
  content[index] = { ...last, content: [...last.content] };
  return {
    ...document,
    package: {
      ...document.package,
      document: { ...document.package.document, content },
    },
  };
};

const parsedPackage = async (
  blocks: readonly GeneratedBlock[],
): Promise<Document> => {
  const prepared = await prepareCanonicalDocxInput(
    await createDocx(generatedDocument(blocks)),
  );
  if (prepared.isErr()) throw prepared.error;
  return parseDocx(prepared.value, {
    detectVariables: false,
    preloadFonts: false,
  });
};

describe("list normalization of a parsed package", () => {
  test(
    "keeps the parser's list renderings through content controls and custom XML",
    async () => {
      await assertProperty(
        fc.asyncProperty(
          fc.array(blockArbitrary, { minLength: 2, maxLength: 8 }),
          async (blocks) => {
            const parsed = await parsedPackage(blocks);
            const expected = bodyRenderings(parsed);
            expect(
              bodyRenderings(
                normalizeListModelRendering(parsed, { inverse: "omit" })
                  .document,
              ),
            ).toEqual(expected);
            // A session's cache, normalizing again past a changed block from
            // the counters its first run kept.
            const cache = new ListNormalizationCache();
            const first = normalizeListModelRendering(parsed, { cache });
            expect(bodyRenderings(first.document)).toEqual(expected);
            const again = normalizeListModelRendering(
              withLastParagraphReplaced(first.document),
              { cache },
            );
            expect(bodyRenderings(again.document)).toEqual(expected);
          },
        ),
        {
          id: "list normalization keeps the parser's list renderings through containers",
          numRuns: 60,
        },
      );
    },
    propertyTestTimeout(30_000),
  );
});

describe("derived list formatting inverses", () => {
  test(
    "restore review fields without authoring derived indentation",
    () => {
      assertProperty(
        fc.property(
          fc.integer({ min: 1, max: 2880 }),
          fc.boolean(),
          fc.option(fc.boolean(), { nil: undefined }),
          fc.option(fc.record({ spaceBefore: fc.nat({ max: 720 }) }), {
            nil: undefined,
          }),
          (indentLeft, styleNumbering, keepNext, currentFormatting) => {
            const paragraph: Paragraph = {
              type: "paragraph",
              paraId: "00000001",
              formatting: {
                ...(styleNumbering
                  ? { styleId: "ListStyle" }
                  : {
                      numPr: { kind: "reference" as const, numId: 1, ilvl: 0 },
                    }),
                ...(keepNext === undefined ? {} : { keepNext }),
              },
              propertyChanges: [
                {
                  type: "paragraphPropertyChange",
                  info: { id: 1, author: "Reviewer" },
                  previousFormatting: { alignment: "center" },
                  ...(currentFormatting === undefined
                    ? {}
                    : { currentFormatting }),
                },
              ],
              content: [],
            };
            const source: Document = {
              package: {
                document: { content: [paragraph] },
                styles: {
                  styles: [
                    {
                      styleId: "ListStyle",
                      type: "paragraph",
                      pPr: { numPr: { kind: "reference", numId: 1, ilvl: 0 } },
                    },
                  ],
                },
                numbering: {
                  abstractNums: [
                    {
                      abstractNumId: 0,
                      levels: [{ ...level(0), pPr: { indentLeft } }],
                    },
                  ],
                  nums: [{ numId: 1, abstractNumId: 0 }],
                },
              },
            };
            const normalized = normalizeListModelRendering(source);
            expect(normalized.inverse).toHaveLength(1);
            const rendered = normalized.document.package.document.content.at(0);
            expect(rendered?.type).toBe("paragraph");
            if (rendered?.type !== "paragraph") return;
            expect(rendered.formatting?.indentLeft).toBe(indentLeft);
            expect(rendered.propertyChanges).toStrictEqual(
              paragraph.propertyChanges,
            );
            // Serialized journal inverses must retain sparse records exactly too.
            const undone = applyDocumentOps(
              normalized.document,
              JSON.parse(JSON.stringify(normalized.inverse)),
            ).unwrap();
            const restored = undone.document.package.document.content.at(0);
            expect(restored?.type).toBe("paragraph");
            if (restored?.type !== "paragraph") return;
            expect(reviewFieldsOf(restored)).toStrictEqual(
              reviewFieldsOf(paragraph),
            );
            const redone = applyDocumentOps(
              undone.document,
              undone.inverse,
            ).unwrap();
            expect(redone.document).toStrictEqual(normalized.document);
            const edited = applyDocumentOps(normalized.document, [
              {
                type: DOCUMENT_OP_TYPES.SET_PARAGRAPH_PROPS,
                story: OP_STORIES.MAIN,
                blockId: "00000001",
                patch: { alignment: "right" },
              },
            ]).unwrap();
            expect(
              applyDocumentOps(edited.document, normalized.inverse).isErr(),
            ).toBe(true);
          },
        ),
        {
          id: "derived list formatting inverses retain pending review snapshots",
          numRuns: 100,
        },
      );
    },
    propertyTestTimeout(30_000),
  );
});
