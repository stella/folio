/**
 * A no-op Document → ProseMirror → Document rebuild is a fixed point on the
 * paragraph model. The per-field round-trip tests next to this file each pin
 * one attr that once went missing; this pins the whole paragraph, so a model
 * field without an inverse projection fails here on the first fixture that
 * carries it instead of waiting for its own bug report (issue #845).
 *
 * `PARAGRAPH_NORMALIZATIONS` lists the fields the rebuild is allowed to
 * change, each with the reason. Growing that list is a review decision.
 */

import { describe, expect, test } from "bun:test";
import fc from "fast-check";
import { readdirSync } from "node:fs";
import path from "node:path";

import { parseDocx } from "../../docx/parser";
import { createDocx } from "../../docx/rezip";
import { toFlowBlocks } from "../../layout-bridge/convert/toFlowBlocks";
import { fromMarkdown } from "../../markdown/fromMarkdown";
import type { BlockContent, Document, ListLevel, Paragraph } from "../../types/document";
import { updateDocumentContent } from "./fromProseDoc";
import { toProseDoc } from "./toProseDoc";
import { paragraphNumberingReferenceId } from "@stll/docx-core/model";

const CORPUS_DIR = path.join(import.meta.dir, "../../docx/__tests__/__fixtures__/corpus");
const VISUAL_FIXTURES_DIR = path.join(import.meta.dir, "../../../../../tests/visual/fixtures");

const CORPUS_FIXTURES = readdirSync(CORPUS_DIR)
  .filter((name) => name.endsWith(".docx"))
  .map((name) => path.join(CORPUS_DIR, name));
const VISUAL_FIXTURES = ["sample.docx", "docx-editor-demo.docx"].map((name) =>
  path.join(VISUAL_FIXTURES_DIR, name),
);

/** Paragraph fields the rebuild may legitimately change, with the reason. */
const PARAGRAPH_NORMALIZATIONS: readonly (keyof Paragraph)[] = [
  // Run content is owned by the inline round-trip tests; marks and runs merge.
  "content",
];

const collectParagraphs = (blocks: readonly BlockContent[], into: Paragraph[]): Paragraph[] => {
  for (const block of blocks) {
    switch (block.type) {
      case "paragraph": {
        into.push(block);
        break;
      }
      case "table": {
        for (const row of block.rows) {
          for (const cell of row.cells) {
            collectParagraphs(cell.content, into);
          }
        }
        break;
      }
      case "blockSdt": {
        collectParagraphs(block.content, into);
        break;
      }
      default: {
        block satisfies never;
      }
    }
  }
  return into;
};

const NORMALIZED_KEYS: ReadonlySet<string> = new Set(PARAGRAPH_NORMALIZATIONS);

const comparableParagraph = (paragraph: Paragraph): Record<string, unknown> =>
  Object.fromEntries(Object.entries(paragraph).filter(([key]) => !NORMALIZED_KEYS.has(key)));

const rebuild = (document: Document): Document =>
  updateDocumentContent(
    document,
    toProseDoc(document, { styles: document.package.styles, theme: document.package.theme }),
  );

const expectParagraphFixedPoint = (document: Document): void => {
  const before = collectParagraphs(document.package.document.content, []);
  const after = collectParagraphs(rebuild(document).package.document.content, []);
  expect(after.length).toBe(before.length);
  for (const [index, paragraph] of before.entries()) {
    expect(comparableParagraph(after[index] ?? paragraph), `paragraph ${index}`).toEqual(
      comparableParagraph(paragraph),
    );
  }
};

const parseFixture = async (fixturePath: string): Promise<Document> =>
  parseDocx(await Bun.file(fixturePath).arrayBuffer(), {
    preloadFonts: false,
    detectVariables: false,
  });

describe("paragraph model fixed point under a no-op rebuild", () => {
  for (const fixturePath of [...CORPUS_FIXTURES, ...VISUAL_FIXTURES]) {
    test(path.basename(fixturePath), async () => {
      expectParagraphFixedPoint(await parseFixture(fixturePath));
    });
  }
});

const NUM_FMTS = ["decimal", "lowerLetter", "upperRoman", "bullet"] as const;

const listLevelArb = (ilvl: number): fc.Arbitrary<ListLevel> =>
  fc
    .record({
      start: fc.integer({ min: 1, max: 9 }),
      numFmt: fc.constantFrom(...NUM_FMTS),
    })
    .map(({ start, numFmt }) => ({
      ilvl,
      start,
      numFmt,
      lvlText: numFmt === "bullet" ? "•" : `%${ilvl + 1}.`,
      pPr: { indentation: { left: 720 * (ilvl + 1), hanging: 360 } },
    }));

const numberedDocumentArb: fc.Arbitrary<Document> = fc
  .record({
    itemLevels: fc.array(fc.integer({ min: 0, max: 2 }), { minLength: 1, maxLength: 5 }),
    levels: fc.tuple(listLevelArb(0), listLevelArb(1), listLevelArb(2)),
    startOverride: fc.option(fc.integer({ min: 1, max: 9 }), { nil: undefined }),
  })
  .map(({ itemLevels, levels, startOverride }) => {
    const markdown = itemLevels.map((_, index) => `${index + 1}. Item ${index + 1}`).join("\n");
    const model = fromMarkdown(`${markdown}\n\nTail.`);
    const paragraphs = collectParagraphs(model.package.document.content, []);
    const numId = paragraphNumberingReferenceId(paragraphs.at(0)?.formatting?.numPr);
    const instance = model.package.numbering?.nums.find((num) => num.numId === numId);
    const abstract = model.package.numbering?.abstractNums.find(
      (definition) => definition.abstractNumId === instance?.abstractNumId,
    );
    if (numId === undefined || !instance || !abstract) {
      throw new Error("generated list must carry a numbering definition");
    }
    abstract.levels = [...levels];
    for (const [index, ilvl] of itemLevels.entries()) {
      const paragraph = paragraphs.at(index);
      if (paragraph === undefined) {
        throw new Error("generated list is shorter than its level list");
      }
      paragraph.formatting = { ...paragraph.formatting, numPr: { kind: "reference", numId, ilvl } };
    }
    if (startOverride !== undefined) {
      instance.levelOverrides = [{ ilvl: 0, startOverride }];
    }
    return model;
  });

const markers = (document: Document): (string | null)[] =>
  toFlowBlocks(toProseDoc(document))
    .filter((block) => block.kind === "paragraph")
    .map((block) => block.attrs?.listMarker ?? null);

describe("list rendering fixed point under a no-op rebuild", () => {
  test("generated numbering definitions keep their metadata and markers", async () => {
    await fc.assert(
      fc.asyncProperty(numberedDocumentArb, async (model) => {
        const document = await parseDocx(await createDocx(model), {
          preloadFonts: false,
          detectVariables: false,
        });
        expectParagraphFixedPoint(document);
        expect(markers(rebuild(document))).toEqual(markers(document));
      }),
      { numRuns: 25 },
    );
  });
});
