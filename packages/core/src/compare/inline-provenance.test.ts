import { describe, expect, setDefaultTimeout, test } from "bun:test";
import fc from "fast-check";
import type { Node as PMNode } from "prosemirror-model";

import { propertyConfig, propertyTestTimeout } from "../../../../test/property-testing";
import type { Paragraph, TextFormatting } from "../types/document";
import { parseDocx } from "../docx/parser";
import { createDocx } from "../docx/rezip";
import { createEmptyDocument } from "../utils/createDocument";
import { FolioDocxReviewer } from "../ai-edits/headless";
import { createFolioAIEditSnapshot, sourceDocumentOf } from "../ai-edits/snapshot";
import type { FolioAIEditSnapshot } from "../ai-edits/types";
import { schema } from "../prosemirror/schema";
import { compareDocx } from "./compare";
import { sameAuthoredInlineProvenance } from "./inline-provenance";

setDefaultTimeout(propertyTestTimeout(30_000));

const OPTIONS = { author: "compare", timestamp: "2026-09-13T00:00:00.000Z" } as const;

type RunSpec = { text: string; formatting?: TextFormatting };

const INLINE_FORMATTING_CATALOG = [
  { bold: false },
  { italic: false },
  { boldCs: true },
  { italicCs: true },
  { fontFamily: { ascii: "Aptos", hAnsi: "Aptos", eastAsia: "Yu Gothic", cs: "Arial" } },
  { fontSize: 23, fontSizeCs: 27 },
  { color: { auto: true } },
  { color: { themeColor: "accent1", themeTint: "80" } },
  { color: { rgb: "C00000" } },
  { underline: { style: "double", color: { rgb: "0070C0" } } },
  { highlight: "yellow", strike: true },
] as const satisfies readonly TextFormatting[];

const partitionFormattingRuns = ({
  prefix,
  cuts,
  catalog = INLINE_FORMATTING_CATALOG,
}: {
  prefix: string;
  cuts: readonly boolean[];
  catalog?: readonly TextFormatting[];
}): RunSpec[] =>
  catalog.flatMap((formatting, index) => {
    const text = `${prefix}${String.fromCharCode(65 + index)}`;
    if (!cuts[index]) return [{ text, formatting }];
    return [
      { text: text.slice(0, 1), formatting },
      { text: text.slice(1), formatting },
    ];
  });

const partitionMatchingRun = (cuts: readonly boolean[]): RunSpec[] => {
  const text = "same";
  const chunks: RunSpec[] = [];
  let start = 0;
  for (let index = 1; index < text.length; index++) {
    if (!cuts[index - 1]) continue;
    chunks.push({ text: text.slice(start, index), formatting: { bold: false } });
    start = index;
  }
  chunks.push({ text: text.slice(start), formatting: { bold: false } });
  return chunks;
};

const MIXED_TARGET_RUNS = [
  {
    text: "Revised ",
    formatting: {
      bold: true,
      boldCs: false,
      italic: true,
      italicCs: false,
      fontFamily: {
        ascii: "Aptos",
        hAnsi: "Aptos",
        eastAsia: "Yu Gothic",
        cs: "Noto Naskh Arabic",
      },
      fontSize: 25,
      fontSizeCs: 27,
      color: { themeColor: "accent1", themeTint: "80" },
      underline: { style: "double", color: { rgb: "C00000" } },
      highlight: "yellow",
    },
  },
  {
    text: "terms",
    formatting: {
      bold: false,
      boldCs: true,
      italic: false,
      italicCs: true,
      fontFamily: { ascii: "Courier New", cs: "Arial" },
      fontSize: 22,
      fontSizeCs: 24,
      color: { auto: true },
      underline: { style: "single", color: { themeColor: "accent2" } },
      highlight: "cyan",
    },
  },
] as const satisfies readonly RunSpec[];

const paragraph = (paraId: string, runs: readonly RunSpec[]): Paragraph => ({
  type: "paragraph",
  paraId,
  textId: paraId,
  content: runs.map(({ text, formatting }) => ({
    type: "run",
    ...(formatting !== undefined && { formatting }),
    content: [{ type: "text", text }],
  })),
});

const documentWith = (paragraphs: readonly Paragraph[]): Promise<ArrayBuffer> => {
  const document = createEmptyDocument();
  document.package.document.content = [...paragraphs];
  return createDocx(document);
};

type FormattedCharacter = {
  paragraphIndex: number;
  text: string;
  formatting: TextFormatting | undefined;
};

const directFormattingCharacters = async (
  buffer: ArrayBuffer,
): Promise<readonly FormattedCharacter[]> => {
  const document = await parseDocx(buffer, { detectVariables: false, preloadFonts: false });
  const characters: FormattedCharacter[] = [];
  for (const [paragraphIndex, block] of document.package.document.content.entries()) {
    if (block.type !== "paragraph") continue;
    for (const run of block.content) {
      if (run.type !== "run") continue;
      for (const content of run.content) {
        if (content.type !== "text") continue;
        for (const text of Array.from(content.text)) {
          characters.push({ paragraphIndex, text, formatting: run.formatting });
        }
      }
    }
  }
  return characters;
};

const expectAcceptedAndRejectedFormatting = async ({
  base,
  revised,
}: {
  base: ArrayBuffer;
  revised: ArrayBuffer;
}): Promise<void> => {
  const [expectedBase, expectedRevised] = await Promise.all([
    directFormattingCharacters(base),
    directFormattingCharacters(revised),
  ]);
  const compared = await compareDocx(base, revised, OPTIONS);
  if (compared.isErr()) throw compared.error;

  const accepting = await FolioDocxReviewer.fromBuffer(compared.value.buffer);
  expect(accepting.acceptAll()).toBeGreaterThan(0);
  expect(await directFormattingCharacters(await accepting.toBuffer())).toEqual(expectedRevised);

  const rejecting = await FolioDocxReviewer.fromBuffer(compared.value.buffer);
  expect(rejecting.rejectAll()).toBeGreaterThan(0);
  expect(await directFormattingCharacters(await rejecting.toBuffer())).toEqual(expectedBase);
};

describe("inline formatting provenance through comparison", () => {
  test("accepting a replacement retains every revised direct run property", async () => {
    await expectAcceptedAndRejectedFormatting({
      base: await documentWith([paragraph("10000001", [{ text: "Base terms" }])]),
      revised: await documentWith([paragraph("20000001", MIXED_TARGET_RUNS)]),
    });
  });

  test("accepting an inserted paragraph retains target direct run properties", async () => {
    await expectAcceptedAndRejectedFormatting({
      base: await documentWith([paragraph("30000001", [{ text: "Anchor" }])]),
      revised: await documentWith([
        paragraph("40000001", MIXED_TARGET_RUNS),
        paragraph("30000001", [{ text: "Anchor" }]),
      ]),
    });
  });

  test("a text-equal direct run-property change has distinct accepted and rejected values", async () => {
    await expectAcceptedAndRejectedFormatting({
      base: await documentWith([
        paragraph("50000001", [
          {
            text: "Stable wording",
            formatting: {
              bold: false,
              boldCs: false,
              italic: false,
              italicCs: false,
              color: { auto: true },
              underline: { style: "none" },
            },
          },
        ]),
      ]),
      revised: await documentWith([
        paragraph("50000001", [
          {
            text: "Stable wording",
            formatting: {
              bold: true,
              boldCs: true,
              italic: true,
              italicCs: true,
              color: { themeColor: "accent3", themeShade: "40" },
              underline: { style: "dash", color: { rgb: "0070C0" } },
              highlight: "green",
            },
          },
        ]),
      ]),
    });
  });

  test("accept and reject preserve every direct formatting catalog value across mixed changes", async () => {
    await fc.assert(
      fc.asyncProperty(
        fc.record({
          baseCuts: fc.array(fc.boolean(), {
            minLength: INLINE_FORMATTING_CATALOG.length,
            maxLength: INLINE_FORMATTING_CATALOG.length,
          }),
          revisedCuts: fc.array(fc.boolean(), {
            minLength: INLINE_FORMATTING_CATALOG.length,
            maxLength: INLINE_FORMATTING_CATALOG.length,
          }),
          baseMatchingCuts: fc.array(fc.boolean(), { minLength: 3, maxLength: 3 }),
          revisedMatchingCuts: fc.array(fc.boolean(), { minLength: 3, maxLength: 3 }),
        }),
        async ({ baseCuts, revisedCuts, baseMatchingCuts, revisedMatchingCuts }) => {
          const base = await documentWith([
            paragraph("60000001", [
              ...partitionMatchingRun(baseMatchingCuts),
              { text: " base " },
              ...partitionFormattingRuns({ prefix: "B", cuts: baseCuts }),
              { text: " tail" },
            ]),
            paragraph("60000002", [{ text: "Anchor" }]),
          ]);
          const revised = await documentWith([
            paragraph("60000001", [
              ...partitionMatchingRun(revisedMatchingCuts),
              { text: " revised " },
              ...partitionFormattingRuns({ prefix: "R", cuts: revisedCuts }),
              { text: " tail" },
            ]),
            paragraph(
              "60000003",
              partitionFormattingRuns({
                prefix: "I",
                cuts: revisedCuts,
                catalog: [...INLINE_FORMATTING_CATALOG].reverse(),
              }),
            ),
            paragraph("60000002", [{ text: "Anchor" }]),
          ]);

          await expectAcceptedAndRejectedFormatting({ base, revised });
        },
      ),
      propertyConfig({ numRuns: 16 }),
    );
  });
});

/**
 * Make the document refuse to resolve a position, so provenance that reaches
 * for one fails here instead of quietly reintroducing an O(runs x blocks) term.
 */
const refusingToResolve = (doc: PMNode): PMNode => {
  Object.defineProperty(doc, "resolve", {
    configurable: true,
    value: () => {
      throw new Error("Inline provenance resolved a position; walk the document instead.");
    },
  });
  return doc;
};

describe("sameAuthoredInlineProvenance", () => {
  const provenanceParagraph = (label: string, bold: boolean, inherited = false): PMNode =>
    schema.node("paragraph", inherited ? { defaultTextFormatting: { bold: true } } : null, [
      schema.text(`${label} plain `),
      schema.text(`${label} bold`, bold ? [schema.mark("bold")] : []),
    ]);
  const tableOf = (content: PMNode[]): PMNode =>
    schema.node("table", null, [
      schema.node("tableRow", null, [schema.node("tableCell", null, content)]),
    ]);
  const documentOf = ({
    nestedBold = true,
    nestedInherited = false,
  }: { nestedBold?: boolean; nestedInherited?: boolean } = {}): PMNode =>
    schema.node("doc", null, [
      ...Array.from({ length: 50 }, (_unused, index) =>
        provenanceParagraph(`body ${String(index)}`, true, index % 2 === 0),
      ),
      tableOf([
        provenanceParagraph("outer cell", true),
        tableOf([provenanceParagraph("nested cell", nestedBold, nestedInherited)]),
        provenanceParagraph("after nested table", true, true),
      ]),
      provenanceParagraph("closing", true),
    ]);
  const snapshotOf = (doc: PMNode): FolioAIEditSnapshot => {
    const snapshot = createFolioAIEditSnapshot(doc);
    refusingToResolve(sourceDocumentOf(snapshot));
    return snapshot;
  };

  test("resolves no positions, so its cost stays linear in run count", () => {
    expect(sameAuthoredInlineProvenance(snapshotOf(documentOf()), snapshotOf(documentOf()))).toBe(
      true,
    );
    expect(
      sameAuthoredInlineProvenance(
        snapshotOf(documentOf()),
        snapshotOf(documentOf({ nestedBold: false })),
      ),
    ).toBe(false);
  });

  test("reads each run against the paragraph that encloses it", () => {
    // The same bold run is direct in one paragraph and inherited in the other;
    // only the nested cell's own paragraph can tell the two apart.
    expect(
      sameAuthoredInlineProvenance(
        snapshotOf(documentOf()),
        snapshotOf(documentOf({ nestedInherited: true })),
      ),
    ).toBe(false);
  });
});
