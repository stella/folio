/**
 * Block-level public operations in the canonical session give the document
 * the legacy engine gives: the same blocks and paragraph properties pending,
 * once every change is accepted, and once every change is rejected.
 */

import { panic } from "better-result";
import { describe, expect, test } from "bun:test";
import { EditorState } from "prosemirror-state";

import { paragraphNumberingReference } from "@stll/docx-core/model";
import { ZipArchive } from "@stll/docx-core/zip";

import type { FolioDocumentOperation } from "../ai-edits";
import { SUGGESTED_SUPPORTED_OPERATION_TYPES } from "../ai-edits/apply";
import { FolioDocxReviewer } from "../ai-edits/headless";
import { createFolioAIEditSnapshot } from "../ai-edits/snapshot";
import {
  type CanonicalCommit,
  createCanonicalSession,
  publishCanonicalProjection,
} from "../canonical/canonicalSession";
import { serializeCanonicalSave } from "../docx/canonicalSave";
import { parseDocx } from "../docx/parser";
import { schema } from "../prosemirror/schema";
import { CanonicalPublicOperations } from "./canonicalPublicOperations";

const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const W14 = "http://schemas.microsoft.com/office/word/2010/wordml";
const HEAD = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';

const paragraph = (id: string, text: string, properties = "") =>
  `<w:p w14:paraId="${id}">${properties ? `<w:pPr>${properties}</w:pPr>` : ""}${text ? `<w:r><w:t xml:space="preserve">${text}</w:t></w:r>` : ""}</w:p>`;
const table = (id: string, text: string) =>
  `<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/></w:tblPr><w:tblGrid><w:gridCol w:w="2000"/></w:tblGrid><w:tr><w:tc><w:tcPr><w:tcW w:w="2000" w:type="dxa"/></w:tcPr>${paragraph(id, text)}</w:tc></w:tr></w:tbl>`;

/** Bodies, each holding the paragraph `10000002` the operations address. */
const BODIES = {
  "between two paragraphs": [
    paragraph("10000001", "Alpha", '<w:jc w:val="center"/>'),
    paragraph("10000002", "Beta", '<w:ind w:left="720"/>'),
    paragraph("10000003", "Gamma", '<w:spacing w:after="240"/>'),
  ],
  "ending the body": [
    paragraph("10000001", "Alpha", '<w:jc w:val="center"/>'),
    paragraph("10000002", "Beta", '<w:ind w:left="720"/>'),
  ],
  "before a table": [
    paragraph("10000001", "Alpha"),
    paragraph("10000002", "Beta", '<w:ind w:left="720"/>'),
    table("10000004", "Cell"),
    paragraph("10000003", "Gamma"),
  ],
  "before a block content control": [
    paragraph("10000001", "Alpha"),
    paragraph("10000002", "Beta", '<w:ind w:left="720"/>'),
    `<w:sdt><w:sdtPr><w:id w:val="5"/></w:sdtPr><w:sdtContent>${paragraph("10000004", "Inside")}</w:sdtContent></w:sdt>`,
    paragraph("10000003", "Gamma"),
  ],
} as const;

const packageOf = async (body: readonly string[]) => {
  const zip = await ZipArchive.read(
    await (await import("../docx/rezip")).createEmptyDocx(),
  );
  const styles = (await zip.entry("word/styles.xml")?.text()) ?? "";
  // A heading style with its own alignment and spacing, for a restyle.
  zip.set(
    "word/styles.xml",
    styles.replace(
      "</w:styles>",
      '<w:style w:type="paragraph" w:styleId="Heading1"><w:name w:val="heading 1"/><w:basedOn w:val="Normal"/><w:pPr><w:jc w:val="center"/><w:spacing w:before="480"/><w:outlineLvl w:val="0"/></w:pPr></w:style></w:styles>',
    ),
  );
  // One decimal list, numId 1, for numbering operations to name.
  zip.set(
    "word/numbering.xml",
    `${HEAD}<w:numbering xmlns:w="${W}"><w:abstractNum w:abstractNumId="0"><w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/><w:pPr><w:ind w:left="720" w:hanging="360"/></w:pPr></w:lvl><w:lvl w:ilvl="1"><w:start w:val="1"/><w:numFmt w:val="lowerLetter"/><w:lvlText w:val="%2)"/><w:pPr><w:ind w:left="1440" w:hanging="360"/></w:pPr></w:lvl></w:abstractNum><w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num></w:numbering>`,
  );
  const rels = (await zip.entry("word/_rels/document.xml.rels")?.text()) ?? "";
  zip.set(
    "word/_rels/document.xml.rels",
    rels.replace(
      "</Relationships>",
      '<Relationship Id="rId90" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering" Target="numbering.xml"/></Relationships>',
    ),
  );
  const types = (await zip.entry("[Content_Types].xml")?.text()) ?? "";
  zip.set(
    "[Content_Types].xml",
    types.replace(
      "</Types>",
      '<Override PartName="/word/numbering.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"/></Types>',
    ),
  );
  zip.set(
    "word/document.xml",
    `${HEAD}<w:document xmlns:w="${W}" xmlns:w14="${W14}"><w:body>${body.join("")}<w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr></w:body></w:document>`,
  );
  return (await zip.generate({ compression: "store" })).buffer;
};

/** Each body block's text and paragraph properties, as a package reads. */
const blocksOf = async (bytes: ArrayBuffer) => {
  const document = await parseDocx(bytes.slice(0), {
    preloadFonts: false,
    detectVariables: false,
  });
  return document.package.document.content.map((block) =>
    block.type === "paragraph"
      ? {
          text: block.content
            .flatMap((item) =>
              item.type === "run"
                ? item.content.flatMap((leaf) => {
                    if (leaf.type === "text") return [leaf.text];
                    if (leaf.type === "tab") return ["\t"];
                    if (leaf.type !== "break") return [];
                    return [leaf.breakType === "page" ? "\f" : "\n"];
                  })
                : [],
            )
            .join(""),
          // Source paragraphs keep their ids; a new one's is the engine's own.
          id: block.paraId?.startsWith("1000000") ? block.paraId : "new",
          pageBreakBefore: block.formatting?.pageBreakBefore,
          runFormatting: block.content.flatMap((item) =>
            item.type === "run" && item.formatting !== undefined
              ? [item.formatting]
              : [],
          ),
          styleId: block.formatting?.styleId,
          numPr: block.formatting?.numPr,
          outlineLevel: block.formatting?.outlineLevel,
          alignment: block.formatting?.alignment,
          indentLeft: block.formatting?.indentLeft,
          indentRight: block.formatting?.indentRight,
          spaceBefore: block.formatting?.spaceBefore,
          spaceAfter: block.formatting?.spaceAfter,
        }
      : { type: block.type },
  );
};

/** The package pending, with every change accepted, and with every change rejected. */
const outcomes = async (bytes: ArrayBuffer) => {
  const resolved = async (decision: "accept" | "reject") => {
    const reviewer = await FolioDocxReviewer.fromBuffer(bytes.slice(0));
    if (decision === "accept") reviewer.acceptAll();
    else reviewer.rejectAll();
    return blocksOf(await reviewer.toBuffer());
  };
  const commented = async (decision?: "accept") => {
    const reviewer = await FolioDocxReviewer.fromBuffer(bytes.slice(0));
    if (decision === "accept") reviewer.acceptAll();
    return reviewer
      .getComments()
      .map(({ anchoredText, text }) => ({ anchoredText, text }));
  };
  const numbering = (
    await parseDocx(bytes.slice(0), {
      preloadFonts: false,
      detectVariables: false,
    })
  ).package.numbering;
  return {
    lists: {
      nums: numbering?.nums.map(({ numId, abstractNumId }) => [
        numId,
        abstractNumId,
      ]),
      formats: numbering?.abstractNums.map(({ abstractNumId, levels }) => [
        abstractNumId,
        levels.map(({ numFmt }) => numFmt),
      ]),
    },
    pending: await blocksOf(bytes),
    accepted: await resolved("accept"),
    rejected: await resolved("reject"),
    comments: {
      pending: await commented(),
      accepted: await commented("accept"),
    },
  };
};

const withoutIds = (blocks: readonly object[]) =>
  blocks.map((block) => ({ ...block, id: undefined }));

/**
 * Shapes where the legacy engine, reading the canonical session's file back,
 * rejects an insertion after the anchor by joining the anchor into the new
 * paragraph under the new paragraph's id. The file records the anchor as the
 * paragraph that stays, as it does at the end of a container; the canonical
 * session's own rejection gives the source back, ids included, which the
 * suggested half of each case checks.
 */
const READ_BACK_UNDER_NEW_ID: ReadonlySet<string> = new Set([
  "before a block content control",
]);

/**
 * The canonical session's outcomes against the legacy engine's. Rejecting a
 * tracked insertion after a container's final paragraph gives the source back
 * from the canonical session's file, ids included: it records which paragraph
 * rejecting the mark keeps. The legacy engine's file does not, so there the
 * anchor's words come back under the new paragraph's id. Everything else is
 * the same.
 */
const expectOutcomes = async (
  actual: Awaited<ReturnType<typeof outcomes>>,
  expected: Awaited<ReturnType<typeof outcomes>>,
  source: ArrayBuffer,
  shape?: string,
) => {
  const sourceBlocks = await blocksOf(source);
  const idsOnly =
    !Bun.deepEquals(actual.rejected, expected.rejected) &&
    Bun.deepEquals(withoutIds(actual.rejected), withoutIds(expected.rejected));
  if (idsOnly && Bun.deepEquals(actual.rejected, sourceBlocks)) {
    expect({ ...actual, rejected: expected.rejected }).toEqual(expected);
    return;
  }
  if (
    idsOnly &&
    shape !== undefined &&
    READ_BACK_UNDER_NEW_ID.has(shape) &&
    Bun.deepEquals(expected.rejected, sourceBlocks)
  ) {
    expect({ ...actual, rejected: expected.rejected }).toEqual(expected);
    return;
  }
  expect(actual).toEqual(expected);
};

type Mode = "direct" | "tracked-changes";

const batchOf = (
  operation: FolioDocumentOperation | readonly FolioDocumentOperation[],
): FolioDocumentOperation[] =>
  Array.isArray(operation)
    ? [...operation]
    : [operation as FolioDocumentOperation];
type Suggestion = "accept" | "reject";

const viaDefaultEngine = async (
  source: ArrayBuffer,
  operation: FolioDocumentOperation | readonly FolioDocumentOperation[],
  mode: Mode,
) => {
  const reviewer = await FolioDocxReviewer.fromBuffer(source.slice(0));
  const result = reviewer.applyOperations(batchOf(operation), {
    mode,
    undefinedReferences: "refuse",
  });
  expect(result.skipped).toEqual([]);
  return await reviewer.toBuffer();
};

/**
 * The canonical session's saved package after the batch; a suggested batch
 * is resolved as `suggestion` says before the save.
 */
/** The legacy engine's package after a suggested batch it then resolves. */
const viaDefaultSuggested = async (
  source: ArrayBuffer,
  operation: FolioDocumentOperation | readonly FolioDocumentOperation[],
  resolution: Suggestion,
) => {
  const reviewer = await FolioDocxReviewer.fromBuffer(source.slice(0));
  const result = reviewer.applyOperations(batchOf(operation), {
    mode: "suggested",
    undefinedReferences: "refuse",
  });
  expect(result.skipped).toEqual([]);
  for (const { id } of batchOf(operation))
    if (resolution === "accept")
      expect(reviewer.acceptSuggestion(id)).toBe(true);
    else expect(reviewer.rejectSuggestion(id)).toBe(true);
  return await reviewer.toBuffer();
};

const canonicalRun = async (
  source: ArrayBuffer,
  operation: FolioDocumentOperation | readonly FolioDocumentOperation[],
  mode: Mode | "suggested",
) => {
  const document = await parseDocx(source.slice(0), {
    preloadFonts: false,
    detectVariables: false,
  });
  const session = createCanonicalSession(
    document,
    document.package.styles,
  ).unwrap();
  let state = EditorState.create({ schema, doc: session.projection.doc });
  const executor = new CanonicalPublicOperations({
    session,
    getState: () => state,
    publish: (commit: CanonicalCommit) => {
      const published = publishCanonicalProjection({ state, commit, session });
      if (published.isErr()) return false;
      state = published.value.state;
      return true;
    },
  });
  const result = executor.apply({
    snapshot: createFolioAIEditSnapshot(state.doc),
    batch: { version: 1, mode, operations: batchOf(operation) },
    undefinedReferences: "refuse",
  });
  return { session, state: () => state, skipped: result.skipped };
};

const viaCanonicalEngine = async (
  source: ArrayBuffer,
  operation: FolioDocumentOperation | readonly FolioDocumentOperation[],
  mode: Mode | "suggested",
  suggestion?: Suggestion,
) => {
  const run = await canonicalRun(source, operation, mode);
  expect(run.skipped).toEqual([]);
  const { session } = run;
  let state = run.state();
  // Each operation of a suggested batch is its own suggestion.
  for (const { id } of batchOf(operation)) {
    if (mode !== "suggested") break;
    const commit = session
      .prepareSuggestion(state, {
        suggestionId: id,
        resolution: suggestion ?? "reject",
        author: "Adopter",
        date: "2026-02-02T00:00:00Z",
      })
      .unwrap();
    const published = publishCanonicalProjection({ state, commit, session });
    if (published.isErr()) throw published.error;
    state = published.value.state;
  }
  return (
    await serializeCanonicalSave({ snapshot: session.captureSaveSnapshot() })
  ).buffer;
};

/** Cases the canonical session declares unavailable, with its reason. */
const REFUSED: Record<string, string> = {};

/** Cases both engines skip, with the reason both give. */
const SKIPPED_BY_BOTH: Record<string, string> = {
  // A direct deletion leaves no text for its comment to cover.
  "deleteBlock commented between two paragraphs, direct": "unsupportedBlock",
  "deleteBlock commented ending the body, direct": "unsupportedBlock",
  "deleteBlock commented before a table, direct": "unsupportedBlock",
  "deleteBlock commented before a block content control, direct":
    "unsupportedBlock",
};

const OPERATIONS = {
  deleteBlock: { id: "op", type: "deleteBlock", blockId: "10000002" },
  insertAfterBlock: {
    id: "op",
    type: "insertAfterBlock",
    blockId: "10000002",
    text: "Inserted",
  },
  insertBeforeBlock: {
    id: "op",
    type: "insertBeforeBlock",
    blockId: "10000002",
    text: "Inserted",
  },
  "insertAfterBlock blank": {
    id: "op",
    type: "insertAfterBlock",
    blockId: "10000002",
    text: "",
  },
  "insertAfterBlock without inheritance": {
    id: "op",
    type: "insertAfterBlock",
    blockId: "10000002",
    text: "Inserted",
    inheritFormatting: false,
  },
  "insertBeforeBlock restyled": {
    id: "op",
    type: "insertBeforeBlock",
    blockId: "10000002",
    text: "Inserted",
    styleId: "Heading1",
    alignment: null,
    spacing: { spaceAfter: 60 },
    pageBreakBefore: true,
  },
  "insertAfterBlock lines": {
    id: "op",
    type: "insertAfterBlock",
    blockId: "10000002",
    text: "Heading\n\nFirst\tbody\r\nSecond body\n",
    alignment: "center",
  },
  "insertBeforeBlock lines formatted alike": {
    id: "op",
    type: "insertBeforeBlock",
    blockId: "10000002",
    text: "One\nTwo",
    styleId: "Heading1",
    formattingScope: "allParagraphs",
    pageBreakBefore: true,
  },
  "insertAfterBlock inline breaks": {
    id: "op",
    type: "insertAfterBlock",
    blockId: "10000002",
    text: "Name\tValue\nNext line",
    lineBreakMode: "inline",
  },
  "insertBeforeBlock page break": {
    id: "op",
    type: "insertBeforeBlock",
    blockId: "10000002",
    text: "",
    hardPageBreak: {},
  },
  "deleteBlock commented": {
    id: "op",
    type: "deleteBlock",
    blockId: "10000002",
    comment: { text: "Why this goes" },
  },
  "insertAfterBlock commented": {
    id: "op",
    type: "insertAfterBlock",
    blockId: "10000002",
    text: "Commented\nSecond line",
    comment: { text: "Why this clause" },
  },
  "insertAfterBlock numbered": {
    id: "op",
    type: "insertAfterBlock",
    blockId: "10000002",
    text: "Item",
    numbering: paragraphNumberingReference({ numId: 1, ilvl: 1 }),
  },
  "insertBeforeBlock numbering cancelled": {
    id: "op",
    type: "insertBeforeBlock",
    blockId: "10000002",
    text: "Plain",
    numbering: { kind: "none" },
  },
  "setBlockParagraphProperties numbered": {
    id: "op",
    type: "setBlockParagraphProperties",
    blockId: "10000002",
    properties: {
      numbering: paragraphNumberingReference({ numId: 1, ilvl: 0 }),
    },
  },
  "setBlockParagraphProperties numbering level": {
    id: "op",
    type: "setBlockParagraphProperties",
    blockId: "10000002",
    properties: { numbering: { kind: "levelOnly", ilvl: 1 } },
  },
  "insertAfterBlock new list": {
    id: "op",
    type: "insertAfterBlock",
    blockId: "10000002",
    text: "First\nSecond",
    numbering: { kind: "newList", format: "numbered" },
    formattingScope: "allParagraphs",
  },
  "setBlockParagraphProperties new list": {
    id: "op",
    type: "setBlockParagraphProperties",
    blockId: "10000002",
    properties: { numbering: { kind: "newList", format: "bullet", level: 1 } },
  },
  mergeBlockWithNext: {
    id: "op",
    type: "mergeBlockWithNext",
    blockId: "10000001",
  },
  splitBlock: {
    id: "op",
    type: "splitBlock",
    blockId: "10000002",
    offset: 2,
  },
  "splitBlock first half's properties": {
    id: "op",
    type: "splitBlock",
    blockId: "10000002",
    offset: 2,
    firstParagraphProperties: { alignment: "right", indentation: null },
  },
  "splitBlock second half's properties": {
    id: "op",
    type: "splitBlock",
    blockId: "10000002",
    offset: 2,
    secondParagraphProperties: {
      styleId: "Heading1",
      spacing: { spaceAfter: 120 },
    },
  },
  "splitBlock both halves' properties": {
    id: "op",
    type: "splitBlock",
    blockId: "10000002",
    offset: 2,
    firstParagraphProperties: {
      numbering: paragraphNumberingReference({ numId: 1, ilvl: 0 }),
    },
    secondParagraphProperties: {
      numbering: paragraphNumberingReference({ numId: 1, ilvl: 1 }),
    },
  },
  "splitBlock second half in a new list": {
    id: "op",
    type: "splitBlock",
    blockId: "10000002",
    offset: 2,
    secondParagraphProperties: {
      numbering: { kind: "newList", format: "bullet" },
    },
  },
  "mergeBlockWithNext separated": {
    id: "op",
    type: "mergeBlockWithNext",
    blockId: "10000001",
    separator: " ",
  },
  "mergeBlockWithNext with properties": {
    id: "op",
    type: "mergeBlockWithNext",
    blockId: "10000001",
    mergedParagraphProperties: {
      alignment: "right",
      spacing: { spaceAfter: 120 },
    },
  },
  "setBlockParagraphProperties alignment": {
    id: "op",
    type: "setBlockParagraphProperties",
    blockId: "10000002",
    properties: { alignment: "right" },
  },
  "setBlockParagraphProperties spacing and cleared indentation": {
    id: "op",
    type: "setBlockParagraphProperties",
    blockId: "10000002",
    properties: { spacing: { spaceBefore: 120 }, indentation: null },
  },
  "setBlockParagraphProperties partial indentation": {
    id: "op",
    type: "setBlockParagraphProperties",
    blockId: "10000002",
    properties: {
      indentation: { indentRight: 360 },
      outlineLevel: { kind: "heading", level: 1 },
    },
  },
  "setBlockParagraphProperties style": {
    id: "op",
    type: "setBlockParagraphProperties",
    blockId: "10000002",
    properties: { styleId: "Heading1" },
  },
} as const satisfies Record<string, FolioDocumentOperation>;

describe("block-level public operations match the legacy engine", () => {
  for (const [kind, operation] of Object.entries(OPERATIONS))
    for (const [shape, body] of Object.entries(BODIES))
      for (const mode of ["direct", "tracked-changes"] as const) {
        const name = `${kind} ${shape}, ${mode}`;
        const refused = REFUSED[name];
        test(name, async () => {
          const source = await packageOf(body);
          const skippedByBoth = SKIPPED_BY_BOTH[name];
          if (skippedByBoth !== undefined) {
            const reviewer = await FolioDocxReviewer.fromBuffer(
              source.slice(0),
            );
            const byDefault = reviewer.applyOperations([operation], {
              mode,
              undefinedReferences: "refuse",
            });
            const { skipped } = await canonicalRun(source, operation, mode);
            expect(byDefault.skipped.map(({ reason }) => reason)).toEqual([
              skippedByBoth,
            ]);
            expect(skipped.map(({ reason }) => reason)).toEqual([
              skippedByBoth,
            ]);
            return;
          }
          if (refused !== undefined) {
            const { skipped } = await canonicalRun(source, operation, mode);
            expect(skipped.map(({ message }) => message)).toEqual([refused]);
            return;
          }
          const expected = await outcomes(
            await viaDefaultEngine(source, operation, mode),
          );
          const actual = await outcomes(
            await viaCanonicalEngine(source, operation, mode),
          );
          await expectOutcomes(actual, expected, source, shape);
          // Kinds a suggestion cannot carry skip in both engines.
          if (
            mode !== "tracked-changes" ||
            !SUGGESTED_SUPPORTED_OPERATION_TYPES.has(operation.type)
          )
            return;
          // Suggested, the edit saves as nothing until it is resolved:
          // rejecting gives the source back, accepting the tracked edit.
          expect(
            await blocksOf(
              await viaCanonicalEngine(
                source,
                operation,
                "suggested",
                "reject",
              ),
            ),
          ).toEqual(await blocksOf(source));
          await expectOutcomes(
            await outcomes(
              await viaCanonicalEngine(
                source,
                operation,
                "suggested",
                "accept",
              ),
            ),
            expected,
            source,
            shape,
          );
          // The legacy engine's suggestion resolves the same way; it
          // carries no splits or merges.
          if (
            operation.type === "splitBlock" ||
            operation.type === "mergeBlockWithNext"
          )
            return;
          expect(
            await blocksOf(
              await viaDefaultSuggested(source, operation, "reject"),
            ),
          ).toEqual(await blocksOf(source));
          expect(
            await outcomes(
              await viaDefaultSuggested(source, operation, "accept"),
            ),
          ).toEqual(expected);
        });
      }
});

describe("a split's second half after an insertion in one batch", () => {
  // The insertion after the later paragraph applies first and takes a new
  // paragraph id; the split's second half takes the next one, and its
  // properties follow it there.
  for (const mode of ["direct", "tracked-changes"] as const)
    test(mode, async () => {
      const source = await packageOf(BODIES["between two paragraphs"]);
      const operations = [
        {
          id: "split",
          type: "splitBlock",
          blockId: "10000001",
          offset: 2,
          secondParagraphProperties: { alignment: "right" },
        },
        {
          id: "insert",
          type: "insertAfterBlock",
          blockId: "10000003",
          text: "Added",
        },
      ] as const satisfies readonly FolioDocumentOperation[];
      const run = await canonicalRun(source, operations, mode);
      expect(run.skipped).toEqual([]);
      await expectOutcomes(
        await outcomes(await viaCanonicalEngine(source, operations, mode)),
        await outcomes(await viaDefaultEngine(source, operations, mode)),
        source,
      );
    });
});

describe("inline line breaks", () => {
  // CRLF and a lone CR are each one line break, as LF is. The legacy engine
  // keeps a carriage return as a text character, which saves as no break.
  for (const mode of ["direct", "tracked-changes"] as const)
    test(`CRLF and CR insert one line break each, ${mode}`, async () => {
      const source = await packageOf(BODIES["ending the body"]);
      const inline = (text: string) =>
        ({
          id: "op",
          type: "insertAfterBlock",
          blockId: "10000002",
          text,
          lineBreakMode: "inline",
        }) as const satisfies FolioDocumentOperation;
      const run = await canonicalRun(
        source,
        inline("Name\r\nNext\rLast"),
        mode,
      );
      expect(run.skipped).toEqual([]);
      expect(
        await outcomes(
          await viaCanonicalEngine(source, inline("Name\r\nNext\rLast"), mode),
        ),
      ).toEqual(
        await outcomes(
          await viaDefaultEngine(source, inline("Name\nNext\nLast"), mode),
        ),
      );
    });
});

describe("a new list beside a dangling list reference", () => {
  // A paragraph names list 2, which the package does not define. A new list
  // takes another id, so that paragraph stays out of it.
  for (const mode of ["direct", "tracked-changes"] as const)
    test(`takes an id no paragraph names, ${mode}`, async () => {
      const source = await packageOf([
        paragraph(
          "10000001",
          "Dangling",
          '<w:numPr><w:ilvl w:val="0"/><w:numId w:val="2"/></w:numPr>',
        ),
        paragraph("10000002", "Beta"),
      ]);
      const operation = {
        id: "op",
        type: "insertAfterBlock",
        blockId: "10000002",
        text: "Item",
        numbering: { kind: "newList", format: "bullet" },
      } as const satisfies FolioDocumentOperation;
      const run = await canonicalRun(source, operation, mode);
      expect(run.skipped).toEqual([]);
      const expected = await outcomes(
        await viaDefaultEngine(source, operation, mode),
      );
      const actual = await outcomes(
        await viaCanonicalEngine(source, operation, mode),
      );
      // The paragraph that named list 2 stays out of the new list.
      for (const state of ["pending", "accepted", "rejected"] as const)
        expect(actual[state].at(0)?.numPr).toEqual({ kind: "none" });
      expect(actual).toEqual(expected);
    });
});

describe("an insertion beside a table cell's paragraph", () => {
  // The legacy engine places it beside the table, not inside the cell.
  for (const type of ["insertAfterBlock", "insertBeforeBlock"] as const)
    for (const mode of ["direct", "tracked-changes"] as const)
      test(`${type}, ${mode}`, async () => {
        const source = await packageOf(BODIES["before a table"]);
        const operation = {
          id: "op",
          type,
          blockId: "10000004",
          text: "Beside",
        } as const satisfies FolioDocumentOperation;
        const run = await canonicalRun(source, operation, mode);
        expect(run.skipped).toEqual([]);
        expect(
          await outcomes(await viaCanonicalEngine(source, operation, mode)),
        ).toEqual(
          await outcomes(await viaDefaultEngine(source, operation, mode)),
        );
      });
});

describe("paragraph identities across undo", () => {
  // An undone insertion's paragraph id stays retired: a later insertion takes
  // another, so an operation read before the undo cannot reach the new one.
  for (const [type, operation] of [
    [
      "insertAfterBlock",
      { id: "op", type: "insertAfterBlock", blockId: "10000002", text: "New" },
    ],
    [
      "splitBlock",
      { id: "op", type: "splitBlock", blockId: "10000002", offset: 2 },
    ],
  ] as const satisfies readonly (readonly [string, FolioDocumentOperation])[])
    for (const mode of ["direct", "tracked-changes"] as const)
      test(`${type} after an undone one takes a fresh id, ${mode}`, async () => {
        const source = await packageOf(BODIES["between two paragraphs"]);
        const document = await parseDocx(source.slice(0), {
          preloadFonts: false,
          detectVariables: false,
        });
        const session = createCanonicalSession(
          document,
          document.package.styles,
        ).unwrap();
        let state = EditorState.create({ schema, doc: session.projection.doc });
        const executor = new CanonicalPublicOperations({
          session,
          getState: () => state,
          publish: (commit: CanonicalCommit) => {
            const published = publishCanonicalProjection({
              state,
              commit,
              session,
            });
            if (published.isErr()) return false;
            state = published.value.state;
            return true;
          },
        });
        const ids = () =>
          session.document.package.document.content.flatMap((block) =>
            block.type === "paragraph" ? [block.paraId] : [],
          );
        const before = ids();
        const apply = () =>
          executor.apply({
            snapshot: createFolioAIEditSnapshot(state.doc),
            batch: { version: 1, mode, operations: [operation] },
            undefinedReferences: "refuse",
          });
        const first = apply();
        expect(first.skipped).toEqual([]);
        const firstIds = ids().filter((id) => !before.includes(id));
        expect(firstIds).toHaveLength(1);
        expect(
          executor.undo(first.undoHandle ?? panic("Missing undo handle"))
            .status,
        ).toBe("undone");
        expect(ids()).toEqual(before);
        expect(apply().skipped).toEqual([]);
        const secondIds = ids().filter((id) => !before.includes(id));
        expect(secondIds).toHaveLength(1);
        expect(secondIds).not.toEqual(firstIds);
      });
});

describe("an insertion beside the paragraph the batch deletes", () => {
  // Both apply, as in the legacy engine, in either order, as a move when
  // both halves name one.
  for (const type of ["insertAfterBlock", "insertBeforeBlock"] as const)
    for (const moveId of [undefined, "move"] as const)
      for (const reversed of [false, true])
        for (const [shape, body] of Object.entries(BODIES))
          for (const mode of ["direct", "tracked-changes"] as const)
            test(`${type}${moveId === undefined ? "" : " as a move"}${reversed ? " listed first" : ""} ${shape}, ${mode}`, async () => {
              const source = await packageOf(body);
              const halves = [
                {
                  id: "from",
                  type: "deleteBlock",
                  blockId: "10000002",
                  ...(moveId === undefined ? {} : { moveId }),
                },
                {
                  id: "to",
                  type,
                  blockId: "10000002",
                  text: "Moved",
                  ...(moveId === undefined ? {} : { moveId }),
                },
              ] as const satisfies readonly FolioDocumentOperation[];
              const operations = reversed ? [...halves].reverse() : halves;
              const run = await canonicalRun(source, operations, mode);
              expect(run.skipped).toEqual([]);
              const actual = await outcomes(
                await viaCanonicalEngine(source, operations, mode),
              );
              const expected = await outcomes(
                await viaDefaultEngine(source, operations, mode),
              );
              // Tracked, a move of the paragraph that ends the body beside
              // itself: the legacy engine keeps both paragraphs pending,
              // each with a mark, the last one on the body's final paragraph,
              // and leaves the moved paragraph behind, empty, once accepted.
              // Here no mark stands on the final paragraph: the moved words
              // and the words moved there share one paragraph. Accepting
              // gives the paragraphs the batch gives directly and rejecting
              // gives the source's back; which of the two paragraph ids the
              // one paragraph keeps differs.
              if (
                moveId !== undefined &&
                shape === "ending the body" &&
                mode === "tracked-changes"
              ) {
                expect(actual.lists).toEqual(expected.lists);
                expect(actual.comments).toEqual(expected.comments);
                expect(withoutIds(actual.accepted)).toEqual(
                  withoutIds(
                    await blocksOf(
                      await viaDefaultEngine(source, operations, "direct"),
                    ),
                  ),
                );
                expect(withoutIds(actual.rejected)).toEqual(
                  withoutIds(await blocksOf(source)),
                );
                expect(
                  withoutIds(
                    await blocksOf(
                      await viaCanonicalEngine(
                        source,
                        operations,
                        "suggested",
                        "reject",
                      ),
                    ),
                  ),
                ).toEqual(withoutIds(await blocksOf(source)));
                return;
              }
              // Tracked after a paragraph a block content control follows,
              // the legacy engine keeps the deleted paragraph apart and
              // leaves it behind, empty, once everything is accepted. Here
              // its struck words join the new paragraph that holds its mark,
              // as at the end of the body: accepting gives what the batch
              // gives directly, and rejecting gives the source's paragraphs
              // back, the anchor's words under the new paragraph's id.
              if (
                type === "insertAfterBlock" &&
                shape === "before a block content control" &&
                mode === "tracked-changes"
              ) {
                expect(actual.lists).toEqual(expected.lists);
                expect(actual.comments).toEqual(expected.comments);
                expect(actual.accepted).toEqual(
                  await blocksOf(
                    await viaDefaultEngine(source, operations, "direct"),
                  ),
                );
                expect(withoutIds(actual.rejected)).toEqual(
                  withoutIds(await blocksOf(source)),
                );
                expect(
                  withoutIds(
                    await blocksOf(
                      await viaCanonicalEngine(
                        source,
                        operations,
                        "suggested",
                        "reject",
                      ),
                    ),
                  ),
                ).toEqual(withoutIds(await blocksOf(source)));
                return;
              }
              expect(actual).toEqual(expected);
            });
});

describe("paragraphs inserted beside the paragraph the batch deletes", () => {
  // Each line its own paragraph, beside a paragraph that ends the body, as
  // the legacy engine places and records them.
  for (const type of ["insertAfterBlock", "insertBeforeBlock"] as const)
    for (const reversed of [false, true])
      for (const mode of ["direct", "tracked-changes"] as const)
        test(`${type}${reversed ? " listed first" : ""}, ${mode}`, async () => {
          const source = await packageOf(BODIES["ending the body"]);
          const halves = [
            { id: "from", type: "deleteBlock", blockId: "10000002" },
            { id: "to", type, blockId: "10000002", text: "Moved\nAgain" },
          ] as const satisfies readonly FolioDocumentOperation[];
          const operations = reversed ? [...halves].reverse() : halves;
          const run = await canonicalRun(source, operations, mode);
          expect(run.skipped).toEqual([]);
          expect(
            await outcomes(await viaCanonicalEngine(source, operations, mode)),
          ).toEqual(
            await outcomes(await viaDefaultEngine(source, operations, mode)),
          );
        });
});

describe("a deletion beside an insertion the batch refuses", () => {
  // The deletion resolves on its own, as it does without the insertion.
  for (const type of ["insertAfterBlock", "insertBeforeBlock"] as const)
    for (const mode of ["direct"] as const)
      test(`${type} refused, ${mode}`, async () => {
        const source = await packageOf(BODIES["ending the body"]);
        const operations = [
          { id: "from", type: "deleteBlock", blockId: "10000002" },
          {
            id: "to",
            type,
            blockId: "10000002",
            text: "Moved",
            precondition: { blockTextHash: "h0stale" },
          },
        ] as const satisfies readonly FolioDocumentOperation[];
        const run = await canonicalRun(source, operations, mode);
        expect(run.skipped.map(({ id, reason }) => [id, reason])).toEqual([
          ["to", "preconditionFailed"],
        ]);
        const alone = await canonicalRun(source, [operations[0]], mode);
        expect(alone.skipped).toEqual([]);
        expect(run.state().doc.toJSON()).toEqual(alone.state().doc.toJSON());
      });
});

describe("a move whose source a content lock keeps", () => {
  // The source half is refused; the destination applies as a plain
  // insertion, reported unpaired, never as half a move.
  for (const mode of ["direct", "tracked-changes"] as const)
    test(`inserts the destination plainly, ${mode}`, async () => {
      const source = await packageOf([
        `<w:sdt><w:sdtPr><w:lock w:val="contentLocked"/></w:sdtPr><w:sdtContent>${paragraph("10000001", "Alpha")}</w:sdtContent></w:sdt>`,
        paragraph("10000002", "Beta"),
        paragraph("10000003", "Gamma"),
      ]);
      const document = await parseDocx(source.slice(0), {
        preloadFonts: false,
        detectVariables: false,
      });
      const session = createCanonicalSession(
        document,
        document.package.styles,
      ).unwrap();
      let state = EditorState.create({ schema, doc: session.projection.doc });
      const executor = new CanonicalPublicOperations({
        session,
        getState: () => state,
        publish: (commit: CanonicalCommit) => {
          const published = publishCanonicalProjection({
            state,
            commit,
            session,
          });
          if (published.isErr()) return false;
          state = published.value.state;
          return true;
        },
      });
      const result = executor.apply({
        snapshot: createFolioAIEditSnapshot(state.doc),
        batch: {
          version: 1,
          mode,
          operations: [
            {
              id: "from",
              type: "deleteBlock",
              blockId: "10000001",
              moveId: "move",
            },
            {
              id: "to",
              type: "insertAfterBlock",
              blockId: "10000003",
              text: "Alpha",
              moveId: "move",
            },
          ],
        },
        undefinedReferences: "refuse",
      });
      expect(result.applied.map(({ id }) => id)).toEqual(["to"]);
      expect(result.skipped.map(({ id }) => id)).toEqual(["from"]);
      expect(result.normalizations).toContainEqual({
        id: "to",
        code: "unpairedMove",
        moveId: "move",
      });
      const saved = (
        await serializeCanonicalSave({
          snapshot: session.captureSaveSnapshot(),
        })
      ).buffer;
      const xml =
        (await (
          await ZipArchive.read(saved)
        )
          .entry("word/document.xml")
          ?.text()) ?? "";
      expect(xml).not.toContain("w:moveTo");
      expect(xml).not.toContain("w:moveFrom");
      expect(xml).toContain("Alpha");
    });
});

describe("a block move matches the legacy engine", () => {
  /** The source paragraph moved after the anchor, both halves naming one move. */
  const move = (destination: "insertAfterBlock" | "insertBeforeBlock") =>
    [
      {
        id: "from",
        type: "deleteBlock",
        blockId: "10000001",
        moveId: "move",
      },
      {
        id: "to",
        type: destination,
        blockId: "10000002",
        text: "Moved",
        moveId: "move",
      },
    ] as const satisfies readonly FolioDocumentOperation[];

  /** The package with only one half of the move accepted or rejected. */
  const halves = async (bytes: ArrayBuffer) => {
    const resolvedHalf = async (
      kind: "insertion" | "deletion",
      decision: "accept" | "reject",
    ) => {
      const reviewer = await FolioDocxReviewer.fromBuffer(bytes.slice(0));
      for (const change of reviewer.getChanges())
        if (change.type === kind)
          if (decision === "accept") reviewer.acceptChange(change);
          else reviewer.rejectChange(change);
      return blocksOf(await reviewer.toBuffer());
    };
    return {
      sourceAccepted: await resolvedHalf("deletion", "accept"),
      sourceRejected: await resolvedHalf("deletion", "reject"),
      destinationAccepted: await resolvedHalf("insertion", "accept"),
      destinationRejected: await resolvedHalf("insertion", "reject"),
    };
  };

  for (const destination of ["insertAfterBlock", "insertBeforeBlock"] as const)
    for (const [shape, body] of Object.entries(BODIES))
      for (const mode of ["direct", "tracked-changes"] as const)
        test(`${destination} ${shape}, ${mode}`, async () => {
          const source = await packageOf(body);
          const byDefault = await viaDefaultEngine(
            source,
            move(destination),
            mode,
          );
          const byCanonical = await viaCanonicalEngine(
            source,
            move(destination),
            mode,
          );
          await expectOutcomes(
            await outcomes(byCanonical),
            await outcomes(byDefault),
            source,
            shape,
          );
          if (mode === "direct") return;
          expect(await halves(byCanonical)).toEqual(await halves(byDefault));
          // The saved file marks the halves as a move.
          const xml =
            (await (
              await ZipArchive.read(byCanonical)
            )
              .entry("word/document.xml")
              ?.text()) ?? "";
          expect(xml).toContain("<w:moveFrom ");
          expect(xml).toContain("<w:moveTo ");
        });
});

describe("a comment on an edit in a table cell matches the legacy engine", () => {
  const operation = {
    id: "op",
    type: "replaceInBlock",
    blockId: "10000004",
    find: "Cell",
    replace: "Box",
    comment: { text: "Inside the table" },
  } as const satisfies FolioDocumentOperation;
  for (const mode of ["direct", "tracked-changes"] as const)
    test(mode, async () => {
      const source = await packageOf(BODIES["before a table"]);
      const expected = await outcomes(
        await viaDefaultEngine(source, operation, mode),
      );
      expect(expected.comments.pending).toHaveLength(1);
      const actual = await outcomes(
        await viaCanonicalEngine(source, operation, mode),
      );
      expect(actual).toEqual(expected);
    });
});

describe("a commented insertion inside a block content control matches the legacy engine", () => {
  for (const type of ["insertAfterBlock", "insertBeforeBlock"] as const)
    for (const mode of ["direct", "tracked-changes"] as const)
      test(`${type}, ${mode}`, async () => {
        const source = await packageOf([
          paragraph("10000001", "Alpha"),
          `<w:sdt><w:sdtPr><w:id w:val="7"/></w:sdtPr><w:sdtContent>${paragraph("10000002", "Beta")}</w:sdtContent></w:sdt>`,
          paragraph("10000003", "Gamma"),
        ]);
        const operation = {
          id: "op",
          type,
          blockId: "10000002",
          text: "Inserted",
          comment: { text: "Why this clause" },
        } as const satisfies FolioDocumentOperation;
        const run = await canonicalRun(source, operation, mode);
        expect(run.skipped).toEqual([]);
        const expected = await outcomes(
          await viaDefaultEngine(source, operation, mode),
        );
        expect(expected.comments.pending).toEqual([
          { anchoredText: "Inserted", text: "Why this clause" },
        ]);
        expect(
          await outcomes(await viaCanonicalEngine(source, operation, mode)),
        ).toEqual(expected);
      });
});

describe("an insertion before a block content control", () => {
  // Rejected, the inserted paragraph goes with its text, as in the default
  // engine: with no paragraph after the new one, the anchor's mark is the
  // inserted one, so rejecting it joins the two.
  const body = [
    paragraph("10000001", "Alpha"),
    paragraph("10000002", "Beta", '<w:ind w:left="720"/>'),
    `<w:sdt><w:sdtPr><w:id w:val="5"/></w:sdtPr><w:sdtContent>${paragraph("10000004", "Inside")}</w:sdtContent></w:sdt>`,
    paragraph("10000003", "Gamma"),
  ];
  for (const [name, operation] of [
    [
      "one paragraph",
      { id: "op", type: "insertAfterBlock", blockId: "10000002", text: "New" },
    ],
    [
      "several paragraphs",
      {
        id: "op",
        type: "insertAfterBlock",
        blockId: "10000002",
        text: "New\nMore",
      },
    ],
    [
      "a blank paragraph",
      { id: "op", type: "insertAfterBlock", blockId: "10000002", text: "" },
    ],
  ] as const satisfies readonly (readonly [string, FolioDocumentOperation])[])
    test(name, async () => {
      const source = await packageOf(body);
      // Read back by the legacy engine, rejecting the canonical file joins
      // the anchor into the new paragraph under the new paragraph's id, as at
      // the end of a container; everything else is the same.
      const actual = await outcomes(
        await viaCanonicalEngine(source, operation, "tracked-changes"),
      );
      const expected = await outcomes(
        await viaDefaultEngine(source, operation, "tracked-changes"),
      );
      expect({ ...actual, rejected: undefined }).toEqual({
        ...expected,
        rejected: undefined,
      });
      expect(withoutIds(actual.rejected)).toEqual(
        withoutIds(await blocksOf(source)),
      );
      expect(withoutIds(expected.rejected)).toEqual(
        withoutIds(await blocksOf(source)),
      );
      // Rejected in the canonical session itself, the source comes back,
      // identities included.
      expect(
        await blocksOf(
          await viaCanonicalEngine(source, operation, "suggested", "reject"),
        ),
      ).toEqual(await blocksOf(source));
      // While the suggestion is pending, the file holds the source.
      const pending = await canonicalRun(source, operation, "suggested");
      expect(
        await blocksOf(
          (
            await serializeCanonicalSave({
              snapshot: pending.session.captureSaveSnapshot(),
            })
          ).buffer,
        ),
      ).toEqual(await blocksOf(source));
    });
});

describe("a save while a suggestion is pending", () => {
  // The file is written as if the suggestion were rejected: the paragraphs it
  // adds are no change to the file, so the save splices only what changed.
  // Where the suggestion touched no source paragraph, the body stays as the
  // source has it, byte for byte.
  for (const [name, operation, untouched] of [
    [
      "insertion",
      { id: "op", type: "insertAfterBlock", blockId: "10000002", text: "New" },
      true,
    ],
    [
      "commented insertion",
      {
        id: "op",
        type: "insertAfterBlock",
        blockId: "10000002",
        text: "New",
        comment: { text: "Why this line" },
      },
      true,
    ],
    [
      "split",
      { id: "op", type: "splitBlock", blockId: "10000002", offset: 2 },
      false,
    ],
  ] as const satisfies readonly (readonly [
    string,
    FolioDocumentOperation,
    boolean,
  ])[])
    test(name, async () => {
      const source = await packageOf(BODIES["between two paragraphs"]);
      const run = await canonicalRun(source, operation, "suggested");
      expect(run.skipped).toEqual([]);
      const saved = await serializeCanonicalSave({
        snapshot: run.session.captureSaveSnapshot(),
        featureFlags: { selectiveSave: true },
      });
      expect(saved.diagnostics).toEqual([]);
      expect(await blocksOf(saved.buffer)).toEqual(await blocksOf(source));
      if (!untouched) return;
      const body = async (bytes: ArrayBuffer) =>
        (await ZipArchive.read(bytes)).entry("word/document.xml")?.text();
      expect(await body(saved.buffer)).toBe(await body(source));
    });
});
