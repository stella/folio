/**
 * Block-level public operations in the canonical session give the document
 * the default engine gives: the same blocks and paragraph properties pending,
 * once every change is accepted, and once every change is rejected.
 */

import { describe, expect, test } from "bun:test";
import { EditorState } from "prosemirror-state";

import { paragraphNumberingReference } from "@stll/docx-core/model";
import { ZipArchive } from "@stll/docx-core/zip";

import type { FolioDocumentOperation } from "../ai-edits";
import { SUGGESTED_SUPPORTED_OPERATION_TYPES } from "../ai-edits/apply";
import { FolioDocxReviewer } from "../ai-edits/headless";
import { createFolioAIEditSnapshot } from "../ai-edits/snapshot";
import { serializeCanonicalSave } from "../docx/canonicalSave";
import { parseDocx } from "../docx/parser";
import { schema } from "../prosemirror/schema";
import { CanonicalPublicOperations } from "./canonicalPublicOperations";
import {
  type CanonicalCommit,
  createCanonicalSession,
  publishCanonicalProjection,
} from "./canonicalSession";

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

/**
 * The canonical session's outcomes against the default engine's. Rejecting a
 * tracked insertion after a container's final paragraph gives the source back
 * from the canonical session's file, ids included: it records which paragraph
 * rejecting the mark keeps. The default engine's file does not, so there the
 * anchor's words come back under the new paragraph's id. Everything else is
 * the same.
 */
const expectOutcomes = async (
  actual: Awaited<ReturnType<typeof outcomes>>,
  expected: Awaited<ReturnType<typeof outcomes>>,
  source: ArrayBuffer,
) => {
  const withoutIds = (blocks: readonly object[]) =>
    blocks.map((block) => ({ ...block, id: undefined }));
  if (
    !Bun.deepEquals(actual.rejected, expected.rejected) &&
    Bun.deepEquals(actual.rejected, await blocksOf(source)) &&
    Bun.deepEquals(withoutIds(actual.rejected), withoutIds(expected.rejected))
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
/** The default engine's package after a suggested batch it then resolves. */
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

describe("block-level public operations match the default engine", () => {
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
          await expectOutcomes(actual, expected, source);
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
          );
          // The default engine's suggestion resolves the same way.
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

describe("a block move matches the default engine", () => {
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

describe("a comment on an edit in a table cell matches the default engine", () => {
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

describe("a commented insertion inside a block content control matches the default engine", () => {
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
