/**
 * A public operation that joins two paragraphs, applied directly, leaves the
 * document that suggesting it and accepting the suggestion leaves, block ids
 * included: the joined paragraph is the one whose mark remains, the later one.
 *
 * Each case merges a paragraph with the next one or deletes one, in either
 * engine, in the body, a table cell, a header or a note, the final paragraph
 * of its container and one that ends its own section among them. Applied
 * directly and applied tracked then accepted, every story reads the same
 * blocks under the same ids, and the direct result names only blocks that
 * exist as updated and only ids that name no block as deleted.
 */

import { describe, expect, mock, test } from "bun:test";
import fc from "fast-check";

import { ZipArchive } from "@stll/docx-core/zip";

import { checkCoverageSample } from "../../../../test/property-coverage";
import {
  assertProperty,
  propertyConfig,
  propertyTestTimeout,
} from "../../../../test/property-testing";
import {
  FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
  type FolioDocumentOperation,
  type FolioDocumentOperationResult,
} from "../document-operations";
import { getParagraphPropertySourceToken } from "../docx/paragraphPropertySource";
import type { HeadlessReviewerSession } from "../types/canonicalCapabilities";
import * as directJoin from "./directJoin";
import { type FolioDocumentStoryHandle, FolioDocxReviewer } from "./headless";

const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const HEAD = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';
const MAIN = "application/vnd.openxmlformats-officedocument.wordprocessingml";

const paragraphXml = (text: string, properties = "", before = ""): string =>
  `<w:p>${properties === "" ? "" : `<w:pPr>${properties}</w:pPr>`}${before}<w:r><w:t xml:space="preserve">${text}</w:t></w:r></w:p>`;

const SECTION = '<w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr>';

/**
 * Two paragraphs in the body, two in a table cell, one that ends its own
 * section and the body's final paragraph; two in a header and two in a note.
 */
const BODY =
  paragraphXml("Body one", "", '<w:r><w:footnoteReference w:id="1"/></w:r>') +
  paragraphXml("Body two") +
  '<w:tbl><w:tblPr><w:tblW w:w="4000" w:type="dxa"/></w:tblPr><w:tblGrid><w:gridCol w:w="4000"/></w:tblGrid><w:tr><w:tc><w:tcPr><w:tcW w:w="4000" w:type="dxa"/></w:tcPr>' +
  paragraphXml("Cell one") +
  paragraphXml("Cell two") +
  "</w:tc></w:tr></w:tbl>" +
  paragraphXml("Before break") +
  paragraphXml("Section end", SECTION) +
  paragraphXml("Body last");

const HEADER = paragraphXml("Header one") + paragraphXml("Header two");

const NOTES =
  '<w:footnote w:type="separator" w:id="-1"><w:p><w:r><w:separator/></w:r></w:p></w:footnote>' +
  '<w:footnote w:type="continuationSeparator" w:id="0"><w:p><w:r><w:continuationSeparator/></w:r></w:p></w:footnote>' +
  '<w:footnote w:id="1"><w:p><w:r><w:footnoteRef/></w:r><w:r><w:t xml:space="preserve"> Note one</w:t></w:r></w:p>' +
  paragraphXml("Note two") +
  "</w:footnote>";

const fixture = async (body = BODY): Promise<ArrayBuffer> => {
  const zip = new ZipArchive();
  zip.set(
    "[Content_Types].xml",
    `${HEAD}<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="${MAIN}.document.main+xml"/><Override PartName="/word/header1.xml" ContentType="${MAIN}.header+xml"/><Override PartName="/word/footnotes.xml" ContentType="${MAIN}.footnotes+xml"/></Types>`,
  );
  zip.set(
    "_rels/.rels",
    `${HEAD}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${R}/officeDocument" Target="word/document.xml"/></Relationships>`,
  );
  zip.set(
    "word/_rels/document.xml.rels",
    `${HEAD}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdH" Type="${R}/header" Target="header1.xml"/><Relationship Id="rIdF" Type="${R}/footnotes" Target="footnotes.xml"/></Relationships>`,
  );
  zip.set(
    "word/header1.xml",
    `${HEAD}<w:hdr xmlns:w="${W}" xmlns:r="${R}">${HEADER}</w:hdr>`,
  );
  zip.set(
    "word/footnotes.xml",
    `${HEAD}<w:footnotes xmlns:w="${W}" xmlns:r="${R}">${NOTES}</w:footnotes>`,
  );
  zip.set(
    "word/document.xml",
    `${HEAD}<w:document xmlns:w="${W}" xmlns:r="${R}"><w:body>${body}<w:sectPr><w:headerReference w:type="default" r:id="rIdH"/></w:sectPr></w:body></w:document>`,
  );
  const bytes = await zip.generate({ compression: "store" });
  return bytes.buffer.slice(
    bytes.byteOffset,
    bytes.byteOffset + bytes.byteLength,
  ) as ArrayBuffer;
};

const open = (buffer: ArrayBuffer, session: HeadlessReviewerSession) =>
  FolioDocxReviewer.fromBuffer(buffer.slice(0), {
    author: "Reviewer",
    session,
  });

type Reviewer = Awaited<ReturnType<typeof open>>;

const MAIN_STORY: FolioDocumentStoryHandle = { type: "main" };
const HEADER_STORY: FolioDocumentStoryHandle = {
  type: "header",
  relationshipId: "rIdH",
};
const NOTE_STORY: FolioDocumentStoryHandle = { type: "footnote", noteId: 1 };
const STORIES = [MAIN_STORY, HEADER_STORY, NOTE_STORY] as const;

const TARGETS = [
  { story: MAIN_STORY, text: "Body one" },
  { story: MAIN_STORY, text: "Body two" },
  { story: MAIN_STORY, text: "Cell one" },
  { story: MAIN_STORY, text: "Cell two" },
  { story: MAIN_STORY, text: "Before break" },
  { story: MAIN_STORY, text: "Section end" },
  { story: MAIN_STORY, text: "Body last" },
  { story: HEADER_STORY, text: "Header one" },
  { story: HEADER_STORY, text: "Header two" },
  { story: NOTE_STORY, text: "Note one" },
  { story: NOTE_STORY, text: "Note two" },
] as const;

type Draw = {
  engine: HeadlessReviewerSession;
  type: "mergeBlockWithNext" | "deleteBlock";
  target: (typeof TARGETS)[number];
};

const drawArbitrary: fc.Arbitrary<Draw> = fc.record({
  engine: fc.constantFrom("legacy" as const, "canonical" as const),
  type: fc.constantFrom("mergeBlockWithNext" as const, "deleteBlock" as const),
  target: fc.constantFrom(...TARGETS),
});

const operationOf = (draw: Draw, blockId: string): FolioDocumentOperation =>
  draw.type === "deleteBlock"
    ? { id: "join", type: "deleteBlock", blockId }
    : { id: "join", type: "mergeBlockWithNext", blockId };

/** The draw applied in `mode`; `null` when the reviewer refuses it. */
const applied = (
  reviewer: Reviewer,
  draw: Draw,
  mode: "direct" | "tracked-changes",
): FolioDocumentOperationResult | null => {
  const { story, text } = draw.target;
  const block = reviewer
    .snapshotStory(story)
    ?.blocks.find((candidate) => candidate.text.includes(text));
  if (block === undefined) throw new Error(`no block reads ${text}`);
  try {
    const result = reviewer.applyDocumentOperationsToStory({
      story,
      batch: {
        version: FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
        operations: [operationOf(draw, block.id)],
        mode,
      },
      undefinedReferences: "refuse",
    });
    return result.applied.length === 1 ? result : null;
  } catch {
    return null;
  }
};

/** Every story's blocks, by id and text. */
const blocksOf = (reviewer: Reviewer) =>
  STORIES.map((story) =>
    (reviewer.snapshotStory(story)?.blocks ?? []).map(({ id, text }) => ({
      id,
      text,
    })),
  );

const law = (buffer: ArrayBuffer, tally: { applied: number }) =>
  fc.asyncProperty(drawArbitrary, async (draw) => {
    const direct = await open(buffer, draw.engine);
    const tracked = await open(buffer, draw.engine);
    const directResult = applied(direct, draw, "direct");
    if (
      directResult === null ||
      applied(tracked, draw, "tracked-changes") === null
    )
      return;
    tally.applied += 1;
    tracked.acceptAll();
    expect(blocksOf(direct)).toEqual(blocksOf(tracked));
    // The direct result names the blocks as they now stand.
    const ids = new Set(
      blocksOf(direct)
        .flat()
        .map(({ id }) => id),
    );
    for (const target of directResult.receipts.flatMap(
      ({ affected }) => affected,
    )) {
      if (target.type !== "block") continue;
      expect({ id: target.blockId, present: ids.has(target.blockId) }).toEqual({
        id: target.blockId,
        present: target.effect !== "deleted",
      });
    }
  });

const RUNS = 40;

/** Deleting a cell's final paragraph, which joins the one before it. */
const FINAL_IN_CELL: Draw = {
  engine: "canonical",
  type: "deleteBlock",
  target: { story: MAIN_STORY, text: "Cell two" },
};
const MERGE_IN_BODY: Draw = {
  engine: "legacy",
  type: "mergeBlockWithNext",
  target: { story: MAIN_STORY, text: "Body one" },
};

describe("a join applied directly", () => {
  test.each([false, true])(
    "a split and join preserve identity and source ownership (join listed first: %s)",
    async (joinFirst) => {
      const reviewer = await open(
        await fixture(
          paragraphXml("First words") + paragraphXml("Second words"),
        ),
        "legacy",
      );
      const [first, second] = reviewer.snapshotStory(MAIN_STORY)?.blocks ?? [];
      if (!first || !second) throw new Error("Missing split/join fixture.");
      const firstModel = reviewer.toDocument().package.document.content.at(0);
      if (firstModel?.type !== "paragraph")
        throw new Error("Missing source paragraph.");
      const firstToken = getParagraphPropertySourceToken(firstModel);
      expect(firstToken).toBeDefined();
      const split = {
        id: "split",
        type: "splitBlock",
        blockId: second.id,
        offset: 6,
      } as const satisfies FolioDocumentOperation;
      const merge = {
        id: "merge",
        type: "mergeBlockWithNext",
        blockId: first.id,
        separator: " ",
      } as const satisfies FolioDocumentOperation;
      const result = reviewer.applyDocumentOperations(
        {
          version: FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
          mode: "direct",
          operations: joinFirst ? [merge, split] : [split, merge],
        },
        { undefinedReferences: "refuse" },
      );
      expect(result.skipped.map(({ id, reason }) => ({ id, reason }))).toEqual(
        joinFirst ? [{ id: "split", reason: "overlappingOperation" }] : [],
      );
      const [joined, sibling] =
        reviewer.snapshotStory(MAIN_STORY)?.blocks ?? [];
      expect(joined?.text).toBe(
        joinFirst ? "First words Second words" : "First words Second",
      );
      expect(joined?.id).toBe(second.id);
      if (joinFirst) expect(sibling).toBeUndefined();
      else {
        expect(sibling?.text).toBe(" words");
        expect(sibling?.id).not.toBe(first.id);
        expect(sibling?.id).not.toBe(second.id);
      }
      const paragraphs = reviewer
        .toDocument()
        .package.document.content.filter((block) => block.type === "paragraph");
      const tokens = paragraphs.map(getParagraphPropertySourceToken);
      expect(tokens.at(0)).toBe(firstToken);
      expect(tokens.at(1)).toBeUndefined();
      const remaining = new Set(
        (reviewer.snapshotStory(MAIN_STORY)?.blocks ?? []).map(({ id }) => id),
      );
      expect(remaining.size).toBe(joinFirst ? 1 : 2);
      for (const target of result.receipts.flatMap(
        ({ affected }) => affected,
      )) {
        if (target.type === "block")
          expect(remaining.has(target.blockId)).toBe(
            target.effect !== "deleted",
          );
      }
    },
    propertyTestTimeout(30_000),
  );

  test(
    "settles chained joins and final deletions after every pre-resolved batch target runs",
    async () => {
      const buffer = await fixture(
        ["Lead", "First", "Second", "Third", "Last"]
          .map((text) => paragraphXml(text))
          .join(""),
      );
      const kinds = ["last", "third", "merge", "insert"] as const;
      await assertProperty(
        fc.asyncProperty(
          fc.shuffledSubarray(kinds, {
            minLength: kinds.length,
            maxLength: kinds.length,
          }),
          async (order) => {
            const direct = await open(buffer, "legacy");
            const tracked = await open(buffer, "legacy");
            const original = new Set(
              blocksOf(direct)
                .flat()
                .map(({ id }) => id),
            );
            const lastId = direct
              .snapshotStory(MAIN_STORY)
              ?.blocks.find(({ text }) => text === "Last")?.id;
            const operations = (reviewer: Reviewer) => {
              const ids = new Map(
                (reviewer.snapshotStory(MAIN_STORY)?.blocks ?? []).map(
                  ({ id, text }) => [text, id],
                ),
              );
              const idOf = (text: string) => {
                const id = ids.get(text);
                if (id === undefined)
                  throw new Error(`Missing fixture paragraph: ${text}`);
                return id;
              };
              const byKind = {
                last: {
                  id: "last",
                  type: "deleteBlock",
                  blockId: idOf("Last"),
                },
                third: {
                  id: "third",
                  type: "deleteBlock",
                  blockId: idOf("Third"),
                },
                merge: {
                  id: "merge",
                  type: "mergeBlockWithNext",
                  blockId: idOf("First"),
                  separator: " ",
                },
                insert: {
                  id: "insert",
                  type: "insertBeforeBlock",
                  blockId: idOf("Lead"),
                  text: "Inserted",
                },
              } satisfies Record<
                (typeof kinds)[number],
                FolioDocumentOperation
              >;
              return order.map((kind) => byKind[kind]);
            };
            const result = direct.applyDocumentOperations(
              {
                version: FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
                mode: "direct",
                operations: operations(direct),
              },
              { undefinedReferences: "refuse" },
            );
            expect(result.skipped).toEqual([]);
            expect(
              tracked.applyDocumentOperations(
                {
                  version: FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
                  mode: "tracked-changes",
                  operations: operations(tracked),
                },
                { undefinedReferences: "refuse" },
              ).skipped,
            ).toEqual([]);
            tracked.acceptAll();
            const existingIdentities = (reviewer: Reviewer) =>
              blocksOf(reviewer).map((blocks) =>
                blocks.map(({ id, text }) => ({
                  id: original.has(id) ? id : "new",
                  text,
                })),
              );
            expect(existingIdentities(direct)).toEqual(
              existingIdentities(tracked),
            );
            expect(
              direct
                .snapshotStory(MAIN_STORY)
                ?.blocks.find(({ text }) => text === "First Second")?.id,
            ).toBe(lastId);
            const remaining = new Set(
              blocksOf(direct)
                .flat()
                .map(({ id }) => id),
            );
            for (const target of result.receipts.flatMap(
              ({ affected }) => affected,
            )) {
              if (target.type === "block")
                expect(remaining.has(target.blockId)).toBe(
                  target.effect !== "deleted",
                );
            }
          },
        ),
        {
          seed: 20261116,
          numRuns: 24,
          id: "chained direct join identities settle after batch targets",
        },
      );
    },
    propertyTestTimeout(60_000),
  );

  test(
    "leaves what the accepted suggestion leaves, ids included",
    async () => {
      const buffer = await fixture();
      await assertProperty(law(buffer, { applied: 0 }), {
        id: "a direct join reads as the accepted suggestion",
        numRuns: RUNS,
        examples: [[FINAL_IN_CELL], [MERGE_IN_BODY]],
      });
      // Most draws of a fixed sample reach the law: both modes apply them.
      const tally = { applied: 0 };
      await checkCoverageSample(law(buffer, tally), RUNS);
      expect(tally.applied).toBeGreaterThan(RUNS / 4);
    },
    propertyTestTimeout(180_000),
  );

  test(
    "the law catches a direct join that keeps the earlier paragraph",
    async () => {
      // The test-only break: the direct join keeps the earlier paragraph's
      // identity, which accepting the suggestion does not.
      const original = { ...directJoin };
      mock.module("./directJoin", () => ({
        ...original,
        directJoinSurvivor: () => "first" as const,
      }));
      const buffer = await fixture();
      try {
        for (const example of [FINAL_IN_CELL, MERGE_IN_BODY]) {
          const result = await fc.check(
            law(buffer, { applied: 0 }),
            propertyConfig({
              numRuns: 1,
              examples: [[example]],
              reporter: () => undefined,
            }),
          );
          expect(result.failed).toBe(true);
        }
      } finally {
        mock.module("./directJoin", () => original);
      }
    },
    propertyTestTimeout(60_000),
  );
});
