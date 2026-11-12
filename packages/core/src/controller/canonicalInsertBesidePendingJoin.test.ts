/**
 * A paragraph inserted beside one whose break a pending join removes, in the
 * canonical session: a batch reads the document as it stands before the
 * batch, so a paragraph put after one that reads as part of the next goes
 * after the paragraph it joins, as the batch read it, even when the same
 * batch splits that paragraph, and whether the joining paragraph keeps its
 * words or was emptied. One put before it keeps its place. Either way the
 * outcome is the one the same batch anchored on the joined paragraph gives,
 * and rejecting gives the document back.
 */

import { describe, expect, test } from "bun:test";
import { EditorState } from "prosemirror-state";

import { ZipArchive } from "@stll/docx-core/zip";

import { FolioDocxReviewer } from "../ai-edits/headless";
import { getTrackedChangeStatsFromDoc } from "../ai-edits/read";
import { createFolioAIEditSnapshot } from "../ai-edits/snapshot";
import type {
  FolioDocumentOperation,
  FolioDocumentOperationMode,
} from "../document-operations";
import { ensureParaIds } from "../docx/ensureParaIds";
import { parseDocx } from "../docx/parser";
import { createDocx, createEmptyDocx } from "../docx/rezip";
import { fromMarkdown } from "../markdown/fromMarkdown";
import { schema } from "../prosemirror/schema";
import { CanonicalPublicOperations } from "./canonicalPublicOperations";
import {
  type CanonicalCommit,
  createCanonicalSession,
  publishCanonicalProjection,
} from "./canonicalSession";

const CLAUSE = "The Buyer pays each invoice within thirty days.";
const SURVIVOR = "Signed in two copies.";
const INSERTED = "Inserted clause.";

/**
 * How the clause comes to join the paragraph after it: a tracked merge
 * leaves its words, and a tracked deletion of which only the words are
 * accepted leaves it empty with its break pending deletion.
 */
const JOIN_SOURCES = ["a tracked merge", "an emptied deletion"] as const;
type JoinSource = (typeof JOIN_SOURCES)[number];

type Fixture = { bytes: ArrayBuffer; clause: string; survivor: string };

/** The clause joining the paragraph after it as `source` makes it, saved. */
const fixture = async (
  source: JoinSource,
  survivorMarkdown = SURVIVOR,
): Promise<Fixture> => {
  const markdown = `Alpha.\n\n${CLAUSE}\n\n${survivorMarkdown}\n\nCharlie.`;
  const reviewer = await FolioDocxReviewer.fromBuffer(
    (await ensureParaIds(await createDocx(fromMarkdown(markdown)))).docx,
    { author: "Test" },
  );
  const [, clause, survivor] = reviewer.snapshot().blocks;
  if (clause?.text !== CLAUSE || survivor?.text !== SURVIVOR)
    throw new Error("Missing fixture");
  const operation: FolioDocumentOperation =
    source === "a tracked merge"
      ? {
          id: "merge",
          type: "mergeBlockWithNext",
          blockId: clause.id,
          separator: " ",
        }
      : { id: "delete", type: "deleteBlock", blockId: clause.id };
  expect(
    reviewer.applyOperations([operation], { mode: "tracked-changes" }).skipped,
  ).toEqual([]);
  if (source === "an emptied deletion") {
    const words = reviewer
      .getChanges()
      .find((change) => change.type === "deletion" && change.text === CLAUSE);
    if (!words) throw new Error("Missing the clause's deleted words");
    expect(reviewer.acceptChange(words)).toBe(true);
  }
  return {
    bytes: await reviewer.toBuffer(),
    clause: clause.id,
    survivor: survivor.id,
  };
};

type Block = { id: string; text: string; styleId: string | undefined };

/** A canonical session over `bytes`, its public operations and its blocks. */
const open = async (bytes: ArrayBuffer) => {
  const document = await parseDocx(bytes.slice(0), {
    preloadFonts: false,
    detectVariables: false,
  });
  const session = createCanonicalSession(
    document,
    document.package.styles,
  ).unwrap();
  let state = EditorState.create({ schema, doc: session.projection.doc });
  const publish = (commit: CanonicalCommit) => {
    const published = publishCanonicalProjection({ state, commit, session });
    if (published.isErr()) return false;
    state = published.value.state;
    return true;
  };
  const executor = new CanonicalPublicOperations({
    session,
    getState: () => state,
    publish,
  });
  return {
    apply: (
      operations: readonly FolioDocumentOperation[],
      mode: FolioDocumentOperationMode,
    ) =>
      executor.apply({
        snapshot: createFolioAIEditSnapshot(state.doc),
        batch: { version: 1, mode, operations: [...operations] },
      }),
    resolveAll: (resolution: "accept" | "reject") => {
      const { ids } = getTrackedChangeStatsFromDoc(session.projection.doc);
      if (ids.length === 0) return;
      expect(
        publish(
          session
            .prepareResolve(state, { revisionIds: ids, resolution })
            .unwrap(),
        ),
      ).toBe(true);
      expect(getTrackedChangeStatsFromDoc(session.projection.doc).ids).toEqual(
        [],
      );
    },
    blocks: (): Block[] =>
      createFolioAIEditSnapshot(session.projection.doc).blocks.map(
        ({ id, text, styleId }) => ({ id, text, styleId }),
      ),
  };
};

/** The texts every change rejected gives: the document before any edit. */
const original = async (bytes: ArrayBuffer) => {
  const editor = await open(bytes);
  editor.resolveAll("reject");
  return editor.blocks().map(({ text }) => text);
};

const insertion = (
  type: "insertAfterBlock" | "insertBeforeBlock",
  blockId: string,
): FolioDocumentOperation => ({ id: "insert", type, blockId, text: INSERTED });

/** Texts and styles, without the ids the batch's new paragraphs take. */
const shapeOf = (blocks: readonly Block[]) =>
  blocks.map(({ text, styleId }) => ({ text, styleId }));

describe("a paragraph beside one that joins the next, in the canonical session", () => {
  for (const source of JOIN_SOURCES)
    for (const mode of ["tracked-changes", "direct"] as const)
      for (const type of ["insertAfterBlock", "insertBeforeBlock"] as const) {
        const name = `(${source}, ${type}, ${mode})`;
        /** The insertion anchored on `anchor`, pending and accepted. */
        const read = async (anchor: "clause" | "survivor") => {
          // The paragraph the join leaves has properties of its own, which
          // the joined paragraph, and so the inserted one, takes.
          const joined = await fixture(source, `## ${SURVIVOR}`);
          const editor = await open(joined.bytes);
          const result = editor.apply([insertion(type, joined[anchor])], mode);
          expect(result.skipped).toEqual([]);
          const pending = editor.blocks();
          editor.resolveAll("accept");
          return { pending, accepted: editor.blocks() };
        };

        test(`is placed as one beside the paragraph it joins ${name}`, async () => {
          const fromClause = await read("clause");
          const fromSurvivor = await read("survivor");
          expect(shapeOf(fromClause.pending)).toEqual(
            shapeOf(fromSurvivor.pending),
          );
          expect(shapeOf(fromClause.accepted)).toEqual(
            shapeOf(fromSurvivor.accepted),
          );
          const texts = fromClause.accepted.map(({ text }) => text);
          const joined = texts.findIndex((text) => text.endsWith(SURVIVOR));
          const inserted = texts.indexOf(INSERTED);
          expect(joined).toBeGreaterThan(-1);
          expect(inserted).toBe(
            type === "insertAfterBlock" ? joined + 1 : joined - 1,
          );
          expect(fromClause.accepted[inserted]?.styleId).toBe(
            fromClause.accepted[joined]?.styleId,
          );
        });
      }

  for (const source of JOIN_SOURCES)
    for (const type of ["insertAfterBlock", "insertBeforeBlock"] as const)
      test(`rejects to the document it started from (${source}, ${type})`, async () => {
        const joined = await fixture(source);
        const editor = await open(joined.bytes);
        expect(
          editor.apply([insertion(type, joined.clause)], "tracked-changes")
            .skipped,
        ).toEqual([]);
        editor.resolveAll("reject");
        expect(editor.blocks().map(({ text }) => text)).toEqual(
          await original(joined.bytes),
        );
      });

  for (const source of JOIN_SOURCES)
    test(`as a suggestion, goes beside the paragraph it names as it stands (${source})`, async () => {
      const joined = await fixture(source);
      const editor = await open(joined.bytes);
      expect(
        editor.apply(
          [insertion("insertAfterBlock", joined.clause)],
          "suggested",
        ).skipped,
      ).toEqual([]);
      const ids = editor.blocks().map(({ id }) => id);
      expect(editor.blocks().at(ids.indexOf(joined.clause) + 1)?.text).toBe(
        INSERTED,
      );
      editor.resolveAll("reject");
      expect(editor.blocks().map(({ text }) => text)).toEqual(
        await original(joined.bytes),
      );
    });
});

/** A join whose paragraphs a block may name: each anchor reads as part of the last. */
type LongerJoin = {
  bytes: ArrayBuffer;
  anchors: Readonly<Record<string, string>> & { survivor: string };
};

/**
 * The clause and an emptied paragraph after it, both with their breaks
 * pending deletion by tracked merges, joining "Signed in two copies.".
 */
const throughAnEmptiedParagraph = async (): Promise<LongerJoin> => {
  const markdown = `Alpha.\n\n${CLAUSE}\n\n${SURVIVOR}\n\nCharlie.`;
  const reviewer = await FolioDocxReviewer.fromBuffer(
    (await ensureParaIds(await createDocx(fromMarkdown(markdown)))).docx,
    { author: "Test" },
  );
  const clause = reviewer.snapshot().blocks.at(1);
  if (clause?.text !== CLAUSE) throw new Error("Missing fixture");
  expect(
    reviewer.applyOperations(
      [{ id: "empty", type: "insertAfterBlock", blockId: clause.id, text: "" }],
      { mode: "direct" },
    ).skipped,
  ).toEqual([]);
  const [, , emptied, survivor] = reviewer.snapshot().blocks;
  if (emptied?.text !== "" || survivor?.text !== SURVIVOR)
    throw new Error("Missing fixture");
  for (const operation of [
    {
      id: "merge-clause",
      type: "mergeBlockWithNext",
      blockId: clause.id,
      separator: " ",
    },
    { id: "merge-emptied", type: "mergeBlockWithNext", blockId: emptied.id },
  ] satisfies FolioDocumentOperation[])
    expect(
      reviewer.applyOperations([operation], { mode: "tracked-changes" })
        .skipped,
    ).toEqual([]);
  return {
    bytes: await reviewer.toBuffer(),
    anchors: {
      clause: clause.id,
      emptied: emptied.id,
      survivor: survivor.id,
    },
  };
};

const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const W14 = "http://schemas.microsoft.com/office/word/2010/wordml";

/**
 * The clause, its break pending deletion, joining "Signed in two copies."
 * across the bookmark that starts between them.
 */
const acrossARangeBoundary = async (): Promise<LongerJoin> => {
  const run = (text: string) =>
    `<w:r><w:t xml:space="preserve">${text}</w:t></w:r>`;
  const zip = await ZipArchive.read(await createEmptyDocx());
  zip.set(
    "word/document.xml",
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document xmlns:w="${W}" xmlns:w14="${W14}"><w:body>` +
      `<w:p w14:paraId="10000001">${run("Alpha.")}</w:p>` +
      `<w:p w14:paraId="10000002"><w:pPr><w:rPr><w:del w:id="1" w:author="Test" w:date="2026-01-01T00:00:00Z"/></w:rPr></w:pPr>${run(CLAUSE)}</w:p>` +
      `<w:bookmarkStart w:id="7" w:name="Signature"/>` +
      `<w:p w14:paraId="10000003">${run(SURVIVOR)}</w:p>` +
      `<w:bookmarkEnd w:id="7"/>` +
      `<w:p w14:paraId="10000004">${run("Charlie.")}</w:p>` +
      `<w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr></w:body></w:document>`,
  );
  return {
    bytes: (await zip.generate({ compression: "store" })).buffer,
    anchors: { clause: "10000002", survivor: "10000003" },
  };
};

const LONGER_JOINS = {
  "through an emptied paragraph": throughAnEmptiedParagraph,
  "across a range boundary": acrossARangeBoundary,
};

describe("a paragraph beside a longer pending join, in the canonical session", () => {
  for (const [shape, make] of Object.entries(LONGER_JOINS))
    for (const side of ["after", "before"] as const) {
      const type = side === "after" ? "insertAfterBlock" : "insertBeforeBlock";
      const name = `(${shape}, ${side})`;

      test(`is placed as one beside the paragraph the join leaves, whichever paragraph of the join it names ${name}`, async () => {
        const join = await make();
        const read = async (anchor: string) => {
          const editor = await open(join.bytes);
          expect(
            editor.apply([insertion(type, anchor)], "tracked-changes").skipped,
          ).toEqual([]);
          const pending = shapeOf(editor.blocks());
          editor.resolveAll("accept");
          return { pending, accepted: editor.blocks().map(({ text }) => text) };
        };
        const expected = await read(join.anchors.survivor);
        const joined = expected.accepted.findIndex((text) =>
          text.endsWith(SURVIVOR),
        );
        expect(joined).toBeGreaterThan(-1);
        expect(expected.accepted.indexOf(INSERTED)).toBe(
          side === "after" ? joined + 1 : joined - 1,
        );
        for (const anchor of Object.values(join.anchors))
          expect(await read(anchor)).toEqual(expected);
      });

      test(`rejects to the document it started from, whichever paragraph it names ${name}`, async () => {
        const join = await make();
        for (const anchor of Object.values(join.anchors)) {
          const editor = await open(join.bytes);
          expect(
            editor.apply([insertion(type, anchor)], "tracked-changes").skipped,
          ).toEqual([]);
          editor.resolveAll("reject");
          expect(editor.blocks().map(({ text }) => text)).toEqual(
            await original(join.bytes),
          );
        }
      });
    }
});

type SplitBatchOptions = {
  source: JoinSource;
  /** Whether the paragraph is put after the joining clause or after the paragraph it joins. */
  anchor: "clause" | "survivor";
  offset: number;
  order: "insertion first" | "split first";
  resolution: "accept" | "reject";
};

/** One batch: a paragraph put after `anchor`, and the joined paragraph split at `offset`. */
const splitBatch = async ({
  source,
  anchor,
  offset,
  order,
  resolution,
}: SplitBatchOptions) => {
  const joined = await fixture(source);
  const editor = await open(joined.bytes);
  const insert = insertion("insertAfterBlock", joined[anchor]);
  const split: FolioDocumentOperation = {
    id: "split",
    type: "splitBlock",
    blockId: joined.survivor,
    offset,
  };
  expect(
    editor.apply(
      order === "insertion first" ? [insert, split] : [split, insert],
      "tracked-changes",
    ).skipped,
  ).toEqual([]);
  editor.resolveAll(resolution);
  return {
    texts: editor.blocks().map(({ text }) => text),
    original: await original(joined.bytes),
  };
};

describe("a paragraph after one that joins a paragraph the batch splits, in the canonical session", () => {
  for (const source of JOIN_SOURCES)
    for (const offset of [1, 14, SURVIVOR.length - 1])
      for (const order of ["insertion first", "split first"] as const) {
        const options = { source, offset, order };
        const name = `(${source}, split at ${offset}, ${order})`;

        test(`follows the whole joined paragraph, as one put after it does ${name}`, async () => {
          const { texts: accepted } = await splitBatch({
            ...options,
            anchor: "clause",
            resolution: "accept",
          });
          const inserted = accepted.indexOf(INSERTED);
          expect(inserted).toBeGreaterThan(0);
          expect(accepted.at(inserted - 1)).toBe(SURVIVOR.slice(offset));
          expect(accepted).toEqual(
            (
              await splitBatch({
                ...options,
                anchor: "survivor",
                resolution: "accept",
              })
            ).texts,
          );
        });

        test(`rejects to the document it started from ${name}`, async () => {
          const rejected = await splitBatch({
            ...options,
            anchor: "clause",
            resolution: "reject",
          });
          expect(rejected.texts).toEqual(rejected.original);
        });
      }
});
