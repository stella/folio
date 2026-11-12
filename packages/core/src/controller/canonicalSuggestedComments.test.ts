/**
 * A comment a suggestion proposes, in the canonical session: pending it is
 * listed with its suggestion and left out of the save; the host's
 * pending-suggestion export carries it into a fresh session; rejecting
 * removes it and accepting keeps it. An edit's attached comment covers the
 * text the edit changed, as in the legacy engine.
 */

import { panic } from "better-result";
import { describe, expect, test } from "bun:test";
import { EditorState } from "prosemirror-state";

import { withoutPendingSuggestions } from "@stll/docx-core/ops";

import { FolioDocxReviewer } from "../ai-edits/headless";
import {
  createPendingSuggestionSourceSnapshot,
  FolioPendingSuggestionRegistry,
} from "../ai-edits/pending-suggestions";
import {
  createFolioAIEditSnapshot,
  createFolioAITextRangeHandle,
} from "../ai-edits/snapshot";
import {
  FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
  type FolioDocumentOperation,
} from "../document-operations";
import { parseDocx } from "../docx/parser";
import { createDocx } from "../docx/rezip";
import { paragraph } from "../docx/server/build";
import { getSuggestions } from "../prosemirror/commands/comments";
import { createEmptyDocument } from "../utils/createDocument";
import { CanonicalPublicOperations } from "./canonicalPublicOperations";
import {
  type CanonicalCommit,
  createCanonicalSession,
  publishCanonicalProjection,
} from "./canonicalSession";

const FIRST_ID = "10000001";
const SECOND_ID = "10000002";
const SECOND_TEXT = "Second clause stays.";
/** The headless reviewer's default author. */
const AUTHOR = "AI";

const sourceBytes = (): Promise<ArrayBuffer> => {
  const document = createEmptyDocument();
  document.package.document.content = [
    { ...paragraph("First clause remains."), paraId: FIRST_ID },
    { ...paragraph(SECOND_TEXT), paraId: SECOND_ID },
  ];
  return createDocx(document);
};

const OPERATIONS = {
  commentOnRange: (): FolioDocumentOperation => ({
    id: "proposal",
    type: "commentOnRange",
    range:
      createFolioAITextRangeHandle({
        blockId: SECOND_ID,
        text: SECOND_TEXT,
        startOffset: 0,
        endOffset: 6,
      }) ?? panic("Missing range"),
    comment: { text: "Check this word." },
  }),
  commentOnBlock: (): FolioDocumentOperation => ({
    id: "proposal",
    type: "commentOnBlock",
    blockId: SECOND_ID,
    comment: { text: "Check this clause." },
  }),
  "replaceInBlock with a comment": (): FolioDocumentOperation => ({
    id: "proposal",
    type: "replaceInBlock",
    blockId: SECOND_ID,
    find: "Second",
    replace: "Third",
    comment: { text: "Check this edit." },
  }),
} as const;

/** A canonical session as a host drives it, with its suggestion registry. */
const canonicalHost = async (bytes: ArrayBuffer) => {
  const session = createCanonicalSession(await parseDocx(bytes)).unwrap();
  let state = EditorState.create({ doc: session.projection.doc });
  const publish = (commit: CanonicalCommit) => {
    const result = publishCanonicalProjection({ state, commit, session });
    if (result.isErr()) return false;
    state = result.value.state;
    return true;
  };
  const executor = new CanonicalPublicOperations({
    session,
    getState: () => state,
    publish,
  });
  const registry = new FolioPendingSuggestionRegistry();
  const active = () =>
    new Set(getSuggestions(state).map(({ suggestionId }) => suggestionId));
  const run = (operation: FolioDocumentOperation, author: string) =>
    executor.apply({
      snapshot: createFolioAIEditSnapshot(state.doc),
      undefinedReferences: "refuse",
      author,
      batch: {
        version: FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
        operations: [operation],
        mode: "suggested",
      },
    });
  return {
    session,
    state: () => state,
    suggest: (operation: FolioDocumentOperation) => {
      const before = active();
      const snapshot = createFolioAIEditSnapshot(state.doc);
      const result = run(operation, AUTHOR);
      expect(result.skipped).toEqual([]);
      registry.recordApplied({
        snapshot,
        story: { type: "main" },
        operations: [operation],
        applied: result.applied,
        author: AUTHOR,
        activeSuggestionIds: before,
      });
    },
    exportPending: () =>
      registry.exportPendingSuggestions({
        activeSuggestionIds: active(),
        snapshotForStory: (_story, commentIds) =>
          createPendingSuggestionSourceSnapshot(state, commentIds),
      }),
    loadPending: (records: readonly unknown[]) =>
      registry.loadPendingSuggestions({
        records,
        snapshotForStory: () => createFolioAIEditSnapshot(state.doc),
        sourceSnapshotForStory: (_story, commentIds) =>
          createPendingSuggestionSourceSnapshot(state, commentIds),
        apply: (record) => run(record.operation, record.author),
        activeSuggestionIds: active,
      }),
    resolve: (resolution: "accept" | "reject") =>
      publish(
        session
          .prepareSuggestion(state, {
            suggestionId: "proposal",
            resolution,
            author: "Adopter",
            date: "2026-02-02T00:00:00Z",
          })
          .unwrap(),
      ),
    undo: () => publish(session.prepareUndo(state).unwrap()),
    save: async () =>
      createDocx(withoutPendingSuggestions(session.document).unwrap()),
  };
};

/** What a reader of the saved file sees: the text and every comment. */
const readSaved = async (bytes: ArrayBuffer) => {
  const reader = await FolioDocxReviewer.fromBuffer(bytes);
  return {
    text: reader.snapshot().blocks.map(({ text }) => text),
    comments: reader.getComments().map(({ text, anchoredText, author }) => ({
      text,
      anchoredText,
      author,
    })),
  };
};

type Outcomes = Record<"pending" | "rejected" | "accepted", unknown>;

const canonicalOutcomes = async (
  operation: FolioDocumentOperation,
): Promise<Outcomes> => {
  const host = await canonicalHost(await sourceBytes());
  host.suggest(operation);
  const suggestion = getSuggestions(host.state()).find(
    ({ suggestionId }) => suggestionId === "proposal",
  );
  expect(suggestion?.kinds).toContain("comment");
  const records = JSON.parse(JSON.stringify(host.exportPending()));
  expect(records).toHaveLength(1);
  expect(records.at(0)?.commentId).toBeDefined();
  const saved = await host.save();
  const restore = async () => {
    const fresh = await canonicalHost(saved);
    expect(fresh.loadPending(records)).toEqual([
      { status: "restaged", suggestionId: "proposal" },
    ]);
    expect(getSuggestions(fresh.state()).map(({ kinds }) => kinds)).toEqual([
      suggestion?.kinds ?? [],
    ]);
    return fresh;
  };
  const rejecting = await restore();
  const beforeReject = rejecting.session.document;
  // Resolving undoes as one step.
  expect(rejecting.resolve("reject")).toBe(true);
  expect(rejecting.undo()).toBe(true);
  expect(rejecting.session.document).toEqual(beforeReject);
  expect(rejecting.resolve("reject")).toBe(true);
  expect(getSuggestions(rejecting.state())).toEqual([]);
  expect(rejecting.session.document.package.document.comments ?? []).toEqual(
    [],
  );
  expect(rejecting.exportPending()).toEqual([]);
  const accepting = await restore();
  const beforeAccept = accepting.session.document;
  expect(accepting.resolve("accept")).toBe(true);
  expect(accepting.undo()).toBe(true);
  expect(accepting.session.document).toEqual(beforeAccept);
  expect(accepting.resolve("accept")).toBe(true);
  expect(getSuggestions(accepting.state())).toEqual([]);
  return {
    pending: await readSaved(saved),
    rejected: await readSaved(await rejecting.save()),
    accepted: await readSaved(await accepting.save()),
  };
};

describe("suggested comments in the canonical session", () => {
  for (const [name, operation] of Object.entries(OPERATIONS)) {
    test(`${name}: pending, restored, rejected and accepted`, async () => {
      const proposed = operation();
      const comment =
        "comment" in proposed ? proposed.comment?.text : undefined;
      const outcomes = await canonicalOutcomes(operation());
      // Pending and rejected, the file holds no comment; accepted, it does.
      expect(outcomes.pending).toMatchObject({ comments: [] });
      expect(outcomes.rejected).toMatchObject({ comments: [] });
      expect(outcomes.accepted).toMatchObject({
        comments: [{ text: comment, author: AUTHOR }],
      });
    });
  }
});

describe("an edit's attached comment covers the same text in both engines", () => {
  // Edits that only insert or only delete words.
  const EDITS = [
    ["clause stays", "clause still stays"],
    ["Second clause", "Second short clause"],
    ["Second clause stays", "Second stays"],
  ] as const;
  for (const mode of ["direct", "tracked-changes"] as const)
    for (const [find, replace] of EDITS)
      test(`${mode}: ${find} -> ${replace}`, async () => {
        const operation = {
          id: "edit",
          type: "replaceInBlock",
          blockId: SECOND_ID,
          find,
          replace,
          comment: { text: "Why?" },
        } as const;
        const reviewer = await FolioDocxReviewer.fromBuffer(
          await sourceBytes(),
        );
        reviewer.applyOperations([operation], {
          undefinedReferences: "refuse",
          mode,
        });
        const expected = await readSaved(await reviewer.toBuffer());
        expect(expected.comments).toHaveLength(1);

        const session = createCanonicalSession(
          await parseDocx(await sourceBytes()),
        ).unwrap();
        let state = EditorState.create({ doc: session.projection.doc });
        const executor = new CanonicalPublicOperations({
          session,
          getState: () => state,
          publish: (commit) => {
            const result = publishCanonicalProjection({
              state,
              commit,
              session,
            });
            if (result.isErr()) return false;
            state = result.value.state;
            return true;
          },
        });
        const result = executor.apply({
          snapshot: createFolioAIEditSnapshot(state.doc),
          undefinedReferences: "refuse",
          batch: {
            version: FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
            operations: [operation],
            mode,
          },
        });
        expect(result.skipped).toEqual([]);
        expect(
          await readSaved(
            await createDocx(
              withoutPendingSuggestions(session.document).unwrap(),
            ),
          ),
        ).toEqual(expected);
      });
});
