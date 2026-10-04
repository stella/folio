import { describe, expect, test, setDefaultTimeout } from "bun:test";
import fc from "fast-check";
import { assertProperty, propertyTestTimeout } from "../../../../test/property-testing";

setDefaultTimeout(propertyTestTimeout(30_000));
import { EditorState } from "prosemirror-state";
import { panic } from "better-result";
import {
  hashFolioAIBlockText,
  createFolioAIEditSnapshot,
  createFolioAITextRangeHandle,
} from "../ai-edits/snapshot";
import { CANONICAL_PUBLIC_OPERATION_DISPOSITIONS } from "./canonicalPublicOperations";
import { createDocx } from "../docx/rezip";
import { parseDocx } from "../docx/parser";
import { toProseDoc } from "../prosemirror/conversion/toProseDoc";
import type { Document } from "../types/document";
import type { FolioDocumentOperation, FolioDocumentOperationBatch } from "../document-operations";
import { CanonicalPublicOperations } from "./canonicalPublicOperations";
import {
  createCanonicalSession,
  publishCanonicalProjection,
  type CanonicalCommit,
} from "./canonicalSession";

const setup = (text = "alpha beta") => {
  const document = {
    package: {
      document: {
        content: [
          {
            type: "paragraph",
            paraId: "12345678",
            content: [{ type: "run", content: [{ type: "text", text }] }],
          },
          { type: "paragraph", paraId: "23456789", content: [] },
        ],
      },
    },
  } satisfies Document;
  const created = createCanonicalSession(document);
  if (created.isErr()) return panic(created.error.message);
  const session = created.value;
  let state = EditorState.create({ doc: session.projection.doc });
  const publish = (commit: CanonicalCommit) => {
    const result = publishCanonicalProjection({ state, commit, session });
    if (result.isErr()) return false;
    state = result.value.state;
    return true;
  };
  const executor = new CanonicalPublicOperations({ session, getState: () => state, publish });
  const snapshot = () => createFolioAIEditSnapshot(state.doc);
  const apply = (
    operations: readonly FolioDocumentOperation[],
    extra: Omit<FolioDocumentOperationBatch, "version" | "operations"> = {},
  ) =>
    executor.apply({
      snapshot: snapshot(),
      batch: { version: 1, mode: "direct", operations, ...extra },
    });
  return { session, executor, snapshot, apply, publish, state: () => state };
};

describe("canonical public operation batches", () => {
  test("generated mixed human and public histories undo through one journal", () => {
    assertProperty(
      fc.property(
        fc.array(fc.stringMatching(/^[a-z]{1,8}$/), { minLength: 1, maxLength: 8 }),
        (words) => {
          const editor = setup("seed");
          const initial = editor.session.document;
          for (const word of words) {
            const before = editor.snapshot().blocks.at(0)?.text ?? panic("Missing paragraph");
            const result = editor.apply([
              { id: word, type: "replaceBlock", blockId: "12345678", text: before + word },
            ]);
            expect(result.applied).toHaveLength(1);
            const human = editor.session.prepareReplace(editor.state(), {
              from: 1,
              to: 1,
              text: word,
            });
            if (human.isErr()) return panic(human.error.message);
            expect(editor.publish(human.value)).toBe(true);
            expect(
              editor.executor.undo(result.undoHandle ?? panic("Missing handle")),
            ).toMatchObject({ status: "rejected", reason: "documentChanged" });
            for (let index = 0; index < 2; index++) {
              const undo = editor.session.prepareUndo(editor.state());
              if (undo.isErr()) return panic(undo.error.message);
              expect(editor.publish(undo.value)).toBe(true);
            }
            expect(editor.session.document).toEqual(initial);
          }
        },
      ),
      { numRuns: 30, seed: 20261004 },
    );
  });

  test("atomic refusal and preview preserve model, projection, version and journal", () => {
    const editor = setup();
    const before = editor.session.document;
    const state = editor.state();
    const operation = {
      id: "replace",
      type: "replaceInBlock",
      blockId: "12345678",
      find: "alpha",
      replace: "gamma",
    } as const;
    const refused = editor.apply(
      [
        operation,
        { id: "comment", type: "commentOnBlock", blockId: "12345678", comment: { text: "review" } },
      ],
      { atomic: true },
    );
    expect(refused.status).toBe("rejected");
    expect(
      refused.issues.find(({ operationId }) => operationId === "comment")?.canonicalRefusal,
    ).toEqual({ gap: "publicOps.comments" });
    expect(editor.session.document).toBe(before);
    expect(editor.state()).toBe(state);
    expect(editor.session.version).toBe(0);
    expect(editor.session.canUndo).toBe(false);
    const preview = editor.apply([operation], { dryRun: true });
    expect(preview.status).toBe("previewed");
    expect(preview.applied).toEqual([{ id: "replace" }]);
    expect(preview.undoHandle).toBeNull();
    expect(editor.session.document).toBe(before);
    expect(editor.state()).toBe(state);
    expect(editor.session.version).toBe(0);
    expect(editor.session.canUndo).toBe(false);
  });

  test("a batch publishes one version and one isolated inverse group", () => {
    const editor = setup();
    const original = editor.session.document;
    const result = editor.apply([
      { id: "a", type: "replaceInBlock", blockId: "12345678", find: "alpha", replace: "first" },
      { id: "b", type: "replaceInBlock", blockId: "12345678", find: "beta", replace: "second" },
    ]);
    expect(result.applied.map(({ id }) => id)).toEqual(["a", "b"]);
    expect(result.receipts.map(({ operationId }) => operationId)).toEqual(["a", "b"]);
    expect(editor.session.version).toBe(1);
    expect(editor.snapshot().blocks.at(0)?.text).toBe("first second");
    expect(editor.executor.undo(result.undoHandle ?? panic("Missing handle")).status).toBe(
      "undone",
    );
    expect(editor.session.document).toEqual(original);
    expect(editor.session.canUndo).toBe(false);
  });

  test("tracked coordinates use visible text across retained deletions", () => {
    const editor = setup();
    const first = editor.apply(
      [
        {
          id: "first",
          type: "replaceInBlock",
          blockId: "12345678",
          find: "alpha",
          replace: "gamma",
        },
      ],
      { mode: "tracked-changes" },
    );
    expect(first.applied).toHaveLength(1);
    const snapshot = editor.snapshot();
    const range = createFolioAITextRangeHandle({
      text: snapshot.blocks.at(0)?.text ?? panic("Missing block"),
      blockId: "12345678",
      startOffset: 6,
      endOffset: 10,
    });
    if (!range) return panic("Missing range");
    const second = editor.executor.apply({
      snapshot,
      batch: {
        version: 1,
        mode: "tracked-changes",
        operations: [{ id: "second", type: "replaceRange", range, replace: "delta" }],
      },
    });
    expect(second.applied).toHaveLength(1);
    expect(editor.snapshot().blocks.at(0)?.text).toBe("gamma delta");
    expect(editor.session.version).toBe(2);
  });
});

type CompiledKind = {
  [Kind in keyof typeof CANONICAL_PUBLIC_OPERATION_DISPOSITIONS]: (typeof CANONICAL_PUBLIC_OPERATION_DISPOSITIONS)[Kind] extends "compile"
    ? Kind
    : never;
}[keyof typeof CANONICAL_PUBLIC_OPERATION_DISPOSITIONS];

test("all supported compiler kinds share direct/tracked undo and typed refusals", () => {
  const snapshot = setup().snapshot();
  const range =
    createFolioAITextRangeHandle({
      blockId: "12345678",
      text: "alpha beta",
      startOffset: 0,
      endOffset: 5,
    }) ?? panic("Missing range");
  const operations = {
    replaceInBlock: {
      id: "replaceInBlock",
      type: "replaceInBlock",
      blockId: "12345678",
      find: "alpha",
      replace: "gamma",
    },
    replaceRange: { id: "replaceRange", type: "replaceRange", range, replace: "gamma" },
    replaceBlock: {
      id: "replaceBlock",
      type: "replaceBlock",
      blockId: "12345678",
      text: "gamma beta",
    },
    formatRange: { id: "formatRange", type: "formatRange", range, formatting: { bold: true } },
    splitBlock: {
      id: "splitBlock",
      type: "splitBlock",
      blockId: "12345678",
      offset: 5,
      separator: " ",
    },
    mergeBlockWithNext: {
      id: "mergeBlockWithNext",
      type: "mergeBlockWithNext",
      blockId: "12345678",
    },
  } as const satisfies Record<CompiledKind, FolioDocumentOperation>;
  for (const mode of ["direct", "tracked-changes"] as const) {
    for (const operation of Object.values(operations)) {
      const editor = setup();
      const original = editor.session.document;
      const result = editor.executor.apply({
        snapshot,
        batch: { version: 1, mode, operations: [operation] },
        revisionStamp: { date: "2026-01-01T00:00:00Z", idSeed: 500 },
      });
      expect(result.applied).toHaveLength(1);
      if (mode === "tracked-changes") {
        expect(result.applied.at(0)?.revisionIds?.every((id) => id >= 500)).toBe(true);
        expect(result.nextRevisionId).toBeGreaterThan(500);
      }
      expect(editor.executor.undo(result.undoHandle ?? panic("Missing handle")).status).toBe(
        "undone",
      );
      expect(editor.session.document).toEqual(original);
    }
  }
});

test("read preconditions and session lifecycle cannot mutate the journal", () => {
  const editor = setup();
  const operation = {
    id: "replace",
    type: "replaceInBlock",
    blockId: "12345678",
    find: "alpha",
    replace: "gamma",
  } as const;
  expect(
    editor
      .apply([{ ...operation, precondition: { blockTextHash: hashFolioAIBlockText("stale") } }])
      .skipped.at(0)?.reason,
  ).toBe("preconditionFailed");
  expect(editor.apply([operation], { mode: "suggested" }).issues.at(0)?.canonicalRefusal?.gap).toBe(
    "publicOps.suggestedMode",
  );
  expect(
    editor.apply([{ ...operation, replace: "**gamma**" }]).issues.at(0)?.canonicalRefusal?.gap,
  ).toBe("publicOps.unsupportedInline");
  expect(
    editor
      .apply([{ id: "table", type: "insertTable", blockId: "12345678", rows: [["cell"]] }])
      .issues.at(0)?.canonicalRefusal?.gap,
  ).toBe("publicOps.tableProjection");
  const before = editor.session.document;
  editor.session.beginComposition();
  expect(
    editor.executor
      .apply({ snapshot: editor.snapshot(), batch: { version: 1, operations: [operation] } })
      .skipped.at(0)?.reason,
  ).toBe("documentNotEditable");
  editor.session.endComposition();
  expect(editor.session.document).toBe(before);
  expect(editor.session.canUndo).toBe(false);
  expect(editor.session.version).toBe(0);
});

test("save observes one immutable committed batch version", async () => {
  const editor = setup();
  editor.apply([{ id: "first", type: "replaceBlock", blockId: "12345678", text: "saved" }]);
  const capturedVersion = editor.session.version;
  const save = createDocx(editor.session.document);
  editor.apply([{ id: "second", type: "replaceBlock", blockId: "12345678", text: "later" }]);
  const reopened = await parseDocx(await save, { preloadFonts: false });
  expect(toProseDoc(reopened).textContent).toBe("saved");
  expect(capturedVersion).toBe(1);
  expect(editor.session.version).toBe(2);
  expect(editor.snapshot().blocks.at(0)?.text).toBe("later");
});
