import { describe, expect, setDefaultTimeout, test } from "bun:test";
import fc from "fast-check";

import {
  assertProperty,
  propertyTestTimeout,
} from "../../../../test/property-testing";

setDefaultTimeout(propertyTestTimeout(30_000));
import { panic } from "better-result";
import { EditorState } from "prosemirror-state";

import { withoutPendingSuggestions } from "@stll/docx-core/ops";

import {
  createFolioAIEditSnapshot,
  createFolioAITextRangeHandle,
  hashFolioAIBlockText,
} from "../ai-edits/snapshot";
import {
  type CanonicalCommit,
  createCanonicalSession,
  publishCanonicalProjection,
} from "../canonical/canonicalSession";
import type {
  FolioDocumentOperation,
  FolioDocumentOperationBatch,
} from "../document-operations";
import { parseDocx } from "../docx/parser";
import { createDocx } from "../docx/rezip";
import { toProseDoc } from "../prosemirror/conversion/toProseDoc";
import type { Document } from "../types/document";
import {
  CANONICAL_PUBLIC_OPERATION_DISPOSITIONS,
  CanonicalPublicOperations,
} from "./canonicalPublicOperations";

const setup = (source: string | Document = "alpha beta") => {
  const document =
    typeof source === "string"
      ? ({
          package: {
            document: {
              content: [
                {
                  type: "paragraph",
                  paraId: "12345678",
                  content: [
                    { type: "run", content: [{ type: "text", text: source }] },
                  ],
                },
                { type: "paragraph", paraId: "23456789", content: [] },
              ],
            },
          },
        } satisfies Document)
      : source;
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
  const executor = new CanonicalPublicOperations({
    session,
    getState: () => state,
    publish,
  });
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
  test("merge with next refuses at table container boundaries", () => {
    const document = {
      package: {
        document: {
          content: [
            {
              type: "table",
              rows: [
                {
                  type: "tableRow",
                  cells: [
                    {
                      type: "tableCell",
                      content: [
                        {
                          type: "paragraph",
                          paraId: "12345678",
                          content: [
                            {
                              type: "run",
                              content: [{ type: "text", text: "A" }],
                            },
                          ],
                        },
                      ],
                    },
                    {
                      type: "tableCell",
                      content: [
                        {
                          type: "paragraph",
                          paraId: "23456789",
                          content: [
                            {
                              type: "run",
                              content: [{ type: "text", text: "B" }],
                            },
                          ],
                        },
                      ],
                    },
                  ],
                },
              ],
            },
            { type: "paragraph", paraId: "34567890", content: [] },
          ],
        },
      },
    } satisfies Document;
    for (const blockId of ["12345678", "23456789"] as const) {
      const editor = setup(document);
      const result = editor.apply([
        { id: "merge", type: "mergeBlockWithNext", blockId },
      ]);
      expect(result.applied).toEqual([]);
      expect(result.issues).toMatchObject([
        { operationId: "merge", code: "missingBlock" },
      ]);
      expect(editor.session.version).toBe(0);
    }
  });

  test("generated mixed human and public histories undo through one journal", () => {
    assertProperty(
      fc.property(
        fc.array(fc.stringMatching(/^[a-z]{1,8}$/), {
          minLength: 1,
          maxLength: 8,
        }),
        (words) => {
          const editor = setup("seed");
          const initial = editor.session.document;
          for (const word of words) {
            const before =
              editor.snapshot().blocks.at(0)?.text ??
              panic("Missing paragraph");
            const result = editor.apply([
              {
                id: word,
                type: "replaceBlock",
                blockId: "12345678",
                text: before + word,
              },
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
              editor.executor.undo(
                result.undoHandle ?? panic("Missing handle"),
              ),
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
      {
        numRuns: 30,
        seed: 20261004,
        id: "generated mixed human and public histories undo through one journal",
      },
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
      [operation, { id: "table", type: "deleteTable", blockId: "12345678" }],
      { atomic: true },
    );
    expect(refused.status).toBe("rejected");
    expect(
      refused.issues.find(({ operationId }) => operationId === "table")
        ?.canonicalRefusal,
    ).toEqual({ gap: "publicOps.tableStructure" });
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
      {
        id: "a",
        type: "replaceInBlock",
        blockId: "12345678",
        find: "alpha",
        replace: "first",
      },
      {
        id: "b",
        type: "replaceInBlock",
        blockId: "12345678",
        find: "beta",
        replace: "second",
      },
    ]);
    expect(result.applied.map(({ id }) => id)).toEqual(["a", "b"]);
    expect(result.receipts.map(({ operationId }) => operationId)).toEqual([
      "a",
      "b",
    ]);
    expect(editor.session.version).toBe(1);
    expect(editor.snapshot().blocks.at(0)?.text).toBe("first second");
    expect(
      editor.executor.undo(result.undoHandle ?? panic("Missing handle")).status,
    ).toBe("undone");
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
        operations: [
          { id: "second", type: "replaceRange", range, replace: "delta" },
        ],
      },
    });
    expect(second.applied).toHaveLength(1);
    expect(editor.snapshot().blocks.at(0)?.text).toBe("gamma delta");
    expect(editor.session.version).toBe(2);
  });
});

type CompiledKind = {
  [
    Kind in keyof typeof CANONICAL_PUBLIC_OPERATION_DISPOSITIONS
  ]: (typeof CANONICAL_PUBLIC_OPERATION_DISPOSITIONS)[Kind] extends "compile"
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
    replaceRange: {
      id: "replaceRange",
      type: "replaceRange",
      range,
      replace: "gamma",
    },
    replaceBlock: {
      id: "replaceBlock",
      type: "replaceBlock",
      blockId: "12345678",
      text: "gamma beta",
    },
    formatRange: {
      id: "formatRange",
      type: "formatRange",
      range,
      formatting: { bold: true },
    },
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
    commentOnBlock: {
      id: "commentOnBlock",
      type: "commentOnBlock",
      blockId: "12345678",
      comment: { text: "note" },
    },
    commentOnRange: {
      id: "commentOnRange",
      type: "commentOnRange",
      range,
      comment: { text: "note" },
    },
    deleteBlock: {
      id: "deleteBlock",
      type: "deleteBlock",
      blockId: "12345678",
    },
    insertAfterBlock: {
      id: "insertAfterBlock",
      type: "insertAfterBlock",
      blockId: "12345678",
      text: "gamma",
    },
    insertBeforeBlock: {
      id: "insertBeforeBlock",
      type: "insertBeforeBlock",
      blockId: "12345678",
      text: "gamma",
    },
    setBlockParagraphProperties: {
      id: "setBlockParagraphProperties",
      type: "setBlockParagraphProperties",
      blockId: "12345678",
      properties: { alignment: "center" },
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
      if (
        operation.type === "commentOnBlock" ||
        operation.type === "commentOnRange"
      ) {
        expect(result.applied.at(0)?.commentId).toBeDefined();
        expect(result.applied.at(0)?.revisionIds).toBeUndefined();
      } else if (mode === "tracked-changes") {
        expect(
          result.applied.at(0)?.revisionIds?.every((id) => id >= 500),
        ).toBe(true);
        expect(result.nextRevisionId).toBeGreaterThan(500);
      }
      expect(
        editor.executor.undo(result.undoHandle ?? panic("Missing handle"))
          .status,
      ).toBe("undone");
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
      .apply([
        {
          ...operation,
          precondition: { blockTextHash: hashFolioAIBlockText("stale") },
        },
      ])
      .skipped.at(0)?.reason,
  ).toBe("preconditionFailed");
  expect(
    editor.apply([{ ...operation, replace: "**gamma**" }]).issues.at(0)
      ?.canonicalRefusal?.gap,
  ).toBe("publicOps.unsupportedInline");
  expect(
    editor
      .apply([
        {
          id: "table",
          type: "insertTableRow",
          blockId: "12345678",
          position: "after",
          cellTexts: ["cell"],
        },
      ])
      .issues.at(0)?.canonicalRefusal?.gap,
  ).toBe("publicOps.tableStructure");
  const before = editor.session.document;
  editor.session.beginComposition();
  expect(
    editor.executor
      .apply({
        snapshot: editor.snapshot(),
        batch: { version: 1, operations: [operation] },
      })
      .skipped.at(0)?.reason,
  ).toBe("documentNotEditable");
  editor.session.endComposition();
  expect(editor.session.document).toBe(before);
  expect(editor.session.canUndo).toBe(false);
  expect(editor.session.version).toBe(0);
});

test("save observes one immutable committed batch version", async () => {
  const editor = setup();
  editor.apply([
    { id: "first", type: "replaceBlock", blockId: "12345678", text: "saved" },
  ]);
  const capturedVersion = editor.session.version;
  const save = createDocx(editor.session.document);
  editor.apply([
    { id: "second", type: "replaceBlock", blockId: "12345678", text: "later" },
  ]);
  const reopened = await parseDocx(await save, { preloadFonts: false });
  expect(toProseDoc(reopened).textContent).toBe("saved");
  expect(capturedVersion).toBe(1);
  expect(editor.session.version).toBe(2);
  expect(editor.snapshot().blocks.at(0)?.text).toBe("later");
});

test("generated attached comments publish and undo atomically in direct and tracked modes", () => {
  assertProperty(
    fc.property(
      fc.stringMatching(/^[a-z]{1,12}$/u),
      fc.constantFrom("direct", "tracked-changes"),
      (replacement, mode) => {
        const editor = setup("A😀éB");
        const initial = editor.session.document;
        const operation = {
          id: "commented",
          type: "replaceInBlock",
          blockId: "12345678",
          find: "😀é",
          replace: replacement,
          comment: { text: "review" },
        } as const;
        const preview = editor.apply([operation], { mode, dryRun: true });
        expect(preview.status).toBe("previewed");
        expect(editor.session.document).toBe(initial);
        const result = editor.apply([operation], { mode });
        expect(result.applied).toHaveLength(1);
        expect(result.applied.at(0)?.commentId).toBeDefined();
        expect(editor.session.document.package.document.comments).toHaveLength(
          1,
        );
        expect(
          editor.executor.undo(result.undoHandle ?? panic("Missing handle"))
            .status,
        ).toBe("undone");
        expect(editor.session.document).toEqual(initial);
      },
    ),
    {
      numRuns: 30,
      seed: 20261008,
      id: "generated attached comments publish and undo atomically in direct and tracked modes",
    },
  );
});

test.each(["splitBlock", "mergeBlockWithNext"] as const)(
  "the public parser refuses an attached comment on %s",
  (type) => {
    const editor = setup();
    const initial = editor.session.document;
    const operation = {
      id: "structural",
      type,
      blockId: "12345678",
      comment: { text: "review" },
      ...(type === "splitBlock" ? { offset: 2 } : {}),
    };
    expect(() => editor.apply([operation])).toThrow(
      "comment: unexpected property",
    );
    expect(editor.session.document).toBe(initial);
    expect(editor.session.version).toBe(0);
    expect(editor.session.canUndo).toBe(false);
  },
);

test("a block's spacing does not make its unchanged text read as changed", () => {
  const spaced = fc.tuple(
    fc.constantFrom("", " ", "  "),
    fc.constantFrom(" ", "  ", "   "),
    fc.constantFrom("", " ", "  "),
  );
  assertProperty(
    fc.property(
      spaced,
      fc.constantFrom("direct", "tracked-changes" as const),
      ([lead, gap, trail], mode) => {
        const editor = setup(`${lead}alpha${gap}beta${trail}`);
        const anchor = editor.snapshot().anchors["12345678"];
        if (!anchor) throw new Error("the paragraph has no anchor");
        const result = editor.apply(
          [
            {
              id: "replace",
              type: "replaceInBlock",
              blockId: "12345678",
              find: "beta",
              replace: "gamma",
              // A host passes back the hash the snapshot gave it.
              precondition: { blockTextHash: anchor.textHash },
            },
          ],
          { mode },
        );
        expect(result.skipped).toEqual([]);
        expect(result.applied.map(({ id }) => id)).toEqual(["replace"]);
      },
    ),
    {
      numRuns: 40,
      seed: 20261107,
      id: "a block's spacing does not make its unchanged text read as changed",
    },
  );
});

test("a suggested operation records one pending suggestion named by its id", () => {
  const editor = setup();
  const replace = {
    id: "replace",
    type: "replaceInBlock",
    blockId: "12345678",
    find: "alpha",
    replace: "gamma",
  } as const;
  const first = editor.apply([replace], { mode: "suggested" });
  expect(first.applied.map(({ suggestionId }) => suggestionId)).toEqual([
    "replace",
  ]);
  const suggestionIds = new Set<unknown>();
  const visit = (value: unknown): void => {
    if (typeof value !== "object" || value === null) return;
    if (Array.isArray(value)) {
      for (const item of value) visit(item);
      return;
    }
    const record = value as Record<string, unknown>;
    const info = record["info"];
    if (typeof info === "object" && info !== null && "id" in info)
      suggestionIds.add((info as { suggestionId?: unknown }).suggestionId);
    for (const child of Object.values(record)) visit(child);
  };
  visit(editor.session.document.package.document.content);
  // Every record the operation made carries its suggestion.
  expect([...suggestionIds]).toEqual(["replace"]);
  // The same operation id again is a second suggestion, not part of the first.
  const second = editor.apply(
    [{ ...replace, id: "replace", find: "beta", replace: "delta" }],
    { mode: "suggested" },
  );
  expect(second.applied.map(({ suggestionId }) => suggestionId)).toEqual([
    "replace#2",
  ]);
  // A comment is proposed as a suggestion of its own, and its record says so.
  const range = createFolioAITextRangeHandle({
    text: editor.snapshot().blocks.at(0)?.text ?? panic("Missing block"),
    blockId: "12345678",
    startOffset: 0,
    endOffset: 1,
  });
  if (!range) return panic("Missing range");
  const commented = editor.apply(
    [
      {
        id: "comment",
        type: "commentOnRange",
        range,
        comment: { text: "Why?" },
      },
    ],
    { mode: "suggested" },
  );
  expect(commented.applied.map(({ suggestionId }) => suggestionId)).toEqual([
    "comment",
  ]);
  expect(
    editor.session.document.package.document.comments?.map(
      ({ suggestionId }) => suggestionId,
    ),
  ).toEqual(["comment"]);
});

describe("suggested operations in the canonical session", () => {
  const range = () =>
    createFolioAITextRangeHandle({
      text: "alpha beta",
      blockId: "12345678",
      startOffset: 0,
      endOffset: 5,
    }) ?? panic("Missing range");
  const kinds = {
    replaceInBlock: () =>
      ({
        id: "s",
        type: "replaceInBlock",
        blockId: "12345678",
        find: "alpha",
        replace: "gamma",
      }) as const,
    replaceRange: () =>
      ({
        id: "s",
        type: "replaceRange",
        range: range(),
        replace: "gamma",
      }) as const,
    replaceBlock: () =>
      ({
        id: "s",
        type: "replaceBlock",
        blockId: "12345678",
        text: "gamma beta",
      }) as const,
    formatRange: () =>
      ({
        id: "s",
        type: "formatRange",
        range: range(),
        formatting: { bold: true },
      }) as const,
  } satisfies Record<string, () => FolioDocumentOperation>;
  /** The text of every paragraph, as accepting every revision leaves it. */
  const acceptedTexts = (editor: ReturnType<typeof setup>) => {
    let state = editor.state();
    const ids = new Set<number>();
    const visit = (value: unknown): void => {
      if (typeof value !== "object" || value === null) return;
      if (Array.isArray(value)) {
        for (const item of value) visit(item);
        return;
      }
      const record = value as Record<string, unknown>;
      const info = record["info"] as { id?: unknown } | undefined;
      if (typeof info?.id === "number") ids.add(info.id);
      for (const child of Object.values(record)) visit(child);
    };
    visit(editor.session.document.package.document.content);
    if (ids.size > 0) {
      const commit = editor.session
        .prepareResolve(state, { revisionIds: [...ids], resolution: "accept" })
        .unwrap();
      expect(editor.publish(commit)).toBe(true);
      state = editor.state();
    }
    return createFolioAIEditSnapshot(state.doc).blocks.map(({ text }) => text);
  };

  for (const [kind, operation] of Object.entries(kinds))
    test(`${kind} records a suggestion the save leaves out, rejecting restores and accepting matches the tracked edit`, () => {
      // The edit as an ordinary tracked change, accepted.
      const tracked = setup();
      tracked.apply([operation()], { mode: "tracked-changes" });
      const expected = acceptedTexts(tracked);

      for (const resolution of ["reject", "accept"] as const) {
        const editor = setup();
        const original = editor.session.document;
        const result = editor.apply([operation()], { mode: "suggested" });
        expect(result.applied.map(({ suggestionId }) => suggestionId)).toEqual([
          "s",
        ]);
        // The file a save writes holds none of it.
        expect(
          withoutPendingSuggestions(editor.session.document).unwrap(),
        ).toEqual(original);
        const commit = editor.session
          .prepareSuggestion(editor.state(), {
            suggestionId: "s",
            resolution,
            author: "Adopter",
            date: "2026-02-02T00:00:00Z",
          })
          .unwrap();
        expect(editor.publish(commit)).toBe(true);
        if (resolution === "reject")
          expect(editor.session.document).toEqual(original);
        else expect(acceptedTexts(editor)).toEqual(expected);
      }
    });

  test("operations a suggestion cannot carry skip as unsupported in this mode", () => {
    const editor = setup();
    const result = editor.apply(
      [
        {
          id: "merge",
          type: "mergeTableCells",
          blockId: "12345678",
          rowCount: 2,
        },
      ],
      { mode: "suggested" },
    );
    expect(result.skipped.map(({ reason }) => reason)).toEqual([
      "unsupportedMode",
    ]);
  });
});

describe("a paragraph after one whose mark is pending deletion", () => {
  /** Every paragraph's words, as resolving every revision leaves them. */
  const resolvedWords = (
    editor: ReturnType<typeof setup>,
    resolution: "accept" | "reject",
  ) => {
    const ids = new Set<number>();
    const visit = (value: unknown): void => {
      if (typeof value !== "object" || value === null) return;
      if (Array.isArray(value)) {
        for (const item of value) visit(item);
        return;
      }
      const record = value as Record<string, unknown>;
      const info = record["info"] as { id?: unknown } | undefined;
      if (typeof info?.id === "number") ids.add(info.id);
      for (const child of Object.values(record)) visit(child);
    };
    visit(editor.session.document.package.document.content);
    if (ids.size > 0) {
      const commit = editor.session
        .prepareResolve(editor.state(), { revisionIds: [...ids], resolution })
        .unwrap();
      expect(editor.publish(commit)).toBe(true);
    }
    return createFolioAIEditSnapshot(editor.state().doc)
      .blocks.map(({ text }) => text)
      .join("")
      .replace(/\s+/gu, "");
  };

  // The first paragraph's mark is deleted or moved away, so in the editing
  // view the two paragraphs show as one; the second ends the body.
  const documentOf = (kind: "del" | "moveFrom", text: string) =>
    ({
      package: {
        document: {
          content: [
            {
              type: "paragraph",
              paraId: "12345678",
              formatting: { alignment: "center" },
              pPrMark: {
                kind,
                info: {
                  id: 1,
                  author: "Earlier",
                  date: "2026-01-02T03:04:05Z",
                },
              },
              content: [
                {
                  type: "run",
                  content: [{ type: "text", text: "alpha beta" }],
                },
              ],
            },
            {
              type: "paragraph",
              paraId: "23456789",
              formatting: { alignment: "end" },
              content:
                text === ""
                  ? []
                  : [{ type: "run", content: [{ type: "text", text }] }],
            },
          ],
        },
      },
    }) satisfies Document;

  for (const kind of ["del", "moveFrom"] as const)
    for (const text of ["gamma", ""])
      describe(`pending as ${kind}, the second ${text === "" ? "empty" : "with words"}`, () => {
        const deletion = {
          id: "delete",
          type: "deleteBlock",
          blockId: "23456789",
        } as const;

        test("deleting the second directly applies and undoes exactly", () => {
          const editor = setup(documentOf(kind, text));
          const original = editor.session.document;
          const result = editor.apply([deletion]);
          expect(result.applied).toHaveLength(1);
          expect(
            editor.executor.undo(result.undoHandle ?? panic("Missing handle"))
              .status,
          ).toBe("undone");
          expect(editor.session.document).toEqual(original);
        });

        if (text === "")
          test("deleting the empty second tracked records nothing", () => {
            // No words to mark deleted, and the mark before it already is.
            const editor = setup(documentOf(kind, text));
            const original = editor.session.document;
            const result = editor.apply([deletion], {
              mode: "tracked-changes",
            });
            expect(result.skipped.map(({ reason }) => reason)).toEqual([
              "noopOperation",
            ]);
            expect(editor.session.document).toEqual(original);
          });
        else
          test("deleting the second tracked: rejecting restores, accepting matches the direct deletion", () => {
            const untouched = setup(documentOf(kind, text));
            const tracked = setup(documentOf(kind, text));
            expect(
              tracked.apply([deletion], { mode: "tracked-changes" }).applied,
            ).toHaveLength(1);
            expect(resolvedWords(tracked, "reject")).toBe(
              resolvedWords(untouched, "reject"),
            );
            const accepted = setup(documentOf(kind, text));
            accepted.apply([deletion], { mode: "tracked-changes" });
            const direct = setup(documentOf(kind, text));
            expect(direct.apply([deletion]).applied).toHaveLength(1);
            expect(resolvedWords(accepted, "accept")).toBe(
              resolvedWords(direct, "accept"),
            );
          });
      });
});

describe("a paragraph the session holds no address in", () => {
  // A text box's paragraph belongs to its shape, not to the body story.
  const document = {
    package: {
      document: {
        content: [
          {
            type: "paragraph",
            paraId: "12345678",
            content: [
              {
                type: "run",
                content: [{ type: "text", text: "The box restates it. " }],
              },
              {
                type: "run",
                content: [
                  {
                    type: "shape",
                    shape: {
                      type: "shape",
                      shapeType: "textBox",
                      id: "box",
                      size: { width: 2_000_000, height: 800_000 },
                      textBody: {
                        content: [
                          {
                            type: "paragraph",
                            paraId: "34567890",
                            content: [
                              {
                                type: "run",
                                content: [
                                  { type: "text", text: "Boxed note." },
                                ],
                              },
                            ],
                          },
                        ],
                      },
                    },
                  },
                ],
              },
            ],
          },
          { type: "paragraph", paraId: "23456789", content: [] },
        ],
      },
    },
  } satisfies Document;

  const operations = [
    {
      id: "insertAfterBlock",
      type: "insertAfterBlock",
      blockId: "34567890",
      text: "Added note.",
    },
    {
      id: "insertBeforeBlock",
      type: "insertBeforeBlock",
      blockId: "34567890",
      text: "Added note.",
    },
    {
      id: "replaceInBlock",
      type: "replaceInBlock",
      blockId: "34567890",
      find: "Boxed",
      replace: "Framed",
    },
    { id: "deleteBlock", type: "deleteBlock", blockId: "34567890" },
  ] as const satisfies readonly FolioDocumentOperation[];

  for (const mode of ["direct", "tracked-changes"] as const)
    for (const operation of operations)
      test(`${operation.type} refuses under a typed gap, ${mode}`, () => {
        const editor = setup(document);
        // A reader lists the paragraph, so the operation names a block that exists.
        expect(
          editor.snapshot().blocks.some(({ id }) => id === "34567890"),
        ).toBe(true);
        const original = editor.session.document;
        const result = editor.apply([operation], { mode });
        expect(result.applied).toEqual([]);
        expect(
          result.skipped.map(({ reason, canonicalRefusal }) => ({
            reason,
            gap: canonicalRefusal?.gap,
          })),
        ).toEqual([
          { reason: "unsupportedBlock", gap: "publicOps.unsupportedInline" },
        ]);
        expect(editor.session.document).toEqual(original);
      });

  test("a block no reader lists is still missing", () => {
    const editor = setup(document);
    const result = editor.apply([
      {
        id: "insertAfterBlock",
        type: "insertAfterBlock",
        blockId: "7FFFFFF0",
        text: "Added note.",
      },
    ]);
    expect(result.skipped.map(({ reason }) => reason)).toEqual([
      "missingBlock",
    ]);
  });
});
