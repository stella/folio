import { describe, expect, setDefaultTimeout, test } from "bun:test";
import fc from "fast-check";

import {
  assertProperty,
  propertyTestTimeout,
} from "../../../../test/property-testing";

setDefaultTimeout(propertyTestTimeout(30_000));
import { panic } from "better-result";
import { EditorState } from "prosemirror-state";
import { CellSelection } from "prosemirror-tables";

import { withoutPendingSuggestions } from "@stll/docx-core/ops";

import {
  getCommentAnchorsFromDoc,
  getTrackedChangesFromDoc,
} from "../ai-edits/read";
import {
  createFolioAIEditSnapshot,
  createFolioAITextRangeHandle,
  hashFolioAIBlockText,
} from "../ai-edits/snapshot";
import {
  CANONICAL_PUBLIC_OPERATION_DISPOSITIONS,
  CanonicalPublicOperations,
} from "../canonical/canonicalPublicOperations";
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
import { toFlowBlocks } from "../layout-bridge/convert/toFlowBlocks";
import type { FlowBlock } from "../layout-engine/types";
import { fromMarkdown } from "../markdown/fromMarkdown";
import { toProseDoc } from "../prosemirror/conversion/toProseDoc";
import type { Document } from "../types/document";

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
    undoLimit: Number.POSITIVE_INFINITY,
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
  test("cell replacement covers every sorted cell independently of selection direction, taking a comment with its words", () => {
    const source = fromMarkdown("| A | B |\n|---|---|\n| C | D |");
    const table = source.package.document.content.at(0);
    if (table?.type !== "table") return panic("Produced table is missing");
    let id = 1;
    for (const row of table.rows)
      for (const cell of row.cells)
        for (const block of cell.content)
          if (block.type === "paragraph")
            block.paraId = (id++).toString(16).padStart(8, "0");
    for (const direction of ["forward", "reverse"] as const) {
      const editor = setup(source);
      expect(
        editor.apply([
          {
            id: "comment",
            type: "commentOnBlock",
            blockId: "00000001",
            quote: "A",
            comment: { text: "Note" },
          },
        ]).applied,
      ).toHaveLength(1);
      const positions = new Map<string, number>();
      editor.state().doc.descendants((node, position) => {
        if (node.type.name === "tableCell" || node.type.name === "tableHeader")
          positions.set(node.textContent, position);
      });
      const start = positions.get("A");
      const end = positions.get("D");
      if (start === undefined || end === undefined)
        return panic("Produced cell is missing");
      const selection = CellSelection.create(
        editor.state().doc,
        direction === "forward" ? start : end,
        direction === "forward" ? end : start,
      );
      const state = editor
        .state()
        .apply(editor.state().tr.setSelection(selection));
      const result = editor.session.prepareReplace(state, {
        from: selection.from,
        to: selection.to,
        text: "replacement",
      });
      expect(result.isOk()).toBe(true);
      expect(editor.publish(result.unwrap())).toBe(true);
      // The text goes into the first cell whichever way the selection ran,
      // and the comment on the cleared words goes with them.
      const cells: string[] = [];
      editor.state().doc.descendants((node) => {
        if (node.type.name === "tableCell" || node.type.name === "tableHeader")
          cells.push(node.textContent);
        return true;
      });
      expect(cells).toEqual(["replacement", "", "", ""]);
      expect(getCommentAnchorsFromDoc(editor.state().doc)).toEqual([]);
      expect(editor.session.document.package.document.comments ?? []).toEqual(
        [],
      );
    }
  });

  test("a direct replacement inside an inline control takes the comment with its words", () => {
    const source = "The Supplier delivers on time.";
    const editor = setup({
      package: {
        document: {
          content: [
            {
              type: "paragraph",
              paraId: "12345678",
              content: [
                {
                  type: "inlineSdt",
                  properties: { sdtType: "richText", id: 9 },
                  content: [
                    { type: "run", content: [{ type: "text", text: source }] },
                  ],
                },
              ],
            },
          ],
        },
      },
    });
    expect(
      editor.apply([
        {
          id: "comment",
          type: "commentOnBlock",
          blockId: "12345678",
          quote: "time",
          comment: { text: "Note" },
        },
      ]).applied,
    ).toHaveLength(1);
    const result = editor.apply(
      [
        {
          id: "replace",
          type: "replaceBlock",
          blockId: "12345678",
          text: "New agreement.",
        },
      ],
      { mode: "direct" },
    );
    expect(result.skipped).toEqual([]);
    expect(result.applied).toHaveLength(1);
    expect(editor.state().doc.textContent).toBe("New agreement.");
    expect(getCommentAnchorsFromDoc(editor.state().doc)).toEqual([]);
    expect(editor.session.document.package.document.comments ?? []).toEqual([]);
  });

  test(
    "generated source comment ranges retain aligned replacement coverage or refuse atomically",
    () => {
      assertProperty(
        fc.property(
          fc.integer({ min: 0, max: 19 }),
          fc.integer({ min: 1, max: 20 }),
          fc.constantFrom("direct", "tracked-changes", "suggested"),
          (start, length, mode) => {
            const source = "abcdefghijklmnopqrst";
            // Every character goes, including the last one inside the comment.
            const replacement = "REVISED!";
            const editor = setup(source);
            const quote = source.slice(start, start + length);
            expect(
              editor.apply([
                {
                  id: "comment",
                  type: "commentOnBlock",
                  blockId: "12345678",
                  quote,
                  comment: { text: "Note" },
                },
              ]).applied,
            ).toHaveLength(1);
            const before = editor.session.document;
            const version = editor.session.version;
            const result = editor.apply(
              [
                {
                  id: "replace",
                  type: "replaceBlock",
                  blockId: "12345678",
                  text: replacement,
                },
              ],
              { mode },
            );
            if (result.applied.length === 0) {
              // Only a suggestion that must move the comment refuses: a
              // direct replacement takes a comment with its words, and a
              // tracked one keeps it on them.
              expect(mode).toBe("suggested");
              expect(result.skipped).toMatchObject([
                {
                  reason: "unsupportedBlock",
                  canonicalRefusal: { gap: "comment-replacement-alignment" },
                },
              ]);
              expect(editor.session.document).toBe(before);
              expect(editor.session.version).toBe(version);
              return;
            }
            expect(result.skipped).toEqual([]);
            if (mode === "tracked-changes") {
              // The comment stays on the words the replacement deletes:
              // rejecting gives it back, accepting keeps it only where the
              // replacement took it over.
              const ids = getTrackedChangesFromDoc(editor.state().doc).map(
                ({ id }) => id,
              );
              const rejected = editor.session
                .prepareResolve(editor.state(), {
                  revisionIds: ids,
                  resolution: "reject",
                })
                .unwrap();
              expect(editor.publish(rejected)).toBe(true);
              expect(
                getCommentAnchorsFromDoc(editor.state().doc).map(
                  ({ quote: anchoredQuote }) => anchoredQuote,
                ),
              ).toEqual([quote]);
              expect(
                editor.publish(
                  editor.session.prepareUndo(editor.state()).unwrap(),
                ),
              ).toBe(true);
            }
            if (mode === "suggested") {
              const adopted = editor.session
                .prepareSuggestion(editor.state(), {
                  suggestionId: "replace",
                  resolution: "accept",
                  author: "Reviewer",
                  date: "2026-01-02T03:04:05Z",
                })
                .unwrap();
              expect(editor.publish(adopted)).toBe(true);
            }
            if (mode !== "direct") {
              const accepted = editor.session
                .prepareResolve(editor.state(), {
                  revisionIds: getTrackedChangesFromDoc(editor.state().doc).map(
                    ({ id }) => id,
                  ),
                  resolution: "accept",
                })
                .unwrap();
              expect(editor.publish(accepted)).toBe(true);
            }
            const anchors = getCommentAnchorsFromDoc(editor.state().doc).map(
              ({ quote: anchoredQuote }) => anchoredQuote,
            );
            // The replacement takes every source word, and with them the
            // comment: the new words never inherit its range.
            expect(anchors).toEqual([]);
          },
        ),
        {
          id: "generated replacement comment alignment is preserved or refused",
          seed: 1683001,
          numRuns: 48,
        },
      );
    },
    propertyTestTimeout(30_000),
  );

  test("a tracked replacement over separate comments keeps each on its deleted words", () => {
    const source = "The Supplier delivers the goods on time and in good order.";
    for (const type of ["replaceBlock", "replaceInBlock"] as const) {
      const editor = setup(source);
      for (const quote of ["time", "The"])
        expect(
          editor.apply([
            {
              id: quote,
              type: "commentOnBlock",
              blockId: "12345678",
              quote,
              comment: { text: quote },
            },
          ]).applied,
        ).toHaveLength(1);
      const text =
        "Invoice delivery payment agreement supplier buyer payment term.";
      const result = editor.apply(
        [
          type === "replaceBlock"
            ? { id: "replace", type, blockId: "12345678", text }
            : {
                id: "replace",
                type,
                blockId: "12345678",
                find: source,
                replace: text,
              },
        ],
        { mode: "tracked-changes" },
      );
      expect(result.skipped).toEqual([]);
      expect(result.applied.map(({ id }) => id)).toEqual(["replace"]);
      const quotes = () =>
        getCommentAnchorsFromDoc(editor.state().doc)
          .map(({ quote }) => quote)
          .toSorted();
      // Pending, both comments still cover their deleted words.
      expect(quotes()).toEqual(["The", "time"]);
      const ids = getTrackedChangesFromDoc(editor.state().doc).map(
        ({ id }) => id,
      );
      const resolve = (resolution: "accept" | "reject") => {
        const resolved = editor.session
          .prepareResolve(editor.state(), { revisionIds: ids, resolution })
          .unwrap();
        expect(editor.publish(resolved)).toBe(true);
      };
      resolve("reject");
      expect(editor.state().doc.textContent).toBe(source);
      expect(quotes()).toEqual(["The", "time"]);
      expect(
        editor.publish(editor.session.prepareUndo(editor.state()).unwrap()),
      ).toBe(true);
      resolve("accept");
      expect(editor.state().doc.textContent).toBe(text);
      expect(quotes()).toEqual([]);
    }
  });

  const REPLACEMENT_SOURCE =
    "The Supplier delivers the goods on time and in good order.";
  const REPLACEMENT_TYPES = [
    "replaceBlock",
    "replaceInBlock",
    "replaceRange",
  ] as const;
  type ReplacementOperationOptions = {
    type: (typeof REPLACEMENT_TYPES)[number];
    text: string;
  };
  const wholeBlockReplacement = ({
    type,
    text,
  }: ReplacementOperationOptions): FolioDocumentOperation => {
    switch (type) {
      case "replaceBlock":
        return { id: "replace", type, blockId: "12345678", text };
      case "replaceInBlock":
        return {
          id: "replace",
          type,
          blockId: "12345678",
          find: REPLACEMENT_SOURCE,
          replace: text,
        };
      case "replaceRange":
        return {
          id: "replace",
          type,
          range: {
            type: "textRange",
            story: "main",
            blockId: "12345678",
            startOffset: 0,
            endOffset: REPLACEMENT_SOURCE.length,
            selectedTextHash: hashFolioAIBlockText(REPLACEMENT_SOURCE),
          },
          replace: text,
        };
      default: {
        const unhandled: never = type;
        return panic(`Unknown replacement type ${unhandled}`);
      }
    }
  };
  const commentedReplacementSource = () => {
    const editor = setup(REPLACEMENT_SOURCE);
    const comment = (quote: string) =>
      editor.apply([
        {
          id: quote,
          type: "commentOnBlock",
          blockId: "12345678",
          quote,
          comment: { text: quote },
        },
      ]);
    expect(comment("time").applied).toHaveLength(1);
    const firstComment = editor.session.document;
    expect(comment("The").applied).toHaveLength(1);
    return { editor, firstComment };
  };

  for (const text of [
    "Invoice delivery payment agreement supplier buyer payment term.",
    "",
  ]) {
    test(`direct/${text === "" ? "empty" : "nonempty"}: a whole-block replacement takes the comments with their words, as accepting the tracked one does`, () => {
      for (const type of REPLACEMENT_TYPES) {
        const operation = wholeBlockReplacement({ type, text });
        const { editor: direct } = commentedReplacementSource();
        const before = direct.session.document;
        const result = direct.apply([operation], { mode: "direct" });
        expect(result.skipped).toEqual([]);
        expect(result.applied).toHaveLength(1);
        const { editor: tracked } = commentedReplacementSource();
        expect(
          tracked.apply([operation], { mode: "tracked-changes" }).skipped,
        ).toEqual([]);
        const accepted = tracked.session
          .prepareResolve(tracked.state(), {
            revisionIds: getTrackedChangesFromDoc(tracked.state().doc).map(
              ({ id }) => id,
            ),
            resolution: "accept",
          })
          .unwrap();
        expect(tracked.publish(accepted)).toBe(true);
        for (const editor of [direct, tracked]) {
          expect(editor.state().doc.textContent).toBe(text);
          expect(getCommentAnchorsFromDoc(editor.state().doc)).toEqual([]);
          expect(
            editor.session.document.package.document.comments ?? [],
          ).toEqual([]);
        }
        expect(
          direct.publish(direct.session.prepareUndo(direct.state()).unwrap()),
        ).toBe(true);
        expect(direct.session.document).toEqual(before);
      }
    });

    test(`suggested/${text === "" ? "empty" : "nonempty"}: comment alignment refusal preserves the live document and journal`, () => {
      for (const type of REPLACEMENT_TYPES) {
        const { editor, firstComment } = commentedReplacementSource();
        const before = editor.session.document;
        const version = editor.session.version;
        const selection = editor.session.selection;
        const save = editor.session.captureSaveSnapshot();
        let committed = 0;
        let dropped = 0;
        editor.session.onCommitted(() => committed++);
        editor.session.onHistoryDrop(() => dropped++);
        const operation = wholeBlockReplacement({ type, text });
        const result = editor.apply([operation], { mode: "suggested" });
        expect(result.applied).toEqual([]);
        expect(result.undoHandle).toBeNull();
        expect(result.skipped).toMatchObject([
          {
            reason: "unsupportedBlock",
            canonicalRefusal: { gap: "comment-replacement-alignment" },
          },
        ]);
        expect(editor.session.document).toBe(before);
        expect(editor.session.version).toBe(version);
        expect(editor.session.selection).toEqual(selection);
        expect(editor.session.captureSaveSnapshot()).toEqual(save);
        expect(editor.session.canUndo).toBe(true);
        expect(editor.session.canRedo).toBe(false);
        expect(committed).toBe(0);
        expect(dropped).toBe(0);
        expect(
          editor.publish(editor.session.prepareUndo(editor.state()).unwrap()),
        ).toBe(true);
        expect(editor.session.document).toEqual(firstComment);
      }
    });
  }

  for (const mode of ["editing", "suggesting"] as const) {
    test(`${mode}: exact covered-edge native replacement takes the comment with its words`, () => {
      const editor = setup("The Supplier delivers on time.");
      expect(
        editor.apply([
          {
            id: "comment",
            type: "commentOnBlock",
            blockId: "12345678",
            quote: "time",
            comment: { text: "Note" },
          },
        ]).applied,
      ).toHaveLength(1);
      editor.session.setMode(
        mode === "editing"
          ? { type: "editing" }
          : { type: "suggesting", author: "Reviewer" },
      );
      let from: number | undefined;
      editor.state().doc.descendants((node, position) => {
        if (node.isText && node.text === "time") from = position;
      });
      if (from === undefined) panic("Source commented text is missing");
      const prepared = editor.session.prepareReplace(editor.state(), {
        from,
        to: from + 4,
        text: "schedule",
      });
      expect(prepared.isOk()).toBe(true);
      expect(editor.publish(prepared.unwrap())).toBe(true);
      if (mode === "suggesting") {
        const accepted = editor.session
          .prepareResolve(editor.state(), {
            revisionIds: getTrackedChangesFromDoc(editor.state().doc).map(
              ({ id }) => id,
            ),
            resolution: "accept",
          })
          .unwrap();
        expect(editor.publish(accepted)).toBe(true);
      }
      // The new words go outside the range of the comment whose words they
      // replace, which leaves it nothing to cover.
      expect(
        getCommentAnchorsFromDoc(editor.state().doc).map(({ quote }) => quote),
      ).toEqual([]);
      expect(editor.state().doc.textContent).toBe(
        "The Supplier delivers on schedule.",
      );
    });

    test(`${mode}: native marker-inclusive whole-paragraph replacement ${mode === "editing" ? "takes" : "refuses"} even a uniform comment`, () => {
      const source = "The Supplier delivers on time.";
      const editor = setup(source);
      expect(
        editor.apply([
          {
            id: "comment",
            type: "commentOnBlock",
            blockId: "12345678",
            quote: source,
            comment: { text: "Note" },
          },
        ]).applied,
      ).toHaveLength(1);
      editor.session.setMode(
        mode === "editing"
          ? { type: "editing" }
          : { type: "suggesting", author: "Reviewer" },
      );
      const before = editor.session.document;
      const version = editor.session.version;
      const paragraph =
        editor.session.projection.firstParagraph ?? panic("Missing paragraph");
      const result = editor.session.prepareReplace(editor.state(), {
        from: paragraph.start,
        to: paragraph.start + paragraph.node.content.size,
        text: "New agreement.",
      });
      if (mode === "editing") {
        // Typed over directly, the comment goes with every word it covered.
        expect(editor.publish(result.unwrap())).toBe(true);
        expect(editor.state().doc.textContent).toBe("New agreement.");
        expect(getCommentAnchorsFromDoc(editor.state().doc)).toEqual([]);
        expect(editor.session.document.package.document.comments ?? []).toEqual(
          [],
        );
        return;
      }
      expect(result.isErr()).toBe(true);
      if (result.isErr())
        expect(result.error.ground).toEqual({
          gap: "comment-replacement-alignment",
        });
      expect(editor.session.document).toBe(before);
      expect(editor.session.version).toBe(version);
    });

    for (const action of ["insert", "delete"] as const) {
      test(`${mode}: native ${action} retains its existing semantics beside comments`, () => {
        const editor = setup(
          "The Supplier delivers the goods on time and in good order.",
        );
        expect(
          editor.apply([
            {
              id: "comment",
              type: "commentOnBlock",
              blockId: "12345678",
              quote: "time",
              comment: { text: "Note" },
            },
          ]).applied,
        ).toHaveLength(1);
        editor.session.setMode(
          mode === "editing"
            ? { type: "editing" }
            : { type: "suggesting", author: "Reviewer" },
        );
        const before = editor.session.document;
        const paragraph =
          editor.session.projection.firstParagraph ??
          panic("Missing paragraph");
        const result = editor.session.prepareReplace(editor.state(), {
          from: paragraph.start,
          to:
            action === "insert"
              ? paragraph.start
              : paragraph.start + paragraph.node.content.size,
          text: action === "insert" ? "New " : "",
        });
        expect(result.isOk()).toBe(true);
        expect(editor.publish(result.unwrap())).toBe(true);
        expect(
          editor.publish(editor.session.prepareUndo(editor.state()).unwrap()),
        ).toBe(true);
        expect(editor.session.document).toEqual(before);
      });
    }
    test(`${mode}: native whole-paragraph replacement over a comment ${mode === "editing" ? "takes it with its words" : "refuses before publishing or changing history"}`, () => {
      const editor = setup(
        "The Supplier delivers the goods on time and in good order.",
      );
      expect(
        editor.apply([
          {
            id: "comment",
            type: "commentOnBlock",
            blockId: "12345678",
            quote: "time",
            comment: { text: "Note" },
          },
        ]).applied,
      ).toHaveLength(1);
      editor.session.setMode(
        mode === "editing"
          ? { type: "editing" }
          : { type: "suggesting", author: "Reviewer" },
      );
      const before = editor.session.document;
      const version = editor.session.version;
      const save = editor.session.captureSaveSnapshot();
      let committed = 0;
      editor.session.onCommitted(() => committed++);
      const paragraph =
        editor.session.projection.firstParagraph ?? panic("Missing paragraph");
      const result = editor.session.prepareReplace(editor.state(), {
        from: paragraph.start,
        to: paragraph.start + paragraph.node.content.size,
        text: "Invoice delivery payment agreement supplier buyer payment term.",
      });
      if (mode === "editing") {
        expect(editor.publish(result.unwrap())).toBe(true);
        expect(committed).toBe(1);
        expect(getCommentAnchorsFromDoc(editor.state().doc)).toEqual([]);
        expect(editor.session.document.package.document.comments ?? []).toEqual(
          [],
        );
        expect(
          editor.publish(editor.session.prepareUndo(editor.state()).unwrap()),
        ).toBe(true);
        expect(editor.session.document).toEqual(before);
        return;
      }
      expect(result.isErr()).toBe(true);
      if (result.isErr())
        expect(result.error.ground).toEqual({
          gap: "comment-replacement-alignment",
        });
      expect(editor.session.document).toBe(before);
      expect(editor.session.version).toBe(version);
      expect(editor.session.captureSaveSnapshot()).toEqual(save);
      expect(committed).toBe(0);
      expect(editor.session.canUndo).toBe(true);
      expect(editor.session.canRedo).toBe(false);
      expect(
        editor.publish(editor.session.prepareUndo(editor.state()).unwrap()),
      ).toBe(true);
      expect(editor.session.document.package.document.comments ?? []).toEqual(
        [],
      );
    });
  }

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
        { operationId: "merge", code: "unsupportedBlock" },
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
      [
        operation,
        {
          id: "table",
          type: "insertTableColumn",
          blockId: "12345678",
          position: "after",
        },
      ],
      { atomic: true },
    );
    // The block is in no table, so the batch is rejected whole.
    expect(refused.status).toBe("rejected");
    expect(
      refused.issues.find(({ operationId }) => operationId === "table")?.code,
    ).toBe("unsupportedBlock");
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
  // A row, column or cell operation on a block in no table is skipped as
  // the default engine skips it, under no gap.
  for (const outsideOperation of [
    {
      id: "row",
      type: "insertTableRow",
      blockId: "12345678",
      position: "after",
      cellTexts: ["cell"],
    },
    {
      id: "column",
      type: "insertTableColumn",
      blockId: "12345678",
      position: "after",
      cellTexts: ["cell"],
    },
    { id: "split", type: "splitTableCell", blockId: "12345678" },
    { id: "merge", type: "mergeTableCells", blockId: "12345678", rowCount: 2 },
  ] as const) {
    const outside = editor.apply([outsideOperation]);
    expect(outside.skipped.at(0)?.reason).toBe("unsupportedBlock");
    expect(outside.issues.at(0)?.canonicalRefusal).toBeUndefined();
  }
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

        if (text === "" && kind === "moveFrom")
          test("tracked deletion of the empty paragraph after a move source records nothing", () => {
            // The second paragraph is empty. The move owns the preceding
            // mark and does not add its words to the deletion's selection.
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
          test("deleting the second tracked: rejecting restores, accepting matches the direct deletion in review order", () => {
            const untouched = setup(documentOf(kind, text));
            const tracked = setup(documentOf(kind, text));
            expect(
              tracked.apply([deletion], { mode: "tracked-changes" }).applied,
            ).toHaveLength(1);
            expect(resolvedWords(tracked, "reject")).toBe(
              resolvedWords(untouched, "reject"),
            );
            const accepted = setup(documentOf(kind, text));
            expect(
              accepted.apply([deletion], { mode: "tracked-changes" }).applied,
            ).toHaveLength(1);
            const direct = setup(documentOf(kind, text));
            // Accept the pending join first so the deletion by the
            // destination's id selects the whole joined paragraph.
            if (kind === "del") resolvedWords(direct, "accept");
            expect(direct.apply([deletion]).applied).toHaveLength(1);
            expect(resolvedWords(accepted, "accept")).toBe(
              resolvedWords(direct, "accept"),
            );
          });
      });
});

describe("a text box paragraph", () => {
  // A text box is a story of its own, its paragraphs named through the body.
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
                          {
                            type: "paragraph",
                            paraId: "3456789A",
                            content: [
                              {
                                type: "run",
                                content: [
                                  { type: "text", text: "Second note." },
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
    {
      id: "commentOnBlock",
      type: "commentOnBlock",
      blockId: "34567890",
      quote: "note",
      comment: { text: "Check the box." },
    },
  ] as const satisfies readonly FolioDocumentOperation[];

  /** The words of the box's paragraphs, deleted ones included. */
  const boxWords = (current: Document): string[] => {
    const host = current.package.document.content[0];
    const run = host?.type === "paragraph" ? host.content[1] : undefined;
    const shape = run?.type === "run" ? run.content[0] : undefined;
    if (shape?.type !== "shape" || shape.shape.textBody === undefined)
      return panic("The host still draws the box.");
    return shape.shape.textBody.content.map(
      (block) =>
        JSON.stringify(block)
          .match(/"text":"[^"]*"/gu)
          ?.map((text) => text.slice(8, -1))
          .join("") ?? "",
    );
  };
  const direct: Record<(typeof operations)[number]["id"], string[]> = {
    insertAfterBlock: ["Boxed note.", "Added note.", "Second note."],
    insertBeforeBlock: ["Added note.", "Boxed note.", "Second note."],
    replaceInBlock: ["Framed note.", "Second note."],
    deleteBlock: ["Second note."],
    commentOnBlock: ["Boxed note.", "Second note."],
  };

  for (const mode of ["direct", "tracked-changes"] as const)
    for (const operation of operations)
      test(`${operation.type} acts in the box, ${mode}`, () => {
        const editor = setup(document);
        expect(
          editor.snapshot().blocks.some(({ id }) => id === "34567890"),
        ).toBe(true);
        const original = editor.session.document;
        const result = editor.apply([operation], { mode });
        expect(result.skipped).toEqual([]);
        expect(result.applied).toHaveLength(1);
        const changed = editor.session.document;
        // The body keeps its other paragraphs as they were.
        expect(changed.package.document.content[1]).toBe(
          original.package.document.content[1],
        );
        if (mode === "direct")
          expect(boxWords(changed)).toEqual(direct[operation.id]);
        else
          expect(changed.package.document.content[0]).not.toEqual(
            original.package.document.content[0],
          );
      });

  // The batch runs its comments first: an edit after a comment it adds in the
  // box lands past that comment's reference, as it does applied after it.
  for (const mode of ["direct", "tracked-changes"] as const)
    test(`an edit after a comment of its batch in the box lands as it would after that comment, ${mode}`, () => {
      const comment = {
        id: "comment",
        type: "commentOnBlock",
        blockId: "34567890",
        quote: "Boxed",
        comment: { text: "Check the box." },
      } as const satisfies FolioDocumentOperation;
      const edit = {
        id: "edit",
        type: "replaceInBlock",
        blockId: "34567890",
        find: "note",
        replace: "line",
      } as const satisfies FolioDocumentOperation;
      const batched = setup(document);
      expect(
        batched
          .apply([edit, comment], { mode })
          .applied.map(({ id }) => id)
          .toSorted(),
      ).toEqual(["comment", "edit"]);
      const sequential = setup(document);
      sequential.apply([comment], { mode });
      sequential.apply([edit], { mode });
      expect(boxWords(batched.session.document)).toEqual(
        boxWords(sequential.session.document),
      );
      if (mode === "direct")
        expect(boxWords(batched.session.document)).toEqual([
          "Boxed line.",
          "Second note.",
        ]);
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

/** Every painted text run's comment ids, table cells included. */
const paintedCommentIds = (blocks: readonly FlowBlock[]): number[][] =>
  blocks.flatMap((block) => {
    switch (block.kind) {
      case "paragraph":
        return block.runs.flatMap((run) =>
          run.kind === "text" && run.commentIds !== undefined
            ? [run.commentIds]
            : [],
        );
      case "table":
        return block.rows.flatMap((row) =>
          row.cells.flatMap((cell) => paintedCommentIds(cell.blocks)),
        );
      default:
        return [];
    }
  });

const tableDocument = (): Document => ({
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
                          content: [{ type: "text", text: "alpha beta" }],
                        },
                      ],
                    },
                  ],
                },
              ],
            },
          ],
        },
        { type: "paragraph", paraId: "23456789", content: [] },
      ],
    },
  },
});

describe("the first comment in a document without any", () => {
  const range =
    createFolioAITextRangeHandle({
      blockId: "12345678",
      text: "alpha beta",
      startOffset: 0,
      endOffset: 5,
    }) ?? panic("Missing range");
  const operations = [
    {
      id: "block",
      type: "commentOnBlock",
      blockId: "12345678",
      comment: { text: "note" },
    },
    {
      id: "range",
      type: "commentOnRange",
      range,
      comment: { text: "note" },
    },
  ] as const satisfies readonly FolioDocumentOperation[];

  for (const [where, source] of [
    ["the body", "alpha beta"],
    ["a table cell", tableDocument()],
  ] as const) {
    for (const operation of operations) {
      test(`${operation.type} in ${where} is comment 0 and is painted through undo, redo and save`, async () => {
        const editor = setup(source);
        const result = editor.apply([operation]);
        expect(result.status).toBe("committed");
        expect(result.applied.at(0)?.commentId).toBe(0);
        expect(paintedCommentIds(toFlowBlocks(editor.state().doc))).toEqual([
          [0],
        ]);

        editor.executor.undo(result.undoHandle ?? panic("Missing handle"));
        expect(paintedCommentIds(toFlowBlocks(editor.state().doc))).toEqual([]);
        const redone = editor.session.prepareRedo(editor.state());
        if (redone.isErr()) return panic(redone.error.message);
        expect(editor.publish(redone.value)).toBe(true);
        expect(paintedCommentIds(toFlowBlocks(editor.state().doc))).toEqual([
          [0],
        ]);

        const reopened = await parseDocx(
          await createDocx(editor.session.document),
          { preloadFonts: false },
        );
        expect(reopened.package.document.comments?.map(({ id }) => id)).toEqual(
          [0],
        );
        expect(paintedCommentIds(toFlowBlocks(toProseDoc(reopened)))).toEqual([
          [0],
        ]);
      });
    }
  }
});
