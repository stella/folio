import { expect, setDefaultTimeout, test } from "bun:test";
import fc from "fast-check";
import { EditorState, TextSelection } from "prosemirror-state";

import { DOCUMENT_OP_TYPES, OP_STORIES } from "@stll/docx-core/ops";

import {
  assertProperty,
  propertyTestTimeout,
} from "../../../../test/property-testing";
import { createCanonicalSession } from "../canonical/canonicalSession";
import type { Document, ParagraphContent } from "../types/document";
import {
  classifyAutocompleteEdit,
  createAutocompleteSuggestion,
} from "./autocompleteSuggestion";

setDefaultTimeout(propertyTestTimeout(30_000));

const BLOCK_ID = "12345678";
const kinds = [
  "plain",
  "bold",
  "tab",
  "bookmarks",
  "insertion",
  "control",
  "nested",
] as const;
const word = fc
  .array(fc.constantFrom("a", "b", "c"), { minLength: 1, maxLength: 5 })
  .map((letters) => letters.join(""));

test("only the exact published typing batch consumes a ghost across source carriers", () => {
  const exercised = new Set<string>();
  const seamCounts = new Set<number>();
  assertProperty(
    fc.property(word, word, (left, right) => {
      for (const kind of kinds)
        for (const suggesting of [false, true]) {
          const run = {
            type: "run",
            content: [{ type: "text", text: left + right }],
          } as const;
          const content: ParagraphContent[] = [];
          switch (kind) {
            case "plain":
              content.push(run);
              break;
            case "bold":
              content.push({ ...run, formatting: { bold: true } });
              break;
            case "tab":
              content.push({ type: "run", content: [{ type: "tab" }] }, run);
              break;
            case "bookmarks":
              content.push(
                { type: "run", content: [{ type: "text", text: left }] },
                { type: "bookmarkStart", id: 1, name: "_Ref" },
                { type: "bookmarkEnd", id: 1 },
                { type: "run", content: [{ type: "text", text: right }] },
              );
              break;
            case "control":
              content.push({
                type: "inlineSdt",
                properties: { sdtType: "richText", id: 7 },
                content: [run],
              });
              break;
            case "nested":
              content.push({
                type: "inlineSdt",
                properties: { sdtType: "richText", id: 7 },
                content: [
                  {
                    type: "inlineWrapper",
                    kind: "smartTag",
                    element: "term",
                    content: [run],
                  },
                ],
              });
              break;
            case "insertion":
              content.push({
                type: "insertion",
                info: {
                  id: 400,
                  author: "Reviewer",
                  date: "2026-10-02T10:00:00Z",
                },
                content: [run],
              });
              break;
          }
          const source: Document = {
            package: {
              document: {
                content: [{ type: "paragraph", paraId: BLOCK_ID, content }],
              },
            },
          };
          const length = left.length + right.length + Number(kind === "tab");
          for (const offset of new Set([0, 1, left.length, length])) {
            const ordinals =
              kind === "bookmarks" && offset === left.length
                ? [0, 1, 2]
                : [undefined];
            for (const zeroWidthBefore of ordinals) {
              const session = createCanonicalSession(source).unwrap();
              if (suggesting)
                session.setMode({ type: "suggesting", author: "Reviewer" });
              const address = {
                story: OP_STORIES.MAIN,
                blockId: BLOCK_ID,
                offset,
                ...(zeroWidthBefore === undefined ? {} : { zeroWidthBefore }),
              };
              const position = session.projection.positionAt(address).unwrap();
              const state = EditorState.create({
                doc: session.projection.doc,
                selection: TextSelection.create(
                  session.projection.doc,
                  position,
                ),
              });
              const owner = createAutocompleteSuggestion();
              owner.start(address, "request");
              owner.append("request", "xyz");
              let published = false;
              session.onCommitted((batch) => {
                published = true;
                const result = classifyAutocompleteEdit(batch);
                expect(result).toMatchObject({
                  type: "typed",
                  text: "x",
                  at: { story: "main", blockId: BLOCK_ID, offset },
                });
                // Operation transport/reconstruction cannot manufacture input meaning.
                expect(classifyAutocompleteEdit({ ...batch })).toEqual({
                  type: "other",
                });
                expect(
                  classifyAutocompleteEdit(structuredClone(batch)),
                ).toEqual({ type: "other" });
                expect(
                  classifyAutocompleteEdit({ ...batch, origin: "remote" }),
                ).toEqual({
                  type: "moved",
                  ops: batch.ops,
                });
                owner.noteEdit(result);
                expect(owner.get()).toMatchObject({
                  text: "yz",
                  anchor: { offset: offset + 1 },
                });
              });
              const commit = session
                .prepareReplace(state, {
                  from: position,
                  to: position,
                  text: "x",
                  semantic: "typing",
                })
                .unwrap();
              expect(published).toBe(false);
              state.apply(commit.transaction);
              commit.publish().unwrap();
              expect(published).toBe(true);
              exercised.add(kind);
              seamCounts.add(
                commit.ops.filter(
                  (op) => op.type === DOCUMENT_OP_TYPES.JOIN_INLINE,
                ).length,
              );
            }
          }
        }
    }),
    { numRuns: 12, id: "autocomplete-typing-normalization" },
  );
  expect(exercised).toEqual(new Set(kinds));
  expect(seamCounts).toEqual(new Set([0, 1, 2]));
});

test("paste, replacement, ranged typing, undo and redo dismiss rather than consume", () => {
  for (const semantic of ["paste", "replacement", "typing"] as const) {
    const session = createCanonicalSession({
      package: {
        document: {
          content: [
            {
              type: "paragraph",
              paraId: BLOCK_ID,
              content: [
                { type: "run", content: [{ type: "text", text: "abc" }] },
              ],
            },
          ],
        },
      },
    }).unwrap();
    let state = EditorState.create({ doc: session.projection.doc });
    const owner = createAutocompleteSuggestion();
    const start = () => {
      owner.start(
        { story: OP_STORIES.MAIN, blockId: BLOCK_ID, offset: 1 },
        "request",
      );
      owner.append("request", "xyz");
    };
    session.onCommitted((batch) => {
      expect(classifyAutocompleteEdit(batch)).toEqual({ type: "other" });
      owner.noteEdit(classifyAutocompleteEdit(batch));
      expect(owner.get()).toEqual({ status: "idle" });
    });
    start();
    const commit = session
      .prepareReplace(state, {
        from: 2,
        to: semantic === "typing" ? 3 : 2,
        text: "x",
        semantic,
      })
      .unwrap();
    state = state.apply(commit.transaction);
    commit.publish().unwrap();
    start();
    const undo = session.prepareUndo(state).unwrap();
    state = state.apply(undo.transaction);
    undo.publish().unwrap();
    start();
    const redo = session.prepareRedo(state).unwrap();
    state = state.apply(redo.transaction);
    redo.publish().unwrap();
    expect(state.doc.textContent).toBe(semantic === "typing" ? "axc" : "axbc");
  }
});
