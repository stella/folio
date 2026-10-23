import { describe, expect, test } from "bun:test";
import fc from "fast-check";
import type { EditorState } from "prosemirror-state";

import {
  applyDocumentOp,
  DOCUMENT_OP_TYPES,
  OP_STORIES,
  REVISION_DECISIONS,
} from "@stll/docx-core/ops";

import {
  assertProperty,
  propertyTestTimeout,
} from "../../../../test/property-testing";
import { createHarnessState } from "../__tests__/editorHarness";
import {
  rejectAIEditRevision,
  rejectAllChanges,
} from "../prosemirror/commands/comments";
import { fromProseDoc } from "../prosemirror/conversion/fromProseDoc";
import type { Document, InlineSdt, Paragraph } from "../types/document";
import { createEmptyDocument } from "../utils/createDocument";

// Rejecting an insertion keeps the range boundaries it held where the
// insertion was, in the editor as in the engine.

const insertion = { id: 3, author: "A", date: "2026-03-04T05:00:03Z" };

const documentWith = (content: Paragraph["content"]): Document => {
  const document = createEmptyDocument();
  document.package.document.content = [
    { type: "paragraph", paraId: "00000001", content },
  ];
  return document;
};

/** The paragraph as its text and the range boundaries between it, in order. */
const outline = (content: readonly unknown[]): string[] => {
  const tokens: string[] = [];
  for (const token of JSON.stringify(content).match(
    /"type":"(?:bookmarkStart|bookmarkEnd|moveToRangeStart|moveToRangeEnd)","id":\d+|"text":"[^"]*"/gu,
  ) ?? []) {
    const text = /^"text":"([^"]*)"$/u.exec(token)?.[1];
    const previous = tokens.at(-1);
    if (text !== undefined && previous?.startsWith("text:") === true)
      tokens[tokens.length - 1] = `${previous}${text}`;
    else
      tokens.push(
        text === undefined ? token.replaceAll('"', "") : `text:${text}`,
      );
  }
  return tokens;
};

const firstParagraphContent = (document: Document) => {
  const first = document.package.document.content[0];
  return first?.type === "paragraph" ? first.content : [];
};

const rejectInState = (state: EditorState, document: Document) => {
  let resolved = state;
  expect(
    rejectAllChanges()(state, (tr) => {
      resolved = state.apply(tr);
    }),
  ).toBe(true);
  return firstParagraphContent(
    fromProseDoc(resolved.doc, document, {
      stylesheetSource: { type: "package" },
    }),
  );
};

const rejectInEditor = (document: Document) =>
  rejectInState(createHarnessState(document, "editing"), document);

const rejectInEngine = (document: Document) => {
  const applied = applyDocumentOp(document, {
    type: DOCUMENT_OP_TYPES.RESOLVE_REVISION,
    story: OP_STORIES.MAIN,
    revisionIds: [insertion.id],
    decision: REVISION_DECISIONS.REJECT,
  });
  if (applied.isErr()) throw applied.error;
  return firstParagraphContent(applied.value.document);
};

describe("rejecting an insertion keeps its range boundaries", () => {
  test("a bookmark inside inserted text stays where the insertion was, in both", () => {
    const document = documentWith([
      { type: "run", content: [{ type: "text", text: "a" }] },
      {
        type: "insertion",
        info: insertion,
        content: [
          { type: "run", content: [{ type: "text", text: "x" }] },
          { type: "bookmarkStart", id: 5, name: "_Ref5" },
          { type: "bookmarkEnd", id: 5 },
          { type: "run", content: [{ type: "text", text: "y" }] },
        ],
      },
      { type: "run", content: [{ type: "text", text: "b" }] },
    ]);
    const editor = outline(rejectInEditor(document));
    expect(editor).toEqual(outline(rejectInEngine(document)));
    expect(editor).toEqual([
      "text:a",
      "type:bookmarkStart,id:5",
      "type:bookmarkEnd,id:5",
      "text:b",
    ]);
  });

  test("an inserted move range end before its existing start moves after it, in both", () => {
    const document = documentWith([
      { type: "run", content: [{ type: "text", text: "one " }] },
      {
        type: "insertion",
        info: insertion,
        content: [{ type: "moveToRangeEnd", id: 6 }],
      },
      { type: "run", content: [{ type: "text", text: "two " }] },
      { type: "moveToRangeStart", id: 6, name: "move6", author: "A" },
      { type: "run", content: [{ type: "text", text: "three" }] },
    ]);
    const editor = outline(rejectInEditor(document));
    expect(editor).toEqual(outline(rejectInEngine(document)));
    expect(editor).toEqual([
      "text:one two ",
      "type:moveToRangeStart,id:6",
      "type:moveToRangeEnd,id:6",
      "text:three",
    ]);
  });

  test("an inserted copy of an existing move range start goes, in both", () => {
    const document = documentWith([
      {
        type: "insertion",
        info: insertion,
        content: [
          { type: "moveToRangeStart", id: 7, name: "move7", author: "A" },
        ],
      },
      { type: "moveToRangeStart", id: 7, name: "move7", author: "A" },
      { type: "run", content: [{ type: "text", text: "two" }] },
      { type: "moveToRangeEnd", id: 7 },
    ]);
    const editor = outline(rejectInEditor(document));
    expect(editor).toEqual(outline(rejectInEngine(document)));
    expect(editor).toEqual([
      "type:moveToRangeStart,id:7",
      "text:two",
      "type:moveToRangeEnd,id:7",
    ]);
  });

  test("rejecting an insertion inside a deletion keeps its bookmark inside the deletion, in both", () => {
    const deletion = { id: 20, author: "B", date: "2026-03-04T05:00:04Z" };
    const document = documentWith([
      { type: "run", content: [{ type: "text", text: "a" }] },
      {
        type: "deletion",
        info: deletion,
        content: [
          { type: "run", content: [{ type: "text", text: "d" }] },
          {
            type: "insertion",
            info: insertion,
            content: [
              { type: "run", content: [{ type: "text", text: "x" }] },
              { type: "bookmarkStart", id: 5, name: "_Ref5" },
              { type: "bookmarkEnd", id: 5 },
            ],
          },
          { type: "run", content: [{ type: "text", text: "e" }] },
        ],
      },
      { type: "run", content: [{ type: "text", text: "b" }] },
    ]);
    const state = createHarnessState(document, "editing");
    let resolved = state;
    expect(
      rejectAIEditRevision(insertion.id)(state, (tr) => {
        resolved = state.apply(tr);
      }),
    ).toBe(true);
    const editor = firstParagraphContent(
      fromProseDoc(resolved.doc, document, {
        stylesheetSource: { type: "package" },
      }),
    );
    /** Each deletion's text and range boundaries, in order. */
    const deleted = (content: Paragraph["content"]) =>
      content.flatMap((item) =>
        item.type === "deletion" ? [outline(item.content)] : [],
      );
    expect(deleted(editor)).toEqual(deleted(rejectInEngine(document)));
    expect(deleted(editor)).toEqual([
      ["text:d", "type:bookmarkStart,id:5", "type:bookmarkEnd,id:5", "text:e"],
    ]);
  });

  test("dropping an inserted copy leaves the next range alone, in both", () => {
    // The copy sits right before another range's end, which comes before its
    // own start; only the range the rejection kept a boundary of is repaired.
    const document = documentWith([
      {
        type: "insertion",
        info: insertion,
        content: [
          { type: "moveToRangeStart", id: 7, name: "move7", author: "A" },
        ],
      },
      { type: "moveToRangeEnd", id: 8 },
      { type: "run", content: [{ type: "text", text: "one" }] },
      { type: "moveToRangeStart", id: 8, name: "move8", author: "A" },
      { type: "moveToRangeStart", id: 7, name: "move7", author: "A" },
      { type: "run", content: [{ type: "text", text: "two" }] },
      { type: "moveToRangeEnd", id: 7 },
    ]);
    const editor = outline(rejectInEditor(document));
    expect(editor).toEqual(outline(rejectInEngine(document)));
    expect(editor).toEqual([
      "type:moveToRangeEnd,id:8",
      "text:one",
      "type:moveToRangeStart,id:8",
      "type:moveToRangeStart,id:7",
      "text:two",
      "type:moveToRangeEnd,id:7",
    ]);
  });

  test(
    "editor and engine reject generated insertions holding every boundary kind alike",
    async () => {
      const run = (text: string) => ({
        type: "run" as const,
        content: [{ type: "text" as const, text }],
      });
      const move = fc.integer({ min: 1, max: 3 }).chain((id) =>
        fc.constantFrom<InlineSdt["content"]>(
          [{ type: "moveToRangeStart", id, name: `move${id}`, author: "A" }],
          [{ type: "moveToRangeEnd", id }],
          // A point range: the editor holds both halves in one node.
          [
            { type: "moveToRangeStart", id, name: `move${id}`, author: "A" },
            { type: "moveToRangeEnd", id },
          ],
        ),
      );
      const moves = fc
        .array(move, { maxLength: 3 })
        .map((items) => items.flat());
      // One bookmark, paired and unique as the editor requires, in any of the
      // places a rejection treats differently.
      const bookmark = fc.constantFrom(
        "none",
        "outside",
        "inside",
        "startOutside",
        "endOutside",
      );
      // Where inline content controls stand: none, one holding what the
      // insertion holds (the rejection removes it whole), two nested so, or
      // one around the insertion (it stays, emptied).
      const controls = fc.constantFrom("none", "inside", "nested", "around");
      const control = (
        id: number,
        content: InlineSdt["content"],
      ): InlineSdt => ({
        type: "inlineSdt",
        properties: { sdtType: "richText", id },
        content: [run(`c${id}`), ...content],
      });
      await assertProperty(
        fc.property(
          moves,
          moves,
          moves,
          bookmark,
          controls,
          (before, inside, after, placement, wrapping) => {
            const start = {
              type: "bookmarkStart" as const,
              id: 9,
              name: "_Ref9",
            };
            const end = { type: "bookmarkEnd" as const, id: 9 };
            const held: InlineSdt["content"] = [
              ...(placement === "inside" || placement === "endOutside"
                ? [start]
                : []),
              ...inside,
              ...(placement === "inside" || placement === "startOutside"
                ? [end]
                : []),
            ];
            let inserted = held;
            if (wrapping === "inside") inserted = [control(41, held)];
            if (wrapping === "nested")
              inserted = [control(41, [control(42, held)])];
            const insertionNode: Paragraph["content"][number] = {
              type: "insertion",
              info: insertion,
              content: [run("x"), ...inserted, run("y")],
            };
            const document = documentWith([
              run("a"),
              ...(placement === "outside" || placement === "startOutside"
                ? [start]
                : []),
              ...before,
              ...(placement === "outside" ? [end] : []),
              wrapping === "around"
                ? control(43, [insertionNode])
                : insertionNode,
              run("b"),
              ...(placement === "endOutside" ? [end] : []),
              ...after,
              run("c"),
            ]);
            let state: EditorState;
            try {
              state = createHarnessState(document, "editing");
            } catch {
              // A document the editor refuses on import holds nothing to reject.
              fc.pre(false);
              return;
            }
            // Both reject the document the editor holds: its import may
            // already have joined a range's two halves into one point.
            const imported = fromProseDoc(state.doc, document, {
              stylesheetSource: { type: "package" },
            });
            expect(outline(rejectInState(state, document))).toEqual(
              outline(rejectInEngine(imported)),
            );
          },
        ),
        {
          numRuns: 200,
          id: "editor and engine reject generated insertions holding every boundary kind alike",
        },
      );
    },
    propertyTestTimeout(60_000),
  );
});
