import { describe, expect, test } from "bun:test";
import fc from "fast-check";

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
import { rejectAllChanges } from "../prosemirror/commands/comments";
import { fromProseDoc } from "../prosemirror/conversion/fromProseDoc";
import type { Document, Paragraph } from "../types/document";
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

const rejectInEditor = (document: Document) => {
  const state = createHarnessState(document, "editing");
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

  test(
    "editor and engine reject generated insertions holding every boundary kind alike",
    async () => {
      const run = (text: string): Paragraph["content"][number] => ({
        type: "run",
        content: [{ type: "text", text }],
      });
      const move = fc.integer({ min: 1, max: 3 }).chain((id) =>
        fc.constantFrom<Paragraph["content"]>(
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
      await assertProperty(
        fc.property(
          moves,
          moves,
          moves,
          bookmark,
          (before, inside, after, placement) => {
            const start = {
              type: "bookmarkStart" as const,
              id: 9,
              name: "_Ref9",
            };
            const end = { type: "bookmarkEnd" as const, id: 9 };
            const document = documentWith([
              run("a"),
              ...(placement === "outside" || placement === "startOutside"
                ? [start]
                : []),
              ...before,
              ...(placement === "outside" ? [end] : []),
              {
                type: "insertion",
                info: insertion,
                content: [
                  run("x"),
                  ...(placement === "inside" || placement === "endOutside"
                    ? [start]
                    : []),
                  ...inside,
                  ...(placement === "inside" || placement === "startOutside"
                    ? [end]
                    : []),
                  run("y"),
                ],
              },
              run("b"),
              ...(placement === "endOutside" ? [end] : []),
              ...after,
              run("c"),
            ]);
            let editor: string[];
            try {
              editor = outline(rejectInEditor(document));
            } catch {
              // A document the editor refuses on import holds nothing to reject.
              fc.pre(false);
              return;
            }
            expect(editor).toEqual(outline(rejectInEngine(document)));
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
