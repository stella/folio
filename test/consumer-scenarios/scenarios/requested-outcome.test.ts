/**
 * The requested-outcome oracle (support/oracle.ts) over the batches most
 * likely to break it: collisions of two operations on one block, edits next
 * to surrogate pairs, table payloads, and edits of pending tracked changes
 * after a save. An `applied` receipt must mean the saved document is what
 * was asked; a refusal must leave it as it was.
 */

import assert from "node:assert/strict";
import { describe, test } from "node:test";

import { fromMarkdown } from "@stll/folio-core/markdown";
import {
  FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
  hashFolioAIBlockText,
} from "@stll/folio-core/server";

import { sequentialGroups } from "../support/batch-replay-order.ts";
import {
  COLLISION_FIXTURES,
  FIXTURES,
  openReviewer,
  packDocument,
  plainDocument,
} from "../support/documents.ts";
import {
  assertHealthy,
  saveAndReopen,
  visibleState,
} from "../support/invariants.ts";
import {
  expectedFailure,
  KNOWN_FAILING_COLLISIONS,
  KNOWN_FAILING_FOLLOW_UPS,
} from "../support/known-issues.ts";
import { captureLinks, comparePreservedLinks } from "../support/link-oracle.ts";
import {
  type AppliedOperation,
  type Block,
  COLLISIONS,
  coreBatch,
  type Mode,
  MODES,
  type Operation,
} from "../support/operations.ts";
import {
  applyChecked,
  assertRequestedOutcome,
  capture,
  compareComments,
  compareWithModel,
  deletionKeptBlock,
  EXPECTATIONS,
  expectOperation,
  KEEPS_NOTE_REFERENCES,
  modelOf,
  OPERATION_TYPES,
  retireNoteReferences,
  type Row,
} from "../support/oracle.ts";
import { compareTableGeometry } from "../support/table-oracle.ts";

type Reviewer = Awaited<ReturnType<typeof openReviewer>>;

const ALL_FIXTURES = { ...FIXTURES, ...COLLISION_FIXTURES };
const blocksOf = (reviewer: Reviewer): Block[] =>
  reviewer.getContent() as Block[];

describe("the requested-outcome oracle", () => {
  test("decides, for every public operation type, what an applied receipt means", () => {
    assert.deepEqual(
      Object.keys(EXPECTATIONS).sort(),
      [...OPERATION_TYPES].sort(),
      "an operation type without an entry in EXPECTATIONS: model it, or mark it `null` (no expectation yet)",
    );
  });

  const row = (id: string, text: string, extra: Partial<Row> = {}): Row => ({
    id,
    text,
    kind: "paragraph",
    ...extra,
  });
  const before = [
    row("a", "First clause."),
    row("b", "Second clause."),
    row("c", "Third."),
  ];
  const expecting = (operations: Operation[]) => {
    const model = modelOf(before);
    for (const operation of operations) expectOperation(model, operation);
    return model;
  };

  test("range formatting exempts only the paragraph it changed from style checks", () => {
    const target = row("target", "😀xx", {
      styleId: "Normal",
      previewRuns: [{ text: "😀xx", bold: false }],
    });
    const heading = row("heading", "Heading", {
      kind: "heading",
      headingLevel: 2,
      styleId: "Heading2",
      previewRuns: [{ text: "Heading", bold: false }],
    });
    const formatted = modelOf([target, heading]);
    formatted.mode = "tracked-changes";
    expectOperation(formatted, {
      type: "formatRange",
      range: {
        type: "textRange",
        story: "main",
        blockId: target.id,
        startOffset: 0,
        endOffset: 2,
      },
      formatting: { bold: true },
    });
    expectOperation(formatted, {
      type: "setBlockParagraphProperties",
      blockId: target.id,
      properties: { styleId: "Heading2" },
    });
    assert.deepEqual(
      compareWithModel(formatted, [
        row("target", "😀xx", {
          kind: "heading",
          headingLevel: 2,
          styleId: "Heading2",
          previewRuns: [
            { text: "😀", bold: true, directFormatting: { bold: true } },
            { text: "xx", bold: false },
          ],
        }),
        heading,
      ]),
      [],
    );

    const other = row("other", "x", {
      styleId: "Normal",
      previewRuns: [{ text: "x", bold: false }],
    });
    const unrelatedFormat = modelOf([target, other, heading]);
    unrelatedFormat.mode = "tracked-changes";
    expectOperation(unrelatedFormat, {
      type: "formatRange",
      range: {
        type: "textRange",
        story: "main",
        blockId: other.id,
        startOffset: 0,
        endOffset: 1,
      },
      formatting: { bold: true },
    });
    expectOperation(unrelatedFormat, {
      type: "setBlockParagraphProperties",
      blockId: target.id,
      properties: { styleId: "Heading2" },
    });
    assert.ok(
      compareWithModel(unrelatedFormat, [
        row("target", "😀xx", {
          kind: "heading",
          headingLevel: 2,
          styleId: "Heading2",
          previewRuns: [{ text: "😀xx", bold: true }],
        }),
        row("other", "x", {
          styleId: "Normal",
          previewRuns: [
            { text: "x", bold: true, directFormatting: { bold: true } },
          ],
        }),
        heading,
      ]).some((problem) =>
        problem.includes("effective bold differs from style Heading2"),
      ),
    );
  });

  test("notices a payload the engine dropped or put on the wrong block", () => {
    const replace = {
      type: "replaceInBlock",
      blockId: "b",
      find: "Second",
      replace: "Next",
    };
    assert.deepEqual(
      compareWithModel(expecting([replace]), [
        row("a", "First clause."),
        row("b", "Next clause."),
        row("c", "Third."),
      ]),
      [],
    );
    // The neighbour rewritten instead.
    assert.equal(
      compareWithModel(
        expecting([replace, { type: "deleteBlock", blockId: "b" }]),
        [row("a", "First clause."), row("b", "Next clause.")],
      ).length,
      1,
    );
    // One of two cell texts dropped.
    const cells = [
      row("t1", "Cell", { table: location(0) }),
      row("t2", "Cell", { table: location(1) }),
    ];
    const model = modelOf(cells);
    expectOperation(model, {
      type: "insertTableRow",
      blockId: "t1",
      cellTexts: ["New one", "New two"],
    });
    assert.equal(
      compareWithModel(model, [...cells, row("n", "New one")]).length,
      1,
    );
  });

  test("notices a paragraph style the request named but the result lacks", () => {
    const restyle = {
      type: "setBlockParagraphProperties",
      blockId: "a",
      properties: { styleId: "Heading2" },
    };
    assert.match(
      compareWithModel(expecting([restyle]), before).join("\n"),
      /styleId is undefined, expected "Heading2"/u,
    );
  });

  test("inherits list properties from the live canonical insertion anchor", () => {
    const beforeRows = [row("anchor", "Anchor."), row("other", "Other.")];
    const properties = {
      id: "properties",
      type: "setBlockParagraphProperties",
      blockId: "anchor",
      properties: {
        alignment: "center",
        numbering: { kind: "newList", format: "numbered" },
      },
    };
    const insertion = {
      id: "insert",
      type: "insertBeforeBlock",
      blockId: "anchor",
      text: "Inserted.",
    };
    const check = (operations: AppliedOperation[], actual: Row[]) => {
      const model = modelOf(beforeRows);
      for (const group of sequentialGroups({
        applied: operations,
        preRows: beforeRows,
        mode: "direct",
      })) {
        for (const operation of group) expectOperation(model, operation);
      }
      return compareWithModel(model, actual);
    };
    const numbered = { numId: 100, level: 0 };

    assert.deepEqual(
      check(
        [properties, insertion],
        [
          row("inserted", "Inserted.", { kind: "paragraph" }),
          row("anchor", "Anchor.", {
            kind: "listItem",
            listLevel: 0,
            directAlignment: "center",
            listReference: numbered,
            displayLabel: "1.",
          }),
          beforeRows[1]!,
        ],
      ),
      [],
    );
    assert.deepEqual(
      check(
        [insertion, properties],
        [
          row("inserted", "Inserted.", {
            kind: "listItem",
            listLevel: 0,
            directAlignment: "center",
            listReference: numbered,
          }),
          row("anchor", "Anchor.", {
            kind: "listItem",
            listLevel: 0,
            directAlignment: "center",
            listReference: numbered,
            displayLabel: "1.",
          }),
          beforeRows[1]!,
        ],
      ),
      [],
    );

    const withoutInheritance = {
      ...insertion,
      inheritFormatting: false,
    };
    const explicitNone = {
      ...insertion,
      numbering: { kind: "none" },
    };
    assert.deepEqual(
      check(
        [withoutInheritance, properties],
        [
          row("inserted", "Inserted.", { kind: "paragraph" }),
          row("anchor", "Anchor.", {
            kind: "listItem",
            listLevel: 0,
            directAlignment: "center",
            listReference: numbered,
            displayLabel: "1.",
          }),
          beforeRows[1]!,
        ],
      ),
      [],
    );
    assert.deepEqual(
      check(
        [explicitNone, properties],
        [
          row("inserted", "Inserted.", {
            kind: "paragraph",
            directAlignment: "center",
          }),
          row("anchor", "Anchor.", {
            kind: "listItem",
            listLevel: 0,
            directAlignment: "center",
            listReference: numbered,
            displayLabel: "1.",
          }),
          beforeRows[1]!,
        ],
      ),
      [],
    );
    assert.match(
      check(
        [insertion, properties],
        [
          row("inserted", "Inserted.", { kind: "paragraph" }),
          row("anchor", "Anchor.", {
            kind: "listItem",
            listLevel: 0,
            directAlignment: "center",
            listReference: numbered,
            displayLabel: "1.",
          }),
          beforeRows[1]!,
        ],
      ).join("\n"),
      /list reference|kind is/u,
    );

    const laterProperties = {
      ...properties,
      id: "later-properties",
    };
    assert.deepEqual(
      check(
        [laterProperties, insertion, properties],
        [
          row("inserted", "Inserted.", {
            kind: "listItem",
            listLevel: 0,
            directAlignment: "center",
            listReference: { numId: 101, level: 0 },
          }),
          row("anchor", "Anchor.", {
            kind: "listItem",
            listLevel: 0,
            directAlignment: "center",
            listReference: { numId: 102, level: 0 },
            displayLabel: "1.",
          }),
          beforeRows[1]!,
        ],
      ),
      [],
    );
    assert.match(
      check(
        [laterProperties, insertion, properties],
        [
          row("inserted", "Inserted.", {
            kind: "listItem",
            listLevel: 0,
            directAlignment: "center",
            listReference: { numId: 102, level: 0 },
          }),
          row("anchor", "Anchor.", {
            kind: "listItem",
            listLevel: 0,
            directAlignment: "center",
            listReference: { numId: 102, level: 0 },
            displayLabel: "1.",
          }),
          beforeRows[1]!,
        ],
      ).join("\n"),
      /aliases another modeled list reference/u,
    );
    assert.deepEqual(
      check(
        [properties, { ...insertion, blockId: "other" }],
        [
          row("anchor", "Anchor.", {
            kind: "listItem",
            listLevel: 0,
            directAlignment: "center",
            listReference: numbered,
            displayLabel: "1.",
          }),
          row("inserted", "Inserted.", { kind: "paragraph" }),
          row("other", "Other."),
        ],
      ),
      [],
    );
  });

  test("keeps direct numbering through restyles and updates inherited numbering at execution time", () => {
    const makeModel = (styleId: string) => {
      const model = modelOf([
        row("anchor", "Anchor.", {
          styleId,
          ...(styleId === "ListStyle"
            ? {
                kind: "listItem",
                listLevel: 0,
                listReference: { numId: 40, level: 0 },
              }
            : {}),
          statedNumbering: { kind: "inherit" },
        }),
      ]);
      model.numberingFacts = {
        styles: new Map([
          ["ListStyle", { kind: "reference", numId: 40, ilvl: 0 }],
          ["PlainStyle", undefined],
        ]),
        direct: new Map([
          ["anchor", undefined],
          ["other", { kind: "none" }],
        ]),
        markerLevels: new Map([[40, new Set([0])]]),
      };
      return model;
    };
    // Operations at one position run in reverse input order, so the batch is
    // stated back to front to name the order they execute in.
    const applyOrdered = (
      model: ReturnType<typeof modelOf>,
      inExecutionOrder: AppliedOperation[],
    ) => {
      for (const group of sequentialGroups({
        applied: inExecutionOrder.toReversed(),
        preRows: [row("anchor", "Anchor.")],
        mode: "direct",
      })) {
        for (const operation of group) expectOperation(model, operation);
      }
      return model;
    };
    const newList = {
      id: "new-list",
      type: "setBlockParagraphProperties",
      blockId: "anchor",
      properties: { numbering: { kind: "newList", format: "numbered" } },
    };
    const restyle = {
      id: "restyle",
      type: "setBlockParagraphProperties",
      blockId: "anchor",
      properties: { styleId: "PlainStyle" },
    };
    const insert = {
      id: "insert",
      type: "insertBeforeBlock",
      blockId: "anchor",
      text: "Inserted.",
    };
    const directFirst = applyOrdered(makeModel("PlainStyle"), [
      newList,
      restyle,
      insert,
    ]);
    assert.deepEqual(
      compareWithModel(directFirst, [
        row("inserted", "Inserted.", {
          kind: "listItem",
          styleId: "PlainStyle",
          listLevel: 0,
          listReference: { numId: 100, level: 0 },
        }),
        row("anchor", "Anchor.", {
          kind: "listItem",
          styleId: "PlainStyle",
          listLevel: 0,
          listReference: { numId: 100, level: 0 },
        }),
      ]),
      [],
    );
    const insertFirst = applyOrdered(makeModel("PlainStyle"), [
      insert,
      restyle,
      newList,
    ]);
    assert.deepEqual(
      compareWithModel(insertFirst, [
        row("inserted", "Inserted.", {
          kind: "paragraph",
          styleId: "PlainStyle",
        }),
        row("anchor", "Anchor.", {
          kind: "listItem",
          styleId: "PlainStyle",
          listLevel: 0,
          listReference: { numId: 100, level: 0 },
        }),
      ]),
      [],
    );

    const styleDerived = applyOrdered(makeModel("ListStyle"), [
      {
        ...restyle,
        properties: { styleId: "PlainStyle" },
      },
      insert,
    ]);
    assert.deepEqual(
      compareWithModel(styleDerived, [
        row("inserted", "Inserted.", {
          kind: "paragraph",
          styleId: "PlainStyle",
        }),
        row("anchor", "Anchor.", {
          kind: "paragraph",
          styleId: "PlainStyle",
        }),
      ]),
      [],
    );

    const explicitNone = applyOrdered(makeModel("ListStyle"), [
      {
        ...newList,
        properties: { numbering: { kind: "none" } },
      },
      {
        ...restyle,
        properties: { styleId: "ListStyle" },
      },
      insert,
    ]);
    assert.deepEqual(
      compareWithModel(explicitNone, [
        row("inserted", "Inserted.", {
          kind: "paragraph",
          styleId: "ListStyle",
        }),
        row("anchor", "Anchor.", {
          kind: "paragraph",
          styleId: "ListStyle",
        }),
      ]),
      [],
    );
  });

  test("inherits from the named cell paragraph, not its table placement host", () => {
    const host = row("host", "Host.", {
      table: location(0),
      directAlignment: "right",
    });
    const named = row("named", "Named.", {
      table: location(1),
      directAlignment: "center",
    });
    const model = modelOf([host, named]);
    expectOperation(model, {
      type: "insertBeforeBlock",
      blockId: named.id,
      text: "Inserted.",
    });

    assert.deepEqual(
      compareWithModel(model, [
        row("inserted", "Inserted.", { directAlignment: "center" }),
        host,
        named,
      ]),
      [],
    );
    assert.match(
      compareWithModel(model, [
        row("inserted", "Inserted.", { directAlignment: "right" }),
        host,
        named,
      ]).join("\n"),
      /directAlignment is "right", expected "center"/u,
    );
  });

  test("notices a table merge that keeps the original grid", () => {
    const cells = [
      row("left", "Party", { table: location(0) }),
      row("right", "Role", { table: location(1) }),
    ];
    const model = modelOf(cells);
    expectOperation(model, {
      type: "mergeTableCells",
      blockId: "left",
      endBlockId: "right",
    });
    assert.deepEqual(model.tableGaps, []);
    assert.match(
      compareTableGeometry(model.tables, cells).join("\n"),
      /geometry/u,
    );
  });

  test("requires comments while any anchored character survives a text edit", () => {
    const text = "abcdef";
    for (let anchorStart = 0; anchorStart < text.length; anchorStart += 1) {
      for (
        let anchorEnd = anchorStart + 1;
        anchorEnd <= text.length;
        anchorEnd += 1
      ) {
        for (let editStart = 0; editStart < text.length; editStart += 1) {
          for (
            let editEnd = editStart + 1;
            editEnd <= text.length;
            editEnd += 1
          ) {
            const model = modelOf([row("b", text)]);
            expectOperation(model, {
              type: "replaceRange",
              range: {
                blockId: "b",
                startOffset: editStart,
                endOffset: editEnd,
              },
              replace: "",
            });
            const comment = {
              id: 1,
              text: "review",
              anchor: text.slice(anchorStart, anchorEnd),
              blockId: "b",
            };
            const removed = editStart <= anchorStart && editEnd >= anchorEnd;
            const problems = compareComments(
              model,
              [comment],
              [],
              [row("b", text.slice(0, editStart) + text.slice(editEnd))],
            );
            assert.equal(
              problems.length === 0,
              removed,
              JSON.stringify({ anchorStart, anchorEnd, editStart, editEnd }),
            );
            assert.equal(
              compareComments(
                model,
                [comment],
                [comment],
                [row("b", text.slice(0, editStart) + text.slice(editEnd))],
              ).length === 0,
              !removed,
              JSON.stringify({ anchorStart, anchorEnd, editStart, editEnd }),
            );
          }
        }
      }
    }
  });

  test("preserves a comment when a split or a different block changes its quoted text", () => {
    const comment = { id: 1, text: "review", anchor: "bcd", blockId: "b" };
    const split = modelOf([row("b", "abcdef")]);
    expectOperation(split, { type: "splitBlock", blockId: "b", offset: 3 });
    assert.equal(compareComments(split, [comment], [], []).length, 1);
    const elsewhere = modelOf([row("a", "bcd"), row("b", "abcdef")]);
    expectOperation(elsewhere, {
      type: "replaceRange",
      range: { blockId: "a", startOffset: 0, endOffset: 3 },
      replace: "",
    });
    assert.equal(compareComments(elsewhere, [comment], [], []).length, 1);
  });

  test("does not guess which repeated quote a comment anchored", () => {
    const comment = { id: 1, text: "review", anchor: "foo", blockId: "b" };
    const model = modelOf([row("b", "foo foo")]);
    expectOperation(model, {
      type: "replaceRange",
      range: { blockId: "b", startOffset: 0, endOffset: 3 },
      replace: "",
    });
    assert.deepEqual(compareComments(model, [comment], [], []), []);
    assert.deepEqual(
      compareComments(model, [comment], [comment], [row("b", " foo")]),
      [],
    );
  });

  test("a nonempty replacement carries a comment, while an empty one removes it", () => {
    const comment = { id: 1, text: "review", anchor: "Signed", blockId: "b" };
    const replaced = modelOf([row("b", "Signed")]);
    expectOperation(replaced, {
      type: "replaceBlock",
      blockId: "b",
      text: "Written",
    });
    assert.deepEqual(
      compareComments(
        replaced,
        [comment],
        [{ ...comment, anchor: "Written" }],
        [row("b", "Written")],
      ),
      [],
    );
    assert.match(
      compareComments(
        replaced,
        [comment],
        [{ ...comment, anchor: "Unrelated" }],
        [row("b", "Written")],
      ).join("\n"),
      /no comment/u,
    );
    assert.equal(
      compareComments(replaced, [comment], [], [row("b", "Written")]).length,
      1,
    );

    const emptied = modelOf([row("b", "Signed")]);
    expectOperation(emptied, { type: "replaceBlock", blockId: "b", text: "" });
    assert.deepEqual(compareComments(emptied, [comment], [], []), []);
    assert.equal(compareComments(emptied, [comment], [comment], []).length, 1);
  });

  test("a pending replacement may anchor both the old text and its insertion", () => {
    const comment = { id: 1, text: "review", anchor: "Old", blockId: "b" };
    const model = modelOf([row("b", "Old")]);
    expectOperation(model, { type: "replaceBlock", blockId: "b", text: "New" });

    assert.deepEqual(
      compareComments(
        model,
        [comment],
        [{ ...comment, anchor: "OldNew" }],
        [row("b", "New")],
        [comment],
        "suggested",
      ),
      [],
    );
    assert.match(
      compareComments(
        model,
        [comment],
        [{ ...comment, anchor: "New" }],
        [row("b", "New")],
        [comment],
        "suggested",
      ).join("\n"),
      /no comment/u,
    );
    assert.match(
      compareComments(
        model,
        [comment],
        [{ ...comment, anchor: "NewOld" }],
        [row("b", "New")],
        [comment],
        "suggested",
      ).join("\n"),
      /no comment/u,
    );
    assert.match(
      compareComments(
        model,
        [comment],
        [{ ...comment, anchor: "OldNew" }],
        [row("b", "New")],
        [comment],
        "direct",
      ).join("\n"),
      /no comment/u,
    );
    assert.match(
      compareComments(
        model,
        [comment],
        [],
        [row("b", "New")],
        [comment],
        "suggested",
      ).join("\n"),
      /no comment/u,
    );

    const shared = { id: 2, text: "review", anchor: "The Buyer", blockId: "c" };
    const sharedModel = modelOf([row("c", "The Buyer")]);
    expectOperation(sharedModel, {
      type: "replaceBlock",
      blockId: "c",
      text: "The Seller",
    });
    assert.deepEqual(
      compareComments(
        sharedModel,
        [shared],
        [{ ...shared, anchor: "The BuyerSeller" }],
        [row("c", "The Seller")],
        [shared],
        "suggested",
      ),
      [],
    );
  });

  test("a deleted comment cannot satisfy a new comment with the same text", () => {
    const old = { id: 1, text: "same note", anchor: "Old", blockId: "a" };
    const model = modelOf([row("a", "Old"), row("b", "New")]);
    expectOperation(model, { type: "deleteBlock", blockId: "a" });
    expectOperation(model, {
      type: "commentOnBlock",
      blockId: "b",
      comment: { text: "same note" },
    });
    assert.deepEqual(
      compareComments(
        model,
        [old],
        [{ id: 2, text: "same note", anchor: "New", blockId: "b" }],
        [row("b", "New")],
      ),
      [],
    );
    const problems = compareComments(
      model,
      [old],
      [{ ...old, anchor: "New", blockId: "b" }],
      [row("b", "New")],
    );
    assert.match(problems.join("\n"), /no comment/u);
  });

  test("a comment retires with its removed note story", () => {
    const comment = {
      id: 1,
      text: "review",
      anchor: "Note text",
      blockId: "note",
    };
    const model = modelOf([row("note", "Note text")]);
    model.story = "removed";
    assert.deepEqual(compareComments(model, [comment], [], []), []);
    assert.match(
      compareComments(model, [comment], [comment], []).join("\n"),
      /outlived the block it anchored/u,
    );
  });

  test("a note comment retires when its body reference is removed", () => {
    const comment = {
      id: 1,
      text: "review",
      anchor: "Note text",
      blockId: "note",
    };
    const model = modelOf([row("body", "Reference")]);
    model.removedNoteStories.add("footnote:1");
    model.noteStoryByBlockId = new Map([["note", "footnote:1"]]);
    assert.deepEqual(compareComments(model, [comment], [], []), []);
  });

  test("a suggested removal of a note reference keeps the note's comment pending", () => {
    const comment = {
      id: 1,
      text: "review",
      anchor: "Note text",
      blockId: "note",
    };
    const model = modelOf([row("body", "Reference")]);
    model.removedNoteStories.add("footnote:1");
    model.noteStoryByBlockId = new Map([["note", "footnote:1"]]);
    assert.deepEqual(
      compareComments(model, [comment], [comment], [], [comment], "suggested"),
      [],
    );
    assert.match(
      compareComments(model, [comment], [], [], [comment], "suggested").join(
        "\n",
      ),
      /no comment/u,
    );
    const removedStory = modelOf([row("note", "Note text")]);
    removedStory.story = "removed";
    assert.deepEqual(
      compareComments(
        removedStory,
        [comment],
        [comment],
        [],
        [comment],
        "suggested",
      ),
      [],
    );
  });

  test("a merge keeps the note whose reference it moves, and a later deletion retires it", () => {
    const footnote = { type: "footnote", noteId: 1 } as const;
    const endnote = { type: "endnote", noteId: 2 } as const;
    const references: Parameters<typeof retireNoteReferences>[0] = new Map([
      ["first", [endnote]],
      ["second", [footnote]],
    ]);
    assert.deepEqual(
      retireNoteReferences(references, "mergeBlockWithNext", "first", [
        "second",
      ]),
      [],
    );
    assert.deepEqual(references.get("first"), [endnote, footnote]);
    assert.equal(references.has("second"), false);
    assert.deepEqual(
      retireNoteReferences(references, "deleteBlock", "first", ["first"]),
      ["endnote:2", "footnote:1"],
    );
    assert.equal(references.size, 0);
  });

  test("a note whose reference an edit may have moved is not predicted either way", () => {
    const comment = {
      id: 1,
      text: "review",
      anchor: "Note text",
      blockId: "note",
    };
    const model = modelOf([row("body", "Reference")]);
    model.removedNoteStories.add("footnote:1");
    model.uncertainNoteStories.add("footnote:1");
    model.noteStoryByBlockId = new Map([["note", "footnote:1"]]);
    assert.deepEqual(compareComments(model, [comment], [comment], []), []);
    assert.deepEqual(compareComments(model, [comment], [], []), []);
  });

  test("every operation that keeps note references in place is a real operation type", () => {
    for (const type of KEEPS_NOTE_REFERENCES)
      assert.ok(OPERATION_TYPES.includes(type), type);
    assert.equal(KEEPS_NOTE_REFERENCES.has("splitBlock"), false);
    assert.equal(KEEPS_NOTE_REFERENCES.has("replaceBlock"), false);
  });

  test("a note referenced from two blocks retires only with its last reference", () => {
    const footnote = { type: "footnote", noteId: 1 } as const;
    const references: Parameters<typeof retireNoteReferences>[0] = new Map([
      ["first", [footnote]],
      ["second", [footnote]],
    ]);
    assert.deepEqual(
      retireNoteReferences(references, "deleteBlock", "first", ["first"]),
      [],
    );
    assert.deepEqual(
      retireNoteReferences(references, "deleteBlock", "second", ["second"]),
      ["footnote:1"],
    );
  });

  test("removing one of two footnote references retires only that note's comment", () => {
    const first = {
      id: 1,
      text: "first review",
      anchor: "First note",
      blockId: "note-1",
    };
    const second = {
      id: 2,
      text: "second review",
      anchor: "Second note",
      blockId: "note-2",
    };
    const model = modelOf([
      row("body-1", "First reference"),
      row("body-2", "Second reference"),
    ]);
    model.removedNoteStories.add("footnote:1");
    model.noteStoryByBlockId = new Map([
      ["note-1", "footnote:1"],
      ["note-2", "footnote:2"],
    ]);

    assert.deepEqual(compareComments(model, [first, second], [second], []), []);
    assert.match(
      compareComments(model, [first, second], [], []).join("\n"),
      /second review/u,
    );
  });

  test("a replaced paragraph's comment survives rejecting its tracked deletion", async () => {
    const reviewer = await openReviewer(await plainDocument());
    const block = reviewer.getContent().find(({ text }) => text.length > 0);
    assert.ok(block);
    const note = "Review this paragraph.";
    const apply = (
      batch: Omit<
        Parameters<Reviewer["applyDocumentOperations"]>[0],
        "version"
      >,
    ) => {
      const result = reviewer.applyDocumentOperations(
        {
          version: FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
          ...batch,
        },
        { undefinedReferences: "refuse" },
      );
      assert.equal(result.applied.length, batch.operations.length);
      assert.deepEqual(result.skipped, []);
    };
    apply({
      mode: "direct",
      operations: [
        {
          id: "comment",
          type: "commentOnBlock",
          blockId: block.id,
          comment: { text: note },
        },
      ],
    });
    apply({
      mode: "tracked-changes",
      operations: [
        {
          id: "replace",
          type: "replaceBlock",
          blockId: block.id,
          text: "Replacement.",
        },
      ],
    });
    apply({
      mode: "tracked-changes",
      operations: [{ id: "delete", type: "deleteBlock", blockId: block.id }],
    });
    const bytes = new Uint8Array(await reviewer.toBuffer());
    const accepted = await openReviewer(bytes);
    accepted.acceptAll();
    assert.equal(
      accepted.getComments().some(({ text }) => text === note),
      false,
    );
    const rejected = await openReviewer(bytes);
    rejected.rejectAll();
    const restored = rejected.getComments().find(({ text }) => text === note);
    assert.ok(restored);
    assert.equal(restored.anchoredText, block.text);
  });

  test("stable comment ids use accepted anchors when live quotes include deleted text", () => {
    const accepted = "Written";
    for (const old of ["Old", "Written before", "before Written after"]) {
      for (const liveAnchor of [old + accepted, accepted + old]) {
        const comment = {
          id: 1,
          text: "review",
          anchor: accepted,
          blockId: "b",
        };
        const live = { ...comment, anchor: liveAnchor };
        const model = modelOf([row("b", accepted)], [row("b", accepted)]);
        expectOperation(model, { type: "deleteBlock", blockId: "b" });
        assert.deepEqual(
          compareComments(model, [comment], [], [], [live], "tracked-changes"),
          [],
        );
        assert.match(
          compareComments(
            model,
            [comment],
            [comment],
            [],
            [live],
            "tracked-changes",
          ).join("\n"),
          /outlived the block it anchored/u,
        );
        assert.match(
          compareComments(model, [comment], [], [], [live], "suggested").join(
            "\n",
          ),
          /no comment/u,
        );
      }
    }
  });

  test("uses live anchors and ids when an accepted pending join moves one of three equal comments", () => {
    const a = { id: 1, text: "same note", anchor: "First", blockId: "a" };
    const moved = { id: 2, text: "same note", anchor: "Signed", blockId: "c" };
    const c = { id: 3, text: "same note", anchor: "Last", blockId: "c" };
    const model = modelOf([
      row("a", "First"),
      row("b", "Signed"),
      row("c", "Signed Last"),
    ]);
    expectOperation(model, { type: "deleteBlock", blockId: "b" });
    const live = [a, { ...moved, blockId: "b" }, c];
    assert.deepEqual(
      compareComments(
        model,
        [a, moved, c],
        [a, c],
        [row("a", "First"), row("c", "Signed Last")],
        live,
      ),
      [],
    );
    assert.deepEqual(
      compareComments(
        model,
        [a, moved, c],
        [a, moved, c],
        [row("a", "First"), row("c", "Signed Last")],
        live,
      ),
      [],
    );
  });

  test("a comment on a pending join survives an edit to another block", () => {
    const joined = { id: 1, text: "review", anchor: "Signed", blockId: "b" };
    const model = modelOf(
      [row("a", "First Signed"), row("c", "Other")],
      [row("a", "First"), row("b", "Signed"), row("c", "Other")],
    );
    expectOperation(model, {
      type: "replaceBlock",
      blockId: "c",
      text: "Revised",
    });
    assert.deepEqual(
      compareComments(model, [joined], [joined], [], [joined]),
      [],
    );
    assert.equal(compareComments(model, [joined], [], [], [joined]).length, 1);
    expectOperation(model, { type: "deleteBlock", blockId: "b" });
    assert.deepEqual(compareComments(model, [joined], [], [], [joined]), []);
    assert.match(
      compareComments(model, [joined], [joined], [], [joined]).join("\n"),
      /outlived the block it anchored/u,
    );
  });

  test("an explicit block deletion removes a live comment even when accepted text differs", () => {
    const comment = {
      id: 1,
      text: "review",
      anchor: "Live quote",
      blockId: "b",
    };
    const model = modelOf(
      [row("b", "Accepted text")],
      [row("b", "Live quote")],
    );
    expectOperation(model, { type: "deleteBlock", blockId: "b" });
    assert.deepEqual(compareComments(model, [comment], [], [], [comment]), []);
    assert.match(
      compareComments(model, [comment], [comment], [], [comment]).join("\n"),
      /outlived the block it anchored/u,
    );
  });

  test("deleting the first block of a spanning comment keeps its surviving anchor", () => {
    const comment = {
      id: 1,
      text: "review",
      anchor: "FirstSecond",
      blockId: "b",
    };
    const model = modelOf([row("b", "First"), row("c", "Second")]);
    expectOperation(model, { type: "deleteBlock", blockId: "b" });
    assert.deepEqual(compareComments(model, [comment], [comment], []), []);
    assert.equal(
      compareComments(model, [comment], [], [row("c", "Second")]).length,
      1,
    );
  });

  test("a suggested block deletion leaves the live comment in place", () => {
    const comment = { id: 1, text: "review", anchor: "Signed", blockId: "b" };
    const model = modelOf([row("b", "Signed")]);
    expectOperation(model, { type: "deleteBlock", blockId: "b" });
    assert.deepEqual(
      compareComments(
        model,
        [comment],
        [comment],
        [row("b", "Signed")],
        [comment],
        "suggested",
      ),
      [],
    );
  });
});

test("a style edit preserves the target of an unchanged link", async () => {
  const reviewer = await openReviewer(
    await packDocument(
      fromMarkdown(
        "# Terms\n\nRead the [schedule](https://example.com/schedule).",
      ),
    ),
  );
  const block = blocksOf(reviewer).find((candidate) =>
    candidate.text.includes("schedule"),
  );
  assert.ok(block);
  assert.ok(
    captureLinks(reviewer)
      .get(block.id)
      ?.targets.includes("https://example.com/schedule"),
  );
  const { applied } = await applyChecked(
    reviewer,
    [
      {
        type: "setBlockParagraphProperties",
        blockId: block.id,
        properties: { alignment: "center" },
      },
    ],
    "direct",
    "link after style edit",
  );
  assert.equal(applied.length, 1);
});

test("a missing saved link projection fails for an unchanged reader block", async () => {
  const reviewer = await openReviewer(
    await packDocument(
      fromMarkdown("Read the [schedule](https://example.com/schedule)."),
    ),
  );
  const before = captureLinks(reviewer);
  const { reopened } = await saveAndReopen(reviewer, "link projection fixture");
  const after = new Map(captureLinks(reopened));
  const linked = [...before.keys()].at(0);
  assert.ok(linked);
  assert.ok(after.delete(linked));

  const unchanged = comparePreservedLinks({
    before,
    after,
    afterRows: blocksOf(reopened),
  });
  assert.equal(unchanged.checked, 1);
  assert.match(
    unchanged.problems.join("\n"),
    /has no matching link projection/u,
  );

  const editedRows = blocksOf(reopened).map(({ id, text }) => ({
    id,
    text: id === linked ? `${text} changed` : text,
  }));
  assert.deepEqual(
    comparePreservedLinks({ before, after, afterRows: editedRows }),
    {
      problems: [],
      checked: 0,
    },
  );
});

test("a signature table has one cell per party after save and reopen", async () => {
  const reviewer = await openReviewer(await plainDocument());
  const block = blocksOf(reviewer).find(
    (candidate) => candidate.text === "Signed in two copies.",
  );
  assert.ok(block);
  const { applied } = await applyChecked(
    reviewer,
    [
      {
        type: "insertSignatureTable",
        blockId: block.id,
        parties: [
          { name: "Supplier", signatory: "A. Example" },
          { name: "Buyer", title: "Director" },
        ],
      },
    ],
    "direct",
    "signature table geometry",
  );
  assert.equal(applied.length, 1);
});

test("cell merge and split predict the saved grid", async () => {
  const reviewer = await openReviewer(await ALL_FIXTURES.tables());
  const left = blocksOf(reviewer).find(
    (candidate) => candidate.text === "Item",
  );
  const right = blocksOf(reviewer).find(
    (candidate) => candidate.text === "Price",
  );
  assert.ok(left);
  assert.ok(right);
  const merged = await applyChecked(
    reviewer,
    [{ type: "mergeTableCells", blockId: left.id, endBlockId: right.id }],
    "direct",
    "merge grid",
  );
  assert.equal(merged.applied.length, 1);
  const split = await applyChecked(
    reviewer,
    [{ type: "splitTableCell", blockId: left.id }],
    "direct",
    "split grid",
  );
  assert.equal(split.applied.length, 1);
});

test("a deleted merge neighbor does not join the following table cell", async () => {
  const reviewer = await openReviewer(await ALL_FIXTURES.tables());
  const heading = blocksOf(reviewer).find(
    (candidate) => candidate.text === "Price Schedule",
  );
  const neighbor = blocksOf(reviewer).find(
    (candidate) => candidate.text === "The prices below apply.",
  );
  assert.ok(heading);
  assert.ok(neighbor);
  const { applied } = await applyChecked(
    reviewer,
    [
      { type: "mergeBlockWithNext", blockId: heading.id, separator: " " },
      { type: "deleteBlock", blockId: neighbor.id },
    ],
    "tracked-changes",
    "join then remove neighbor",
  );
  assert.equal(applied.length, 2);
});

function location(cellIndex: number) {
  return {
    outerTableIndex: 0,
    tableIndex: 0,
    rowIndex: 0,
    cellIndex,
    gridColumnIndex: cellIndex,
    columnSpan: 1,
    rowSpan: 1,
  };
}

/**
 * `docxToMarkdown` writes a table with merged cells as HTML, which the reader
 * comparison does not parse; that fixture is saved and reopened only.
 */
const READERS_COMPARABLE = (fixture: string) => fixture !== "mergedTable";

const persistedAfterBulkAcceptance = async (reviewer: Reviewer) => {
  const saved = await openReviewer(new Uint8Array(await reviewer.toBuffer()));
  saved.acceptAll();
  return visibleState(saved);
};

/** Run one collision on a fresh reviewer; what applied, or null when none fits. */
const runCollision = async (
  fixture: string,
  bytes: Uint8Array,
  mode: Mode,
  name: string,
): Promise<string | null> => {
  const context = `${fixture} / ${mode}: ${name}`;
  const reviewer = await openReviewer(bytes);
  const persisted =
    mode === "suggested"
      ? { persisted: await persistedAfterBulkAcceptance(reviewer) }
      : {};
  const operations = COLLISIONS[name]?.(blocksOf(reviewer));
  if (!operations) return null;
  const { applied, issues } = await applyChecked(
    reviewer,
    operations,
    mode,
    context,
  );
  if (mode === "suggested") reviewer.acceptAll();
  if (READERS_COMPARABLE(fixture))
    await assertHealthy(reviewer, context, persisted);
  else await saveAndReopen(reviewer, context, persisted);
  return `${name}: applied ${applied.length}${issues.length > 0 ? `, refused ${issues.join(", ")}` : ""}`;
};

describe("collision batches", () => {
  for (const [fixture, build] of Object.entries(ALL_FIXTURES)) {
    for (const mode of MODES) {
      const known = KNOWN_FAILING_COLLISIONS.filter(
        (entry) => entry.fixture === fixture && entry.mode === mode,
      );
      test(`${fixture} / ${mode}: every collision does what its applied operations asked`, async () => {
        const bytes = await build();
        const ran: string[] = [];
        for (const name of Object.keys(COLLISIONS)) {
          if (known.some((entry) => entry.collision === name)) continue;
          const outcome = await runCollision(fixture, bytes, mode, name);
          if (outcome !== null) ran.push(outcome);
        }
        if (process.env["FOLIO_ORACLE_GAPS"])
          console.log(`${fixture} / ${mode}:\n  ${ran.join("\n  ")}`);
        assert.ok(ran.length > 0, "no collision fits this fixture");
      });
      for (const { collision, finding, symptom } of known) {
        expectedFailure(
          finding,
          `${fixture} / ${mode}: ${collision}`,
          symptom,
          async () => {
            await runCollision(fixture, await build(), mode, collision);
          },
        );
      }
    }
  }
});

/**
 * A first edit, saved and reopened, then a second one that names what the
 * first left pending.
 */
type FollowUp = {
  first: (blocks: readonly Block[]) => Operation[];
  second: (blocks: readonly Block[]) => Operation[] | null;
  /** The issue code every mode must refuse the second edit with. */
  refusedWith?: string;
};

const INSERTED = "An inserted clause about delivery terms.";
const anchorOf = (blocks: readonly Block[]) =>
  blocks.find((block) => block.text.startsWith("The Supplier")) as Block;
const insertedOf = (blocks: readonly Block[]) =>
  blocks.find((block) => block.text === INSERTED);
const insertFirst = (blocks: readonly Block[]): Operation[] => [
  { type: "insertAfterBlock", blockId: anchorOf(blocks).id, text: INSERTED },
];
const onInserted =
  (build: (block: Block) => Operation) =>
  (blocks: readonly Block[]): Operation[] | null => {
    const block = insertedOf(blocks);
    return block ? [build(block)] : null;
  };
const rangeOver = (block: Block, word: string) => {
  const start = block.text.indexOf(word);
  return createRange(block, start, start + word.length);
};

const FOLLOW_UPS: Record<string, FollowUp> = {
  replaceInPendingInsertion: {
    first: insertFirst,
    second: onInserted((block) => ({
      type: "replaceInBlock",
      blockId: block.id,
      find: "delivery",
      replace: "payment",
    })),
  },
  replaceRangeInPendingInsertion: {
    first: insertFirst,
    second: onInserted((block) => ({
      type: "replaceRange",
      range: rangeOver(block, "clause"),
      replace: "term",
    })),
  },
  deletePendingInsertion: {
    first: insertFirst,
    second: onInserted((block) => ({ type: "deleteBlock", blockId: block.id })),
  },
  formatPendingInsertion: {
    first: insertFirst,
    second: onInserted((block) => ({
      type: "formatRange",
      range: rangeOver(block, "inserted"),
      formatting: { bold: true },
    })),
  },
  splitPendingInsertion: {
    first: insertFirst,
    second: onInserted((block) => ({
      type: "splitBlock",
      blockId: block.id,
      offset: block.text.indexOf("about"),
    })),
  },
  replaceOverPendingReplacement: {
    first: (blocks) => [
      {
        type: "replaceInBlock",
        blockId: anchorOf(blocks).id,
        find: "goods",
        replace: "products",
      },
    ],
    second: (blocks) => {
      const block = blocks.find((candidate) =>
        candidate.text.includes("products"),
      );
      return block
        ? [
            {
              type: "replaceRange",
              range: rangeOver(block, "products"),
              replace: "wares",
            },
          ]
        : null;
    },
  },
  rewritePendingReplacement: {
    first: (blocks) => [
      {
        type: "replaceInBlock",
        blockId: anchorOf(blocks).id,
        find: "goods",
        replace: "products",
      },
    ],
    second: (blocks) => {
      const block = blocks.find((candidate) =>
        candidate.text.includes("products"),
      );
      return block
        ? [
            {
              type: "replaceBlock",
              blockId: block.id,
              text: "The Supplier delivers promptly.",
            },
          ]
        : null;
    },
  },
  // A paragraph pending deletion reads as a blank block, which a model may name.
  // Text typed there stays joined to the next paragraph once accepted, so
  // rewriting it is refused rather than glued onto that paragraph.
  rewritePendingDeletion: {
    first: deleteFirst,
    second: onDeleted((block) => ({
      type: "replaceBlock",
      blockId: block.id,
      text: "Rewritten.",
    })),
    refusedWith: "pendingDeletion",
  },
  insertAfterPendingDeletion: {
    first: deleteFirst,
    second: onDeleted((block) => ({
      type: "insertAfterBlock",
      blockId: block.id,
      text: "Added.",
    })),
  },
  insertBeforePendingDeletion: {
    first: deleteFirst,
    second: onDeleted((block) => ({
      type: "insertBeforeBlock",
      blockId: block.id,
      text: "Added.",
    })),
  },
};

function deleteFirst(blocks: readonly Block[]): Operation[] {
  return [{ type: "deleteBlock", blockId: anchorOf(blocks).id }];
}

function onDeleted(build: (block: Block) => Operation) {
  return (blocks: readonly Block[]): Operation[] | null => {
    const block = blocks.find((candidate) => candidate.text === "");
    return block ? [build(block)] : null;
  };
}

function createRange(block: Block, start: number, end: number) {
  return {
    type: "textRange",
    story: "main",
    blockId: block.id,
    startOffset: start,
    endOffset: end,
    selectedTextHash: hashFolioAIBlockText(block.text.slice(start, end)),
  };
}

const runFollowUp = async (name: string, mode: Mode): Promise<void> => {
  const followUp = FOLLOW_UPS[name] as FollowUp;
  const context = `tracked, saved, then ${mode}: ${name}`;
  const reviewer = await openReviewer(await plainDocument());
  const first = await applyChecked(
    reviewer,
    followUp.first(blocksOf(reviewer)),
    "tracked-changes",
    context,
  );
  assert.ok(
    first.applied.length > 0,
    `${context}: the first edit was refused: ${first.issues}`,
  );
  const reopened = await openReviewer(
    new Uint8Array(await reviewer.toBuffer()),
  );
  const persisted =
    mode === "suggested"
      ? { persisted: await persistedAfterBulkAcceptance(reopened) }
      : {};
  const operations = followUp.second(blocksOf(reopened));
  assert.ok(operations, `${context}: the second edit found nothing to name`);
  const second = await applyChecked(reopened, operations, mode, context);
  if (followUp.refusedWith !== undefined) {
    assert.deepEqual(
      second.applied,
      [],
      `${context}: applied what it must refuse`,
    );
    assert.ok(
      second.issues.every((issue) =>
        issue.endsWith(`: ${followUp.refusedWith}`),
      ),
      `${context}: refused with ${second.issues.join(", ")}, not ${followUp.refusedWith}`,
    );
  }
  if (mode === "suggested") reopened.acceptAll();
  await assertHealthy(reopened, context, persisted);
};

describe("edits of what a saved tracked edit left pending", () => {
  for (const name of Object.keys(FOLLOW_UPS)) {
    for (const mode of MODES) {
      const known = KNOWN_FAILING_FOLLOW_UPS.find(
        (entry) => entry.followUp === name && entry.mode === mode,
      );
      const title = `${name} (${mode})`;
      if (known) {
        expectedFailure(known.finding, title, known.symptom, () =>
          runFollowUp(name, mode),
        );
      } else {
        test(title, () => runFollowUp(name, mode));
      }
    }
  }
});

/** Style ids the package does not define as paragraph styles, named by an operation. */
const STYLE_REQUESTS: Record<string, (block: Block) => Operation> = {
  "insertAfterBlock with an undefined style": (block) => ({
    type: "insertAfterBlock",
    blockId: block.id,
    text: "Styled clause.",
    styleId: "NoSuchStyle",
  }),
  "replaceBlock with an undefined style": (block) => ({
    type: "replaceBlock",
    blockId: block.id,
    text: "Styled clause.",
    styleId: "NoSuchStyle",
  }),
  "setBlockParagraphProperties with an undefined style": (block) => ({
    type: "setBlockParagraphProperties",
    blockId: block.id,
    properties: { styleId: "NoSuchStyle" },
  }),
  "setBlockParagraphProperties with a table style": (block) => ({
    type: "setBlockParagraphProperties",
    blockId: block.id,
    properties: { styleId: "TableGrid" },
  }),
};

const runStyleRequest = async (name: string, mode: Mode): Promise<void> => {
  const reviewer = await openReviewer(await ALL_FIXTURES.tables());
  const block = blocksOf(reviewer).find(
    (candidate) => candidate.text === "Taxes are extra.",
  );
  assert.ok(block);
  const build = STYLE_REQUESTS[name] as (block: Block) => Operation;
  await applyChecked(reviewer, [build(block)], mode, `${name} (${mode})`);
};

describe("a style an operation names is a paragraph style of the saved package, or the operation is refused", () => {
  for (const name of Object.keys(STYLE_REQUESTS)) {
    for (const mode of ["direct", "tracked-changes"] as const) {
      test(`${name} (${mode})`, () => runStyleRequest(name, mode));
    }
  }
});

describe("a paragraph property edit replayed after its block's deletion", () => {
  const row = (id: string, text: string, cellIndex?: number): Row => ({
    id,
    text,
    kind: "paragraph",
    ...(cellIndex === undefined ? {} : { table: location(cellIndex) }),
  });

  test("formats a block only when the deletion kept it, emptied and last", () => {
    const before = [row("a", "Alpha."), row("b", "Bravo.")];
    // The story's last block, emptied and still last.
    assert.equal(
      deletionKeptBlock({
        before,
        after: [row("a", "Alpha."), row("b", "")],
        id: "b",
      }),
      true,
    );
    // Not the last block before the batch: the deletion removes it, whatever
    // the result holds.
    assert.equal(
      deletionKeptBlock({
        before,
        after: [row("a", ""), row("b", "Bravo.")],
        id: "a",
      }),
      false,
    );
    // The last block, but the result no longer holds it.
    assert.equal(
      deletionKeptBlock({ before, after: [row("a", "Alpha.")], id: "b" }),
      false,
    );
    // Still there, but not emptied.
    assert.equal(deletionKeptBlock({ before, after: before, id: "b" }), false);
    // Emptied, but no longer last.
    assert.equal(
      deletionKeptBlock({
        before,
        after: [row("a", "Alpha."), row("b", ""), row("c", "Charlie.")],
        id: "b",
      }),
      false,
    );
    // A table cell ends with its own last block, whatever follows the table.
    const cells = [row("x", "Item", 0), row("y", "Price", 1), row("z", "End.")];
    assert.equal(
      deletionKeptBlock({
        before: cells,
        after: [row("x", "", 0), row("y", "Price", 1), row("z", "End.")],
        id: "x",
      }),
      true,
    );
  });

  const finalParagraph = (reviewer: Reviewer): Block => {
    const blocks = blocksOf(reviewer);
    const block = blocks.at(-1);
    assert.ok(block);
    assert.equal(block.text, "Schedules may change by agreement.");
    assert.ok(
      block.table === undefined && blocks.at(-2)?.table !== undefined,
      "the fixture ends with a paragraph after a table",
    );
    return block;
  };
  const center = (block: Block): Operation => ({
    type: "setBlockParagraphProperties",
    blockId: block.id,
    properties: { alignment: "center" },
  });
  const apply = (
    reviewer: Reviewer,
    operations: readonly Operation[],
  ): AppliedOperation[] => {
    const batch = coreBatch(operations, "direct");
    const result = reviewer.applyDocumentOperations(batch as never, {
      undefinedReferences: "refuse",
    });
    assert.equal(result.applied.length, operations.length);
    return batch.operations;
  };

  for (const mode of ["direct", "tracked-changes"] as const) {
    test(`formats the closing paragraph after a table that its deletion empties (${mode})`, async () => {
      const reviewer = await openReviewer(await ALL_FIXTURES.tables());
      const closing = finalParagraph(reviewer);
      const { applied } = await applyChecked(
        reviewer,
        [{ type: "deleteBlock", blockId: closing.id }, center(closing)],
        mode,
        `delete then format the closing paragraph (${mode})`,
      );
      assert.deepEqual(applied, ["op-1", "op-2"]);
    });
  }

  test("reports a requested property the kept block lacks", async () => {
    const reviewer = await openReviewer(await ALL_FIXTURES.tables());
    const closing = finalParagraph(reviewer);
    const pre = await capture(reviewer, "direct");
    // Only the deletion runs; the receipt claims the format too.
    const [deletion] = apply(reviewer, [
      { type: "deleteBlock", blockId: closing.id },
    ]);
    assert.ok(deletion);
    const [, format] = coreBatch(
      [{ type: "deleteBlock", blockId: closing.id }, center(closing)],
      "direct",
    ).operations;
    assert.ok(format);
    await assert.rejects(
      assertRequestedOutcome(
        reviewer,
        pre,
        { applied: [deletion, format] },
        "a format the engine dropped",
      ),
      /the emptied block .*directAlignment is undefined, expected "center"/u,
    );
  });

  test("reports a format of a block its deletion really removed", async () => {
    const reviewer = await openReviewer(await ALL_FIXTURES.tables());
    const closing = finalParagraph(reviewer);
    const pre = await capture(reviewer, "direct");
    // A paragraph after it first, so the deletion takes the block out; the
    // receipt claims the deletion and a format of the removed block.
    apply(reviewer, [
      { type: "insertAfterBlock", blockId: closing.id, text: "Added." },
    ]);
    apply(reviewer, [{ type: "deleteBlock", blockId: closing.id }]);
    assert.ok(!blocksOf(reviewer).some((block) => block.id === closing.id));
    const claimed = coreBatch(
      [{ type: "deleteBlock", blockId: closing.id }, center(closing)],
      "direct",
    ).operations;
    await assert.rejects(
      assertRequestedOutcome(
        reviewer,
        pre,
        { applied: claimed },
        "a format of a removed block",
      ),
      /which an earlier one of the batch removed/u,
    );
  });
});

describe("findings the oracle pins", () => {
  for (const mode of ["direct", "tracked-changes"] as const) {
    test(`two paragraph-property operations on one block in one batch: the later one is refused (${mode})`, async () => {
      const reviewer = await openReviewer(await plainDocument());
      const block = blocksOf(reviewer).at(-1) as Block;
      const align = (alignment: string): Operation => ({
        type: "setBlockParagraphProperties",
        blockId: block.id,
        properties: { alignment },
      });
      const { applied, issues } = await applyChecked(
        reviewer,
        [align("center"), align("right")],
        mode,
        `alignment twice (${mode})`,
      );
      assert.deepEqual(applied, ["op-1"]);
      assert.deepEqual(issues, ["op-2: overlappingOperation"]);
    });
  }

  for (const mode of ["tracked-changes", "suggested"] as const) {
    test(`a row deletion over a pending word leaves the reader no text of that row (${mode})`, async () => {
      const reviewer = await openReviewer(await ALL_FIXTURES.tables());
      const cell = blocksOf(reviewer).find(
        (block) => block.text === "Price",
      ) as Block;
      await applyChecked(
        reviewer,
        [
          {
            type: "replaceInBlock",
            blockId: cell.id,
            find: "Price",
            replace: "amended",
          },
        ],
        mode,
        "a pending word",
      );
      await applyChecked(
        reviewer,
        [{ type: "deleteTableRow", blockId: cell.id }],
        mode,
        "then its row's deletion",
      );
      assert.ok(
        !blocksOf(reviewer).some((block) => block.text.includes("amended")),
        "the reader still lists the deleted row's pending word",
      );
    });
  }
});
