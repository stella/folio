/**
 * Tracked operations and their resolution on small fixed documents: the
 * construction rules one by one, and the review classes an editor has got
 * wrong (formatting applied directly instead of tracked, Enter and Delete
 * marks not recorded, a split paragraph given resolved attributes, an undo
 * that silently dropped a mark).
 */

import { describe, expect, test } from "bun:test";

import type {
  BlockContent,
  Document,
  Paragraph,
  ParagraphContent,
  TextFormatting,
} from "../../model/document";
import { applyDocumentOp, applyDocumentOps, type AppliedDocumentOp } from "../apply";
import { contractViolation } from "../contract";
import { planTrackedDeletion, revisionIdDemand } from "../plan";
import { DOCUMENT_OP_REFUSAL_REASONS } from "../refusal";
import { DOCUMENT_OP_TYPES, type DocumentOp, OP_STORIES, type RevisionStamp } from "../types";

const DATE = "2026-05-06T07:08:09Z";

const stamp = (id: number, author = "Reviewer"): RevisionStamp => ({ id, author, date: DATE });

const run = (text: string, formatting?: TextFormatting): ParagraphContent =>
  formatting === undefined
    ? { type: "run", content: [{ type: "text", text }] }
    : { type: "run", formatting, content: [{ type: "text", text }] };

const paragraph = (
  paraId: string,
  content: ParagraphContent[],
  fields: Partial<Paragraph> = {},
): Paragraph => ({ ...fields, type: "paragraph", paraId, content });

const documentOf = (...content: BlockContent[]): Document => ({
  package: { document: { content } },
});

const at = (blockId: string, offset: number) => ({ story: OP_STORIES.MAIN, blockId, offset });

const applied = (document: Document, op: DocumentOp): AppliedDocumentOp => {
  const result = applyDocumentOp(document, op);
  if (result.isErr()) throw result.error;
  expect(contractViolation(result.value.document)).toBeUndefined();
  // Every tracked operation is undone exactly.
  const undone = applyDocumentOps(result.value.document, result.value.inverse);
  if (undone.isErr()) throw undone.error;
  expect(undone.value.document).toStrictEqual(document);
  return result.value;
};

/** The operation without its stamp: what it does directly. */
const directly = (op: DocumentOp): DocumentOp => {
  const direct = { ...op };
  Reflect.deleteProperty(direct, "revision");
  return direct;
};

const refusalOf = (document: Document, op: DocumentOp): string | undefined => {
  const result = applyDocumentOp(document, op);
  return result.isErr() ? result.error.reason : undefined;
};

const blocks = (document: Document): BlockContent[] => document.package.document.content;

describe("tracked text", () => {
  test("a tracked insertion is wrapped in an insertion carrying the stamp", () => {
    const document = documentOf(paragraph("00000001", [run("Hello")]));
    const insert: DocumentOp = {
      type: DOCUMENT_OP_TYPES.INSERT_TEXT,
      at: at("00000001", 5),
      text: " world",
      runProps: "inherit",
      newIds: { revision: [2] },
      revision: stamp(1),
    };
    const tracked = applied(document, insert);
    expect(tracked.revisions).toEqual([1]);
    expect(blocks(tracked.document)).toEqual([
      paragraph("00000001", [
        run("Hello"),
        {
          type: "insertion",
          info: { id: 1, author: "Reviewer", date: DATE },
          content: [run(" world")],
        },
      ]),
    ]);
  });

  test("a tracked insertion with the same stamp joins the wrapper it follows", () => {
    const document = documentOf(paragraph("00000001", [run("Hello")]));
    const first = applied(document, {
      type: DOCUMENT_OP_TYPES.INSERT_TEXT,
      at: at("00000001", 5),
      text: " wo",
      runProps: { bold: true },
      revision: stamp(1),
    });
    const second = applied(first.document, {
      type: DOCUMENT_OP_TYPES.INSERT_TEXT,
      at: at("00000001", 8),
      text: "rld",
      runProps: { italic: true },
      revision: stamp(2),
    });
    expect(second.revisions).toEqual([]);
    const [only] = blocks(second.document);
    expect(only?.type === "paragraph" && only.content.map(({ type }) => type)).toEqual([
      "run",
      "insertion",
    ]);
    // Another date is another change.
    const later = applied(first.document, {
      type: DOCUMENT_OP_TYPES.INSERT_TEXT,
      at: at("00000001", 8),
      text: "rld",
      runProps: { italic: true },
      revision: { ...stamp(2), date: "2026-05-06T07:08:10Z" },
    });
    expect(later.revisions).toEqual([2]);
  });

  test("a tracked insertion splits another author's insertion instead of nesting in it", () => {
    const theirs: ParagraphContent = {
      type: "insertion",
      info: { id: 5, author: "Other", date: "2026-01-01T00:00:00Z" },
      content: [run("ab")],
    };
    const document = documentOf(paragraph("00000001", [theirs]));
    const tracked = applied(document, {
      type: DOCUMENT_OP_TYPES.INSERT_TEXT,
      at: at("00000001", 1),
      text: "x",
      runProps: "inherit",
      newIds: { revision: [7] },
      revision: stamp(6),
    });
    const [only] = blocks(tracked.document);
    expect(only?.type === "paragraph" && only.content).toEqual([
      { ...theirs, content: [run("a")] },
      { type: "insertion", info: { id: 6, author: "Reviewer", date: DATE }, content: [run("x")] },
      { ...theirs, info: { ...theirs.info, id: 7 }, content: [run("b")] },
    ]);
  });

  test("a tracked insertion inside a deletion is refused", () => {
    const document = documentOf(
      paragraph("00000001", [
        {
          type: "deletion",
          info: { id: 5, author: "Other" },
          content: [run("a"), run("b", { bold: true })],
        },
      ]),
    );
    expect(
      refusalOf(document, {
        type: DOCUMENT_OP_TYPES.INSERT_TEXT,
        at: at("00000001", 1),
        text: "x",
        runProps: "inherit",
        revision: stamp(6),
      }),
    ).toBe(DOCUMENT_OP_REFUSAL_REASONS.INSIDE_TRACKED_DELETION);
    // Content whose open ends continue the deletion lands inside it.
    expect(
      refusalOf(document, {
        type: DOCUMENT_OP_TYPES.INSERT_CONTENT,
        at: at("00000001", 1),
        slice: {
          content: [{ type: "deletion", info: { id: 5, author: "Other" }, content: [run("x")] }],
          openStart: 1,
          openEnd: 1,
        },
        revision: stamp(6),
      }),
    ).toBe(DOCUMENT_OP_REFUSAL_REASONS.INSIDE_TRACKED_DELETION);
  });

  test("a tracked deletion removes nothing, nests inside insertions and skips deleted text", () => {
    const control: ParagraphContent = {
      type: "inlineSdt",
      properties: { sdtType: "richText", id: 3 },
      content: [run("d")],
    };
    const document = documentOf(
      paragraph("00000001", [
        { type: "insertion", info: { id: 5, author: "Other" }, content: [run("a")] },
        run("b"),
        { type: "deletion", info: { id: 6, author: "Other" }, content: [run("c")] },
        control,
      ]),
    );
    const tracked = applied(document, {
      type: DOCUMENT_OP_TYPES.DELETE_RANGE,
      from: at("00000001", 0),
      to: at("00000001", 4),
      newIds: { revision: [11, 12] },
      revision: stamp(10),
    });
    const info = (id: number) => ({ id, author: "Reviewer", date: DATE });
    expect(blocks(tracked.document)).toEqual([
      paragraph("00000001", [
        {
          type: "insertion",
          info: { id: 5, author: "Other" },
          content: [{ type: "deletion", info: info(10), content: [run("a")] }],
        },
        { type: "deletion", info: info(11), content: [run("b")] },
        { type: "deletion", info: { id: 6, author: "Other" }, content: [run("c")] },
        { type: "deletion", info: info(12), content: [control] },
      ]),
    ]);
    expect(tracked.revisions).toEqual([10, 11, 12]);
  });

  test("a tracked deletion refuses comment anchors; the plan deletes around them", () => {
    const document = documentOf(
      paragraph("00000001", [
        run("ab"),
        { type: "commentRangeStart", id: 1 },
        run("cd", { bold: true }),
        { type: "insertion", info: { id: 5, author: "Reviewer" }, content: [run("ef")] },
      ]),
    );
    const range = { from: at("00000001", 1), to: at("00000001", 6) };
    expect(
      refusalOf(document, { type: DOCUMENT_OP_TYPES.DELETE_RANGE, ...range, revision: stamp(9) }),
    ).toBe(DOCUMENT_OP_REFUSAL_REASONS.UNTRACKABLE);
    const plan = planTrackedDeletion(document, {
      ...range,
      revision: stamp(9),
      newIds: { revision: [20, 21, 22] },
    });
    if (plan.isErr()) throw plan.error;
    // The own insertion goes directly; the rest is tracked in two pieces around the anchor.
    expect(
      plan.value.map((op) => op.type === DOCUMENT_OP_TYPES.DELETE_RANGE && op.revision?.id),
    ).toEqual([undefined, 9, 20]);
    const after = applyDocumentOps(document, plan.value);
    if (after.isErr()) throw after.error;
    expect(blocks(after.value.document)).toEqual([
      paragraph("00000001", [
        run("a"),
        { type: "deletion", info: { id: 20, author: "Reviewer", date: DATE }, content: [run("b")] },
        { type: "commentRangeStart", id: 1 },
        {
          type: "deletion",
          info: { id: 9, author: "Reviewer", date: DATE },
          content: [run("cd", { bold: true })],
        },
      ]),
    ]);
  });

  test("revisionIdDemand counts the ids a tracked operation takes past its stamp", () => {
    const withChange: ParagraphContent = {
      type: "run",
      propertyChanges: [{ type: "runPropertyChange", info: { id: 4, author: "Other" } }],
      content: [{ type: "text", text: "abc" }],
    };
    const document = documentOf(paragraph("00000001", [withChange, run("de")]));
    const op: DocumentOp = {
      type: DOCUMENT_OP_TYPES.DELETE_RANGE,
      from: at("00000001", 1),
      to: at("00000001", 4),
      revision: stamp(9),
    };
    const demand = revisionIdDemand(document, op);
    // The piece of the cut run the deletion wraps takes a new id for its property change.
    expect(demand.isOk() && demand.value).toBe(1);
    expect(refusalOf(document, op)).toBe(DOCUMENT_OP_REFUSAL_REASONS.NEEDS_NEW_IDS);
    expect(refusalOf(document, { ...op, newIds: { revision: [30] } })).toBeUndefined();
  });
});

describe("C8: tracked formatting is recorded, not applied directly", () => {
  test("a tracked run patch records each run's own formatting", () => {
    const document = documentOf(paragraph("00000001", [run("ab", { italic: true }), run("cd")]));
    const patch: DocumentOp = {
      type: DOCUMENT_OP_TYPES.SET_RUN_PROPS,
      from: at("00000001", 1),
      to: at("00000001", 3),
      patch: { bold: true },
      newIds: { revision: [2] },
      revision: stamp(1),
    };
    const tracked = applied(document, patch);
    const info = (id: number) => ({ id, author: "Reviewer", date: DATE });
    expect(blocks(tracked.document)).toEqual([
      paragraph("00000001", [
        run("a", { italic: true }),
        {
          type: "run",
          formatting: { italic: true, bold: true },
          propertyChanges: [
            { type: "runPropertyChange", info: info(1), previousFormatting: { italic: true } },
          ],
          content: [{ type: "text", text: "b" }],
        },
        {
          type: "run",
          formatting: { bold: true },
          propertyChanges: [{ type: "runPropertyChange", info: info(2) }],
          content: [{ type: "text", text: "c" }],
        },
        run("d"),
      ]),
    ]);
    expect(tracked.revisions).toEqual([1, 2]);
  });

  test("a run already carrying a property change keeps its baseline", () => {
    const earlier = {
      type: "runPropertyChange" as const,
      info: { id: 4, author: "Other" },
      previousFormatting: { italic: true },
    };
    const document = documentOf(
      paragraph("00000001", [
        {
          type: "run",
          formatting: { bold: true },
          propertyChanges: [earlier],
          content: [{ type: "text", text: "ab" }],
        },
      ]),
    );
    const tracked = applied(document, {
      type: DOCUMENT_OP_TYPES.SET_RUN_PROPS,
      from: at("00000001", 0),
      to: at("00000001", 2),
      patch: { underline: { style: "single" } },
      revision: stamp(1),
    });
    expect(tracked.revisions).toEqual([]);
    expect(blocks(tracked.document)).toEqual([
      paragraph("00000001", [
        {
          type: "run",
          formatting: { bold: true, underline: { style: "single" } },
          propertyChanges: [earlier],
          content: [{ type: "text", text: "ab" }],
        },
      ]),
    ]);
    // Rejecting the change that was there restores what it started from.
  });

  test("a tracked paragraph patch records the paragraph properties, not the mark's", () => {
    const document = documentOf(
      paragraph("00000001", [run("ab")], {
        formatting: { styleId: "Heading1", runProperties: { bold: true } },
      }),
      paragraph("00000002", [run("cd")]),
    );
    const tracked = applied(document, {
      type: DOCUMENT_OP_TYPES.SET_PARAGRAPH_PROPS,
      story: OP_STORIES.MAIN,
      blockId: "00000001",
      patch: { alignment: "center" },
      revision: stamp(1),
    });
    const [first] = blocks(tracked.document);
    expect(first?.type === "paragraph" && first.propertyChanges).toEqual([
      {
        type: "paragraphPropertyChange",
        info: { id: 1, author: "Reviewer", date: DATE },
        previousFormatting: { styleId: "Heading1" },
      },
    ]);
    expect(
      refusalOf(document, {
        type: DOCUMENT_OP_TYPES.SET_PARAGRAPH_PROPS,
        story: OP_STORIES.MAIN,
        blockId: "00000001",
        patch: { runProperties: { italic: true } },
        revision: stamp(1),
      }),
    ).toBe(DOCUMENT_OP_REFUSAL_REASONS.UNTRACKABLE);
  });
});

describe("C9: Enter and Delete record paragraph marks", () => {
  const source = paragraph("00000001", [run("Hello"), run("World", { bold: true })], {
    formatting: { styleId: "Heading1" },
    textId: "77777777",
  });

  test("a tracked split inside a paragraph gives the new id to the first half", () => {
    const document = documentOf(source, paragraph("00000009", [run("next")]));
    const split: DocumentOp = {
      type: DOCUMENT_OP_TYPES.SPLIT_BLOCK,
      at: at("00000001", 5),
      newBlockId: "00000002",
      revision: stamp(1),
    };
    const tracked = applied(document, split);
    expect(blocks(tracked.document)).toEqual([
      paragraph("00000002", [run("Hello")], {
        formatting: { styleId: "Heading1" },
        pPrMark: { kind: "ins", info: { id: 1, author: "Reviewer", date: DATE } },
      }),
      { ...source, content: [run("World", { bold: true })] },
      paragraph("00000009", [run("next")]),
    ]);
  });

  test("a tracked split at the end adds a paragraph that records the source's properties", () => {
    const pending = {
      type: "paragraphPropertyChange" as const,
      info: { id: 7, author: "Other" },
      previousFormatting: { alignment: "end" as const },
    };
    const document = documentOf(
      { ...source, propertyChanges: [pending] },
      paragraph("00000009", [run("next")]),
    );
    const split: DocumentOp = {
      type: DOCUMENT_OP_TYPES.SPLIT_BLOCK,
      at: at("00000001", 10),
      newBlockId: "00000002",
      newParagraph: {},
      revision: stamp(1),
      newIds: { revision: [2] },
    };
    const tracked = applied(document, split);
    // The new paragraph takes the pending change, which keeps the properties it started from.
    expect(blocks(tracked.document)).toEqual([
      {
        ...source,
        pPrMark: { kind: "ins", info: { id: 1, author: "Reviewer", date: DATE } },
      },
      paragraph("00000002", [], { propertyChanges: [pending] }),
      paragraph("00000009", [run("next")]),
    ]);
    const plain = applied(documentOf(source, paragraph("00000009", [run("next")])), split);
    expect(blocks(plain.document)[1]).toEqual(
      paragraph("00000002", [], {
        propertyChanges: [
          {
            type: "paragraphPropertyChange",
            info: { id: 2, author: "Reviewer", date: DATE },
            previousFormatting: { styleId: "Heading1" },
          },
        ],
      }),
    );
  });

  test("a tracked join marks the first paragraph's mark as deleted and moves nothing", () => {
    const next = paragraph("00000002", [run("World")], { formatting: { alignment: "end" } });
    const document = documentOf(source, next);
    const join: DocumentOp = {
      type: DOCUMENT_OP_TYPES.JOIN_BLOCKS,
      story: OP_STORIES.MAIN,
      blockId: "00000001",
      nextBlockId: "00000002",
      newIds: { revision: [2] },
      revision: stamp(1),
    };
    const tracked = applied(document, join);
    expect(blocks(tracked.document)).toEqual([
      { ...source, pPrMark: { kind: "del", info: { id: 1, author: "Reviewer", date: DATE } } },
      {
        ...next,
        formatting: { styleId: "Heading1" },
        propertyChanges: [
          {
            type: "paragraphPropertyChange",
            info: { id: 2, author: "Reviewer", date: DATE },
            previousFormatting: { alignment: "end" },
          },
        ],
      },
    ]);
    // A mark already carrying a change is not overwritten.
    expect(refusalOf(tracked.document, { ...join, revision: stamp(3) })).toBe(
      DOCUMENT_OP_REFUSAL_REASONS.REVISION_CONFLICT,
    );
  });

  test("a mark is never given to a paragraph that ends its container", () => {
    const document = documentOf(source);
    expect(
      refusalOf(document, {
        type: DOCUMENT_OP_TYPES.SET_PARAGRAPH_REVIEW,
        story: OP_STORIES.MAIN,
        blockId: "00000001",
        expected: { formatting: { styleId: "Heading1" } },
        review: {
          formatting: { styleId: "Heading1" },
          pPrMark: { kind: "del", info: { id: 1, author: "Reviewer" } },
        },
      }),
    ).toBe(DOCUMENT_OP_REFUSAL_REASONS.CONTAINER_FINAL_MARK);
  });
});

describe("C6: a tracked split's new paragraph has exactly the direct split's fields", () => {
  test("nothing resolved is added; a property change records the source's own formatting", () => {
    const document = documentOf(source6(), paragraph("00000009", [run("next")]));
    const split: DocumentOp = {
      type: DOCUMENT_OP_TYPES.SPLIT_BLOCK,
      at: at("00000001", 3),
      newBlockId: "00000002",
      newParagraph: { formatting: { styleId: "BodyText" } },
      revision: stamp(1),
      newIds: { revision: [2] },
    };
    const tracked = blocks(applied(document, split).document);
    const direct = blocks(applied(document, directly(split)).document);
    const [trackedFirst, trackedSecond] = tracked;
    const [directFirst, directSecond] = direct;
    // The new paragraph is the first half here: it records the change and the mark.
    expect(trackedFirst).toEqual({
      ...directFirst,
      propertyChanges: [
        {
          type: "paragraphPropertyChange",
          info: { id: 1, author: "Reviewer", date: DATE },
          previousFormatting: { styleId: "Heading1", spaceBefore: 120 },
        },
      ],
      pPrMark: { kind: "ins", info: { id: 2, author: "Reviewer", date: DATE } },
    });
    expect(trackedSecond).toEqual(directSecond);
  });
});

const source6 = (): Paragraph =>
  paragraph("00000001", [run("abcdef")], {
    formatting: { styleId: "Heading1", spaceBefore: 120, runProperties: { bold: true } },
  });

describe("C5: undoing a tracked split after a later edit never drops a mark silently", () => {
  const original = paragraph("00000001", [run("HelloWorld")], {
    pPrMark: { kind: "del", info: { id: 5, author: "Other" } },
  });
  const document = documentOf(original, paragraph("00000009", [run("next")]));
  const split: DocumentOp = {
    type: DOCUMENT_OP_TYPES.SPLIT_BLOCK,
    at: at("00000001", 5),
    newBlockId: "00000002",
    revision: stamp(1),
  };
  const format: DocumentOp = {
    type: DOCUMENT_OP_TYPES.SET_PARAGRAPH_PROPS,
    story: OP_STORIES.MAIN,
    blockId: "00000001",
    patch: { alignment: "center" },
  };

  test("undoing in reverse order restores exactly", () => {
    const first = applied(document, split);
    const second = applied(first.document, format);
    const undone = applyDocumentOps(second.document, [...second.inverse, ...first.inverse]);
    if (undone.isErr()) throw undone.error;
    expect(undone.value.document).toStrictEqual(document);
  });

  test("undoing the split alone after an edit its join would drop is refused as stale", () => {
    const first = applied(document, split);
    const second = applied(first.document, format);
    const undone = applyDocumentOps(second.document, first.inverse);
    expect(undone.isErr() && undone.error.reason).toBe(DOCUMENT_OP_REFUSAL_REASONS.STALE);
  });

  test("undoing the split alone after an edit of its text keeps the edit and the mark", () => {
    const first = applied(document, split);
    const typed = applied(first.document, {
      type: DOCUMENT_OP_TYPES.INSERT_TEXT,
      at: at("00000001", 5),
      text: "!",
      runProps: "inherit",
    });
    const undone = applyDocumentOps(typed.document, first.inverse);
    if (undone.isErr()) throw undone.error;
    expect(blocks(undone.value.document)).toEqual([
      { ...original, content: [run("HelloWorld!")] },
      paragraph("00000009", [run("next")]),
    ]);
  });
});
