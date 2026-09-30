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
import { contractViolation, normalizeForOps } from "../contract";
import { planTrackedDeletion, revisionIdDemand } from "../plan";
import { mergeAtSeam } from "../resolve";
import { DOCUMENT_OP_REFUSAL_REASONS } from "../refusal";
import {
  DOCUMENT_OP_TYPES,
  type DocumentOp,
  OP_STORIES,
  REVISION_DECISIONS,
  type RevisionDecision,
  type RevisionStamp,
} from "../types";

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

const resolved = (
  document: Document,
  revisionIds: readonly number[],
  decision: RevisionDecision,
): Document =>
  applied(document, {
    type: DOCUMENT_OP_TYPES.RESOLVE_REVISION,
    story: OP_STORIES.MAIN,
    revisionIds,
    decision,
  }).document;

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
  // The L1 generator found an empty revision acting as a zero-width deletion
  // target. The input normalization must remove it for both operation modes.
  test.each(["insertion", "deletion", "moveFrom", "moveTo"] as const)(
    "deleting text beside an empty %s agrees after tracked acceptance",
    (type) => {
      const document = normalizeForOps(
        documentOf(
          paragraph("00000001", [{ type, info: { id: 1, author: "A" }, content: [] }, run("x")]),
        ),
      );
      const op = {
        type: DOCUMENT_OP_TYPES.DELETE_RANGE,
        from: { ...at("00000001", 0), zeroWidthBefore: 0 },
        to: at("00000001", 1),
        revision: stamp(2),
      } as const satisfies DocumentOp;
      const tracked = applied(document, op);
      const direct = applied(document, directly(op));
      expect(resolved(tracked.document, tracked.revisions, REVISION_DECISIONS.ACCEPT)).toEqual(
        direct.document,
      );
      expect(blocks(direct.document)).toEqual([paragraph("00000001", [])]);
    },
  );

  test("a tracked insertion is wrapped and accepts to the direct insertion", () => {
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
    const direct = applied(document, directly(insert));
    expect(resolved(tracked.document, [1], REVISION_DECISIONS.ACCEPT)).toStrictEqual(
      direct.document,
    );
    expect(resolved(tracked.document, [1], REVISION_DECISIONS.REJECT)).toStrictEqual(document);
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
    expect(resolved(tracked.document, [6], REVISION_DECISIONS.REJECT)).toStrictEqual(document);
  });

  // The L1 generator found an accepted insertion whose hyperlink merged into
  // the next one while the seam inside was left unmerged. Resolution merges
  // the seam inside the records it merges as it merges the outer one: the
  // inserted wrapper held nothing before, so it stays, as it does directly.
  test("accepting an insertion merges the seams inside the records it merges", () => {
    const wrapper = (content: ParagraphContent[]): ParagraphContent => ({
      type: "inlineWrapper",
      kind: "bidi",
      control: "embedding",
      content,
    });
    const link = (children: ParagraphContent[]): ParagraphContent => ({
      type: "hyperlink",
      rId: "rId9",
      href: "https://example.org",
      children,
    });
    const document = documentOf(paragraph("00000001", [link([wrapper([run("a")])])]));
    const op = {
      type: DOCUMENT_OP_TYPES.INSERT_CONTENT,
      at: at("00000001", 0),
      slice: { content: [link([wrapper([])])], openStart: 0, openEnd: 0 },
      revision: stamp(1),
    } as const satisfies DocumentOp;
    const tracked = applied(document, op);
    const direct = applied(document, directly(op));
    const accepted = resolved(tracked.document, tracked.revisions, REVISION_DECISIONS.ACCEPT);
    const merged = link([wrapper([]), wrapper([run("a")])]);
    expect(blocks(accepted)).toEqual([paragraph("00000001", [merged])]);
    // Directly, the inserted hyperlink stands beside the other; merged, they agree.
    const [only] = blocks(direct.document);
    const [first, second] = only?.type === "paragraph" ? only.content : [];
    expect(first).toEqual(link([wrapper([])]));
    expect(second !== undefined && mergeAtSeam(link([wrapper([])]), second)).toEqual([merged]);
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
    expect(resolved(tracked.document, [10, 11, 12], REVISION_DECISIONS.REJECT)).toStrictEqual(
      document,
    );
    expect(resolved(tracked.document, [10, 11, 12, 6], REVISION_DECISIONS.ACCEPT)).toStrictEqual(
      documentOf(paragraph("00000001", [])),
    );
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
    expect(resolved(tracked.document, [1, 2], REVISION_DECISIONS.REJECT)).toStrictEqual(document);
    const direct = applied(document, directly(patch));
    expect(resolved(tracked.document, [1, 2], REVISION_DECISIONS.ACCEPT)).toStrictEqual(
      direct.document,
    );
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
    expect(blocks(resolved(tracked.document, [4], REVISION_DECISIONS.REJECT))).toEqual([
      paragraph("00000001", [run("ab", { italic: true })]),
    ]);
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
    expect(resolved(tracked.document, [1], REVISION_DECISIONS.REJECT)).toStrictEqual(document);
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
    const direct = applied(document, directly(split));
    expect(resolved(tracked.document, [1], REVISION_DECISIONS.ACCEPT)).toStrictEqual(
      direct.document,
    );
    expect(resolved(tracked.document, [1], REVISION_DECISIONS.REJECT)).toStrictEqual(document);
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
    const direct = applied(document, directly(join));
    expect(resolved(tracked.document, [1, 2], REVISION_DECISIONS.ACCEPT)).toStrictEqual(
      direct.document,
    );
    expect(resolved(tracked.document, [1, 2], REVISION_DECISIONS.REJECT)).toStrictEqual(document);
    // A mark already carrying a change is not overwritten.
    expect(refusalOf(tracked.document, { ...join, revision: stamp(3) })).toBe(
      DOCUMENT_OP_REFUSAL_REASONS.REVISION_CONFLICT,
    );
  });

  test("accepting a join of an empty paragraph leaves the next one as it was", () => {
    const empty = paragraph("00000001", [{ type: "bookmarkStart", id: 1, name: "_Ref" }], {
      formatting: { styleId: "Heading1" },
    });
    const next = paragraph("00000002", [run("World")], { formatting: { alignment: "end" } });
    const document = documentOf(empty, next, paragraph("00000003", [run("end")]));
    const tracked = applied(document, {
      type: DOCUMENT_OP_TYPES.JOIN_BLOCKS,
      story: OP_STORIES.MAIN,
      blockId: "00000001",
      nextBlockId: "00000002",
      newIds: { revision: [2] },
      revision: stamp(1),
    });
    expect(blocks(resolved(tracked.document, [1, 2], REVISION_DECISIONS.ACCEPT))).toEqual([
      { ...next, content: [{ type: "bookmarkStart", id: 1, name: "_Ref" }, run("World")] },
      paragraph("00000003", [run("end")]),
    ]);
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

  test("joining at a section break is refused", () => {
    const ends = paragraph("00000001", [run("a")], {
      sectionProperties: { pageWidth: 12240 },
      pPrMark: { kind: "del", info: { id: 1, author: "Other" } },
    });
    const document = documentOf(ends, paragraph("00000002", [run("b")]));
    expect(
      refusalOf(document, {
        type: DOCUMENT_OP_TYPES.RESOLVE_REVISION,
        story: OP_STORIES.MAIN,
        revisionIds: [1],
        decision: REVISION_DECISIONS.ACCEPT,
      }),
    ).toBe(DOCUMENT_OP_REFUSAL_REASONS.UNTRACKABLE);
    // Rejecting keeps the break, which needs no section operation.
    expect(blocks(resolved(document, [1], REVISION_DECISIONS.REJECT))).toEqual([
      paragraph("00000001", [run("a")], { sectionProperties: { pageWidth: 12240 } }),
      paragraph("00000002", [run("b")]),
    ]);
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

describe("a tracked split's new second half keeps the paragraph mark's run properties", () => {
  const pending = {
    type: "paragraphPropertyChange" as const,
    info: { id: 1, author: "Other" },
    previousFormatting: { alignment: "end" as const },
  };
  const source = (fields: Partial<Paragraph>): Paragraph =>
    paragraph("00000001", [run("abc")], {
      formatting: { styleId: "Heading1", runProperties: { bold: true } },
      ...fields,
    });
  const splitAtEnd = (formatting: Paragraph["formatting"]): DocumentOp => ({
    type: DOCUMENT_OP_TYPES.SPLIT_BLOCK,
    at: at("00000001", 3),
    newBlockId: "00000002",
    newParagraph: { formatting },
    revision: stamp(2),
    newIds: { revision: [3] },
  });

  test("other mark run properties are refused: rejecting could not give them back", () => {
    for (const fields of [{}, { propertyChanges: [pending] }]) {
      const document = documentOf(source(fields), paragraph("00000009", [run("next")]));
      for (const formatting of [
        { styleId: "Heading1", runProperties: { italic: true } },
        { styleId: "Heading1" },
        { styleId: "Heading1", runProperties: { bold: true }, runInWithNext: true },
      ]) {
        expect(refusalOf(document, splitAtEnd(formatting))).toBe(
          DOCUMENT_OP_REFUSAL_REASONS.UNTRACKABLE,
        );
      }
    }
  });

  test("rejecting a split with the same mark run properties gives the paragraph back", () => {
    for (const fields of [{}, { propertyChanges: [pending] }]) {
      const original = source(fields);
      const document = documentOf(original, paragraph("00000009", [run("next")]));
      const split = applied(
        document,
        splitAtEnd({ styleId: "Heading1", runProperties: { bold: true } }),
      );
      const [rejected, next] = blocks(
        resolved(split.document, split.revisions, REVISION_DECISIONS.REJECT),
      );
      // The mark ending the paragraph is the new half's, under its id.
      expect(rejected).toEqual({ ...original, paraId: "00000002" });
      expect(next).toEqual(paragraph("00000009", [run("next")]));
    }
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

  test("undoing a tracked edit after its paragraph's mark was resolved is refused as stale", () => {
    const first = applied(document, split);
    // The first half is the new paragraph, which carries the inserted mark.
    const formatted = applied(first.document, {
      ...format,
      blockId: "00000002",
      revision: stamp(3),
    });
    const accepted = resolved(formatted.document, [1], REVISION_DECISIONS.ACCEPT);
    const undone = applyDocumentOps(accepted, formatted.inverse);
    expect(undone.isErr() && undone.error.reason).toBe(DOCUMENT_OP_REFUSAL_REASONS.STALE);
  });
});

describe("an empty content control is content: resolution never merges it away", () => {
  const control = (id: number, content: ParagraphContent[]): ParagraphContent => ({
    type: "inlineSdt",
    properties: { sdtType: "richText", id, tag: "clause" },
    content,
  });

  test("rejecting a tracked deletion keeps an empty control beside an alike one", () => {
    const document = documentOf(
      paragraph("00000001", [control(1, []), control(2, [run("ab")])]),
      paragraph("00000009", [run("next")]),
    );
    const op = {
      type: DOCUMENT_OP_TYPES.DELETE_RANGE,
      from: { ...at("00000001", 0), zeroWidthBefore: 0 },
      to: at("00000001", 1),
      revision: stamp(1),
      newIds: { revision: [2] },
    } as const satisfies DocumentOp;
    const tracked = applied(document, op);
    expect(resolved(tracked.document, tracked.revisions, REVISION_DECISIONS.REJECT)).toStrictEqual(
      document,
    );
    expect(resolved(tracked.document, tracked.revisions, REVISION_DECISIONS.ACCEPT)).toStrictEqual(
      applied(document, directly(op)).document,
    );
  });

  test("rejecting a tracked split keeps an empty control that ends the first half", () => {
    const document = documentOf(
      paragraph("00000001", [control(1, []), control(2, [run("a")])]),
      paragraph("00000009", [run("next")]),
    );
    const tracked = applied(document, {
      type: DOCUMENT_OP_TYPES.SPLIT_BLOCK,
      at: { ...at("00000001", 0), zeroWidthBefore: 1 },
      newBlockId: "00000002",
      revision: stamp(1),
    });
    expect(resolved(tracked.document, tracked.revisions, REVISION_DECISIONS.REJECT)).toStrictEqual(
      document,
    );
  });

  test("a piece of a cut control that resolution empties still goes back into it", () => {
    const document = documentOf(
      paragraph("00000001", [control(1, [run("a")])]),
      paragraph("00000009", [run("next")]),
    );
    const typed = applied(document, {
      type: DOCUMENT_OP_TYPES.INSERT_TEXT,
      at: at("00000001", 0),
      text: "x",
      runProps: {},
      revision: stamp(1),
    });
    // The field goes between the control's two pieces, the first holding only the insertion.
    const field = applied(typed.document, {
      type: DOCUMENT_OP_TYPES.INSERT_CONTENT,
      at: at("00000001", 1),
      slice: {
        content: [{ type: "simpleField", instruction: "PAGE", content: [run("3")] }],
        openStart: 0,
        openEnd: 0,
      },
      revision: { ...stamp(2), date: "2026-05-06T07:08:10Z" },
      newIds: { control: [7] },
    });
    const ids = [...typed.revisions, ...field.revisions];
    expect(resolved(field.document, ids, REVISION_DECISIONS.REJECT)).toStrictEqual(document);
  });

  // Two alike containers a resolution merges meet inside too; a piece there
  // is folded back only if the resolution emptied it.
  test("an emptied piece inside merged containers goes back into its neighbour", () => {
    const nested = (content: ParagraphContent[]): ParagraphContent =>
      control(1, [control(2, content)]);
    const document = documentOf(
      paragraph("00000001", [nested([run("ab")])]),
      paragraph("00000009", [run("next")]),
    );
    const typed = applied(document, {
      type: DOCUMENT_OP_TYPES.INSERT_TEXT,
      at: at("00000001", 1),
      text: "x",
      runProps: {},
      revision: stamp(1),
    });
    // Split on each side of the insertion: in the middle paragraph the inner
    // control's piece holds only it, inside a piece of the outer control, so
    // rejecting empties the inner piece, not the paragraph's first record.
    const first = applied(typed.document, {
      type: DOCUMENT_OP_TYPES.SPLIT_BLOCK,
      at: at("00000001", 1),
      newBlockId: "00000002",
      revision: stamp(2),
      newIds: { control: [7, 17] },
    });
    const second = applied(first.document, {
      type: DOCUMENT_OP_TYPES.SPLIT_BLOCK,
      at: at("00000001", 1),
      newBlockId: "00000003",
      revision: stamp(3),
      newIds: { control: [8, 18] },
    });
    const ids = [...typed.revisions, ...first.revisions, ...second.revisions];
    expect(resolved(second.document, ids, REVISION_DECISIONS.REJECT)).toStrictEqual(document);
  });

  // The L3 generator accepted a deletion that emptied one piece of a cut
  // wrapper before the insertion cutting it was resolved: kept, the empty
  // piece later met the other one and stayed beside it, where accepting
  // everything at once folded it in. Accepting leaves no record it empties.
  test("accepting in turn and at once agree on a piece an acceptance empties", () => {
    const wrapper = (content: ParagraphContent[]): ParagraphContent => ({
      type: "inlineWrapper",
      kind: "bidi",
      control: "embedding",
      content,
    });
    const later = (id: number, second: number): RevisionStamp => ({
      ...stamp(id),
      date: `2026-05-06T07:08:${String(second).padStart(2, "0")}Z`,
    });
    const document = documentOf(paragraph("00000001", [wrapper([run("ab")])]));
    const deleted = applied(document, {
      type: DOCUMENT_OP_TYPES.DELETE_RANGE,
      from: at("00000001", 0),
      to: at("00000001", 1),
      revision: later(1, 10),
    });
    // A closed slice cuts the wrapper: its first piece holds only the deletion.
    const inserted = applied(deleted.document, {
      type: DOCUMENT_OP_TYPES.INSERT_CONTENT,
      at: at("00000001", 1),
      slice: { content: [run("x")], openStart: 0, openEnd: 0 },
      revision: later(2, 11),
    });
    const removed = applied(inserted.document, {
      type: DOCUMENT_OP_TYPES.DELETE_RANGE,
      from: at("00000001", 1),
      to: at("00000001", 2),
      revision: later(3, 12),
    });
    expect(blocks(removed.document)).toEqual([
      paragraph("00000001", [
        wrapper([{ type: "deletion", info: later(1, 10), content: [run("a")] }]),
        {
          type: "insertion",
          info: later(2, 11),
          content: [{ type: "deletion", info: later(3, 12), content: [run("x")] }],
        },
        wrapper([run("b")]),
      ]),
    ]);
    const runs = [deleted.revisions, inserted.revisions, removed.revisions];
    const atOnce = resolved(removed.document, runs.flat(), REVISION_DECISIONS.ACCEPT);
    let inTurn = removed.document;
    for (const ids of runs) inTurn = resolved(inTurn, ids, REVISION_DECISIONS.ACCEPT);
    expect(blocks(atOnce)).toEqual([paragraph("00000001", [wrapper([run("b")])])]);
    expect(inTurn).toStrictEqual(atOnce);
  });

  // The L3 generator pasted a slice open at its start inside a control: the
  // pasted content went into the control's first piece, the rest of the
  // control into a new piece after it. Rejecting emptied the first piece with
  // no change resolved between the two, so it stayed beside the other.
  test("rejecting a paste that cut a control gives the control back", () => {
    const document = documentOf(
      paragraph("00000001", [control(1, [run("ab")])]),
      paragraph("00000009", [run("next")]),
    );
    const typed = applied(document, {
      type: DOCUMENT_OP_TYPES.INSERT_TEXT,
      at: at("00000001", 0),
      text: "xx",
      runProps: {},
      revision: stamp(1),
    });
    // Pasted inside the typed insertion, the slice open at its start.
    const typing = { type: "insertion", info: stamp(1), content: [run("y")] } as const;
    const pasted = applied(typed.document, {
      type: DOCUMENT_OP_TYPES.INSERT_CONTENT,
      at: at("00000001", 1),
      slice: { content: [control(1, [typing])], openStart: 4, openEnd: 0 },
      revision: { ...stamp(2), date: "2026-05-06T07:08:10Z" },
      newIds: { revision: [3], control: [7] },
    });
    const typing3 = { ...typing, info: stamp(3), content: [run("x")] };
    expect(blocks(pasted.document)).toEqual([
      paragraph("00000001", [
        control(1, [
          { ...typing, content: [run("x")] },
          { ...typing, info: { ...stamp(2), date: "2026-05-06T07:08:10Z" } },
        ]),
        control(7, [typing3, run("ab")]),
      ]),
      paragraph("00000009", [run("next")]),
    ]);
    // The typing is revision 1 and the piece of it the paste cut off, 3.
    const typings = [...typed.revisions, 3];
    const ids = [...typings, ...pasted.revisions];
    expect(resolved(pasted.document, ids, REVISION_DECISIONS.REJECT)).toStrictEqual(document);
    let inTurn = pasted.document;
    for (const turn of [pasted.revisions, typings]) {
      inTurn = resolved(inTurn, turn, REVISION_DECISIONS.REJECT);
    }
    expect(inTurn).toStrictEqual(document);
  });

  // A container a resolution empties meets its neighbours whatever its kind,
  // however deep the emptied piece sits at the edge, and whichever of two
  // pieces empties first.
  describe("pieces of any container a rejection empties go back together", () => {
    const wrapper = (content: ParagraphContent[]): ParagraphContent => ({
      type: "inlineWrapper",
      kind: "bidi",
      control: "embedding",
      content,
    });
    const link = (children: ParagraphContent[]): ParagraphContent => ({
      type: "hyperlink",
      rId: "rId9",
      href: "https://example.org",
      children,
    });
    const inserted = (id: number, text: string): ParagraphContent => ({
      type: "insertion",
      info: stamp(id),
      content: [run(text)],
    });
    const kinds: [string, (id: number, content: ParagraphContent[]) => ParagraphContent][] = [
      ["a content control", control],
      ["a bidi wrapper", (_, content) => wrapper(content)],
      ["a hyperlink holding a bidi wrapper", (_, content) => link([wrapper(content)])],
    ];
    for (const [name, piece] of kinds) {
      test(name, () => {
        const withNext = (content: ParagraphContent[]): Document =>
          documentOf(paragraph("00000001", content), paragraph("00000009", [run("next")]));
        const cut = withNext([piece(1, [inserted(1, "x")]), piece(7, [run("ab")])]);
        expect(resolved(cut, [1], REVISION_DECISIONS.REJECT)).toStrictEqual(
          withNext([piece(1, [run("ab")])]),
        );
        // Both pieces emptied, at once or one after the other, leave one.
        const both = withNext([piece(1, [inserted(1, "x")]), piece(7, [inserted(2, "ab")])]);
        const one = withNext([piece(1, [])]);
        expect(resolved(both, [1, 2], REVISION_DECISIONS.REJECT)).toStrictEqual(one);
        for (const order of [
          [1, 2],
          [2, 1],
        ]) {
          let inTurn = both;
          for (const id of order) inTurn = resolved(inTurn, [id], REVISION_DECISIONS.REJECT);
          expect(inTurn).toStrictEqual(one);
        }
      });
    }
  });
});
