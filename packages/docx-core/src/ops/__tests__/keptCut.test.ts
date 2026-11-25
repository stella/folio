import { describe, expect, test } from "bun:test";

import type {
  Document,
  KeptCutHolders,
  ParagraphContent,
} from "../../model/document";
import { contractViolation } from "../contract";
import { isKeptCut, settleKeptCuts } from "../keptCut";
import type { InlineNode } from "../leaves";

const CUT = { keptFrom: "2026-03-04T05:00:01Z" };

const link = (text: string) =>
  ({
    type: "hyperlink",
    rId: "rId9",
    href: "https://example.org",
    children: [{ type: "run", content: [{ type: "text", text }] }],
  }) satisfies ParagraphContent;

const documentWith = (content: ParagraphContent[]): Document => ({
  package: {
    document: { content: [{ type: "paragraph", paraId: "00000001", content }] },
  },
});

const cuts = (nodes: readonly InlineNode[]): unknown[] =>
  nodes.flatMap((node) =>
    node.type === "bookmarkEnd" && node.keptCut !== undefined
      ? [node.keptCut]
      : [],
  );

describe("a kept range marker's cut", () => {
  test("the contract admits a cut with exactly its fields and refuses another", () => {
    const marked = (cut: typeof CUT) =>
      documentWith([
        link("a"),
        { type: "bookmarkEnd", id: 1, keptCut: cut },
        link("b"),
      ]);
    expect(contractViolation(marked(CUT))).toBeUndefined();
    const other = { ...CUT, side: "before" };
    expect(contractViolation(marked(other))).toBeDefined();
  });

  test("a cut names a revision on either side, both, or neither", () => {
    expect(isKeptCut(CUT)).toBe(true);
    expect(isKeptCut({ ...CUT, holders: { before: 0 } })).toBe(true);
    expect(isKeptCut({ ...CUT, holders: { after: 4 } })).toBe(true);
    expect(isKeptCut({ ...CUT, holders: { before: 0, after: 4 } })).toBe(true);
    expect(isKeptCut({ ...CUT, holders: {} })).toBe(false);
    expect(isKeptCut({ ...CUT, holders: { after: -1 } })).toBe(false);
    expect(isKeptCut({ ...CUT, piece: "after" })).toBe(false);
  });

  test("stays while a pending change waits beside the markers, and goes once none does", () => {
    const marker: InlineNode = { type: "bookmarkEnd", id: 1, keptCut: CUT };
    const pending: InlineNode = {
      type: "insertion",
      info: { id: 4, author: "C", date: "2026-03-04T05:00:02Z" },
      content: [{ type: "run", content: [{ type: "text", text: "x" }] }],
    };
    const before = link("a");
    const after = link("b");
    expect(cuts(settleKeptCuts([before, marker, pending, after]))).toEqual([
      CUT,
    ]);
    expect(cuts(settleKeptCuts([before, marker, after]))).toEqual([]);
  });

  test("inside a record, stays while a pending change meets the markers across its edge", () => {
    const pending: InlineNode = {
      type: "insertion",
      info: { id: 4, author: "C", date: "2026-03-04T05:00:02Z" },
      content: [link("x")],
    };
    const holding = (): ParagraphContent => ({
      type: "hyperlink",
      rId: "rId9",
      href: "https://example.org",
      children: [
        { type: "bookmarkEnd", id: 1, keptCut: CUT },
        { type: "run", content: [{ type: "text", text: "a" }] },
      ],
    });
    const inner = (nodes: readonly InlineNode[]): unknown[] =>
      nodes.flatMap((node) =>
        node.type === "hyperlink" ? cuts(node.children) : [],
      );
    expect(inner(settleKeptCuts([pending, holding()]))).toEqual([CUT]);
    expect(inner(settleKeptCuts([link("b"), holding()]))).toEqual([]);
  });

  test("a piece stays named while the pending change meeting the markers stands at the near edge of a record beside them", () => {
    const pending: InlineNode = {
      type: "insertion",
      info: { id: 4, author: "C", date: "2026-03-04T05:00:02Z" },
      content: [{ type: "run", content: [{ type: "text", text: "x" }] }],
    };
    const holding: ParagraphContent = {
      type: "hyperlink",
      rId: "rId9",
      href: "https://example.org",
      children: [
        pending,
        { type: "run", content: [{ type: "text", text: "a" }] },
      ],
    };
    const marker = (holders: KeptCutHolders): InlineNode => ({
      type: "bookmarkEnd",
      id: 1,
      keptCut: { ...CUT, holders },
    });
    // The cut and its piece wait on the same change.
    expect(
      cuts(settleKeptCuts([link("b"), marker({ after: 4 }), holding])),
    ).toEqual([{ ...CUT, holders: { after: 4 } }]);
    expect(
      cuts(settleKeptCuts([link("b"), marker({ after: 5 }), holding])),
    ).toEqual([CUT]);
    expect(
      cuts(
        settleKeptCuts([link("b"), marker({ before: 3, after: 4 }), holding]),
      ),
    ).toEqual([{ ...CUT, holders: { after: 4 } }]);
    // A piece naming a side where no change waits is consumed.
    expect(
      cuts(settleKeptCuts([link("b"), marker({ before: 4 }), holding])),
    ).toEqual([CUT]);
  });

  test("markers standing together while a change waits keep the revision each was kept from", () => {
    const pending: InlineNode = {
      type: "insertion",
      info: { id: 4, author: "C", date: "2026-03-04T05:00:02Z" },
      content: [link("x")],
    };
    const later = { keptFrom: "2026-03-04T05:00:03Z" };
    const settled = settleKeptCuts([
      pending,
      { type: "bookmarkStart", id: 1, name: "b", keptCut: CUT },
      { type: "bookmarkEnd", id: 1, keptCut: later },
      { type: "bookmarkEnd", id: 2 },
    ]);
    // The authored marker gains no cut.
    expect(settled.slice(1)).toEqual([
      { type: "bookmarkStart", id: 1, name: "b", keptCut: CUT },
      { type: "bookmarkEnd", id: 1, keptCut: later },
      { type: "bookmarkEnd", id: 2 },
    ]);
  });
});
