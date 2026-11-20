import { describe, expect, test } from "bun:test";

import type { Document, ParagraphContent } from "../../model/document";
import { contractViolation } from "../contract";
import { settleKeptCuts } from "../keptCut";
import type { InlineNode } from "../leaves";

const CUT = { depth: 2, keptFrom: "2026-03-04T05:00:01Z" };

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
  test("the contract admits a cut at least one level deep and refuses another", () => {
    const marked = (depth: number) =>
      documentWith([
        link("a"),
        { type: "bookmarkEnd", id: 1, keptCut: { ...CUT, depth } },
        link("b"),
      ]);
    expect(contractViolation(marked(2))).toBeUndefined();
    expect(contractViolation(marked(0))).toBeDefined();
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

  test("markers standing together while a change waits stand in the deepest cut any records", () => {
    const pending: InlineNode = {
      type: "insertion",
      info: { id: 4, author: "C", date: "2026-03-04T05:00:02Z" },
      content: [link("x")],
    };
    const shallow = { depth: 1, keptFrom: "2026-03-04T05:00:03Z" };
    const settled = settleKeptCuts([
      pending,
      { type: "bookmarkStart", id: 1, name: "b", keptCut: CUT },
      { type: "bookmarkEnd", id: 1, keptCut: shallow },
      { type: "bookmarkEnd", id: 2 },
    ]);
    // Each keeps the revision it was kept from; the authored marker gains no cut.
    expect(settled.slice(1)).toEqual([
      { type: "bookmarkStart", id: 1, name: "b", keptCut: CUT },
      { type: "bookmarkEnd", id: 1, keptCut: { ...shallow, depth: CUT.depth } },
      { type: "bookmarkEnd", id: 2 },
    ]);
  });
});
