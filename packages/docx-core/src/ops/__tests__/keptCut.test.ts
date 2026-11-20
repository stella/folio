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
});
