import { expect, test } from "bun:test";
import fc from "fast-check";

import { assertProperty } from "../../../../../test/property-testing";
import type {
  BlockContent,
  Document,
  Paragraph,
  ParagraphContent,
} from "../../model/document";
import { applyDocumentOp } from "../apply";
import { childNodes, type InlineNode } from "../leaves";
import { DOCUMENT_OP_TYPES, OP_STORIES, REVISION_DECISIONS } from "../types";

// Resolving a revision never moves text across a range's boundary: what a
// bookmark covers can only lose text that goes, and what it leaves out can
// only lose text that goes. Each paragraph's text is its own letter, so a
// letter counted on the wrong side of a range is text that crossed it.

const info = (id: number) => ({ id, author: "A", date: "2026-10-23" });
const run = (text: string): ParagraphContent => ({
  type: "run",
  content: [{ type: "text", text }],
});
const start = (id: number) =>
  ({ type: "bookmarkStart", id, name: `_Ref${id}` }) as const;
const end = (id: number) => ({ type: "bookmarkEnd", id }) as const;

const wrap = (
  content: ParagraphContent,
  kind: "none" | "insertion" | "deletion",
): ParagraphContent => {
  if (kind === "none") return content;
  return {
    type: kind,
    info: info(kind === "insertion" ? 10 : 11),
    content: [content],
  };
};

type Sides = Map<
  number,
  { inside: Map<string, number>; outside: Map<string, number> }
>;

/** For each bookmark with one start and one end, the letters inside and outside it. */
const coverage = (document: Document): Sides => {
  const starts = new Map<number, number>();
  const ends = new Map<number, number>();
  const texts: { open: ReadonlySet<number>; text: string }[] = [];
  const open = new Set<number>();
  const marker = (node: { type: string; id?: number }): void => {
    if (node.type === "bookmarkStart" && node.id !== undefined) {
      starts.set(node.id, (starts.get(node.id) ?? 0) + 1);
      open.add(node.id);
    }
    if (node.type === "bookmarkEnd" && node.id !== undefined) {
      ends.set(node.id, (ends.get(node.id) ?? 0) + 1);
      open.delete(node.id);
    }
  };
  const inline = (nodes: readonly InlineNode[]): void => {
    for (const node of nodes) {
      marker(node);
      if (node.type === "run")
        for (const item of node.content)
          if (item.type === "text")
            texts.push({ open: new Set(open), text: item.text });
      inline(node.type === "run" ? [] : (childNodes(node) ?? []));
    }
  };
  const blocks = (list: readonly BlockContent[]): void => {
    for (const block of list) {
      marker(block);
      if (block.type === "paragraph") inline(block.content);
      if (block.type === "table")
        for (const row of block.rows)
          for (const cell of row.cells) blocks(cell.content);
    }
  };
  blocks(document.package.document.content);
  const sides: Sides = new Map();
  for (const [id, count] of starts) {
    if (count !== 1 || ends.get(id) !== 1) continue;
    const inside = new Map<string, number>();
    const outside = new Map<string, number>();
    for (const { open: around, text } of texts)
      for (const letter of text) {
        const side = around.has(id) ? inside : outside;
        side.set(letter, (side.get(letter) ?? 0) + 1);
      }
    sides.set(id, { inside, outside });
  }
  return sides;
};

const placement = fc.constantFrom("none", "start", "end");

const MARKERS = {
  none: [],
  start: [start(1)],
  end: [end(1)],
} as const;

const markersAt = (where: keyof typeof MARKERS) => MARKERS[where];

test("resolving a revision never moves text across a range boundary", () => {
  assertProperty(
    fc.property(
      fc.constantFrom("none", "ins", "del"),
      fc.constantFrom("none", "insertion", "deletion"),
      placement,
      fc.constantFrom("none", "start", "end"),
      placement,
      placement,
      fc.constantFrom(REVISION_DECISIONS.ACCEPT, REVISION_DECISIONS.REJECT),
      (mark, wrapper, inFirst, between, inCell, inLast, decision) => {
        const first: Paragraph = {
          type: "paragraph",
          paraId: "00000001",
          content: [...markersAt(inFirst), wrap(run("aaa"), wrapper)],
          ...(mark === "none"
            ? {}
            : { pPrMark: { kind: mark, info: info(12) } }),
        };
        const document: Document = {
          package: {
            document: {
              content: [
                first,
                ...markersAt(between),
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
                              paraId: "00000002",
                              content: [run("bb"), ...markersAt(inCell)],
                            },
                          ],
                        },
                      ],
                    },
                  ],
                },
                {
                  type: "paragraph",
                  paraId: "00000003",
                  content: [run("cc"), ...markersAt(inLast)],
                },
              ],
            },
          },
        };
        const before = coverage(document);
        fc.pre(before.size > 0);
        const ids = [
          ...(mark === "none" ? [] : [12]),
          ...(wrapper === "none" ? [] : [wrapper === "insertion" ? 10 : 11]),
        ];
        fc.pre(ids.length > 0);
        const applied = applyDocumentOp(document, {
          type: DOCUMENT_OP_TYPES.RESOLVE_REVISION,
          story: OP_STORIES.MAIN,
          revisionIds: ids,
          decision,
        });
        if (applied.isErr()) return;
        const after = coverage(applied.value.document);
        for (const [id, sides] of after) {
          const was = before.get(id);
          if (was === undefined) continue;
          for (const side of ["inside", "outside"] as const)
            for (const [letter, count] of sides[side])
              expect({ id, side, letter, count }).toEqual({
                id,
                side,
                letter,
                count: Math.min(count, was[side].get(letter) ?? 0),
              });
        }
      },
    ),
    {
      numRuns: 300,
      id: "resolving a revision never moves text across a range boundary",
    },
  );
});
