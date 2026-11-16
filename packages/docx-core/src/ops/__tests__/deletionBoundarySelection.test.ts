import { panic } from "better-result";
import { afterEach, expect, mock, spyOn, test } from "bun:test";

import type { Document, ParagraphContent } from "../../model/document";
import { applyDocumentOp, applyDocumentOps } from "../apply";
import { normalizeForOps } from "../contract";
import * as deletionBoundaries from "../deletionBoundaries";
import { childNodes, type InlineNode, leafPlaces } from "../leaves";
import { DOCUMENT_OP_TYPES, OP_STORIES } from "../types";

afterEach(() => mock.restore());

const run = (text: string): ParagraphContent => ({
  type: "run",
  content: [{ type: "text", text }],
});

const producerChain = (length: number) => {
  const original = normalizeForOps({
    package: {
      document: {
        content: [
          {
            type: "paragraph",
            paraId: "00000001",
            content: [
              {
                type: "hyperlink",
                anchor: "source",
                children: [
                  {
                    type: "inlineWrapper",
                    kind: "bidi",
                    control: "embedding",
                    content: [
                      {
                        type: "moveToRangeStart",
                        id: 1,
                        name: "_Move1",
                        author: "Author",
                      },
                      run("a"),
                      {
                        type: "insertion",
                        info: { id: 5, author: "Author" },
                        content: [
                          run("bcdef"),
                          { type: "moveToRangeEnd", id: 1 },
                          run("tail"),
                        ],
                      },
                    ],
                  },
                ],
              },
            ],
          },
        ],
      },
    },
  } satisfies Document);
  let document = original;
  let targetId = 5;
  const deletionIds: number[] = [];
  for (let index = 0; index < length; index++) {
    const revisionId = 7 + index;
    const applied = applyDocumentOp(document, {
      type: DOCUMENT_OP_TYPES.DELETE_RANGE,
      from: {
        story: OP_STORIES.MAIN,
        blockId: "00000001",
        offset: index === 0 ? 0 : index + 1,
        zeroWidthBefore: 0,
      },
      to: { story: OP_STORIES.MAIN, blockId: "00000001", offset: index + 2 },
      revision: { id: revisionId, author: "Reviewer" },
      newIds: {
        revision: Array.from({ length: 20 }, (_, n) => 1000 + index * 100 + n),
        control: Array.from({ length: 20 }, (_, n) => 2000 + index * 100 + n),
      },
    });
    if (applied.isErr()) throw applied.error;
    document = applied.value.document;
    deletionIds.push(...applied.value.revisions);
    const block = document.package.document.content.at(0);
    if (block?.type !== "paragraph")
      panic("The producer keeps its source paragraph.");
    const targets = new Set<number>();
    const visit = (nodes: readonly InlineNode[]): void => {
      for (const node of nodes) {
        if (node.type === "deletion" && node.info.id === revisionId)
          for (const transfer of node.resolutionJoins?.retainedAfter ?? [])
            for (const slot of transfer.target)
              if (slot.space === "revision") targets.add(slot.id);
        visit(childNodes(node) ?? []);
      }
    };
    visit(block.content);
    expect(targets.size).toBe(1);
    targetId =
      [...targets].at(0) ??
      panic("A produced handoff has its target revision.");
  }
  return { document, deletionIds, targetId };
};

for (const length of [0, 1, 2, 3])
  test(`boundary planning reaches the actual identity-selection fixed point (${length} producer handoffs)`, () => {
    const { document, deletionIds, targetId } = producerChain(length);
    const original = structuredClone(document);
    const planner = spyOn(deletionBoundaries, "keptDeletionBoundaryPaths");
    const operation = {
      type: DOCUMENT_OP_TYPES.RESOLVE_REVISION,
      story: OP_STORIES.MAIN,
      revisionIds: [...deletionIds, targetId],
      decision: "accept",
    } as const;
    const resolved = applyDocumentOp(document, operation);
    if (resolved.isErr()) throw resolved.error;
    expect(planner).toHaveBeenCalledTimes(length === 0 ? 1 : 2);
    planner.mockClear();
    const expanded = applyDocumentOp(document, {
      ...operation,
      revisionIds: [...operation.revisionIds, 5],
    });
    if (expanded.isErr()) throw expanded.error;
    expect(expanded.value.document).toStrictEqual(resolved.value.document);
    expect(planner).toHaveBeenCalledTimes(1);
    const block = resolved.value.document.package.document.content.at(0);
    if (block?.type !== "paragraph")
      panic("Resolution keeps its source paragraph.");
    expect(
      leafPlaces(block.content).flatMap(({ node }) =>
        node.type === "moveToRangeStart" || node.type === "moveToRangeEnd"
          ? [node.type]
          : [],
      ),
    ).toEqual(["moveToRangeStart", "moveToRangeEnd"]);
    const repeated = applyDocumentOp(resolved.value.document, operation);
    if (repeated.isErr()) throw repeated.error;
    expect(repeated.value.inverse).toEqual([]);
    const undone = applyDocumentOps(
      resolved.value.document,
      JSON.parse(JSON.stringify(resolved.value.inverse)),
    );
    if (undone.isErr()) throw undone.error;
    expect(undone.value.document).toStrictEqual(document);
    expect(document).toStrictEqual(original);
  });
