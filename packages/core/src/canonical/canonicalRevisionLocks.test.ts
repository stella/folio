import { expect, spyOn, test } from "bun:test";

import type { Document, Paragraph } from "@stll/docx-core/model";
import * as ops from "@stll/docx-core/ops";

import { canonicalRevisionLocks } from "./canonicalRevisionLocks";

test.each(["accept", "reject"] as const)(
  "%s probes each revision group only when a source control is locked",
  (decision) => {
    const groups = Array.from({ length: 16 }, (_, i) => [i + 1]);
    const paragraphs = groups.map(([id], i): Paragraph => ({
      type: "paragraph",
      paraId: (0x10000000 + i).toString(16),
      content: [
        {
          type: "insertion",
          info: { id, author: "Reviewer" },
          content: [{ type: "run", content: [{ type: "text", text: "Text" }] }],
        },
      ],
    }));
    const classify = (document: Document) =>
      canonicalRevisionLocks({ document, story: "main", groups, decision });
    const apply = spyOn(ops, "applyDocumentOps");
    try {
      expect(
        classify({ package: { document: { content: paragraphs } } }),
      ).toEqual({
        held: [],
        unlocked: groups,
      });
      expect(apply.mock.calls).toHaveLength(0);
      const first = paragraphs.at(0);
      if (first === undefined) throw new Error("Expected a revision group");
      const locked = classify({
        package: {
          document: {
            content: [
              {
                type: "blockSdt",
                properties: { lock: "contentLocked", id: 100 },
                content: [first],
              },
              ...paragraphs.slice(1),
            ],
          },
        },
      });
      expect(apply.mock.calls).toHaveLength(groups.length);
      expect(locked.held.map(({ ids }) => ids)).toEqual([groups[0]]);
      expect(locked.unlocked).toEqual(groups.slice(1));
    } finally {
      apply.mockRestore();
    }
  },
);
