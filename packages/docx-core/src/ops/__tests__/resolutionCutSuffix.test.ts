import { expect, test } from "bun:test";

import type { Document } from "../../model/document";
import { applyDocumentOp, applyDocumentOps } from "../apply";
import {
  DOCUMENT_OP_TYPES,
  OP_STORIES,
  REVISION_DECISIONS,
  type SplitBlockOp,
} from "../types";

test.each([32, 96])(
  "rejecting %i tracked splits reuses the physical paragraph-list suffix",
  (count) => {
    const blockId = "00000001";
    const original: Document = {
      package: {
        document: {
          content: [
            {
              type: "paragraph",
              paraId: blockId,
              content: [
                { type: "run", content: [{ type: "text", text: "tail" }] },
              ],
            },
          ],
        },
      },
    };
    const splits = Array.from(
      { length: count },
      (_, index) =>
        ({
          type: DOCUMENT_OP_TYPES.SPLIT_BLOCK,
          at: { story: OP_STORIES.MAIN, blockId, offset: 0 },
          newBlockId: (0x1000_0000 + index).toString(16),
          revision: { id: index + 1, author: "Reviewer" },
        }) satisfies SplitBlockOp,
    );
    const pending = applyDocumentOps(original, splits).unwrap();
    const content = pending.document.package.document.content;
    const marks = content.flatMap((block) =>
      block.type === "paragraph" && block.pPrMark !== undefined
        ? [block.pPrMark]
        : [],
    );
    expect(marks).toHaveLength(count);
    expect(marks.every(({ resolutionJoin }) => resolutionJoin === 0)).toBe(
      true,
    );

    let visits = 0;
    const observed = new Proxy(content, {
      get(target, property, receiver): unknown {
        if (typeof property === "string" && /^[0-9]+$/u.test(property))
          visits += 1;
        return Reflect.get(target, property, receiver);
      },
    });
    const measured = {
      ...pending.document,
      package: {
        ...pending.document.package,
        document: { ...pending.document.package.document, content: observed },
      },
    };
    const rejected = applyDocumentOp(measured, {
      type: DOCUMENT_OP_TYPES.RESOLVE_REVISION,
      story: OP_STORIES.MAIN,
      revisionIds: pending.revisions,
      decision: REVISION_DECISIONS.REJECT,
    }).unwrap();
    // Native validation and inverse capture also read this list. The budget
    // allows their linear work and rejects a walk of every qualifying suffix.
    expect(visits).toBeLessThan(count * 40);
    expect(rejected.document).toStrictEqual(original);
    expect(
      applyDocumentOps(rejected.document, rejected.inverse).unwrap().document,
    ).toStrictEqual(pending.document);
  },
);
