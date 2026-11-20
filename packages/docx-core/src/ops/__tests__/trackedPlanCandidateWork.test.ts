import { expect, spyOn, test } from "bun:test";

import type { Document, Paragraph } from "../../model/document";
import { validateOpsDocument } from "../contract";
import * as identityKeys from "../ids";
import { createTrackedPlan } from "../plan";
import { DOCUMENT_OP_TYPES, OP_STORIES } from "../types";

const revision = { id: 5, author: "Reviewer", date: "2026-01-01T00:00:00Z" };
const paraId = (offset: number) => (offset + 1).toString(16).padStart(8, "0");

test.each([16, 256])(
  "planning with %i occupied ids examines only supplied candidates",
  (width) => {
    const content = Array.from({ length: width }, (_, offset): Paragraph => ({
      type: "paragraph",
      paraId: paraId(offset),
      content: [
        {
          type: "insertion",
          info: { ...revision, id: 1_000 + offset },
          content: [{ type: "run", content: [{ type: "text", text: "kept" }] }],
        },
      ],
    }));
    const document: Document = { package: { document: { content } } };
    const validation = validateOpsDocument(document);
    if (validation.isErr()) throw validation.error;
    const census = spyOn(identityKeys, "packageIdentityKeys");
    try {
      const pool = [revision.id, 40, 40, 41, 1_000 + width - 1];
      const plan = createTrackedPlan({
        document,
        revision,
        newIds: { revision: pool, control: [30] },
      });
      for (const offset of [0, 1, 2]) {
        const result = plan.append({
          type: DOCUMENT_OP_TYPES.SET_PARAGRAPH_PROPS,
          story: OP_STORIES.MAIN,
          blockId: paraId(offset),
          patch: { alignment: "center" },
          revision,
        });
        if (result.isErr()) throw result.error;
      }
      expect(plan.ops.map((op) => "revision" in op && op.revision?.id)).toEqual(
        [revision.id, 41, 40],
      );
      expect(census).not.toHaveBeenCalled();
    } finally {
      census.mockRestore();
    }
  },
);
