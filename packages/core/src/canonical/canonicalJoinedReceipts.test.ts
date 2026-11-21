import { expect, test } from "bun:test";

import type { FolioAIEditAppliedOperation } from "../ai-edits/types";
import { addJoinedParagraphReceipts } from "./canonicalJoinedReceipts";

test("each operation reads its address once across independent and shared joins", () => {
  const count = 128;
  const ids = Array.from({ length: count + 1 }, (_, i) =>
    (0x10000000 + i).toString(16),
  );
  const survivor = "10000000";
  const retired = ids.slice(1);
  const absorptions = retired.map((absorbed) => ({ absorbed, survivor }));
  let reads = 0;
  const operations = ids.map((id) => ({
    id,
    type: "deleteBlock" as const,
    get blockId() {
      reads += 1;
      return id;
    },
  }));
  const applied = new Map<string, FolioAIEditAppliedOperation>(
    ids.map((id) => [id, { id }]),
  );
  const retained = "10000001";
  applied.set(retained, {
    id: retained,
    paragraphIdentity: { type: "retained", blockId: retained },
  });
  const skipped = "10000002";
  applied.delete(skipped);

  addJoinedParagraphReceipts({ absorptions: [], operations, applied });
  expect(reads).toBe(0);
  addJoinedParagraphReceipts({ absorptions, operations, applied });
  expect(reads).toBeLessThanOrEqual(operations.length);
  expect(applied.get(survivor)?.paragraphIdentity).toEqual({
    type: "joined",
    blockId: survivor,
    retiredBlockId: retained,
  });
  expect(applied.get(retained)?.paragraphIdentity).toEqual({
    type: "retained",
    blockId: retained,
  });
  expect(applied.has(skipped)).toBe(false);
  for (const id of retired.slice(2)) {
    expect(applied.get(id)?.paragraphIdentity).toEqual({
      type: "joined",
      blockId: survivor,
      retiredBlockId: id,
    });
  }
});
