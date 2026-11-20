import type { ParagraphAbsorption } from "@stll/docx-core/ops";

import type { FolioAIEditAppliedOperation } from "../ai-edits/types";
import type { FolioDocumentOperation } from "../document-operations";

type JoinedReceiptOptions = {
  absorptions: readonly ParagraphAbsorption[] | undefined;
  operations: readonly FolioDocumentOperation[];
  applied: ReadonlyMap<string, FolioAIEditAppliedOperation>;
};

/** A receipt names the first join that absorbed or kept its addressed paragraph. */
export const addJoinedParagraphReceipts = ({
  absorptions,
  operations,
  applied,
}: JoinedReceiptOptions): void => {
  if (!absorptions?.length) return;
  const firstJoin = new Map<string, ParagraphAbsorption>();
  for (const join of absorptions) {
    if (!firstJoin.has(join.absorbed)) firstJoin.set(join.absorbed, join);
    if (!firstJoin.has(join.survivor)) firstJoin.set(join.survivor, join);
  }
  for (const operation of operations) {
    const entry = applied.get(operation.id);
    if (
      entry === undefined ||
      entry.paragraphIdentity !== undefined ||
      !("blockId" in operation)
    )
      continue;
    const join = firstJoin.get(operation.blockId);
    if (join === undefined) continue;
    entry.paragraphIdentity = {
      type: "joined",
      blockId: join.survivor,
      retiredBlockId: join.absorbed,
    };
  }
};
