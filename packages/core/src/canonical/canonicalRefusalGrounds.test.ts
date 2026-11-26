import { expect, test } from "bun:test";

import { DOCUMENT_OP_REFUSAL_REASONS } from "@stll/docx-core/ops";

import {
  CANONICAL_CAPABILITIES,
  CANONICAL_CONDITION,
  CANONICAL_GAP,
  CANONICAL_RULE,
} from "../types/canonicalCapabilities";
import { groundForDocumentOpRefusal } from "./canonicalRefusalGrounds";

const R = DOCUMENT_OP_REFUSAL_REASONS;

test("a moved-on document, an exhausted id space, a link suggestion and an untrackable change keep their own grounds", () => {
  expect(groundForDocumentOpRefusal(R.STALE)).toEqual({
    condition: CANONICAL_CONDITION.stale,
  });
  expect(groundForDocumentOpRefusal(R.SECTIONS_OUT_OF_STEP)).toEqual({
    condition: CANONICAL_CONDITION.stale,
  });
  expect(groundForDocumentOpRefusal(R.NEEDS_NEW_IDS)).toEqual({
    condition: CANONICAL_CONDITION.idSpace,
  });
  expect(groundForDocumentOpRefusal(R.UNRECORDABLE_LINK_SUGGESTION)).toEqual({
    rule: CANONICAL_RULE.linkSuggestionShape,
  });
  expect(groundForDocumentOpRefusal(R.UNTRACKABLE)).toEqual({
    gap: CANONICAL_GAP.trackedChangeRecords,
  });
});

test("every other document operation refusal is a request the document does not allow", () => {
  const named = new Set<string>([
    R.STALE,
    R.SECTIONS_OUT_OF_STEP,
    R.NEEDS_NEW_IDS,
    R.UNRECORDABLE_LINK_SUGGESTION,
    R.UNTRACKABLE,
  ]);
  const others = Object.values(R).filter((reason) => !named.has(reason));
  expect(others.length).toBeGreaterThan(0);
  for (const reason of others)
    expect(groundForDocumentOpRefusal(reason)).toEqual({
      condition: CANONICAL_CONDITION.invalidRequest,
    });
});

test("the gap a document operation refusal names is in the ledger", () => {
  for (const reason of Object.values(R)) {
    const { gap } = groundForDocumentOpRefusal(reason);
    if (gap !== undefined)
      expect(Object.hasOwn(CANONICAL_CAPABILITIES, gap)).toBe(true);
  }
});
