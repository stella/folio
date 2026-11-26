import { describe, expect, test } from "bun:test";

import {
  DOCUMENT_OP_REFUSAL_REASONS,
  DocumentOpRefusal,
  type DocumentOpRefusalReason,
} from "@stll/docx-core/ops";

import { compilerRefusalSkip } from "./canonicalCompilerRefusal";
import { groundForDocumentOpRefusal } from "./canonicalRefusalGrounds";

const R = DOCUMENT_OP_REFUSAL_REASONS;

const refusal = (reason: DocumentOpRefusalReason) =>
  new DocumentOpRefusal({
    reason,
    message: `Refused: ${reason}.`,
    opType: undefined,
  });

describe("compiler refusal skips", () => {
  test("a refusal reports the public reason that fits it and its reason's ground", () => {
    expect(compilerRefusalSkip(refusal(R.STALE))).toEqual({
      reason: "staleRange",
      message: "Refused: stale.",
      refusal: { condition: "stale" },
    });
    expect(compilerRefusalSkip(refusal(R.SECTIONS_OUT_OF_STEP)).reason).toBe(
      "staleRange",
    );
    expect(compilerRefusalSkip(refusal(R.SPLITS_SURROGATE_PAIR)).reason).toBe(
      "splitsCharacter",
    );
    expect(compilerRefusalSkip(refusal(R.BLOCK_NOT_FOUND)).reason).toBe(
      "missingBlock",
    );
    expect(
      compilerRefusalSkip(refusal(R.UNRECORDABLE_LINK_SUGGESTION)),
    ).toEqual({
      reason: "unsupportedBlock",
      message: "Refused: unrecordableLinkSuggestion.",
      refusal: { rule: "linkSuggestionShape" },
    });
    expect(compilerRefusalSkip(refusal(R.NEEDS_NEW_IDS)).refusal).toEqual({
      condition: "idSpace",
    });
  });

  test("a deliberate refusal and a no-op report their reason alone", () => {
    expect(compilerRefusalSkip(refusal(R.PENDING_TABLE_CHANGE))).toEqual({
      reason: "pendingTableChange",
      message: "Refused: pendingTableChange.",
    });
    expect(compilerRefusalSkip(refusal(R.NO_CHANGE))).toEqual({
      reason: "noopOperation",
      message: "Refused: noChange.",
    });
  });

  test("every grounded refusal reports the ground of its reason", () => {
    for (const reason of Object.values(R)) {
      const skip = compilerRefusalSkip(refusal(reason));
      if (skip.refusal !== undefined)
        expect(skip.refusal).toEqual(groundForDocumentOpRefusal(reason));
    }
  });
});
