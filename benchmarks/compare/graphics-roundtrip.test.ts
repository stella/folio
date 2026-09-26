import { expect, test } from "bun:test";
import { compareDocx } from "@stll/folio-core";

import { buildDocumentPackage } from "./documents";
import { checkInvariants } from "./invariants";
import { zipPackage } from "./package-xml";
import { applyVariant } from "./variants";

const OPTIONS = { author: "folio compare benchmark", timestamp: "2000-01-01T00:00:00.000Z" };

test.each(["light", "churn", "reorder", "structural", "rewrite"] as const)(
  "graphics %s round trip preserves inline atoms",
  async (variant) => {
    const parts = buildDocumentPackage({ documentClass: "graphics", size: "s" });
    const targetParts = applyVariant({ parts, variant });
    if (!targetParts) throw new Error(`Graphics fixture does not support ${variant}`);
    const base = await zipPackage(parts);
    const target = await zipPackage(targetParts);
    const compared = await compareDocx(base, target, OPTIONS);
    if (compared.isErr()) throw compared.error;
    const result = await checkInvariants({
      base,
      target,
      redlined: compared.value.buffer,
      changes: compared.value.changes,
      unsupported: compared.value.unsupported.map(({ reason }) => reason),
      expectation: "different",
      options: OPTIONS,
      validate: null,
    });
    expect(result.outcomes.filter(({ status }) => status === "failed")).toEqual([]);
  },
);
