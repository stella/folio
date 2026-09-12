import { describe, expect, test } from "bun:test";

import type { CompareDocxOptions } from "@stll/folio-core/compare/types";

import { compareDocx } from "../../packages/core/src/compare/compare";
import { buildDocumentPackage } from "./documents";
import { zipPackage } from "./package-xml";
import { compareCorpusPairInBothModes } from "./refusal-modes";
import { applyVariant } from "./variants";

describe("corpus refusal modes", () => {
  test("runs strict and best-effort as distinct public policies", async () => {
    const base = new ArrayBuffer(1);
    const target = new ArrayBuffer(2);
    const options = {
      author: "benchmark",
      timestamp: "2000-01-01T00:00:00.000Z",
    } as const satisfies CompareDocxOptions;
    const calls: CompareDocxOptions[] = [];

    const outcomes = await compareCorpusPairInBothModes({
      base,
      target,
      options,
      compare: async (receivedBase, receivedTarget, receivedOptions) => {
        expect(receivedBase).toBe(base);
        expect(receivedTarget).toBe(target);
        calls.push(receivedOptions);
        return calls.length;
      },
    });

    expect(outcomes).toEqual({ strict: 1, bestEffort: 2 });
    expect(calls).toEqual([options, { ...options, mode: "bestEffort" }]);
  });

  test("reaches both policies on a known-unverified corpus pair", async () => {
    const baseParts = buildDocumentPackage({ documentClass: "lists", size: "s" });
    const targetParts = applyVariant({ parts: baseParts, variant: "numbering" });
    if (targetParts === null) {
      throw new Error("The lists fixture did not carry a numbering definition");
    }

    const { strict, bestEffort } = await compareCorpusPairInBothModes({
      compare: compareDocx,
      base: await zipPackage(baseParts),
      target: await zipPackage(targetParts),
      options: {
        author: "benchmark",
        timestamp: "2000-01-01T00:00:00.000Z",
      },
    });

    expect(strict.isErr()).toBe(true);
    if (strict.isErr()) {
      expect(strict.error._tag).toBe("CompareDocxUnsupportedError");
    }
    expect(bestEffort.isOk()).toBe(true);
    if (bestEffort.isOk()) {
      expect(bestEffort.value.verification.status).toBe("unverified");
      expect(bestEffort.value.unsupported.map(({ reason }) => reason)).toContain(
        "numbering-definition",
      );
    }
  });
});
