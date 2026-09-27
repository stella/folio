/**
 * The benchmark's correctness invariants over every generated family.
 *
 * `run.ts` checks accept, reject, self-compare, determinism, revision ids and
 * schema validity for every configuration, but it is a timing run: nothing
 * runs it on a pull request. The hand-picked round-trip tests beside this file
 * cover a few variants of a few classes, so a change that broke one class
 * under one edit shape (the notes class under `churn`, say) could merge green.
 * This runs the same invariants over the smallest size of every class and
 * every edit variant that applies to it, without timing anything.
 */

import { expect, test } from "bun:test";
import { compareDocx } from "@stll/folio-core";

import { buildDocumentPackage, DOCUMENT_CLASSES } from "./documents";
import { checkInvariants } from "./invariants";
import { zipPackage } from "./package-xml";
import { resolvePackageValidator } from "./validator";
import { applyVariant, EDIT_VARIANTS } from "./variants";

const OPTIONS = { author: "folio compare benchmark", timestamp: "2000-01-01T00:00:00.000Z" };

const validate = resolvePackageValidator();

const exercised = new Set<string>();

test.each([...DOCUMENT_CLASSES])(
  "generated %s documents satisfy every invariant under each applicable edit",
  async (documentClass) => {
    const parts = buildDocumentPackage({ documentClass, size: "s" });
    const base = await zipPackage(parts);
    const failures: string[] = [];
    for (const variant of EDIT_VARIANTS) {
      const targetParts = applyVariant({ parts, variant });
      if (targetParts === null) {
        continue;
      }
      exercised.add(variant);
      const id = `${documentClass}/s/${variant}`;
      const target = await zipPackage(targetParts);
      const compared = await compareDocx(base, target, OPTIONS);
      if (compared.isErr()) {
        failures.push(`${id}: ${compared.error.name}: ${compared.error.message}`);
        continue;
      }
      const { outcomes } = await checkInvariants({
        base,
        target,
        redlined: compared.value.buffer,
        changes: compared.value.changes,
        unsupported: compared.value.unsupported.map(({ reason }) => reason),
        expectation: variant === "identical" ? "identical" : "different",
        options: OPTIONS,
        validate,
      });
      for (const outcome of outcomes) {
        if (outcome.status === "failed") {
          failures.push(`${id}: ${outcome.invariant}: ${outcome.detail}`);
        }
      }
    }
    expect(failures).toEqual([]);
  },
  120_000,
);

test("every edit variant applies to at least one generated family", () => {
  expect([...exercised].toSorted()).toEqual([...EDIT_VARIANTS].toSorted());
});
