/**
 * What counts as truncation, and what counts as a crash.
 *
 * Both distinctions decide whether a file's gating findings are believed, so
 * both are easy to get wrong in the direction that hides a defect: a file
 * wrongly called truncated has its real failures discarded, and a worker that
 * died wrongly called slow lets the gate pass over a crash.
 */

import { describe, expect, test } from "bun:test";

import { buildBodySequenceDocx } from "@stll/folio-core/compare/__fixtures__/body-sequence";
import { parseDocx } from "@stll/folio-core/docx/parser";

import { INVARIANT_ORDER, runExtendedChecks } from "./lib/corpus-extended";
import {
  EXTENDED_CORPUS_INVARIANTS,
  familyOf,
  isGatingFamily,
} from "./lib/corpus-invariants/contract";
import { CORPUS_INVARIANTS } from "./lib/corpus-signature";

describe("extended invariant registration", () => {
  test("the execution order contains every runnable invariant exactly once", () => {
    const expected = Object.values(EXTENDED_CORPUS_INVARIANTS).filter(
      (invariant) => invariant !== EXTENDED_CORPUS_INVARIANTS.performance,
    );
    expect([...INVARIANT_ORDER].sort()).toEqual(expected.sort());
  });
});

describe("a skipped operation invariant cannot silently pass", () => {
  test("a file budget exhausted before the runners records both missing verdicts", async () => {
    const empty = new Uint8Array(0);
    const result = await runExtendedChecks({
      bytes: empty,
      buffer: empty.buffer as ArrayBuffer,
      // No runner should receive this fixture: the exhausted budget skips every stage.
      parsed: { package: {} } as never,
      documentPart: "word/document.xml",
      invariantBudgetMs: 30_000,
      fileBudgetMs: -1,
    });
    const missing = result.failures.filter(
      (failure) =>
        failure.message ===
        "the file time budget prevented an operation invariant verdict",
    );
    expect(missing.map((failure) => failure.invariant).sort()).toEqual(
      [
        EXTENDED_CORPUS_INVARIANTS.opInverse,
        EXTENDED_CORPUS_INVARIANTS.opLocality,
      ].sort(),
    );
    expect(result.truncatedAt).toBe(EXTENDED_CORPUS_INVARIANTS.opInverse);
  });
});

describe("a stage that merely ran slowly does not truncate the file", () => {
  // An invariant budget of zero makes every stage overrun, while a file budget
  // large enough that nothing is ever skipped.
  test("every invariant still runs, and the file is not truncated", async () => {
    const empty = new Uint8Array(0);
    const result = await runExtendedChecks({
      bytes: empty,
      buffer: empty.buffer as ArrayBuffer,
      // The invariants fail on this input; what matters here is that the loop
      // runs them all and reports no truncation.
      parsed: { package: {} } as never,
      documentPart: "word/document.xml",
      invariantBudgetMs: 0,
      fileBudgetMs: Number.MAX_SAFE_INTEGER,
    });

    expect(result.truncatedAt).toBeUndefined();
    const overruns = result.failures.filter((failure) =>
      failure.message.includes("exceeded its per-file time budget"),
    );
    expect(overruns.length).toBeGreaterThan(0);
    // Nothing was skipped: no stage reports the file budget running out.
    expect(
      result.failures.some((failure) =>
        failure.message.includes("file time budget was exhausted"),
      ),
    ).toBe(false);
  });
});

describe("the two ways a worker stops answering are not the same finding", () => {
  test("a watchdog expiry reports as timing, so it cannot gate", () => {
    expect(
      isGatingFamily(familyOf(EXTENDED_CORPUS_INVARIANTS.performance)),
    ).toBe(false);
  });

  test("a worker that exits without a verdict still gates", () => {
    expect(isGatingFamily(familyOf(CORPUS_INVARIANTS.completes))).toBe(true);
  });
});

describe("an operation invariant that spends its budget stops between steps", () => {
  test("it reports no verdict, never a pass, and the file's later invariants still run", async () => {
    const buffer = await buildBodySequenceDocx([
      { kind: "paragraph", text: "First paragraph." },
      { kind: "paragraph", text: "Final paragraph." },
    ]);
    const result = await runExtendedChecks({
      bytes: new Uint8Array(buffer),
      buffer,
      parsed: await parseDocx(buffer, { preloadFonts: false }),
      documentPart: "word/document.xml",
      // Spent before the first sequence: each operation invariant stops there.
      invariantBudgetMs: -1,
      fileBudgetMs: Number.MAX_SAFE_INTEGER,
      only: new Set([
        EXTENDED_CORPUS_INVARIANTS.opInverse,
        EXTENDED_CORPUS_INVARIANTS.opLocality,
      ]),
    });
    const noVerdict = result.failures.filter(
      (failure) =>
        failure.message ===
        "the invariant time budget stopped it before an operation invariant verdict",
    );
    expect(noVerdict.map((failure) => failure.invariant).sort()).toEqual(
      [
        EXTENDED_CORPUS_INVARIANTS.opInverse,
        EXTENDED_CORPUS_INVARIANTS.opLocality,
      ].sort(),
    );
    expect(
      result.failures
        .filter(
          (failure) =>
            failure.invariant === EXTENDED_CORPUS_INVARIANTS.performance,
        )
        .map((failure) => failure.message)
        .sort(),
    ).toEqual(
      [
        `${EXTENDED_CORPUS_INVARIANTS.opInverse} stopped at its per-file time budget after N of N sequences`,
        `${EXTENDED_CORPUS_INVARIANTS.opLocality} stopped at its per-file time budget after N of N sequences`,
      ].sort(),
    );
    expect(result.truncatedAt).toBeUndefined();
  });
});
