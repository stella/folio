import { describe, expect, test } from "bun:test";
import { createHash } from "node:crypto";

import { diffShape, failureMarker } from "../test/consumer-scenarios/support/failure-fingerprints";
import {
  classifyFailureMarkers,
  extractFailureMarkers,
  parseKnownFailures,
} from "./failure-fingerprints";

const marker = (seed: number, failure: Error) =>
  failureMarker({
    test: "comments / suggested",
    seed,
    repro: `FOLIO_SCENARIO_SEED=${seed} bun scripts/consumer-scenarios.ts --only 'comments'`,
    failure,
  });

describe("failure fingerprints", () => {
  test("the same mismatch keeps its fingerprint across seeds and document text", () => {
    const first = marker(
      1250352731,
      new Error(
        'step 8: the result is not what was asked (replaceBlock):\n  no comment {"id":1790619452379,"anchor":"The Buyer pays thirty days."} among []',
      ),
    );
    const second = marker(
      1250359999,
      new Error(
        'step 3: the result is not what was asked (formatRange, replaceBlock):\n  no comment {"id":1790619459911,"anchor":"Le contrat est signé."} among []',
      ),
    );
    expect(first.fingerprint).toBe(second.fingerprint);
    expect(first.assertion).toBe("no comment");
  });

  test("different mismatch kinds have different fingerprints", () => {
    const missingComment = marker(1, new Error("step 1: no comment"));
    const changedReader = marker(
      2,
      new Error("step 1: [readerStability] getChanges reads otherwise after the save"),
    );
    expect(missingComment.fingerprint).not.toBe(changedReader.fingerprint);
  });

  test("a non-error cause does not replace the error message", () => {
    const failure = new Error("step 1: no comment", { cause: "request context" });
    expect(marker(1, failure).assertion).toBe("no comment");
  });

  test("extracts CI-prefixed markers and classifies known fingerprints", () => {
    const first = marker(1, new Error("step 1: no comment"));
    const second = marker(2, new Error("step 2: no comment"));
    const other = marker(3, new Error("block texts differ"));
    const log = [first, second, other]
      .map((entry) => `job 2026-09-28 FOLIO_FAILURE ${JSON.stringify(entry)}`)
      .join("\n");
    const classified = classifyFailureMarkers(extractFailureMarkers(log), [
      { fingerprint: first.fingerprint, issueOrPr: "#1191", firstSeen: "2026-09-28" },
    ]);
    expect(classified).toMatchObject([
      { status: "known", seeds: [1, 2], issueOrPr: "#1191" },
      { status: "new", seeds: [3] },
    ]);
  });

  test("rejects malformed known entries", () => {
    expect(() => parseKnownFailures({ known: [{ fingerprint: "bad" }] })).toThrow(
      "Invalid known failure fingerprint entry",
    );
  });

  describe("the shape of the differences tells two bugs behind one symptom apart", () => {
    const change = (type: string, text: string, blockId: string) =>
      JSON.stringify({ type, author: "Consumer Scenario", text, blockId });
    const readerStability = (lines: string[]) =>
      new Error(
        [
          "step 4: [readerStability] getChanges reads otherwise after the save than before it (before → after):",
          ...lines.map((line) => `    ${line}`),
        ].join("\n"),
      );
    // A tracked column change the reopened package lists once instead of per cell.
    const perCell = (at: number, blockId: string) =>
      readerStability([
        `[${at}].blockId: "${blockId}" → "0F88C890"`,
        `[${at + 1}]: ${change("tableColumnInsertion", "", blockId)} → undefined`,
      ]);
    // A split inline insertion the reopened package lists as one change.
    const split = (at: number, text: string) =>
      readerStability([
        `[${at}].text: "${text}" → "${text} and more"`,
        `[${at + 1}]: ${change("insertion", " and more", "1508FAF4")} → undefined`,
      ]);

    test("same bug, other seeds, positions, ids and text: one fingerprint", () => {
      expect(marker(1, perCell(0, "1108F4A8")).fingerprint).toBe(
        marker(2, perCell(3, "5FCE4B43")).fingerprint,
      );
      expect(marker(1, split(1, "payment")).fingerprint).toBe(
        marker(9, split(4, "delivery term")).fingerprint,
      );
      expect(diffShape(split(1, "payment"))).toBe("[] {insertion}→∅; [].text text→text");
    });

    test("different bugs with the same symptom line: two fingerprints", () => {
      const column = marker(1, perCell(0, "1108F4A8"));
      const inline = marker(1, split(1, "payment"));
      expect(column.assertion).toBe(inline.assertion);
      expect(column.fingerprint).not.toBe(inline.fingerprint);
      expect(column.diff).toBe("[] {tableColumnInsertion}→∅; [].blockId text→text");
    });

    test("a typed entry's type stays in the shape, its position does not", () => {
      const typed = (at: number) =>
        readerStability([
          `[${at}:insertion].text: "payment" → "payment and more"`,
          `.blocks[${at}:paragraph].runs[0].text: "a" → "b"`,
          `(root): "x" → "y"`,
        ]);
      expect(diffShape(typed(2))).toBe(
        "(root) text→text; .blocks[paragraph].runs[].text text→text; [insertion].text text→text",
      );
      expect(marker(1, typed(2)).fingerprint).toBe(marker(2, typed(5)).fingerprint);
      expect(marker(1, typed(2)).fingerprint).not.toBe(
        marker(1, readerStability([`[2:deletion].text: "payment" → "payment and more"`]))
          .fingerprint,
      );
    });

    test("a failure with no differences keeps the fingerprint it always had", () => {
      const plain = marker(1, new Error("step 1: no comment"));
      expect(plain.diff).toBeUndefined();
      expect(plain.fingerprint).toBe(
        createHash("sha256").update("comments / suggested\0no comment").digest("hex").slice(0, 16),
      );
    });

    test("a minimized flow's operations refine the fingerprint and keep the primary one", () => {
      const failure = split(1, "payment");
      const unshrunk = marker(1, failure);
      const shrunk = failureMarker({
        test: "comments / suggested",
        seed: 1,
        repro: "replay",
        failure,
        flow: "insertAfterBlock > replaceRange",
      });
      expect(shrunk.primary).toBe(unshrunk.fingerprint);
      expect(shrunk.fingerprint).not.toBe(unshrunk.fingerprint);
      const [extracted] = extractFailureMarkers(`FOLIO_FAILURE ${JSON.stringify(shrunk)}`);
      expect(extracted).toEqual(shrunk);
    });
  });
});
