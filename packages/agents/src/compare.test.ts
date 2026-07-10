/**
 * Version-diff engine tests: real-paraId alignment, the deterministic-fallback-id
 * regression guard (same text, shifted ordinal must still pair as unchanged via
 * the text-LCS pass), identical-buffer no-op, the LLM text format, and the
 * as-accepted semantics over a revised document that carries pending tracked
 * changes.
 *
 * Buffers are built directly from the typed `Document` model (`createEmptyDocument`
 * as a template, `createDocx` to serialize) rather than a fixture file, since each
 * case needs precise control over paragraph text/paraId and no existing corpus
 * fixture has more than two named paragraphs.
 */

import { describe, expect, test } from "bun:test";

import { createDocx, createEmptyDocument, FolioDocxReviewer } from "@stll/folio-core/server";

import { compareDocxVersions, exceedsLcsBudget, formatVersionDiffForLLM } from "./compare";
import type { FolioAgentBlockDiff } from "./compare";

type ParagraphSpec = { text: string; paraId?: string };

const buildDocxBuffer = (paragraphs: readonly ParagraphSpec[]): Promise<ArrayBuffer> => {
  const template = createEmptyDocument();
  return createDocx({
    ...template,
    package: {
      ...template.package,
      document: {
        ...template.package.document,
        content: paragraphs.map(({ text, paraId }) => ({
          type: "paragraph",
          content: [{ type: "run", content: [{ type: "text", text }] }],
          ...(paraId !== undefined && { paraId }),
        })),
      },
    },
  });
};

const findChange = (
  changes: readonly FolioAgentBlockDiff[],
  blockId: string,
): FolioAgentBlockDiff => {
  const change = changes.find((c) => c.blockId === blockId);
  if (!change) {
    throw new Error(`no change for block ${blockId}`);
  }
  return change;
};

describe("compareDocxVersions: real w14:paraId alignment", () => {
  test("classifies unchanged, modified, added, and deleted blocks by stable id", async () => {
    const base = await buildDocxBuffer([
      { text: "Alpha paragraph.", paraId: "00000001" },
      { text: "Beta paragraph.", paraId: "00000002" },
      { text: "Gamma paragraph.", paraId: "00000003" },
      { text: "Zeta paragraph.", paraId: "00000005" },
    ]);
    const revised = await buildDocxBuffer([
      { text: "Alpha paragraph.", paraId: "00000001" },
      { text: "Beta clause.", paraId: "00000002" },
      { text: "Zeta paragraph.", paraId: "00000005" },
      { text: "Epsilon paragraph.", paraId: "00000006" },
    ]);

    const diff = await compareDocxVersions(base, revised);

    expect(diff.summaryCounts).toEqual({ added: 1, deleted: 1, modified: 1, unchanged: 2 });
    // Unchanged blocks (Alpha, Zeta) never appear in `changes`.
    expect(diff.changes).toHaveLength(3);
    expect(diff.changes.map((c) => c.type)).toEqual(["modified", "deleted", "added"]);

    const modified = findChange(diff.changes, "00000002");
    if (modified.type !== "modified") {
      throw new Error("expected a modified change");
    }
    expect(modified.kind).toBe("paragraph");
    expect(modified.segments).toEqual([
      { type: "equal", text: "Beta " },
      { type: "del", text: "paragraph." },
      { type: "ins", text: "clause." },
    ]);

    const deleted = findChange(diff.changes, "00000003");
    expect(deleted).toEqual({
      type: "deleted",
      blockId: "00000003",
      kind: "paragraph",
      text: "Gamma paragraph.",
    });

    const added = findChange(diff.changes, "00000006");
    expect(added).toEqual({
      type: "added",
      blockId: "00000006",
      kind: "paragraph",
      text: "Epsilon paragraph.",
    });
  });
});

describe("compareDocxVersions: deterministic fallback ids (no w14:paraId)", () => {
  test("a same-text block whose ordinal shifted still pairs as unchanged via the text-LCS pass", async () => {
    // Neither buffer carries a w14:paraId, so FolioDocxReviewer assigns each
    // block a deterministic fallback id derived from hash(text + ordinal). The
    // "Epsilon" paragraph inserted before "Gamma" shifts Gamma's ordinal (3rd
    // -> 4th), which changes its fallback id even though its text is
    // untouched. Pass 1 (stable-id pairing) therefore CANNOT pair Gamma across
    // versions; only the text-LCS pass (pass 2) recovers it as unchanged. This
    // is the regression guard for the deterministic-id pitfall.
    const base = await buildDocxBuffer([
      { text: "Alpha paragraph." },
      { text: "Beta paragraph." },
      { text: "Gamma paragraph." },
    ]);
    const revised = await buildDocxBuffer([
      { text: "Alpha paragraph." },
      { text: "Beta paragraph modified." },
      { text: "Epsilon paragraph." },
      { text: "Gamma paragraph." },
    ]);

    const diff = await compareDocxVersions(base, revised);

    // Alpha (unshifted, same fallback id both sides) and Gamma (shifted,
    // recovered via text-LCS) are both unchanged and absent from `changes`.
    expect(diff.summaryCounts).toEqual({ added: 1, deleted: 0, modified: 1, unchanged: 2 });
    expect(diff.changes.map((c) => c.type)).toEqual(["modified", "added"]);

    const modified = diff.changes[0];
    if (!modified || modified.type !== "modified") {
      throw new Error("expected a modified change");
    }
    // Reconstruct the revised text from the `equal` + `ins` segments (the
    // `del` segments carry the superseded base-side text, not part of either
    // whole string on their own).
    const reconstructedRevised = modified.segments
      .filter((s) => s.type !== "del")
      .map((s) => s.text)
      .join("");
    expect(reconstructedRevised).toBe("Beta paragraph modified.");

    const added = diff.changes[1];
    if (!added || added.type !== "added") {
      throw new Error("expected an added change");
    }
    expect(added.text).toBe("Epsilon paragraph.");
  });
});

describe("compareDocxVersions: no-op", () => {
  test("identical documents produce zero changes and every block counts as unchanged", async () => {
    const paragraphs: ParagraphSpec[] = [{ text: "Alpha paragraph." }, { text: "Beta paragraph." }];
    const base = await buildDocxBuffer(paragraphs);
    const revised = await buildDocxBuffer(paragraphs);

    const diff = await compareDocxVersions(base, revised);

    expect(diff.changes).toEqual([]);
    expect(diff.summaryCounts).toEqual({ added: 0, deleted: 0, modified: 0, unchanged: 2 });
  });
});

describe("formatVersionDiffForLLM", () => {
  test("renders a stable header + one line per change, using word-diff porcelain markers", async () => {
    const base = await buildDocxBuffer([
      { text: "Alpha paragraph.", paraId: "00000001" },
      { text: "Beta paragraph.", paraId: "00000002" },
      { text: "Gamma paragraph.", paraId: "00000003" },
      { text: "Zeta paragraph.", paraId: "00000005" },
    ]);
    const revised = await buildDocxBuffer([
      { text: "Alpha paragraph.", paraId: "00000001" },
      { text: "Beta clause.", paraId: "00000002" },
      { text: "Zeta paragraph.", paraId: "00000005" },
      { text: "Epsilon paragraph.", paraId: "00000006" },
    ]);

    const diff = await compareDocxVersions(base, revised);
    const formatted = formatVersionDiffForLLM(diff);
    const lines = formatted.split("\n");

    expect(lines[0]).toBe("Version diff: 1 added, 1 deleted, 1 modified, 2 unchanged");
    expect(lines).toContain("~ [00000002] modified: Beta [-paragraph.-]{+clause.+}");
    expect(lines).toContain("- [00000003] deleted: Gamma paragraph.");
    expect(lines).toContain("+ [00000006] added: Epsilon paragraph.");
  });
});

describe("exceedsLcsBudget: pass 2's LCS cell-budget guard", () => {
  test("flags unpaired-block counts whose product would exceed the LCS cell budget (4,000,000)", () => {
    // A degenerate/adversarial document with no w14:paraIds can leave
    // thousands of blocks unpaired on both sides; this guard is what stops
    // pairByExactText from allocating an O(m*n) table for it. Exercise the
    // predicate directly rather than constructing a multi-million-block
    // fixture, which would make this test slow for no extra coverage.
    expect(exceedsLcsBudget(2000, 2000)).toBe(false); // exactly at budget: 4,000,000 cells
    expect(exceedsLcsBudget(2001, 2001)).toBe(true); // just over budget: 4,004,001 cells
  });
});

describe("compareDocxVersions: as-accepted semantics", () => {
  test("a pending tracked insertion in the revised document counts as already applied", async () => {
    const base = await buildDocxBuffer([{ text: "Payment is due." }]);

    // Build the revised buffer by applying a tracked-changes (default mode)
    // replace against a fresh reviewer over `base` — the revised buffer still
    // carries real w:ins/w:del marks on disk.
    const revisedReviewer = await FolioDocxReviewer.fromBuffer(base, { author: "AI" });
    const target = revisedReviewer.snapshot().blocks[0];
    if (!target) {
      throw new Error("expected a block in the base document");
    }
    revisedReviewer.applyOperations([
      {
        id: "t1",
        type: "replaceInBlock",
        blockId: target.id,
        find: "due.",
        replace: "due promptly.",
      },
    ]);
    const revised = await revisedReviewer.toBuffer();

    // Sanity: the revised snapshot's clean text is already the accepted view.
    expect(revisedReviewer.snapshot().blocks[0]?.text).toBe("Payment is due promptly.");

    const diff = await compareDocxVersions(base, revised);

    expect(diff.summaryCounts).toEqual({ added: 0, deleted: 0, modified: 1, unchanged: 0 });
    const [change] = diff.changes;
    if (!change || change.type !== "modified") {
      throw new Error("expected a single modified change");
    }
    // The diff runs over the as-accepted text: reconstructing the revised
    // side from `equal` + `ins` segments reproduces the clean accepted
    // string, not raw tracked-change markup.
    const reconstructedRevised = change.segments
      .filter((s) => s.type !== "del")
      .map((s) => s.text)
      .join("");
    expect(reconstructedRevised).toBe("Payment is due promptly.");
    expect(change.segments.some((s) => s.type === "ins" && s.text.includes("promptly"))).toBe(true);
  });
});
