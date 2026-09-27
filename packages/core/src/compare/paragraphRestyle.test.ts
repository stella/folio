import { describe, expect, test } from "bun:test";

import { FolioDocxReviewer } from "../ai-edits/headless";
import { createDocx } from "../docx/rezip";
import { fromMarkdown } from "../markdown/fromMarkdown";
import { compareDocx } from "./compare";

const OPTIONS = { author: "compare", timestamp: "2026-09-27T00:00:00.000Z" } as const;

const markdownDocx = (markdown: string): Promise<ArrayBuffer> => createDocx(fromMarkdown(markdown));

const restyleProjection = (reviewer: FolioDocxReviewer) =>
  reviewer
    .getContent()
    .map(({ kind, headingLevel, styleId, directOutlineLevel, text, previewRuns }) => ({
      kind,
      headingLevel,
      styleId,
      directOutlineLevel,
      text,
      previewRuns,
    }));

const expectVerifiedRestyle = async (
  base: ArrayBuffer,
  target: ArrayBuffer,
  kinds: readonly string[] = ["paragraph-format"],
): Promise<void> => {
  const compared = await compareDocx(base, target, OPTIONS);
  if (compared.isErr()) throw compared.error;
  expect(compared.value.verification).toEqual({ status: "verified" });
  expect(compared.value.unsupported).toEqual([]);
  expect(compared.value.changes.map(({ kind }) => kind)).toEqual(kinds);

  const accepting = await FolioDocxReviewer.fromBuffer(compared.value.buffer);
  expect(accepting.acceptAll()).toBeGreaterThan(0);
  const accepted = await FolioDocxReviewer.fromBuffer(await accepting.toBuffer());
  const expected = await FolioDocxReviewer.fromBuffer(target);
  expect(restyleProjection(accepted)).toEqual(restyleProjection(expected));

  const rejecting = await FolioDocxReviewer.fromBuffer(compared.value.buffer);
  expect(rejecting.rejectAll()).toBeGreaterThan(0);
  const rejected = await FolioDocxReviewer.fromBuffer(await rejecting.toBuffer());
  const original = await FolioDocxReviewer.fromBuffer(base);
  expect(restyleProjection(rejected)).toEqual(restyleProjection(original));
};

describe("comparison paragraph restyles", () => {
  test("verifies a plain-to-heading Markdown restyle without a run change", async () => {
    await expectVerifiedRestyle(
      await markdownDocx("Scope\n\nBody."),
      await markdownDocx("# Scope\n\nBody."),
    );
  });

  test("verifies a plain-to-heading editor restyle without a run change", async () => {
    const base = await markdownDocx("Scope\n\nBody.");
    const reviewer = await FolioDocxReviewer.fromBuffer(base);
    const blockId = reviewer.snapshot().blocks.at(0)?.id;
    expect(blockId).toBeDefined();
    if (blockId === undefined) return;
    const applied = reviewer.applyOperations(
      [
        {
          id: "restyle",
          type: "setBlockParagraphProperties",
          blockId,
          properties: { styleId: "Heading1" },
        },
      ],
      { mode: "direct" },
    );
    expect(applied.skipped).toEqual([]);
    await expectVerifiedRestyle(base, await reviewer.toBuffer());
  });

  test.each([
    { label: "heading-to-plain", base: "# Scope\n\nBody.", target: "Scope\n\nBody." },
    { label: "Heading1-to-Heading2", base: "# Scope\n\nBody.", target: "## Scope\n\nBody." },
  ])("verifies a $label restyle without a run change", async ({ base, target }) => {
    await expectVerifiedRestyle(await markdownDocx(base), await markdownDocx(target));
  });

  test("keeps an authored bold change alongside a paragraph restyle", async () => {
    await expectVerifiedRestyle(
      await markdownDocx("Scope\n\nBody."),
      await markdownDocx("# **Scope**\n\nBody."),
      ["paragraph-format", "format"],
    );
  });
});

describe("comparison insertions beside a heading that states its outline level", () => {
  // Markdown headings carry a direct `w:outlineLvl`. An inserted paragraph
  // that did not state its own took the anchor's, so an ordinary paragraph
  // placed before such a heading became a heading once accepted.
  test.each([
    {
      label: "an inserted paragraph before a heading",
      base: "# Scope\n\nBody.",
      target: "Preamble.\n\n# Scope\n\nBody.",
      kinds: ["insert"],
    },
    {
      label: "an inserted paragraph between body text and a heading",
      base: "Intro.\n\n# Scope\n\nBody.",
      target: "Intro.\n\nPreamble.\n\n# Scope\n\nBody.",
      kinds: ["insert"],
    },
    {
      label: "sections swapped before a trailing paragraph",
      base: "Intro.\n\n## Fees\n\nPay.\n\n## Term\n\nEither.\n\nTail.",
      target: "Intro.\n\n## Term\n\nEither.\n\n## Fees\n\nPay.\n\nTail.",
      kinds: ["delete", "delete", "insert", "insert"],
    },
  ])("keeps $label ordinary once accepted", async ({ base, target, kinds }) => {
    await expectVerifiedRestyle(await markdownDocx(base), await markdownDocx(target), kinds);
  });

  test("keeps a relocated paragraph ordinary beside edited table rows", async () => {
    const table = (rows: string) => `| Item | Price |\n|---|---|\n${rows}`;
    await expectVerifiedRestyle(
      await markdownDocx(
        [
          "# Scope",
          "The supplier provides services.",
          "## Fees",
          "Payment term is 30 days.",
          table("| Service A | 100 |\n| Service B | 200 |"),
          "## Term",
          "Either party may terminate on notice.",
          "The parties will maintain confidentiality.",
        ].join("\n\n"),
      ),
      await markdownDocx(
        [
          "# Scope",
          "The supplier provides services.",
          "## Term",
          "Either party may terminate on notice.",
          "## Fees",
          "Payment term is 45 days.",
          table("| Service A | 150 |\n| Service C | 300 |"),
          "The parties will maintain confidentiality.",
        ].join("\n\n"),
      ),
      ["insert", "move", "replace", "replace", "table-row-delete", "table-row-insert", "delete"],
    );
  });
});
