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
