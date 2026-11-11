import { expect, type Page } from "@playwright/test";

import { parseDocx } from "../../packages/core/src/docx/parser";
import { createDocx } from "../../packages/core/src/docx/rezip";
import { createEmptyDocument } from "../../packages/core/src/utils/createDocument";
import {
  clearRefusals,
  expectNoRefusals,
  loadReady,
  MODIFIER,
  reactPort,
  snapshot,
  vuePort,
} from "../parity/canonicalInteractionAssertions";
import { canonicalLoadFixture } from "../parity/canonicalLoadFixture";
import { test } from "../parity/canonicalTest";

/**
 * Dragging a column border or a row border resizes the table in the
 * canonical session, in both hosts: the saved grid and row height change,
 * nothing is refused, and undo restores the grid.
 */

const cell = (text: string) => ({
  type: "tableCell" as const,
  content: [
    {
      type: "paragraph" as const,
      content: [
        { type: "run" as const, content: [{ type: "text" as const, text }] },
      ],
    },
  ],
});

const savedTable = async (page: Page) => {
  const saved = await page.evaluate(async () => {
    const buffer = await globalThis.__folioCanonical?.save();
    return buffer ? [...new Uint8Array(buffer)] : null;
  });
  if (!saved) throw new TypeError("The canonical document did not save.");
  const document = await parseDocx(new Uint8Array(saved), {
    preloadFonts: false,
    detectVariables: false,
  });
  const table = document.package.document.content.find(
    (block) => block.type === "table",
  );
  if (table?.type !== "table") throw new TypeError("The table is absent.");
  return {
    grid: table.columnWidths ?? [],
    firstRowHeight: table.rows[0]?.formatting?.height?.value,
  };
};

/**
 * Drag a handle from a point of its own: the column and row border handles
 * cross where the borders meet, so the grab avoids the middle of either.
 */
const drag = async (
  page: Page,
  selector: string,
  grab: { x: number; y: number },
  delta: { dx: number; dy: number },
) => {
  const handle = page.locator(selector).first();
  const box = await handle.boundingBox();
  if (!box) throw new TypeError(`No handle ${selector}.`);
  const x = box.x + box.width * grab.x;
  const y = box.y + box.height * grab.y;
  expect(
    await page.evaluate(
      ({ px, py, wanted }) =>
        document.elementFromPoint(px, py)?.matches(wanted) ?? false,
      { px: x, py: y, wanted: selector },
    ),
    `the grab point is on ${selector}`,
  ).toBe(true);
  await page.mouse.move(x, y);
  await page.mouse.down();
  await page.mouse.move(x + delta.dx / 2, y + delta.dy / 2);
  await page.mouse.move(x + delta.dx, y + delta.dy);
  await page.mouse.up();
};

test("dragging a column border or a row border resizes the table, and undo restores it", async ({
  page,
}) => {
  test.setTimeout(90_000);
  const document = createEmptyDocument({ initialText: "Before" });
  document.package.document.content.push({
    type: "table",
    columnWidths: [2000, 2000],
    rows: [
      { type: "tableRow", cells: [cell("A"), cell("B")] },
      { type: "tableRow", cells: [cell("C"), cell("D")] },
    ],
  });
  const source = await canonicalLoadFixture(await createDocx(document));
  for (const [host, port] of [
    ["React", reactPort],
    ["Vue", vuePort],
  ] as const) {
    await page.goto(`http://localhost:${port}/?session=canonical`);
    await page.waitForSelector(".layout-page");
    await clearRefusals(page);
    await loadReady(page, source);
    const before = await savedTable(page);
    expect(before.grid, host).toEqual([2000, 2000]);

    await drag(
      page,
      '.layout-table-resize-handle[data-column-index="0"]',
      { x: 0.5, y: 0.15 },
      { dx: 40, dy: 0 },
    );
    await expectNoRefusals(page);
    expect((await snapshot(page))?.projectionMatchesCanonical, host).toBe(true);
    const widened = await savedTable(page);
    const [left = 0, right = 0] = widened.grid;
    expect([host, left > 2000, right < 2000, left + right], host).toEqual([
      host,
      true,
      true,
      4000,
    ]);

    // Back in the text, undo takes the resize back.
    await page
      .locator(".layout-page .layout-run-text", { hasText: "Before" })
      .first()
      .click();
    await page.keyboard.press(`${MODIFIER}+z`);
    await expectNoRefusals(page);
    expect((await savedTable(page)).grid, host).toEqual([2000, 2000]);

    await drag(
      page,
      '.layout-table-row-resize-handle[data-row-index="0"]',
      { x: 0.1, y: 0.5 },
      { dx: 0, dy: 30 },
    );
    await expectNoRefusals(page);
    expect((await snapshot(page))?.projectionMatchesCanonical, host).toBe(true);
    const taller = await savedTable(page);
    expect(
      [host, (taller.firstRowHeight ?? 0) > (before.firstRowHeight ?? 0)],
      host,
    ).toEqual([host, true]);
  }
});
