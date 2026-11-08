import { expect } from "@playwright/test";

import { parseDocx } from "../../packages/core/src/docx/parser";
import { createDocx } from "../../packages/core/src/docx/rezip";
import { createEmptyDocument } from "../../packages/core/src/utils/createDocument";
import {
  clearRefusals,
  expectNoRefusals,
  loadReady,
  MODIFIER,
  reactPort,
  select,
  snapshot,
  vuePort,
} from "../parity/canonicalInteractionAssertions";
import { canonicalLoadFixture } from "../parity/canonicalLoadFixture";
import { test } from "../parity/canonicalTest";

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

const cellTexts = async (page: Parameters<typeof select>[0]) => {
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
    grid: table.columnWidths?.length,
    rows: table.rows.map((row) =>
      row.cells.map((each) =>
        each.content
          .flatMap((block) =>
            block.type === "paragraph"
              ? block.content.flatMap((item) =>
                  item.type === "run"
                    ? item.content.flatMap((leaf) =>
                        leaf.type === "text" ? [leaf.text] : [],
                      )
                    : [],
                )
              : [],
          )
          .join(""),
      ),
    ),
  };
};

const run = (page: Parameters<typeof select>[0], name: string) =>
  page.evaluate(
    (command) => globalThis.__folioCanonical?.executeNamedCommand(command, []),
    name,
  );

test("column commands from the editor change the saved table and undo", async ({
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
  for (const port of [reactPort, vuePort]) {
    await page.goto(`http://localhost:${port}/?session=canonical`);
    await page.waitForSelector(".layout-page");
    await clearRefusals(page);
    await loadReady(page, source);
    // "Before" spans 0-8; the table opens at 8, its row at 9, the first cell
    // at 10 and its paragraph at 11, so "A" ends at 13.
    await select(page, 13);
    expect(await run(page, "addColumnRight")).toBe(true);
    await expectNoRefusals(page);
    expect((await snapshot(page))?.projectionMatchesCanonical).toBe(true);
    expect(await cellTexts(page)).toEqual({
      grid: 3,
      rows: [
        ["A", "", "B"],
        ["C", "", "D"],
      ],
    });
    await select(page, 13);
    expect(await run(page, "deleteColumn")).toBe(true);
    await expectNoRefusals(page);
    expect(await cellTexts(page)).toEqual({
      grid: 2,
      rows: [
        ["", "B"],
        ["", "D"],
      ],
    });
    await page.keyboard.press(`${MODIFIER}+z`);
    await expectNoRefusals(page);
    expect(await cellTexts(page)).toEqual({
      grid: 3,
      rows: [
        ["A", "", "B"],
        ["C", "", "D"],
      ],
    });
  }
});
