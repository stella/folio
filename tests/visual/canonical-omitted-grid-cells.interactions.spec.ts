import { expect } from "@playwright/test";

import { parseDocx } from "../../packages/core/src/docx/parser";
import { createDocx } from "../../packages/core/src/docx/rezip";
import { createEmptyDocument } from "../../packages/core/src/utils/createDocument";
import {
  clearRefusals,
  expectNoRefusals,
  loadReady,
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

test("a table whose rows stop short of its grid opens and its cells take typing", async ({
  page,
}) => {
  test.setTimeout(60_000);
  const document = createEmptyDocument({ initialText: "Before" });
  document.package.document.content.push({
    type: "table",
    rows: [
      { type: "tableRow", formatting: { gridAfter: 1 }, cells: [cell("A")] },
      { type: "tableRow", cells: [cell("E"), cell("F")] },
    ],
  });
  const source = await canonicalLoadFixture(await createDocx(document));
  for (const port of [reactPort, vuePort]) {
    await page.goto(`http://localhost:${port}/?session=canonical`);
    await page.waitForSelector(".layout-page");
    await clearRefusals(page);
    await loadReady(page, source);
    await expectNoRefusals(page);
    // "Before" spans 0-8; the table opens at 8, its row at 9, the first cell at
    // 10 and its paragraph at 11, so "A" ends at 13.
    const end = 13;
    await select(page, end);
    await page.keyboard.type("x");
    expect((await snapshot(page))?.text).toContain("Ax");
    await expectNoRefusals(page);
    const saved = await page.evaluate(async () => {
      const buffer = await globalThis.__folioCanonical?.save();
      return buffer ? [...new Uint8Array(buffer)] : null;
    });
    if (!saved) throw new TypeError("The canonical document did not save.");
    const reopened = await parseDocx(new Uint8Array(saved), {
      preloadFonts: false,
      detectVariables: false,
    });
    const table = reopened.package.document.content.find(
      (block) => block.type === "table",
    );
    if (table?.type !== "table") throw new TypeError("The table is absent.");
    expect(
      table.rows.map(({ formatting }) => formatting?.gridAfter ?? 0),
    ).toEqual([1, 0]);
    expect(
      table.rows[0]?.cells[0]?.content
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
    ).toBe("Ax");
  }
});
