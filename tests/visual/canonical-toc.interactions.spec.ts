import { expect, test, type Page } from "@playwright/test";
import { createEmptyDocument } from "../../packages/core/src/utils/createDocument";
import { createDocx } from "../../packages/core/src/docx/rezip";
import { parseDocx } from "../../packages/core/src/docx/parser";
import type {} from "../parity/canonicalBridge";
import type {} from "../parity/canonicalFuzzErrors";

const MODIFIER = process.platform === "darwin" ? "Meta" : "Control";
const snapshot = async (page: Page) => {
  await page.waitForFunction(() => globalThis.__folioCanonical?.canSnapshot());
  const value = await page.evaluate(() => globalThis.__folioCanonical?.snapshot());
  expect(value?.active).toBe(true);
  expect(value?.projectionMatchesCanonical).toBe(true);
  expect(value?.projectionJSON).toEqual(value?.canonicalProjectionJSON);
  if (!value?.document) throw new TypeError("Missing canonical TOC document");
  return { ...value, document: value.document };
};
for (const { adapter, port } of [
  { adapter: "React", port: Number(process.env["FOLIO_PLAYGROUND_PORT"]) || 4200 },
  { adapter: "Vue", port: Number(process.env["FOLIO_PLAYGROUND_VUE_PORT"]) || 4201 },
]) {
  test(`canonical TOC toolbar retains bookmark targets and exact history in ${adapter}`, async ({
    page,
  }) => {
    test.setTimeout(90_000);
    const document = createEmptyDocument({ initialText: "Heading😀" });
    const heading = document.package.document.content.at(0);
    if (heading?.type !== "paragraph") throw new TypeError("Missing TOC heading");
    heading.paraId = "12345678";
    heading.formatting = { outlineLevel: { kind: "heading", level: 0 } };
    document.package.document.content.push({
      type: "paragraph",
      paraId: "23456789",
      content: [{ type: "run", content: [{ type: "text", text: "Body text" }] }],
    });
    const source = await createDocx(document);
    await page.goto(`http://localhost:${port}/?session=canonical`);
    await page.waitForSelector(".layout-page");
    expect(
      await page.evaluate(
        (bytes) => globalThis.__folioCanonical?.load(bytes),
        [...new Uint8Array(source)],
      ),
    ).toBe(true);
    expect(await page.evaluate(() => globalThis.__folioCanonical?.setMode("editing"))).toBe(true);
    expect(await page.evaluate(() => globalThis.__folioCanonical?.select(14, 17))).toBe(true);
    const insert = async () => {
      if (adapter === "React")
        await page.getByRole("button", { name: "Insert table of contents", exact: true }).click();
      else {
        await page.getByRole("button", { name: "Insert", exact: true }).click();
        await page.getByRole("button", { name: "Table of Contents", exact: true }).click();
      }
    };
    const before = await snapshot(page);
    await insert();
    const after = await snapshot(page);
    expect(after.document).not.toEqual(before.document);
    const fields = after.document.package.document.content.flatMap((paragraph) =>
      paragraph.type === "paragraph"
        ? paragraph.content.filter((item) => item.type === "complexField")
        : [],
    );
    expect(fields).toHaveLength(1);
    await page.keyboard.press(`${MODIFIER}+z`);
    const undone = await snapshot(page);
    expect(undone.document).toEqual(before.document);
    expect(undone.selectionJSON).toEqual(before.selectionJSON);
    await page.keyboard.press(`${MODIFIER}+Shift+z`);
    const redone = await snapshot(page);
    expect(redone.document).toEqual(after.document);
    expect(redone.selectionJSON).toEqual(after.selectionJSON);
    await insert();
    const regenerated = await snapshot(page);
    const currentFields = regenerated.document.package.document.content.flatMap((paragraph) =>
      paragraph.type === "paragraph"
        ? paragraph.content.filter((item) => item.type === "complexField")
        : [],
    );
    expect(currentFields).toHaveLength(2);
    expect(currentFields.map((field) => field.instruction)).toEqual([
      fields.at(0)?.instruction,
      fields.at(0)?.instruction,
    ]);
    const bytes = await page.evaluate(() => globalThis.__folioCanonical?.save());
    expect(bytes).toBeTruthy();
    if (!bytes) throw new TypeError("Missing TOC save");
    const reopened = await parseDocx(new Uint8Array(bytes).buffer);
    const reopenedFields = reopened.package.document.content.flatMap((paragraph) =>
      paragraph.type === "paragraph"
        ? paragraph.content.filter((item) => item.type === "complexField")
        : [],
    );
    expect(reopenedFields.map((field) => field.instruction)).toEqual(
      currentFields.map((field) => field.instruction),
    );
    expect(await page.evaluate((saved) => globalThis.__folioCanonical?.load(saved), bytes)).toBe(
      true,
    );
    expect(
      (await snapshot(page)).document.package.document.content.flatMap((paragraph) =>
        paragraph.type === "paragraph"
          ? paragraph.content
              .filter((item) => item.type === "complexField")
              .map((field) => field.instruction)
          : [],
      ),
    ).toEqual(currentFields.map((field) => field.instruction));
    expect(await page.evaluate(() => globalThis.__folioCanonical?.setMode("suggesting"))).toBe(
      true,
    );
    await page.evaluate(() => {
      globalThis.__folioCanonicalFuzzErrors = [];
    });
    const beforeRefusal = await snapshot(page);
    await insert();
    const refused = await snapshot(page);
    expect(refused.document).toEqual(beforeRefusal.document);
    expect(refused.selectionJSON).toEqual(beforeRefusal.selectionJSON);
    if (adapter === "React") {
      const errors = await page.evaluate(() => globalThis.__folioCanonicalFuzzErrors);
      expect(errors).toContainEqual(
        expect.objectContaining({ gap: "tracked-hyperlink-resolution" }),
      );
    }
  });
}
