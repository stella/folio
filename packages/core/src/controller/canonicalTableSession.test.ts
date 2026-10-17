import { expect, test } from "bun:test";

import { createCanonicalEditorHarness } from "../../../../test/canonicalEditorHarness";
import { parseShapeDocument } from "../__tests__/editorHarness";
import { shapeArrayBuffer } from "../__tests__/documentShapes";
import { createEmptyDocument } from "../utils/createDocument";
import { cloneDocumentWithParagraphPropertySources } from "../docx/paragraphPropertySource";
import { assertExactModel } from "../../../../test/exactModel";

const tableFixture = async (story: "main" | "header" | "footer") => {
  const parsed = await parseShapeDocument(new Uint8Array(await shapeArrayBuffer("tables")));
  if (story === "main") return parsed;
  const content = parsed.package.document.content.filter((block) => block.type === "table");
  const source = createEmptyDocument({ initialText: "Body" });
  if (story === "header")
    source.package.headers = new Map([
      ["rIdHeader", { type: "header", hdrFtrType: "default", content }],
    ]);
  else
    source.package.footers = new Map([
      ["rIdFooter", { type: "footer", hdrFtrType: "default", content }],
    ]);
  return source;
};

test.each(["main", "header", "footer"] as const)(
  "canonical sessions activate with the repository table fixture in the %s story",
  async (story) => {
    const driver = createCanonicalEditorHarness(await tableFixture(story), "editing");
    try {
      expect(driver.authority).toBe("canonical");
      expect(driver.refusals).toEqual([]);
      expect(driver.snapshot().package.document.content.length).toBeGreaterThan(0);
    } finally {
      driver.dispose();
    }
  },
);

test.each([
  { key: "Backspace", edge: "start" },
  { key: "Delete", edge: "end" },
] as const)("$key at a table-cell $edge does not cross its container", async ({ key, edge }) => {
  const driver = createCanonicalEditorHarness(await tableFixture("main"), "editing");
  try {
    let start = -1;
    let end = -1;
    driver.state.doc.descendants((node, position) => {
      if (start >= 0 || node.type.name !== "paragraph" || node.textContent !== "Cell C2") return;
      start = position + 1;
      end = start + node.content.size;
    });
    expect(start).toBeGreaterThan(0);
    const before = cloneDocumentWithParagraphPropertySources(driver.snapshot());
    driver.history.setSelection(edge === "start" ? start : end);
    driver.pressKey(key);
    assertExactModel(driver.snapshot(), before);
    expect(driver.refusals).toEqual([]);
  } finally {
    driver.dispose();
  }
});
