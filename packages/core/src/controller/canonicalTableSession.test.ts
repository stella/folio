import { expect, spyOn, test } from "bun:test";

import { createCanonicalEditorHarness } from "../../../../test/canonicalEditorHarness";
import { assertExactModel } from "../../../../test/exactModel";
import { shapeArrayBuffer } from "../__tests__/documentShapes";
import { parseShapeDocument } from "../__tests__/editorHarness";
import { cloneDocumentWithParagraphPropertySources } from "../docx/paragraphPropertySource";
import * as conversion from "../prosemirror/conversion/toProseDoc";
import { schema } from "../prosemirror/schema";
import { createEmptyDocument } from "../utils/createDocument";
import { createCanonicalSession } from "./canonicalSession";

const tableFixture = async (story: "main" | "header" | "footer") => {
  const parsed = await parseShapeDocument(
    new Uint8Array(await shapeArrayBuffer("tables")),
  );
  if (story === "main") return parsed;
  const content = parsed.package.document.content.filter(
    (block) => block.type === "table",
  );
  const source = createEmptyDocument({ initialText: "Body" });
  const body = source.package.document.content.at(0);
  if (body?.type === "paragraph") body.paraId = "10000000";
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
    const driver = createCanonicalEditorHarness(
      await tableFixture(story),
      "editing",
    );
    try {
      expect(driver.authority).toBe("canonical");
      expect(driver.refusals).toEqual([]);
      expect(driver.snapshot().package.document.content.length).toBeGreaterThan(
        0,
      );
    } finally {
      driver.dispose();
    }
  },
);

test("canonical sessions project a vertical-merge continuation without its source paragraph", async () => {
  const document = await tableFixture("main");
  const continuation = document.package.document.content
    .filter((block) => block.type === "table")
    .flatMap((table) => table.rows)
    .flatMap((row) => row.cells)
    .find((cell) => cell.formatting?.vMerge === "continue");
  expect(continuation).toBeDefined();

  const driver = createCanonicalEditorHarness(document, "editing");
  try {
    expect(driver.authority).toBe("canonical");
    expect(driver.refusals).toEqual([]);
    const projectedIds = new Set<string>();
    driver.state.doc.descendants((node) => {
      const paraId = node.attrs["paraId"];
      if (node.type.name === "paragraph" && typeof paraId === "string")
        projectedIds.add(paraId);
    });
    const continuationParagraphIds = continuation?.content.flatMap((block) =>
      block.type === "paragraph" && block.paraId !== undefined
        ? [block.paraId]
        : [],
    );
    expect(continuationParagraphIds?.length).toBeGreaterThan(0);
    expect(
      continuationParagraphIds?.every((paraId) => !projectedIds.has(paraId)),
    ).toBe(true);
  } finally {
    driver.dispose();
  }
});

test("canonical sessions refuse a projection that drops a non-continuation paragraph", () => {
  const document = createEmptyDocument({ initialText: "Projected" });
  const firstSource = document.package.document.content.at(0);
  if (firstSource?.type === "paragraph") firstSource.paraId = "10000000";
  document.package.document.content.push({
    type: "paragraph",
    paraId: "20000000",
    content: [],
  });
  const projected = conversion.toProseDoc(document);
  const first = projected.firstChild;
  expect(first).toBeDefined();
  const injected = spyOn(conversion, "toProseDoc").mockReturnValue(
    schema.node("doc", null, first),
  );
  try {
    const result = createCanonicalSession(document);
    expect(result.isErr()).toBe(true);
    if (result.isErr())
      expect(result.error.message).toBe(
        "The projection changed paragraph structure.",
      );
  } finally {
    injected.mockRestore();
  }
});

test.each([
  { key: "Backspace", edge: "start" },
  { key: "Delete", edge: "end" },
] as const)(
  "$key at a table-cell $edge does not cross its container",
  async ({ key, edge }) => {
    const driver = createCanonicalEditorHarness(
      await tableFixture("main"),
      "editing",
    );
    try {
      let start = -1;
      let end = -1;
      driver.state.doc.descendants((node, position) => {
        if (
          start >= 0 ||
          node.type.name !== "paragraph" ||
          node.textContent !== "Cell C2"
        )
          return;
        start = position + 1;
        end = start + node.content.size;
      });
      expect(start).toBeGreaterThan(0);
      const before = cloneDocumentWithParagraphPropertySources(
        driver.snapshot(),
      );
      driver.history.setSelection(edge === "start" ? start : end);
      driver.pressKey(key);
      assertExactModel(driver.snapshot(), before);
      expect(driver.refusals).toEqual([]);
    } finally {
      driver.dispose();
    }
  },
);
