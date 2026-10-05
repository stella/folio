import { expect, test } from "bun:test";
import { documentStories, OP_STORIES, storyBody } from "@stll/docx-core/ops";
import type { Document } from "@stll/docx-core/model";
import { cloneDocumentWithParagraphPropertySources } from "../docx/documentClone";
import { Fragment, Slice } from "prosemirror-model";
import { AllSelection, NodeSelection } from "prosemirror-state";

import {
  validateHarnessRefusalRows,
  harnessRefusalProblems,
} from "../../../../test/canonical-refusal-rows";
import { assertExactModel } from "../../../../test/exactModel";
import { createEmptyDocument } from "../utils/createDocument";
import { modelMarkdown, parseShapeDocument } from "./editorHarness";
import { documentShape, shapeArrayBuffer } from "./documentShapes";
import { PASTED_LIST, PASTED_TABLE } from "./editorCommandConformance";
import { CANONICAL_CAPABILITIES, CANONICAL_GAP } from "../types/canonicalCapabilities";
import {
  createCanonicalEditorHarness,
  createCanonicalHarnessCase,
  resolveCanonicalHarnessDocument,
  saveCanonicalHarnessDocument,
} from "../../../../test/canonicalEditorHarness";

const tableActivationSource = async (
  location: "main" | "header" | "footer" | "footnote" | "endnote",
): Promise<Document> => {
  const parsed = await parseShapeDocument(new Uint8Array(await shapeArrayBuffer("tables")));
  if (location === OP_STORIES.MAIN) return parsed;
  const content = parsed.package.document.content.filter((block) => block.type === "table");
  if (content.length === 0) throw new TypeError("The production table fixture has no table.");
  const source = createEmptyDocument({ initialText: "Valid main story" });
  const paragraph = source.package.document.content.at(0);
  if (paragraph?.type !== "paragraph")
    throw new TypeError("Table activation fixture lost its body.");
  paragraph.paraId = "FFFFFFFF";
  switch (location) {
    case "header":
      source.package.headers = new Map([
        ["rIdTableHeader", { type: "header", hdrFtrType: "default", content }],
      ]);
      break;
    case "footer":
      source.package.footers = new Map([
        ["rIdTableFooter", { type: "footer", hdrFtrType: "default", content }],
      ]);
      break;
    case "footnote":
      source.package.footnotes = [{ type: "footnote", id: 7, content }];
      break;
    case "endnote":
      source.package.endnotes = [{ type: "endnote", id: 7, content }];
      break;
  }
  return source;
};

const TABLE_ACTIVATION_CASES = (
  ["main", "header", "footer", "footnote", "endnote"] as const
).flatMap((location) => (["editing", "suggesting"] as const).map((mode) => ({ location, mode })));

test.each(TABLE_ACTIVATION_CASES)(
  "canonical $location table activation is an exact ledger precondition in $mode",
  async ({ location, mode }) => {
    const source = await tableActivationSource(location);
    const tableStories = documentStories(source).filter((story) =>
      storyBody(source, story).content.some((block) => block.type === "table"),
    );
    expect(tableStories).toHaveLength(1);
    expect(tableStories.map((story) => (story === OP_STORIES.MAIN ? story : story.kind))).toEqual([
      location,
    ]);
    const before = cloneDocumentWithParagraphPropertySources(source);
    const ambientDom = typeof document;
    const hosts = globalThis.document?.body.childElementCount ?? 0;
    const result = createCanonicalHarnessCase(source, mode);
    expect(result.type).toBe("activationRefused");
    if (result.type !== "activationRefused") {
      result.driver.dispose();
      throw new TypeError("Table activation row must become strict");
    }
    expect(result.refusal).toEqual({
      gap: CANONICAL_GAP.tableActivation,
      message: "Canonical sessions cannot activate documents containing tables.",
      expectation: "declared",
      row: "table-session-activation",
    });
    assertExactModel(source, before);
    expect(globalThis.document?.body.childElementCount ?? 0).toBe(hosts);
    expect(typeof document).toBe(ambientDom);
  },
);

test("canonical refusal rows cannot survive a retired ledger id", () => {
  const row = { id: "clipboard-table", gap: CANONICAL_GAP.dispatch, message: "refused" };
  expect(() => validateHarnessRefusalRows([row])).not.toThrow();
  const retired = Object.fromEntries(
    Object.entries(CANONICAL_CAPABILITIES).filter(([gap]) => gap !== row.gap),
  );
  expect(() => validateHarnessRefusalRows([row], retired)).toThrow("must become strict");
  const refusal = { ...row, row: row.id, expectation: "declared" } as const;
  expect(harnessRefusalProblems({ rows: [row], refusals: [refusal] })).toEqual([]);
  expect(harnessRefusalProblems({ rows: [], refusals: [refusal] })).toHaveLength(1);
  expect(
    harnessRefusalProblems({ rows: [row], refusals: [{ ...refusal, message: "different" }] }),
  ).toHaveLength(1);
});

test.each(["editing", "suggesting"] as const)(
  "canonical modified deletion carries an exact ledger refusal in %s",
  (mode) => {
    const source = createEmptyDocument({ initialText: "alpha beta" });
    const paragraph = source.package.document.content.at(0);
    if (paragraph?.type !== "paragraph") throw new TypeError("Refusal fixture lost paragraph");
    paragraph.paraId = "12345678";
    const driver = createCanonicalEditorHarness(source, mode);
    try {
      driver.history.setSelection(4, 4);
      const before = driver.snapshot();
      const state = driver.state;
      driver.pressKey("Mod-Backspace");
      expect(driver.refusals).toEqual([
        {
          gap: CANONICAL_GAP.dispatch,
          message: "Only plain character deletion is available in this session.",
          expectation: "declared",
          row: "modified-deletion",
        },
      ]);
      expect(driver.refusalRows).toHaveLength(1);
      expect(driver.state).toBe(state);
      assertExactModel(driver.snapshot(), before);
      expect(driver.history.canUndo()).toBe(false);
      expect(driver.history.canRedo()).toBe(false);
      driver.paste(PASTED_TABLE(driver.state.doc));
      expect(driver.refusals.at(-1)).toEqual({
        gap: CANONICAL_GAP.dispatch,
        message: "Clipboard tables and embedded blocks require canonical table editing.",
        expectation: "declared",
        row: "clipboard-table",
      });
      expect(driver.state).toBe(state);
      assertExactModel(driver.snapshot(), before);
      driver.history.setSelection(2, 4);
      expect(driver.cut()).toBe(true);
      expect(driver.refusals).toHaveLength(2);
      expect(driver.history.canUndo()).toBe(true);
      expect(driver.history.undo()).toBe(true);
      assertExactModel(driver.snapshot(), before);
    } finally {
      driver.dispose();
    }
  },
);

test.each(["editing", "suggesting"] as const)(
  "canonical harness owns native input, refusal and exact history in %s",
  (mode) => {
    const source = createEmptyDocument({ initialText: "alpha😀café東京" });
    const paragraph = source.package.document.content.at(0);
    if (paragraph?.type !== "paragraph")
      throw new TypeError("The native input fixture has no paragraph.");
    paragraph.paraId = "12345678";
    const driver = createCanonicalEditorHarness(source, mode);
    try {
      driver.dispatch(driver.state.tr.setSelection(NodeSelection.create(driver.state.doc, 0)));
      // Keep text replacement inside the paragraph; node replacement has its own image case.
      driver.history.setSelection(2, 4);
      const baseline = driver.snapshot();
      const selection = driver.state.selection.toJSON();
      driver.typeText("界");
      expect(driver.refusals).toEqual([]);
      const edited = driver.snapshot();
      expect(modelMarkdown(edited)).not.toBe(modelMarkdown(baseline));
      expect(driver.pressKey("Mod-z")).toBe(true);
      assertExactModel(driver.snapshot(), baseline);
      expect(driver.state.selection.toJSON()).toEqual(selection);
      expect(driver.pressKey("Mod-Shift-z")).toBe(true);
      assertExactModel(driver.snapshot(), edited);
      const beforeBypass = driver.snapshot();
      const state = driver.state;
      driver.dispatch(state.tr.insertText("bypass", 1));
      expect(driver.refusals.at(-1)?.expectation).toBe("unexpected");
      assertExactModel(driver.snapshot(), beforeBypass);
      expect(driver.state).toBe(state);
      if (mode === "suggesting") {
        expect(modelMarkdown(resolveCanonicalHarnessDocument(edited, "reject").model)).toBe(
          modelMarkdown(baseline),
        );
      }
    } finally {
      driver.dispose();
    }
  },
);

test.each(["editing", "suggesting"] as const)(
  "canonical image/node replacement routes native text and paste in %s",
  async (mode) => {
    const source = await parseShapeDocument(
      new Uint8Array(await shapeArrayBuffer(documentShape("image"))),
    );
    for (const kind of ["text", "paste", "list"] as const) {
      for (const placement of ["node", "document"] as const) {
        const driver = createCanonicalEditorHarness(source, mode);
        try {
          let position: number | undefined;
          driver.state.doc.descendants((node, offset) => {
            if (node.type.name === "image") position ??= offset;
          });
          if (position === undefined) throw new TypeError("The image/node fixture has no image.");
          driver.dispatch(
            driver.state.tr.setSelection(
              placement === "node"
                ? NodeSelection.create(driver.state.doc, position)
                : new AllSelection(driver.state.doc),
            ),
          );
          const baseline = driver.snapshot();
          const selection = driver.state.selection.toJSON();
          if (kind === "text") driver.typeText("xyz");
          else if (kind === "list") driver.paste(PASTED_LIST(driver.state.doc));
          else
            driver.paste(new Slice(Fragment.from(driver.state.schema.text("Inline paste")), 0, 0));
          expect(driver.refusals).toEqual([]);
          const edited = driver.snapshot();
          expect(driver.history.canUndo()).toBe(true);
          while (driver.history.canUndo()) expect(driver.history.undo()).toBe(true);
          assertExactModel(driver.snapshot(), baseline);
          expect(driver.state.selection.toJSON()).toEqual(selection);
          expect(driver.history.canRedo()).toBe(true);
          while (driver.history.canRedo()) expect(driver.history.redo()).toBe(true);
          assertExactModel(driver.snapshot(), edited);
          if (mode === "suggesting")
            expect(modelMarkdown(resolveCanonicalHarnessDocument(edited, "reject").model)).toBe(
              modelMarkdown(baseline),
            );
        } finally {
          driver.dispose();
        }
      }
    }
  },
);

for (const mode of ["editing", "suggesting"] as const) {
  for (const shape of ["notes", "comments", "image"] as const) {
    test(`copied ${shape} blocks use the resource preflight and saved canonical authority in ${mode}`, async () => {
      const source = await parseShapeDocument(new Uint8Array(await shapeArrayBuffer(shape)));
      const driver = createCanonicalEditorHarness(source, mode);
      try {
        const slice = new Slice(driver.state.doc.content, 0, 0);
        const last = driver.state.doc.lastChild;
        if (last === null) throw new TypeError("Copied-block fixture has no final paragraph");
        const end = driver.state.doc.content.size - 1;
        driver.history.setSelection(end, end);
        const before = driver.snapshot();
        const state = driver.state;
        driver.paste(slice);
        if (shape !== "image") {
          expect(driver.refusals).toHaveLength(1);
          expect(driver.refusals.at(0)).toMatchObject({
            expectation: "declared",
            row: "clipboard-story-parts",
          });
          assertExactModel(driver.snapshot(), before);
          expect(driver.state).toBe(state);
          expect(driver.history.canUndo()).toBe(false);
          expect(driver.history.canRedo()).toBe(false);
          return;
        }
        expect(driver.refusals).toEqual([]);
        const edited = driver.snapshot();
        const saved = await saveCanonicalHarnessDocument(edited);
        const reopened = await parseShapeDocument(saved.bytes);
        expect(modelMarkdown(reopened)).toBe(modelMarkdown(edited));
        expect(driver.history.undo()).toBe(true);
        assertExactModel(driver.snapshot(), before);
        expect(driver.history.redo()).toBe(true);
        assertExactModel(driver.snapshot(), edited);
      } finally {
        driver.dispose();
      }
    });
  }
}
