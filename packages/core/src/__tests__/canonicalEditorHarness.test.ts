import { expect, test } from "bun:test";
import { Fragment, Slice } from "prosemirror-model";
import { NodeSelection } from "prosemirror-state";

import { assertExactModel } from "../../../../test/exactModel";
import { createEmptyDocument } from "../utils/createDocument";
import { modelMarkdown, parseShapeDocument } from "./editorHarness";
import { documentShape, shapeArrayBuffer } from "./documentShapes";
import { PASTED_LIST } from "./editorCommandConformance";
import {
  createCanonicalEditorHarness,
  resolveCanonicalHarnessDocument,
} from "../../../../test/canonicalEditorHarness";

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
      const driver = createCanonicalEditorHarness(source, mode);
      try {
        let position: number | undefined;
        driver.state.doc.descendants((node, offset) => {
          if (node.type.name === "image") position ??= offset;
        });
        if (position === undefined) throw new TypeError("The image/node fixture has no image.");
        driver.dispatch(
          driver.state.tr.setSelection(NodeSelection.create(driver.state.doc, position)),
        );
        const baseline = driver.snapshot();
        const selection = driver.state.selection.toJSON();
        if (kind === "text") driver.typeText("xyz");
        else if (kind === "list") driver.paste(PASTED_LIST(driver.state.doc));
        else driver.paste(new Slice(Fragment.from(driver.state.schema.text("Inline paste")), 0, 0));
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
  },
);
