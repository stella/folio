import { describe, expect, test } from "bun:test";
import { NodeSelection, TextSelection } from "prosemirror-state";

import type { Document, Paragraph } from "@stll/docx-core/model";
import { paragraphLogicalText } from "@stll/docx-core/ops";

import { createCanonicalEditorHarness } from "../../../../test/canonicalEditorHarness";
import { createEmptyDocument } from "../utils/createDocument";
import type { EditorMode } from "./editorHarness";

const MODES: readonly EditorMode[] = ["editing", "suggesting"];

/** A contents entry: a link holding the heading text and its page field whole. */
const contentsEntry = (): Document => {
  const document = createEmptyDocument();
  const paragraph: Paragraph = {
    type: "paragraph",
    paraId: "12345678",
    content: [
      {
        type: "hyperlink",
        anchor: "_Toc1",
        children: [
          { type: "run", content: [{ type: "text", text: "Overview " }] },
          {
            type: "complexField",
            instruction: " PAGEREF _Toc1 \\h ",
            fieldType: "PAGEREF",
            fieldCode: [],
            // A result of differently formatted runs keeps its own content.
            fieldResult: [
              {
                type: "run",
                formatting: { bold: true },
                content: [{ type: "text", text: "1" }],
              },
              { type: "run", content: [{ type: "text", text: "2" }] },
            ],
          },
        ],
      },
      { type: "run", content: [{ type: "text", text: " tail" }] },
    ],
  };
  document.package.document.content = [paragraph];
  return document;
};

const hasLink = (document: Document) =>
  JSON.stringify(document.package.document.content).includes(
    '"type":"hyperlink"',
  );

describe("link removal over a field the link holds", () => {
  test.each(
    MODES.flatMap((mode) =>
      (["caret in the entry text", "field selected"] as const).map(
        (selection) => [`${selection}, ${mode}`, selection, mode] as const,
      ),
    ),
  )("%s takes the link off the field whole", (_name, selection, mode) => {
    const driver = createCanonicalEditorHarness(contentsEntry(), mode);
    try {
      let field = -1;
      driver.state.doc.descendants((node, position) => {
        if (node.type.name === "structuredField") field = position;
      });
      expect(field).toBeGreaterThan(0);
      driver.dispatch(
        driver.state.tr.setSelection(
          selection === "field selected"
            ? NodeSelection.create(driver.state.doc, field)
            : TextSelection.create(driver.state.doc, 3),
        ),
      );
      const before = driver.snapshot();
      expect(
        driver.execute(
          driver.commandManager.requireCommand("removeHyperlink")(),
        ),
      ).toBe(true);
      const after = driver.snapshot();
      const paragraph = after.package.document.content.at(0);
      if (paragraph?.type !== "paragraph")
        throw new TypeError("Expected the entry paragraph.");
      if (mode === "editing") {
        // The field holds no link; the caret's whole link leaves the text too.
        expect(JSON.stringify(paragraph.content)).toContain('"complexField"');
        expect(hasLink(after)).toBe(selection === "field selected");
        // The field stands as one unit in the paragraph text.
        expect(paragraphLogicalText(paragraph)).toBe("Overview \uFFFC tail");
      } else {
        // The linked content is struck and a plain copy follows it.
        expect(JSON.stringify(paragraph.content)).toContain('"deletion"');
        expect(JSON.stringify(paragraph.content)).toContain('"insertion"');
        // Rejecting brings the link back as it was.
        expect(driver.resolve("reject")).toBe(true);
        expect(driver.snapshot().package.document.content).toStrictEqual(
          before.package.document.content,
        );
        expect(driver.history.undo()).toBe(true);
        // Accepting leaves the field, unlinked, where it was.
        expect(driver.resolve("accept")).toBe(true);
        const accepted = driver.snapshot().package.document.content.at(0);
        if (accepted?.type !== "paragraph")
          throw new TypeError("Expected the entry paragraph.");
        expect(paragraphLogicalText(accepted)).toBe("Overview ￼ tail");
        expect(JSON.stringify(accepted.content)).toContain('"complexField"');
        expect(driver.history.undo()).toBe(true);
      }
      expect(driver.history.undo()).toBe(true);
      expect(driver.snapshot()).toStrictEqual(before);
    } finally {
      driver.dispose();
    }
  });
});
