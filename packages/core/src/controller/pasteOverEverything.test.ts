/**
 * Pasting over the whole document, then typing: accepted, the suggested edit
 * leaves the blocks the direct edit does, in every document shape. Markers
 * and pending deletions left beside the caret keep no paragraph of their own.
 */
import { expect, test } from "bun:test";
import { type Node as PMNode, Slice } from "prosemirror-model";
import { AllSelection } from "prosemirror-state";

import {
  createCanonicalEditorHarness,
  resolveCanonicalHarnessDocument,
} from "../../../../test/canonicalEditorHarness";
import { DOCUMENT_SHAPES, shapeArrayBuffer } from "../__tests__/documentShapes";
import {
  PASTED_PARAGRAPHS,
  PASTED_TABLE,
} from "../__tests__/editorCommandConformance";
import { parseShapeDocument } from "../__tests__/editorHarness";
import type { Document } from "../types/document";

const CLIPBOARDS = {
  table: PASTED_TABLE,
  "open paragraphs": PASTED_PARAGRAPHS,
  "closed paragraphs": (doc: PMNode) =>
    new Slice(PASTED_PARAGRAPHS(doc).content, 0, 0),
} as const satisfies Record<string, (doc: PMNode) => Slice>;

/** Each body block's kind; a paragraph's as the kinds of what it holds. */
const blockKinds = (document: Document) =>
  document.package.document.content.map((block) =>
    block.type === "paragraph"
      ? block.content.map((item) => item.type)
      : block.type,
  );

for (const shape of DOCUMENT_SHAPES)
  for (const [name, clipboard] of Object.entries(CLIPBOARDS))
    test(`${shape.id}: ${name} pasted over everything reviews as the direct paste`, async () => {
      const source = await parseShapeDocument(
        new Uint8Array(await shapeArrayBuffer(shape)),
      );
      const edit = (mode: "editing" | "suggesting") => {
        const driver = createCanonicalEditorHarness(source, mode);
        try {
          driver.dispatch(
            driver.state.tr.setSelection(new AllSelection(driver.state.doc)),
          );
          driver.paste(clipboard(driver.state.doc));
          driver.typeText("alpha");
          return { model: driver.snapshot(), refusals: driver.refusals.length };
        } finally {
          driver.dispose();
        }
      };
      const direct = edit("editing");
      const suggested = edit("suggesting");
      expect(suggested.refusals).toBe(direct.refusals);
      if (direct.refusals > 0) return;
      // A shape may hold changes already pending; both sides accept them.
      const accepted = (model: Document) =>
        blockKinds(resolveCanonicalHarnessDocument(model, "accept").model);
      expect(accepted(suggested.model)).toEqual(accepted(direct.model));
    });
