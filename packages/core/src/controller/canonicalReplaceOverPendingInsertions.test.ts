/**
 * Typing over everything while suggesting, after a suggested paste in front
 * of a comment: the hosts the replacement retires hold the paste's pending
 * insertions, which go with them, so the comment retires as in the direct
 * edit instead of the session refusing the replacement.
 */
import { describe, expect, test } from "bun:test";
import { AllSelection, TextSelection } from "prosemirror-state";

import {
  createCanonicalHarnessCase,
  resolveReopenedCanonicalHarnessDocument,
  saveCanonicalHarnessDocument,
} from "../../../../test/canonicalEditorHarness";
import { shapeArrayBuffer } from "../__tests__/documentShapes";
import { parseShapeDocument } from "../__tests__/editorHarness";
import { FolioDocxReviewer } from "../ai-edits/headless";
import type { Document } from "../types/document";

const read = async (document: Document) => {
  const { bytes } = await saveCanonicalHarnessDocument(document);
  const reviewer = await FolioDocxReviewer.fromBuffer(bytes.slice().buffer, {
    session: "legacy",
  });
  return {
    blocks: reviewer.snapshot().blocks.map(({ text }) => text),
    comments: reviewer.getComments().map(({ id }) => id),
  };
};

const typeOverEverythingAfterPaste = async (
  shape: string,
  pasted: string,
  mode: "editing" | "suggesting",
) => {
  const created = createCanonicalHarnessCase(
    await parseShapeDocument(new Uint8Array(await shapeArrayBuffer(shape))),
    mode,
  );
  if (created.type !== "ready")
    throw new TypeError("The shape opens in a canonical session.");
  const { driver } = created;
  try {
    driver.dispatch(
      driver.state.tr.setSelection(TextSelection.create(driver.state.doc, 1)),
    );
    driver.pasteText(pasted);
    driver.dispatch(
      driver.state.tr.setSelection(new AllSelection(driver.state.doc)),
    );
    driver.typeText("alpha");
    expect(driver.refusals).toEqual([]);
    return driver.snapshot();
  } finally {
    driver.dispose();
  }
};

describe("typing over everything after a suggested paste beside a comment", () => {
  for (const shape of ["comments", "invisible-markers"])
    for (const pasted of ["Title", "Title\nBody"])
      test(`${shape}, pasting ${JSON.stringify(pasted)}`, async () => {
        const source = await read(
          await parseShapeDocument(
            new Uint8Array(await shapeArrayBuffer(shape)),
          ),
        );
        const direct = await read(
          await typeOverEverythingAfterPaste(shape, pasted, "editing"),
        );
        const suggested = await typeOverEverythingAfterPaste(
          shape,
          pasted,
          "suggesting",
        );
        const accepted = await resolveReopenedCanonicalHarnessDocument(
          suggested,
          "accept",
        );
        expect(accepted.stranded).toEqual([]);
        expect(await read(accepted.model)).toEqual(direct);
        const rejected = await resolveReopenedCanonicalHarnessDocument(
          suggested,
          "reject",
        );
        expect(rejected.stranded).toEqual([]);
        expect(await read(rejected.model)).toEqual(source);
      });
});
