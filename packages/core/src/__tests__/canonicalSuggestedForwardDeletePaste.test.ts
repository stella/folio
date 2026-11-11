/**
 * Text pasted at the caret after forward deletions, while suggesting, is an
 * insertion beside the struck text: accepting every change reads as the same
 * keys pressed while editing.
 */

import { expect, test } from "bun:test";
import { Fragment, Slice } from "prosemirror-model";

import {
  createCanonicalEditorHarness,
  resolveCanonicalHarnessDocument,
} from "../../../../test/canonicalEditorHarness";
import { shapeArrayBuffer } from "./documentShapes";
import {
  type EditorMode,
  findTextblock,
  modelMarkdown,
  parseShapeDocument,
} from "./editorHarness";

type Keys = { deletes: number; pasted: string };

const pressThenPaste = async (mode: EditorMode, { deletes, pasted }: Keys) => {
  const source = await parseShapeDocument(
    new Uint8Array(await shapeArrayBuffer("mixed-lists")),
  );
  const driver = createCanonicalEditorHarness(source, mode);
  try {
    const target = findTextblock(driver.state.doc, "Intro paragraph.");
    if (!target) throw new TypeError("The mixed-lists shape has no intro.");
    driver.history.setSelection(target.pos + 1, target.pos + 1);
    for (let index = 0; index < deletes; index += 1) driver.pressKey("Delete");
    driver.paste(
      new Slice(Fragment.from(driver.state.schema.text(pasted)), 0, 0),
    );
    expect(driver.refusals).toEqual([]);
    return driver.snapshot();
  } finally {
    driver.dispose();
  }
};

for (const keys of [
  { deletes: 1, pasted: "alpha" },
  { deletes: 2, pasted: "alpha" },
  { deletes: 3, pasted: "a" },
]) {
  test(`${keys.deletes} forward deletions then a paste, accepted, equal the direct keys`, async () => {
    const edited = await pressThenPaste("editing", keys);
    const suggested = await pressThenPaste("suggesting", keys);
    expect(
      modelMarkdown(resolveCanonicalHarnessDocument(suggested, "accept").model),
    ).toBe(modelMarkdown(edited));
    expect(modelMarkdown(edited)).toContain(`${keys.pasted}`);
  });
}
