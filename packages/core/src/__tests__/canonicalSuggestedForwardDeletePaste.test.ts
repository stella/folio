/**
 * Clipboard text pasted beside or inside struck text, while suggesting, is an
 * insertion: accepting every change reads as the same keys pressed while
 * editing, and rejecting every change restores the source. The paste goes
 * through the view's own paste event, so the editor parses the text against
 * the caret as a browser paste does.
 */

import { expect, test } from "bun:test";

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

const INTRO = "Intro paragraph.";

const introDriver = async (mode: EditorMode) => {
  const source = await parseShapeDocument(
    new Uint8Array(await shapeArrayBuffer("mixed-lists")),
  );
  const driver = createCanonicalEditorHarness(source, mode);
  const target = findTextblock(driver.state.doc, INTRO);
  if (!target) {
    driver.dispose();
    throw new TypeError("The mixed-lists shape has no intro.");
  }
  return { driver, intro: target.pos + 1 };
};

type Keys = { deletes: number; pasted: string };

const pressThenPaste = async (mode: EditorMode, { deletes, pasted }: Keys) => {
  const { driver, intro } = await introDriver(mode);
  try {
    driver.history.setSelection(intro, intro);
    for (let index = 0; index < deletes; index += 1) driver.pressKey("Delete");
    expect(driver.pasteText(pasted)).toBe(true);
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
    expect(modelMarkdown(edited)).toContain(
      `${keys.pasted}${INTRO.slice(keys.deletes)}`,
    );
  });
}

test("text pasted inside struck text is inserted, not struck", async () => {
  const { driver, intro } = await introDriver("suggesting");
  try {
    const baseline = driver.snapshot();
    driver.history.setSelection(intro, intro + 2);
    driver.pressKey("Delete");
    driver.history.setSelection(intro + 1, intro + 1);
    expect(driver.pasteText("alpha")).toBe(true);
    expect(driver.refusals).toEqual([]);
    const suggested = driver.snapshot();
    expect(
      modelMarkdown(resolveCanonicalHarnessDocument(suggested, "accept").model),
    ).toContain(`alpha${INTRO.slice(2)}`);
    expect(
      modelMarkdown(resolveCanonicalHarnessDocument(suggested, "reject").model),
    ).toBe(modelMarkdown(baseline));
  } finally {
    driver.dispose();
  }
});
