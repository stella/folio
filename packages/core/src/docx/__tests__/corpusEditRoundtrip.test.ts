/**
 * Every corpus fixture must survive a real document mutation, save, and reopen.
 * The inserted paragraph proves the model changed; byte comparisons prove that
 * package parts unrelated to the body edit were preserved.
 */

import { describe, expect, test } from "bun:test";
import JSZip from "jszip";
import { readFileSync, readdirSync } from "node:fs";
import path from "node:path";
import { EditorState } from "prosemirror-state";

import { fromProseDoc } from "../../prosemirror/conversion/fromProseDoc";
import { toProseDoc } from "../../prosemirror/conversion/toProseDoc";
import { parseDocx } from "../parser";
import { repackDocx } from "../rezip";

const FIXTURES_DIR = path.join(import.meta.dir, "__fixtures__", "corpus");
const EDIT_MARKER_PREFIX = "Folio corpus edit: ";
const BODY_PART = "word/document.xml";

const FIXTURE_FILES = readdirSync(FIXTURES_DIR)
  .filter((name) => name.endsWith(".docx"))
  .sort();

const readFixture = (filename: string): ArrayBuffer => {
  const bytes = readFileSync(path.join(FIXTURES_DIR, filename));
  return new Uint8Array(bytes).buffer;
};

const packageFileNames = (zip: JSZip): string[] =>
  Object.values(zip.files)
    .filter((file) => !file.dir)
    .map((file) => file.name)
    .sort();

type ExpectUnrelatedPartsPreservedOptions = {
  originalBytes: ArrayBuffer;
  savedBytes: ArrayBuffer;
};

const expectUnrelatedPartsPreserved = async ({
  originalBytes,
  savedBytes,
}: ExpectUnrelatedPartsPreservedOptions): Promise<void> => {
  const [originalZip, savedZip] = await Promise.all([
    JSZip.loadAsync(originalBytes),
    JSZip.loadAsync(savedBytes),
  ]);
  const originalNames = packageFileNames(originalZip);
  expect(packageFileNames(savedZip)).toEqual(originalNames);

  for (const filename of originalNames) {
    if (filename === BODY_PART) {
      continue;
    }

    const originalFile = originalZip.file(filename);
    const savedFile = savedZip.file(filename);
    if (!originalFile || !savedFile) {
      throw new Error(`Expected package part is missing: ${filename}`);
    }

    const [originalPart, savedPart] = await Promise.all([
      originalFile.async("nodebuffer"),
      savedFile.async("nodebuffer"),
    ]);
    expect(savedPart.equals(originalPart)).toBe(true);
  }
};

describe("corpus edit/save/reopen", () => {
  test("corpus fixtures are present", () => {
    expect(FIXTURE_FILES.length).toBeGreaterThan(0);
  });

  test.each(FIXTURE_FILES)("preserves package content after a body edit (%s)", async (filename) => {
    const originalBytes = readFixture(filename);
    const parsed = await parseDocx(originalBytes);
    const originalPm = toProseDoc(parsed);
    const marker = `${EDIT_MARKER_PREFIX}${filename}`;
    const markerParagraph = originalPm.type.schema.nodes.paragraph.create(
      null,
      originalPm.type.schema.text(marker),
    );
    const state = EditorState.create({ doc: originalPm });
    const editedPm = state.apply(state.tr.insert(state.doc.content.size, markerParagraph)).doc;
    const savedBytes = await repackDocx(fromProseDoc(editedPm, parsed), {
      updateModifiedDate: false,
    });
    const reopenedPm = toProseDoc(await parseDocx(savedBytes));

    expect(reopenedPm.textContent).toBe(`${originalPm.textContent}${marker}`);
    expect(reopenedPm.childCount).toBe(originalPm.childCount + 1);
    await expectUnrelatedPartsPreserved({ originalBytes, savedBytes });
  });
});
