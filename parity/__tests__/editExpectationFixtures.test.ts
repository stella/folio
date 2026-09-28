import { describe, expect, test } from "bun:test";
import path from "node:path";

import JSZip from "jszip";

import { buildEditExpectationSeeds } from "../fixtures/build-edit-expectation-fixtures";
import { EDIT_OPERATION_SCRIPTS } from "../fixtures/edit-operation-scripts";

describe("synthetic edit expectation seeds", () => {
  test("builds deterministic DOCX seeds with the declared feature coverage", async () => {
    const first = await buildEditExpectationSeeds();
    const second = await buildEditExpectationSeeds();
    expect([...first.keys()]).toEqual([...second.keys()]);
    expect([...first.keys()].toSorted()).toEqual(
      [...new Set(EDIT_OPERATION_SCRIPTS.map(({ seed }) => seed))].toSorted(),
    );

    for (const [name, bytes] of first) {
      expect(Buffer.from(bytes).equals(Buffer.from(second.get(name) ?? []))).toBeTrue();
      const committed = new Uint8Array(
        await Bun.file(path.join(import.meta.dir, "../fixtures", name)).arrayBuffer(),
      );
      expect(Buffer.from(bytes).equals(Buffer.from(committed))).toBeTrue();
      const zip = await JSZip.loadAsync(bytes);
      const document = await zip.file("word/document.xml")?.async("text");
      if (!document) throw new TypeError(`${name} is missing word/document.xml`);
      if (name === "edit-final-paragraph-seed.docx") {
        expect(document.match(/<w:p>/g)).toHaveLength(3);
        expect(document).toContain("Target final paragraph Omega 47.");
      }
      if (name === "edit-paragraph-boundary-seed.docx") {
        expect(document.match(/<w:p>/g)).toHaveLength(2);
        expect(document).toContain("Paragraph before boundary Cedar 19.");
        expect(document).toContain("Paragraph after boundary Maple 38.");
      }
      if (name === "edit-numbering-seed.docx") {
        expect(document).toContain("w:numPr");
        expect(document).toContain('w:numId w:val="1"');
      }
      if (name === "edit-merged-table-seed.docx") {
        expect(document).toContain("w:vMerge");
        expect(document.match(/<w:tr>/g)).toHaveLength(2);
      }
      if (name === "edit-comment-range-seed.docx") {
        expect(document).toContain("w:commentRangeStart");
        expect(document).toContain("w:commentRangeEnd");
        expect(document).toContain("w:commentReference");
        expect(zip.file("word/comments.xml")).not.toBeNull();
      }
      if (name === "edit-notes-fields-sections-seed.docx") {
        expect(document).toContain("w:footnoteReference");
        expect(document).toContain("w:endnoteReference");
        expect(document).toContain("w:fldChar");
        expect(document).toContain("w:fldSimple");
        expect(document).toContain('w:type w:val="nextPage"');
        expect(zip.file("word/footnotes.xml")).not.toBeNull();
        expect(zip.file("word/endnotes.xml")).not.toBeNull();
      }
    }
  });
});
