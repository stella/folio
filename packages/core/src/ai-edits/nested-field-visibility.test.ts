/**
 * A field's instruction (code) region is never displayed — only the cached
 * result between its `separate` and `end` is. A nested field placed inside an
 * outer field's instruction region (`{ IF { REF x } = "a" "yes" "no" }`) is
 * OOXML the field-character definition explicitly allows, and its own result
 * must stay just as hidden as any other instruction-region text: it has not
 * been evaluated into the outer field's displayed result, so showing it is
 * showing something the page never renders.
 *
 * `field-results.test.ts` and `story-text-parity.test.ts` only cover
 * non-nested fields; this file is the nested case for every reader that walks
 * a `ComplexField` (`getContent`, `snapshotStory`/`readStory`'s unloaded model
 * walk, and `docxToMarkdown`), plus that an edit and a save/reopen keep the
 * nested structure intact.
 */

import { describe, expect, test } from "bun:test";
import JSZip from "jszip";

import { RELATIONSHIP_TYPES } from "../docx/relsParser";
import { docxToMarkdown } from "../server";
import { FolioDocxReviewer } from "./headless";

const XML_DECLARATION = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';

const run = (text: string): string => `<w:r><w:t xml:space="preserve">${text}</w:t></w:r>`;
const instrRun = (text: string): string =>
  `<w:r><w:instrText xml:space="preserve">${text}</w:instrText></w:r>`;
const fldChar = (charType: "begin" | "separate" | "end"): string =>
  `<w:r><w:fldChar w:fldCharType="${charType}"/></w:r>`;

/** A complex field's raw run sequence, its code and result each free text. */
const complexField = (code: string, result: string): string =>
  fldChar("begin") + code + fldChar("separate") + result + fldChar("end");

const createDocx = async (
  mainInner: string,
  {
    footnoteInner,
    secondParagraphInner,
  }: { footnoteInner?: string; secondParagraphInner?: string } = {},
): Promise<ArrayBuffer> => {
  const zip = new JSZip();
  const hasFootnote = footnoteInner !== undefined;
  zip.file(
    "[Content_Types].xml",
    `${XML_DECLARATION}
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
  ${hasFootnote ? '<Override PartName="/word/footnotes.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml"/>' : ""}
</Types>`,
  );
  zip.file(
    "_rels/.rels",
    `${XML_DECLARATION}
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="${RELATIONSHIP_TYPES.officeDocument}" Target="word/document.xml"/>
</Relationships>`,
  );
  if (hasFootnote) {
    zip.file(
      "word/_rels/document.xml.rels",
      `${XML_DECLARATION}
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId30" Type="${RELATIONSHIP_TYPES.footnotes}" Target="footnotes.xml"/>
</Relationships>`,
    );
    zip.file(
      "word/footnotes.xml",
      `${XML_DECLARATION}
<w:footnotes xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:footnote w:id="2"><w:p>${footnoteInner}</w:p></w:footnote>
</w:footnotes>`,
    );
  }
  const anchor = hasFootnote
    ? `<w:p><w:r><w:footnoteReference w:id="2"/></w:r><w:r><w:t>anchor</w:t></w:r></w:p>`
    : "";
  const second = secondParagraphInner !== undefined ? `<w:p>${secondParagraphInner}</w:p>` : "";
  zip.file(
    "word/document.xml",
    `${XML_DECLARATION}
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p>${mainInner}</w:p>
    ${second}
    ${anchor}
    <w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr>
  </w:body>
</w:document>`,
  );
  zip.file(
    "word/styles.xml",
    `${XML_DECLARATION}
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"/>`,
  );
  return zip.generateAsync({ type: "arraybuffer" });
};

const mainBlockText = async (buffer: ArrayBuffer): Promise<string> => {
  const reviewer = await FolioDocxReviewer.fromBuffer(buffer);
  return reviewer.getContent().at(0)?.text ?? "";
};

describe("a nested field's instruction-region result stays hidden", () => {
  test("nested field before the outer field's own separator (instruction region)", async () => {
    // { IF { PAGE } = 99 "Shown" "Hidden" } — the nested PAGE's cached "99"
    // sits before the outer IF's own separator, in its instruction region.
    const nested = complexField(instrRun(" PAGE "), run("99"));
    const outerCode = instrRun(" IF ") + nested + instrRun(' = 99 "Shown" "Hidden" ');
    const inner = run("Before ") + complexField(outerCode, run("Shown")) + run(" after.");

    const buffer = await createDocx(inner);
    expect(await mainBlockText(buffer)).toBe("Before Shown after.");

    const md = await docxToMarkdown(buffer);
    expect(md).toContain("Before Shown after.");
    expect(md).not.toContain("99");
  });

  test("nested field after the outer field's own separator (result region) still shows", async () => {
    // { IF 1=1 "Outer { PAGE } tail" "" } — the nested field's cached result
    // sits inside the outer's own displayed result, so it must show.
    const outerCode = instrRun(" IF 1=1 ");
    const nested = complexField(instrRun(" PAGE "), run("42"));
    const outerResult = run("Outer ") + nested + run(" tail");
    const inner = run("Before ") + complexField(outerCode, outerResult) + run(" after.");

    const buffer = await createDocx(inner);
    expect(await mainBlockText(buffer)).toBe("Before Outer 42 tail after.");
  });

  test("two nesting levels: a field's own result is hidden if any ancestor's code region holds it", async () => {
    // A{ IF B{ IF C{ PAGE } = 1 "BVAL" "" } = 1 "AVAL" "" } — C sits in B's
    // code region and B sits in A's code region, so both C's and B's own
    // results are hidden however their own separators read, and only A's is
    // shown.
    const fieldC = complexField(instrRun(" PAGE "), run("CVAL"));
    const fieldB = complexField(instrRun(" IF ") + fieldC + instrRun(' = 1 "BVAL" '), run("BVAL"));
    const fieldA = complexField(instrRun(" IF ") + fieldB + instrRun(' = 1 "AVAL" '), run("AVAL"));
    const inner = run("Before ") + fieldA + run(" after.");

    const buffer = await createDocx(inner);
    const text = await mainBlockText(buffer);
    expect(text).toBe("Before AVAL after.");
    expect(text).not.toContain("BVAL");
    expect(text).not.toContain("CVAL");
  });

  test("a field spanning many runs assembles the same way nested", async () => {
    // The nested field's own instruction, result and the outer's instruction
    // are each split across extra runs, mirroring how a source document
    // sometimes splits a single logical run at formatting or revision
    // boundaries.
    const nested = complexField(instrRun(" PA") + instrRun("GE "), run("9") + run("9"));
    const outerCode =
      instrRun(" I") + instrRun("F ") + nested + instrRun(' = 99 "Sh') + instrRun('own" "Hidden" ');
    const inner = run("Before ") + complexField(outerCode, run("Sh") + run("own")) + run(" after.");

    const buffer = await createDocx(inner);
    expect(await mainBlockText(buffer)).toBe("Before Shown after.");
  });

  test("a fldSimple nested in a complex field's instruction region stays hidden", async () => {
    const nestedSimple = `<w:fldSimple w:instr=" DOCPROPERTY X "><w:r><w:t>SIMPLEVAL</w:t></w:r></w:fldSimple>`;
    const outerCode = instrRun(" IF ") + nestedSimple + instrRun(' = 1 "Shown2" ');
    const inner = run("Before ") + complexField(outerCode, run("Shown2")) + run(" after.");

    const buffer = await createDocx(inner);
    const text = await mainBlockText(buffer);
    expect(text).toBe("Before Shown2 after.");
    expect(text).not.toContain("SIMPLEVAL");
  });

  test("a fldSimple nested in a complex field's result region still shows", async () => {
    const nestedSimple = `<w:fldSimple w:instr=" DOCPROPERTY X "><w:r><w:t>SIMPLEVAL</w:t></w:r></w:fldSimple>`;
    const outerResult = run("Outer ") + nestedSimple + run(" tail");
    const inner = run("Before ") + complexField(instrRun(" IF 1=1 "), outerResult) + run(" after.");

    const buffer = await createDocx(inner);
    expect(await mainBlockText(buffer)).toBe("Before Outer SIMPLEVAL tail after.");
  });

  test("a footnote's unloaded model-walk text hides nested instruction-region content too", async () => {
    const nested = complexField(instrRun(" PAGE "), run("99"));
    const outerCode = instrRun(" IF ") + nested + instrRun(' = 99 "Shown" "Hidden" ');
    const footnoteInner = run("Before ") + complexField(outerCode, run("Shown")) + run(" after.");

    const buffer = await createDocx(run("body"), { footnoteInner });
    const reviewer = await FolioDocxReviewer.fromBuffer(buffer);
    const stories = reviewer.listStories();
    const footnote = stories.find(({ handle }) => handle.type === "footnote");
    expect(footnote?.text).toBe("Before Shown after.");
  });

  test("replaceInBlock in another paragraph and a save/reopen keep the nested structure", async () => {
    const nested = complexField(instrRun(" PAGE "), run("99"));
    const outerCode = instrRun(" IF ") + nested + instrRun(' = 99 "Shown" "Hidden" ');
    const fieldParagraph = run("Before ") + complexField(outerCode, run("Shown")) + run(" after.");

    // The edit lands in the first paragraph; the field lives untouched in the
    // second, so a save patches only the edited paragraph and the field
    // paragraph's original bytes — nested field characters included — are
    // never rebuilt through ProseMirror's flatter `field` node round-trip.
    const buffer = await createDocx(run("Editable text."), {
      secondParagraphInner: fieldParagraph,
    });
    const reviewer = await FolioDocxReviewer.fromBuffer(buffer);
    const blocks = reviewer.getContent();
    expect(blocks.map((b) => b.text)).toEqual(["Editable text.", "Before Shown after."]);
    const editableBlock = blocks[0];
    if (!editableBlock) {
      throw new Error("expected a block");
    }

    const result = reviewer.applyOperations(
      [
        {
          id: "edit",
          type: "replaceInBlock",
          blockId: editableBlock.id,
          find: "Editable",
          replace: "Edited",
        },
      ],
      { mode: "direct" },
    );
    expect(result.applied).toHaveLength(1);
    const afterEdit = reviewer.getContent().map((b) => b.text);
    expect(afterEdit).toEqual(["Edited text.", "Before Shown after."]);
    expect(afterEdit.join("")).not.toContain("99");

    const saved = await reviewer.toBuffer();
    const reopened = await FolioDocxReviewer.fromBuffer(saved);
    const reopenedTexts = reopened.getContent().map((b) => b.text);
    expect(reopenedTexts).toEqual(["Edited text.", "Before Shown after."]);
    expect(reopenedTexts.join("")).not.toContain("99");

    // The untouched field paragraph's own structure — both begin/separate/end
    // triples — survives the save byte-identifiably, not merely by
    // coincidence of text.
    const zip = await JSZip.loadAsync(saved);
    const xml = (await zip.file("word/document.xml")?.async("text")) ?? "";
    expect(xml.match(/w:fldCharType="begin"/g)).toHaveLength(2);
    expect(xml.match(/w:fldCharType="separate"/g)).toHaveLength(2);
    expect(xml.match(/w:fldCharType="end"/g)).toHaveLength(2);
    expect(xml).toContain("PAGE");
  });
});
