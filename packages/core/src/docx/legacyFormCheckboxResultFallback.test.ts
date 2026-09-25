/**
 * A legacy `FORMCHECKBOX` field with no cached result run must not gain one
 * on save.
 *
 * `w:ffData/w:checkBox` states a checkbox's default and current state, but a
 * document is free to leave the field's result region (between the
 * `separate` and `end` field characters) empty — a compliant reader
 * recomputes the glyph on open. The paragraph parser synthesizes that glyph
 * into
 * `ComplexField.fieldResult` so the editor has something to paint, and
 * flags it `fieldResultIsFallback` because the source never authored it.
 * Serializing that run anyway on an unedited save turned an absent result
 * into a literal "☐"/"☒" text run the next open would treat as authored
 * content.
 *
 * A `w:sdt`/`w14:checkbox` content control never synthesizes at parse time —
 * its `w:sdtContent` is ordinary block content, preserved as authored. The
 * second describe block guards that this stays true.
 */

import { describe, expect, test } from "bun:test";
import JSZip from "jszip";

import { fromProseDoc } from "../prosemirror/conversion/fromProseDoc";
import { toProseDoc } from "../prosemirror/conversion/toProseDoc";
import { parseDocx } from "./parser";
import { createEmptyDocx, repackDocx } from "./rezip";

const XML_DECLARATION = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';
const W_NAMESPACE = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const W14_NAMESPACE = "http://schemas.microsoft.com/office/word/2010/wordml";

const buildDocx = async (body: string): Promise<ArrayBuffer> => {
  const zip = await JSZip.loadAsync(await createEmptyDocx());
  zip.file(
    "word/document.xml",
    `${XML_DECLARATION}<w:document xmlns:w="${W_NAMESPACE}" xmlns:w14="${W14_NAMESPACE}"><w:body>${body}<w:sectPr/></w:body></w:document>`,
  );
  return zip.generateAsync({ type: "arraybuffer" });
};

const documentXmlOf = async (saved: ArrayBuffer): Promise<string> =>
  (await (await JSZip.loadAsync(saved)).file("word/document.xml")?.async("text")) ?? "";

/** The part a plain save writes, and the part the editor round trip writes. */
const savedAndProjected = async (body: string): Promise<{ saved: string; projected: string }> => {
  const parsed = await parseDocx(await buildDocx(body), { preloadFonts: false });
  const saved = await documentXmlOf(await repackDocx(parsed, { updateModifiedDate: false }));
  const roundTripped = fromProseDoc(toProseDoc(parsed), parsed);
  const projected = await documentXmlOf(
    await repackDocx(roundTripped, { updateModifiedDate: false }),
  );
  return { saved, projected };
};

describe("a resultless FORMCHECKBOX keeps no result on save", () => {
  const uncheckedNoRpr = `<w:p w14:paraId="10000001"><w:r><w:fldChar w:fldCharType="begin"><w:ffData><w:name w:val="Check1"/><w:enabled/><w:calcOnExit w:val="0"/><w:checkBox><w:sizeAuto/><w:default w:val="0"/></w:checkBox></w:ffData></w:fldChar></w:r><w:r><w:instrText xml:space="preserve"> FORMCHECKBOX </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r></w:p>`;

  test("no glyph run appears between separate and end", async () => {
    const { saved, projected } = await savedAndProjected(uncheckedNoRpr);
    for (const xml of [saved, projected]) {
      expect(xml).not.toContain("☐");
      expect(xml).not.toContain("☒");
      expect(xml).not.toContain("<w:t");
      expect(xml).toContain(
        '<w:fldChar w:fldCharType="separate"/></w:r><w:r><w:fldChar w:fldCharType="end"/>',
      );
    }
  });

  test("an explicit checkbox size does not leak a w:rPr onto the structural runs", async () => {
    const explicitSize = `<w:p w14:paraId="10000002"><w:r><w:fldChar w:fldCharType="begin"><w:ffData><w:name w:val="Check2"/><w:enabled/><w:calcOnExit w:val="0"/><w:checkBox><w:size w:val="24"/><w:default w:val="0"/></w:checkBox></w:ffData></w:fldChar></w:r><w:r><w:instrText xml:space="preserve"> FORMCHECKBOX </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r></w:p>`;
    const { saved, projected } = await savedAndProjected(explicitSize);
    for (const xml of [saved, projected]) {
      expect(xml).not.toContain("w:sz");
      expect(xml).toContain(
        '<w:fldChar w:fldCharType="separate"/></w:r><w:r><w:fldChar w:fldCharType="end"/>',
      );
    }
  });

  test("a checked default field still leaks nothing", async () => {
    const checkedDefault = `<w:p w14:paraId="10000003"><w:r><w:fldChar w:fldCharType="begin"><w:ffData><w:name w:val="Check3"/><w:enabled/><w:calcOnExit w:val="0"/><w:checkBox><w:sizeAuto/><w:default w:val="1"/></w:checkBox></w:ffData></w:fldChar></w:r><w:r><w:instrText xml:space="preserve"> FORMCHECKBOX </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r></w:p>`;
    const { saved, projected } = await savedAndProjected(checkedDefault);
    for (const xml of [saved, projected]) {
      expect(xml).not.toContain("☐");
      expect(xml).not.toContain("☒");
    }
  });

  test("a field with its own cached result run keeps it verbatim (not treated as a fallback)", async () => {
    const cachedChecked = `<w:p w14:paraId="10000004"><w:r><w:fldChar w:fldCharType="begin"><w:ffData><w:name w:val="Check4"/><w:enabled/><w:calcOnExit w:val="0"/><w:checkBox><w:sizeAuto/><w:default w:val="0"/><w:checked/></w:checkBox></w:ffData></w:fldChar></w:r><w:r><w:instrText xml:space="preserve"> FORMCHECKBOX </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:rPr><w:rFonts w:ascii="MS Gothic"/></w:rPr><w:t>☒</w:t></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r></w:p>`;
    const { saved, projected } = await savedAndProjected(cachedChecked);
    for (const xml of [saved, projected]) {
      expect(xml).toContain("<w:t>☒</w:t>");
    }
  });
});

describe("a w14:checkbox content control keeps its authored sdtContent on save", () => {
  const checkboxSdt = `<w:p w14:paraId="20000001"><w:sdt><w:sdtPr><w:id w:val="1"/><w14:checkbox><w14:checked w14:val="0"/><w14:checkedState w14:val="2612" w14:font="MS Gothic"/><w14:uncheckedState w14:val="2610" w14:font="MS Gothic"/></w14:checkbox></w:sdtPr><w:sdtContent><w:r><w:rPr><w:rFonts w:ascii="MS Gothic" w:hAnsi="MS Gothic" w:hint="eastAsia"/></w:rPr><w:t>☐</w:t></w:r></w:sdtContent></w:sdt></w:p>`;

  test("the authored glyph run round-trips unchanged", async () => {
    const { saved, projected } = await savedAndProjected(checkboxSdt);
    for (const xml of [saved, projected]) {
      expect(xml).toContain("<w:t>☐</w:t>");
    }
  });
});
