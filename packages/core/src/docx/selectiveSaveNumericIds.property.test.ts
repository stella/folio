import { expect, test, setDefaultTimeout } from "bun:test";
import { assertValidOoxmlNumericIds } from "@stll/docx-core";
import { panic } from "better-result";
import fc from "fast-check";
import JSZip from "jszip";

import { assertProperty, propertyTestTimeout } from "../../../../test/property-testing";
import { parseDocx } from "./parser";
import { createEmptyDocx } from "./rezip";
import { attemptSelectiveSave } from "./selectiveSave";

setDefaultTimeout(propertyTestTimeout(30_000));

const WORD_NAMESPACE = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const invalidId = fc
  .oneof(
    fc.bigInt({ min: 2_147_483_648n, max: 99_999_999_999_999_999_999n }),
    fc.bigInt({ min: -99_999_999_999_999_999_999n, max: -2_147_483_649n }),
  )
  .map(String);

test("selective saves share imported numeric identities across original and canonical baselines", async () => {
  await assertProperty(
    fc.asyncProperty(invalidId, async (id) => {
      const zip = await JSZip.loadAsync(await createEmptyDocx());
      zip.file(
        "word/document.xml",
        `<w:document xmlns:w="${WORD_NAMESPACE}" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"><w:body><w:p w14:paraId="10000001"><w:bookmarkStart w:id="${id}" w:name="anchor"/><w:r><w:t>Before</w:t></w:r><w:bookmarkEnd w:id="${id}"/><w:r><w:footnoteReference w:id="${id}"/></w:r></w:p><w:p w14:paraId="10000002"><w:bookmarkStart w:id="1" w:name="occupied"/><w:r><w:t>Untouched</w:t></w:r><w:bookmarkEnd w:id="1"/></w:p></w:body></w:document>`,
      );
      zip.file(
        "word/footnotes.xml",
        `<w:footnotes xmlns:w="${WORD_NAMESPACE}" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"><w:footnote w:id="${id}"><w:p w14:paraId="20000001"><w:r><w:t>Note</w:t></w:r></w:p></w:footnote></w:footnotes>`,
      );
      // An opaque OOXML part participates in the same package identity spaces.
      zip.file(
        "word/opaque.xml",
        `<w:opaque xmlns:w="${WORD_NAMESPACE}"><w:bookmarkStart w:id="${id}" w:name="opaque"/><w:bookmarkEnd w:id="${id}"/></w:opaque>`,
      );
      const source = await zip.generateAsync({ type: "arraybuffer" });
      const document = await parseDocx(source, { preloadFonts: false });
      const canonicalZip = await JSZip.loadAsync(document.originalBuffer);
      const canonicalOpaqueXml = await canonicalZip.file("word/opaque.xml")?.async("text");
      const canonicalNoteXml = await canonicalZip.file("word/footnotes.xml")?.async("text");
      expect(canonicalOpaqueXml).toBeDefined();
      expect(canonicalNoteXml).toBeDefined();
      expect(canonicalOpaqueXml).not.toContain(`w:id="${id}"`);

      const paragraph = document.package.document.content.at(0);
      if (paragraph?.type !== "paragraph") panic("Synthetic source has no body paragraph");
      const run = paragraph.content.find((content) => content.type === "run");
      const text = run?.content.at(0);
      if (text?.type !== "text") panic("Synthetic source has no editable text");

      for (const edited of [false, true]) {
        text.text = edited ? "After" : "Before";
        for (const baseline of [source, document.originalBuffer]) {
          // oxlint-disable-next-line no-await-in-loop -- exercise both baseline forms for each edit mode
          const saved = await attemptSelectiveSave(document, baseline, {
            changedParaIds: new Set(edited ? ["10000001"] : []),
            structuralChange: false,
            hasUntrackedChanges: false,
          });
          expect(saved).not.toBeNull();
          if (saved === null) panic("Selective save declined a normalized synthetic source");
          // oxlint-disable-next-line no-await-in-loop -- inspect the result of this baseline's save
          const savedZip = await JSZip.loadAsync(saved);
          for (const [path, file] of Object.entries(savedZip.files)) {
            if (file.dir || !path.startsWith("word/") || !path.endsWith(".xml")) continue;
            // oxlint-disable-next-line no-await-in-loop -- the emitted package oracle checks each OOXML part
            const xml = await file.async("text");
            expect(() => assertValidOoxmlNumericIds(xml, path)).not.toThrow();
          }
          // oxlint-disable-next-line no-await-in-loop -- untouched opaque XML is byte-identical to canonical import
          expect(await savedZip.file("word/opaque.xml")?.async("text")).toBe(canonicalOpaqueXml);
          // oxlint-disable-next-line no-await-in-loop -- untouched note XML is byte-identical to canonical import
          expect(await savedZip.file("word/footnotes.xml")?.async("text")).toBe(canonicalNoteXml);
          // oxlint-disable-next-line no-await-in-loop -- compare each saved model against its edit mode
          const reopened = await parseDocx(saved, { preloadFonts: false });
          expect(reopened.package.document.content).toEqual(document.package.document.content);
          expect(reopened.package.footnotes).toEqual(document.package.footnotes);
        }
      }
    }),
    { numRuns: 8 },
  );
});
