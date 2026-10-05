import { expect, test } from "bun:test";
import {
  buildPatchedDocumentXml,
  collectParaIds,
  findParagraphOffsets,
  scanParagraphs,
} from "./selectiveXmlPatch";

const part = (content: string) =>
  `<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"><w:body>${content}</w:body></w:document>`;

for (const quote of ["'", '"']) {
  for (const assignment of ["=", " = ", "\t=\n"]) {
    test(`paragraph identity survives XML attribute quote ${quote} and assignment ${JSON.stringify(assignment)}`, () => {
      const sourceParagraph = `<w:p w14:paraId${assignment}${quote}11111111${quote}><w:r><w:t>before</w:t></w:r></w:p>`;
      const tail = '<w:p w14:paraId="22222222"><w:r><w:t>unchanged</w:t></w:r></w:p>';
      const source = part(sourceParagraph + "\n<!-- gap -->\n" + tail);
      const edited = '<w:p w14:paraId="11111111"><w:r><w:t>after</w:t></w:r></w:p>';
      const generated = part(edited + tail);
      expect(scanParagraphs(source).map(({ paraId }) => paraId)).toEqual(["11111111", "22222222"]);
      expect(collectParaIds(source)).toEqual(
        new Map([
          ["11111111", 1],
          ["22222222", 1],
        ]),
      );
      const range = findParagraphOffsets(source, "11111111");
      expect(range && source.slice(range.start, range.end)).toBe(sourceParagraph);
      expect(buildPatchedDocumentXml(source, generated, new Set(["11111111"]))).toBe(
        part(edited + "\n<!-- gap -->\n" + tail),
      );
      const anonymousSource = part(
        `<w:p w14:textId${assignment}${quote}33333333${quote}><w:r><w:t>before</w:t></w:r></w:p>`,
      );
      const anonymousGenerated = part(
        '<w:p w14:paraId="11111111" w14:textId="33333333"><w:r><w:t>after</w:t></w:r></w:p>',
      );
      expect(
        buildPatchedDocumentXml(anonymousSource, anonymousGenerated, new Set(["11111111"])),
      ).toBe(part('<w:p w14:textId="33333333"><w:r><w:t>after</w:t></w:r></w:p>'));
    });
  }
}
