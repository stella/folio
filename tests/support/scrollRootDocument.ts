import JSZip from "jszip";
import {
  SCROLL_CONTROL_TAG,
  SCROLL_REVISION_ID,
  SCROLL_TARGET_PARA_ID,
  SCROLL_TARGET_TEXT,
} from "../../packages/playground/src/scrollParityBridge";

/** Explicit page boundaries and OOXML paragraph identifiers make navigation targets stable. */
export const buildScrollRootDocument = () => {
  const zip = new JSZip();
  const options = { date: new Date("2026-01-01T00:00:00Z") };
  zip.file(
    "[Content_Types].xml",
    `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>`,
    options,
  );
  zip.file(
    "_rels/.rels",
    `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>`,
    options,
  );
  zip.file(
    "word/document.xml",
    `<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"><w:body>
<w:p w14:paraId="13300100"><w:r><w:t>First page</w:t></w:r></w:p>
<w:p w14:paraId="13300200"><w:pPr><w:pageBreakBefore/></w:pPr><w:r><w:t>Second page</w:t></w:r></w:p>
<w:p w14:paraId="${SCROLL_TARGET_PARA_ID}"><w:pPr><w:pageBreakBefore/></w:pPr><w:r><w:t>${SCROLL_TARGET_TEXT}</w:t></w:r><w:ins w:id="${SCROLL_REVISION_ID}" w:author="Reviewer" w:date="2026-01-01T00:00:00Z"><w:r><w:t> revised</w:t></w:r></w:ins></w:p>
<w:sdt><w:sdtPr><w:id w:val="1330"/><w:tag w:val="${SCROLL_CONTROL_TAG}"/><w:text/></w:sdtPr><w:sdtContent><w:p w14:paraId="13300301"><w:r><w:t>Content control destination</w:t></w:r></w:p></w:sdtContent></w:sdt>
<w:p w14:paraId="13300400"><w:pPr><w:pageBreakBefore/></w:pPr><w:r><w:t>Fourth page</w:t></w:r></w:p>
<w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/></w:sectPr>
</w:body></w:document>`,
    options,
  );
  return zip.generateAsync({ type: "uint8array", compression: "DEFLATE" });
};
