import JSZip from "jszip";
import { panic } from "better-result";

import { createEmptyDocx } from "../../docx/rezip";

const NUMBERING_CONTENT_TYPE =
  "application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml";
const NUMBERING_RELATIONSHIP =
  "http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering";

/** A minimal DOCX with outline signals at paragraph, style, and numbering tiers. */
export const createHeadingCollectorOoxmlFixture = async (): Promise<ArrayBuffer> => {
  const zip = await JSZip.loadAsync(await createEmptyDocx());
  const contentTypes = await zip.file("[Content_Types].xml")?.async("text");
  const relationships = await zip.file("word/_rels/document.xml.rels")?.async("text");
  if (!contentTypes || !relationships) {
    return panic("The empty DOCX fixture is missing package relationships.");
  }

  zip.file(
    "[Content_Types].xml",
    contentTypes.replace(
      "</Types>",
      `<Override PartName="/word/numbering.xml" ContentType="${NUMBERING_CONTENT_TYPE}"/></Types>`,
    ),
  );
  zip.file(
    "word/_rels/document.xml.rels",
    relationships.replace(
      "</Relationships>",
      `<Relationship Id="rIdOutlineFixture" Type="${NUMBERING_RELATIONSHIP}" Target="numbering.xml"/></Relationships>`,
    ),
  );
  zip.file(
    "word/numbering.xml",
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:abstractNum w:abstractNumId="7">
    <w:lvl w:ilvl="0">
      <w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/>
      <w:pPr><w:outlineLvl w:val="2"/></w:pPr>
    </w:lvl>
  </w:abstractNum>
  <w:num w:numId="7"><w:abstractNumId w:val="7"/></w:num>
</w:numbering>`,
  );
  zip.file(
    "word/styles.xml",
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>
  <w:style w:type="paragraph" w:styleId="Heading1"><w:name w:val="heading 1"/></w:style>
  <w:style w:type="paragraph" w:styleId="Heading2"><w:name w:val="heading 2"/></w:style>
  <w:style w:type="paragraph" w:styleId="ClauseHeading">
    <w:name w:val="Clause heading"/><w:basedOn w:val="Heading2"/>
  </w:style>
  <w:style w:type="paragraph" w:styleId="NumberedBody">
    <w:name w:val="Numbered body"/><w:basedOn w:val="Normal"/>
  </w:style>
</w:styles>`,
  );
  zip.file(
    "word/document.xml",
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr><w:r><w:t>Real Heading 1</w:t></w:r></w:p>
    <w:p><w:pPr><w:pStyle w:val="Heading2"/></w:pPr><w:r><w:t>Real Heading 2</w:t></w:r></w:p>
    <w:p><w:pPr><w:pStyle w:val="ClauseHeading"/></w:pPr><w:r><w:t>Inherited Heading</w:t></w:r></w:p>
    <w:p><w:pPr><w:pStyle w:val="NumberedBody"/><w:numPr><w:ilvl w:val="0"/><w:numId w:val="7"/></w:numPr></w:pPr>
      <w:r><w:t>Numbered body with numbering-only outline level</w:t></w:r></w:p>
    <w:p><w:pPr><w:pStyle w:val="Normal"/><w:outlineLvl w:val="3"/></w:pPr>
      <w:r><w:t>Direct paragraph outline</w:t></w:r></w:p>
    <w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr>
  </w:body>
</w:document>`,
  );

  return zip.generateAsync({ type: "arraybuffer" });
};
