import type { ImportedIdCollisionKind } from "@stll/docx-core/model";

const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const paragraph = (text: string) => `<w:p><w:r><w:t>${text}</w:t></w:r></w:p>`;
const numbering = (definitions: string) =>
  `<w:numbering xmlns:w="${W}">${definitions}</w:numbering>`;
const abstract = (format: string) =>
  `<w:abstractNum w:abstractNumId="41"><w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="${format}"/><w:lvlText w:val="%1."/><w:lvlJc w:val="left"/></w:lvl></w:abstractNum>`;
const num = (id: string) =>
  `<w:num w:numId="73"><w:abstractNumId w:val="${id}"/></w:num>`;
const drawing = `<w:r><w:drawing><wp:inline xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"><wp:extent cx="914400" cy="914400"/><wp:docPr id="317" name="Collision fixture"/><a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:graphicData uri="urn:fixture"/></a:graphic></wp:inline></w:drawing></w:r>`;

type CollisionFixture = {
  body: string;
  parts?: Record<string, string>;
  contentTypes?: { partName: string; contentType: string }[];
  documentRelationships?: { id: string; type: string; target: string }[];
};
const noteFixture = (kind: "footnote" | "endnote"): CollisionFixture => ({
  body: `<w:p><w:r><w:t>Note</w:t></w:r><w:r><w:${kind}Reference w:id="23"/></w:r></w:p>`,
  parts: {
    [`word/${kind}s.xml`]: `<w:${kind}s xmlns:w="${W}"><w:${kind} w:id="23">${paragraph("First note")}</w:${kind}><w:${kind} w:id="23">${paragraph("Later note")}</w:${kind}></w:${kind}s>`,
  },
  contentTypes: [
    {
      partName: `/word/${kind}s.xml`,
      contentType: `application/vnd.openxmlformats-officedocument.wordprocessingml.${kind}s+xml`,
    },
  ],
  documentRelationships: [
    {
      id: `rId${kind}`,
      type: `http://schemas.openxmlformats.org/officeDocument/2006/relationships/${kind}s`,
      target: `${kind}s.xml`,
    },
  ],
});
const numberingFixture = (xml: string): CollisionFixture => ({
  body: `<w:p><w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="73"/></w:numPr></w:pPr><w:r><w:t>Numbered text</w:t></w:r></w:p>`,
  parts: { "word/numbering.xml": xml },
  contentTypes: [
    {
      partName: "/word/numbering.xml",
      contentType:
        "application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml",
    },
  ],
  documentRelationships: [
    {
      id: "rIdNumbering",
      type: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering",
      target: "numbering.xml",
    },
  ],
});

export const IMPORTED_ID_COLLISION_FIXTURES = {
  bookmark: {
    body: `<w:p><w:bookmarkStart w:id="37" w:name="first"/><w:r><w:t>First</w:t></w:r><w:bookmarkEnd w:id="37"/><w:bookmarkStart w:id="37" w:name="later"/><w:r><w:t>Later</w:t></w:r><w:bookmarkEnd w:id="37"/></w:p>`,
  },
  comment: {
    body: `<w:p><w:commentRangeStart w:id="23"/><w:r><w:t>Commented</w:t></w:r><w:commentRangeEnd w:id="23"/><w:r><w:commentReference w:id="23"/></w:r></w:p>`,
    parts: {
      "word/comments.xml": `<w:comments xmlns:w="${W}"><w:comment w:id="23" w:author="Fixture">${paragraph("First comment")}</w:comment><w:comment w:id="23" w:author="Fixture">${paragraph("Later comment")}</w:comment></w:comments>`,
    },
    contentTypes: [
      {
        partName: "/word/comments.xml",
        contentType:
          "application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml",
      },
    ],
    documentRelationships: [
      {
        id: "rIdComments",
        type: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments",
        target: "comments.xml",
      },
    ],
  },
  footnote: noteFixture("footnote"),
  endnote: noteFixture("endnote"),
  revision: {
    body: `<w:p><w:ins w:id="83" w:author="Fixture" w:date="2024-01-01T00:00:00Z"><w:r><w:t>First revision</w:t></w:r></w:ins><w:ins w:id="83" w:author="Fixture" w:date="2024-01-01T00:00:00Z"><w:r><w:t>Later revision</w:t></w:r></w:ins></w:p>`,
  },
  paraId: {
    body: `<w:p xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" w14:paraId="15AB23CD"><w:r><w:t>First</w:t></w:r></w:p><w:p xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" w14:paraId="15AB23CD"><w:r><w:t>Later</w:t></w:r></w:p>`,
  },
  textId: {
    body: `<w:p xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" w14:textId="15AB23CD"><w:r><w:t>First</w:t></w:r></w:p><w:p xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" w14:textId="15AB23CD"><w:r><w:t>Later</w:t></w:r></w:p>`,
  },
  numId: numberingFixture(
    numbering(
      abstract("decimal") +
        abstract("upperRoman").replace(
          'abstractNumId="41"',
          'abstractNumId="42"',
        ) +
        num("41") +
        num("42"),
    ),
  ),
  abstractNumId: numberingFixture(
    numbering(abstract("decimal") + abstract("upperRoman") + num("41")),
  ),
  drawing: { body: `<w:p>${drawing}${drawing}</w:p>` },
  relationship: {
    body: `<w:p><w:hyperlink xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:id="rIdLink"><w:r><w:t>Linked text</w:t></w:r></w:hyperlink></w:p>`,
    parts: {
      "word/_rels/document.xml.rels": `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdStyles" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/><Relationship Id="rIdLink" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="https://example.org/first" TargetMode="External"/><Relationship Id="rIdLink" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="https://example.org/later" TargetMode="External"/></Relationships>`,
    },
  },
} as const satisfies Record<ImportedIdCollisionKind, CollisionFixture>;

export const importedIdCollisionFixtures = Object.entries(
  IMPORTED_ID_COLLISION_FIXTURES,
).map(([kind, fixture]) =>
  Object.assign({ filename: `duplicate-imported-${kind}.docx` }, fixture),
);
