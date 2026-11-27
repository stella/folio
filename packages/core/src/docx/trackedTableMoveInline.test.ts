import { expect, test } from "bun:test";

import {
  applyDocumentOp,
  applyDocumentOps,
  compileEditorIntent,
  createEditorIntentIdAllocator,
  DOCUMENT_OP_REFUSAL_REASONS,
  DOCUMENT_OP_TYPES,
  type DocumentOp,
  OP_STORIES,
  REVISION_DECISIONS,
} from "@stll/docx-core/ops";
import { ZipArchive } from "@stll/docx-core/zip";
import { base64ToBytes } from "@stll/utf8-boundaries";

import { FolioDocxReviewer } from "../ai-edits/headless";
import { FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION } from "../document-operations";
import type { BlockContent, Document, Table } from "../types/document";
import { RELATIONSHIP_TYPES } from "./relsParser";
import { FOLIO_REVIEW_HISTORY_NAMESPACE } from "./reviewHistoryNamespace";
import { createEmptyDocx } from "./rezip";
import { serializeParagraph } from "./serializer/paragraphSerializer";
import { serializeTable } from "./serializer/tableSerializer";
import { parseTable } from "./tableParser";
import { parseXmlDocument } from "./xmlParser";

const ownerId = "00000001";
const sourceId = "00000002";
const plain = "<w:r><w:t>Owner</w:t></w:r>";
const direction = '<w:dir w:val="rtl">';
const link = '<w:hyperlink w:anchor="Target">';
const atoms =
  '<w:r><w:t>Linked</w:t><w:tab/><w:br/><w:sym w:font="Symbol" w:char="F020"/><w:softHyphen/><w:noBreakHyphen/></w:r>';
const inlinePicture =
  '<w:r><w:drawing><wp:inline><wp:extent cx="914400" cy="914400"/><wp:docPr id="27" name="Inline logo"/><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic><pic:nvPicPr><pic:cNvPr id="0" name="logo.png"/><pic:cNvPicPr/></pic:nvPicPr><pic:blipFill><a:blip r:embed="rIdLogo"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill><pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="914400" cy="914400"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r>';
const PNG =
  "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg==";

const readTable = (xml: string): Table => {
  const element = parseXmlDocument(
    xml.replace(
      "<w:tbl",
      `<w:tbl xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:folio="${FOLIO_REVIEW_HISTORY_NAMESPACE}"`,
    ),
  );
  if (element === null) throw new TypeError("Expected well-formed XML.");
  const table = parseTable(element, new Map(), null, null, null, null);
  if (table === undefined) throw new TypeError("Expected a table.");
  return table;
};

const tableOf = (document: Document): Table => {
  const table = document.package.document.content.at(0);
  if (table?.type !== "table") throw new TypeError("Expected a table.");
  return table;
};

const savedContentOf = (block: BlockContent): BlockContent => {
  if (block.type === "blockCustomXml")
    return { ...block, content: block.content.map(savedContentOf) };
  if (block.type !== "paragraph") return block;
  const { paraId, content } = block;
  return {
    type: "paragraph",
    ...(paraId === undefined ? {} : { paraId }),
    content,
  };
};
const inlineContentOf = (table: Table) =>
  table.rows.map(({ cells }) =>
    cells.map(({ content }) => content.map(savedContentOf)),
  );

type SourceWithOptions = {
  content: string;
  owner?: string;
  wrap?: (paragraph: string) => string;
};
const sourceWith = ({
  content,
  owner = plain,
  wrap = (paragraph) => paragraph,
}: SourceWithOptions): Document => {
  const absorbed = wrap(`<w:p w14:paraId="${sourceId}">${content}</w:p>`);
  const table = readTable(
    '<w:tbl><w:tblPr/><w:tblGrid><w:gridCol w:w="1000"/><w:gridCol w:w="1000"/></w:tblGrid>' +
      `<w:tr><w:tc><w:p w14:paraId="${ownerId}">${owner}</w:p></w:tc>` +
      `<w:tc>${absorbed}</w:tc></w:tr></w:tbl>`,
  );
  return {
    package: {
      document: {
        content: [
          table,
          { type: "paragraph", paraId: "00000003", content: [] },
        ],
      },
    },
  };
};
const directMerge = {
  type: DOCUMENT_OP_TYPES.MERGE_CELLS,
  story: OP_STORIES.MAIN,
  blockId: ownerId,
  top: 0,
  bottom: 1,
  left: 0,
  right: 2,
  newBlockIds: [],
} as const;
const trackedMerge = {
  ...directMerge,
  newBlockIds: ["00000004"],
  revision: { id: 100, author: "Reviewer", date: "2026-11-27T00:00:00Z" },
  newIds: { revision: [101, 102, 103, 104, 105, 106] },
} as const;

const allocateTrackedMerge = (document: Document) => {
  const { newBlockIds: _fresh, ...operation } = directMerge;
  const intent = { type: "table", operation } as const;
  const { revisionId, newIds } = createEditorIntentIdAllocator()(
    document,
    intent,
  );
  const { ops } = compileEditorIntent(document, {
    intent,
    mode: {
      type: "suggesting",
      revision: { id: revisionId, author: "Reviewer" },
      newIds,
    },
  }).unwrap();
  const merge = ops.at(0);
  if (merge?.type !== DOCUMENT_OP_TYPES.MERGE_CELLS)
    throw new TypeError("Expected a tracked merge.");
  return { merge, controlIds: newIds.control ?? [] };
};

test.each([
  '<m:oMath><w:bookmarkStart w:id="42" w:name="Formula"/><m:r><m:t>x</m:t></m:r><w:bookmarkEnd w:id="42"/></m:oMath>',
  '<w:proofErr w:type="spellStart"/><w:r><w:t>A</w:t></w:r>',
])(
  "a tracked move preserves range ownership when its payload cannot move as a unit: %s",
  (content) => {
    const document = sourceWith({ content });
    const before = structuredClone(document);
    const result = applyDocumentOp(document, trackedMerge);
    expect(result.isErr() && result.error.reason).toBe(
      DOCUMENT_OP_REFUSAL_REASONS.UNTRACKABLE,
    );
    expect(document).toStrictEqual(before);
  },
);

const assertSavedMove = (
  original: Document,
  operation: Extract<DocumentOp, { type: "mergeCells" }> = trackedMerge,
) => {
  const direct = applyDocumentOp(original, directMerge).unwrap().document;
  const pending = applyDocumentOp(original, operation).unwrap();
  const grid = tableOf(pending.document).formatting?.gridChange;
  if (grid?.geometryChange?.type !== "mergeCells")
    throw new TypeError("Expected a tracked merge geometry.");
  const move = grid.geometryChange.moves?.at(0);
  const reviewId =
    (move?.kind === "empty" ? move.destinationMark : move?.fromRevision) ??
    grid.id;
  expect(
    applyDocumentOps(pending.document, pending.inverse).unwrap().document,
  ).toEqual(original);
  const xml = serializeTable(tableOf(pending.document), serializeParagraph);
  const reopened = {
    ...original,
    package: {
      document: {
        content: [
          readTable(xml),
          ...original.package.document.content.slice(1),
        ],
      },
    },
  };
  for (const decision of [
    REVISION_DECISIONS.ACCEPT,
    REVISION_DECISIONS.REJECT,
  ]) {
    const expected = decision === REVISION_DECISIONS.ACCEPT ? direct : original;
    for (const document of [pending.document, reopened]) {
      const resolved = applyDocumentOp(document, {
        type: DOCUMENT_OP_TYPES.RESOLVE_REVISION,
        story: OP_STORIES.MAIN,
        revisionIds: [reviewId],
        decision,
      }).unwrap().document;
      if (document === pending.document) {
        expect(resolved).toEqual(expected);
      } else {
        expect(inlineContentOf(tableOf(resolved))).toEqual(
          inlineContentOf(tableOf(expected)),
        );
      }
    }
  }
  return xml;
};

test.each([
  `${direction}${link}${atoms}</w:hyperlink></w:dir>`,
  `${link}${direction}${atoms}</w:dir></w:hyperlink>`,
  `${atoms}${link}${atoms}</w:hyperlink>${direction}${atoms}</w:dir>`,
  "<w:r><w:tab/></w:r>",
  inlinePicture,
  inlinePicture.replace('<pic:cNvPr id="0"', '<pic:cNvPr id="27"'),
  '<w:r><w:ruby><w:rubyPr><w:rubyAlign w:val="center"/><w:hps w:val="10"/><w:hpsRaise w:val="12"/><w:hpsBaseText w:val="20"/><w:lid w:val="ja-JP"/></w:rubyPr><w:rt><w:r><w:t>とう</w:t></w:r></w:rt><w:rubyBase><w:r><w:t>東</w:t></w:r></w:rubyBase></w:ruby></w:r>',
  "<m:oMath><m:f><m:num><m:r><m:t>1</m:t></m:r></m:num><m:den><m:r><m:t>2</m:t></m:r></m:den></m:f></m:oMath>",
  '<q:oMathPara xmlns:q="http://purl.oclc.org/ooxml/officeDocument/math"><q:oMath><q:r><q:t xml:space="preserve"> 1 + 2 </q:t></q:r></q:oMath></q:oMathPara>',
  '<w:proofErr w:type="gramStart"/><w:proofErr w:type="spellStart"/><w:r><w:t>Teh text.</w:t></w:r><w:proofErr w:type="spellEnd"/><w:proofErr w:type="gramEnd"/>',
  '<w:customXml w:element="clause" w:uri="urn:example"><w:r><w:t>A clause.</w:t></w:r></w:customXml>',
  '<w:smartTag w:element="term" w:uri="urn:example"><w:r><w:t>A term.</w:t></w:r></w:smartTag>',
])("tracked table merges preserve moved inline content: %s", (content) => {
  const picture = content.includes("rIdLogo");
  const original = sourceWith({
    content,
    ...(picture
      ? { owner: inlinePicture.replace('docPr id="27"', 'docPr id="1"') }
      : {}),
  });
  const xml = assertSavedMove(original);
  expect(xml).toContain("<w:moveFrom ");
  expect(xml).toContain("<w:moveTo ");
  if (picture) {
    const drawingIds = [...xml.matchAll(/<wp:docPr\b[^>]*\bid="(\d+)"/gu)].map(
      (match) => match.at(1),
    );
    expect(drawingIds).toHaveLength(3);
    expect(new Set(drawingIds).size).toBe(3);
    expect(drawingIds).toContain("1");
    expect(drawingIds).toContain("27");
    expect(xml.match(/r:embed="rIdLogo"/gu)).toHaveLength(3);
  }
});

test("a tracked merge moves each enclosing custom XML block with its contents", () => {
  const original = sourceWith({
    content: plain,
    wrap: (paragraph) =>
      `<w:customXml w:uri="urn:test" w:element="clause"><w:customXml w:uri="urn:test" w:element="term">${paragraph}</w:customXml></w:customXml>`,
  });
  const xml = assertSavedMove(original);
  expect(xml.match(/<w:customXml\b/gu)).toHaveLength(4);
  expect(xml).toMatch(/<w:customXml[^>]*><w:p[^>]*>.*?<w:moveTo /u);
  expect(xml).toMatch(/<w:customXml[^>]*><w:p[^>]*>.*?<w:moveFrom /u);
});

test("a copied nested table keeps its tagged cell paragraph and stable identity", () => {
  const original = sourceWith({
    content: plain,
    wrap: (paragraph) =>
      `<w:tbl><w:tblPr/><w:tblGrid><w:gridCol w:w="500"/></w:tblGrid><w:tr><w:tc><w:customXml w:element="tag">${paragraph}</w:customXml></w:tc></w:tr></w:tbl><w:p w14:paraId="00000005"/>`,
  });
  const xml = assertSavedMove(original, {
    ...trackedMerge,
    newBlockIds: ["00000004", "00000006"],
  });
  expect(xml.match(/<w:customXml\b/gu)).toHaveLength(2);
});

test("a moved nested table preserves drawing and rich-text control identities", () => {
  const original = sourceWith({
    owner: inlinePicture.replace('docPr id="27"', 'docPr id="1"'),
    content: `<w:sdt><w:sdtPr><w:id w:val="72"/></w:sdtPr><w:sdtContent>${inlinePicture}</w:sdtContent></w:sdt>`,
    wrap: (paragraph) =>
      '<w:tbl><w:tblGrid><w:gridCol w:w="500"/></w:tblGrid><w:tr><w:tc>' +
      `<w:sdt><w:sdtPr><w:id w:val="71"/></w:sdtPr><w:sdtContent>${paragraph}</w:sdtContent></w:sdt>` +
      '</w:tc></w:tr></w:tbl><w:p w14:paraId="00000005"/>',
  });
  const { merge, controlIds } = allocateTrackedMerge(original);
  expect(controlIds).toHaveLength(2);
  const xml = assertSavedMove(original, merge);
  expect(xml.match(/r:embed="rIdLogo"/gu)).toHaveLength(3);
  expect(xml.match(/<wp:docPr id="1"/gu)).toHaveLength(1);
  expect(xml.match(/<wp:docPr id="2"/gu)).toHaveLength(1);
  expect(xml.match(/<wp:docPr id="27"/gu)).toHaveLength(1);
  for (const id of [71, 72, ...controlIds])
    expect(xml.match(new RegExp(`<w:id w:val="${id}"`, "gu"))).toHaveLength(1);
});

test("a moved nested table preserves controls shared by rows and cells", () => {
  const original = sourceWith({
    content: plain,
    wrap: (paragraph) =>
      '<w:tbl><w:tblGrid><w:gridCol w:w="500"/><w:gridCol w:w="500"/></w:tblGrid>' +
      '<w:customXml w:element="controlledRows"><w:sdt><w:sdtPr><w:id w:val="73"/></w:sdtPr><w:sdtContent>' +
      '<w:tr><w:sdt><w:sdtPr><w:id w:val="74"/></w:sdtPr><w:sdtContent>' +
      `<w:tc>${paragraph}</w:tc><w:tc><w:p w14:paraId="00000007"/></w:tc>` +
      "</w:sdtContent></w:sdt></w:tr>" +
      '<w:tr><w:tc><w:p w14:paraId="00000008"/></w:tc><w:tc><w:p w14:paraId="00000009"/></w:tc></w:tr>' +
      '</w:sdtContent></w:sdt></w:customXml></w:tbl><w:p w14:paraId="00000005"/>',
  });
  const { merge, controlIds } = allocateTrackedMerge(original);
  expect(controlIds).toHaveLength(2);
  const xml = assertSavedMove(original, merge);
  expect(xml.match(/<w:customXml w:element="controlledRows"/gu)).toHaveLength(
    2,
  );
  for (const id of [73, 74, ...controlIds])
    expect(xml.match(new RegExp(`<w:id w:val="${id}"`, "gu"))).toHaveLength(1);
});

test("the saved picture is available after each public table review decision", async () => {
  const source = tableOf(sourceWith({ content: inlinePicture }));
  const tableXml = serializeTable(source, serializeParagraph).replace(
    "<w:tbl",
    '<w:tbl xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"',
  );
  const seed = await ZipArchive.read(await createEmptyDocx());
  seed.set(
    "word/document.xml",
    `<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"><w:body>${tableXml}</w:body></w:document>`,
  );
  const path = "word/media/logo.png";
  const relsPath = "word/_rels/document.xml.rels";
  const rels = await seed.entry(relsPath)?.text();
  const typesPath = "[Content_Types].xml";
  const types = await seed.entry(typesPath)?.text();
  if (rels === undefined || types === undefined)
    throw new TypeError("Expected a DOCX package.");
  seed.set(
    relsPath,
    rels.replace(
      "</Relationships>",
      `<Relationship Id="rIdLogo" Type="${RELATIONSHIP_TYPES.image}" Target="media/logo.png"/></Relationships>`,
    ),
  );
  if (!types.includes('Extension="png"'))
    seed.set(
      typesPath,
      types.replace(
        "</Types>",
        '<Default Extension="png" ContentType="image/png"/></Types>',
      ),
    );
  seed.set(path, base64ToBytes(PNG));
  const sourceReviewer = await FolioDocxReviewer.fromBuffer(
    (await seed.generate({ compression: "store" })).buffer,
    { session: "canonical" },
  );
  const result = sourceReviewer.applyDocumentOperations(
    {
      version: FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
      mode: "tracked-changes",
      operations: [
        {
          id: "merge-logo-cell",
          type: "mergeTableCells",
          blockId: ownerId,
          endBlockId: sourceId,
        },
      ],
    },
    { undefinedReferences: "refuse" },
  );
  expect(result.skipped).toEqual([]);
  expect(result.applied).toHaveLength(1);
  const pending = await sourceReviewer.toBuffer();
  for (const decision of ["accept", "reject"] as const) {
    const reviewer = await FolioDocxReviewer.fromBuffer(pending.slice(0), {
      session: "canonical",
    });
    if (decision === "accept") reviewer.acceptAll();
    else reviewer.rejectAll();
    const archive = await ZipArchive.read(await reviewer.toBuffer());
    const main = archive.entry("word/document.xml");
    if (main === undefined)
      throw new TypeError("Expected the main document part.");
    const xml = await main.text();
    expect(xml.match(/<wp:docPr\b[^>]*\bid="27"/gu)).toHaveLength(1);
    expect(xml).not.toContain("<w:moveTo ");
    expect(xml).not.toContain("<w:moveFrom ");
    expect(archive.entry(path)).toBeDefined();
  }
});
