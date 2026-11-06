/**
 * Rejecting a paragraph a tracked edit added takes it out again, wherever it
 * landed: after a table that ends a content control, or after a paragraph with
 * a text box drawn in it.
 */

import { describe, expect, test } from "bun:test";

import { ZipArchive } from "@stll/docx-core/zip";

import { createDocx } from "../docx/rezip";
import { createEmptyDocument } from "../utils/createDocument";
import { FolioDocxReviewer } from "./headless";

const NS =
  'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"';

const paragraph = (id: string, text: string) =>
  `<w:p w14:paraId="${id}"><w:r><w:t>${text}</w:t></w:r></w:p>`;

const TABLE = `<w:tbl><w:tblPr/><w:tblGrid><w:gridCol w:w="1000"/></w:tblGrid><w:tr><w:tc><w:tcPr/>${paragraph("10000001", "Cell")}</w:tc></w:tr></w:tbl>`;

const TEXT_BOX = `<w:r><w:drawing><wp:anchor distT="0" distB="0" distL="0" distR="0" simplePos="0" relativeHeight="1" behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1"><wp:simplePos x="0" y="0"/><wp:positionH relativeFrom="column"><wp:posOffset>0</wp:posOffset></wp:positionH><wp:positionV relativeFrom="paragraph"><wp:posOffset>0</wp:posOffset></wp:positionV><wp:extent cx="914400" cy="914400"/><wp:wrapSquare wrapText="bothSides"/><wp:docPr id="1" name="Box"/><a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"><wps:wsp><wps:cNvSpPr txBox="1"/><wps:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="914400" cy="914400"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></wps:spPr><wps:txbx><w:txbxContent>${paragraph("10000005", "Boxed")}</w:txbxContent></wps:txbx><wps:bodyPr/></wps:wsp></a:graphicData></a:graphic></wp:anchor></w:drawing></w:r>`;

const control = (content: string) =>
  `<w:sdt><w:sdtPr><w:tag w:val="control"/></w:sdtPr><w:sdtContent>${content}</w:sdtContent></w:sdt>`;

const docx = async (body: string): Promise<ArrayBuffer> => {
  const zip = await ZipArchive.read(await createDocx(createEmptyDocument()));
  zip.set(
    "word/document.xml",
    `<?xml version="1.0" encoding="UTF-8"?><w:document ${NS}><w:body>${body}</w:body></w:document>`,
  );
  return (await zip.generate()).slice().buffer;
};

const texts = (reviewer: FolioDocxReviewer) =>
  reviewer.getContent().map(({ text }) => text);

const CASES = [
  {
    name: "a table that ends a content control",
    body: `${control(TABLE)}${paragraph("10000002", "Tail")}`,
    anchor: "10000001",
  },
  {
    name: "a paragraph with a text box drawn in it, ending the body",
    body: `<w:p w14:paraId="10000003">${TEXT_BOX}<w:r><w:t>Host</w:t></w:r></w:p>`,
    anchor: "10000003",
  },
  {
    name: "a table that ends the body",
    body: `${paragraph("10000002", "Head")}${TABLE}`,
    anchor: "10000001",
  },
  {
    name: "a table that ends the body but for a text box drawn before it",
    body: `<w:p w14:paraId="10000003">${TEXT_BOX}<w:r><w:t>Host</w:t></w:r></w:p>${TABLE}`,
    anchor: "10000001",
  },
] as const;

describe("rejecting a paragraph added after", () => {
  for (const { name, body, anchor } of CASES) {
    const edit = async () => {
      const bytes = await docx(body);
      const source = texts(await FolioDocxReviewer.fromBuffer(bytes.slice(0)));
      const reviewer = await FolioDocxReviewer.fromBuffer(bytes.slice(0));
      const result = reviewer.applyDocumentOperations(
        {
          version: 1,
          mode: "tracked-changes",
          operations: [
            {
              id: "add",
              type: "insertAfterBlock",
              blockId: anchor,
              text: "Added",
            },
          ],
        },
        { undefinedReferences: "refuse" },
      );
      expect(result.applied).toHaveLength(1);
      expect(texts(reviewer)).toContain("Added");
      return { reviewer, source };
    };

    test(`${name}: all at once`, async () => {
      const { reviewer, source } = await edit();
      reviewer.rejectAll();
      expect(texts(reviewer)).toEqual(source);
      const reopened = await FolioDocxReviewer.fromBuffer(
        await reviewer.toBuffer(),
      );
      expect(texts(reopened)).toEqual(source);
    });

    test(`${name}: change by change`, async () => {
      const { reviewer, source } = await edit();
      // The words first: with them still there, a break with nothing to join
      // to keeps its place.
      const changes = reviewer.getChanges();
      for (const change of changes.filter(({ type }) => type === "insertion")) {
        reviewer.rejectChange(change);
      }
      for (const change of reviewer.getChanges()) {
        reviewer.rejectChange(change);
      }
      expect(texts(reviewer)).toEqual(source);
    });

    test(`${name}: after a reopen`, async () => {
      const { reviewer, source } = await edit();
      const reopened = await FolioDocxReviewer.fromBuffer(
        await reviewer.toBuffer(),
      );
      reopened.rejectAll();
      expect(texts(reopened)).toEqual(source);
    });

    test(`${name}: accepting keeps it`, async () => {
      const { reviewer, source } = await edit();
      reviewer.acceptAll();
      expect(texts(reviewer)).toContain("Added");
      for (const text of source) expect(texts(reviewer)).toContain(text);
    });
  }
});

test("an unrelated rejection leaves an added final paragraph alone", async () => {
  const reviewer = await FolioDocxReviewer.fromBuffer(
    await docx(`${paragraph("10000002", "Head")}${TABLE}`),
  );
  reviewer.applyDocumentOperations(
    {
      version: 1,
      mode: "tracked-changes",
      operations: [
        {
          id: "add",
          type: "insertAfterBlock",
          blockId: "10000001",
          text: "Added",
        },
        {
          id: "head",
          type: "replaceInBlock",
          blockId: "10000002",
          find: "Head",
          replace: "Top",
        },
      ],
    },
    { undefinedReferences: "refuse" },
  );
  const headChange = reviewer.getChanges().find(({ text }) => text === "Top");
  if (!headChange) throw new Error("The heading change is missing");
  reviewer.rejectChange(headChange);
  expect(texts(reviewer)).toContain("Added");
  reviewer.rejectAll();
  expect(texts(reviewer)).toEqual(["Head", "Cell"]);
});

test("accepting an added final paragraph keeps it as an ordinary one", async () => {
  const reviewer = await FolioDocxReviewer.fromBuffer(
    await docx(`${paragraph("10000002", "Head")}${TABLE}`),
  );
  reviewer.applyDocumentOperations(
    {
      version: 1,
      mode: "tracked-changes",
      operations: [
        {
          id: "add",
          type: "insertAfterBlock",
          blockId: "10000001",
          text: "Added",
        },
      ],
    },
    { undefinedReferences: "refuse" },
  );
  reviewer.acceptAll();
  const reopened = await FolioDocxReviewer.fromBuffer(
    await reviewer.toBuffer(),
  );
  expect(texts(reopened)).toEqual(["Head", "Cell", "Added"]);
  expect(reopened.getChanges()).toEqual([]);
});

describe("a tracked removed break on a paragraph with a text box drawn in it", () => {
  const BODY = `<w:p w14:paraId="10000003"><w:pPr><w:rPr><w:del w:id="9" w:author="A" w:date="2026-01-01T00:00:00Z"/></w:rPr></w:pPr>${TEXT_BOX}<w:r><w:t>Host</w:t></w:r></w:p>${paragraph("10000004", "Next")}${paragraph("10000006", "Tail")}`;

  const open = async () => FolioDocxReviewer.fromBuffer(await docx(BODY));

  test("accepting joins the two, and the box follows", async () => {
    const reviewer = await open();
    reviewer.acceptAll();
    expect(texts(reviewer)).toEqual(["HostNext", "Boxed", "Tail"]);
    const reopened = await FolioDocxReviewer.fromBuffer(
      await reviewer.toBuffer(),
    );
    expect(texts(reopened)).toEqual(["HostNext", "Boxed", "Tail"]);
  });

  test("rejecting keeps the two apart", async () => {
    const reviewer = await open();
    reviewer.rejectAll();
    expect(texts(reviewer)).toEqual(["Host", "Boxed", "Next", "Tail"]);
  });
});
