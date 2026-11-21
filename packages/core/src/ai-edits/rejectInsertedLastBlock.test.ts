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

const SESSIONS = ["legacy", "canonical"] as const;
// The canonical reader supplies a closing paragraph when the body's source
// ends with a table. Editing the body writes that paragraph.
const TABLE_BODY_END = { legacy: [], canonical: [""] } as const;

const CASES = [
  {
    name: "a table that ends a content control",
    body: `${control(TABLE)}${paragraph("10000002", "Tail")}`,
    anchor: "10000001",
    accepted: ["Cell", "Added", "Tail"],
    bodyEnd: "paragraph",
  },
  {
    name: "a paragraph with a text box drawn in it, ending the body",
    body: `<w:p w14:paraId="10000003">${TEXT_BOX}<w:r><w:t>Host</w:t></w:r></w:p>`,
    anchor: "10000003",
    accepted: ["Host", "Boxed", "Added"],
    bodyEnd: "paragraph",
  },
  {
    name: "a table that ends the body",
    body: `${paragraph("10000002", "Head")}${TABLE}`,
    anchor: "10000001",
    accepted: ["Head", "Cell", "Added"],
    bodyEnd: "table",
  },
  {
    name: "a table that ends the body but for a text box drawn before it",
    body: `<w:p w14:paraId="10000003">${TEXT_BOX}<w:r><w:t>Host</w:t></w:r></w:p>${TABLE}`,
    anchor: "10000001",
    accepted: ["Host", "Boxed", "Cell", "Added"],
    bodyEnd: "table",
  },
] as const;

for (const { name, body, anchor, accepted, bodyEnd } of CASES)
  describe.each(SESSIONS)(
    `on the %s session, a paragraph added after ${name}`,
    (session) => {
      const open = (bytes: ArrayBuffer) =>
        FolioDocxReviewer.fromBuffer(bytes, { session });
      const edit = async () => {
        const bytes = await docx(body);
        const source = texts(await open(bytes.slice(0)));
        const reviewer = await open(bytes.slice(0));
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
        return { reviewer, source, undoHandle: result.undoHandle };
      };

      test("rejecting all at once takes it out", async () => {
        const { reviewer, source } = await edit();
        reviewer.rejectAll();
        expect(texts(reviewer)).toEqual(source);
        const reopened = await open(await reviewer.toBuffer());
        expect(texts(reopened)).toEqual(source);
      });

      test("rejecting change by change takes it out", async () => {
        const { reviewer, source } = await edit();
        // The words first: with them still there, a break with nothing to join
        // to keeps its place.
        const changes = reviewer.getChanges();
        for (const change of changes.filter(
          ({ type }) => type === "insertion",
        )) {
          reviewer.rejectChange(change);
        }
        for (const change of reviewer.getChanges()) {
          reviewer.rejectChange(change);
        }
        expect(texts(reviewer)).toEqual(source);
      });

      test("rejecting after a reopen takes it out", async () => {
        const { reviewer, source } = await edit();
        const reopened = await open(await reviewer.toBuffer());
        reopened.rejectAll();
        expect(texts(reopened)).toEqual(source);
      });

      test("accepting keeps it as an ordinary paragraph", async () => {
        const { reviewer } = await edit();
        const expected =
          bodyEnd === "table"
            ? [...accepted, ...TABLE_BODY_END[session]]
            : [...accepted];
        const saved = await reviewer.toBuffer();
        reviewer.acceptAll();
        expect(texts(reviewer)).toEqual(expected);
        const reopened = await open(await reviewer.toBuffer());
        expect(texts(reopened)).toEqual(expected);
        expect(reopened.getChanges()).toEqual([]);
        const acceptedAfterReopen = await open(saved);
        acceptedAfterReopen.acceptAll();
        expect(texts(acceptedAfterReopen)).toEqual(expected);
        expect(acceptedAfterReopen.getChanges()).toEqual([]);
      });

      test("undoing the edit restores the document", async () => {
        const { reviewer, source, undoHandle } = await edit();
        if (undoHandle === null) throw new Error("The edit has no undo handle");
        expect(reviewer.undoDocumentOperations(undoHandle).status).toBe(
          "undone",
        );
        expect(texts(reviewer)).toEqual(source);
        expect(reviewer.getChanges()).toEqual([]);
        const reopened = await open(await reviewer.toBuffer());
        expect(texts(reopened)).toEqual(source);
      });
    },
  );

const openLegacy = (bytes: ArrayBuffer) =>
  FolioDocxReviewer.fromBuffer(bytes, { session: "legacy" });

test.each(SESSIONS)(
  "an unrelated rejection leaves an added final paragraph alone (%s)",
  async (session) => {
    const reviewer = await FolioDocxReviewer.fromBuffer(
      await docx(`${paragraph("10000002", "Head")}${TABLE}`),
      { session },
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
    expect(texts(reviewer)).toEqual([
      "Head",
      "Cell",
      ...TABLE_BODY_END[session],
    ]);
  },
);

test.each(SESSIONS)(
  "accepting an added final paragraph keeps it as an ordinary one (%s)",
  async (session) => {
    const open = (bytes: ArrayBuffer) =>
      FolioDocxReviewer.fromBuffer(bytes, { session });
    const reviewer = await open(
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
    const reopened = await open(await reviewer.toBuffer());
    expect(texts(reopened)).toEqual([
      "Head",
      "Cell",
      "Added",
      ...TABLE_BODY_END[session],
    ]);
    expect(reopened.getChanges()).toEqual([]);
  },
);

describe("a tracked removed break on a paragraph with a text box drawn in it", () => {
  const BODY = `<w:p w14:paraId="10000003"><w:pPr><w:rPr><w:del w:id="9" w:author="A" w:date="2026-01-01T00:00:00Z"/></w:rPr></w:pPr>${TEXT_BOX}<w:r><w:t>Host</w:t></w:r></w:p>${paragraph("10000004", "Next")}${paragraph("10000006", "Tail")}`;

  const open = async () => openLegacy(await docx(BODY));

  test("accepting joins the two, and the box follows", async () => {
    const reviewer = await open();
    reviewer.acceptAll();
    expect(texts(reviewer)).toEqual(["HostNext", "Boxed", "Tail"]);
    const reopened = await openLegacy(await reviewer.toBuffer());
    expect(texts(reopened)).toEqual(["HostNext", "Boxed", "Tail"]);
  });

  test("rejecting keeps the two apart", async () => {
    const reviewer = await open();
    reviewer.rejectAll();
    expect(texts(reviewer)).toEqual(["Host", "Boxed", "Next", "Tail"]);
  });
});
