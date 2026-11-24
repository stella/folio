/**
 * A story that ends in a table (the body, a header, a note) ends with a
 * paragraph in the canonical session's model, the one a reader shows after
 * the table: the session adds it, gives the story back as read while it is
 * unedited, and writes it once the story is edited. Removing every row of the
 * story's last table then never empties the story.
 */
import { expect, test } from "bun:test";

import { ZipArchive } from "@stll/docx-core/zip";

import { buildRawPackage, SECTION } from "../__tests__/documentShapes";
import { FolioDocxReviewer } from "../ai-edits/headless";

const W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const R_NS =
  "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const WML_CT = "application/vnd.openxmlformats-officedocument.wordprocessingml";
const XML_DECL = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';

const paragraph = (value: string) =>
  `<w:p><w:r><w:t>${value}</w:t></w:r></w:p>`;
const TABLE = `<w:tbl><w:tblGrid><w:gridCol w:w="4000"/></w:tblGrid><w:tr><w:tc>${paragraph("Cell")}</w:tc></w:tr></w:tbl>`;

const HEADER = "rIdHeader1";

const STORIES = {
  body: {
    part: "word/document.xml",
    build: () =>
      buildRawPackage({
        styles: null,
        body: paragraph("Opening") + TABLE + SECTION(),
      }),
  },
  header: {
    part: "word/header1.xml",
    build: () =>
      buildRawPackage({
        styles: null,
        body:
          paragraph("Body") +
          SECTION(`<w:headerReference w:type="default" r:id="${HEADER}"/>`),
        parts: [
          {
            path: "word/header1.xml",
            contentType: `${WML_CT}.header+xml`,
            relationshipType: "header",
            relationshipId: HEADER,
            body: `${XML_DECL}<w:hdr xmlns:w="${W_NS}" xmlns:r="${R_NS}">${paragraph("Opening")}${TABLE}</w:hdr>`,
          },
        ],
      }),
  },
} as const;

type Story = keyof typeof STORIES;

const partText = async (bytes: ArrayBuffer | Uint8Array, part: string) =>
  (await ZipArchive.read(bytes)).entry(part)?.text();

for (const story of Object.keys(STORIES) as Story[]) {
  const { part, build } = STORIES[story];

  test(`a ${story} that ends in a table reads a final paragraph and saves unedited as read`, async () => {
    const source = await build();
    const reviewer = await FolioDocxReviewer.fromBuffer(source.slice().buffer, {
      session: "canonical",
    });
    const handle =
      story === "body"
        ? ({ type: "main" } as const)
        : ({ type: "header", relationshipId: HEADER } as const);
    expect(
      reviewer.snapshotStory(handle)?.blocks.map(({ text }) => text),
    ).toEqual(["Opening", "Cell", ""]);
    expect(await partText(await reviewer.toBuffer(), part)).toBe(
      await partText(source, part),
    );
  });
}

for (const session of ["canonical"] as const)
  test(`removing every row of the body's last table keeps the body a paragraph (${session})`, async () => {
    const source = await STORIES.body.build();
    const reviewer = await FolioDocxReviewer.fromBuffer(source.slice().buffer, {
      session,
    });
    const cell = reviewer.getContent().find(({ text }) => text === "Cell");
    if (cell === undefined) throw new Error("expected the cell");
    const result = reviewer.applyOperations(
      [{ id: "row", type: "deleteTableRow", blockId: cell.id }],
      { undefinedReferences: "refuse" },
    );
    expect(result.skipped).toEqual([]);
    reviewer.acceptAll();
    const texts = reviewer.getContent().map(({ text }) => text);
    expect(texts.at(0)).toBe("Opening");
    expect(texts).not.toContain("Cell");
    // The saved package reopens with the body's paragraphs.
    const reopened = await FolioDocxReviewer.fromBuffer(
      await reviewer.toBuffer(),
      { session },
    );
    expect(reopened.getContent().at(0)?.text).toBe("Opening");
  });
