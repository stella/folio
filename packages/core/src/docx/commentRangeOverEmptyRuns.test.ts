/**
 * A comment whose range start sits alone in an inline content control, with
 * only empty runs between it and the range end and reference at paragraph
 * level. The shape comes from a real document; the package is generated here
 * so the input stays legible in the diff.
 */

import { describe, expect, test } from "bun:test";

import { validateDocumentModel } from "@stll/docx-core";
import {
  normalizeForOps,
  survivingComments,
  withBodyContent,
} from "@stll/docx-core/ops";
import { ZipArchive } from "@stll/docx-core/zip";

import type {
  BlockContent,
  Document,
  ParagraphContent,
} from "../types/document";
import { withoutOrphanCommentRanges } from "./commentRangeIntegrity";
import { visitDocxContentSlots } from "./paragraphTraversal";
import { parseDocx } from "./parser";
import { repackDocx } from "./rezip";

const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const RELATIONSHIPS =
  "http://schemas.openxmlformats.org/package/2006/relationships";
const OFFICE_RELATIONSHIPS =
  "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const WORDPROCESSING =
  "application/vnd.openxmlformats-officedocument.wordprocessingml";

const CELL_PARAGRAPH =
  `<w:p><w:sdt><w:sdtPr><w:tag w:val="clause"/></w:sdtPr>` +
  `<w:sdtContent><w:commentRangeStart w:id="1"/></w:sdtContent></w:sdt>` +
  `<w:r><w:t></w:t></w:r><w:r><w:t></w:t></w:r>` +
  `<w:commentRangeEnd w:id="1"/>` +
  `<w:r><w:rPr><w:rStyle w:val="CommentReference"/></w:rPr>` +
  `<w:commentReference w:id="1"/></w:r></w:p>`;

const commentOverEmptyRunsDocx = async (): Promise<ArrayBuffer> => {
  const parts: Record<string, string> = {
    "[Content_Types].xml":
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
      `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">` +
      `<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>` +
      `<Default Extension="xml" ContentType="application/xml"/>` +
      `<Override PartName="/word/document.xml" ContentType="${WORDPROCESSING}.document.main+xml"/>` +
      `<Override PartName="/word/styles.xml" ContentType="${WORDPROCESSING}.styles+xml"/>` +
      `<Override PartName="/word/comments.xml" ContentType="${WORDPROCESSING}.comments+xml"/>` +
      `</Types>`,
    "_rels/.rels":
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
      `<Relationships xmlns="${RELATIONSHIPS}">` +
      `<Relationship Id="rId1" Type="${OFFICE_RELATIONSHIPS}/officeDocument" Target="word/document.xml"/>` +
      `</Relationships>`,
    "word/_rels/document.xml.rels":
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
      `<Relationships xmlns="${RELATIONSHIPS}">` +
      `<Relationship Id="rId1" Type="${OFFICE_RELATIONSHIPS}/styles" Target="styles.xml"/>` +
      `<Relationship Id="rId2" Type="${OFFICE_RELATIONSHIPS}/comments" Target="comments.xml"/>` +
      `</Relationships>`,
    "word/styles.xml":
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
      `<w:styles xmlns:w="${W}">` +
      `<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>` +
      `</w:styles>`,
    "word/comments.xml":
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
      `<w:comments xmlns:w="${W}">` +
      `<w:comment w:id="1" w:author="Reviewer" w:date="2026-01-01T00:00:00Z">` +
      `<w:p><w:r><w:t>Check this clause.</w:t></w:r></w:p></w:comment>` +
      `</w:comments>`,
    "word/document.xml":
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
      `<w:document xmlns:w="${W}"><w:body>` +
      `<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/></w:tblPr>` +
      `<w:tblGrid><w:gridCol w:w="9360"/></w:tblGrid>` +
      `<w:tr><w:tc><w:tcPr><w:tcW w:w="9360" w:type="dxa"/></w:tcPr>` +
      `${CELL_PARAGRAPH}</w:tc></w:tr></w:tbl>` +
      `<w:p><w:r><w:t>After the table.</w:t></w:r></w:p>` +
      `<w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr>` +
      `</w:body></w:document>`,
  };

  const zip = new ZipArchive();
  for (const name of Object.keys(parts).toSorted()) {
    zip.set(name, parts[name] ?? "");
  }
  return (await zip.generate({ compression: "store" })).buffer;
};

type MarkerFact = { type: string; id: number; inControl: boolean };

/** Each comment marker of the body in order, and whether a content control holds it. */
const commentMarkers = (document: Document): MarkerFact[] => {
  const facts: MarkerFact[] = [];
  const controlArrays = new Set<ParagraphContent[]>();
  visitDocxContentSlots({ documentBody: document.package.document }, (slot) => {
    if (slot.type !== "inline") return;
    if (slot.item.type === "inlineSdt") controlArrays.add(slot.item.content);
    if (
      slot.item.type === "commentRangeStart" ||
      slot.item.type === "commentRangeEnd" ||
      slot.item.type === "commentReference"
    ) {
      facts.push({
        type: slot.item.type,
        id: slot.item.id,
        inControl: controlArrays.has(slot.content),
      });
    }
  });
  return facts;
};

const commentIssues = (document: Document): string[] =>
  validateDocumentModel(document)
    .issues.map(({ message }) => message)
    .filter((message) => /comment/iu.test(message));

const SOURCE_MARKERS: MarkerFact[] = [
  { type: "commentRangeStart", id: 1, inControl: true },
  { type: "commentRangeEnd", id: 1, inControl: false },
  { type: "commentReference", id: 1, inControl: false },
];

const withoutComments = (document: Document): Document => ({
  ...document,
  package: {
    ...document.package,
    document: { ...document.package.document, comments: [] },
  },
});

/**
 * The removal this save used to make: comment markers dropped from each
 * paragraph's own content only, never from inside an inline container.
 */
const withoutTopLevelCommentMarkers = (document: Document): Document => {
  const blocks = (content: BlockContent[]): BlockContent[] =>
    content.map((block) => {
      if (block.type === "table") {
        return {
          ...block,
          rows: block.rows.map((row) => ({
            ...row,
            cells: row.cells.map((cell) => ({
              ...cell,
              content: blocks(cell.content),
            })),
          })),
        };
      }
      if (block.type !== "paragraph") return block;
      return {
        ...block,
        content: block.content.filter(
          (item) =>
            item.type !== "commentRangeStart" &&
            item.type !== "commentRangeEnd" &&
            item.type !== "commentReference",
        ),
      };
    });
  return {
    ...document,
    package: {
      ...document.package,
      document: withBodyContent(
        document.package.document,
        blocks(document.package.document.content),
      ),
    },
  };
};

describe("a comment anchored over empty runs", () => {
  test("survives normalizing for operations, and saves and reopens whole", async () => {
    const source = await parseDocx(await commentOverEmptyRunsDocx(), {
      preloadFonts: false,
    });
    expect(commentMarkers(source)).toEqual(SOURCE_MARKERS);

    const normalized = normalizeForOps(source);
    expect(survivingComments(normalized).map(({ id }) => id)).toEqual([1]);
    expect(commentIssues(normalized)).toEqual([]);

    const reopened = await parseDocx(await repackDocx(normalized), {
      preloadFonts: false,
    });
    expect(
      (reopened.package.document.comments ?? []).map(({ id }) => id),
    ).toEqual([1]);
    expect(commentMarkers(reopened)).toEqual(SOURCE_MARKERS);
    expect(commentIssues(reopened)).toEqual([]);
  });

  test("once retired, loses its start inside the content control too", async () => {
    const source = await parseDocx(await commentOverEmptyRunsDocx(), {
      preloadFonts: false,
    });
    const retired = withoutComments(source);

    const cleaned = withoutOrphanCommentRanges(retired);
    expect(commentMarkers(cleaned)).toEqual([]);
    expect(commentIssues(cleaned)).toEqual([]);

    // Break proof: removal from paragraph content alone leaves the start the
    // content control holds, and the model validation refuses the save.
    const topLevelOnly = withoutTopLevelCommentMarkers(retired);
    expect(commentMarkers(topLevelOnly)).toEqual([SOURCE_MARKERS[0]]);
    expect(commentIssues(topLevelOnly)).toContain(
      "Comment 1 is referenced but not present in comments.xml.",
    );
  });
});
