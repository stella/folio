/**
 * A comment body keeps what `Comment.content` cannot hold.
 *
 * The survival census cannot see this: `w:comment` is reachable only under
 * `w:comments`, a part root its fixture builder does not synthesise, so all 35
 * of the container's pairs are counted unrepresentable rather than run. The
 * loss is real all the same — the schema lets a comment body hold everything a
 * document body can, and folio modelled only `w:p` — so the guard is here.
 */

import { describe, expect, test } from "bun:test";
import JSZip from "jszip";

import { parseDocx } from "./parser";
import { createEmptyDocx, repackDocx } from "./rezip";

const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const XML_DECLARATION = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';
const COMMENTS_CONTENT_TYPE =
  "application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml";
const COMMENTS_RELATIONSHIP =
  "http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments";

const buildDocx = async (commentBody: string): Promise<ArrayBuffer> => {
  const zip = await JSZip.loadAsync(await createEmptyDocx());
  zip.file(
    "word/document.xml",
    `${XML_DECLARATION}<w:document xmlns:w="${W}"><w:body><w:p>` +
      '<w:commentRangeStart w:id="1"/><w:r><w:t>anchor</w:t></w:r>' +
      '<w:commentRangeEnd w:id="1"/><w:r><w:commentReference w:id="1"/></w:r>' +
      "</w:p><w:sectPr/></w:body></w:document>",
  );
  zip.file(
    "word/comments.xml",
    `${XML_DECLARATION}<w:comments xmlns:w="${W}">` +
      `<w:comment w:id="1" w:author="Reviewer" w:date="2024-01-01T00:00:00Z">${commentBody}</w:comment>` +
      "</w:comments>",
  );

  const types = await zip.file("[Content_Types].xml")?.async("text");
  const rels = await zip.file("word/_rels/document.xml.rels")?.async("text");
  if (types === undefined || rels === undefined) {
    throw new Error("the empty package lost its packaging parts");
  }
  zip.file(
    "[Content_Types].xml",
    types.replace(
      "</Types>",
      `<Override PartName="/word/comments.xml" ContentType="${COMMENTS_CONTENT_TYPE}"/></Types>`,
    ),
  );
  zip.file(
    "word/_rels/document.xml.rels",
    rels.replace(
      "</Relationships>",
      `<Relationship Id="rIdComments" Type="${COMMENTS_RELATIONSHIP}" Target="comments.xml"/></Relationships>`,
    ),
  );
  return zip.generateAsync({ type: "arraybuffer" });
};

const commentsPart = async (buffer: ArrayBuffer): Promise<string> =>
  (await (await JSZip.loadAsync(buffer)).file("word/comments.xml")?.async("text")) ?? "";

const save = async (commentBody: string): Promise<string> =>
  commentsPart(
    await repackDocx(await parseDocx(await buildDocx(commentBody), { preloadFonts: false }), {
      updateModifiedDate: false,
    }),
  );

describe("a comment body survives a rebuild", () => {
  test("a table in a comment is kept, between the paragraphs it sat between", async () => {
    const saved = await save(
      "<w:p><w:r><w:t>before</w:t></w:r></w:p>" +
        "<w:tbl><w:tr><w:tc><w:p><w:r><w:t>cell</w:t></w:r></w:p></w:tc></w:tr></w:tbl>" +
        "<w:p><w:r><w:t>after</w:t></w:r></w:p>",
    );
    expect(saved).toContain("<w:tbl>");
    expect(saved.indexOf("before")).toBeLessThan(saved.indexOf("<w:tbl>"));
    expect(saved.indexOf("<w:tbl>")).toBeLessThan(saved.indexOf("after"));
  });

  test("a bookmark and a range marker in a comment body are kept", async () => {
    const saved = await save(
      '<w:bookmarkStart w:id="9" w:name="inComment"/>' +
        "<w:p><w:r><w:t>text</w:t></w:r></w:p>" +
        '<w:bookmarkEnd w:id="9"/>',
    );
    expect(saved).toContain('<w:bookmarkStart w:id="9" w:name="inComment"/>');
    expect(saved).toContain('<w:bookmarkEnd w:id="9"/>');
  });

  test("a content control in a comment body is kept", async () => {
    const saved = await save(
      '<w:sdt><w:sdtPr><w:alias w:val="pick"/></w:sdtPr>' +
        "<w:sdtContent><w:p><w:r><w:t>controlled</w:t></w:r></w:p></w:sdtContent></w:sdt>",
    );
    expect(saved).toContain("<w:sdt>");
    expect(saved).toContain('<w:alias w:val="pick"/>');
  });

  test("a comment holding only unmodelled markup keeps it, and does not gain a paragraph", async () => {
    const saved = await save(
      "<w:tbl><w:tr><w:tc><w:p><w:r><w:t>only</w:t></w:r></w:p></w:tc></w:tr></w:tbl>",
    );
    expect(saved).toContain("<w:tbl>");
    expect(saved).not.toContain("<w:annotationRef/>");
  });
});
