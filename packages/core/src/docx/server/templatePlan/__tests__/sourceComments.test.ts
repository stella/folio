import { beforeAll, describe, expect, test } from "bun:test";

import { buildTemplate, NS, p, r, readPart, XML } from "./fixtures";
import { applyFor, buildSourceWith, unwrap } from "./planning";

const W14 = "http://schemas.microsoft.com/office/word/2010/wordml";
const W15 = "http://schemas.microsoft.com/office/word/2012/wordml";
const W16CID = "http://schemas.microsoft.com/office/word/2016/wordml/cid";
const REL = "http://schemas.openxmlformats.org/package/2006/relationships";
const OFFICE_REL =
  "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

const comment = (id: number, paraId: string, text: string): string =>
  `<w:comment w:id="${id}" w:author="Reviewer" w:initials="R"><w:p w14:paraId="${paraId}"><w:pPr><w:pStyle w:val="CommentText"/></w:pPr><w:r><w:annotationRef/></w:r>${r(text)}</w:p></w:comment>`;

/** A source with three comments: one on a word of a numbered clause, one
 *  that is a point in the second clause, and one on an empty line the plan
 *  may omit; plus the parts that extend comments. */
const sourceWithComments = async (): Promise<Uint8Array> =>
  await buildSourceWith({
    body: [
      p(
        `${r("1.1 The Provider shall deliver the ")}<w:commentRangeStart w:id="0"/>${r("services")}<w:commentRangeEnd w:id="0"/><w:r><w:commentReference w:id="0"/></w:r>${r(" on time.")}`,
      ),
      p(
        `${r("1.2 Payment is due monthly.")}<w:r><w:commentReference w:id="1"/></w:r>`,
      ),
      // An empty line a plan may omit, holding only a comment.
      p(`<w:r><w:commentReference w:id="2"/></w:r>`),
    ].join(""),
    parts: {
      "word/_rels/document.xml.rels": `${XML}<Relationships xmlns="${REL}"><Relationship Id="rId40" Type="${OFFICE_REL}/comments" Target="comments.xml"/><Relationship Id="rId41" Type="http://schemas.microsoft.com/office/2011/relationships/commentsExtended" Target="commentsExtended.xml"/><Relationship Id="rId42" Type="http://schemas.microsoft.com/office/2016/09/relationships/commentsIds" Target="commentsIds.xml"/></Relationships>`,
      "word/comments.xml": `${XML}<w:comments ${NS} xmlns:w14="${W14}">${comment(0, "0A000001", "Which services?")}${comment(1, "0A000002", "Monthly or quarterly?")}${comment(2, "0A000003", "Delete before signing.")}</w:comments>`,
      "word/commentsExtended.xml": `${XML}<w15:commentsEx xmlns:w15="${W15}"><w15:commentEx w15:paraId="0A000001" w15:done="0"/><w15:commentEx w15:paraId="0A000002" w15:done="1"/><w15:commentEx w15:paraId="0A000003" w15:done="0"/></w15:commentsEx>`,
      "word/commentsIds.xml": `${XML}<w16cid:commentsIds xmlns:w16cid="${W16CID}"><w16cid:commentId w16cid:paraId="0A000001" w16cid:durableId="11111111"/><w16cid:commentId w16cid:paraId="0A000002" w16cid:durableId="22222222"/><w16cid:commentId w16cid:paraId="0A000003" w16cid:durableId="33333333"/></w16cid:commentsIds>`,
    },
  });

let template: Uint8Array;

beforeAll(async () => {
  template = await buildTemplate();
});

describe("the source's comments", () => {
  test("are left out unless the plan keeps them", async () => {
    const applied = unwrap(
      await applyFor({ source: await sourceWithComments(), template }),
    );
    expect(applied.report).not.toHaveProperty("sourceComments");
    expect(applied.report.dropped.comments).toBe(3);
    const document = (await readPart(applied.docx, "word/document.xml")) ?? "";
    expect(document).not.toContain("comment");
  });

  test("travel with their ranges and extensions; one in an omitted block is counted", async () => {
    const source = await sourceWithComments();
    const applied = unwrap(
      await applyFor({
        source,
        template,
        items: (blocks) =>
          blocks.slice(0, 2).map(({ id, hash }) => ({
            kind: "source" as const,
            blockId: id,
            hash,
          })),
        extra: {
          comments: "keep",
          omit: [{ blockId: "p2", reason: "empty" }],
        },
      }),
    ).report;
    expect(applied.sourceComments).toEqual({ kept: 2, leftOut: 1 });
    expect(applied.dropped.comments).toBe(1);
  });

  test("anchor on the same text in the output", async () => {
    const source = await sourceWithComments();
    const applied = unwrap(
      await applyFor({ source, template, extra: { comments: "keep" } }),
    );
    expect(applied.report.sourceComments).toEqual({ kept: 3, leftOut: 0 });
    const document = (await readPart(applied.docx, "word/document.xml")) ?? "";
    expect(document).toMatch(
      /<w:commentRangeStart w:id="0"\/><w:r>(?:<w:rPr>.*?<\/w:rPr>)?<w:t xml:space="preserve">services<\/w:t><\/w:r><w:commentRangeEnd w:id="0"\/><w:r>(?:<w:rPr>.*?<\/w:rPr>)?<w:commentReference w:id="0"\/><\/w:r>/u,
    );
    expect(document).toContain('<w:commentReference w:id="1"/>');
    const comments = (await readPart(applied.docx, "word/comments.xml")) ?? "";
    expect(comments).toContain("Which services?");
    expect(comments).toContain('w14:paraId="0A000002"');
    const extended =
      (await readPart(applied.docx, "word/commentsExtended.xml")) ?? "";
    expect(extended).toContain('w15:done="1"');
    const relationships =
      (await readPart(applied.docx, "word/_rels/document.xml.rels")) ?? "";
    expect(relationships).toContain("comments.xml");
    expect(relationships).toContain("commentsIds.xml");
    const types = (await readPart(applied.docx, "[Content_Types].xml")) ?? "";
    expect(types).toContain("wordprocessingml.comments+xml");
    expect(types).toContain("wordprocessingml.commentsExtended+xml");
  });
});
