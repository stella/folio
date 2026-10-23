import JSZip from "jszip";

import { buildPanelLayoutDocument } from "./panelLayoutDocument";
import { generateDocxFixture } from "./validatedDocxFixture";

export const COMMENT_SEARCH_REPLY_LENGTHS = [8, 40] as const;

export const REPLY_SEARCH_TARGET = "reply-tail-search-target";

export const buildCommentSearchDocument = async (replyLength: number) => {
  const zip = await JSZip.loadAsync(
    await buildPanelLayoutDocument("comment-and-changes"),
  );
  const longReply = `${"This reply explains the disclosure obligation in detail. ".repeat(replyLength)}${REPLY_SEARCH_TARGET}`;
  const comments = [
    { id: 1, paraId: "2AAA0001", text: "Review the disclosure obligation." },
    { id: 2, paraId: "2AAA0002", text: "An earlier reply in the same thread." },
    { id: 3, paraId: "2AAA0003", text: longReply },
  ];
  zip.file(
    "word/comments.xml",
    `<w:comments xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml">${comments.map(({ id, paraId, text }) => `<w:comment w:id="${id}" w:author="Reviewer" w:date="2026-01-01T00:00:00Z"><w:p w14:paraId="${paraId}"><w:r><w:t>${text}</w:t></w:r></w:p></w:comment>`).join("")}</w:comments>`,
  );
  zip.file(
    "word/commentsExtended.xml",
    '<w15:commentsEx xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml"><w15:commentEx w15:paraId="2AAA0001" w15:done="0"/><w15:commentEx w15:paraId="2AAA0002" w15:paraIdParent="2AAA0001" w15:done="0"/><w15:commentEx w15:paraId="2AAA0003" w15:paraIdParent="2AAA0001" w15:done="0"/></w15:commentsEx>',
  );
  const contentTypes = await zip.file("[Content_Types].xml")?.async("string");
  const relationships = await zip
    .file("word/_rels/document.xml.rels")
    ?.async("string");
  if (!contentTypes || !relationships)
    throw new TypeError("Panel fixture parts are missing");
  zip.file(
    "[Content_Types].xml",
    contentTypes.replace(
      "</Types>",
      '<Override PartName="/word/commentsExtended.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.commentsExtended+xml"/></Types>',
    ),
  );
  zip.file(
    "word/_rels/document.xml.rels",
    relationships.replace(
      "</Relationships>",
      '<Relationship Id="rId3" Type="http://schemas.microsoft.com/office/2011/relationships/commentsExtended" Target="commentsExtended.xml"/></Relationships>',
    ),
  );
  return generateDocxFixture(zip, "comment-search-long-reply");
};
