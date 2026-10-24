import { describe, expect, test } from "bun:test";

import { fromProseDoc } from "../prosemirror/conversion/fromProseDoc";
import { toProseDoc } from "../prosemirror/conversion/toProseDoc";
import { schema } from "../prosemirror/schema";
import type { Document, Paragraph } from "../types/document";
import { createEmptyDocument } from "../utils/createDocument";
import { parseParagraph } from "./paragraphParser";
import { visitInlineContentSlotsInContent } from "./paragraphTraversal";
import { serializeParagraph } from "./serializer/paragraphSerializer";
import { parseXmlDocument } from "./xmlParser";

const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

const paragraphFromXml = (xml: string): Paragraph => {
  const root = parseXmlDocument(xml);
  if (!root) throw new Error("Failed to parse the paragraph fixture");
  return parseParagraph(root, null, null, null);
};

const roundTrip = (source: Document): Document =>
  fromProseDoc(toProseDoc(source), source, {
    stylesheetSource: { type: "package" },
  });

const documentWithCommentParagraph = (paragraph: Paragraph): Document => {
  const template = createEmptyDocument();
  return {
    ...template,
    package: {
      ...template.package,
      document: {
        ...template.package.document,
        content: [paragraph],
        comments: [
          {
            id: 9,
            author: "Reviewer",
            date: "2026-01-01T00:00:00Z",
            content: [],
          },
        ],
      },
    },
  };
};

const commentParagraph = (content: string): Paragraph =>
  paragraphFromXml(
    `<w:p xmlns:w="${W}">${content}` +
      `<w:r><w:commentReference w:id="9"/></w:r></w:p>`,
  );

describe("comment range boundaries in inline content controls", () => {
  test("parsing and saving keeps direct anchors inside sdtContent", () => {
    const paragraph = paragraphFromXml(
      `<w:p xmlns:w="${W}"><w:commentRangeStart w:id="9"/>` +
        `<w:sdt><w:sdtPr><w:tag w:val="clause"/></w:sdtPr>` +
        `<w:sdtContent><w:r><w:t>covered</w:t></w:r>` +
        `<w:commentRangeEnd w:id="9"/></w:sdtContent></w:sdt>` +
        `<w:r><w:commentReference w:id="9"/></w:r></w:p>`,
    );
    const control = paragraph.content.find((item) => item.type === "inlineSdt");

    expect(control).toMatchObject({
      type: "inlineSdt",
      content: [
        { type: "run", content: [{ type: "text", text: "covered" }] },
        { type: "commentRangeEnd", id: 9 },
      ],
    });

    const saved = serializeParagraph(paragraph);
    const sdtContentStart = saved.indexOf("<w:sdtContent>");
    const sdtContentEnd = saved.indexOf("</w:sdtContent>", sdtContentStart);
    expect(sdtContentStart).toBeGreaterThanOrEqual(0);
    expect(saved.slice(sdtContentStart, sdtContentEnd)).toContain(
      '<w:commentRangeEnd w:id="9"/>',
    );

    const reparsed = paragraphFromXml(
      saved.replace("<w:p>", `<w:p xmlns:w="${W}">`),
    );
    expect(reparsed.content).toEqual(paragraph.content);
  });

  test("comment references stay in their legal run inside the control", () => {
    const paragraph = paragraphFromXml(
      `<w:p xmlns:w="${W}"><w:sdt><w:sdtPr/><w:sdtContent>` +
        `<w:r><w:commentReference w:id="9"/></w:r>` +
        `</w:sdtContent></w:sdt></w:p>`,
    );
    const control = paragraph.content.find((item) => item.type === "inlineSdt");

    expect(control).toMatchObject({
      type: "inlineSdt",
      content: [{ type: "commentReference", id: 9 }],
    });

    const saved = serializeParagraph(paragraph);
    expect(saved).toContain(
      `<w:sdtContent><w:r><w:rPr><w:rStyle w:val="CommentReference"/></w:rPr>` +
        `<w:commentReference w:id="9"/></w:r></w:sdtContent>`,
    );

    const reparsed = paragraphFromXml(
      saved.replace("<w:p>", `<w:p xmlns:w="${W}">`),
    );
    expect(reparsed.content).toEqual(paragraph.content);

    const source = documentWithCommentParagraph(paragraph);
    expect(roundTrip(source).package.document.content).toEqual(
      source.package.document.content,
    );
  });

  test("tracked field content keeps comment boundaries inside nested controls", () => {
    const paragraph = paragraphFromXml(
      `<w:p xmlns:w="${W}"><w:fldSimple w:instr="PAGE">` +
        `<w:ins w:id="14" w:author="Reviewer">` +
        `<w:sdt><w:sdtPr/><w:sdtContent>` +
        `<w:commentRangeStart w:id="9"/><w:r><w:t>inside</w:t></w:r>` +
        `<w:commentRangeEnd w:id="9"/></w:sdtContent></w:sdt>` +
        `</w:ins></w:fldSimple><w:r><w:commentReference w:id="9"/></w:r></w:p>`,
    );
    let inlineSdtCount = 0;
    const boundaryIds: number[] = [];
    const textParts: string[] = [];
    visitInlineContentSlotsInContent(paragraph.content, ({ item }) => {
      if (item.type === "inlineSdt") inlineSdtCount += 1;
      if (
        item.type === "commentRangeStart" ||
        item.type === "commentRangeEnd"
      ) {
        boundaryIds.push(item.id);
      }
      if (item.type === "run") {
        for (const part of item.content) {
          if (part.type === "text") textParts.push(part.text);
        }
      }
    });
    expect(inlineSdtCount).toBe(1);
    expect(boundaryIds).toEqual([9, 9]);
    expect(textParts.join("")).toContain("inside");

    const source = documentWithCommentParagraph(paragraph);

    expect(roundTrip(source).package.document.content).toEqual(
      source.package.document.content,
    );
  });

  test("a nonempty range inside the control closes before outside text", () => {
    const paragraph = paragraphFromXml(
      `<w:p xmlns:w="${W}"><w:sdt><w:sdtPr>` +
        `<w:tag w:val="clause"/></w:sdtPr><w:sdtContent>` +
        `<w:commentRangeStart w:id="9"/><w:r><w:t>inside</w:t></w:r>` +
        `<w:commentRangeEnd w:id="9"/></w:sdtContent></w:sdt>` +
        `<w:r><w:t> outside</w:t></w:r>` +
        `<w:r><w:commentReference w:id="9"/></w:r></w:p>`,
    );
    const source = documentWithCommentParagraph(paragraph);

    const editor = toProseDoc(source);
    let insideHasComment = false;
    let outsideHasComment = false;
    editor.descendants((node) => {
      if (!node.isText) return true;
      if (node.text?.includes("inside")) {
        insideHasComment = node.marks.some(
          (mark) => mark.type.name === "comment",
        );
      }
      if (node.text?.includes("outside")) {
        outsideHasComment = node.marks.some(
          (mark) => mark.type.name === "comment",
        );
      }
      return true;
    });
    expect(insideHasComment).toBe(true);
    expect(outsideHasComment).toBe(false);
    expect(roundTrip(source).package.document.content).toEqual(
      source.package.document.content,
    );
  });

  test("a range entering and leaving the control stays outside at both ends", () => {
    const paragraph = commentParagraph(
      `<w:commentRangeStart w:id="9"/><w:r><w:t>before</w:t></w:r>` +
        `<w:sdt><w:sdtPr/><w:sdtContent><w:r><w:t>inside</w:t></w:r></w:sdtContent></w:sdt>` +
        `<w:r><w:t>after</w:t></w:r><w:commentRangeEnd w:id="9"/>`,
    );
    const source = documentWithCommentParagraph(paragraph);

    expect(roundTrip(source).package.document.content).toEqual(
      source.package.document.content,
    );
  });

  test("a range that starts in the control and continues outside keeps its start inside", () => {
    const paragraph = commentParagraph(
      `<w:sdt><w:sdtPr/><w:sdtContent><w:commentRangeStart w:id="9"/>` +
        `<w:r><w:t>inside</w:t></w:r></w:sdtContent></w:sdt>` +
        `<w:r><w:t>after</w:t></w:r><w:commentRangeEnd w:id="9"/>`,
    );
    const source = documentWithCommentParagraph(paragraph);

    expect(roundTrip(source).package.document.content).toEqual(
      source.package.document.content,
    );
  });

  test("a range that ends in the control keeps its end inside", () => {
    const paragraph = commentParagraph(
      `<w:commentRangeStart w:id="9"/><w:r><w:t>before</w:t></w:r>` +
        `<w:sdt><w:sdtPr/><w:sdtContent><w:r><w:t>inside</w:t></w:r>` +
        `<w:commentRangeEnd w:id="9"/></w:sdtContent></w:sdt>` +
        `<w:r><w:t>after</w:t></w:r>`,
    );
    const source = documentWithCommentParagraph(paragraph);

    expect(roundTrip(source).package.document.content).toEqual(
      source.package.document.content,
    );
  });

  test("nested controls retain a range wholly inside the inner control", () => {
    const paragraph = commentParagraph(
      `<w:sdt><w:sdtPr><w:tag w:val="outer"/></w:sdtPr><w:sdtContent>` +
        `<w:sdt><w:sdtPr><w:tag w:val="inner"/></w:sdtPr><w:sdtContent>` +
        `<w:commentRangeStart w:id="9"/><w:r><w:t>inside</w:t></w:r>` +
        `<w:commentRangeEnd w:id="9"/></w:sdtContent></w:sdt>` +
        `</w:sdtContent></w:sdt><w:r><w:t>after</w:t></w:r>`,
    );
    const source = documentWithCommentParagraph(paragraph);

    expect(roundTrip(source).package.document.content).toEqual(
      source.package.document.content,
    );
  });

  test.each([
    {
      name: "starts before the outer control and ends inside the inner control",
      content:
        `<w:commentRangeStart w:id="9"/><w:r><w:t>before</w:t></w:r>` +
        `<w:sdt><w:sdtPr/><w:sdtContent><w:sdt><w:sdtPr/><w:sdtContent>` +
        `<w:r><w:t>inside</w:t></w:r><w:commentRangeEnd w:id="9"/>` +
        `</w:sdtContent></w:sdt></w:sdtContent></w:sdt>`,
    },
    {
      name: "starts inside the inner control and ends after the outer control",
      content:
        `<w:sdt><w:sdtPr/><w:sdtContent><w:sdt><w:sdtPr/><w:sdtContent>` +
        `<w:commentRangeStart w:id="9"/><w:r><w:t>inside</w:t></w:r>` +
        `</w:sdtContent></w:sdt></w:sdtContent></w:sdt>` +
        `<w:r><w:t>after</w:t></w:r><w:commentRangeEnd w:id="9"/>`,
    },
    {
      name: "encloses both nested controls",
      content:
        `<w:commentRangeStart w:id="9"/><w:r><w:t>before</w:t></w:r>` +
        `<w:sdt><w:sdtPr/><w:sdtContent><w:sdt><w:sdtPr/><w:sdtContent>` +
        `<w:r><w:t>inside</w:t></w:r></w:sdtContent></w:sdt>` +
        `</w:sdtContent></w:sdt><w:r><w:t>after</w:t></w:r>` +
        `<w:commentRangeEnd w:id="9"/>`,
    },
  ])("nested controls preserve a range that $name", ({ content }) => {
    const source = documentWithCommentParagraph(commentParagraph(content));

    expect(roundTrip(source).package.document.content).toEqual(
      source.package.document.content,
    );
  });

  test("an unmarked reference inside a control does not split an outer range", () => {
    const source = documentWithCommentParagraph(
      commentParagraph(`<w:r><w:t>before</w:t></w:r>`),
    );
    const comment = schema.marks.comment.create({ commentId: 9 });
    const editor = schema.node("doc", null, [
      schema.node("paragraph", null, [
        schema.text("before", [comment]),
        schema.node("sdt", { sdtType: "richText", tag: "clause" }, [
          schema.node("commentReference", { commentId: 9 }),
        ]),
        schema.text("after", [comment]),
      ]),
    ]);

    const projected = fromProseDoc(editor, source, {
      stylesheetSource: { type: "package" },
    });
    const paragraph = projected.package.document.content.at(0);
    expect(paragraph?.type).toBe("paragraph");
    if (paragraph?.type !== "paragraph") return;

    const control = paragraph.content.find((item) => item.type === "inlineSdt");
    expect(control).toMatchObject({
      type: "inlineSdt",
      content: [{ type: "commentReference", id: 9 }],
    });
    expect(paragraph.content.map((item) => item.type)).toEqual([
      "commentRangeStart",
      "run",
      "inlineSdt",
      "run",
      "commentRangeEnd",
    ]);
  });
});
