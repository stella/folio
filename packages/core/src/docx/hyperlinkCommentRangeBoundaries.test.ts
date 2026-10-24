import { expect, test } from "bun:test";

import { validateDocumentModel } from "@stll/docx-core";

import { expectCommentRangeBoundaryAttrs } from "../prosemirror/commentRangeBoundaryAttrs";
import { fromProseDoc } from "../prosemirror/conversion/fromProseDoc";
import { toProseDoc } from "../prosemirror/conversion/toProseDoc";
import type { Comment, Document, Paragraph } from "../types/document";
import { createEmptyDocument } from "../utils/createDocument";
import { completeCommentReferences } from "./commentReferenceCompletion";
import { parseParagraph } from "./paragraphParser";
import { serializeParagraph } from "./serializer/paragraphSerializer";
import { parseXmlDocument } from "./xmlParser";

const WORDPROCESSING_NAMESPACE =
  "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

const paragraphFromXml = (xml: string): Paragraph => {
  const root = parseXmlDocument(xml);
  if (!root) throw new Error("Failed to parse hyperlink comment fixture");
  return parseParagraph(root, null, null, null);
};

const documentWithParagraph = (paragraph: Paragraph): Document => {
  const template = createEmptyDocument();
  return {
    ...template,
    package: {
      ...template.package,
      document: {
        ...template.package.document,
        content: [paragraph],
      },
    },
  };
};

const roundTripParagraph = (paragraph: Paragraph): Paragraph => {
  const source = documentWithParagraph(paragraph);
  const saved = fromProseDoc(toProseDoc(source), source, {
    stylesheetSource: { type: "package" },
  });
  const result = saved.package.document.content.at(0);
  if (result?.type !== "paragraph") {
    throw new Error("The round trip lost the hyperlink paragraph");
  }
  return result;
};

const hyperlinksIn = (paragraph: Paragraph) => {
  const hyperlinks = paragraph.content.filter(
    (item) => item.type === "hyperlink",
  );
  if (hyperlinks.length !== 2) {
    throw new Error(`Expected two source hyperlinks, got ${hyperlinks.length}`);
  }
  return hyperlinks;
};

const commentIdsByText = (document: Document): Map<string, number[]> => {
  const idsByText = new Map<string, number[]>();
  toProseDoc(document).descendants((node) => {
    if (!node.isText || !node.text) return true;
    idsByText.set(
      node.text,
      node.marks
        .filter((mark) => mark.type.name === "comment")
        .flatMap((mark) => {
          const commentId = mark.attrs["commentId"];
          return typeof commentId === "number" ? [commentId] : [];
        }),
    );
    return true;
  });
  return idsByText;
};

const rangeAcrossSameHrefLinks = (commentId: number): string =>
  `<w:hyperlink w:anchor="target"><w:r><w:t>beforeStart</w:t></w:r>` +
  `<w:commentRangeStart w:id="${commentId}"/><w:r><w:t>insideFirst</w:t></w:r></w:hyperlink>` +
  `<w:hyperlink w:anchor="target"><w:r><w:t>insideSecond</w:t></w:r>` +
  `<w:commentRangeEnd w:id="${commentId}"/><w:r><w:t>afterEnd</w:t></w:r></w:hyperlink>` +
  `<w:hyperlink w:anchor="target"><w:r><w:t>sameHrefNeighbor</w:t></w:r></w:hyperlink>`;

test("round-trips nonempty comment endpoints authored inside a hyperlink", () => {
  const paragraph = paragraphFromXml(
    `<w:p xmlns:w="${WORDPROCESSING_NAMESPACE}"><w:hyperlink w:anchor="target">` +
      `<w:r><w:t>before</w:t></w:r><w:commentRangeStart w:id="41"/>` +
      `<w:r><w:t>covered</w:t></w:r><w:commentRangeEnd w:id="41"/>` +
      `<w:r><w:t>after</w:t></w:r></w:hyperlink>` +
      `<w:hyperlink w:anchor="target"><w:r><w:t>neighbor</w:t></w:r></w:hyperlink>` +
      `<w:r><w:commentReference w:id="41"/></w:r></w:p>`,
  );

  expect(
    hyperlinksIn(paragraph)
      .at(0)
      ?.children.map((child) => child.type),
  ).toEqual(["run", "commentRangeStart", "run", "commentRangeEnd", "run"]);

  const roundTripped = roundTripParagraph(paragraph);
  const [link, neighbor] = hyperlinksIn(roundTripped);
  const children = link.children;
  expect(children.map((child) => child.type)).toEqual([
    "run",
    "commentRangeStart",
    "run",
    "commentRangeEnd",
    "run",
  ]);
  expect(neighbor.children).toMatchObject([
    { type: "run", content: [{ type: "text", text: "neighbor" }] },
  ]);
  expect(serializeParagraph(roundTripped)).toContain(
    `<w:hyperlink w:anchor="target"><w:r><w:t>before</w:t></w:r>` +
      `<w:commentRangeStart w:id="41"/><w:r><w:t>covered</w:t></w:r>` +
      `<w:commentRangeEnd w:id="41"/><w:r><w:t>after</w:t></w:r></w:hyperlink>` +
      `<w:hyperlink w:anchor="target"><w:r><w:t>neighbor</w:t></w:r></w:hyperlink>`,
  );
});

test("round-trips an empty comment range authored inside a hyperlink", () => {
  const paragraph = paragraphFromXml(
    `<w:p xmlns:w="${WORDPROCESSING_NAMESPACE}"><w:hyperlink w:anchor="target">` +
      `<w:r><w:t>before</w:t></w:r><w:commentRangeStart w:id="42"/>` +
      `<w:commentRangeEnd w:id="42"/><w:r><w:t>after</w:t></w:r>` +
      `</w:hyperlink><w:hyperlink w:anchor="target">` +
      `<w:r><w:t>neighbor</w:t></w:r></w:hyperlink>` +
      `<w:r><w:commentReference w:id="42"/></w:r></w:p>`,
  );

  const source = documentWithParagraph(paragraph);
  const projected = toProseDoc(source);
  const boundaries: { type: string; id: number; underLink: boolean }[] = [];
  projected.descendants((node) => {
    if (node.type.name !== "commentRangeBoundary") return;
    const { marker } = expectCommentRangeBoundaryAttrs(node);
    boundaries.push({
      type: marker.type,
      id: marker.id,
      underLink: node.marks.some((mark) => mark.type.name === "hyperlink"),
    });
  });
  expect(boundaries).toEqual([
    { type: "commentRangeStart", id: 42, underLink: true },
    { type: "commentRangeEnd", id: 42, underLink: true },
  ]);

  const roundTripped = roundTripParagraph(paragraph);
  const [link, neighbor] = hyperlinksIn(roundTripped);
  expect(link.children.map((child) => child.type)).toEqual([
    "run",
    "commentRangeStart",
    "commentRangeEnd",
    "run",
  ]);
  expect(neighbor.children).toMatchObject([
    { type: "run", content: [{ type: "text", text: "neighbor" }] },
  ]);
});

test("retains run payload and every reference inside a hyperlink", () => {
  const paragraph = paragraphFromXml(
    `<w:p xmlns:w="${WORDPROCESSING_NAMESPACE}"><w:hyperlink w:anchor="target">` +
      `<w:r><w:t>linked text</w:t><w:commentReference w:id="44"/>` +
      `<w:commentReference w:id="45"/></w:r></w:hyperlink>` +
      `<w:hyperlink w:anchor="target"><w:r><w:t>neighbor</w:t></w:r></w:hyperlink></w:p>`,
  );
  expect(hyperlinksIn(paragraph).at(0)?.children).toMatchObject([
    { type: "run", content: [{ type: "text", text: "linked text" }] },
    { type: "commentReference", id: 44 },
    { type: "commentReference", id: 45 },
  ]);

  const roundTripped = roundTripParagraph(paragraph);
  const [link, neighbor] = hyperlinksIn(roundTripped);
  expect(link.children).toMatchObject([
    { type: "run", content: [{ type: "text", text: "linked text" }] },
    { type: "commentReference", id: 44 },
    { type: "commentReference", id: 45 },
  ]);
  expect(neighbor.children).toMatchObject([
    { type: "run", content: [{ type: "text", text: "neighbor" }] },
  ]);
});

test("keeps comment coverage local when ranges cross hyperlink boundaries", () => {
  const startsInside = paragraphFromXml(
    `<w:p xmlns:w="${WORDPROCESSING_NAMESPACE}"><w:hyperlink w:anchor="target">` +
      `<w:r><w:t>beforeStart</w:t></w:r><w:commentRangeStart w:id="51"/>` +
      `<w:r><w:t>insideLink</w:t></w:r></w:hyperlink>` +
      `<w:r><w:t>insideOutsideLink</w:t></w:r>` +
      `<w:commentRangeEnd w:id="51"/><w:r><w:t>afterEnd</w:t></w:r>` +
      `<w:hyperlink w:anchor="target"><w:r><w:t>sameHrefNeighbor</w:t></w:r></w:hyperlink></w:p>`,
  );
  expect(commentIdsByText(documentWithParagraph(startsInside))).toEqual(
    new Map([
      ["beforeStart", []],
      ["insideLink", [51]],
      ["insideOutsideLink", [51]],
      ["afterEnd", []],
      ["sameHrefNeighbor", []],
    ]),
  );

  const endsInside = paragraphFromXml(
    `<w:p xmlns:w="${WORDPROCESSING_NAMESPACE}"><w:commentRangeStart w:id="52"/>` +
      `<w:r><w:t>beforeLink</w:t></w:r><w:hyperlink w:anchor="target">` +
      `<w:r><w:t>insideLink</w:t></w:r><w:commentRangeEnd w:id="52"/>` +
      `<w:r><w:t>afterEndInLink</w:t></w:r></w:hyperlink>` +
      `<w:r><w:t>afterLink</w:t></w:r>` +
      `<w:hyperlink w:anchor="target"><w:r><w:t>sameHrefNeighbor</w:t></w:r></w:hyperlink></w:p>`,
  );
  expect(commentIdsByText(documentWithParagraph(endsInside))).toEqual(
    new Map([
      ["beforeLink", [52]],
      ["insideLink", [52]],
      ["afterEndInLink", []],
      ["afterLink", []],
      ["sameHrefNeighbor", []],
    ]),
  );
});

test("keeps child-specific hyperlink coverage in every container", () => {
  const contexts = [
    {
      name: "paragraph",
      content: rangeAcrossSameHrefLinks(61),
      commentId: 61,
    },
    {
      name: "simple field",
      content: `<w:fldSimple w:instr="PAGE">${rangeAcrossSameHrefLinks(62)}</w:fldSimple>`,
      commentId: 62,
    },
    {
      name: "inline SDT",
      content: `<w:sdt><w:sdtPr/><w:sdtContent>${rangeAcrossSameHrefLinks(63)}</w:sdtContent></w:sdt>`,
      commentId: 63,
    },
    {
      name: "tracked wrapper",
      content: `<w:ins w:id="64" w:author="Reviewer" w:date="2026-10-24T00:00:00Z">${rangeAcrossSameHrefLinks(64)}</w:ins>`,
      commentId: 64,
    },
  ];

  for (const context of contexts) {
    const paragraph = paragraphFromXml(
      `<w:p xmlns:w="${WORDPROCESSING_NAMESPACE}">${context.content}</w:p>`,
    );
    const commentIds = commentIdsByText(documentWithParagraph(paragraph));
    expect(commentIds, context.name).toEqual(
      new Map([
        ["beforeStart", []],
        ["insideFirst", [context.commentId]],
        ["insideSecond", [context.commentId]],
        ["afterEnd", []],
        ["sameHrefNeighbor", []],
      ]),
    );
  }
});

test("keeps comment endpoints in their tracked hyperlink source host", () => {
  const wrappers = {
    ins: "insertion",
    del: "deletion",
    moveFrom: "moveFrom",
    moveTo: "moveTo",
  } as const;
  for (const [wrapper, type] of Object.entries(wrappers)) {
    for (const referenceHost of ["hyperlink", "paragraph"] as const) {
      const reference = `<w:r><w:commentReference w:id="72"/></w:r>`;
      const textTag =
        type === "deletion" || type === "moveFrom" ? "delText" : "t";
      const paragraph = paragraphFromXml(
        `<w:p xmlns:w="${WORDPROCESSING_NAMESPACE}">` +
          `<w:${wrapper} w:id="71" w:author="Reviewer"><w:hyperlink w:anchor="target">` +
          `<w:r><w:${textTag}>before</w:${textTag}></w:r><w:commentRangeStart w:id="72"/>` +
          `<w:r><w:${textTag}>covered</w:${textTag}></w:r><w:commentRangeEnd w:id="72"/>` +
          (referenceHost === "hyperlink" ? reference : "") +
          `<w:r><w:${textTag}>after</w:${textTag}></w:r></w:hyperlink>` +
          `<w:hyperlink w:anchor="target"><w:r><w:${textTag}>neighbor</w:${textTag}></w:r></w:hyperlink></w:${wrapper}>` +
          (referenceHost === "paragraph" ? reference : "") +
          `</w:p>`,
      );
      const saved = roundTripParagraph(paragraph);
      const hyperlinks = saved.content.flatMap((item) =>
        item.type === type
          ? item.content.filter((child) => child.type === "hyperlink")
          : [],
      );
      expect(
        hyperlinks.map((link) => link.children.map((child) => child.type)),
      ).toEqual([
        [
          "run",
          "commentRangeStart",
          "run",
          "commentRangeEnd",
          ...(referenceHost === "hyperlink" ? ["commentReference"] : []),
          "run",
        ],
        ["run"],
      ]);
      expect(hyperlinks.at(1)?.children).toMatchObject([
        { type: "run", content: [{ type: "text", text: "neighbor" }] },
      ]);
    }
  }
});

test("completes a missing reference inside its hyperlink source host", () => {
  const paragraph = paragraphFromXml(
    `<w:p xmlns:w="${WORDPROCESSING_NAMESPACE}"><w:hyperlink w:anchor="target">` +
      `<w:r><w:t>before</w:t></w:r><w:commentRangeStart w:id="43"/>` +
      `<w:r><w:t>covered</w:t></w:r><w:commentRangeEnd w:id="43"/>` +
      `<w:r><w:t>after</w:t></w:r></w:hyperlink>` +
      `<w:hyperlink w:anchor="target"><w:r><w:t>neighbor</w:t></w:r></w:hyperlink></w:p>`,
  );
  const source = documentWithParagraph(paragraph);
  const comment: Comment = {
    id: 43,
    author: "Reviewer",
    content: [{ type: "paragraph", content: [{ type: "run", content: [] }] }],
  };
  source.package.document.comments = [comment];

  const completed = fromProseDoc(toProseDoc(source), source, {
    stylesheetSource: { type: "package" },
  });
  const completedParagraph = completed.package.document.content.at(0);
  if (completedParagraph?.type !== "paragraph") {
    throw new Error("The completed hyperlink paragraph is missing");
  }
  const [completedLink, completedNeighbor] = hyperlinksIn(completedParagraph);
  expect(completedLink.children.map((child) => child.type)).toEqual([
    "run",
    "commentRangeStart",
    "run",
    "commentRangeEnd",
    "commentReference",
    "run",
  ]);
  expect(completedNeighbor.children).toMatchObject([
    { type: "run", content: [{ type: "text", text: "neighbor" }] },
  ]);
  expect(validateDocumentModel(completed).valid).toBe(true);
  expect(completeCommentReferences(completed.package.document.content)).toBe(0);

  const saved = fromProseDoc(toProseDoc(completed), completed, {
    stylesheetSource: { type: "package" },
  });
  const savedParagraph = saved.package.document.content.at(0);
  if (savedParagraph?.type !== "paragraph") {
    throw new Error("The saved hyperlink paragraph is missing");
  }
  expect(
    hyperlinksIn(savedParagraph).map((link) =>
      link.children.map((child) => child.type),
    ),
  ).toEqual([
    completedLink.children.map((child) => child.type),
    completedNeighbor.children.map((child) => child.type),
  ]);
  expect(validateDocumentModel(saved).valid).toBe(true);
  const serialized = serializeParagraph(savedParagraph);
  expect(serialized).toContain(
    `<w:hyperlink w:anchor="target"><w:r><w:t>before</w:t></w:r>` +
      `<w:commentRangeStart w:id="43"/><w:r><w:t>covered</w:t></w:r>` +
      `<w:commentRangeEnd w:id="43"/><w:r><w:rPr><w:rStyle w:val="CommentReference"/></w:rPr>` +
      `<w:commentReference w:id="43"/></w:r><w:r><w:t>after</w:t></w:r></w:hyperlink>` +
      `<w:hyperlink w:anchor="target"><w:r><w:t>neighbor</w:t></w:r></w:hyperlink>`,
  );
  const reopenedRoot = parseXmlDocument(serialized);
  if (!reopenedRoot) throw new Error("Failed to reparse saved hyperlink XML");
  const reopenedParagraph = parseParagraph(reopenedRoot, null, null, null);
  expect(
    hyperlinksIn(reopenedParagraph).map((link) =>
      link.children.map((child) => child.type),
    ),
  ).toEqual([
    completedLink.children.map((child) => child.type),
    completedNeighbor.children.map((child) => child.type),
  ]);
  const reopened = documentWithParagraph(reopenedParagraph);
  reopened.package.document.comments = [comment];
  expect(validateDocumentModel(reopened).valid).toBe(true);
});
