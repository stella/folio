import { panic } from "better-result";
import { describe, expect, setDefaultTimeout, test } from "bun:test";
import fc from "fast-check";
import type { Node as PMNode } from "prosemirror-model";
import { EditorState } from "prosemirror-state";

import { applyDocumentOp, DOCUMENT_OP_TYPES } from "@stll/docx-core/ops";
import { ZipArchive } from "@stll/docx-core/zip";

import {
  assertProperty,
  propertyTestTimeout,
} from "../../../../test/property-testing";
import { resolveAllChangesInHeadlessState } from "../prosemirror/commands/comments";
import { expectCommentRangeBoundaryAttrs } from "../prosemirror/commentRangeBoundaryAttrs";
import { fromProseDoc } from "../prosemirror/conversion/fromProseDoc";
import { toProseDoc } from "../prosemirror/conversion/toProseDoc";
import type {
  Comment,
  Document,
  Paragraph,
  ParagraphContent,
} from "../types/document";
import { createEmptyDocument } from "../utils/createDocument";
import { parseDocx } from "./parser";
import { createDocx, repackDocx } from "./rezip";
import {
  getChildElements,
  getLocalName,
  getNamespaceUri,
  parseXmlDocument,
} from "./xmlParser";

setDefaultTimeout(propertyTestTimeout(30_000));

const COMMENT_ID = 9;
const TRANSITIONAL_WORDPROCESSINGML =
  "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const STRICT_WORDPROCESSINGML =
  "http://purl.oclc.org/ooxml/wordprocessingml/main";
const FOREIGN_NAMESPACE = "urn:folio:shadowed-prefix";
const WORDPROCESSINGML_NAMESPACES = new Set([
  TRANSITIONAL_WORDPROCESSINGML,
  STRICT_WORDPROCESSINGML,
]);
const XML_DECLARATION =
  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';

const NAMESPACE_PROFILES = [
  TRANSITIONAL_WORDPROCESSINGML,
  STRICT_WORDPROCESSINGML,
] as const;

const SLOT_SDT_DEPTHS = [0, 0, 1, 1, 2, 2, 1, 1, 0, 0] as const;
type RangeRoute = {
  startSlot: number;
  endSlot: number;
  startDepth: number;
  endDepth: number;
};

const crossingRoutes = (): RangeRoute[] => {
  const routes: RangeRoute[] = [];
  for (const startSlot of SLOT_SDT_DEPTHS.keys()) {
    for (
      let endSlot = startSlot + 1;
      endSlot < SLOT_SDT_DEPTHS.length;
      endSlot += 1
    ) {
      const startDepth = SLOT_SDT_DEPTHS.at(startSlot);
      const endDepth = SLOT_SDT_DEPTHS.at(endSlot);
      if (startDepth === undefined || endDepth === undefined) continue;
      const crossesControl =
        startDepth !== endDepth ||
        SLOT_SDT_DEPTHS.slice(startSlot, endSlot + 1).some(
          (depth) => depth !== startDepth,
        );
      if (!crossesControl) continue;
      routes.push({ startSlot, endSlot, startDepth, endDepth });
    }
  }
  return routes;
};

const CROSSING_ROUTES = crossingRoutes();

const TEXT_SEGMENTS = [
  { text: "before", startSlot: 0, endSlot: 1 },
  { text: "outer-before", startSlot: 2, endSlot: 3 },
  { text: "inner", startSlot: 4, endSlot: 5 },
  { text: "outer-after", startSlot: 6, endSlot: 7 },
  { text: "after", startSlot: 8, endSlot: 9 },
] as const;

type MarkerKind = "start" | "end";

const markerAt = (slot: number, route: RangeRoute, prefix: string): string => {
  const markers: string[] = [];
  if (route.startSlot === slot) {
    markers.push(`<${prefix}:commentRangeStart ${prefix}:id="${COMMENT_ID}"/>`);
  }
  if (route.endSlot === slot) {
    markers.push(`<${prefix}:commentRangeEnd ${prefix}:id="${COMMENT_ID}"/>`);
  }
  return markers.join("");
};

const xmlForRoute = (
  namespace: (typeof NAMESPACE_PROFILES)[number],
  route: RangeRoute,
  rebindAliases: boolean,
): string => {
  const slotPrefix = (slot: number): string =>
    SLOT_SDT_DEPTHS.at(slot) === 1 ? "q" : "w";
  const marker = (slot: number): string =>
    markerAt(slot, route, slotPrefix(slot));
  const outerElement = rebindAliases ? "x" : "w";
  const innerElement = rebindAliases ? "x" : "q";
  const outerBindings = rebindAliases
    ? ` xmlns:w="${FOREIGN_NAMESPACE}" xmlns:q="${namespace}"`
    : "";
  const innerBindings = rebindAliases
    ? ` xmlns:q="${FOREIGN_NAMESPACE}" xmlns:w="${namespace}"`
    : "";

  return (
    `${XML_DECLARATION}<w:document xmlns:w="${namespace}" ` +
    `xmlns:q="${rebindAliases ? FOREIGN_NAMESPACE : namespace}" ` +
    `xmlns:x="${namespace}"><w:body><w:p>` +
    `${marker(0)}<w:r><w:t>before</w:t></w:r>${marker(1)}` +
    `<${outerElement}:sdt${outerBindings}><q:sdtPr><q:tag q:val="outer"/>` +
    `</q:sdtPr><q:sdtContent>${marker(2)}` +
    `<q:r><q:t>outer-before</q:t></q:r>${marker(3)}` +
    `<${innerElement}:sdt${innerBindings}><w:sdtPr><w:tag w:val="inner"/>` +
    `</w:sdtPr><w:sdtContent>${marker(4)}` +
    `<w:r><w:t>inner</w:t></w:r>` +
    `<w:r><w:commentReference w:id="${COMMENT_ID}"/></w:r>${marker(5)}` +
    `</w:sdtContent></${innerElement}:sdt>${marker(6)}` +
    `<q:r><q:t>outer-after</q:t></q:r>${marker(7)}` +
    `</q:sdtContent></${outerElement}:sdt>${marker(8)}` +
    `<w:r><w:t>after</w:t></w:r>${marker(9)}` +
    `</w:p><w:sectPr/></w:body></w:document>`
  );
};

const sourceDocument = (
  additionalComments: readonly Comment[] = [],
): Document => {
  const empty = createEmptyDocument();
  return {
    ...empty,
    package: {
      ...empty.package,
      document: {
        ...empty.package.document,
        comments: [
          {
            id: COMMENT_ID,
            author: "Reviewer",
            date: "2026-01-01T00:00:00Z",
            content: [],
          },
          ...additionalComments,
        ],
      },
    },
  };
};

const sourcePackage = async (
  documentXml: string,
  additionalComments: readonly Comment[] = [],
): Promise<ArrayBuffer> => {
  const zip = await ZipArchive.read(
    await createDocx(sourceDocument(additionalComments)),
  );
  zip.set("word/document.xml", documentXml);
  return zip.generate({ compression: "store" }).then((bytes) => bytes.buffer);
};

const paragraphOf = (document: Document): Paragraph => {
  const paragraph = document.package.document.content.at(0);
  if (paragraph?.type !== "paragraph") {
    panic("Expected one paragraph in the generated document");
  }
  return paragraph;
};

const expectedCoveredText = (route: RangeRoute): string =>
  TEXT_SEGMENTS.filter(
    ({ startSlot, endSlot }) =>
      route.startSlot <= startSlot && route.endSlot >= endSlot,
  )
    .map(({ text }) => text)
    .join("");

const markedCommentText = (
  document: ReturnType<typeof toProseDoc>,
  commentId = COMMENT_ID,
): string => {
  let text = "";
  document.descendants((node) => {
    if (
      node.isText &&
      node.marks.some(
        (mark) =>
          mark.type.name === "comment" && mark.attrs["commentId"] === commentId,
      )
    ) {
      text += node.text ?? "";
    }
    return true;
  });
  return text;
};

const projectedRangeFacts = (
  document: ReturnType<typeof toProseDoc>,
): RangeFact[] => {
  const facts: RangeFact[] = [];
  const visit = (node: PMNode, sdtDepth: number): void => {
    if (node.type.name === "commentRangeBoundary") {
      const { marker } = expectCommentRangeBoundaryAttrs(node);
      facts.push({
        kind: marker.type === "commentRangeStart" ? "start" : "end",
        sdtDepth,
      });
    }
    const childSdtDepth = sdtDepth + (node.type.name === "sdt" ? 1 : 0);
    node.forEach((child) => visit(child, childSdtDepth));
  };
  visit(document, 0);
  return facts;
};

type RangeFact = { kind: MarkerKind; sdtDepth: number };

const rawRangeFacts = (xml: string): RangeFact[] => {
  const root = parseXmlDocument(xml);
  if (!root) panic("The generated XML document did not parse");
  const facts: RangeFact[] = [];
  const visit = (element: NonNullable<typeof root>, parentSdtDepth: number) => {
    const isWordElement = WORDPROCESSINGML_NAMESPACES.has(
      getNamespaceUri(element) ?? "",
    );
    const depth =
      parentSdtDepth +
      (isWordElement && getLocalName(element.name) === "sdt" ? 1 : 0);
    const localName = isWordElement ? getLocalName(element.name) : "";
    if (localName === "commentRangeStart" || localName === "commentRangeEnd") {
      facts.push({
        kind: localName === "commentRangeStart" ? "start" : "end",
        sdtDepth: parentSdtDepth,
      });
    }
    for (const child of getChildElements(element)) visit(child, depth);
  };
  visit(root, 0);
  return facts;
};

const modelRangeFacts = (content: readonly ParagraphContent[]): RangeFact[] => {
  const facts: RangeFact[] = [];
  const visit = (items: readonly ParagraphContent[], sdtDepth: number) => {
    for (const item of items) {
      if (
        item.type === "commentRangeStart" ||
        item.type === "commentRangeEnd"
      ) {
        facts.push({
          kind: item.type === "commentRangeStart" ? "start" : "end",
          sdtDepth,
        });
      }
      switch (item.type) {
        case "inlineSdt":
          visit(item.content, sdtDepth + 1);
          break;
        case "insertion":
        case "deletion":
        case "moveFrom":
        case "moveTo":
        case "inlineWrapper":
          visit(item.content, sdtDepth);
          break;
        case "hyperlink":
          visit(item.children, sdtDepth);
          break;
      }
    }
  };
  visit(content, 0);
  return facts;
};

const modelReferenceDepths = (
  content: readonly ParagraphContent[],
): number[] => {
  const depths: number[] = [];
  const visit = (items: readonly ParagraphContent[], sdtDepth: number) => {
    for (const item of items) {
      if (item.type === "commentReference") depths.push(sdtDepth);
      if (item.type === "inlineSdt") visit(item.content, sdtDepth + 1);
    }
  };
  visit(content, 0);
  return depths;
};

const xmlDocumentOf = async (buffer: ArrayBuffer): Promise<string> => {
  const zip = await ZipArchive.read(buffer);
  const documentXml = await zip.entry("word/document.xml")?.text();
  if (documentXml === undefined) {
    panic("The saved package lost document.xml");
  }
  return documentXml;
};

describe("balanced comment ranges crossing nested inline SDTs", () => {
  test("a tracked SDT keeps an interior source boundary after uncovered text", async () => {
    const documentXml =
      `${XML_DECLARATION}<w:document ` +
      `xmlns:w="${TRANSITIONAL_WORDPROCESSINGML}"><w:body><w:p>` +
      `<w:sdt><w:sdtPr><w:id w:val="42"/><w:richText/></w:sdtPr>` +
      `<w:sdtContent><w:ins w:id="7" w:author="Reviewer">` +
      `<w:r><w:t>before</w:t></w:r>` +
      `<w:commentRangeStart w:id="${COMMENT_ID}"/>` +
      `<w:r><w:t>covered</w:t></w:r>` +
      `</w:ins></w:sdtContent></w:sdt>` +
      `<w:commentRangeEnd w:id="${COMMENT_ID}"/>` +
      `<w:r><w:commentReference w:id="${COMMENT_ID}"/></w:r>` +
      `</w:p><w:sectPr/></w:body></w:document>`;
    const parsed = await parseDocx(await sourcePackage(documentXml), {
      detectVariables: false,
      preloadFonts: false,
    });
    const parsedContent = paragraphOf(parsed).content;
    const editor = toProseDoc(parsed);
    expect(markedCommentText(editor)).toBe("covered");
    expect(projectedRangeFacts(editor)).toEqual([
      { kind: "start", sdtDepth: 1 },
      { kind: "end", sdtDepth: 0 },
    ]);

    const projected = fromProseDoc(editor, parsed, {
      stylesheetSource: { type: "package" },
    });
    expect(paragraphOf(projected).content).toStrictEqual(parsedContent);

    const saved = await repackDocx(projected, { updateModifiedDate: false });
    expect(rawRangeFacts(await xmlDocumentOf(saved))).toEqual([
      { kind: "start", sdtDepth: 1 },
      { kind: "end", sdtDepth: 0 },
    ]);
    const reopened = await parseDocx(saved, {
      detectVariables: false,
      preloadFonts: false,
    });
    expect(modelRangeFacts(paragraphOf(reopened).content)).toEqual([
      { kind: "start", sdtDepth: 1 },
      { kind: "end", sdtDepth: 0 },
    ]);

    const rejected = resolveAllChangesInHeadlessState(
      EditorState.create({ doc: editor }),
      "reject",
    );
    expect(projectedRangeFacts(rejected.doc)).toEqual([
      { kind: "end", sdtDepth: 0 },
    ]);
  });

  test("paragraph endpoints retain their owner around an SDT inside a tracked wrapper", async () => {
    const documentXml =
      `${XML_DECLARATION}<w:document ` +
      `xmlns:w="${TRANSITIONAL_WORDPROCESSINGML}"><w:body><w:p>` +
      `<w:commentRangeStart w:id="${COMMENT_ID}"/>` +
      `<w:ins w:id="7" w:author="Reviewer"><w:sdt>` +
      `<w:sdtPr><w:id w:val="42"/></w:sdtPr><w:sdtContent>` +
      `<w:r><w:t>covered</w:t></w:r>` +
      `</w:sdtContent></w:sdt></w:ins>` +
      `<w:commentRangeEnd w:id="${COMMENT_ID}"/>` +
      `</w:p><w:sectPr/></w:body></w:document>`;
    const parsed = await parseDocx(await sourcePackage(documentXml), {
      detectVariables: false,
      preloadFonts: false,
    });

    const editor = toProseDoc(parsed);
    expect(projectedRangeFacts(editor)).toEqual([
      { kind: "start", sdtDepth: 0 },
      { kind: "end", sdtDepth: 0 },
    ]);
    expect(markedCommentText(editor)).toBe("covered");

    const projected = fromProseDoc(editor, parsed, {
      stylesheetSource: { type: "package" },
    });
    expect(modelRangeFacts(paragraphOf(projected).content)).toEqual([
      { kind: "start", sdtDepth: 0 },
      { kind: "end", sdtDepth: 0 },
    ]);
  });

  test("explicit endpoints keep their source host while text is inserted around an SDT", async () => {
    const route = {
      startSlot: 1,
      endSlot: 8,
      startDepth: 0,
      endDepth: 0,
    };
    const parsed = await parseDocx(
      await sourcePackage(
        xmlForRoute(TRANSITIONAL_WORDPROCESSINGML, route, false),
      ),
      { detectVariables: false, preloadFonts: false },
    );
    const original = toProseDoc(parsed);
    let outerSdtPosition: number | undefined;
    original.descendants((node, position) => {
      if (node.type.name === "sdt" && outerSdtPosition === undefined) {
        outerSdtPosition = position;
        return false;
      }
      return true;
    });
    if (outerSdtPosition === undefined) {
      panic("The projected document lost its outer inline SDT");
    }

    const beforeControl = EditorState.create({ doc: original });
    const editedBefore = beforeControl.apply(
      beforeControl.tr.insertText("before-control", outerSdtPosition),
    ).doc;
    expect(projectedRangeFacts(editedBefore)).toEqual([
      { kind: "start", sdtDepth: 0 },
      { kind: "end", sdtDepth: 0 },
    ]);

    const insideControl = EditorState.create({ doc: editedBefore });
    const paragraph = insideControl.doc.child(0);
    expect(paragraph.type.name).toBe("paragraph");
    let innerInsertPosition: number | undefined;
    paragraph.descendants((node, position) => {
      if (node.isText && node.text?.includes("outer-before")) {
        innerInsertPosition = position + 2;
        return false;
      }
      return true;
    });
    if (innerInsertPosition === undefined) {
      panic("The paragraph lost its outer SDT text");
    }
    const editedInside = insideControl.apply(
      insideControl.tr.insertText("inside-control", innerInsertPosition),
    ).doc;
    expect(markedCommentText(editedInside)).toContain("inside-control");
    expect(projectedRangeFacts(editedInside)).toEqual([
      { kind: "start", sdtDepth: 0 },
      { kind: "end", sdtDepth: 0 },
    ]);

    const projected = fromProseDoc(editedInside, parsed, {
      stylesheetSource: { type: "package" },
    });
    expect(modelRangeFacts(paragraphOf(projected).content)).toEqual([
      { kind: "start", sdtDepth: 0 },
      { kind: "end", sdtDepth: 0 },
    ]);
  });

  test("ordinary paragraphs continue to use comment marks without boundary points", async () => {
    const documentXml =
      `${XML_DECLARATION}<w:document ` +
      `xmlns:w="${TRANSITIONAL_WORDPROCESSINGML}"><w:body><w:p>` +
      `<w:commentRangeStart w:id="${COMMENT_ID}"/>` +
      `<w:r><w:t>covered</w:t></w:r>` +
      `<w:commentRangeEnd w:id="${COMMENT_ID}"/>` +
      `</w:p><w:sectPr/></w:body></w:document>`;
    const parsed = await parseDocx(await sourcePackage(documentXml), {
      detectVariables: false,
      preloadFonts: false,
    });
    const editor = toProseDoc(parsed);
    expect(projectedRangeFacts(editor)).toEqual([]);
    expect(markedCommentText(editor)).toBe("covered");
    const projected = fromProseDoc(editor, parsed, {
      stylesheetSource: { type: "package" },
    });
    expect(modelRangeFacts(paragraphOf(projected).content)).toEqual([
      { kind: "start", sdtDepth: 0 },
      { kind: "end", sdtDepth: 0 },
    ]);
  });

  test("deleting a thread removes its explicit SDT endpoints without changing another thread", async () => {
    const otherCommentId = 10;
    const documentXml =
      `${XML_DECLARATION}<w:document ` +
      `xmlns:w="${TRANSITIONAL_WORDPROCESSINGML}" ` +
      `xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml">` +
      `<w:body><w:p w14:paraId="00000001">` +
      `<w:commentRangeStart w:id="${COMMENT_ID}"/>` +
      `<w:r><w:t>before</w:t></w:r>` +
      `<w:sdt><w:sdtPr><w:id w:val="42"/></w:sdtPr><w:sdtContent>` +
      `<w:r><w:t>inside</w:t></w:r>` +
      `<w:commentRangeEnd w:id="${COMMENT_ID}"/>` +
      `</w:sdtContent></w:sdt>` +
      `<w:commentRangeStart w:id="${otherCommentId}"/>` +
      `<w:r><w:t>other</w:t></w:r>` +
      `<w:commentRangeEnd w:id="${otherCommentId}"/>` +
      `</w:p><w:sectPr/></w:body></w:document>`;
    const parsed = await parseDocx(
      await sourcePackage(documentXml, [
        {
          id: otherCommentId,
          author: "Reviewer",
          date: "2026-01-02T00:00:00Z",
          content: [],
        },
      ]),
      { detectVariables: false, preloadFonts: false },
    );
    expect(parsed.package.document.comments?.map(({ id }) => id)).toEqual([
      COMMENT_ID,
      otherCommentId,
    ]);
    const deleted = applyDocumentOp(parsed, {
      type: DOCUMENT_OP_TYPES.DELETE_COMMENT,
      id: COMMENT_ID,
      scope: "thread",
    }).unwrap().document;
    expect(deleted.package.document.comments?.map(({ id }) => id)).toEqual([
      otherCommentId,
    ]);
    expect(modelRangeFacts(paragraphOf(deleted).content)).toEqual([
      { kind: "start", sdtDepth: 0 },
      { kind: "end", sdtDepth: 0 },
    ]);

    const editor = toProseDoc(deleted);
    expect(projectedRangeFacts(editor)).toEqual([
      { kind: "start", sdtDepth: 0 },
      { kind: "end", sdtDepth: 0 },
    ]);
    expect(markedCommentText(editor, otherCommentId)).toBe("other");
    const saved = await repackDocx(
      fromProseDoc(editor, deleted, { stylesheetSource: { type: "package" } }),
      { updateModifiedDate: false },
    );
    const reopened = await parseDocx(saved, {
      detectVariables: false,
      preloadFonts: false,
    });
    expect(reopened.package.document.comments?.map(({ id }) => id)).toEqual([
      otherCommentId,
    ]);
    expect(modelRangeFacts(paragraphOf(reopened).content)).toEqual([
      { kind: "start", sdtDepth: 0 },
      { kind: "end", sdtDepth: 0 },
    ]);
  });

  test("range owners and covered text survive namespace and editor round trips", async () => {
    await assertProperty(
      fc.asyncProperty(
        fc.constantFrom(...NAMESPACE_PROFILES),
        fc.boolean(),
        fc.constantFrom(...CROSSING_ROUTES),
        async (namespace, rebindAliases, route) => {
          const documentXml = xmlForRoute(namespace, route, rebindAliases);
          expect(rawRangeFacts(documentXml)).toEqual([
            { kind: "start", sdtDepth: route.startDepth },
            { kind: "end", sdtDepth: route.endDepth },
          ]);

          const parsed = await parseDocx(await sourcePackage(documentXml), {
            detectVariables: false,
            preloadFonts: false,
          });
          const parsedParagraph = paragraphOf(parsed);
          expect(modelRangeFacts(parsedParagraph.content)).toEqual([
            { kind: "start", sdtDepth: route.startDepth },
            { kind: "end", sdtDepth: route.endDepth },
          ]);
          expect(modelReferenceDepths(parsedParagraph.content)).toEqual([2]);

          const editor = toProseDoc(parsed);
          expect(markedCommentText(editor)).toBe(expectedCoveredText(route));
          expect(projectedRangeFacts(editor)).toEqual([
            { kind: "start", sdtDepth: route.startDepth },
            { kind: "end", sdtDepth: route.endDepth },
          ]);

          const projected = fromProseDoc(editor, parsed, {
            stylesheetSource: { type: "package" },
          });
          expect(paragraphOf(projected).content).toStrictEqual(
            parsedParagraph.content,
          );
          expect(modelRangeFacts(paragraphOf(projected).content)).toEqual([
            { kind: "start", sdtDepth: route.startDepth },
            { kind: "end", sdtDepth: route.endDepth },
          ]);
          expect(modelReferenceDepths(paragraphOf(projected).content)).toEqual([
            2,
          ]);

          const firstSave = await repackDocx(projected, {
            updateModifiedDate: false,
          });
          const reopened = await parseDocx(firstSave, {
            detectVariables: false,
            preloadFonts: false,
          });
          const reopenedEditor = toProseDoc(reopened);
          expect(markedCommentText(reopenedEditor)).toBe(
            expectedCoveredText(route),
          );
          expect(modelRangeFacts(paragraphOf(reopened).content)).toEqual([
            { kind: "start", sdtDepth: route.startDepth },
            { kind: "end", sdtDepth: route.endDepth },
          ]);
          expect(modelReferenceDepths(paragraphOf(reopened).content)).toEqual([
            2,
          ]);

          const reopenedProjection = fromProseDoc(reopenedEditor, reopened, {
            stylesheetSource: { type: "package" },
          });
          const secondSave = await repackDocx(reopenedProjection, {
            updateModifiedDate: false,
          });
          expect(await xmlDocumentOf(secondSave)).toBe(
            await xmlDocumentOf(firstSave),
          );
        },
      ),
      {
        id: "balanced-comment-ranges-cross-nested-inline-sdts",
        numRuns: 48,
        examples: [
          [
            TRANSITIONAL_WORDPROCESSINGML,
            false,
            { startSlot: 5, endSlot: 8, startDepth: 2, endDepth: 0 },
          ],
          [
            TRANSITIONAL_WORDPROCESSINGML,
            true,
            { startSlot: 5, endSlot: 8, startDepth: 2, endDepth: 0 },
          ],
          [
            STRICT_WORDPROCESSINGML,
            false,
            { startSlot: 5, endSlot: 8, startDepth: 2, endDepth: 0 },
          ],
          [
            STRICT_WORDPROCESSINGML,
            true,
            { startSlot: 5, endSlot: 8, startDepth: 2, endDepth: 0 },
          ],
        ],
      },
    );
  });
});
