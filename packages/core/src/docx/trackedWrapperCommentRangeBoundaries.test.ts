/**
 * A comment-range boundary inside a run-level tracked change stays inside it.
 *
 * `CT_RunTrackChange` admits comment range markers through
 * `EG_RunLevelElts > EG_RangeMarkupElements`. The marker is a position: lifting
 * it beside the revision changes which side of the tracked text it occupies
 * after the revision is resolved.
 */

import { describe, expect, test } from "bun:test";
import { applyDocumentOps, OP_STORIES, REVISION_DECISIONS } from "@stll/docx-core/ops";
import fc from "fast-check";
import { EditorState } from "prosemirror-state";

import {
  assertProperty,
  propertyConfig,
  propertyTestTimeout,
} from "../../../../test/property-testing";

import {
  acceptAllChanges,
  rejectAllChanges,
  resolveAllChangesInHeadlessState,
} from "../prosemirror/commands/comments";
import { fromProseDoc } from "../prosemirror/conversion/fromProseDoc";
import { toProseDoc } from "../prosemirror/conversion/toProseDoc";
import { schema } from "../prosemirror/schema";
import type { Document, Paragraph } from "../types/document";

import { parseParagraph } from "./paragraphParser";
import { serializeParagraph } from "./serializer/paragraphSerializer";
import { parseXmlDocument } from "./xmlParser";

const WORDPROCESSINGML_NAMESPACE = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

const TRACKED_WRAPPERS = [
  { tag: "ins", type: "insertion" },
  { tag: "del", type: "deletion" },
  { tag: "moveFrom", type: "moveFrom" },
  { tag: "moveTo", type: "moveTo" },
] as const;

const COMMENT_MARKERS = ["commentRangeStart", "commentRangeEnd"] as const;

const parseParagraphXml = (xml: string): Paragraph => {
  const element = parseXmlDocument(xml);
  if (!element) throw new Error("Failed to parse the tracked comment boundary fixture");
  return parseParagraph(element, null, null, null, null, null);
};

const REVISION_ATTRIBUTES = 'w:id="7" w:author="Reviewer"';
const COMMENT_REFERENCE =
  '<w:r><w:rPr><w:rStyle w:val="CommentReference"/></w:rPr><w:commentReference w:id="5"/></w:r>';

const serializeInnerXml = (inner: string): string =>
  serializeParagraph(
    parseParagraphXml(`<w:p xmlns:w="${WORDPROCESSINGML_NAMESPACE}">${inner}</w:p>`),
  ).replace(/^<w:p>|<\/w:p>$/g, "");

const wrapParagraph = (paragraph: Paragraph): Document => ({
  package: {
    document: {
      comments: [{ id: 5, author: "Reviewer", content: [] }],
      content: [{ ...paragraph, paraId: "A1B2C3D4" }],
    },
  },
});

const markerTypes = (document: Document): string[] => {
  const paragraph = document.package.document.content.at(0);
  if (paragraph?.type !== "paragraph") throw new Error("Expected a paragraph");
  return paragraph.content.flatMap((item) =>
    item.type === "commentRangeStart" ||
    item.type === "commentRangeEnd" ||
    item.type === "commentReference"
      ? [item.type]
      : [],
  );
};

const textLength = (content: Paragraph["content"]): number => {
  let length = 0;
  for (const item of content) {
    if (item.type === "run") {
      for (const child of item.content) {
        if (child.type === "text") length += child.text.length;
      }
      continue;
    }
    if (
      item.type === "insertion" ||
      item.type === "deletion" ||
      item.type === "moveFrom" ||
      item.type === "moveTo"
    ) {
      length += textLength(item.content);
    }
  }
  return length;
};

const commentStartOffset = (content: Paragraph["content"], id: number): number | undefined => {
  let offset = 0;
  for (const item of content) {
    if (item.type === "commentRangeStart" && item.id === id) return offset;
    if (item.type === "run") {
      offset += textLength([item]);
      continue;
    }
    if (
      item.type === "insertion" ||
      item.type === "deletion" ||
      item.type === "moveFrom" ||
      item.type === "moveTo"
    ) {
      const nested = commentStartOffset(item.content, id);
      if (nested !== undefined) return offset + nested;
      offset += textLength(item.content);
    }
  }
  return undefined;
};

const markedText = (document: ReturnType<typeof toProseDoc>, commentId: number): string => {
  let text = "";
  document.descendants((node) => {
    if (
      node.isText &&
      node.marks.some(
        (mark) => mark.type.name === "comment" && mark.attrs["commentId"] === commentId,
      )
    ) {
      text += node.text ?? "";
    }
  });
  return text;
};

describe("comment-range boundaries inside tracked changes", () => {
  for (const { name, xml } of [
    {
      name: "w:commentRangeStart inside w:ins and the closing range markup outside it",
      xml:
        `<w:ins ${REVISION_ATTRIBUTES}><w:r><w:t>X</w:t></w:r>` +
        '<w:commentRangeStart w:id="5"/><w:r><w:t>Y</w:t></w:r></w:ins>' +
        `<w:commentRangeEnd w:id="5"/>${COMMENT_REFERENCE}`,
    },
    {
      name: "a following w:ins after the closing range markup",
      xml:
        `<w:ins ${REVISION_ATTRIBUTES}><w:r><w:t>X</w:t></w:r>` +
        '<w:commentRangeStart w:id="5"/><w:r><w:t>Y</w:t></w:r></w:ins>' +
        `<w:commentRangeEnd w:id="5"/>${COMMENT_REFERENCE}` +
        `<w:ins ${REVISION_ATTRIBUTES}><w:r><w:t>W</w:t></w:r></w:ins>`,
    },
    {
      name: "w:ins before w:commentRangeStart",
      xml:
        `<w:ins ${REVISION_ATTRIBUTES}><w:r><w:t>Z</w:t></w:r></w:ins>` +
        '<w:commentRangeStart w:id="5"/><w:r><w:t>beta</w:t></w:r>' +
        `<w:commentRangeEnd w:id="5"/>${COMMENT_REFERENCE}`,
    },
    {
      name: "range markup and w:commentReference in one w:del",
      xml:
        `<w:del ${REVISION_ATTRIBUTES}><w:r><w:delText>b</w:delText></w:r>` +
        '<w:commentRangeStart w:id="5"/><w:r><w:delText>et</w:delText></w:r>' +
        `<w:commentRangeEnd w:id="5"/>${COMMENT_REFERENCE}` +
        "<w:r><w:delText>a</w:delText></w:r></w:del>",
    },
    {
      name: "closing range markup between two w:del elements",
      xml:
        `<w:del ${REVISION_ATTRIBUTES}><w:r><w:delText>b</w:delText></w:r>` +
        '<w:commentRangeStart w:id="5"/><w:r><w:delText>et</w:delText></w:r></w:del>' +
        `<w:commentRangeEnd w:id="5"/>${COMMENT_REFERENCE}` +
        `<w:del ${REVISION_ATTRIBUTES}><w:r><w:delText>a</w:delText></w:r></w:del>`,
    },
  ]) {
    test(`${name} survives parse, projection, and serialization`, () => {
      const parsed = parseParagraphXml(`<w:p xmlns:w="${WORDPROCESSINGML_NAMESPACE}">${xml}</w:p>`);
      const reopened = fromProseDoc(toProseDoc(wrapParagraph(parsed)), wrapParagraph(parsed), {
        stylesheetSource: { type: "package" },
      });
      const paragraph = reopened.package.document.content.at(0);
      if (paragraph?.type !== "paragraph") throw new Error("Expected a paragraph");

      expect(serializeInnerXml(xml)).toBe(xml);
      expect(serializeParagraph(paragraph).replace(/^<w:p[^>]*>|<\/w:p>$/g, "")).toBe(xml);
    });
  }

  test(
    "random marker offsets keep live and reopened anchor text equal across nested revisions",
    () => {
      assertProperty(
        fc.property(
          fc
            .tuple(fc.integer({ min: 1, max: 12 }), fc.integer({ min: 2, max: 15 }))
            .filter(([start, end]) => start < end),
          ([start, end]) => {
            const source = "0123456789abcdef";
            const split = Math.min(start + 1, end);
            const input = wrapParagraph({
              type: "paragraph",
              content: [
                {
                  type: "deletion",
                  info: { id: 7, author: "Reviewer" },
                  content: [
                    {
                      type: "insertion",
                      info: { id: 8, author: "Reviewer" },
                      content: [
                        {
                          type: "run",
                          formatting: {},
                          content: [{ type: "text", text: source.slice(0, start) }],
                        },
                        { type: "commentRangeStart", id: 5 },
                        {
                          type: "run",
                          formatting: {},
                          content: [{ type: "text", text: source.slice(start, split) }],
                        },
                      ],
                    },
                    {
                      type: "run",
                      formatting: {},
                      content: [{ type: "text", text: source.slice(split, end) }],
                    },
                    { type: "commentRangeEnd", id: 5 },
                    {
                      type: "run",
                      formatting: {},
                      content: [{ type: "text", text: source.slice(end) }],
                    },
                  ],
                },
              ],
            });
            const projected = toProseDoc(input);
            const reopened = fromProseDoc(projected, input, {
              stylesheetSource: { type: "package" },
            });

            expect(markedText(projected, 5)).toBe(source.slice(start, end));
            expect(markedText(toProseDoc(reopened), 5)).toBe(markedText(projected, 5));
          },
        ),
        {
          ...propertyConfig({ numRuns: 100 }),
          id: "random marker offsets keep live and reopened anchor text equal across nested revisions",
        },
      );
    },
    propertyTestTimeout(10_000),
  );

  for (const fixture of [
    {
      name: "between identically formatted runs",
      content: [
        { type: "run", formatting: {}, content: [{ type: "text", text: "before" }] },
        { type: "commentRangeStart", id: 5 },
        { type: "run", formatting: {}, content: [{ type: "text", text: "anchor" }] },
        { type: "commentRangeEnd", id: 5 },
      ],
    },
    {
      name: "before the second run in a deletion",
      content: [
        {
          type: "deletion",
          info: { id: 7, author: "Reviewer" },
          content: [
            { type: "run", formatting: {}, content: [{ type: "text", text: "before" }] },
            { type: "commentRangeStart", id: 5 },
            { type: "run", formatting: {}, content: [{ type: "text", text: "anchor" }] },
            { type: "commentRangeEnd", id: 5 },
          ],
        },
      ],
    },
    {
      name: "at the start of an insertion",
      content: [
        {
          type: "insertion",
          info: { id: 7, author: "Reviewer" },
          content: [
            { type: "commentRangeStart", id: 5 },
            { type: "run", formatting: {}, content: [{ type: "text", text: "anchor" }] },
            { type: "commentRangeEnd", id: 5 },
          ],
        },
      ],
    },
  ] satisfies { name: string; content: Paragraph["content"] }[]) {
    test(`projection keeps a range start ${fixture.name}`, () => {
      const input = wrapParagraph({ type: "paragraph", content: fixture.content });
      const projected = toProseDoc(input);
      const paragraph = projected.firstChild;
      if (!paragraph) throw new Error("Expected a projected paragraph");
      let projectedOffset: number | undefined;
      paragraph.forEach((node, offset) => {
        if (node.marks.some((mark) => mark.type.name === "comment")) {
          projectedOffset ??= offset;
        }
      });
      const reopened = fromProseDoc(projected, input, {
        stylesheetSource: { type: "package" },
      });
      const reopenedParagraph = reopened.package.document.content.at(0);
      if (reopenedParagraph?.type !== "paragraph") throw new Error("Expected a paragraph");

      expect(projectedOffset).toBe(commentStartOffset(fixture.content, 5));
      expect(commentStartOffset(reopenedParagraph.content, 5)).toBe(
        commentStartOffset(fixture.content, 5),
      );
    });
  }

  for (const wrapper of TRACKED_WRAPPERS) {
    for (const marker of COMMENT_MARKERS) {
      test(`w:${wrapper.tag} keeps w:${marker} in its run sequence`, () => {
        const paragraph = parseParagraphXml(
          `<w:p xmlns:w="${WORDPROCESSINGML_NAMESPACE}">` +
            `<w:${wrapper.tag} w:id="7" w:author="Reviewer">` +
            `<w:r><w:${wrapper.tag === "del" || wrapper.tag === "moveFrom" ? "delText" : "t"} xml:space="preserve"> </w:${wrapper.tag === "del" || wrapper.tag === "moveFrom" ? "delText" : "t"}></w:r>` +
            `<w:${marker} w:id="5"/>` +
            `<w:r><w:${wrapper.tag === "del" || wrapper.tag === "moveFrom" ? "delText" : "t"}>packed</w:${wrapper.tag === "del" || wrapper.tag === "moveFrom" ? "delText" : "t"}></w:r>` +
            `</w:${wrapper.tag}></w:p>`,
        );

        const tracked = paragraph.content.at(0);
        expect(tracked?.type).toBe(wrapper.type);
        if (tracked?.type !== wrapper.type) throw new Error(`Expected ${wrapper.type}`);
        expect(tracked.content.map(({ type }) => type)).toEqual(["run", marker, "run"]);

        const serialized = serializeParagraph(paragraph);
        const wrapperXml = serialized.slice(
          serialized.indexOf(`<w:${wrapper.tag} `),
          serialized.indexOf(`</w:${wrapper.tag}>`) + `</w:${wrapper.tag}>`.length,
        );
        expect(wrapperXml.indexOf("> </")).toBeLessThan(
          wrapperXml.indexOf(`<w:${marker} w:id="5"/>`),
        );
        expect(wrapperXml.indexOf(`<w:${marker} w:id="5"/>`)).toBeLessThan(
          wrapperXml.indexOf(">packed</"),
        );
      });
    }
  }

  for (const fixture of [
    { tag: "ins", type: "insertion", removingDecision: "reject" },
    { tag: "del", type: "deletion", removingDecision: "accept" },
  ] as const) {
    for (const decision of ["accept", "reject"] as const) {
      test(`${decision} of w:${fixture.tag} resolves its complete comment range`, () => {
        const paragraph = parseParagraphXml(
          `<w:p xmlns:w="${WORDPROCESSINGML_NAMESPACE}"><w:${fixture.tag} ${REVISION_ATTRIBUTES}>` +
            `<w:r><w:${fixture.tag === "del" ? "delText" : "t"}>b</w:${fixture.tag === "del" ? "delText" : "t"}></w:r>` +
            '<w:commentRangeStart w:id="5"/>' +
            `<w:r><w:${fixture.tag === "del" ? "delText" : "t"}>et</w:${fixture.tag === "del" ? "delText" : "t"}></w:r>` +
            `<w:commentRangeEnd w:id="5"/>${COMMENT_REFERENCE}</w:${fixture.tag}></w:p>`,
        );
        const input = wrapParagraph(paragraph);
        const operation = applyDocumentOps(input, [
          {
            type: "resolveRevision",
            story: OP_STORIES.MAIN,
            decision: decision === "accept" ? REVISION_DECISIONS.ACCEPT : REVISION_DECISIONS.REJECT,
            revisionIds: [7],
          },
        ]);
        if (operation.isErr()) throw operation.error;

        const state = resolveAllChangesInHeadlessState(
          EditorState.create({ schema, doc: toProseDoc(input) }),
          decision,
        );
        const editor = fromProseDoc(state.doc, input, {
          stylesheetSource: { type: "package" },
        });
        const expected =
          decision === fixture.removingDecision
            ? []
            : ["commentRangeStart", "commentRangeEnd", "commentReference"];

        expect(markerTypes(operation.value.document)).toEqual(expected);
        expect(markerTypes(editor)).toEqual(expected);
      });
    }
  }

  test("an empty comment range stays inside its insertion", () => {
    const input = wrapParagraph({
      type: "paragraph",
      content: [
        {
          type: "insertion",
          info: { id: 7, author: "Reviewer" },
          content: [
            { type: "commentRangeStart", id: 5 },
            { type: "commentRangeEnd", id: 5 },
            { type: "commentReference", id: 5 },
          ],
        },
      ],
    });

    const reopened = fromProseDoc(toProseDoc(input), input, {
      stylesheetSource: { type: "package" },
    });
    const paragraph = reopened.package.document.content.at(0);
    if (paragraph?.type !== "paragraph") throw new Error("Expected a paragraph");
    const insertion = paragraph.content.at(0);
    if (insertion?.type !== "insertion") throw new Error("Expected an insertion");

    expect(insertion.content.map(({ type }) => type)).toEqual([
      "commentRangeStart",
      "commentRangeEnd",
      "commentReference",
    ]);
  });

  test("a range end between adjacent revisions stays outside both revision wrappers", () => {
    const input = wrapParagraph({
      type: "paragraph",
      content: [
        {
          type: "insertion",
          info: { id: 7, author: "Reviewer" },
          content: [
            { type: "commentRangeStart", id: 5 },
            { type: "run", content: [{ type: "text", text: "first" }] },
          ],
        },
        { type: "commentRangeEnd", id: 5 },
        {
          type: "insertion",
          info: { id: 8, author: "Other reviewer" },
          content: [{ type: "run", content: [{ type: "text", text: "second" }] }],
        },
        { type: "commentReference", id: 5 },
      ],
    });

    const reopened = fromProseDoc(toProseDoc(input), input, {
      stylesheetSource: { type: "package" },
    });
    const paragraph = reopened.package.document.content.at(0);
    if (paragraph?.type !== "paragraph") throw new Error("Expected a paragraph");

    expect(paragraph.content.map(({ type }) => type)).toEqual([
      "insertion",
      "commentRangeEnd",
      "insertion",
      "commentReference",
    ]);
    const first = paragraph.content.at(0);
    if (first?.type !== "insertion") throw new Error("Expected the first insertion");
    expect(first.content.map(({ type }) => type)).toEqual(["commentRangeStart", "run"]);
  });

  for (const wrapper of TRACKED_WRAPPERS) {
    for (const decision of ["accept", "reject"] as const) {
      test(`${decision} ${wrapper.type} follows the comment lifecycle when its text is removed`, () => {
        const paragraph = parseParagraphXml(
          `<w:p xmlns:w="${WORDPROCESSINGML_NAMESPACE}"><w:${wrapper.tag} w:id="7" w:author="Reviewer">` +
            `<w:r><w:t>before</w:t></w:r><w:commentRangeStart w:id="5"/>` +
            `<w:r><w:t>covered</w:t></w:r><w:commentRangeEnd w:id="5"/>` +
            `<w:r><w:t>after</w:t></w:r></w:${wrapper.tag}>${COMMENT_REFERENCE}</w:p>`,
        );
        const input = wrapParagraph(paragraph);

        const operation = applyDocumentOps(input, [
          {
            type: "resolveRevision",
            story: OP_STORIES.MAIN,
            decision: decision === "accept" ? REVISION_DECISIONS.ACCEPT : REVISION_DECISIONS.REJECT,
            revisionIds: [7],
          },
        ]);
        if (operation.isErr()) throw operation.error;

        let state = EditorState.create({ schema, doc: toProseDoc(input) });
        const command = decision === "accept" ? acceptAllChanges() : rejectAllChanges();
        command(state, (transaction) => {
          state = state.apply(transaction);
        });
        const editor = fromProseDoc(state.doc, input, {
          stylesheetSource: { type: "package" },
        });

        const expected = ["commentRangeStart", "commentRangeEnd", "commentReference"];
        const removesContent =
          (decision === "reject" && (wrapper.type === "insertion" || wrapper.type === "moveTo")) ||
          (decision === "accept" && (wrapper.type === "deletion" || wrapper.type === "moveFrom"));
        expect(markerTypes(operation.value.document)).toEqual(removesContent ? [] : expected);
        const exhaustsEditorComment =
          decision === "reject" && (wrapper.type === "insertion" || wrapper.type === "moveTo");
        expect(markerTypes(editor)).toEqual(
          exhaustsEditorComment ? ["commentReference"] : expected,
        );
      });
    }
  }
});
