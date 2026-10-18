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
import { EditorState } from "prosemirror-state";

import { acceptAllChanges, rejectAllChanges } from "../prosemirror/commands/comments";
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
    item.type === "commentRangeStart" || item.type === "commentRangeEnd" ? [item.type] : [],
  );
};

describe("comment-range boundaries inside tracked changes", () => {
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

  for (const wrapper of TRACKED_WRAPPERS) {
    for (const decision of ["accept", "reject"] as const) {
      test(`${decision} ${wrapper.type} agrees across resolution paths and retains marker order`, () => {
        const paragraph = parseParagraphXml(
          `<w:p xmlns:w="${WORDPROCESSINGML_NAMESPACE}"><w:${wrapper.tag} w:id="7" w:author="Reviewer">` +
            `<w:r><w:t>before</w:t></w:r><w:commentRangeStart w:id="5"/>` +
            `<w:r><w:t>covered</w:t></w:r><w:commentRangeEnd w:id="5"/>` +
            `<w:r><w:t>after</w:t></w:r></w:${wrapper.tag}></w:p>`,
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

        const expected = ["commentRangeStart", "commentRangeEnd"];
        expect(markerTypes(operation.value.document)).toEqual(expected);
        const removesInsertedContent =
          decision === "reject" && (wrapper.type === "insertion" || wrapper.type === "moveTo");
        expect(markerTypes(editor)).toEqual(removesInsertedContent ? [] : expected);
      });
    }
  }
});
