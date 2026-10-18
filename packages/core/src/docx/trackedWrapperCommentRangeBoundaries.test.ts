/**
 * A comment-range boundary inside a run-level tracked change stays inside it.
 *
 * `CT_RunTrackChange` admits comment range markers through
 * `EG_RunLevelElts > EG_RangeMarkupElements`. The marker is a position: lifting
 * it beside the revision changes which side of the tracked text it occupies
 * after the revision is resolved.
 */

import { describe, expect, test } from "bun:test";

import type { Paragraph } from "../types/document";

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
});
