/**
 * A move-range boundary inside a run-level tracked change stays inside it.
 *
 * `CT_RunTrackChange` reaches all four move-range markers through
 * `EG_RunLevelElts > EG_RangeMarkupElements`. Lifting one beside the revision
 * changes which review operation owns the marker and can leave the move range
 * crossing content it did not originally cover.
 */

import { describe, expect, test } from "bun:test";

import { fromProseDoc } from "../prosemirror/conversion/fromProseDoc";
import { toProseDoc } from "../prosemirror/conversion/toProseDoc";
import type { Document, Paragraph, TrackedRunContent } from "../types/document";

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

const MOVE_RANGE_MARKERS = [
  {
    tag: "moveFromRangeStart",
    xml: '<w:moveFromRangeStart w:id="41" w:displacedByCustomXml="next" w:colFirst="1" w:colLast="2" w:name="source" w:author="Reviewer" w:date="2026-09-22T08:00:00Z"/>',
    pairXml: '<w:moveFromRangeEnd w:id="41"/>',
    starts: true,
  },
  {
    tag: "moveFromRangeEnd",
    xml: '<w:moveFromRangeEnd w:id="42" w:displacedByCustomXml="prev"/>',
    pairXml: '<w:moveFromRangeStart w:id="42" w:name="source-end" w:author="Reviewer"/>',
    starts: false,
  },
  {
    tag: "moveToRangeStart",
    xml: '<w:moveToRangeStart w:id="43" w:displacedByCustomXml="next" w:colFirst="3" w:colLast="4" w:name="destination" w:author="Reviewer" w:date="2026-09-22T08:00:00Z"/>',
    pairXml: '<w:moveToRangeEnd w:id="43"/>',
    starts: true,
  },
  {
    tag: "moveToRangeEnd",
    xml: '<w:moveToRangeEnd w:id="44" w:displacedByCustomXml="prev"/>',
    pairXml: '<w:moveToRangeStart w:id="44" w:name="destination-end" w:author="Reviewer"/>',
    starts: false,
  },
] as const satisfies ReadonlyArray<{
  tag: TrackedRunContent["type"];
  xml: string;
  pairXml: string;
  starts: boolean;
}>;

const parseParagraphXml = (xml: string): Paragraph => {
  const namespaced = xml.includes("xmlns:w=")
    ? xml
    : xml.replace("<w:p", `<w:p xmlns:w="${WORDPROCESSINGML_NAMESPACE}"`);
  const element = parseXmlDocument(namespaced);
  if (!element) {
    throw new Error("Failed to parse the tracked-change fixture");
  }
  return parseParagraph(element, null, null, null, null, null);
};

const wrapParagraph = (paragraph: Paragraph): Document => ({
  package: { document: { content: [paragraph] } },
});

const throughEditor = (paragraph: Paragraph): Paragraph => {
  const document = wrapParagraph(paragraph);
  const rebuilt = fromProseDoc(toProseDoc(document), document).package.document.content.at(0);
  if (rebuilt?.type !== "paragraph") {
    throw new Error("Expected the editor round trip to rebuild one paragraph");
  }
  return rebuilt;
};

const fixtureXml = (
  wrapper: (typeof TRACKED_WRAPPERS)[number],
  marker: (typeof MOVE_RANGE_MARKERS)[number],
): string => {
  const revision =
    `<w:${wrapper.tag} w:id="7" w:author="Reviewer" w:date="2026-09-22T08:00:00Z">` +
    `<w:r><w:t>before</w:t></w:r>${marker.xml}<w:r><w:t>after</w:t></w:r>` +
    `</w:${wrapper.tag}>`;
  return (
    `<w:p xmlns:w="${WORDPROCESSINGML_NAMESPACE}">` +
    `${marker.starts ? revision + marker.pairXml : marker.pairXml + revision}</w:p>`
  );
};

const trackedWrapperXml = (xml: string, tag: string): string => {
  const start = xml.indexOf(`<w:${tag} `);
  const end = xml.indexOf(`</w:${tag}>`, start);
  if (start === -1 || end === -1) {
    throw new Error(`Expected serialized w:${tag}`);
  }
  return xml.slice(start, end + `</w:${tag}>`.length);
};

describe("move-range boundaries inside tracked changes", () => {
  for (const wrapper of TRACKED_WRAPPERS) {
    for (const marker of MOVE_RANGE_MARKERS) {
      test(`w:${wrapper.tag} keeps w:${marker.tag} through parse, edit, and repeated save`, () => {
        const parsed = parseParagraphXml(fixtureXml(wrapper, marker));
        const tracked = parsed.content.find((content) => content.type === wrapper.type);
        expect(tracked?.type).toBe(wrapper.type);
        if (tracked?.type !== wrapper.type) {
          throw new Error(`Expected parsed ${wrapper.type}`);
        }
        expect(tracked.content.map(({ type }) => type)).toEqual(["run", marker.tag, "run"]);

        const serialized = serializeParagraph(parsed);
        const serializedWrapper = trackedWrapperXml(serialized, wrapper.tag);
        expect(serializedWrapper.indexOf("before")).toBeLessThan(
          serializedWrapper.indexOf(`<w:${marker.tag} `),
        );
        expect(serializedWrapper.indexOf(`<w:${marker.tag} `)).toBeLessThan(
          serializedWrapper.indexOf("after"),
        );
        expect(serializedWrapper).toContain(marker.xml);

        const edited = serializeParagraph(throughEditor(parsed));
        expect(trackedWrapperXml(edited, wrapper.tag)).toContain(marker.xml);

        const fixedPoint = serializeParagraph(throughEditor(parseParagraphXml(edited)));
        expect(fixedPoint).toBe(edited);
      });
    }
  }
});
