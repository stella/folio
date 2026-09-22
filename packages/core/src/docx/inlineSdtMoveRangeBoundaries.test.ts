/**
 * A tracked-move range boundary inside an inline content control stays there.
 *
 * `CT_SdtContentRun` reaches all four move-range markers through
 * `EG_RunLevelElts > EG_RangeMarkupElements`. Lifting a marker beside the
 * control changes the range's authored extent, while ProseMirror can carry the
 * position directly with its zero-width move-boundary and empty-range nodes.
 */

import { describe, expect, test } from "bun:test";

import { fromProseDoc } from "../prosemirror/conversion/fromProseDoc";
import { toProseDoc } from "../prosemirror/conversion/toProseDoc";
import type { Document, InlineSdt, Paragraph } from "../types/document";

import { parseParagraph } from "./paragraphParser";
import { serializeParagraph } from "./serializer/paragraphSerializer";
import { parseXmlDocument } from "./xmlParser";

const WORDPROCESSINGML_NAMESPACE = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

const MOVE_RANGES = [
  {
    name: "move-from",
    start:
      '<w:moveFromRangeStart w:id="41" w:displacedByCustomXml="next" w:colFirst="1" w:colLast="2" w:name="source" w:author="Reviewer" w:date="2026-09-22T08:00:00Z"/>',
    startType: "moveFromRangeStart",
    end: '<w:moveFromRangeEnd w:id="41" w:displacedByCustomXml="prev"/>',
    endType: "moveFromRangeEnd",
  },
  {
    name: "move-to",
    start:
      '<w:moveToRangeStart w:id="42" w:displacedByCustomXml="next" w:colFirst="3" w:colLast="4" w:name="destination" w:author="Reviewer" w:date="2026-09-22T08:00:00Z"/>',
    startType: "moveToRangeStart",
    end: '<w:moveToRangeEnd w:id="42" w:displacedByCustomXml="prev"/>',
    endType: "moveToRangeEnd",
  },
] as const;

const parseParagraphXml = (xml: string): Paragraph => {
  const element = parseXmlDocument(
    xml.replace("<w:p", `<w:p xmlns:w="${WORDPROCESSINGML_NAMESPACE}"`),
  );
  if (!element) {
    throw new Error("Failed to parse the inline content-control fixture");
  }
  return parseParagraph(element, null, null, null, null, null);
};

const throughEditor = (paragraph: Paragraph): Paragraph => {
  const document: Document = {
    package: { document: { content: [paragraph] } },
  };
  const rebuilt = fromProseDoc(toProseDoc(document), document).package.document.content.at(0);
  if (rebuilt?.type !== "paragraph") {
    throw new Error("Expected the editor round trip to rebuild one paragraph");
  }
  return rebuilt;
};

const inlineControl = (paragraph: Paragraph): InlineSdt => {
  const control = paragraph.content.find(({ type }) => type === "inlineSdt");
  if (control?.type !== "inlineSdt") {
    throw new Error("Expected one inline content control");
  }
  return control;
};

const fixtureXml = (start: string, end: string, inner = "<w:r><w:t>inside</w:t></w:r>"): string =>
  `<w:p><w:sdt><w:sdtPr><w:tag w:val="move-range"/></w:sdtPr><w:sdtContent>` +
  `${start}${inner}${end}` +
  "</w:sdtContent></w:sdt></w:p>";

const controlXml = (paragraphXml: string): string => {
  const start = paragraphXml.indexOf("<w:sdtContent>");
  const end = paragraphXml.indexOf("</w:sdtContent>", start);
  if (start === -1 || end === -1) {
    throw new Error("Expected serialized inline content-control content");
  }
  return paragraphXml.slice(start, end + "</w:sdtContent>".length);
};

describe("tracked-move boundaries inside inline content controls", () => {
  for (const range of MOVE_RANGES) {
    test(`${range.name} keeps both boundaries through parse, edit, and repeated save`, () => {
      const parsed = parseParagraphXml(fixtureXml(range.start, range.end));
      expect(inlineControl(parsed).content.map(({ type }) => type)).toEqual([
        range.startType,
        "run",
        range.endType,
      ]);

      const edited = throughEditor(parsed);
      expect(inlineControl(edited).content.map(({ type }) => type)).toEqual([
        range.startType,
        "run",
        range.endType,
      ]);

      const serialized = serializeParagraph(edited);
      const serializedControl = controlXml(serialized);
      expect(serializedControl).toContain(range.start);
      expect(serializedControl).toContain(range.end);
      expect(serializedControl.indexOf(range.start)).toBeLessThan(
        serializedControl.indexOf("inside"),
      );
      expect(serializedControl.indexOf("inside")).toBeLessThan(
        serializedControl.indexOf(range.end),
      );

      const fixedPoint = serializeParagraph(throughEditor(parseParagraphXml(serialized)));
      expect(fixedPoint).toBe(serialized);
    });

    test(`${range.name} keeps an empty range atomic inside the control`, () => {
      const parsed = parseParagraphXml(fixtureXml(range.start, range.end, ""));
      const projected = toProseDoc({
        package: { document: { content: [parsed] } },
      });
      let anchors = 0;
      projected.descendants((node) => {
        if (node.type.name === "rangeAnchor") {
          anchors += 1;
        }
      });
      expect(anchors).toBe(1);

      const serializedControl = controlXml(serializeParagraph(throughEditor(parsed)));
      expect(serializedControl).toContain(`${range.start}${range.end}`);
    });
  }
});
