/**
 * A kept range marker's cut (`keptCut`) is review state the docx-core
 * resolution keeps between resolutions. It is no attribute of the marker:
 * saving never writes it, and the editor model does not carry it.
 */

import { describe, expect, test } from "bun:test";

import { fromProseDoc } from "../prosemirror/conversion/fromProseDoc";
import { toProseDoc } from "../prosemirror/conversion/toProseDoc";
import type { Document, Paragraph } from "../types/document";
import { parseParagraph } from "./paragraphParser";
import { serializeParagraph } from "./serializer/paragraphSerializer";
import { parseXmlDocument, type XmlElement } from "./xmlParser";

const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const CUT = { depth: 2, keptFrom: "2026-03-04T05:00:01Z" };

const paragraph: Paragraph = {
  type: "paragraph",
  paraId: "00000001",
  content: [
    { type: "bookmarkStart", id: 1, name: "kept", keptCut: CUT },
    { type: "run", content: [{ type: "text", text: "a" }] },
    { type: "bookmarkEnd", id: 1, keptCut: CUT },
    { type: "commentRangeStart", id: 2, keptCut: CUT },
    {
      type: "moveToRangeStart",
      id: 3,
      name: "move3",
      author: "A",
      keptCut: CUT,
    },
    { type: "run", content: [{ type: "text", text: "b" }] },
    { type: "moveToRangeEnd", id: 3, keptCut: CUT },
    { type: "commentRangeEnd", id: 2, keptCut: CUT },
  ],
};

const keptCuts = (value: unknown): unknown[] => {
  if (Array.isArray(value)) return value.flatMap(keptCuts);
  if (typeof value !== "object" || value === null) return [];
  return Object.entries(value).flatMap(([key, field]) =>
    key === "keptCut" ? [field] : keptCuts(field),
  );
};

describe("a kept range marker's cut is never written", () => {
  test("saving and reopening the paragraph keeps the markers and drops the cut", () => {
    const saved = serializeParagraph(paragraph);
    expect(saved).not.toContain(CUT.keptFrom);
    expect(saved).not.toContain("keptCut");
    const root = parseXmlDocument(
      saved.replace("<w:p", `<w:p xmlns:w="${W}"`),
    ) as XmlElement | null;
    if (!root) throw new Error("The saved paragraph parses.");
    const reopened = parseParagraph(root, null, null, null, null, null);
    expect(reopened.content.map(({ type }) => type)).toEqual(
      paragraph.content.map(({ type }) => type),
    );
    expect(keptCuts(reopened)).toEqual([]);
  });

  test("the editor model does not carry the cut", () => {
    const input: Document = { package: { document: { content: [paragraph] } } };
    const back = fromProseDoc(toProseDoc(input), input, {
      stylesheetSource: { type: "package" },
    });
    expect(keptCuts(back)).toEqual([]);
  });
});
