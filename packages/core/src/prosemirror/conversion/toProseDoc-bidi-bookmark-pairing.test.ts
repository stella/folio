/**
 * A bookmark pair inside a bidirectional wrapper is still a pair.
 *
 * The editor flattens `w:bdo`/`w:dir` out of the inline list, so a boundary
 * written inside one is converted at the paragraph's own level. The pairing
 * pass that decides which boundaries become nodes walked every other
 * transparent wrapper and not this one, so the pair looked unmatched: the
 * start fell back to the legacy paragraph attribute and the end became
 * nothing at all.
 */

import { describe, expect, test } from "bun:test";

import type { Document, Paragraph } from "../../types/document";
import { createEmptyDocument } from "../../utils/createDocument";
import { toProseDoc } from "./toProseDoc";

const documentWith = (content: Paragraph["content"]): Document => {
  const template = createEmptyDocument();
  return {
    ...template,
    package: {
      ...template.package,
      document: {
        ...template.package.document,
        content: [{ type: "paragraph", paraId: "C0000001", content }],
      },
    },
  };
};

const boundaryNames = (content: Paragraph["content"]): string[] => {
  const names: string[] = [];
  toProseDoc(documentWith(content)).descendants((node) => {
    if (node.type.name === "bookmarkBoundary") {
      names.push(`${String(node.attrs["type"])}:${String(node.attrs["id"])}`);
    }
  });
  return names;
};

const pair: Paragraph["content"] = [
  { type: "bookmarkStart", id: 7, name: "clause" },
  { type: "run", content: [{ type: "text", text: "abc" }] },
  { type: "bookmarkEnd", id: 7 },
];

describe("bookmark pairing reads through a bidirectional wrapper", () => {
  test("a pair written inside an override becomes two boundary nodes", () => {
    expect(
      boundaryNames([{ type: "inlineWrapper", kind: "bidi", control: "override", content: pair }]),
    ).toEqual(["start:7", "end:7"]);
  });

  test("a pair split across the wrapper's edge still pairs", () => {
    expect(
      boundaryNames([
        { type: "bookmarkStart", id: 7, name: "clause" },
        {
          type: "inlineWrapper",
          kind: "bidi",
          control: "embedding",
          direction: "rtl",
          content: [{ type: "run", content: [{ type: "text", text: "abc" }] }],
        },
        { type: "bookmarkEnd", id: 7 },
      ]),
    ).toEqual(["start:7", "end:7"]);
  });

  test("the wrapper does not invent a pair out of a lone start", () => {
    expect(
      boundaryNames([
        {
          type: "inlineWrapper",
          kind: "bidi",
          control: "override",
          direction: "rtl",
          content: [{ type: "bookmarkStart", id: 7, name: "clause" }],
        },
      ]),
    ).toEqual([]);
  });
});
