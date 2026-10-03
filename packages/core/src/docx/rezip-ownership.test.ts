import { expect, test } from "bun:test";
import JSZip from "jszip";

import { getSourceReplayToken, inheritSourceReplayToken } from "@stll/docx-core/ops";
import type { Document, Paragraph } from "../types/document";
import { parseDocx } from "./parser";
import { RELATIONSHIP_TYPES } from "./relsParser";
import { createDocx, createEmptyDocx, repackDocx, repackDocxFromRaw } from "./rezip";
import { unzipDocx } from "./unzip";

const IMAGE =
  "data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+/p9sAAAAASUVORK5CYII=";

const resourceParagraph = (paraId: string) =>
  ({
    type: "paragraph",
    paraId,
    content: [
      {
        type: "hyperlink",
        href: "https://example.org/resource",
        children: [{ type: "run", content: [{ type: "text", text: "link" }] }],
      },
      {
        type: "run",
        content: [
          {
            type: "drawing",
            image: {
              type: "image",
              src: IMAGE,
              size: { width: 9525, height: 9525 },
              wrap: { type: "inline" },
            },
          },
        ],
      },
    ],
  }) satisfies Paragraph;

const resources = () =>
  ({
    package: {
      document: {
        content: [
          {
            ...resourceParagraph("00000001"),
            content: [
              { type: "commentRangeStart", id: 7 },
              ...resourceParagraph("00000001").content,
              { type: "commentRangeEnd", id: 7 },
              { type: "commentReference", id: 7 },
            ],
          },
        ],
        finalSectionProperties: {
          headerReferences: [{ type: "default", rId: "rId1" }],
          footerReferences: [{ type: "default", rId: "rIdFooter" }],
        },
        comments: [
          { id: 7, author: "Reviewer", content: [{ type: "paragraph", content: [] }] },
          {
            id: 8,
            author: "Reviewer",
            parentId: 7,
            content: [{ type: "paragraph", content: [] }],
          },
        ],
      },
      // rId1 belongs to styles; materialization must rebind the header reference.
      relationships: new Map([
        ["rId1", { id: "rId1", type: RELATIONSHIP_TYPES.styles, target: "styles.xml" }],
      ]),
      headers: new Map([
        [
          "rId1",
          {
            type: "header",
            hdrFtrType: "default",
            content: [
              {
                type: "blockSdt",
                properties: { sdtType: "richText", id: 9 },
                content: [resourceParagraph("00000002")],
              },
            ],
          },
        ],
      ]),
      footers: new Map([
        [
          "rIdFooter",
          { type: "footer", hdrFtrType: "default", content: [resourceParagraph("00000003")] },
        ],
      ]),
      footnotes: [{ type: "footnote", id: 2, content: [resourceParagraph("00000004")] }],
      endnotes: [{ type: "endnote", id: 3, content: [resourceParagraph("00000005")] }],
    },
  }) satisfies Document;

const freezeGraph = (value: unknown, seen = new WeakSet<object>()): void => {
  if (value === null || typeof value !== "object" || seen.has(value)) return;
  seen.add(value);
  if (value instanceof ArrayBuffer || ArrayBuffer.isView(value)) return;
  if (value instanceof Map)
    for (const [key, entry] of value) {
      freezeGraph(key, seen);
      freezeGraph(entry, seen);
    }
  else if (value instanceof Set) for (const entry of value) freezeGraph(entry, seen);
  else for (const entry of Object.values(value)) freezeGraph(entry, seen);
  Object.freeze(value);
};

test.each(["fresh", "createParsed", "repack", "repackRaw"] as const)(
  "%s exports resource relationships without changing caller ownership",
  async (path) => {
    const source = await createEmptyDocx();
    const parsed = await parseDocx(source, { preloadFonts: false, sourceReplay: "tracked" });
    const authored = resources();
    const document = (
      path === "fresh"
        ? authored
        : {
            ...parsed,
            package: {
              ...parsed.package,
              ...authored.package,
              document: {
                ...parsed.package.document,
                ...authored.package.document,
                content: [...parsed.package.document.content, ...authored.package.document.content],
              },
            },
          }
    ) satisfies Document;
    if (path !== "fresh") inheritSourceReplayToken(document, parsed);
    const token = getSourceReplayToken(document);
    if (path !== "fresh") expect(token).toBeDefined();
    const before = structuredClone(document);
    freezeGraph(document);
    const raw = path === "repackRaw" ? await unzipDocx(source) : undefined;
    let output: ArrayBuffer;
    switch (path) {
      case "fresh":
      case "createParsed":
        output = await createDocx(document);
        break;
      case "repack":
        output = await repackDocx(document, { sourceReplay: token, updateModifiedDate: false });
        break;
      case "repackRaw":
        if (!raw) throw new Error("Raw export requires its source package");
        output = await repackDocxFromRaw(document, raw, { updateModifiedDate: false });
        break;
    }
    expect(document).toEqual(before);
    expect(getSourceReplayToken(document)).toBe(token);
    const zip = await JSZip.loadAsync(output);
    const relationshipParts = Object.values(zip.files).filter(({ name }) => name.endsWith(".rels"));
    const relationships = await Promise.all(relationshipParts.map((part) => part.async("text")));
    expect(relationships.filter((xml) => xml.includes(RELATIONSHIP_TYPES.hyperlink)).length).toBe(
      5,
    );
    expect(relationships.filter((xml) => xml.includes(RELATIONSHIP_TYPES.image)).length).toBe(5);
    const reopened = await parseDocx(output, { preloadFonts: false });
    expect(reopened.package.headers?.size).toBe(1);
    expect(reopened.package.footers?.size).toBe(1);
    expect(reopened.package.footnotes?.some(({ id }) => id === 2)).toBe(true);
    expect(reopened.package.endnotes?.some(({ id }) => id === 3)).toBe(true);
    expect(reopened.package.document.comments?.some(({ parentId }) => parentId === 7)).toBe(true);
  },
);
