import { describe, expect, test } from "bun:test";
import fc from "fast-check";
import JSZip from "jszip";
import { readFileSync } from "node:fs";

import {
  propertyConfig,
  propertyTestTimeout,
} from "../../../../test/property-testing";
import { buildBookmarkText } from "../fields/bookmarkText";
import { toFlowBlocks } from "../layout-bridge/convert/toFlowBlocks";
import { fromProseDoc } from "../prosemirror/conversion/fromProseDoc";
import { toProseDoc } from "../prosemirror/conversion/toProseDoc";
import { normalizeBookmarkIdsInXmlParts } from "./bookmarkIdNormalization";
import { addBookmark, createBookmarkMap } from "./bookmarkParser";
import { parseDocx } from "./parser";
import { repackDocx } from "./rezip";

const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

// Source capability tokens intentionally bind to each saved ZIP's digest.
const semanticProjection = (doc: ReturnType<typeof toProseDoc>) => {
  const nodes: unknown[] = [];
  doc.descendants((node, position) => {
    const { _docxParagraphSourceToken, ...attrs } = node.attrs;
    nodes.push({
      position,
      type: node.type.name,
      attrs,
      text: node.text,
      marks: node.marks.map((mark) => mark.toJSON()),
    });
  });
  return nodes;
};

describe("imported bookmark collisions", () => {
  test(
    "every range survives with unique ids and repair is a fixed point",
    () => {
      fc.assert(
        fc.property(
          fc.array(fc.integer({ min: 0, max: 12 }), {
            minLength: 1,
            maxLength: 40,
          }),
          (ids) => {
            const xml = `<x:document xmlns:x="${W}"><x:body>${ids.map((id, index) => `<x:p><x:bookmarkStart x:id="${id}" x:name="b${index}"/><x:r><x:t>text ${index}</x:t></x:r><x:bookmarkEnd x:id="${id}"/></x:p>`).join("")}</x:body></x:document>`;
            const parts = new Map([["word/document.xml", xml]]);
            const repaired = normalizeBookmarkIdsInXmlParts(parts);
            const output = repaired.get("word/document.xml") ?? "";
            const starts = [
              ...output.matchAll(/bookmarkStart x:id="(\d+)"/gu),
            ].map((match) => Number(match.at(1)));
            const ends = [...output.matchAll(/bookmarkEnd x:id="(\d+)"/gu)].map(
              (match) => Number(match.at(1)),
            );
            expect(starts.length).toBe(ids.length);
            expect(new Set(starts).size).toBe(ids.length);
            expect(ends).toEqual(starts);
            expect(output.replace(/x:id="\d+"/gu, "")).toBe(
              xml.replace(/x:id="\d+"/gu, ""),
            );
            const seen = new Set<number>();
            for (const [index, id] of ids.entries()) {
              if (!seen.has(id)) expect(starts.at(index)).toBe(id);
              seen.add(id);
            }
            expect(normalizeBookmarkIdsInXmlParts(repaired)).toBe(repaired);
          },
        ),
        propertyConfig(),
      );
    },
    propertyTestTimeout(5_000),
  );

  test("overlapping duplicates close in encounter order and reserve future ids", () => {
    const xml = `<x:p xmlns:x="${W}"><x:bookmarkStart x:id="7" x:name="first"/><x:bookmarkStart x:id="7" x:name="later"/><x:bookmarkEnd x:id="7"/><x:bookmarkEnd x:id="7"/><x:bookmarkStart x:id="0" x:name="reserved"/><x:bookmarkEnd x:id="0"/></x:p>`;
    const result = normalizeBookmarkIdsInXmlParts(
      new Map([["word/document.xml", xml]]),
    );
    expect(result.get("word/document.xml")).toBe(
      xml
        .replace('x:id="7" x:name="later"', 'x:id="1" x:name="later"')
        .replace(
          '<x:bookmarkEnd x:id="7"/><x:bookmarkStart',
          '<x:bookmarkEnd x:id="1"/><x:bookmarkStart',
        ),
    );
  });

  test("lookup retains the first bookmark by id and name", () => {
    const map = createBookmarkMap();
    const first = { type: "bookmarkStart", id: 7, name: "target" } as const;
    addBookmark(map, first);
    addBookmark(map, { type: "bookmarkStart", id: 7, name: "target" });
    expect(map.byId.get(7)).toBe(first);
    expect(map.byName.get("target")).toBe(first);
    expect(map.bookmarks.length).toBe(2);
  });

  test("mutually exclusive compatibility branches do not collide", () => {
    const range = `<x:p><x:bookmarkStart x:id="7" x:name="target"/><x:r><x:t>Target</x:t></x:r><x:bookmarkEnd x:id="7"/></x:p>`;
    for (const branches of [
      `<mc:Choice Requires="x">${range}</mc:Choice><mc:Fallback>${range}</mc:Fallback>`,
      `<mc:Fallback>${range}</mc:Fallback><mc:Choice Requires="x">${range}</mc:Choice>`,
      `<mc:Fallback>${range}</mc:Fallback>`,
    ]) {
      const xml = `<x:document xmlns:x="${W}" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"><x:body><mc:AlternateContent>${branches}</mc:AlternateContent></x:body></x:document>`;
      const parts = new Map([["word/document.xml", xml]]);
      expect(normalizeBookmarkIdsInXmlParts(parts)).toBe(parts);
    }
  });

  test("corpus document opens, projects both ranges, resolves REF first, and saves stably", async () => {
    const bytes = readFileSync(
      new URL(
        "./__tests__/__fixtures__/corpus/duplicate-bookmark-ids.docx",
        import.meta.url,
      ),
    );
    const document = await parseDocx(
      bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength),
    );
    expect(
      document.parseWarnings?.some(
        ({ code }) => code === "duplicate-imported-id",
      ),
    ).toBe(true);
    const pm = toProseDoc(document);
    expect(pm.textContent).toContain("First target");
    expect(pm.textContent).toContain("Later target");
    const boundaries: number[] = [];
    pm.descendants((node) => {
      if (
        node.type.name === "bookmarkBoundary" &&
        node.attrs["type"] === "start"
      )
        boundaries.push(node.attrs["id"]);
    });
    expect(new Set(boundaries).size).toBe(2);
    expect(buildBookmarkText(toFlowBlocks(pm)).get("target")).toBe(
      "First target",
    );
    const save = async (input: typeof document) =>
      repackDocx(
        fromProseDoc(toProseDoc(input), input, {
          stylesheetSource: { type: "package" },
        }),
        { updateModifiedDate: false },
      );
    const saved = await save(document);
    const reopened = await parseDocx(saved);
    expect(semanticProjection(toProseDoc(reopened))).toEqual(
      semanticProjection(pm),
    );
    const savedAgain = await save(reopened);
    const firstZip = await JSZip.loadAsync(saved);
    const secondZip = await JSZip.loadAsync(savedAgain);
    expect(await secondZip.file("word/document.xml")?.async("string")).toBe(
      await firstZip.file("word/document.xml")?.async("string"),
    );
  });
});
