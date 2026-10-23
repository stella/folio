/**
 * A tracked move's range markers written between blocks are typed blocks.
 *
 * `w:body`, `w:tc` and a block content control admit `w:moveFromRangeStart`,
 * `w:moveToRangeStart` and their ends beside paragraphs and tables. The model
 * keeps each as a block between the same two siblings, the editor keeps it as
 * a block-level boundary node, and the save writes it back there, so the
 * passes that pair a start with its end can see both halves.
 */

import { describe, expect, setDefaultTimeout, test } from "bun:test";
import fc from "fast-check";
import JSZip from "jszip";

import {
  propertyConfig,
  propertyTestTimeout,
} from "../../../../test/property-testing";
import { fromProseDoc } from "../prosemirror/conversion/fromProseDoc";
import { toProseDoc } from "../prosemirror/conversion/toProseDoc";
import type { BlockContent, Document } from "../types/document";
import { parseDocx } from "./parser";
import { createEmptyDocx, repackDocx } from "./rezip";

setDefaultTimeout(propertyTestTimeout(30_000));

const XML_DECLARATION =
  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';
const W_NAMESPACE =
  "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

const MOVES = ["moveFrom", "moveTo"] as const;
const PLACEMENTS = ["bodyStart", "betweenBlocks", "bodyEnd", "inCell"] as const;

const markup = (move: (typeof MOVES)[number], id: number) => ({
  start: `<w:${move}RangeStart w:id="${id}" w:name="mv${id}" w:author="A" w:date="2024-01-01T00:00:00Z"/>`,
  end: `<w:${move}RangeEnd w:id="${id}"/>`,
});

const paragraph = (text: string): string =>
  `<w:p><w:r><w:t>${text}</w:t></w:r></w:p>`;

const bodyFor = (
  placement: (typeof PLACEMENTS)[number],
  { start, end }: { start: string; end: string },
): string => {
  switch (placement) {
    case "bodyStart":
      return `${start}${paragraph("one")}${end}${paragraph("two")}`;
    case "betweenBlocks":
      return `${paragraph("one")}${start}${paragraph("two")}${end}${paragraph("three")}`;
    case "bodyEnd":
      return `${paragraph("one")}${start}${paragraph("two")}${end}`;
    case "inCell":
      return (
        `${paragraph("one")}<w:tbl><w:tblPr/><w:tblGrid><w:gridCol w:w="2000"/></w:tblGrid>` +
        `<w:tr><w:tc><w:tcPr/>${start}${paragraph("cell")}${end}</w:tc></w:tr></w:tbl>${paragraph("two")}`
      );
    default: {
      const unreachable: never = placement;
      return unreachable;
    }
  }
};

const buildDocx = async (body: string): Promise<ArrayBuffer> => {
  const zip = await JSZip.loadAsync(await createEmptyDocx());
  zip.file(
    "word/document.xml",
    `${XML_DECLARATION}<w:document xmlns:w="${W_NAMESPACE}"><w:body>${body}<w:sectPr/></w:body></w:document>`,
  );
  return zip.generateAsync({ type: "arraybuffer" });
};

const parse = (bytes: ArrayBuffer) => parseDocx(bytes, { preloadFonts: false });

/** The blocks in reading order, with a cell's own blocks after their table. */
const outline = (blocks: readonly BlockContent[]): string[] =>
  blocks.flatMap((block) => {
    switch (block.type) {
      case "paragraph":
        return [
          `p:${block.content.map((item) => JSON.stringify(item)).join("")}`,
        ];
      case "table":
        return [
          "table",
          ...block.rows.flatMap((row) =>
            row.cells.flatMap((cell) => outline(cell.content)),
          ),
        ];
      case "moveFromRangeStart":
      case "moveFromRangeEnd":
      case "moveToRangeStart":
      case "moveToRangeEnd":
        return [JSON.stringify(block)];
      default:
        return [block.type];
    }
  });

const throughEditor = (document: Document): Document =>
  fromProseDoc(toProseDoc(document), document, {
    stylesheetSource: { type: "package" },
  });

describe("move range markers between blocks", () => {
  test("parse as typed blocks, survive the editor and the save in place, and reparse alike", async () => {
    await fc.assert(
      fc.asyncProperty(
        fc.constantFrom(...MOVES),
        fc.constantFrom(...PLACEMENTS),
        fc.integer({ min: 1, max: 5000 }),
        async (move, placement, id) => {
          const marker = markup(move, id);
          const parsed = await parse(
            await buildDocx(bodyFor(placement, marker)),
          );
          const blocks = parsed.package.document.content;
          const typed = outline(blocks).filter((entry) =>
            entry.includes(`"type":"${move}Range`),
          );
          // Both halves are typed blocks, not captured bytes.
          expect(typed.map((entry) => JSON.parse(entry).type)).toEqual([
            `${move}RangeStart`,
            `${move}RangeEnd`,
          ]);

          const edited = throughEditor(parsed);
          expect(outline(edited.package.document.content)).toEqual(
            outline(blocks),
          );

          const saved = await repackDocx(edited, { updateModifiedDate: false });
          const xml =
            (await (
              await JSZip.loadAsync(saved)
            )
              .file("word/document.xml")
              ?.async("text")) ?? "";
          const startAt = xml.indexOf(marker.start);
          const endAt = xml.indexOf(marker.end);
          expect({ start: startAt !== -1, end: endAt !== -1 }).toEqual({
            start: true,
            end: true,
          });
          expect(startAt).toBeLessThan(endAt);
          // Written as a block: no paragraph or run encloses it.
          expect(xml.slice(Math.max(0, startAt - 6), startAt)).not.toMatch(
            /<w:r>|<w:p>$/u,
          );

          const reparsed = await parse(saved);
          expect(outline(reparsed.package.document.content)).toEqual(
            outline(blocks),
          );
        },
      ),
      propertyConfig({ numRuns: 40 }),
    );
  });
});
