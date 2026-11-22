/**
 * A saved part states every section break the model holds, and none inside a
 * table cell, whatever operation was asked to put one there.
 */

import { panic } from "better-result";
import { expect, setDefaultTimeout, test } from "bun:test";
import fc from "fast-check";

import {
  applyDocumentOp,
  applyDocumentOps,
  DOCUMENT_OP_REFUSAL_REASONS,
  DOCUMENT_OP_TYPES,
  type DocumentOp,
  OP_STORIES,
} from "@stll/docx-core/ops";
import { ZipArchive } from "@stll/docx-core/zip";

import { assertExactModel } from "../../../../test/exactModel";
import { checkCoverageSample } from "../../../../test/property-coverage";
import {
  assertProperty,
  propertyTestTimeout,
} from "../../../../test/property-testing";
import type {
  BlockContent,
  Document,
  Paragraph,
  Table,
} from "../types/document";
import { createEmptyDocument } from "../utils/createDocument";
import { createDocx } from "./rezip";

setDefaultTimeout(propertyTestTimeout(60_000));

const SECTION = { pageWidth: 11_906, pageHeight: 16_838 };

type BlockShape =
  | { kind: "paragraph"; endsSection: boolean }
  | {
      kind: "table";
      cells: { paragraphs: number; holder: "none" | "control" | "table" }[];
    };

const shapeArbitrary: fc.Arbitrary<BlockShape> = fc.oneof(
  fc.record({
    kind: fc.constant("paragraph" as const),
    endsSection: fc.boolean(),
  }),
  fc.record({
    kind: fc.constant("table" as const),
    cells: fc.array(
      fc.record({
        paragraphs: fc.integer({ min: 1, max: 2 }),
        holder: fc.constantFrom(
          "none" as const,
          "control" as const,
          "table" as const,
        ),
      }),
      { minLength: 1, maxLength: 2 },
    ),
  }),
);

/** A paragraph of the document, the row directly holding it, and that row's table. */
type Place = { paragraph: Paragraph; table: Table | undefined };

const drawn = (shapes: readonly BlockShape[]) => {
  let next = 1;
  let control = 1;
  const places: Place[] = [];
  const paragraphIn = (table: Table | undefined): Paragraph => {
    const id = next++;
    const paragraph: Paragraph = {
      type: "paragraph",
      paraId: id.toString(16).padStart(8, "0").toUpperCase(),
      content: [
        { type: "run", content: [{ type: "text", text: `Text ${id}` }] },
      ],
    };
    places.push({ paragraph, table });
    return paragraph;
  };
  const tableHolding = (
    cells: Extract<BlockShape, { kind: "table" }>["cells"],
  ): Table => {
    const table: Table = {
      type: "table",
      columnWidths: cells.map(() => 2400),
      rows: [],
    };
    table.rows.push({
      type: "tableRow",
      cells: cells.map(({ paragraphs, holder }) => {
        const held = (): BlockContent[] =>
          Array.from({ length: paragraphs }, () => paragraphIn(table));
        switch (holder) {
          case "none":
            return { type: "tableCell", content: held() };
          case "control":
            return {
              type: "tableCell",
              content: [
                {
                  type: "blockSdt",
                  properties: { sdtType: "richText", id: control++ },
                  content: held(),
                },
                paragraphIn(table),
              ],
            };
          case "table":
            return {
              type: "tableCell",
              content: [
                tableHolding([{ paragraphs, holder: "none" }]),
                paragraphIn(table),
              ],
            };
          default: {
            const unreachable: never = holder;
            return unreachable;
          }
        }
      }),
    });
    return table;
  };
  const content = shapes.map((shape): BlockContent => {
    if (shape.kind === "table") return tableHolding(shape.cells);
    const paragraph = paragraphIn(undefined);
    if (shape.endsSection) paragraph.sectionProperties = SECTION;
    return paragraph;
  });
  content.push(paragraphIn(undefined));
  const document = createEmptyDocument();
  document.package.document.content = content;
  document.package.document.sections = undefined;
  return { document, places };
};

const story = OP_STORIES.MAIN;
const WAYS = ["replace", "endpoint", "split", "row", "rows", "table"] as const;
type Way = (typeof WAYS)[number];

/** An operation that ends a section at a place, or at a new paragraph beside it. */
const sectionOp = (
  way: Way,
  { paragraph, table }: Place,
): DocumentOp | undefined => {
  const blockId = paragraph.paraId ?? panic("A drawn paragraph has an id.");
  const ending = { ...paragraph, sectionProperties: SECTION };
  const row = {
    type: "tableRow" as const,
    cells: [
      {
        type: "tableCell" as const,
        content: [
          {
            type: "paragraph" as const,
            paraId: "0000A000",
            content: [],
            sectionProperties: SECTION,
          },
        ],
      },
    ],
  };
  switch (way) {
    case "replace":
      return {
        type: DOCUMENT_OP_TYPES.REPLACE_BLOCKS,
        story,
        expected: [paragraph],
        blocks: [ending],
        sectionBoundaries: "replace",
      };
    case "endpoint":
      return {
        type: DOCUMENT_OP_TYPES.SET_SECTION_ENDPOINT,
        endpoint: { type: "paragraph", blockId },
        properties: { type: "present", value: SECTION },
      };
    case "split":
      return {
        type: DOCUMENT_OP_TYPES.SPLIT_BLOCK,
        at: { story, blockId, offset: 2 },
        newBlockId: "0000A000",
        newHalf: "first",
        firstSectionProperties: SECTION,
      };
    case "row":
      return table === undefined
        ? undefined
        : {
            type: DOCUMENT_OP_TYPES.INSERT_ROW,
            story,
            blockId,
            at: 0,
            row,
          };
    case "rows":
      return table === undefined
        ? undefined
        : {
            type: DOCUMENT_OP_TYPES.SET_TABLE_ROWS,
            story,
            blockId,
            expected: table.rows,
            rows: [...table.rows, row],
          };
    case "table":
      return table === undefined
        ? undefined
        : {
            type: DOCUMENT_OP_TYPES.SET_TABLE,
            story,
            blockId,
            expected: table,
            table: { ...table, rows: [...table.rows, row] },
          };
    default: {
      const unreachable: never = way;
      return unreachable;
    }
  }
};

/** Section breaks of a saved body, and those of them inside a table cell. */
const savedSections = async (
  document: Document,
): Promise<{ total: number; inCells: number }> => {
  const zip = await ZipArchive.read(
    await createDocx(structuredClone(document)),
  );
  const xml = await zip.entry("word/document.xml")?.text();
  if (xml === undefined) panic("The saved package holds its document part.");
  let depth = 0;
  let total = 0;
  let inCells = 0;
  for (const [tag] of xml.matchAll(/<w:tc>|<\/w:tc>|<w:sectPr[\s/>]/gu)) {
    if (tag === "<w:tc>") depth += 1;
    else if (tag === "</w:tc>") depth -= 1;
    else {
      total += 1;
      if (depth > 0) inCells += 1;
    }
  }
  return { total, inCells };
};

type Outcome = "refusedInCell" | "appliedOutside" | "refusedOutside" | "none";

const sectionSaves = (tally: Map<Outcome, number>) =>
  fc.asyncProperty(
    fc.array(shapeArbitrary, { minLength: 1, maxLength: 4 }),
    fc.nat(),
    fc.constantFrom(...WAYS),
    async (shapes, at, way) => {
      const { document, places } = drawn(shapes);
      const place = places[at % places.length];
      const op = place === undefined ? undefined : sectionOp(way, place);
      if (place === undefined || op === undefined) {
        tally.set("none", (tally.get("none") ?? 0) + 1);
        return;
      }
      const before = structuredClone(document);
      const result = applyDocumentOp(document, op);
      assertExactModel(document, before);
      const saved = result.isOk() ? result.value.document : document;
      const breaks = saved.package.document.content.filter(
        (block) =>
          block.type === "paragraph" && block.sectionProperties !== undefined,
      ).length;
      // The body states its final section as well.
      expect(await savedSections(saved)).toEqual({
        total: breaks + 1,
        inCells: 0,
      });
      const outcome = ((): Outcome => {
        if (place.table === undefined)
          return result.isOk() ? "appliedOutside" : "refusedOutside";
        expect(result.isErr() ? result.error.reason : "applied").toBe(
          DOCUMENT_OP_REFUSAL_REASONS.SECTION_BOUNDARY,
        );
        return "refusedInCell";
      })();
      tally.set(outcome, (tally.get(outcome) ?? 0) + 1);
      if (result.isErr()) return;
      const inverse: DocumentOp[] = JSON.parse(
        JSON.stringify(result.value.inverse),
      );
      assertExactModel(
        applyDocumentOps(result.value.document, inverse).unwrap().document,
        before,
      );
    },
  );

test("a saved part holds no section break inside a table cell", async () => {
  await assertProperty(sectionSaves(new Map()), {
    id: "a saved part holds no section break inside a table cell",
    numRuns: 100,
  });
});

test("the cases reach cells and paragraphs outside them", async () => {
  const tally = new Map<Outcome, number>();
  await checkCoverageSample(sectionSaves(tally), 100);
  expect({
    refusedInCell: (tally.get("refusedInCell") ?? 0) >= 20,
    appliedOutside: (tally.get("appliedOutside") ?? 0) >= 5,
    tally: Object.fromEntries(tally),
  }).toMatchObject({ refusedInCell: true, appliedOutside: true });
});
