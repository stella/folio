/**
 * A container's verbatim sink survives the editor, not only a save.
 *
 * `dispatchChildren` keeps the children a container's model cannot hold and
 * records where they sat, and `serializeWithPreservedChildren` puts them back
 * between the same siblings. That closed the save leg. The editor leg was a
 * separate question for every container whose sink is a record field rather
 * than a member of a content union: `toProseDoc` had nowhere to put a table's
 * or a row's sink, so a document that was merely opened and saved through the
 * editor came back without it. A `w:bookmarkEnd` written beside a table's rows
 * is the corpus instance — losing it leaves the `w:bookmarkStart` in a cell
 * with no end, which is a bookmark Word cannot resolve.
 *
 * The property runs over arbitrary marker subsets and arbitrary positions
 * because the defect is per capture and per slot: a carrier that keeps the
 * first capture and drops the rest, or keeps them all at the end, passes any
 * single fixture. The editor assertions are the boundary the design turns on —
 * a record the editor creates has no sink, and a copy does not inherit one.
 */

import { describe, expect, test } from "bun:test";
import fc from "fast-check";
import JSZip from "jszip";

import { propertyConfig } from "../../../../test/property-testing";

import { PRESERVED_SINK_CARRIERS } from "../prosemirror/conversion/preservedSinkCarriers";
import { fromProseDoc, proseDocToBlocks } from "../prosemirror/conversion/fromProseDoc";
import { toProseDoc } from "../prosemirror/conversion/toProseDoc";
import { schema } from "../prosemirror/schema";
import type { Document, PreservedChild, Table } from "../types/document";
import { parseDocx } from "./parser";
import { createEmptyDocx, repackDocx } from "./rezip";

const XML_DECLARATION = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';
const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

/**
 * Self-contained markup the two sinks capture.
 *
 * Each is an empty marker that needs no other part, so a fixture carrying it
 * is schema-valid on its own and the capture is exactly the bytes below.
 *
 * Bookmark markers are modelled beside both tables and rows, so this property
 * covers the remaining verbatim sink rather than asserting that a typed
 * marker must also appear there.
 */
const MARKERS = {
  bookmarkEnd: '<w:bookmarkEnd w:id="41"/>',
  permStart: '<w:permStart w:id="42" w:edGrp="everyone"/>',
  permEnd: '<w:permEnd w:id="42"/>',
  proofErr: '<w:proofErr w:type="spellStart"/>',
} as const satisfies Record<string, string>;

type MarkerName = keyof typeof MARKERS;
const TABLE_MARKERS: readonly MarkerName[] = ["permStart", "proofErr"];
const ROW_MARKERS: readonly MarkerName[] = ["permStart", "permEnd", "proofErr"];

const ROWS = 2;
const CELLS = 2;

const cellXml = (text: string): string =>
  `<w:tc><w:tcPr/><w:p><w:r><w:t>${text}</w:t></w:r></w:p></w:tc>`;

/** One container's children, with the chosen markers spliced in at `index`. */
const withMarkers = (
  modelled: readonly string[],
  markers: readonly MarkerName[],
  index: number,
): string =>
  modelled
    .flatMap((child, at) =>
      at === index ? [...markers.map((name) => MARKERS[name]), child] : child,
    )
    .join("") + (index >= modelled.length ? markers.map((name) => MARKERS[name]).join("") : "");

type Placement = {
  tableMarkers: readonly MarkerName[];
  tableIndex: number;
  rowMarkers: readonly MarkerName[];
  rowIndex: number;
};

const rowXml = ({ rowMarkers, rowIndex }: Placement, row: number): string => {
  const cells = Array.from({ length: CELLS }, (_, cell) => cellXml(`r${row}c${cell}`));
  // Only the first row carries the row-level sink, so the two sinks cannot be
  // confused for each other when the model comes back.
  return `<w:tr>${row === 0 ? withMarkers(cells, rowMarkers, rowIndex) : cells.join("")}</w:tr>`;
};

const documentXml = (placement: Placement): string => {
  const rows = Array.from({ length: ROWS }, (_, row) => rowXml(placement, row));
  return (
    `${XML_DECLARATION}<w:document xmlns:w="${W}"><w:body>` +
    '<w:tbl><w:tblPr/><w:tblGrid><w:gridCol w:w="2400"/><w:gridCol w:w="2400"/></w:tblGrid>' +
    `${withMarkers(rows, placement.tableMarkers, placement.tableIndex)}</w:tbl>` +
    "<w:p><w:r><w:t>tail</w:t></w:r></w:p>" +
    '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/></w:sectPr>' +
    "</w:body></w:document>"
  );
};

const packageFor = async (xml: string): Promise<ArrayBuffer> => {
  const zip = await JSZip.loadAsync(await createEmptyDocx());
  zip.file("word/document.xml", xml);
  return zip.generateAsync({ type: "arraybuffer" });
};

const open = (xml: string): Promise<Document> =>
  packageFor(xml).then((buffer) => parseDocx(buffer, { preloadFonts: false }));

const save = (document: Document): Promise<ArrayBuffer> =>
  repackDocx(document, { updateModifiedDate: false });

const reopen = async (document: Document): Promise<Document> =>
  parseDocx(await save(document), { preloadFonts: false });

const firstTable = (document: Document): Table => {
  const table = document.package.document.content.find((block) => block.type === "table");
  if (table === undefined) {
    throw new Error("the fixture produced no table");
  }
  return table;
};

const expected = (markers: readonly MarkerName[], index: number): PreservedChild[] =>
  markers.map((name) => ({ index, xml: MARKERS[name] }));

/** Where one carried record's sink is read back from, and what belongs there. */
type SinkSite = {
  read: (document: Document) => PreservedChild[] | undefined;
  wrote: (placement: Placement) => PreservedChild[];
};

/**
 * One site per key of `PRESERVED_SINK_CARRIERS`, so the map that claims a
 * record is carried and the assertions that prove it cannot drift apart: a
 * record that gains a sink joins the carrier map, and the map's key set is
 * what this table must be total over.
 */
const SINK_SITES = {
  table: {
    read: (document) => firstTable(document).preserved?.children,
    wrote: ({ tableMarkers, tableIndex }) => expected(tableMarkers, tableIndex),
  },
  tableRow: {
    read: (document) => firstTable(document).rows.at(0)?.preserved?.children,
    wrote: ({ rowMarkers, rowIndex }) => expected(rowMarkers, rowIndex),
  },
} as const satisfies Record<keyof typeof PRESERVED_SINK_CARRIERS, SinkSite>;

const SITES: readonly SinkSite[] = Object.values(SINK_SITES);

const expectSinksIntact = (document: Document, placement: Placement): void => {
  for (const site of SITES) {
    expect(site.read(document)).toEqual(site.wrote(placement));
  }
};

const placementArbitrary = fc.record({
  tableMarkers: fc.subarray([...TABLE_MARKERS], { minLength: 1 }),
  tableIndex: fc.integer({ min: 0, max: ROWS }),
  rowMarkers: fc.subarray([...ROW_MARKERS], { minLength: 1 }),
  rowIndex: fc.integer({ min: 0, max: CELLS }),
});

describe("a container's verbatim sink survives the editor projection", () => {
  test("every marker comes back between the siblings it was written between", async () => {
    await fc.assert(
      fc.asyncProperty(placementArbitrary, async (placement) => {
        const parsed = await open(documentXml(placement));
        // The save leg is the premise: a sink the parser never filled would
        // make the editor assertion below vacuous.
        expectSinksIntact(parsed, placement);
        expectSinksIntact(await reopen(fromProseDoc(toProseDoc(parsed), parsed)), placement);
      }),
      propertyConfig({ numRuns: 20 }),
    );
  }, 180_000);

  test("an edit elsewhere in the document leaves the sink where it was", async () => {
    await fc.assert(
      fc.asyncProperty(placementArbitrary, async (placement) => {
        const parsed = await open(documentXml(placement));
        const projected = toProseDoc(parsed);
        const created = schema.nodes["paragraph"]?.createAndFill();
        if (!created) {
          throw new Error("the schema refused an empty paragraph");
        }
        const edited = schema.node("doc", projected.attrs, [...projected.content.content, created]);

        expectSinksIntact(await reopen(fromProseDoc(edited, parsed)), placement);
      }),
      propertyConfig({ numRuns: 20 }),
    );
  }, 180_000);
});

describe("the sink follows the record", () => {
  const onlyTableMarker: Placement = {
    tableMarkers: ["permStart"],
    tableIndex: ROWS,
    rowMarkers: ["proofErr"],
    rowIndex: CELLS,
  };

  test("a table the editor creates from scratch has no sink", async () => {
    const parsed = await open(documentXml(onlyTableMarker));
    const projected = toProseDoc(parsed);
    const authored = projected.content.content.find((node) => node.type.name === "table");
    if (!authored) {
      throw new Error("the projection produced no table");
    }
    const created = authored.type.createAndFill();
    if (!created) {
      throw new Error("the schema refused an empty table");
    }

    const blocks = proseDocToBlocks(
      schema.node("doc", projected.attrs, [created, ...projected.content.content]),
      parsed.package.document.content,
      parsed.package.styles,
    );
    const [fresh] = blocks;
    expect(fresh?.type).toBe("table");
    expect(fresh?.type === "table" ? fresh.preserved : undefined).toBeUndefined();
  });

  test("a copied table keeps the sink on the copy that comes first", async () => {
    const parsed = await open(documentXml(onlyTableMarker));
    const projection = toProseDoc(parsed);
    const authored = projection.content.content.find((node) => node.type.name === "table");
    if (!authored) {
      throw new Error("the projection produced no table");
    }

    // What a copy produces: two nodes over one attrs object, so both hold the
    // very sink the authored table carried.
    const copies = [
      authored.type.create(authored.attrs, authored.content),
      authored.type.create(authored.attrs, authored.content),
    ];
    const blocks = proseDocToBlocks(
      schema.node("doc", projection.attrs, copies),
      parsed.package.document.content,
      parsed.package.styles,
    );
    const [head, tail] = blocks;

    expect(head?.type === "table" ? head.preserved : undefined).toBeDefined();
    expect(tail?.type === "table" ? tail.preserved : undefined).toBeUndefined();
    // The rows are copies of the same nodes, so the rule has to hold one level
    // down as well: a second copy of a row's markup is markup nobody wrote.
    expect(head?.type === "table" ? head.rows.at(0)?.preserved : undefined).toBeDefined();
    expect(tail?.type === "table" ? tail.rows.at(0)?.preserved : undefined).toBeUndefined();
  });
});
