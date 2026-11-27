/**
 * A rebuilt container must keep the tracked-change record it arrived with.
 *
 * `w:tblGridChange`, `w:numberingChange`, and `w:numPr/w:ins` are revisions:
 * they record the replaced grid, replaced numbering, and inserted numbering
 * properties. Nothing in the editable model derives them, so a serializer that
 * rebuilds their container from the model must carry their records explicitly.
 * In a document under review that is not a fidelity detail; it accepts or
 * discards somebody's edit without telling anyone.
 *
 * Both used to survive only by verbatim replay, which is exactly the path an
 * edited document does not take. The tests therefore force the rebuild:
 * resizing a column makes the captured `w:tblGrid` stop matching the model,
 * and the `w:numberingChange` case is checked both through replay and through
 * a rebuild, because refusing the capture used to be what forced a rebuild
 * that could not write it.
 */

import { describe, expect, setDefaultTimeout, test } from "bun:test";
import fc from "fast-check";

import {
  applyDocumentOp,
  applyDocumentOps,
  DOCUMENT_OP_TYPES,
  OP_STORIES,
  REVISION_DECISIONS,
} from "@stll/docx-core/ops";

import {
  propertyConfig,
  propertyTestTimeout,
} from "../../../../../test/property-testing";
import type { Document, Table } from "../../types/document";
import { parseParagraph } from "../paragraphParser";
import { FOLIO_REVIEW_HISTORY_NAMESPACE } from "../reviewHistoryNamespace";
import { serializeParagraph } from "../serializer/paragraphSerializer";
import { serializeTable } from "../serializer/tableSerializer";
import { parseTable } from "../tableParser";
import { parseXmlDocument, type XmlElement } from "../xmlParser";

setDefaultTimeout(propertyTestTimeout(30_000));

const W_NS =
  'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';

const GRID_CHANGE =
  '<w:tblGridChange w:id="42"><w:tblGrid><w:gridCol w:w="1000"/><w:gridCol w:w="5000"/></w:tblGrid></w:tblGridChange>';

const TABLE_XML =
  `<w:tbl ${W_NS}><w:tblPr/><w:tblGrid><w:gridCol w:w="2400"/><w:gridCol w:w="3600"/>${GRID_CHANGE}</w:tblGrid>` +
  "<w:tr><w:tc><w:p/></w:tc><w:tc><w:p/></w:tc></w:tr></w:tbl>";

const NUMBERING_CHANGE =
  '<w:numberingChange w:id="7" w:author="Reviewer" w:date="2024-01-01T00:00:00Z" w:original="1)."/>';
const NUMBERING_INSERTION =
  '<w:ins w:id="8" w:author="Editor" w:date="2024-01-02T00:00:00Z"/>';

const PARAGRAPH_XML =
  `<w:p ${W_NS}><w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="3"/>` +
  `${NUMBERING_CHANGE}${NUMBERING_INSERTION}</w:numPr></w:pPr>` +
  "<w:r><w:t>item</w:t></w:r></w:p>";

/**
 * The table with every verbatim capture cleared, which is what the survival
 * law's carrier probe does: whatever survives this, the typed model holds.
 */
const modelOnly = (table: Table | undefined): Table => {
  if (!table) {
    throw new Error("the fixture did not parse into a table");
  }
  const {
    sourceXml: _source,
    gridSourceXml: _grid,
    ...formatting
  } = table.formatting ?? {};
  return { ...table, formatting };
};

const parseElement = (xml: string): XmlElement => {
  const node = parseXmlDocument(xml) as XmlElement | null;
  if (!node) {
    throw new Error("the fixture did not parse");
  }
  return node;
};

describe("w:tblGridChange survives a grid the serializer has to rebuild", () => {
  const namespaces = `${W_NS} xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" xmlns:folio="${FOLIO_REVIEW_HISTORY_NAMESPACE}"`;
  const read = (fragment: string): Table => {
    const table = parseTable(
      parseElement(fragment.replace("<w:tbl>", `<w:tbl ${namespaces}>`)),
      new Map(),
      null,
      null,
      null,
      null,
    );
    if (table === undefined) throw new Error("Expected a table.");
    return table;
  };
  const tableOf = (document: Document): Table => {
    const table = document.package.document.content.at(0);
    if (table?.type !== "table") throw new Error("Expected a table.");
    return table;
  };
  const withTable = (table: Table): Document => ({
    package: {
      document: {
        content: [
          table,
          { type: "paragraph", paraId: "00000010", content: [] },
        ],
      },
    },
  });
  const twoCells =
    '<w:tr><w:tc><w:p w14:paraId="00000001"/></w:tc><w:tc><w:p w14:paraId="00000002"/></w:tc></w:tr>';

  test("a resized column keeps the record of the grid it replaced", () => {
    const table = parseTable(
      parseElement(TABLE_XML),
      new Map(),
      null,
      null,
      null,
      null,
    );
    expect(table).toBeDefined();
    if (!table) {
      return;
    }
    expect(table.formatting?.gridChange).toEqual({
      id: 42,
      columnWidths: [1000, 5000],
    });

    // Resizing a column is what makes the captured grid stop matching the
    // model, so the serializer rebuilds it instead of replaying the capture.
    table.columnWidths = [3000, 3000];
    const xml = serializeTable(table, serializeParagraph);

    expect(xml).toContain('<w:gridCol w:w="3000"/><w:gridCol w:w="3000"/>');
    expect(xml).toContain('<w:tblGridChange w:id="42">');
    expect(xml).toContain('<w:gridCol w:w="5000"/>');
  });

  test("an unchanged grid still replays, record and all", () => {
    const table = parseTable(
      parseElement(TABLE_XML),
      new Map(),
      null,
      null,
      null,
      null,
    );
    expect(table).toBeDefined();
    if (!table) {
      return;
    }
    expect(serializeTable(table, serializeParagraph)).toContain(GRID_CHANGE);
  });

  // A `w:tblGridChange` holds a `w:tblGrid` and that grid holds its own
  // `w:gridCol` children: a container nested in one of its own kind. While the
  // snapshot travelled as bytes, a save with nothing verbatim left wrote the
  // outer grid and nothing under the change, which is what the survival
  // census's carrier probe was reading when it called the pair captured.
  test("the snapshot's own grid and columns survive a save the model alone builds", () => {
    fc.assert(
      fc.property(
        fc.array(
          fc.option(fc.integer({ min: 1, max: 20_000 }), { nil: undefined }),
          {
            minLength: 1,
            maxLength: 6,
          },
        ),
        fc.integer({ min: 0, max: 9999 }),
        (snapshotWidths, changeId) => {
          const snapshot = snapshotWidths
            .map((width) =>
              width === undefined
                ? "<w:gridCol/>"
                : `<w:gridCol w:w="${width}"/>`,
            )
            .join("");
          const xml =
            `<w:tbl ${W_NS}><w:tblPr/><w:tblGrid><w:gridCol w:w="2400"/>` +
            `<w:tblGridChange w:id="${changeId}"><w:tblGrid>${snapshot}</w:tblGrid></w:tblGridChange>` +
            "</w:tblGrid><w:tr><w:tc><w:p/></w:tc></w:tr></w:tbl>";

          const table = modelOnly(
            parseTable(parseElement(xml), new Map(), null, null, null, null),
          );
          expect(table.formatting?.gridChange).toEqual({
            id: changeId,
            columnWidths: snapshotWidths,
          });

          const saved = serializeTable(table, serializeParagraph);
          expect(saved).toContain(
            `<w:tblGridChange w:id="${changeId}"><w:tblGrid>${snapshot}</w:tblGrid></w:tblGridChange>`,
          );

          // A fixed point: the second parse reads the same snapshot the first
          // did, so nothing about it depends on the bytes it arrived as. The
          // serializer returns a fragment, and the revision id is resolved
          // against the namespace the prefix is bound to, so the fragment is
          // read back declaring the binding the saved part carries.
          const reparsed = parseTable(
            parseElement(saved.replace("<w:tbl>", `<w:tbl ${W_NS}>`)),
            new Map(),
            null,
            null,
            null,
            null,
          );
          expect(reparsed?.formatting?.gridChange).toEqual(
            table.formatting?.gridChange,
          );
        },
      ),
      propertyConfig({ numRuns: 100 }),
    );
  });

  test("rejecting an imported grid change restores missing widths separately from zero widths", () => {
    const oldGrid = '<w:gridCol/><w:gridCol w:w="0"/><w:gridCol w:w="3400"/>';
    const pending = withTable(
      read(
        '<w:tbl><w:tblPr/><w:tblGrid><w:gridCol w:w="1100"/><w:gridCol w:w="1200"/><w:gridCol w:w="1300"/>' +
          `<w:tblGridChange w:id="42"><w:tblGrid>${oldGrid}</w:tblGrid></w:tblGridChange></w:tblGrid>` +
          '<w:tr><w:tc><w:p w14:paraId="00000001"/></w:tc><w:tc><w:p w14:paraId="00000002"/></w:tc><w:tc><w:p w14:paraId="00000003"/></w:tc></w:tr></w:tbl>',
      ),
    );
    const reject = {
      type: DOCUMENT_OP_TYPES.RESOLVE_REVISION,
      story: OP_STORIES.MAIN,
      revisionIds: [42],
      decision: REVISION_DECISIONS.REJECT,
    } as const;
    const rejected = applyDocumentOp(pending, reject).unwrap();
    expect(
      applyDocumentOps(rejected.document, rejected.inverse).unwrap().document,
    ).toEqual(pending);
    const restored = tableOf(rejected.document);
    expect(restored.columnWidths).toEqual([0, 0, 3400]);
    expect(restored.omittedColumnWidthIndices).toEqual([0]);
    const restoredXml = serializeTable(restored, serializeParagraph);
    expect(restoredXml).toContain(`<w:tblGrid>${oldGrid}</w:tblGrid>`);
    const reopened = read(restoredXml);
    expect(reopened.omittedColumnWidthIndices).toEqual([0]);

    const inserted = applyDocumentOp(withTable(reopened), {
      type: DOCUMENT_OP_TYPES.INSERT_COLUMN,
      story: OP_STORIES.MAIN,
      blockId: "00000001",
      column: 0,
      width: 900,
      newBlockIds: ["00000005"],
    }).unwrap().document;
    const insertedTable = tableOf(inserted);
    expect(insertedTable.columnWidths).toEqual([900, 0, 0, 3400]);
    expect(insertedTable.omittedColumnWidthIndices).toEqual([1]);
    expect(serializeTable(insertedTable, serializeParagraph)).toContain(
      `<w:tblGrid><w:gridCol w:w="900"/>${oldGrid}</w:tblGrid>`,
    );

    const resize = applyDocumentOp(inserted, {
      type: DOCUMENT_OP_TYPES.SET_TABLE_GRID,
      story: OP_STORIES.MAIN,
      blockId: "00000001",
      columnWidths: [500, 600, 700, 800],
      revision: { id: 43, author: "Reviewer", date: "2026-11-27T00:00:00Z" },
      newIds: { revision: [44] },
    }).unwrap().document;
    expect(tableOf(resize).omittedColumnWidthIndices).toBeUndefined();
    expect(tableOf(resize).formatting?.gridChange?.columnWidths).toEqual([
      900,
      undefined,
      0,
      3400,
    ]);
    const rejectedResize = applyDocumentOp(resize, {
      ...reject,
      revisionIds: [43, 44],
    }).unwrap().document;
    expect(rejectedResize).toEqual(inserted);

    const deleted = applyDocumentOp(rejectedResize, {
      type: DOCUMENT_OP_TYPES.DELETE_COLUMN,
      story: OP_STORIES.MAIN,
      blockId: "00000001",
      column: 2,
    }).unwrap().document;
    const deletedTable = tableOf(deleted);
    expect(deletedTable.columnWidths).toEqual([900, 0, 3400]);
    expect(deletedTable.omittedColumnWidthIndices).toEqual([1]);
    expect(serializeTable(deletedTable, serializeParagraph)).toContain(
      '<w:tblGrid><w:gridCol w:w="900"/><w:gridCol/><w:gridCol w:w="3400"/></w:tblGrid>',
    );
  });

  test.each([1, 2, 4])(
    "rejection restores %i declared unmeasured columns across save and reopen",
    (declared) => {
      const priorGrid = "<w:gridCol/>".repeat(declared);
      const original = withTable(
        read(
          '<w:tbl><w:tblPr/><w:tblGrid><w:gridCol w:w="900"/><w:gridCol w:w="1100"/>' +
            `<w:tblGridChange w:id="42"><w:tblGrid>${priorGrid}</w:tblGrid></w:tblGridChange></w:tblGrid>` +
            `${twoCells}</w:tbl>`,
        ),
      );
      const reject = {
        type: DOCUMENT_OP_TYPES.RESOLVE_REVISION,
        story: OP_STORIES.MAIN,
        revisionIds: [42],
        decision: REVISION_DECISIONS.REJECT,
      } as const;
      const result = applyDocumentOp(original, reject).unwrap();
      expect(
        applyDocumentOps(result.document, result.inverse).unwrap().document,
      ).toStrictEqual(original);
      const restored = tableOf(result.document);
      const xml = serializeTable(restored, serializeParagraph);
      expect(xml).toContain(`<w:tblGrid>${priorGrid}</w:tblGrid>`);
      const reopened = read(xml);
      for (const table of [restored, reopened]) {
        expect(table.columnWidths).toBeUndefined();
        expect(table.omittedColumnWidthIndices).toBeUndefined();
        expect(Object.hasOwn(table, "unmeasuredColumnCount")).toBe(
          declared !== 2,
        );
        expect(table.unmeasuredColumnCount).toBe(
          declared === 2 ? undefined : declared,
        );
      }

      const inserted = applyDocumentOp(withTable(reopened), {
        type: DOCUMENT_OP_TYPES.INSERT_COLUMN,
        story: OP_STORIES.MAIN,
        blockId: "00000001",
        column: 0,
        width: 500,
        newBlockIds: ["00000004"],
      }).unwrap().document;
      const expanded = tableOf(inserted);
      expect(expanded.columnWidths).toBeUndefined();
      expect(serializeTable(expanded, serializeParagraph)).toContain(
        `<w:tblGrid>${"<w:gridCol/>".repeat(declared + 1)}</w:tblGrid>`,
      );
    },
  );

  test("a pending merge of an inferred grid reopens without an explicit count", () => {
    const original = withTable(
      read(
        `<w:tbl><w:tblPr/><w:tblGrid><w:gridCol/><w:gridCol/></w:tblGrid>${twoCells}</w:tbl>`,
      ),
    );
    const pending = applyDocumentOp(original, {
      type: DOCUMENT_OP_TYPES.MERGE_CELLS,
      story: OP_STORIES.MAIN,
      blockId: "00000001",
      top: 0,
      bottom: 1,
      left: 0,
      right: 2,
      newBlockIds: [],
      revision: { id: 50, author: "Reviewer", date: "2026-11-27T00:00:00Z" },
      newIds: { revision: [51, 52, 53] },
    }).unwrap().document;
    const reopened = read(serializeTable(tableOf(pending), serializeParagraph));
    expect(reopened.formatting?.gridChange?.geometryChange?.type).toBe(
      "mergeCells",
    );
    expect(Object.hasOwn(reopened, "unmeasuredColumnCount")).toBe(false);
  });
});

describe("numbering revision records survive a w:pPr the serializer has to rebuild", () => {
  test("both records reach the model", () => {
    const paragraph = parseParagraph(
      parseElement(PARAGRAPH_XML),
      new Map(),
      null,
      null,
    );
    expect(paragraph.formatting?.numberingChangeXml).toContain(
      'w:original="1)."',
    );
    expect(paragraph.formatting?.numberingInsertionXml).toContain(
      'w:author="Editor"',
    );
  });

  test("a rebuilt w:numPr writes both records back in schema order", () => {
    const paragraph = parseParagraph(
      parseElement(PARAGRAPH_XML),
      new Map(),
      null,
      null,
    );
    // Changing the numbering is what a reviewer's next edit does, and it is
    // what stops the captured `w:pPr` from being replayed.
    if (paragraph.formatting?.numPr) {
      paragraph.formatting.numPr = { kind: "reference", numId: 9 };
    }
    const xml = serializeParagraph(paragraph);
    expect(xml).toContain('<w:numId w:val="9"/>');
    expect(xml).toContain('w:author="Reviewer"');
    expect(xml).toContain('w:original="1)."');
    expect(xml).toContain(`${NUMBERING_CHANGE}${NUMBERING_INSERTION}`);
  });

  test("an insertion record keeps an otherwise empty w:numPr", () => {
    const xml =
      `<w:p ${W_NS}><w:pPr><w:numPr>${NUMBERING_INSERTION}</w:numPr></w:pPr>` +
      "<w:r><w:t>item</w:t></w:r></w:p>";
    const paragraph = parseParagraph(parseElement(xml), new Map(), null, null);

    expect(paragraph.formatting?.numPr).toBeUndefined();
    expect(serializeParagraph(paragraph)).toContain(
      `<w:numPr>${NUMBERING_INSERTION}</w:numPr>`,
    );
  });
});
