/**
 * Accept/reject of tracked property changes (w:pPrChange, w:sectPrChange,
 * w:tblPrChange, w:trPrChange, w:tcPrChange) against the REAL editor schema,
 * built via toProseDoc so the attr names under test are the ones the parser
 * actually produces — and round-tripped through fromProseDoc so the
 * serializer provably no longer re-emits a resolved change.
 */

import { describe, expect, test } from "bun:test";
import { EditorState } from "prosemirror-state";
import type { Transaction } from "prosemirror-state";

import type {
  Document,
  Paragraph,
  Table,
  TableCellPropertyChange,
  TablePropertyChange,
  TableRowPropertyChange,
} from "../../types/document";
import { fromProseDoc } from "../conversion/fromProseDoc";
import { toProseDoc } from "../conversion/toProseDoc";
import { acceptChange, rejectChange } from "./comments";

const CHANGE_INFO = { id: 42, author: "Reviewer", date: "2026-05-15T12:00:00Z" };

const makeDocument = (content: (Paragraph | Table)[]): Document =>
  ({
    package: {
      document: {
        content,
        finalSectionProperties: {},
      },
    },
  }) as never;

const makeState = (content: (Paragraph | Table)[]): EditorState =>
  EditorState.create({ doc: toProseDoc(makeDocument(content)) });

const dispatcher = (state: EditorState) => {
  const view = {
    state,
    dispatch(tr: Transaction) {
      view.state = view.state.apply(tr);
    },
  };
  return view;
};

const paragraphText = (text: string): Paragraph["content"] => [
  { type: "run", content: [{ type: "text", text }] },
];

describe("pPrChange accept/reject (real schema)", () => {
  const makeParagraph = (): Paragraph => ({
    type: "paragraph",
    // The tracked change set alignment AND indentLeft; the stored old pPr
    // only had indentLeft — so alignment was ADDED, indentLeft was CHANGED.
    formatting: { alignment: "center", indentLeft: 720 },
    propertyChanges: [
      {
        type: "paragraphPropertyChange",
        info: CHANGE_INFO,
        previousFormatting: { indentLeft: 1440 },
      },
    ],
    content: paragraphText("body"),
  });

  test("reject restores the old pPr wholesale and the save path drops the pPrChange", () => {
    const view = dispatcher(makeState([makeParagraph()]));

    expect(rejectChange(0, view.state.doc.content.size)(view.state, view.dispatch)).toBe(true);

    const attrs = view.state.doc.child(0).attrs;
    // Changed key: restored to the stored old value.
    expect(attrs["indentLeft"]).toBe(1440);
    // Added key: reset — the old pPr did not carry it.
    expect(attrs["alignment"]).toBeNull();
    expect(attrs["_propertyChanges"]).toBeNull();

    const roundtripped = fromProseDoc(view.state.doc).package.document.content[0] as Paragraph;
    expect(roundtripped.propertyChanges).toBeUndefined();
    expect(roundtripped.formatting?.indentLeft).toBe(1440);
    expect(roundtripped.formatting?.alignment).toBeUndefined();
  });

  test("accept keeps the live pPr and only clears the record", () => {
    const view = dispatcher(makeState([makeParagraph()]));

    expect(acceptChange(0, view.state.doc.content.size)(view.state, view.dispatch)).toBe(true);

    const attrs = view.state.doc.child(0).attrs;
    expect(attrs["alignment"]).toBe("center");
    expect(attrs["indentLeft"]).toBe(720);
    expect(attrs["_propertyChanges"]).toBeNull();

    const roundtripped = fromProseDoc(view.state.doc).package.document.content[0] as Paragraph;
    expect(roundtripped.propertyChanges).toBeUndefined();
    expect(roundtripped.formatting?.alignment).toBe("center");
    expect(roundtripped.formatting?.indentLeft).toBe(720);
  });

  test("reject is a no-op when the paragraph carries no change records", () => {
    const plain: Paragraph = {
      type: "paragraph",
      formatting: { alignment: "center" },
      content: paragraphText("body"),
    };
    const view = dispatcher(makeState([plain]));
    const before = view.state.doc;

    expect(rejectChange(0, view.state.doc.content.size)(view.state, view.dispatch)).toBe(true);

    expect(view.state.doc.eq(before)).toBe(true);
  });
});

describe("sectPrChange accept/reject (real schema)", () => {
  const makeSectionParagraph = (): Paragraph => ({
    type: "paragraph",
    content: paragraphText("section end"),
    sectionProperties: {
      sectionStart: "nextPage",
      pageWidth: 12_240,
      pageHeight: 15_840,
      marginTop: 1440,
      headerReferences: [{ type: "default", relationshipId: "rId7" }],
      propertyChanges: [
        {
          type: "sectionPropertyChange",
          info: CHANGE_INFO,
          previousProperties: {
            sectionStart: "continuous",
            pageWidth: 11_906,
            pageHeight: 16_838,
          },
        },
      ],
    },
  });

  test("reject restores the old sectPr, keeps live header references, updates sectionBreakType", () => {
    const view = dispatcher(makeState([makeSectionParagraph()]));

    expect(rejectChange(0, view.state.doc.content.size)(view.state, view.dispatch)).toBe(true);

    const attrs = view.state.doc.child(0).attrs;
    expect(attrs["_sectionProperties"]).toEqual({
      sectionStart: "continuous",
      pageWidth: 11_906,
      pageHeight: 16_838,
      headerReferences: [{ type: "default", relationshipId: "rId7" }],
    });
    expect(attrs["sectionBreakType"]).toBe("continuous");

    const roundtripped = fromProseDoc(view.state.doc).package.document.content[0] as Paragraph;
    expect(roundtripped.sectionProperties?.propertyChanges).toBeUndefined();
    expect(roundtripped.sectionProperties?.pageWidth).toBe(11_906);
    expect(roundtripped.sectionProperties?.headerReferences).toEqual([
      { type: "default", relationshipId: "rId7" },
    ]);
  });

  test("accept keeps the live sectPr and clears the record", () => {
    const view = dispatcher(makeState([makeSectionParagraph()]));

    expect(acceptChange(0, view.state.doc.content.size)(view.state, view.dispatch)).toBe(true);

    const attrs = view.state.doc.child(0).attrs;
    expect(attrs["_sectionProperties"]).toEqual({
      sectionStart: "nextPage",
      pageWidth: 12_240,
      pageHeight: 15_840,
      marginTop: 1440,
      headerReferences: [{ type: "default", relationshipId: "rId7" }],
    });
    expect(attrs["sectionBreakType"]).toBe("nextPage");

    const roundtripped = fromProseDoc(view.state.doc).package.document.content[0] as Paragraph;
    expect(roundtripped.sectionProperties?.propertyChanges).toBeUndefined();
    expect(roundtripped.sectionProperties?.pageWidth).toBe(12_240);
  });
});

describe("table property-change accept/reject (real schema)", () => {
  const tblChange: TablePropertyChange = {
    type: "tablePropertyChange",
    info: CHANGE_INFO,
    // The old tblPr had a narrower width and no justification (the change
    // ADDED the centering).
    previousFormatting: { width: { value: 4000, type: "pct" } },
  };
  const trChange: TableRowPropertyChange = {
    type: "tableRowPropertyChange",
    info: CHANGE_INFO,
    previousFormatting: { height: { value: 300, type: "dxa" }, heightRule: "atLeast" },
  };
  const tcChange: TableCellPropertyChange = {
    type: "tableCellPropertyChange",
    info: CHANGE_INFO,
    previousFormatting: { shading: { fill: { rgb: "00FF00" } } },
  };

  const makeTable = (): Table => ({
    type: "table",
    formatting: { width: { value: 5000, type: "pct" }, justification: "center" },
    propertyChanges: [tblChange],
    columnWidths: [4680],
    rows: [
      {
        type: "tableRow",
        formatting: { height: { value: 500, type: "dxa" }, heightRule: "exact" },
        propertyChanges: [trChange],
        cells: [
          {
            type: "tableCell",
            formatting: { shading: { fill: { rgb: "FF0000" } } },
            propertyChanges: [tcChange],
            content: [{ type: "paragraph", content: paragraphText("cell") }],
          },
        ],
      },
    ],
  });

  test("toProseDoc/fromProseDoc round-trips unresolved table property changes", () => {
    const state = makeState([makeTable()]);
    const roundtripped = fromProseDoc(state.doc).package.document.content[0] as Table;

    expect(roundtripped.propertyChanges).toEqual([tblChange]);
    expect(roundtripped.rows[0]?.propertyChanges).toEqual([trChange]);
    expect(roundtripped.rows[0]?.cells[0]?.propertyChanges).toEqual([tcChange]);
  });

  test("reject restores previous table/row/cell formatting and drops the records", () => {
    const view = dispatcher(makeState([makeTable()]));

    expect(rejectChange(0, view.state.doc.content.size)(view.state, view.dispatch)).toBe(true);

    const table = view.state.doc.child(0);
    expect(table.attrs["width"]).toBe(4000);
    expect(table.attrs["widthType"]).toBe("pct");
    // The change ADDED the justification — reject resets it.
    expect(table.attrs["justification"]).toBeNull();
    expect(table.attrs["tblPrChange"]).toBeNull();

    const row = table.child(0);
    expect(row.attrs["height"]).toBe(300);
    expect(row.attrs["heightRule"]).toBe("atLeast");
    expect(row.attrs["trPrChange"]).toBeNull();

    const cell = row.child(0);
    expect(cell.attrs["backgroundColor"]).toBe("00FF00");
    expect(cell.attrs["tcPrChange"]).toBeNull();

    const roundtripped = fromProseDoc(view.state.doc).package.document.content[0] as Table;
    expect(roundtripped.propertyChanges).toBeUndefined();
    expect(roundtripped.formatting?.width).toEqual({ value: 4000, type: "pct" });
    expect(roundtripped.formatting?.justification).toBeUndefined();
    expect(roundtripped.rows[0]?.propertyChanges).toBeUndefined();
    expect(roundtripped.rows[0]?.formatting?.height).toEqual({ value: 300, type: "dxa" });
    expect(roundtripped.rows[0]?.cells[0]?.propertyChanges).toBeUndefined();
    expect(roundtripped.rows[0]?.cells[0]?.formatting?.shading).toEqual({
      fill: { rgb: "00FF00" },
    });
  });

  test("accept keeps current table/row/cell formatting and drops the records", () => {
    const view = dispatcher(makeState([makeTable()]));

    expect(acceptChange(0, view.state.doc.content.size)(view.state, view.dispatch)).toBe(true);

    const table = view.state.doc.child(0);
    expect(table.attrs["width"]).toBe(5000);
    expect(table.attrs["justification"]).toBe("center");
    expect(table.attrs["tblPrChange"]).toBeNull();

    const row = table.child(0);
    expect(row.attrs["height"]).toBe(500);
    expect(row.attrs["heightRule"]).toBe("exact");
    expect(row.attrs["trPrChange"]).toBeNull();

    const cell = row.child(0);
    expect(cell.attrs["backgroundColor"]).toBe("FF0000");
    expect(cell.attrs["tcPrChange"]).toBeNull();

    const roundtripped = fromProseDoc(view.state.doc).package.document.content[0] as Table;
    expect(roundtripped.propertyChanges).toBeUndefined();
    expect(roundtripped.formatting?.justification).toBe("center");
    expect(roundtripped.rows[0]?.propertyChanges).toBeUndefined();
    expect(roundtripped.rows[0]?.cells[0]?.propertyChanges).toBeUndefined();
  });
});
