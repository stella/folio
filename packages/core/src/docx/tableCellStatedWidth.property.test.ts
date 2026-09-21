/**
 * A cell states its own preferred width, or states none.
 *
 * `TableCellAttrs.width` is the width the cell *renders* at, which the table
 * resolves from its grid when the cell declares no `w:tcW`, and the way back
 * wrote it into `w:tcPr` unconditionally: opening a document and saving it
 * again gave every cell in every table a preferred width its author never
 * wrote. `TableCellAttrs._authoredWidth` is the record of what the cell
 * states, the way `_resolvedBorders` and `_resolvedMargins` keep an inherited
 * border and margin apart from a stated one, and the save leg writes `w:tcW`
 * from it.
 *
 * A command that moves a cell's width states one, so the properties below
 * cover both directions: a width the source never stated stays unstated, and a
 * width a resize or a merge states is written.
 */

import {
  afterAll,
  afterEach,
  beforeAll,
  describe,
  expect,
  setDefaultTimeout,
  test,
} from "bun:test";
import { GlobalRegistrator } from "@happy-dom/global-registrator";
import fc from "fast-check";
import { EditorState } from "prosemirror-state";
import { EditorView } from "prosemirror-view";

import { propertyConfig, propertyTestTimeout } from "../../../../test/property-testing";

import { fromProseDoc } from "../prosemirror/conversion/fromProseDoc";
import { toProseDoc } from "../prosemirror/conversion/toProseDoc";
import { commitColumnResize } from "../prosemirror/tableResize";
import type { Table } from "../types/document";
import { serializeParagraph } from "./serializer/paragraphSerializer";
import { serializeTable } from "./serializer/tableSerializer";
import { parseTable } from "./tableParser";
import { parseXmlDocument, type XmlElement } from "./xmlParser";

setDefaultTimeout(propertyTestTimeout(30_000));

const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

/** What a column's cell states, `null` for a cell that states no width. */
type StatedWidth = { readonly value: number; readonly type: "dxa" | "pct" } | null;

const statedWidth = fc.option(
  fc.record({
    value: fc.integer({ min: 1, max: 5000 }),
    type: fc.constantFrom("dxa" as const, "pct" as const),
  }),
  { nil: null },
);

const cellXml = (stated: StatedWidth, index: number): string => {
  const width = stated === null ? "" : `<w:tcW w:w="${stated.value}" w:type="${stated.type}"/>`;
  // A property set on every cell, so the ones that state no width still reach
  // the serializer with a `w:tcPr` to acquire one in.
  return (
    `<w:tc><w:tcPr>${width}<w:vAlign w:val="center"/></w:tcPr>` +
    `<w:p><w:r><w:t>c${index}</w:t></w:r></w:p></w:tc>`
  );
};

const tableXml = (columns: readonly StatedWidth[]): string =>
  `<w:tbl xmlns:w="${W}"><w:tblPr><w:tblW w:w="9600" w:type="dxa"/></w:tblPr><w:tblGrid>` +
  columns.map(() => '<w:gridCol w:w="2400"/>').join("") +
  "</w:tblGrid><w:tr>" +
  columns.map((stated, index) => cellXml(stated, index)).join("") +
  "</w:tr></w:tbl>";

const parsed = (xml: string): Table => {
  const root = parseXmlDocument(xml) as XmlElement | null;
  if (!root) {
    throw new Error("the table fixture did not parse");
  }
  const table = parseTable(root, null, null, null, null, null);
  if (!table) {
    throw new Error("the table fixture parsed to nothing");
  }
  return table;
};

const documentOf = (table: Table) =>
  ({ package: { document: { content: [table], finalSectionProperties: {} } } }) as never;

/** The `w:tcW` each cell carries after the save, in column order. */
const statedWidthsOf = (saved: string): (string | null)[] =>
  saved
    .split("<w:tc>")
    .slice(1)
    .map((cell) => /<w:tcW [^/]*\/>/u.exec(cell)?.[0] ?? null);

const throughEditor = (columns: readonly StatedWidth[]): string => {
  const document = documentOf(parsed(tableXml(columns)));
  const projected = fromProseDoc(toProseDoc(document), document).package.document
    .content[0] as Table;
  return serializeTable(projected, serializeParagraph);
};

const views: EditorView[] = [];

const mount = (columns: readonly StatedWidth[]): EditorView => {
  const view = new EditorView(document.body.appendChild(document.createElement("div")), {
    state: EditorState.create({ doc: toProseDoc(documentOf(parsed(tableXml(columns)))) }),
  });
  views.push(view);
  return view;
};

beforeAll(() => GlobalRegistrator.register());

afterEach(() => {
  for (const view of views.splice(0)) {
    const host = view.dom.parentElement;
    view.destroy();
    host?.remove();
  }
  document.body.replaceChildren();
});

afterAll(() => GlobalRegistrator.unregister());

describe("a cell's preferred width is the one it states", () => {
  test("a save writes `w:tcW` on exactly the cells that stated one", () => {
    fc.assert(
      fc.property(fc.array(statedWidth, { minLength: 1, maxLength: 4 }), (columns) => {
        expect(statedWidthsOf(throughEditor(columns)).map((width) => width !== null)).toEqual(
          columns.map((stated) => stated !== null),
        );
      }),
      propertyConfig({ numRuns: 100 }),
    );
  });

  test("a stated width comes back with its own value and unit", () => {
    fc.assert(
      fc.property(fc.array(statedWidth, { minLength: 1, maxLength: 4 }), (columns) => {
        expect(statedWidthsOf(throughEditor(columns))).toEqual(
          columns.map((stated) =>
            stated === null ? null : `<w:tcW w:w="${stated.value}" w:type="${stated.type}"/>`,
          ),
        );
      }),
      propertyConfig({ numRuns: 100 }),
    );
  });

  test("a column resize states a width on the columns it moved, and only those", () => {
    const columns: StatedWidth[] = [null, null, null];
    const view = mount(columns);
    const document_ = documentOf(parsed(tableXml(columns)));

    commitColumnResize(view, { pmStart: 0, colIdx: 0, newLeft: 3000, newRight: 1800 });

    const projected = fromProseDoc(view.state.doc, document_).package.document.content[0] as Table;
    expect(statedWidthsOf(serializeTable(projected, serializeParagraph))).toEqual([
      '<w:tcW w:w="3000" w:type="dxa"/>',
      '<w:tcW w:w="1800" w:type="dxa"/>',
      null,
    ]);
  });
});
