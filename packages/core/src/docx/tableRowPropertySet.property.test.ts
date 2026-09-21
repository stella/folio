/**
 * A row's property set survives a rebuild, over its declared children.
 *
 * `w:trPr` was read by name: ten children had an `if` and the other five had
 * nothing, so `w:cnfStyle`, `w:divId` and `w:tblCellSpacing` went on every
 * save, and so did a `w:trHeight` whose value the reader refuses. The set is
 * now dispatched, and the universe here is the generated declared-child list
 * rather than a list somebody kept: `Record<DeclaredChild<"row-properties">,
 * string>` is total, so a child the schema gains cannot reach this test
 * without a sample.
 *
 * `CT_TrPrBase` is a repeated choice, so the row's properties have no order a
 * consumer enforces; `CT_TrPr` closes a sequence over it, so the structural
 * revision and `w:trPrChange` do have to come last. The order assertion is
 * about that, and about the sink landing a capture between the same two
 * neighbours it was read between.
 */

import { describe, expect, setDefaultTimeout, test } from "bun:test";
import fc from "fast-check";

import { propertyConfig, propertyTestTimeout } from "../../../../test/property-testing";

import { fromProseDoc } from "../prosemirror/conversion/fromProseDoc";
import { toProseDoc } from "../prosemirror/conversion/toProseDoc";
import type { Table } from "../types/document";
import { CONTAINER_CHILDREN, type DeclaredChild } from "./containerChildren.gen";
import { serializeParagraph } from "./serializer/paragraphSerializer";
import { serializeTable } from "./serializer/tableSerializer";
import { parseTable } from "./tableParser";
import { parseXmlDocument, type XmlElement } from "./xmlParser";

setDefaultTimeout(propertyTestTimeout(30_000));

const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

const DECLARED = CONTAINER_CHILDREN["row-properties"];

/** One authored instance per declared child, each stating something. */
const SAMPLES = {
  cnfStyle: '<w:cnfStyle w:val="100000000000"/>',
  divId: '<w:divId w:val="7"/>',
  gridBefore: '<w:gridBefore w:val="2"/>',
  gridAfter: '<w:gridAfter w:val="1"/>',
  wBefore: '<w:wBefore w:w="900" w:type="dxa"/>',
  wAfter: '<w:wAfter w:w="450" w:type="dxa"/>',
  cantSplit: "<w:cantSplit/>",
  trHeight: '<w:trHeight w:val="480" w:hRule="atLeast"/>',
  tblHeader: "<w:tblHeader/>",
  tblCellSpacing: '<w:tblCellSpacing w:w="15" w:type="dxa"/>',
  jc: '<w:jc w:val="center"/>',
  hidden: "<w:hidden/>",
  ins: '<w:ins w:id="4" w:author="Reviewer" w:date="2026-05-15T12:00:00Z"/>',
  del: '<w:del w:id="5" w:author="Reviewer" w:date="2026-05-15T12:00:00Z"/>',
  trPrChange:
    '<w:trPrChange w:id="6" w:author="Reviewer" w:date="2026-05-15T12:00:00Z">' +
    "<w:trPr><w:tblHeader/></w:trPr></w:trPrChange>",
} as const satisfies Record<DeclaredChild<"row-properties">, string>;

/**
 * `w:ins` and `w:del` are one slot in the model: a row is inserted or deleted,
 * not both, and the parser takes the insertion. So the order assertion authors
 * every child but the deletion, rather than asking the model to hold two
 * structural revisions it has no place for.
 */
const ORDERABLE = DECLARED.filter((name) => name !== "del");

/**
 * Children the reader states nothing about, one per way of stating nothing.
 *
 * An element with no attributes at all, one whose value the reader's
 * enumeration does not admit, one whose value it normalises away, and one from
 * a namespace the content model does not name. None can be decided by a map
 * keyed on the child's name.
 */
const UNREAD_CHILDREN = [
  "<w:cnfStyle/>",
  '<w:jc w:val="end"/>',
  '<w:trHeight w:val="0" w:hRule="atLeast"/>',
  '<w:gridBefore w:val="0"/>',
  '<x:hint xmlns:x="urn:example:vendor" x:kind="row"/>',
] as const;

const CELL = "<w:tc><w:p><w:r><w:t>cell</w:t></w:r></w:p></w:tc>";

const tableXml = (rowProperties: string): string =>
  `<w:tbl xmlns:w="${W}"><w:tblPr/><w:tblGrid><w:gridCol w:w="2400"/></w:tblGrid>` +
  `<w:tr><w:trPr>${rowProperties}</w:trPr>${CELL}</w:tr></w:tbl>`;

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

/**
 * The row's own capture cleared, so the serializer runs.
 *
 * `TableRowFormatting.sourceXml` replays the authored element whenever the
 * model still agrees with it, so a round trip that keeps it exercises the
 * capture machinery instead of the serializer — the same forcing the survival
 * law applies.
 */
const withoutReplay = (table: Table): Table => ({
  ...table,
  rows: table.rows.map((row) => {
    if (!row.formatting) {
      return row;
    }
    const { sourceXml: _source, ...formatting } = row.formatting;
    return { ...row, formatting };
  }),
});

const rebuild = (rowProperties: string): string =>
  serializeTable(withoutReplay(parsed(tableXml(rowProperties))), serializeParagraph);

/** The same table after a no-op pass through the editor's document model. */
const throughEditor = (rowProperties: string): string => {
  const table = parsed(tableXml(rowProperties));
  const document = {
    package: { document: { content: [table], finalSectionProperties: {} } },
  } as never;
  const projected = fromProseDoc(toProseDoc(document), document).package.document
    .content[0] as Table;
  return serializeTable(withoutReplay(projected), serializeParagraph);
};

/** The outer `w:trPr`'s own children: a `w:trPrChange` nests a second one. */
const rowPropertiesOf = (saved: string): string =>
  saved.slice(saved.indexOf("<w:trPr>") + "<w:trPr>".length, saved.lastIndexOf("</w:trPr>"));

describe("a row's property set survives a rebuild", () => {
  test("every declared child comes back, and the save after it is a fixed point", () => {
    fc.assert(
      fc.property(fc.constantFrom(...DECLARED), (name) => {
        const saved = rebuild(SAMPLES[name]);

        expect(saved).toContain(`<w:${name}`);
        expect(rebuild(rowPropertiesOf(saved))).toBe(saved);
      }),
      propertyConfig({ numRuns: 100 }),
    );
  });

  test("a child the reader takes nothing from keeps its bytes", () => {
    fc.assert(
      fc.property(fc.constantFrom(...UNREAD_CHILDREN), (child) => {
        const saved = rebuild(`<w:tblHeader/>${child}`);

        expect(saved).toContain(child);
        expect(rebuild(rowPropertiesOf(saved))).toBe(saved);
      }),
      propertyConfig({ numRuns: 100 }),
    );
  });

  test("the children come back in the order the content model declares", () => {
    // Authored backwards: a serializer that wrote them in the order it read
    // them, or in the order of its own statements, would put `w:trPrChange`
    // among the properties it records a change to.
    const saved = rebuild(
      [...ORDERABLE]
        .reverse()
        .map((name) => SAMPLES[name])
        .join(""),
    );

    const positions = ORDERABLE.map((name) => saved.indexOf(`<w:${name}`));
    expect(positions.every((at) => at > -1)).toBe(true);
    expect(positions).toEqual([...positions].sort((left, right) => left - right));
  });

  test("a capture comes back between the same two modelled properties", () => {
    const saved = rebuild('<w:gridBefore w:val="2"/><w:trHeight/><w:tblHeader/>');

    expect(saved.indexOf("<w:gridBefore")).toBeLessThan(saved.indexOf("<w:trHeight/>"));
    expect(saved.indexOf("<w:trHeight/>")).toBeLessThan(saved.indexOf("<w:tblHeader/>"));
  });

  test("every declared child survives the editor projection", () => {
    // `TableRowAttrs._originalFormatting` carries the whole record through
    // ProseMirror, so the sink rides it; the assertion is that the way back
    // does not rebuild the element from the handful of attrs the editor
    // surfaces.
    fc.assert(
      fc.property(fc.constantFrom(...DECLARED), (name) => {
        expect(rowPropertiesOf(throughEditor(SAMPLES[name]))).toBe(
          rowPropertiesOf(rebuild(SAMPLES[name])),
        );
      }),
      propertyConfig({ numRuns: 100 }),
    );
  });

  test("an empty element survives, because its presence is the value", () => {
    // `w:trPr` is optional on `CT_Row`, so a row that wrote an empty one said
    // something an absent element does not. Reading it as "no properties"
    // deleted it on save.
    expect(rebuild("")).toContain("<w:trPr/>");
    expect(throughEditor("")).toContain("<w:trPr/>");
  });

  test("a row that never wrote one still writes none", () => {
    const table = parsed(
      `<w:tbl xmlns:w="${W}"><w:tblPr/><w:tblGrid><w:gridCol w:w="2400"/></w:tblGrid>` +
        `<w:tr>${CELL}</w:tr></w:tbl>`,
    );

    expect(serializeTable(withoutReplay(table), serializeParagraph)).not.toContain("<w:trPr");
  });

  test("a property change survives a row that states no properties of its own", () => {
    const saved = rebuild(SAMPLES.trPrChange);

    expect(saved).toContain('<w:trPrChange w:id="6" w:author="Reviewer"');
    expect(saved).toContain("<w:tblHeader/>");
  });
});
