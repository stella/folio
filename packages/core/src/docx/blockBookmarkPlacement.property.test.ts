/**
 * A bookmark marker written at block level keeps its block-level place.
 *
 * `CT_Body`, `CT_Tc`, `CT_SdtContentBlock`, `CT_Row` and `CT_Tbl` each declare
 * `w:bookmarkStart` and `w:bookmarkEnd` beside their own children, and Word
 * writes them there: a bookmark that selects whole rows opens as a child of
 * `w:tr`, one that selects a whole table closes as a child of `w:tbl`, and
 * `_GoBack` lands on `w:body` or on a block content control's `w:sdtContent`.
 *
 * A marker re-anchored into a neighbouring paragraph still saves, so a probe
 * that asks whether the element is in the part says nothing. What changes is
 * the range: a bookmark that spanned a row comes back inside one cell, and a
 * `REF` field or a link that resolves it then covers the wrong text. So the
 * assertion is the marker's **parent element and its ordinal among that
 * parent's children**, both through the save and through the editor.
 */

import { describe, expect, setDefaultTimeout, test } from "bun:test";
import fc from "fast-check";
import JSZip from "jszip";

import { propertyConfig, propertyTestTimeout } from "../../../../test/property-testing";
import { fromProseDoc } from "../prosemirror/conversion/fromProseDoc";
import { toProseDoc } from "../prosemirror/conversion/toProseDoc";

import { CONTAINER_CHILDREN, type DispatchedContainer } from "./containerChildren.gen";
import { parseDocx } from "./parser";
import { createEmptyDocx, repackDocx } from "./rezip";
import { getChildElements, getLocalName, parseXmlDocument, type XmlElement } from "./xmlParser";

setDefaultTimeout(propertyTestTimeout(30_000));

const XML_DECLARATION = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';
const W_NAMESPACE = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

/**
 * Every block-level container the dispatcher covers that declares a bookmark.
 *
 * Derived rather than listed: a container that starts declaring
 * `w:bookmarkStart` joins this union, and the totality check below then fails
 * until a placement exercises it. The inline containers are excluded by name
 * because a marker inside a paragraph is the case that already worked.
 */
type BlockLevelContainer = Extract<
  DispatchedContainer,
  "block-content" | "row-content" | "table-content"
>;

const BLOCK_LEVEL_CONTAINERS = [
  "block-content",
  "row-content",
  "table-content",
] as const satisfies readonly BlockLevelContainer[];

const paragraph = (text: string): string => `<w:p><w:r><w:t>${text}</w:t></w:r></w:p>`;

const start = (id: number, name: string): string =>
  `<w:bookmarkStart w:id="${id}" w:name="${name}"/>`;
const end = (id: number): string => `<w:bookmarkEnd w:id="${id}"/>`;

/**
 * One body, and where in it the pair is expected to come back.
 *
 * `parent` is the element the marker is a direct child of, and `ordinal`
 * counts that parent's element children before it — the two facts that say
 * the bookmark still covers what it covered.
 */
type Placement = {
  /** The dispatched container whose declared child this marker is. */
  readonly container: BlockLevelContainer;
  readonly body: (id: number) => string;
  readonly parent: string;
  readonly startOrdinal: number;
  readonly endOrdinal: number;
};

/**
 * A cell holding two paragraphs, so a marker between them has an ordinal that
 * only a block-level position can produce.
 */
const cell = (inner: string): string => `<w:tc><w:tcPr/>${inner}</w:tc>`;

const grid = `<w:tblPr/><w:tblGrid><w:gridCol w:w="2000"/><w:gridCol w:w="2000"/></w:tblGrid>`;

const PLACEMENTS = {
  /** `_GoBack` and a table-of-contents range: a body's own children. */
  body: {
    container: "block-content",
    body: (id) => `${paragraph("one")}${start(id, `b${id}`)}${paragraph("two")}${end(id)}`,
    parent: "body",
    startOrdinal: 1,
    endOrdinal: 3,
  },
  /** A block content control whose range opens before the content it binds. */
  sdtContentBlock: {
    container: "block-content",
    body: (id) =>
      `<w:sdt><w:sdtPr/><w:sdtContent>` +
      `${start(id, `b${id}`)}${paragraph("one")}${end(id)}${paragraph("two")}` +
      `</w:sdtContent></w:sdt>`,
    parent: "sdtContent",
    startOrdinal: 0,
    endOrdinal: 2,
  },
  /** A form field's range, opened on the cell rather than in its paragraph. */
  cell: {
    container: "block-content",
    body: (id) =>
      `<w:tbl>${grid}<w:tr>` +
      cell(`${start(id, `b${id}`)}${paragraph("a")}${end(id)}${paragraph("b")}`) +
      cell(paragraph("c")) +
      `</w:tr></w:tbl>`,
    parent: "tc",
    startOrdinal: 0,
    endOrdinal: 2,
  },
  /** A bookmark that selects whole rows: both halves are children of `w:tr`. */
  row: {
    container: "row-content",
    body: (id) =>
      `<w:tbl>${grid}<w:tr>` +
      `${start(id, `b${id}`)}${cell(paragraph("a"))}${cell(paragraph("b"))}${end(id)}` +
      `</w:tr></w:tbl>`,
    parent: "tr",
    startOrdinal: 0,
    endOrdinal: 3,
  },
  /** A bookmark that selects a whole table: both halves are children of `w:tbl`. */
  table: {
    container: "table-content",
    body: (id) =>
      `<w:tbl>${grid}${start(id, `b${id}`)}` +
      `<w:tr>${cell(paragraph("a"))}${cell(paragraph("b"))}</w:tr>` +
      `${end(id)}</w:tbl>`,
    parent: "tbl",
    startOrdinal: 0,
    endOrdinal: 2,
  },
} as const satisfies Record<string, Placement>;

type PlacementKey = keyof typeof PLACEMENTS;

const PLACEMENT_KEYS = Object.keys(PLACEMENTS) as PlacementKey[];

const buildDocx = async (body: string): Promise<ArrayBuffer> => {
  const zip = await JSZip.loadAsync(await createEmptyDocx());
  zip.file(
    "word/document.xml",
    `${XML_DECLARATION}<w:document xmlns:w="${W_NAMESPACE}"><w:body>${body}<w:sectPr/></w:body></w:document>`,
  );
  return zip.generateAsync({ type: "arraybuffer" });
};

type Anchor = {
  /** Local name of the element the marker is a direct child of. */
  parent: string;
  /** Content children of that parent before the marker. */
  ordinal: number;
};

/**
 * Children that precede all content in their container's model.
 *
 * `w:tcPr`, `w:tblPr` and their siblings are written from the record rather
 * than from the child list, and whether a save writes an empty one says
 * nothing about where a bookmark stands. The ordinal counts what the content
 * model orders the marker against, which is the same index the sink counts.
 */
const PROPERTY_CHILDREN: ReadonlySet<string> = new Set([
  "sectPr",
  "tblGrid",
  "tblPr",
  "tblPrEx",
  "tcPr",
  "trPr",
  "sdtPr",
  "sdtEndPr",
]);

/** Where every `w:bookmarkStart`/`w:bookmarkEnd` with this id sits. */
const anchorsOf = (xml: string, id: number): { start: Anchor[]; end: Anchor[] } => {
  const root = parseXmlDocument(xml);
  const found: { start: Anchor[]; end: Anchor[] } = { start: [], end: [] };
  if (!root) {
    return found;
  }
  const visit = (element: XmlElement): void => {
    const parent = getLocalName(element.name);
    let ordinal = 0;
    for (const child of getChildElements(element)) {
      const local = getLocalName(child.name);
      if (PROPERTY_CHILDREN.has(local)) {
        visit(child);
        continue;
      }
      if (local === "bookmarkStart" || local === "bookmarkEnd") {
        const attributes = child.attributes ?? {};
        const written = attributes["w:id"] ?? attributes["id"];
        if (written !== undefined && Number(written) === id) {
          found[local === "bookmarkStart" ? "start" : "end"].push({ parent, ordinal });
        }
      }
      ordinal += 1;
      visit(child);
    }
  };
  visit(root);
  return found;
};

const documentXmlOf = async (saved: ArrayBuffer): Promise<string> => {
  const zip = await JSZip.loadAsync(saved);
  return (await zip.file("word/document.xml")?.async("text")) ?? "";
};

const expectedAnchors = (placement: Placement): { start: Anchor[]; end: Anchor[] } => ({
  start: [{ parent: placement.parent, ordinal: placement.startOrdinal }],
  end: [{ parent: placement.parent, ordinal: placement.endOrdinal }],
});

describe("a bookmark marker at block level keeps its container and its ordinal", () => {
  test("every block container the schema lets hold one is exercised", () => {
    const declaring = BLOCK_LEVEL_CONTAINERS.filter((container) =>
      (CONTAINER_CHILDREN[container] as readonly string[]).includes("bookmarkStart"),
    );
    const exercised = new Set(PLACEMENT_KEYS.map((key) => PLACEMENTS[key].container));
    expect(declaring.filter((container) => !exercised.has(container))).toEqual([]);
  });

  test("every block container that declares one, through a save", async () => {
    await fc.assert(
      fc.asyncProperty(
        fc.constantFrom(...PLACEMENT_KEYS),
        fc.integer({ min: 1, max: 5000 }),
        async (key, id) => {
          const placement: Placement = PLACEMENTS[key];
          const parsed = await parseDocx(await buildDocx(placement.body(id)), {
            preloadFonts: false,
          });
          const saved = await repackDocx(parsed, { updateModifiedDate: false });

          expect({ key, ...anchorsOf(await documentXmlOf(saved), id) }).toEqual({
            key,
            ...expectedAnchors(placement),
          });
        },
      ),
      propertyConfig({ numRuns: 30 }),
    );
  });

  test("every block container that declares one, through the editor", async () => {
    await fc.assert(
      fc.asyncProperty(
        fc.constantFrom(...PLACEMENT_KEYS),
        fc.integer({ min: 1, max: 5000 }),
        async (key, id) => {
          const placement: Placement = PLACEMENTS[key];
          const parsed = await parseDocx(await buildDocx(placement.body(id)), {
            preloadFonts: false,
          });
          const document = fromProseDoc(toProseDoc(parsed), parsed);
          const saved = await repackDocx(document, { updateModifiedDate: false });

          expect({ key, ...anchorsOf(await documentXmlOf(saved), id) }).toEqual({
            key,
            ...expectedAnchors(placement),
          });
        },
      ),
      propertyConfig({ numRuns: 30 }),
    );
  });
});

describe("an edit inside a spanned cell keeps the bookmark's extent", () => {
  test("a row-level bookmark still opens and closes on the row", async () => {
    const id = 77;
    const parsed = await parseDocx(await buildDocx(PLACEMENTS.row.body(id)), {
      preloadFonts: false,
    });
    const proseDoc = toProseDoc(parsed);
    // Type into the first spanned cell: the edit ProseMirror's own mapping has
    // to carry the boundary through.
    const edited = proseDoc.type.schema.nodeFromJSON(
      JSON.parse(JSON.stringify(proseDoc.toJSON()).replace('"text":"a"', '"text":"a edited"')),
    );
    const saved = await repackDocx(fromProseDoc(edited, parsed), { updateModifiedDate: false });
    const xml = await documentXmlOf(saved);

    expect(xml).toContain("a edited");
    expect(anchorsOf(xml, id)).toEqual(expectedAnchors(PLACEMENTS.row));
  });
});
