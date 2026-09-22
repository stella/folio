/**
 * A row- or cell-level content control keeps its wrapper, and what it wrapped.
 *
 * `CT_SdtRow` and `CT_SdtCell` are transparent, and folio used to prove it by
 * throwing the wrapper away: the rows or cells were spliced into the table or
 * the row and the control — its tag, alias, lock, data binding and
 * `w:sdtEndPr` — went with the `w:sdt` element. A template whose repeating
 * section or bound cell was merely opened and saved came back unbound.
 *
 * The subject is the control, so the variable is what sits beside it. The
 * generator walks the declared children of `table-content` and `row-content`
 * from `CONTAINER_CHILDREN` — the same generated set the dispatcher's handler
 * map is total over — and puts one of them next to the control, because a
 * sibling is what a walk that splices reorders and a walk that recurses loses
 * track of. Three legs, for three different fixes:
 *
 * - **save**, which a replay of the captured `w:sdtPr` would pass on its own;
 * - **capture-free save**, with every replayable slot cleared so the
 *   serializers have to rebuild the control from the model, which is the leg
 *   that failed before `SdtProperties.endProperties` existed;
 * - **editor projection**, through `toProseDoc`/`fromProseDoc`, which is where
 *   a record carried by an index rather than by the row would drift.
 */

import { describe, expect, setDefaultTimeout, test } from "bun:test";
import fc from "fast-check";
import JSZip from "jszip";

import { propertyConfig, propertyTestTimeout } from "../../../../test/property-testing";

import type { Document } from "../types/document";
import { fromProseDoc } from "../prosemirror/conversion/fromProseDoc";
import { toProseDoc } from "../prosemirror/conversion/toProseDoc";
import { CONTAINER_CHILDREN } from "./containerChildren.gen";
import { parseDocx } from "./parser";
import { createEmptyDocx, repackDocx } from "./rezip";

setDefaultTimeout(propertyTestTimeout(30_000));

const XML_DECLARATION = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';
const W_NAMESPACE = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const AUTHORED = 'w:author="A" w:date="2024-01-01T00:00:00Z"';

const cell = (text: string): string =>
  `<w:tc><w:tcPr/><w:p><w:r><w:t>${text}</w:t></w:r></w:p></w:tc>`;

const row = (text: string): string => `<w:tr>${cell(text)}</w:tr>`;

/**
 * The control under test.
 *
 * Every field the model projects out of `w:sdtPr` is spelled, so a leg that
 * rebuilds the properties rather than replaying them is checked against all of
 * them rather than against the element's name. `w:sdtEndPr` carries a `w:rPr`
 * because an empty one would pass whether or not the record holds anything.
 */
const CONTROL_PROPERTIES =
  '<w:sdtPr><w:id w:val="4242"/><w:alias w:val="Party block"/>' +
  '<w:tag w:val="party"/><w:lock w:val="sdtLocked"/></w:sdtPr>' +
  "<w:sdtEndPr><w:rPr><w:b/></w:rPr></w:sdtEndPr>";

const control = (inner: string): string =>
  `<w:sdt>${CONTROL_PROPERTIES}<w:sdtContent>${inner}</w:sdtContent></w:sdt>`;

/**
 * Markup for one declared child of a table or a row.
 *
 * A child the generated set names but this table cannot spell is not skipped:
 * the record is total, so a child added to the content model later arrives
 * here as a compile error rather than as a silently untested sibling.
 */
const DECLARED_CHILD_MARKUP = {
  bookmarkStart: '<w:bookmarkStart w:id="7" w:name="bm"/>',
  bookmarkEnd: '<w:bookmarkEnd w:id="7"/>',
  commentRangeStart: '<w:commentRangeStart w:id="8"/>',
  commentRangeEnd: '<w:commentRangeEnd w:id="8"/>',
  customXml: '<w:customXml w:uri="urn:example" w:element="e"/>',
  customXmlDelRangeStart: `<w:customXmlDelRangeStart w:id="9" ${AUTHORED}/>`,
  customXmlDelRangeEnd: '<w:customXmlDelRangeEnd w:id="9"/>',
  customXmlInsRangeStart: `<w:customXmlInsRangeStart w:id="10" ${AUTHORED}/>`,
  customXmlInsRangeEnd: '<w:customXmlInsRangeEnd w:id="10"/>',
  customXmlMoveFromRangeStart: `<w:customXmlMoveFromRangeStart w:id="11" ${AUTHORED}/>`,
  customXmlMoveFromRangeEnd: '<w:customXmlMoveFromRangeEnd w:id="11"/>',
  customXmlMoveToRangeStart: `<w:customXmlMoveToRangeStart w:id="12" ${AUTHORED}/>`,
  customXmlMoveToRangeEnd: '<w:customXmlMoveToRangeEnd w:id="12"/>',
  customXmlPr: "<w:customXmlPr/>",
  del: `<w:del w:id="13" ${AUTHORED}/>`,
  ins: `<w:ins w:id="14" ${AUTHORED}/>`,
  moveFrom: `<w:moveFrom w:id="15" ${AUTHORED}/>`,
  moveFromRangeStart: `<w:moveFromRangeStart w:id="16" w:name="mv16" ${AUTHORED}/>`,
  moveFromRangeEnd: '<w:moveFromRangeEnd w:id="16"/>',
  moveTo: `<w:moveTo w:id="17" ${AUTHORED}/>`,
  moveToRangeStart: `<w:moveToRangeStart w:id="18" w:name="mv18" ${AUTHORED}/>`,
  moveToRangeEnd: '<w:moveToRangeEnd w:id="18"/>',
  permStart: '<w:permStart w:id="19" w:edGrp="everyone"/>',
  permEnd: '<w:permEnd w:id="19"/>',
  proofErr: '<w:proofErr w:type="spellStart"/>',
  sdt: control(""),
  tbl: "",
  tblGrid: "",
  tblPr: "",
  tblPrEx: "",
  tc: "",
  tr: "",
  trPr: "",
} as const satisfies Record<
  | (typeof CONTAINER_CHILDREN)["table-content"][number]
  | (typeof CONTAINER_CHILDREN)["row-content"][number]
  | "tbl",
  string
>;

/**
 * The siblings worth generating: the markers, and not the structural children.
 *
 * `w:tblPr`, `w:tblGrid`, `w:trPr` and `w:tblPrEx` are property elements the
 * fixture already writes in their declared position, `w:tr` and `w:tc` are the
 * content the control wraps, and the nested `w:sdt` has a test of its own
 * below. What is left is exactly the markers a splicing walk misplaces.
 */
const siblingsOf = (container: "table-content" | "row-content"): readonly string[] =>
  CONTAINER_CHILDREN[container]
    .map((name) => DECLARED_CHILD_MARKUP[name])
    .filter((markup) => markup.length > 0 && markup !== DECLARED_CHILD_MARKUP.sdt);

const TABLE_SIBLINGS = siblingsOf("table-content");
const ROW_SIBLINGS = siblingsOf("row-content");

const GRID = '<w:tblGrid><w:gridCol w:w="2000"/><w:gridCol w:w="2000"/></w:tblGrid>';

/** A table whose first row sits inside a row-level control, with a sibling. */
const tableWithRowControl = (sibling: string): string =>
  `<w:tbl><w:tblPr/>${GRID}${sibling}${control(row("controlled"))}${row("plain")}</w:tbl>`;

/** A table whose first cell sits inside a cell-level control, with a sibling. */
const tableWithCellControl = (sibling: string): string =>
  `<w:tbl><w:tblPr/>${GRID}<w:tr>${sibling}${control(cell("controlled"))}${cell("plain")}</w:tr></w:tbl>`;

const buildDocx = async (body: string): Promise<ArrayBuffer> => {
  const zip = await JSZip.loadAsync(await createEmptyDocx());
  zip.file(
    "word/document.xml",
    `${XML_DECLARATION}<w:document xmlns:w="${W_NAMESPACE}">` +
      `<w:body>${body}<w:sectPr/></w:body></w:document>`,
  );
  return zip.generateAsync({ type: "arraybuffer" });
};

/**
 * Every slot that holds replayable markup rather than a parsed shape.
 *
 * The same forcing `scripts/lib/container-survival/laws.ts` applies, restated
 * here rather than imported so this suite reads no file outside its package:
 * whatever survives with these cleared, the typed model holds.
 */
const CAPTURE_SLOT_NAMES: ReadonlySet<string> = new Set([
  "gridSourceXml",
  "rawEndPropertiesXml",
  "sourceXml",
]);

const clearCaptures = (value: unknown, seen: WeakSet<object>): void => {
  if (Array.isArray(value)) {
    for (const item of value) {
      clearCaptures(item, seen);
    }
    return;
  }
  if (value instanceof Map) {
    for (const item of value.values()) {
      clearCaptures(item, seen);
    }
    return;
  }
  if (typeof value !== "object" || value === null || seen.has(value)) {
    return;
  }
  seen.add(value);
  const record = value as Record<string, unknown>;
  for (const key of Object.keys(record)) {
    if (CAPTURE_SLOT_NAMES.has(key)) {
      record[key] = undefined;
      continue;
    }
    clearCaptures(record[key], seen);
  }
};

const LEGS = {
  /** Parse and save: the captured `w:sdtPr` is replayed. */
  save: "save",
  /** Parse, clear every capture, save: the serializers rebuild the control. */
  captureFree: "capture-free save",
  /** Parse, through ProseMirror, save. */
  editor: "editor projection",
} as const;

type Leg = (typeof LEGS)[keyof typeof LEGS];

const LEG_VALUES = Object.values(LEGS);

const documentForLeg = (parsed: Document, leg: Leg): Document => {
  switch (leg) {
    case LEGS.save:
      return parsed;
    case LEGS.captureFree:
      clearCaptures(parsed.package, new WeakSet());
      return parsed;
    case LEGS.editor:
      return fromProseDoc(toProseDoc(parsed), parsed);
    default: {
      const unreachable: never = leg;
      return unreachable;
    }
  }
};

const savedDocumentXml = async (body: string, leg: Leg): Promise<string> => {
  const parsed = await parseDocx(await buildDocx(body), { preloadFonts: false });
  const saved = await repackDocx(documentForLeg(parsed, leg), { updateModifiedDate: false });
  const zip = await JSZip.loadAsync(saved);
  return (await zip.file("word/document.xml")?.async("text")) ?? "";
};

/** The control's wrapper, and what each of its modelled fields writes. */
const CONTROL_MARKS = [
  "<w:sdt>",
  '<w:id w:val="4242"/>',
  '<w:alias w:val="Party block"/>',
  '<w:tag w:val="party"/>',
  '<w:lock w:val="sdtLocked"/>',
  "<w:sdtEndPr>",
] as const;

const controlSurvives = (xml: string): Record<string, boolean> =>
  Object.fromEntries(CONTROL_MARKS.map((mark) => [mark, xml.includes(mark)]));

const ALL_PRESENT: Record<string, boolean> = Object.fromEntries(
  CONTROL_MARKS.map((mark) => [mark, true]),
);

/** What the control holds, between its own start and end tags. */
const insideFirstControl = (xml: string): string => {
  const opened = xml.indexOf("<w:sdtContent>");
  return opened === -1 ? "" : xml.slice(opened, xml.indexOf("</w:sdtContent>"));
};

describe("a table content control keeps its wrapper", () => {
  test("a row-level control survives every declared sibling, on every leg", async () => {
    await fc.assert(
      fc.asyncProperty(
        fc.constantFrom(...TABLE_SIBLINGS),
        fc.constantFrom(...LEG_VALUES),
        async (sibling, leg) => {
          const saved = await savedDocumentXml(tableWithRowControl(sibling), leg);

          expect({ leg, ...controlSurvives(saved) }).toEqual({ leg, ...ALL_PRESENT });
          // The control still holds the row it held, and only that row: a walk
          // that spliced would leave the wrapper empty or swallow the sibling.
          expect(insideFirstControl(saved)).toContain("<w:t>controlled</w:t>");
          expect(insideFirstControl(saved)).not.toContain("<w:t>plain</w:t>");
        },
      ),
      propertyConfig({ numRuns: 60 }),
    );
    // Each run builds, parses and repacks a package; sixty of them do not fit
    // the default five-second budget on a loaded machine.
  }, 120_000);

  test("a cell-level control survives every declared sibling, on every leg", async () => {
    await fc.assert(
      fc.asyncProperty(
        fc.constantFrom(...ROW_SIBLINGS),
        fc.constantFrom(...LEG_VALUES),
        async (sibling, leg) => {
          const saved = await savedDocumentXml(tableWithCellControl(sibling), leg);

          expect({ leg, ...controlSurvives(saved) }).toEqual({ leg, ...ALL_PRESENT });
          expect(insideFirstControl(saved)).toContain("<w:t>controlled</w:t>");
          expect(insideFirstControl(saved)).not.toContain("<w:t>plain</w:t>");
        },
      ),
      propertyConfig({ numRuns: 60 }),
    );
  }, 120_000);

  test("saving is a fixed point: the second save writes what the first did", async () => {
    for (const body of [tableWithRowControl(""), tableWithCellControl("")]) {
      const once = await savedDocumentXml(body, LEGS.editor);
      const table = once.slice(once.indexOf("<w:tbl>"), once.indexOf("</w:tbl>") + 8);
      const twice = await savedDocumentXml(table, LEGS.editor);

      expect(twice.slice(twice.indexOf("<w:tbl>"), twice.indexOf("</w:tbl>") + 8)).toBe(table);
    }
  }, 60_000);

  test("a control over two rows comes back as one wrapper, not two", async () => {
    const body =
      `<w:tbl><w:tblPr/>${GRID}` +
      `${control(`${row("first")}${row("second")}`)}${row("plain")}</w:tbl>`;

    for (const leg of LEG_VALUES) {
      const saved = await savedDocumentXml(body, leg);

      // One wrapper: rebuilding one per row would mint a second `w:id` for a
      // control the author wrote once.
      expect({ leg, wrappers: saved.split("<w:sdt>").length - 1 }).toEqual({ leg, wrappers: 1 });
      const inside = insideFirstControl(saved);
      expect(inside).toContain("<w:t>first</w:t>");
      expect(inside).toContain("<w:t>second</w:t>");
      expect(inside).not.toContain("<w:t>plain</w:t>");
    }
  }, 60_000);

  test("a preserved child inside a control keeps that control's stack", async () => {
    const opaque = '<w:proofErr w:type="spellStart"/>';
    const start = '<w:bookmarkStart w:id="80" w:name="inside"/>';
    const end = '<w:bookmarkEnd w:id="80"/>';
    const bodies = [
      `<w:tbl><w:tblPr/>${GRID}${control(
        `${row("first")}${opaque}${start}${row("second")}${end}`,
      )}</w:tbl>`,
      `<w:tbl><w:tblPr/>${GRID}<w:tr>${control(
        `${cell("first")}${opaque}${start}${cell("second")}${end}`,
      )}</w:tr></w:tbl>`,
    ];

    for (const body of bodies) {
      for (const leg of LEG_VALUES) {
        const saved = await savedDocumentXml(body, leg);

        expect({ leg, wrappers: saved.split("<w:sdt>").length - 1 }).toEqual({
          leg,
          wrappers: 1,
        });
        expect(insideFirstControl(saved)).toContain(opaque);
        expect(insideFirstControl(saved)).toContain(start);
        expect(insideFirstControl(saved)).toContain(end);
      }
    }
  }, 60_000);

  test("an empty or opaque table control remains positioned in its parent", async () => {
    const empty = control("");
    const opaque = `<w:sdt>${CONTROL_PROPERTIES}</w:sdt>`;
    const bodies = [
      `<w:tbl><w:tblPr/>${GRID}${empty}${row("plain")}</w:tbl>`,
      `<w:tbl><w:tblPr/>${GRID}<w:tr>${empty}${cell("plain")}</w:tr></w:tbl>`,
      `<w:tbl><w:tblPr/>${GRID}<w:tr><w:tc><w:tcPr/>${empty}<w:p/></w:tc></w:tr></w:tbl>`,
      `<w:tbl><w:tblPr/>${GRID}<w:tr><w:tc><w:tcPr/>${opaque}<w:p/></w:tc></w:tr></w:tbl>`,
    ];

    for (const body of bodies) {
      for (const leg of LEG_VALUES) {
        const saved = await savedDocumentXml(body, leg);

        expect({ leg, wrappers: saved.split("<w:sdt>").length - 1 }).toEqual({
          leg,
          wrappers: 1,
        });
        expect(saved).toContain('<w:tag w:val="party"/>');
        expect(insideFirstControl(saved)).not.toContain("plain");
      }
    }
  }, 60_000);

  test("direct SDT siblings keep their side of the content", async () => {
    const before = '<w:bookmarkStart w:id="81" w:name="before"/>';
    const after = '<w:bookmarkEnd w:id="81"/>';
    const wrapped =
      `<w:sdt>${CONTROL_PROPERTIES}${before}` +
      `<w:sdtContent>${row("controlled")}</w:sdtContent>${after}</w:sdt>`;
    const body = `<w:tbl><w:tblPr/>${GRID}${wrapped}</w:tbl>`;

    for (const leg of LEG_VALUES) {
      const saved = await savedDocumentXml(body, leg);
      const contentStart = saved.indexOf("<w:sdtContent>");
      const contentEnd = saved.indexOf("</w:sdtContent>");

      expect(saved.indexOf(before)).toBeGreaterThanOrEqual(0);
      expect(saved.indexOf(before)).toBeLessThan(contentStart);
      expect(saved.indexOf(after)).toBeGreaterThan(contentEnd);
    }
  }, 60_000);

  test("a foreign same-named property element is not read as WordprocessingML", async () => {
    const body =
      `<w:tbl><w:tblPr/>${GRID}` +
      `<w:sdt xmlns:x="urn:foreign"><x:sdtPr><x:tag x:val="foreign"/></x:sdtPr>` +
      `${CONTROL_PROPERTIES}<w:sdtContent>${row("controlled")}</w:sdtContent></w:sdt></w:tbl>`;

    for (const leg of LEG_VALUES) {
      const saved = await savedDocumentXml(body, leg);

      expect(saved).toContain('<w:tag w:val="party"/>');
      expect(saved).not.toContain('<w:tag w:val="foreign"/>');
      expect(saved).toContain("urn:foreign");
    }
  }, 60_000);

  test("a row control inside a table inside a cell control keeps both", async () => {
    const inner =
      `<w:tbl><w:tblPr/>${GRID}` +
      `<w:sdt><w:sdtPr><w:tag w:val="inner"/></w:sdtPr>` +
      `<w:sdtContent>${row("deep")}</w:sdtContent></w:sdt></w:tbl>`;
    const body =
      `<w:tbl><w:tblPr/>${GRID}<w:tr>` +
      `<w:sdt><w:sdtPr><w:tag w:val="outer"/></w:sdtPr>` +
      `<w:sdtContent><w:tc><w:tcPr/>${inner}</w:tc></w:sdtContent></w:sdt>` +
      `${cell("plain")}</w:tr></w:tbl>`;

    for (const leg of LEG_VALUES) {
      const saved = await savedDocumentXml(body, leg);

      expect({ leg, outer: saved.indexOf('<w:tag w:val="outer"/>') }).not.toEqual({
        leg,
        outer: -1,
      });
      // The inner control is inside the outer one, which is what nesting means
      // and what a flat record on the row could not have said.
      expect(saved.indexOf('<w:tag w:val="outer"/>')).toBeLessThan(
        saved.indexOf('<w:tag w:val="inner"/>'),
      );
      expect(saved).toContain("<w:t>deep</w:t>");
    }
  }, 60_000);

  test("a control directly inside a control keeps both, outermost first", async () => {
    const body =
      `<w:tbl><w:tblPr/>${GRID}` +
      `<w:sdt><w:sdtPr><w:tag w:val="section"/></w:sdtPr><w:sdtContent>` +
      `<w:sdt><w:sdtPr><w:tag w:val="item"/></w:sdtPr>` +
      `<w:sdtContent>${row("bound")}</w:sdtContent></w:sdt>` +
      `</w:sdtContent></w:sdt></w:tbl>`;

    for (const leg of LEG_VALUES) {
      const saved = await savedDocumentXml(body, leg);

      expect({ leg, wrappers: saved.split("<w:sdt>").length - 1 }).toEqual({ leg, wrappers: 2 });
      expect(saved.indexOf('<w:tag w:val="section"/>')).toBeLessThan(
        saved.indexOf('<w:tag w:val="item"/>'),
      );
      expect(saved).toContain("<w:t>bound</w:t>");
    }
  }, 60_000);
});
