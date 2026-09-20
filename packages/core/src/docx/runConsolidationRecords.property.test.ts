/**
 * Consolidation merges two runs only when their whole record agrees.
 *
 * `w:r` carries more than the formatting folio models. It carries the editing
 * session its text was written in (`w:rsidR` and its family, kept as the
 * attribute remainder) and the `w:rPr` children no reader took a typed value
 * from (kept as the property set's verbatim sink). A merge produces one run,
 * and one run holds one of each, so merging two runs that disagree on either
 * discards the loser's copy — at parse time, before the editor exists and
 * before any round-trip gate can compare anything.
 *
 * Two thirds of the merges folio performs on real documents are merges across
 * editing sessions, so this is not an edge case: it is most of what the
 * consolidator does. The price of not lying about editing history is roughly a
 * tenth more run records, which is why the merge rule is asserted here as a
 * law rather than left to a fixture somebody might relax.
 *
 * The properties run over subsets rather than examples because the defect is
 * per record: a predicate that compares the fields somebody thought to write
 * down keeps every example in this file and still merges across the one field
 * it forgot.
 */

import { describe, expect, test } from "bun:test";
import fc from "fast-check";
import JSZip from "jszip";

import { propertyConfig } from "../../../../test/property-testing";

import type { Document, Paragraph, Run } from "../types/document";
import { parseDocx } from "./parser";
import { createEmptyDocx, repackDocx } from "./rezip";
import { consolidateRuns, runsMergeable } from "./runConsolidator";

const XML_DECLARATION = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';
const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

/** The `w:rsid*` attributes `CT_R` declares, all of them `ST_LongHexNumber`. */
const RUN_RSID_ATTRIBUTES = ["rsidR", "rsidRPr", "rsidDel"] as const;

/**
 * `w:rPr` children folio models no field for, so each lands in the sink.
 *
 * One per shape the sink has to keep: an empty element, one with attributes,
 * and one whose schema ordinal sits at the far end of `EG_RPrBase`.
 */
const SINK_CHILDREN = {
  snapToGrid: '<w:snapToGrid w:val="0"/>',
  webHidden: "<w:webHidden/>",
  eastAsianLayout: '<w:eastAsianLayout w:id="5" w:combine="1"/>',
} as const;

type SinkChild = keyof typeof SINK_CHILDREN;

/**
 * One run's authored record: which editing session, which of its attributes
 * name that session, and which sink children the property set carries.
 */
type RunRecord = {
  session: number;
  rsids: readonly string[];
  sink: readonly SinkChild[];
};

/** `ST_LongHexNumber`, distinct per (attribute, session) pair. */
const rsidValue = (attribute: string, session: number): string =>
  `00${(((attribute.length * 131 + session * 7919) * 2_654_435_761) % 0xff_ff_ff)
    .toString(16)
    .padStart(6, "0")
    .toUpperCase()}`;

const runXml = ({ session, rsids, sink }: RunRecord, text: string): string => {
  const attributes = rsids.map((name) => ` w:${name}="${rsidValue(name, session)}"`).join("");
  const children = sink.map((name) => SINK_CHILDREN[name]).join("");
  // `w:b` is shared by every run, so typed formatting never decides a merge
  // here: only the remainder and the sink can.
  return `<w:r${attributes}><w:rPr><w:b/>${children}</w:rPr><w:t>${text}</w:t></w:r>`;
};

const documentXml = (records: readonly RunRecord[]): string =>
  `${XML_DECLARATION}<w:document xmlns:w="${W}"><w:body><w:p>` +
  records.map((record, index) => runXml(record, `r${index}`)).join("") +
  "</w:p></w:body></w:document>";

/** What a merge compares, flattened: two runs agree iff these strings match. */
const recordKey = ({ session, rsids, sink }: RunRecord): string =>
  `${rsids.length > 0 ? session : ""}|${rsids.join()}|${sink.join()}`;

const packageFor = async (xml: string): Promise<ArrayBuffer> => {
  const zip = await JSZip.loadAsync(await createEmptyDocx());
  zip.file("word/document.xml", xml);
  return zip.generateAsync({ type: "arraybuffer" });
};

const documentPartOf = async (buffer: ArrayBuffer): Promise<string> =>
  (await (await JSZip.loadAsync(buffer)).file("word/document.xml")?.async("text")) ?? "";

const save = (document: Document): Promise<ArrayBuffer> =>
  repackDocx(document, { updateModifiedDate: false });

const open = (xml: string): Promise<Document> =>
  packageFor(xml).then((buffer) => parseDocx(buffer, { preloadFonts: false }));

const firstParagraph = (document: Document): Paragraph => {
  const block = document.package.document.content.at(0);
  if (block?.type !== "paragraph") {
    throw new Error("the fixture's first block is a paragraph");
  }
  return block;
};

const runsOf = (document: Document): Run[] =>
  firstParagraph(document).content.filter((item): item is Run => item.type === "run");

const textOf = (run: Run): string =>
  run.content.map((content) => (content.type === "text" ? content.text : "")).join("");

/** The run's whole record, as the saved part spells it. */
const savedRunTags = (xml: string): string[] => [
  ...(xml.match(/<w:r(?:\s[^>]*)?>.*?<\/w:r>/gsu) ?? []),
];

/** Every `w:t` in the saved part, joined: the text however the runs were cut. */
const savedText = (xml: string): string =>
  [...xml.matchAll(/<w:t(?:\s[^>]*)?>(.*?)<\/w:t>/gsu)].map(([, text]) => text ?? "").join("");

const runRecordArbitrary = fc.record({
  session: fc.integer({ min: 0, max: 2 }),
  rsids: fc.subarray([...RUN_RSID_ATTRIBUTES] as string[], { minLength: 0 }),
  sink: fc.subarray(Object.keys(SINK_CHILDREN) as SinkChild[], { minLength: 0 }),
});

const plainRun = (text: string): Run => ({
  type: "run",
  formatting: { bold: true },
  content: [{ type: "text", text }],
});

describe("a merge needs the whole record to agree", () => {
  test("two sessions stay two runs, and both keep their attributes", async () => {
    const xml =
      `${XML_DECLARATION}<w:document xmlns:w="${W}"><w:body><w:p>` +
      '<w:r w:rsidR="00AAAAAA"><w:t>one</w:t></w:r>' +
      '<w:r w:rsidR="00BBBBBB"><w:t>two</w:t></w:r>' +
      "</w:p></w:body></w:document>";
    const saved = await documentPartOf(await save(await open(xml)));

    expect(savedRunTags(saved)).toEqual([
      '<w:r w:rsidR="00AAAAAA"><w:t>one</w:t></w:r>',
      '<w:r w:rsidR="00BBBBBB"><w:t>two</w:t></w:r>',
    ]);
  });

  test("two property sets stay two runs, so neither run wears the other's bytes", async () => {
    const xml =
      `${XML_DECLARATION}<w:document xmlns:w="${W}"><w:body><w:p>` +
      "<w:r><w:rPr><w:b/><w:webHidden/></w:rPr><w:t>one</w:t></w:r>" +
      "<w:r><w:rPr><w:b/></w:rPr><w:t>two</w:t></w:r>" +
      "</w:p></w:body></w:document>";
    const saved = await documentPartOf(await save(await open(xml)));

    expect(savedRunTags(saved)).toEqual([
      "<w:r><w:rPr><w:b/><w:webHidden/></w:rPr><w:t>one</w:t></w:r>",
      "<w:r><w:rPr><w:b/></w:rPr><w:t>two</w:t></w:r>",
    ]);
  });

  test("one session and one property set become one run, stated once", async () => {
    const record = { session: 1, rsids: ["rsidR", "rsidDel"], sink: ["webHidden"] } as const;
    const xml =
      `${XML_DECLARATION}<w:document xmlns:w="${W}"><w:body><w:p>` +
      runXml(record, "one") +
      runXml(record, "two") +
      "</w:p></w:body></w:document>";
    const saved = await documentPartOf(await save(await open(xml)));

    expect(savedRunTags(saved)).toHaveLength(1);
    expect(saved).toContain("<w:t>onetwo</w:t>");
    expect([...saved.matchAll(/w:rsidR=/gu)]).toHaveLength(1);
    expect([...saved.matchAll(/<w:webHidden\/>/gu)]).toHaveLength(1);
  });

  test("a remainder spelled in another order is the same remainder", () => {
    const a: Run = {
      ...plainRun("one"),
      preservedAttributes: [
        { namespace: W, name: "rsidR", value: "00AAAAAA" },
        { namespace: W, name: "rsidDel", value: "00CCCCCC" },
      ],
    };
    const b: Run = {
      ...plainRun("two"),
      preservedAttributes: [
        { namespace: W, name: "rsidDel", value: "00CCCCCC" },
        { namespace: W, name: "rsidR", value: "00AAAAAA" },
      ],
    };

    // Attribute order in XML says nothing, so the set decides, and the merged
    // run keeps the survivor's spelling of it.
    expect(runsMergeable(a, b)).toBe(true);
    expect(consolidateRuns([a, b])).toEqual([
      { ...a, content: [{ type: "text", text: "onetwo" }] },
    ]);
  });

  test("a sink in another order is another sink", () => {
    const withSink = (children: { index: number; xml: string }[]): Run => ({
      ...plainRun("x"),
      formatting: { bold: true, preserved: { children } },
    });
    const a = withSink([
      { index: 21, xml: "<w:webHidden/>" },
      { index: 40, xml: "<w:specVanish/>" },
    ]);
    const b = withSink([
      { index: 40, xml: "<w:specVanish/>" },
      { index: 21, xml: "<w:webHidden/>" },
    ]);

    // The sink is ordered: its order is what puts the markup back between the
    // same modelled siblings.
    expect(runsMergeable(a, b)).toBe(false);
  });
});

describe("the merge rule is what parse converges on", () => {
  test("adjacent runs merge exactly when their records agree", async () => {
    await fc.assert(
      fc.asyncProperty(
        fc.array(runRecordArbitrary, { minLength: 2, maxLength: 5 }),
        async (records) => {
          const parsed = runsOf(await open(documentXml(records)));

          // Every surviving run's text is the concatenation of one maximal
          // block of records that a merge may join, in order.
          const expectedTexts: string[] = [];
          for (const [index, record] of records.entries()) {
            const previous = records[index - 1];
            const joins = previous !== undefined && recordKey(record) === recordKey(previous);
            const last = expectedTexts.at(-1);
            if (joins && last !== undefined) {
              expectedTexts[expectedTexts.length - 1] = last + `r${index}`;
              continue;
            }
            expectedTexts.push(`r${index}`);
          }

          expect(parsed.map(textOf)).toEqual(expectedTexts);
        },
      ),
      propertyConfig({ numRuns: 40 }),
    );
  }, 120_000);

  test("parse is a fixed point over arbitrary records", async () => {
    await fc.assert(
      fc.asyncProperty(
        fc.array(runRecordArbitrary, { minLength: 1, maxLength: 5 }),
        async (records) => {
          const saved = await documentPartOf(await save(await open(documentXml(records))));

          // The second save is where a consolidation that merges more than the
          // next parse would stops being a fixed point: the text has already
          // been joined, so the run count can only fall again.
          expect(await documentPartOf(await save(await open(saved)))).toBe(saved);
        },
      ),
      propertyConfig({ numRuns: 25 }),
    );
  }, 120_000);

  test("the text and the record survive, whatever the merge decides", async () => {
    await fc.assert(
      fc.asyncProperty(
        fc.array(runRecordArbitrary, { minLength: 1, maxLength: 5 }),
        async (records) => {
          const saved = await documentPartOf(await save(await open(documentXml(records))));

          expect(savedText(saved)).toBe(records.map((_, index) => `r${index}`).join(""));
          for (const { session, rsids, sink } of records) {
            for (const attribute of rsids) {
              expect(saved).toContain(`w:${attribute}="${rsidValue(attribute, session)}"`);
            }
            for (const child of sink) {
              expect(saved).toContain(SINK_CHILDREN[child]);
            }
          }
        },
      ),
      propertyConfig({ numRuns: 25 }),
    );
  }, 120_000);
});
