/**
 * A save may not write a run the next parse drops.
 *
 * Whether a run is kept is one question — `runHoldsPayload` — and three sites
 * ask it: the parser's keep rule, the consolidator, and the serializer. While
 * they agreed only by coincidence, a run that folio itself had emptied made
 * the saves oscillate: save 1 wrote `<w:r><w:rPr…/></w:r>`, the next parse
 * dropped it, and save 2 differed from save 1 over a run neither save showed.
 * The public corpus found it on a text box, where enrichment lifts the shape
 * into a run of its own and leaves the carrier holding only its properties.
 *
 * The property generates run sequences rather than examples because the defect
 * is about which run stands next to which: a carrier alone in a paragraph
 * survived consolidation on a length check, and the same carrier beside a text
 * run did not.
 */

import { describe, expect, setDefaultTimeout, test } from "bun:test";
import fc from "fast-check";
import JSZip from "jszip";

import { propertyConfig, propertyTestTimeout } from "../../../../test/property-testing";
import type { Document, Paragraph, Run } from "../types/document";

import { parseDocx } from "./parser";
import { createEmptyDocx, repackDocx } from "./rezip";

setDefaultTimeout(propertyTestTimeout(30_000));

const XML_DECLARATION = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';
const W_NAMESPACE = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const WP_NAMESPACE = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing";
const A_NAMESPACE = "http://schemas.openxmlformats.org/drawingml/2006/main";
const WPS_NAMESPACE = "http://schemas.microsoft.com/office/word/2010/wordprocessingShape";

/**
 * A `wps` text box, whose run the parser keeps empty on purpose: the shape is
 * claimed by `enrichParagraphTextBoxes`, a second pass over the same paragraph,
 * so between the two passes the run that carries it holds nothing.
 */
const textBoxDrawing = (id: number): string =>
  '<w:drawing><wp:inline><wp:extent cx="914400" cy="457200"/>' +
  `<wp:docPr id="${id}" name="Text Box ${id}"/><a:graphic>` +
  `<a:graphicData uri="${WPS_NAMESPACE}"><wps:wsp><wps:spPr>` +
  '<a:xfrm><a:off x="0" y="0"/><a:ext cx="914400" cy="457200"/></a:xfrm>' +
  '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></wps:spPr>' +
  "<wps:txbx><w:txbxContent><w:p><w:r><w:t>Text in box</w:t></w:r></w:p></w:txbxContent></wps:txbx>" +
  "<wps:bodyPr/></wps:wsp></a:graphicData></a:graphic></wp:inline></w:drawing>";

/**
 * The run shapes that hold nothing, nearly nothing, or nothing yet.
 *
 * `bare` and `properties` hold no payload at all. `emptyText` holds one: a
 * `w:t` that parses to the empty string is still a text node the model keeps,
 * and conflating the two is how the class starts. `textBox` holds one the
 * model has not been given yet.
 */
const RUN_BODIES = {
  bare: () => "",
  properties: () => "",
  emptyText: () => "<w:t></w:t>",
  text: () => "<w:t>x</w:t>",
  tab: () => "<w:tab/>",
  break: () => "<w:br/>",
  textBox: (index: number) => textBoxDrawing(index + 1),
} as const;

type RunKind = keyof typeof RUN_BODIES;

/** Which kinds state a `w:rPr` whatever the generated case says. */
const KINDS_ALWAYS_STATING_PROPERTIES: ReadonlySet<RunKind> = new Set(["properties"]);

type GeneratedRun = { kind: RunKind; bold: boolean };

const runXml = ({ kind, bold }: GeneratedRun, index: number): string => {
  const statesProperties = bold || KINDS_ALWAYS_STATING_PROPERTIES.has(kind);
  const properties = statesProperties ? `<w:rPr>${bold ? "<w:b/>" : ""}</w:rPr>` : "";
  return `<w:r>${properties}${RUN_BODIES[kind](index)}</w:r>`;
};

/**
 * A bookmark pair before the paragraph, or none.
 *
 * folio moves a body-level bookmark into the paragraph that follows it, which
 * shifts the position enrichment matches a lifted shape at. That is what turns
 * a carrier the enrichment fills back in place into a carrier left behind, so
 * it is the generated variable rather than a fixture's decision.
 */
const generatedCase = fc.record({
  runs: fc.array(
    fc.record({
      kind: fc.constantFrom(...(Object.keys(RUN_BODIES) as RunKind[])),
      bold: fc.boolean(),
    }),
    { minLength: 1, maxLength: 8 },
  ),
  bookmarked: fc.boolean(),
});

const BOOKMARK_PAIR = '<w:bookmarkStart w:id="0" w:name="_GoBack"/><w:bookmarkEnd w:id="0"/>';

const bodyXml = ({
  runs,
  bookmarked,
}: {
  runs: readonly GeneratedRun[];
  bookmarked: boolean;
}): string => `${bookmarked ? BOOKMARK_PAIR : ""}<w:p>${runs.map(runXml).join("")}</w:p>`;

const buildDocx = async (body: string): Promise<ArrayBuffer> => {
  const zip = await JSZip.loadAsync(await createEmptyDocx());
  zip.file(
    "word/document.xml",
    `${XML_DECLARATION}<w:document xmlns:w="${W_NAMESPACE}" xmlns:wp="${WP_NAMESPACE}" ` +
      `xmlns:a="${A_NAMESPACE}" xmlns:wps="${WPS_NAMESPACE}">` +
      `<w:body>${body}<w:sectPr/></w:body></w:document>`,
  );
  return zip.generateAsync({ type: "arraybuffer" });
};

/** Every part of a package, by path, so a comparison names the part that drifted. */
const packageParts = async (buffer: ArrayBuffer): Promise<Map<string, string>> => {
  const zip = await JSZip.loadAsync(buffer);
  return new Map(
    await Promise.all(
      Object.entries(zip.files)
        .filter(([, file]) => !file.dir)
        .map(
          async ([path, file]): Promise<readonly [string, string]> => [
            path,
            await file.async("base64"),
          ],
        ),
    ),
  );
};

const bodyParagraphs = ({ package: { document } }: Document): Paragraph[] =>
  document.content.filter((block): block is Paragraph => block.type === "paragraph");

const bodyRuns = (document: Document): Run[] =>
  bodyParagraphs(document).flatMap((paragraph) =>
    paragraph.content.filter((item): item is Run => item.type === "run"),
  );

/**
 * What each run in the body holds, in order.
 *
 * A run's identity here is its payload plus the properties that payload is
 * shown with, so a run that survived with its `w:rPr` stripped reads as a
 * different run rather than as the same one. Shape and drawing payloads carry
 * generated ids, so only their kind is compared.
 */
const runRecords = (document: Document): string[] =>
  bodyRuns(document).map((run) =>
    JSON.stringify({
      formatting: run.formatting ?? null,
      content: run.content.map((content) =>
        content.type === "text" ? { type: content.type, text: content.text } : content.type,
      ),
    }),
  );

type TwoSaves = {
  first: ArrayBuffer;
  afterFirst: Document;
  afterSecond: Document;
};

/**
 * Save, parse, save, parse, and assert the law on the way through.
 *
 * The first save is allowed to normalise; from there on the package is fixed.
 * Byte stability alone is also what dropping everything achieves, so the runs
 * the first save wrote are compared as well: every one of them is still there
 * after the second.
 */
const expectFixedPoint = async (source: ArrayBuffer): Promise<TwoSaves> => {
  const parsed = await parseDocx(source, { preloadFonts: false });
  const first = await repackDocx(parsed, { updateModifiedDate: false });
  const afterFirst = await parseDocx(first, { preloadFonts: false });
  const second = await repackDocx(afterFirst, { updateModifiedDate: false });
  const afterSecond = await parseDocx(second, { preloadFonts: false });

  const firstParts = await packageParts(first);
  const secondParts = await packageParts(second);
  expect([...secondParts.keys()].sort()).toEqual([...firstParts.keys()].sort());
  for (const [path, content] of firstParts) {
    expect({ path, content }).toEqual({ path, content: secondParts.get(path) });
  }
  expect(runRecords(afterSecond)).toEqual(runRecords(afterFirst));

  return { first, afterFirst, afterSecond };
};

describe("a run that holds no payload is never written", () => {
  test("any sequence of empty, near-empty and not-yet-filled runs is a fixed point after the first save", async () => {
    await fc.assert(
      fc.asyncProperty(generatedCase, async (generated) => {
        await expectFixedPoint(await buildDocx(bodyXml(generated)));
      }),
      propertyConfig({ numRuns: 40 }),
    );
  }, 60_000); // Each case zips, parses and repacks a package twice over.

  test("a paragraph of nothing but empty runs writes no run at all", async () => {
    const { afterFirst } = await expectFixedPoint(
      await buildDocx("<w:p><w:r/><w:r><w:rPr><w:b/></w:rPr></w:r></w:p>"),
    );
    expect(runRecords(afterFirst)).toEqual([]);
  });
});

/**
 * The corpus construct, minimised: a bookmark pair before a paragraph whose
 * only run carries run properties and a `wps` text box. This is `fdo66929`
 * from the `libreoffice-core` collection with everything else taken away.
 */
const CARRIER_RUN = `<w:r><w:rPr><w:noProof/></w:rPr>${textBoxDrawing(11)}</w:r>`;

describe("a text box does not leave its carrier run behind", () => {
  test("the carrier emptied by enrichment is not written", async () => {
    const { first } = await expectFixedPoint(
      await buildDocx(`${BOOKMARK_PAIR}<w:p>${CARRIER_RUN}</w:p>`),
    );
    const zip = await JSZip.loadAsync(first);
    expect(await zip.file("word/document.xml")?.async("text")).not.toContain(
      "<w:rPr><w:noProof/></w:rPr></w:r>",
    );
  });

  test("a carrier between two mergeable runs is a boundary the file can hold", async () => {
    // A carrier the save does not write is a merge boundary only the model
    // sees: the runs on either side stay apart on the first parse and merge on
    // the second, so save 2 loses a run save 1 wrote.
    const { afterFirst } = await expectFixedPoint(
      await buildDocx(
        `${BOOKMARK_PAIR}<w:p><w:r><w:rPr><w:b/></w:rPr><w:t>x</w:t></w:r>${CARRIER_RUN}` +
          "<w:r><w:rPr><w:b/></w:rPr><w:t></w:t></w:r></w:p>",
      ),
    );
    expect(runRecords(afterFirst)).toHaveLength(3);
  });

  test("a carrier standing after a text run keeps its own run properties", async () => {
    const { afterFirst } = await expectFixedPoint(
      await buildDocx(`<w:p><w:r><w:t>A</w:t></w:r>${CARRIER_RUN}</w:p>`),
    );
    const shapeRun = bodyRuns(afterFirst).find((run) =>
      run.content.some((content) => content.type === "shape"),
    );
    expect(shapeRun?.formatting?.noProof).toBe(true);
  });
});
