import { describe, expect, test } from "bun:test";
import path from "node:path";

import JSZip from "jszip";

import {
  buildLayoutCorpus,
  buildLayoutInteractionCaseFixture,
  layoutInteractionMatrixManifest,
} from "../fixtures/build-layout-corpus";
import { buildLayoutInteractionMatrix } from "../fixtures/layout-interaction-matrix";

const FIXTURES_DIR = path.join(import.meta.dir, "..", "fixtures");
const FIXTURE_NAMES = [
  "isolated-page-furniture.docx",
  "pairwise-layout-interactions.docx",
  "layout-kitchen-sink.docx",
] as const;

const loadFixture = async (name: (typeof FIXTURE_NAMES)[number]): Promise<JSZip> =>
  JSZip.loadAsync(await Bun.file(path.join(FIXTURES_DIR, name)).arrayBuffer());

describe("synthetic layout corpus", () => {
  test("committed fixtures are deterministic generator outputs", async () => {
    const generated = await buildLayoutCorpus(FIXTURES_DIR);

    for (const name of FIXTURE_NAMES) {
      const committed = new Uint8Array(await Bun.file(path.join(FIXTURES_DIR, name)).arrayBuffer());
      const expected = generated.get(name);
      if (!expected) {
        throw new TypeError(`generator omitted ${name}`);
      }
      expect(Buffer.from(committed).equals(Buffer.from(expected))).toBeTrue();
    }
    expect(await Bun.file(path.join(FIXTURES_DIR, "layout-interaction-matrix.json")).text()).toBe(
      layoutInteractionMatrixManifest(),
    );
  });

  test("keeps isolated, pairwise, and kitchen-sink coverage structurally distinct", async () => {
    const requiredParts = [
      "word/document.xml",
      "word/header1.xml",
      "word/footer1.xml",
      "word/media/synthetic-band.png",
    ];

    for (const name of FIXTURE_NAMES) {
      const zip = await loadFixture(name);
      for (const part of requiredParts) {
        expect(zip.file(part), `${name} must contain ${part}`).not.toBeNull();
      }
    }

    const pairwise = await loadFixture("pairwise-layout-interactions.docx");
    const pairwiseDocument = await pairwise.file("word/document.xml")!.async("text");
    const scenarios = buildLayoutInteractionMatrix();
    const authoredFlowBreaks = scenarios.filter(({ flow }) => flow === "hardPageBreak").length;
    const interCaseBreaks = scenarios
      .slice(1)
      .filter((_, index) => scenarios[index]?.section !== "nextPage").length;
    expect(pairwiseDocument.match(/<w:br w:type="page"\/>/g)).toHaveLength(
      authoredFlowBreaks + interCaseBreaks,
    );
    expect(pairwiseDocument).toContain("<w:lastRenderedPageBreak/>");
    expect(pairwiseDocument).toContain('<w:type w:val="continuous"/>');
    expect(pairwiseDocument).toContain('relativeFrom="page"');
    expect(pairwiseDocument).toContain('relativeFrom="margin"');
    expect(pairwiseDocument).toContain('relativeFrom="line"');
    expect(pairwiseDocument).toContain("<wp:inline");
    expect(pairwiseDocument).toContain("<wp:wrapNone");
    expect(pairwiseDocument).toContain("<wp:wrapSquare");
    expect(pairwiseDocument).toContain("<wp:wrapTopAndBottom");
    expect(pairwiseDocument).toContain('<w:tblLayout w:type="fixed"');
    expect(pairwiseDocument).toContain('<w:tblLayout w:type="autofit"');
    expect(pairwiseDocument).toContain("<w:vMerge");
    expect(pairwiseDocument).toContain('<w:lang w:eastAsia="ja-JP"');
    expect(pairwiseDocument).toContain("<w:bidi");
    expect(pairwiseDocument).toContain("<w:tab/>");
    expect(pairwiseDocument).toContain("<w:numPr>");
    expect(pairwiseDocument).toContain('<w:type w:val="nextPage"');
    expect(pairwiseDocument).toContain("<w:cols");
    for (const scenario of scenarios) {
      expect(pairwiseDocument).toContain(scenario.id);
    }

    const kitchenSink = await loadFixture("layout-kitchen-sink.docx");
    const kitchenDocument = await kitchenSink.file("word/document.xml")!.async("text");
    expect(kitchenDocument).toContain("<w:tbl>");
    expect(kitchenDocument).toContain("<w:footnoteReference");
    expect(kitchenDocument).toContain("<w:bidi/>");
    expect(kitchenDocument).toContain("<w:cols");
    expect(kitchenSink.file("word/footnotes.xml")).not.toBeNull();
  });

  test("isolates every matrix case with an observable section boundary", async () => {
    for (const scenario of buildLayoutInteractionMatrix()) {
      // oxlint-disable-next-line no-await-in-loop -- validates each bounded generated case independently
      const zip = await JSZip.loadAsync(await buildLayoutInteractionCaseFixture(scenario));
      // oxlint-disable-next-line no-await-in-loop -- JSZip exposes async part reads
      const document = await zip.file("word/document.xml")!.async("text");
      expect(document).toContain(`${scenario.id} post-boundary sentinel.`);
      if (scenario.section === "single") {
        expect(document).not.toContain("Section boundary");
      } else {
        expect(document).toContain("Section boundary");
      }
    }
  });
});
