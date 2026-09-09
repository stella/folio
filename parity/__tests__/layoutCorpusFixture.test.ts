import { describe, expect, test } from "bun:test";
import path from "node:path";

import JSZip from "jszip";

import { buildLayoutCorpus } from "../fixtures/build-layout-corpus";

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
    expect(pairwiseDocument).toContain("<w:lastRenderedPageBreak/>");
    expect(pairwiseDocument).toContain('<w:type w:val="continuous"/>');
    expect(pairwiseDocument).toContain('relativeFrom="page"');

    const kitchenSink = await loadFixture("layout-kitchen-sink.docx");
    const kitchenDocument = await kitchenSink.file("word/document.xml")!.async("text");
    expect(kitchenDocument).toContain("<w:tbl>");
    expect(kitchenDocument).toContain("<w:footnoteReference");
    expect(kitchenDocument).toContain("<w:bidi/>");
    expect(kitchenDocument).toContain("<w:cols");
    expect(kitchenSink.file("word/footnotes.xml")).not.toBeNull();
  });
});
