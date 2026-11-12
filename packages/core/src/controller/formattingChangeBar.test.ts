import { panic } from "better-result";
import { expect, test } from "bun:test";

import { ZipArchive } from "@stll/docx-core/zip";

import { createCanonicalSession } from "../canonical/canonicalSession";
import { parseDocx } from "../docx/parser";
import { createEmptyDocx } from "../docx/rezip";
import { toFlowBlocks } from "../layout-bridge/convert/toFlowBlocks";
import type { ParagraphBlock } from "../layout-engine/types";
import { paragraphHasTrackedChanges } from "../layout-painter/renderParagraph";
import { toProseDoc } from "../prosemirror/conversion/toProseDoc";
import { schema } from "../prosemirror/schema";

const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const W14 = "http://schemas.microsoft.com/office/word/2010/wordml";
const STAMP = 'w:author="Reviewer" w:date="2026-01-01T00:00:00Z"';

const packageOf = async (body: string) => {
  const zip = await ZipArchive.read(await createEmptyDocx());
  zip.set(
    "word/document.xml",
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document xmlns:w="${W}" xmlns:w14="${W14}"><w:body>${body}<w:sectPr/></w:body></w:document>`,
  );
  return (await zip.generate({ compression: "store" })).buffer;
};

// A pending formatting change paints the text as it now is; the change bar
// is what tells the reader the paragraph holds a change, in either engine.
test("a paragraph holding a pending formatting change paints the change bar", async () => {
  const bytes = await packageOf(
    [
      `<w:p w14:paraId="10000001"><w:r><w:t xml:space="preserve">Plain </w:t></w:r><w:r><w:rPr><w:b/><w:rPrChange w:id="1" ${STAMP}><w:rPr/></w:rPrChange></w:rPr><w:t>bolded</w:t></w:r></w:p>`,
      `<w:p w14:paraId="10000002"><w:pPr><w:jc w:val="center"/><w:pPrChange w:id="2" ${STAMP}><w:pPr/></w:pPrChange></w:pPr><w:r><w:t>Centered</w:t></w:r></w:p>`,
      `<w:p w14:paraId="10000003"><w:r><w:t>Untouched</w:t></w:r></w:p>`,
    ].join(""),
  );
  const document = await parseDocx(bytes.slice(0), {
    preloadFonts: false,
    detectVariables: false,
  });
  const engines = {
    default: toProseDoc(
      document,
      document.package.styles === undefined
        ? {}
        : { styles: document.package.styles },
    ),
    canonical: createCanonicalSession(
      document,
      document.package.styles,
    ).unwrap().projection.doc,
  };
  for (const [engine, doc] of Object.entries(engines)) {
    const painted = toFlowBlocks(doc)
      .filter((block): block is ParagraphBlock => block.kind === "paragraph")
      .map((block) => [block.paraId, paragraphHasTrackedChanges(block)]);
    expect([engine, painted]).toEqual([
      engine,
      [
        ["10000001", true],
        ["10000002", true],
        ["10000003", false],
      ],
    ]);
  }
});

// The bar marks what the review list lists: a run property change mark that
// records no change has nothing to review.
test("a run property change mark that records no change paints no bar", () => {
  const doc = schema.node("doc", null, [
    schema.node("paragraph", { paraId: "10000001" }, [
      schema.text("Text", [
        schema.marks["runPropertyChange"]?.create({ changes: [] }) ??
          panic("No run property change mark."),
      ]),
    ]),
  ]);
  const painted = toFlowBlocks(doc)
    .filter((block): block is ParagraphBlock => block.kind === "paragraph")
    .map((block) => paragraphHasTrackedChanges(block));
  expect(painted).toEqual([false]);
});
