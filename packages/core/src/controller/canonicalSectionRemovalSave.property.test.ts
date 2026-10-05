import { expect, setDefaultTimeout, test } from "bun:test";
import fc from "fast-check";
import JSZip from "jszip";
import { OP_STORIES } from "@stll/docx-core/ops";
import { panic } from "better-result";
import { EditorState } from "prosemirror-state";
import { createCanonicalSession, publishCanonicalProjection } from "./canonicalSession";
import { schema } from "../prosemirror/schema";
import { createEmptyDocument } from "../utils/createDocument";
import type { Document, Paragraph } from "../types/document";
import { repackWithCanonicalStoryRemovals } from "../docx/canonicalStoryRepack";
import { parseDocx } from "../docx/parser";
import { createDocx, repackDocx } from "../docx/rezip";
import { readDocumentSectionFacts } from "../docx/documentSectionFacts";
import { assertProperty, propertyTestTimeout } from "../../../../test/property-testing";

setDefaultTimeout(propertyTestTimeout(10_000));

type ParagraphOptions = {
  index: number;
  text: string;
  endsSection: boolean;
};

const paragraph = ({ index, text, endsSection }: ParagraphOptions): Paragraph => ({
  type: "paragraph",
  paraId: index.toString(16).padStart(8, "0").toUpperCase(),
  content: [{ type: "run", content: [{ type: "text", text }] }],
  ...(endsSection ? { sectionProperties: { pageWidth: 12240 } } : {}),
});

const documentWithSectionEndpoints = (removableCount: number): Document => {
  const document = createEmptyDocument();
  document.package.document.content = Array.from({ length: removableCount * 2 + 1 }, (_, index) => {
    const paragraphIndex = index + 1;
    return paragraph({
      index: paragraphIndex,
      text: `Paragraph ${paragraphIndex}`,
      endsSection: index % 2 === 0 && index < removableCount * 2,
    });
  });
  document.package.document.finalSectionProperties = { pageWidth: 12240 };
  return document;
};

const removeCanonicalSectionEndpoints = async (
  document: Document,
  removableCount: number,
): Promise<Document> => {
  const session = createCanonicalSession(document).unwrap();
  let state = EditorState.create({
    schema,
    doc: session.projection.doc,
  });
  for (let index = 0; index < removableCount; index += 1) {
    const blockId = (index * 2 + 1).toString(16).padStart(8, "0").toUpperCase();
    const nextBlockId = (index * 2 + 2).toString(16).padStart(8, "0").toUpperCase();
    const commit = session
      .prepareIntent(state, {
        type: "joinParagraphs",
        story: OP_STORIES.MAIN,
        blockId,
        nextBlockId,
      })
      .unwrap();
    const published = publishCanonicalProjection({ state, session, commit });
    if (published.isErr()) panic("Expected canonical section join to publish");
    state = published.value.state;
  }
  return session.document;
};

const documentXmlOf = async (buffer: ArrayBuffer): Promise<string> => {
  const zip = await JSZip.loadAsync(buffer);
  const xml = await zip.file("word/document.xml")?.async("text");
  if (!xml) panic("Expected word/document.xml");
  return xml;
};

test("canonical joins save the exact generated section removal count", async () => {
  await assertProperty(
    fc.asyncProperty(fc.integer({ min: 1, max: 4 }), async (count) => {
      const source = await createDocx(documentWithSectionEndpoints(count), {
        updateModifiedDate: false,
      });
      const original = await parseDocx(source, { preloadFonts: false });
      const edited = await removeCanonicalSectionEndpoints(original, count);

      await expect(repackDocx(edited, { updateModifiedDate: false })).rejects.toThrow(
        /drop section properties/u,
      );
      const saved = await repackWithCanonicalStoryRemovals({
        document: edited,
        repack: () => repackDocx(edited, { updateModifiedDate: false }),
      });
      const reopened = await parseDocx(saved, { preloadFonts: false });
      expect(readDocumentSectionFacts(await documentXmlOf(saved)).sectionCount).toBe(1);
      expect(reopened.package.document.content).toHaveLength(count + 1);
    }),
    { numRuns: 8 },
  );
});

test("canonical section removal does not authorize loss already present in the source", async () => {
  const generated = await createDocx(documentWithSectionEndpoints(1), {
    updateModifiedDate: false,
  });
  const zip = await JSZip.loadAsync(generated);
  const originalXml = await zip.file("word/document.xml")?.async("text");
  if (!originalXml) panic("Expected source document XML");
  zip.file(
    "word/document.xml",
    originalXml.replace("</w:body>", "<w:customXml><w:sectPr/></w:customXml></w:body>"),
  );
  const sourceWithUnmodeledEndpoint = await zip.generateAsync({ type: "arraybuffer" });
  expect(
    readDocumentSectionFacts(await documentXmlOf(sourceWithUnmodeledEndpoint)).sectionCount,
  ).toBe(3);
  const original = await parseDocx(generated, { preloadFonts: false });
  original.originalBuffer = sourceWithUnmodeledEndpoint;
  const edited = await removeCanonicalSectionEndpoints(original, 1);

  await expect(
    repackWithCanonicalStoryRemovals({
      document: edited,
      repack: () => repackDocx(edited, { updateModifiedDate: false }),
    }),
  ).rejects.toThrow(/drop section properties/u);
});
