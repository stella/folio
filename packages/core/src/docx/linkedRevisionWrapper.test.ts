import { expect, test } from "bun:test";

import { ZipArchive } from "@stll/docx-core/zip";

import { FolioDocxReviewer } from "../ai-edits/headless";
import { toProseDoc } from "../prosemirror/conversion/toProseDoc";
import type { Paragraph } from "../types/document";
import { createEmptyDocument } from "../utils/createDocument";
import { parseParagraph } from "./paragraphParser";
import { FOLIO_REVIEW_HISTORY_NAMESPACE } from "./reviewHistoryNamespace";
import { createDocx } from "./rezip";
import { serializeParagraph } from "./serializer/paragraphSerializer";
import { findChild, parseXmlDocument } from "./xmlParser";

const NAMESPACES =
  'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math"';
const AUTHORED = 'w:author="Reviewer" w:date="2026-01-01T00:00:00Z"';

const read = (xml: string): Paragraph => {
  const root = parseXmlDocument(
    `<root ${NAMESPACES} xmlns:folio="${FOLIO_REVIEW_HISTORY_NAMESPACE}">${xml}</root>`,
  );
  const paragraph = root === null ? null : findChild(root, "w", "p");
  if (paragraph === null) throw new TypeError("The paragraph disappeared.");
  return parseParagraph(paragraph, null, null, null);
};

/** The editor's inline nodes for one paragraph: name, text and mark names. */
const editorInline = (paragraph: Paragraph) => {
  const document = createEmptyDocument();
  document.package.document.content = [paragraph];
  const nodes: { node: string; text: string; marks: string[] }[] = [];
  toProseDoc(document)
    .child(0)
    .forEach((node) => {
      nodes.push({
        node: node.type.name,
        text: node.textContent,
        marks: node.marks
          .map(({ type }) => type.name)
          .filter((name) => !name.startsWith("font")),
      });
    });
  return nodes;
};

const linked = (held: string) =>
  `<w:p><w:hyperlink w:anchor="target"><w:ins w:id="1" ${AUTHORED}><w:dir w:val="rtl"><w:r><w:t>a</w:t></w:r>${held}</w:dir></w:ins></w:hyperlink></w:p>`;

// A transparent wrapper inside a revision inside a link holds what the link
// cannot hold. The revision is hoisted around the link; the wrapper then owns
// the link's pieces, as a nested revision does, so the editor shows the
// content instead of dropping it.
for (const [name, held, expected] of [
  [
    "an equation",
    "<m:oMath><m:r><m:t>x</m:t></m:r></m:oMath>",
    { node: "math", text: "", marks: ["inlineWrapper", "insertion"] },
  ],
  [
    "a nested deletion",
    `<w:del w:id="2" ${AUTHORED}><w:r><w:delText>x</w:delText></w:r></w:del>`,
    {
      node: "text",
      text: "x",
      marks: ["inlineWrapper", "deletion", "hyperlink"],
    },
  ],
  [
    "a simple field",
    '<w:fldSimple w:instr="PAGE"><w:r><w:t>1</w:t></w:r></w:fldSimple>',
    { node: "field", text: "1", marks: ["inlineWrapper", "insertion"] },
  ],
] as const)
  test(`a wrapper holding ${name} under a link revision shows it`, () => {
    const paragraph = read(linked(held));
    const [change] = paragraph.content;
    expect(change?.type).toBe("insertion");
    expect(
      change?.type === "insertion" && change.content.map(({ type }) => type),
    ).toEqual(["inlineWrapper"]);
    expect(editorInline(paragraph)).toEqual([
      {
        node: "text",
        text: "a",
        marks: ["inlineWrapper", "insertion", "hyperlink"],
      },
      expected,
    ]);
    // Saving and reading again keeps the same structure.
    const xml = serializeParagraph(paragraph);
    expect(serializeParagraph(read(xml))).toBe(xml);
  });

test("a canonical reviewer opens an equation in a wrapper under a link revision", async () => {
  const zip = await ZipArchive.read(await createDocx(createEmptyDocument()));
  zip.set(
    "word/document.xml",
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document ${NAMESPACES}><w:body>${linked(
      "<m:oMath><m:r><m:t>x</m:t></m:r></m:oMath>",
    ).replace(
      "<w:p>",
      '<w:p w14:paraId="00000001">',
    )}<w:sectPr/></w:body></w:document>`,
  );
  const reviewer = await FolioDocxReviewer.fromBuffer(
    (await zip.generate({ compression: "store" })).buffer,
    { session: "canonical" },
  );
  const [story] = reviewer.listStories();
  if (story === undefined) throw new TypeError("The story disappeared.");
  expect(
    reviewer.readReviewedStory({ story: story.handle, view: "final" })?.text,
  ).toContain("a");
});
