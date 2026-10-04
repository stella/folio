import { expect, test } from "bun:test";
import JSZip from "jszip";
import { parseDocx } from "./parser";
import { createDocx, repackDocx } from "./rezip";
import { createSimpleDocument, serializeDocument } from "./serializer/documentSerializer";
import { cloneDocumentWithParagraphPropertySources } from "./paragraphPropertySource";

const WORD = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const untouched =
  '<q:p xmlns:p14="http://schemas.microsoft.com/office/word/2010/wordml" p14:paraId="11111111">\n<q:r>  <q:t>untouched</q:t> </q:r></q:p>';
const source = `<?xml version='1.0' encoding='UTF-8'?>\n<q:document xmlns:q='${WORD}'><q:background q:color='FFFFFF'/>\n<q:body>\n${untouched}\n<!-- authored gap -->\t<q:p><q:r><q:t>before</q:t></q:r></q:p>\n<q:sectPr><q:pgSz q:w='12240' q:h='15840'/></q:sectPr>\n</q:body></q:document>`;

const parsedSource = async () => {
  const zip = await JSZip.loadAsync(await createDocx(createSimpleDocument([{ text: "seed" }])));
  zip.file("word/document.xml", source);
  return parseDocx(await zip.generateAsync({ type: "arraybuffer" }));
};

test("trusted body replay preserves authored shell, metadata, gaps and untouched blocks after cloning", async () => {
  const document = cloneDocumentWithParagraphPropertySources(await parsedSource());
  const paragraph = document.package.document.content.at(1);
  if (paragraph?.type !== "paragraph") throw new Error("Missing paragraph");
  paragraph.content = [{ type: "run", content: [{ type: "text", text: "after" }] }];
  const saved = await repackDocx(document);
  const xml = await (await JSZip.loadAsync(saved)).file("word/document.xml")?.async("text");
  expect(xml).toContain(untouched);
  expect(xml).toContain("<!-- authored gap -->\t");
  expect(xml).toStartWith(`<?xml version='1.0' encoding='UTF-8'?>\n<q:document xmlns:q='${WORD}'>`);
  expect(xml).toContain("<q:background q:color='FFFFFF'/>");
  expect(xml).toContain("<q:sectPr><q:pgSz q:w='12240' q:h='15840'/></q:sectPr>");
  const reopened = await parseDocx(saved);
  const changed = reopened.package.document.content.at(1);
  expect(changed?.type === "paragraph" && changed.content.at(0)).toMatchObject({
    type: "run",
    content: [{ type: "text", text: "after" }],
  });
});

test("document metadata changes serialize from the current model without losing body source", async () => {
  const document = await parsedSource();
  document.package.document.background = { color: { rgb: "ABCDEF" } };
  document.package.document.finalSectionProperties = undefined;
  const xml = serializeDocument(document);
  expect(xml).toContain(untouched);
  expect(xml).toContain('w:color="ABCDEF"');
  expect(xml).not.toContain("<q:sectPr>");
});

test("trusted source correspondence failures emit a body diagnostic", async () => {
  const document = await parsedSource();
  const diagnostics: unknown[] = [];
  serializeDocument(document, undefined, {
    xml: source + " ",
    onDiagnostic: (value) => diagnostics.push(value),
  });
  expect(diagnostics).toEqual([{ type: "sourceReplayMismatch", part: "word/document.xml" }]);
});
